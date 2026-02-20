from __future__ import annotations

import asyncio
import io
import logging
import os
import re
import secrets
import shutil
import time
import uuid
import zipfile
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Deque, Dict, Optional

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .report_generator import generate_quarterly_report
from .api import reports_router


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR.parent / "data"
JOBS_DIR = DATA_DIR / "jobs"
JOBS_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_CSV_PATH = Path(os.getenv("DEFAULT_CSV_PATH", str(DATA_DIR / "default_budget.csv")))
WEBAPP_OVERRIDE_DIR = DATA_DIR / "webapp_override"

# Upload size limits
MAX_CSV_SIZE = 500 * 1024 * 1024   # 500 MB – CSV kann sehr groß sein
MAX_XML_SIZE = 100 * 1024 * 1024   # 100 MB
MAX_ZIP_SIZE =  50 * 1024 * 1024   #  50 MB – OTA-Update

# Admin credentials – set via environment variables (never hardcode)
_ADMIN_USER = os.getenv("ADMIN_USER", "admin")
_ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "")

_http_basic = HTTPBasic()

CLEANUP_RETENTION_DAYS = 7
CLEANUP_INTERVAL_SECONDS = 24 * 60 * 60  # einmal täglich prüfen

# ── Audit-Logger ──────────────────────────────────────────────────────────────
_audit_logger = logging.getLogger("admin_audit")
_audit_logger.setLevel(logging.INFO)
_audit_logger.propagate = False
_audit_handler = logging.FileHandler(str(DATA_DIR / "admin_audit.log"), encoding="utf-8")
_audit_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
_audit_logger.addHandler(_audit_handler)

# ── Rate-Limiter (In-Memory Sliding-Window) ───────────────────────────────────
# Tracks timestamps of failed auth attempts per IP; max 5 per 60 seconds.
_admin_rate: Dict[str, deque] = {}
_RATE_LIMIT = 5
_RATE_WINDOW = 60.0  # seconds


@dataclass
class Job:
    id: str
    created_at: datetime
    csv_path: Path
    xml_path: Path
    output_dir: Path
    requested_quarter: Optional[str]
    output_prefix: Optional[str] = None
    status: str = "queued"  # queued | processing | finished | failed
    progress: int = 0
    message: str = "In Warteschlange"
    result_path: Optional[Path] = None
    error: Optional[str] = None

    def to_dict(self, queue_position: Optional[int]) -> Dict[str, object]:
        return {
            "job_id": self.id,
            "status": self.status,
            "progress": self.progress,
            "message": self.message,
            "queue_position": queue_position,
            "download_available": self.status == "finished" and self.result_path is not None,
            "result_filename": self.result_path.name if self.result_path else None,
            "error": self.error,
        }


app = FastAPI(title="Quartalsreport Generator")

# Include API routers
app.include_router(reports_router)

# Ensure UTF-8 encoding for all responses
@app.middleware("http")
async def add_charset_header(request: Request, call_next):
    response = await call_next(request)
    if "text/html" in response.headers.get("content-type", ""):
        response.headers["content-type"] = "text/html; charset=utf-8"
    elif "application/json" in response.headers.get("content-type", ""):
        response.headers["content-type"] = "application/json; charset=utf-8"
    return response

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

job_store: Dict[str, Job] = {}
job_queue: "asyncio.Queue[Job]" = asyncio.Queue()
pending_jobs: Deque[str] = deque()
queue_lock = asyncio.Lock()
current_job_id: Optional[str] = None


async def _save_upload(upload: UploadFile, destination: Path, max_bytes: int = MAX_XML_SIZE) -> None:
    """Save an uploaded file with a size limit. Raises HTTP 413 if exceeded."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    bytes_written = 0
    try:
        with destination.open("wb") as buffer:
            while True:
                chunk = await upload.read(1024 * 1024)
                if not chunk:
                    break
                bytes_written += len(chunk)
                if bytes_written > max_bytes:
                    raise HTTPException(
                        status_code=413,
                        detail=f"Datei zu groß (max. {max_bytes // (1024 * 1024)} MB)",
                    )
                buffer.write(chunk)
    except HTTPException:
        destination.unlink(missing_ok=True)
        raise
    finally:
        await upload.close()


def _safe_filename(name: str, fallback: str) -> str:
    """Strip path components and allow only safe characters in a filename."""
    stem = Path(name).name  # removes any directory traversal
    safe = re.sub(r"[^\w.\-]", "_", stem)
    return safe[:128] or fallback


def _job_progress_updater(job: Job):
    def _update(progress: int, message: str) -> None:
        job.progress = max(0, min(progress, 100))
        job.message = message
    return _update



async def worker() -> None:
    global current_job_id
    while True:
        job = await job_queue.get()
        async with queue_lock:
            if pending_jobs and pending_jobs[0] == job.id:
                pending_jobs.popleft()
            current_job_id = job.id

        job.status = "processing"
        job.message = "Starte Verarbeitung"
        job.progress = 15

        try:
            result_path = generate_quarterly_report(
                csv_path=job.csv_path,
                xml_path=job.xml_path,
                output_dir=job.output_dir,
                output_name_prefix=job.output_prefix,
                requested_quarter=job.requested_quarter,
                progress_cb=_job_progress_updater(job),
            )
            job.result_path = result_path
            job.status = "finished"
            job.progress = 100
            job.message = "Fertig"
        except Exception as exc:  # pylint: disable=broad-except
            job.status = "failed"
            job.progress = 100
            job.message = "Fehlgeschlagen"
            job.error = str(exc)
        finally:
            job_queue.task_done()
            async with queue_lock:
                current_job_id = None


@app.on_event("startup")
async def on_startup() -> None:
    asyncio.create_task(worker())
    asyncio.create_task(garbage_collector())
    # Einmalig beim Start Altlasten bereinigen
    await cleanup_stale_jobs()


async def garbage_collector() -> None:
    while True:
        await asyncio.sleep(CLEANUP_INTERVAL_SECONDS)
        await cleanup_stale_jobs()


async def cleanup_stale_jobs() -> None:
    cutoff = datetime.utcnow() - timedelta(days=CLEANUP_RETENTION_DAYS)
    async with queue_lock:
        removable_ids = [
            job_id
            for job_id, job in list(job_store.items())
            if job.created_at < cutoff and job.status in {"finished", "failed"}
        ]
        for job_id in removable_ids:
            job = job_store.pop(job_id, None)
            if job_id in pending_jobs:
                pending_jobs.remove(job_id)
            if job and job.output_dir.exists():
                shutil.rmtree(job.output_dir, ignore_errors=True)

    # Auch verwaiste Verzeichnisse entfernen, die keinen Job mehr haben
    for child in JOBS_DIR.iterdir():
        if not child.is_dir():
            continue
        job_id = child.name
        if job_id in job_store:
            continue
        try:
            mtime = datetime.utcfromtimestamp(child.stat().st_mtime)
        except OSError:
            continue
        if mtime < cutoff:
            shutil.rmtree(child, ignore_errors=True)


def _queue_position(job_id: str) -> Optional[int]:
    if job_id not in job_store:
        return None
    job = job_store[job_id]
    if job.status == "processing":
        return 0
    if job.status != "queued":
        return 0
    try:
        idx = pending_jobs.index(job_id)
        return idx + 1
    except ValueError:
        return None


def _require_admin(
    request: Request,
    credentials: HTTPBasicCredentials = Depends(_http_basic),
) -> str:
    """Dependency that enforces admin credentials via HTTP Basic Auth.
    Returns the client IP for use in audit log entries.
    """
    ip = request.client.host if request.client else "unknown"

    # ── Rate limit: check before verifying credentials ────────────────────────
    now = time.monotonic()
    window = _admin_rate.setdefault(ip, deque())
    while window and now - window[0] > _RATE_WINDOW:
        window.popleft()
    if len(window) >= _RATE_LIMIT:
        _audit_logger.warning("Rate-limit reached ip=%s", ip)
        raise HTTPException(
            status_code=429,
            detail="Zu viele fehlgeschlagene Versuche. Bitte warten.",
        )

    if not _ADMIN_PASSWORD:
        raise HTTPException(status_code=503, detail="Admin-Bereich nicht konfiguriert (ADMIN_PASSWORD fehlt)")

    user_ok = secrets.compare_digest(credentials.username.encode(), _ADMIN_USER.encode())
    pass_ok = secrets.compare_digest(credentials.password.encode(), _ADMIN_PASSWORD.encode())
    if not (user_ok and pass_ok):
        window.append(time.monotonic())
        _audit_logger.warning("Auth failed ip=%s user=%r", ip, credentials.username)
        raise HTTPException(
            status_code=401,
            detail="Ungültige Zugangsdaten",
            headers={"WWW-Authenticate": "Basic realm=\"Quartalsreport Admin\""},
        )

    _audit_logger.info("Auth ok ip=%s user=%r", ip, credentials.username)
    return ip


@app.get("/admin/budget/info")
async def admin_budget_info(ip: str = Depends(_require_admin)):
    """Return metadata about the currently stored default budget CSV."""
    if not DEFAULT_CSV_PATH.exists():
        return JSONResponse({"exists": False})
    stat = DEFAULT_CSV_PATH.stat()
    return JSONResponse({
        "exists": True,
        "filename": DEFAULT_CSV_PATH.name,
        "size_bytes": stat.st_size,
        "last_modified": datetime.utcfromtimestamp(stat.st_mtime).isoformat() + "Z",
    })


@app.post("/admin/budget")
async def admin_upload_budget(csv_file: UploadFile = File(...), ip: str = Depends(_require_admin)):
    """Replace the default budget CSV. Only accessible to admins."""
    if csv_file.content_type not in {
        "text/csv", "application/vnd.ms-excel", "text/plain", "application/octet-stream"
    }:
        raise HTTPException(status_code=400, detail="CSV-Datei wird erwartet")
    DEFAULT_CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    await _save_upload(csv_file, DEFAULT_CSV_PATH, max_bytes=MAX_CSV_SIZE)
    size = DEFAULT_CSV_PATH.stat().st_size
    _audit_logger.info("Budget CSV uploaded ip=%s size=%d", ip, size)
    return JSONResponse({"status": "ok", "filename": DEFAULT_CSV_PATH.name})


@app.post("/admin/update")
async def admin_ota_update(zip_file: UploadFile = File(...), ip: str = Depends(_require_admin)):
    """
    OTA code update: upload a zip created via `git archive HEAD --format=zip`.
    Files under the `webapp/` prefix are extracted to /app/webapp/ and persisted
    in /app/data/webapp_override/ so they survive container restarts.
    uvicorn's --reload mode picks up the changed files automatically.
    """
    if not (zip_file.filename or "").lower().endswith(".zip"):
        raise HTTPException(status_code=400, detail="ZIP-Datei wird erwartet")

    # Read with size limit into memory
    raw = b""
    while True:
        chunk = await zip_file.read(1024 * 1024)
        if not chunk:
            break
        raw += chunk
        if len(raw) > MAX_ZIP_SIZE:
            await zip_file.close()
            raise HTTPException(
                status_code=413,
                detail=f"ZIP zu groß (max. {MAX_ZIP_SIZE // (1024 * 1024)} MB)",
            )
    await zip_file.close()

    try:
        zf = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        raise HTTPException(status_code=400, detail="Ungültige ZIP-Datei")

    # Security: reject symlinks (Unix external_attr bits)
    for info in zf.infolist():
        if (info.external_attr >> 16) & 0o170000 == 0o120000:
            raise HTTPException(status_code=400, detail=f"ZIP enthält Symlinks: {info.filename}")

    webapp_entries = [n for n in zf.namelist() if n.startswith("webapp/") and not n.endswith("/")]
    if not webapp_entries:
        raise HTTPException(status_code=400, detail="ZIP enthält kein 'webapp/'-Verzeichnis")

    # Security: reject any path that would escape the target directory
    for name in webapp_entries:
        parts = Path(name).parts
        if ".." in parts or any(p.startswith("/") for p in parts):
            raise HTTPException(status_code=400, detail=f"Unzulässiger Pfad im ZIP: {name}")

    WEBAPP_OVERRIDE_DIR.mkdir(parents=True, exist_ok=True)

    for name in webapp_entries:
        # Strip leading "webapp/" prefix → relative path within the webapp package
        rel = Path(*Path(name).parts[1:])
        for target_root in (BASE_DIR, WEBAPP_OVERRIDE_DIR):
            dest = target_root / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(zf.read(name))

    zf.close()
    _audit_logger.info("OTA update applied ip=%s files=%d", ip, len(webapp_entries))
    return JSONResponse({"status": "reloading", "files_updated": len(webapp_entries)})


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/jobs", response_class=JSONResponse)
async def create_job(
    csv_file: Optional[UploadFile] = File(default=None),
    xml_file: UploadFile = File(...),
    quarter: Optional[str] = Form(default=None),
):
    if csv_file and csv_file.content_type not in {"text/csv", "application/vnd.ms-excel", "text/plain", "application/octet-stream"}:
        raise HTTPException(status_code=400, detail="CSV-Datei wird erwartet")
    if not xml_file.filename or not xml_file.filename.lower().endswith(".xml"):
        raise HTTPException(status_code=400, detail="XML-Datei wird erwartet")
    if csv_file is None and not DEFAULT_CSV_PATH.exists():
        raise HTTPException(status_code=400, detail="Keine Standard-CSV hinterlegt. Bitte CSV-Datei hochladen.")

    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    csv_path = job_dir / DEFAULT_CSV_PATH.name
    xml_path = job_dir / _safe_filename(xml_file.filename or "", "source.xml")

    if csv_file:
        # Use uploaded CSV for this job only – the global default is managed via /admin/budget
        csv_path = job_dir / _safe_filename(csv_file.filename or "", "source.csv")
        await _save_upload(csv_file, csv_path, max_bytes=MAX_CSV_SIZE)
    else:
        shutil.copy2(DEFAULT_CSV_PATH, csv_path)
    await _save_upload(xml_file, xml_path, max_bytes=MAX_XML_SIZE)

    job = Job(
        id=job_id,
        created_at=datetime.utcnow(),
        csv_path=csv_path,
        xml_path=xml_path,
        output_dir=job_dir,
        requested_quarter=quarter if quarter else None,
        output_prefix=xml_path.stem or None,
    )

    async with queue_lock:
        job_store[job_id] = job
        pending_jobs.append(job_id)
        await job_queue.put(job)
        position = _queue_position(job_id)

    return JSONResponse({
        "job_id": job_id,
        "status": job.status,
        "queue_position": position,
        "message": job.message,
    })


@app.get("/api/jobs/{job_id}", response_class=JSONResponse)
async def job_status(job_id: str):
    job = job_store.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job nicht gefunden")
    async with queue_lock:
        position = _queue_position(job_id)
    return JSONResponse(job.to_dict(position))


@app.get("/api/jobs/{job_id}/download")
async def job_download(job_id: str):
    job = job_store.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job nicht gefunden")
    if job.status != "finished" or not job.result_path:
        raise HTTPException(status_code=409, detail="Job ist noch nicht abgeschlossen")

    filename = job.result_path.name
    return FileResponse(path=job.result_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.delete("/api/jobs/{job_id}")
async def delete_job(job_id: str):
    job = job_store.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job nicht gefunden")
    if job.status == "processing":
        raise HTTPException(status_code=409, detail="Job wird gerade verarbeitet")

    async with queue_lock:
        if job_id in pending_jobs:
            pending_jobs.remove(job_id)
        job_store.pop(job_id, None)

    if job.output_dir.exists():
        shutil.rmtree(job.output_dir, ignore_errors=True)

    return JSONResponse({"status": "deleted"})


@app.get("/healthz")
async def healthcheck():
    return {"status": "ok"}


__all__ = ["app"]
