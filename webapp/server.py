from __future__ import annotations

import asyncio
import shutil
import uuid
from collections import deque
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Deque, Dict, Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .report_generator import generate_quarterly_report


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR.parent / "data"
JOBS_DIR = DATA_DIR / "jobs"
JOBS_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class Job:
    id: str
    created_at: datetime
    csv_path: Path
    xml_path: Path
    output_dir: Path
    requested_quarter: Optional[str]
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
            "error": self.error,
        }


app = FastAPI(title="Quartalsreport Generator")

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

job_store: Dict[str, Job] = {}
job_queue: "asyncio.Queue[Job]" = asyncio.Queue()
pending_jobs: Deque[str] = deque()
queue_lock = asyncio.Lock()
current_job_id: Optional[str] = None


async def _save_upload(upload: UploadFile, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as buffer:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            buffer.write(chunk)
    await upload.close()


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


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/jobs", response_class=JSONResponse)
async def create_job(
    csv_file: UploadFile = File(...),
    xml_file: UploadFile = File(...),
    quarter: Optional[str] = Form(default=None),
):
    if csv_file.content_type not in {"text/csv", "application/vnd.ms-excel", "text/plain", "application/octet-stream"}:
        raise HTTPException(status_code=400, detail="CSV-Datei wird erwartet")
    if not xml_file.filename.lower().endswith(".xml"):
        raise HTTPException(status_code=400, detail="XML-Datei wird erwartet")

    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    csv_path = job_dir / (csv_file.filename or "source.csv")
    xml_path = job_dir / (xml_file.filename or "source.xml")

    await _save_upload(csv_file, csv_path)
    await _save_upload(xml_file, xml_path)

    job = Job(
        id=job_id,
        created_at=datetime.utcnow(),
        csv_path=csv_path,
        xml_path=xml_path,
        output_dir=job_dir,
        requested_quarter=quarter if quarter else None,
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

