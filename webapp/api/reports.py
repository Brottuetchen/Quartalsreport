"""API endpoints for flexible report generation."""

from __future__ import annotations

import asyncio
import logging
import re
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

logger = logging.getLogger(__name__)

# Upload size limits (kept in sync with server.py constants)
_MAX_CSV_SIZE = 500 * 1024 * 1024   # 500 MB
_MAX_XML_SIZE = 100 * 1024 * 1024   # 100 MB

from ..models import ReportConfig, ReportType, TimeGrouping  # noqa: E402
from ..services import FlexibleReportGenerator


router = APIRouter(prefix="/api/reports", tags=["reports"])


@router.post("/flexible")
async def generate_flexible_report(
    # Report type and date range
    report_type: str = Form(..., description="Type of report: quarterly, custom_period, monthly, yearly"),
    start_date: str = Form(..., description="Start date (YYYY-MM-DD)"),
    end_date: str = Form(..., description="End date (YYYY-MM-DD)"),
    time_grouping: str = Form("monthly", description="Time grouping: monthly, period, weekly, none"),

    # File uploads
    csv_file: UploadFile = File(..., description="CSV budget file"),
    xml_file: UploadFile = File(..., description="XML timesheet file"),

    # Optional filters
    projects: Optional[str] = Form(None, description="Comma-separated project codes"),
    employees: Optional[str] = Form(None, description="Comma-separated employee names"),

    # Report options
    include_bonus_calc: bool = Form(True),
    include_budget_overview: bool = Form(True),
    include_summary_sheet: bool = Form(True),
    include_quarterly_summary: bool = Form(True),
    exclude_special_projects: bool = Form(False),
) -> FileResponse:
    """
    Generate a flexible report with custom configuration.

    Supports:
    - Custom date ranges (e.g., 15.08-15.09)
    - Different time groupings (monthly blocks, single period, weekly, or total)
    - Project and employee filtering
    - Configurable report components

    Returns the generated Excel file.
    """

    # Parse dates
    try:
        start = datetime.strptime(start_date, "%Y-%m-%d").date()
        end = datetime.strptime(end_date, "%Y-%m-%d").date()
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"Invalid date format: {e}")

    # Parse report type
    try:
        report_type_enum = ReportType(report_type)
    except ValueError:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid report type. Must be one of: {', '.join([t.value for t in ReportType])}"
        )

    # Parse time grouping
    try:
        time_grouping_enum = TimeGrouping(time_grouping)
    except ValueError:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid time grouping. Must be one of: {', '.join([t.value for t in TimeGrouping])}"
        )

    # Parse optional filters
    project_list = [p.strip() for p in projects.split(",")] if projects else None
    employee_list = [e.strip() for e in employees.split(",")] if employees else None

    # Create config
    try:
        config = ReportConfig(
            report_type=report_type_enum,
            start_date=start,
            end_date=end,
            time_grouping=time_grouping_enum,
            projects=project_list,
            employees=employee_list,
            include_bonus_calc=include_bonus_calc,
            include_budget_overview=include_budget_overview,
            include_summary_sheet=include_summary_sheet,
            include_quarterly_summary=include_quarterly_summary,
            exclude_special_projects=exclude_special_projects,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    # Create temporary directory for this job
    job_id = str(uuid.uuid4())
    job_dir = Path("data/jobs") / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Save uploaded files with size limits and safe filenames
        csv_path = job_dir / _safe_filename(csv_file.filename or "", "budget.csv")
        xml_path = job_dir / _safe_filename(xml_file.filename or "", "timesheets.xml")

        await _save_upload(csv_file, csv_path, max_bytes=_MAX_CSV_SIZE)
        await _save_upload(xml_file, xml_path, max_bytes=_MAX_XML_SIZE)

        # Generate output filename
        if report_type_enum == ReportType.QUARTERLY:
            quarter_str = f"Q{(start.month - 1) // 3 + 1}-{start.year}"
            output_filename_base = f"{quarter_str}.xlsm"
        else:
            output_filename_base = f"Report_{start.strftime('%Y%m%d')}-{end.strftime('%Y%m%d')}.xlsx"

        xml_prefix = Path(_safe_filename(xml_file.filename or "", "")).stem
        if xml_prefix:
            output_filename_base = f"{xml_prefix}_{output_filename_base}"

        output_path = job_dir / output_filename_base

        # Generate report
        generator = FlexibleReportGenerator(
            config=config,
            csv_path=csv_path,
            xml_path=xml_path,
        )

        result_path = await asyncio.to_thread(generator.generate, output_path)

        # Determine final filename from the result path
        final_output_filename = result_path.name

        # Return the file
        return FileResponse(
            path=result_path,
            filename=final_output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except HTTPException:
        raise
    except Exception as e:
        # Clean up on error
        import shutil
        logger.error("Report generation failed", exc_info=True)
        if job_dir.exists():
            shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail="Interner Fehler bei der Report-Erstellung")


@router.get("/types")
async def get_report_types() -> JSONResponse:
    """Get available report types."""
    return JSONResponse({
        "report_types": [
            {"value": t.value, "label": t.name.replace("_", " ").title()}
            for t in ReportType
        ],
        "time_groupings": [
            {"value": t.value, "label": t.name.replace("_", " ").title()}
            for t in TimeGrouping
        ],
    })


async def _save_upload(upload: UploadFile, destination: Path, max_bytes: int = _MAX_XML_SIZE) -> None:
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
    stem = Path(name).name
    safe = re.sub(r"[^\w.\-]", "_", stem)
    return safe[:128] or fallback
