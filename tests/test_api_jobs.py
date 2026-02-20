"""Tests for the standard-report job API (/api/jobs)."""

from __future__ import annotations

import io

import pytest
from fastapi.testclient import TestClient


# ── Job creation ──────────────────────────────────────────────────────────────

def test_create_job_no_csv_no_default(client: TestClient, sample_xml_bytes):
    """Without a default CSV and without uploading one, job creation should fail (400)."""
    import os, webapp.server as srv
    # Temporarily point DEFAULT_CSV_PATH to a non-existent file
    original = srv.DEFAULT_CSV_PATH
    from pathlib import Path
    srv.DEFAULT_CSV_PATH = Path("/tmp/nonexistent_budget_xyz.csv")
    try:
        res = client.post(
            "/api/jobs",
            files={"xml_file": ("data.xml", io.BytesIO(sample_xml_bytes), "application/xml")},
        )
        assert res.status_code == 400
        assert "CSV" in res.json().get("detail", "")
    finally:
        srv.DEFAULT_CSV_PATH = original


def test_create_job_with_both_files(client: TestClient, sample_csv_bytes, sample_xml_bytes):
    """Submitting both CSV and XML should create a job and return a job_id."""
    res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
            "xml_file": ("data.xml", io.BytesIO(sample_xml_bytes), "application/xml"),
        },
    )
    assert res.status_code == 200
    data = res.json()
    assert "job_id" in data
    assert data["status"] in ("queued", "processing", "finished", "failed")


def test_create_job_xml_wrong_extension(client: TestClient, sample_csv_bytes):
    """Non-XML file as xml_file should be rejected with 400."""
    res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
            "xml_file": ("data.txt", io.BytesIO(b"<not xml>"), "text/plain"),
        },
    )
    assert res.status_code == 400


# ── Job status ────────────────────────────────────────────────────────────────

def test_job_status_not_found(client: TestClient):
    """Fetching status of a non-existent job should return 404."""
    res = client.get("/api/jobs/doesnotexist123")
    assert res.status_code == 404


def test_job_status_after_creation(client: TestClient, sample_csv_bytes, sample_xml_bytes):
    """After creating a job, its status should be fetchable."""
    create_res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
            "xml_file": ("data.xml", io.BytesIO(sample_xml_bytes), "application/xml"),
        },
    )
    assert create_res.status_code == 200
    job_id = create_res.json()["job_id"]

    status_res = client.get(f"/api/jobs/{job_id}")
    assert status_res.status_code == 200
    assert status_res.json()["job_id"] == job_id


# ── Download ──────────────────────────────────────────────────────────────────

def test_download_not_finished(client: TestClient, sample_csv_bytes, sample_xml_bytes):
    """Attempting to download a job that isn't finished should return 409."""
    # Create a job
    create_res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
            "xml_file": ("data.xml", io.BytesIO(sample_xml_bytes), "application/xml"),
        },
    )
    job_id = create_res.json()["job_id"]

    # Immediately try to download (likely still queued/processing)
    dl_res = client.get(f"/api/jobs/{job_id}/download")
    # Either 409 (not finished) or 200 (finished very fast) – both are valid
    assert dl_res.status_code in (200, 409)


def test_download_nonexistent_job(client: TestClient):
    """Downloading a non-existent job should return 404."""
    res = client.get("/api/jobs/nonexistent/download")
    assert res.status_code == 404


# ── Delete ────────────────────────────────────────────────────────────────────

def test_delete_nonexistent_job(client: TestClient):
    """Deleting a non-existent job should return 404."""
    res = client.delete("/api/jobs/nonexistent")
    assert res.status_code == 404


def test_delete_job(client: TestClient, sample_csv_bytes, sample_xml_bytes):
    """A queued job should be deletable."""
    create_res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
            "xml_file": ("data.xml", io.BytesIO(sample_xml_bytes), "application/xml"),
        },
    )
    job_id = create_res.json()["job_id"]

    del_res = client.delete(f"/api/jobs/{job_id}")
    # Either 200 (deleted) or 409 (currently processing)
    assert del_res.status_code in (200, 409)
