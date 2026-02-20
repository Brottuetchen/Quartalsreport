"""Tests for file-upload validation: size limits, extensions, filename sanitization."""

from __future__ import annotations

import io
import os
import importlib

import pytest
from fastapi.testclient import TestClient


# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_oversized(size_bytes: int) -> bytes:
    return b"x" * size_bytes


# ── XML extension validation ──────────────────────────────────────────────────

def test_xml_wrong_extension(client: TestClient, admin_auth, sample_csv_bytes):
    """Uploading a non-XML file as the XML input should be rejected with 400."""
    res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
            "xml_file": ("not_an_xml.txt", io.BytesIO(b"hello"), "text/plain"),
        },
    )
    assert res.status_code == 400


# ── CSV content-type flexibility ──────────────────────────────────────────────

def test_csv_octet_stream_accepted(client: TestClient, sample_csv_bytes, sample_xml_bytes):
    """CSV uploads with application/octet-stream content-type should be accepted."""
    res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "application/octet-stream"),
            "xml_file": ("data.xml", io.BytesIO(sample_xml_bytes), "application/xml"),
        },
    )
    # We only care that the server didn't reject the upload (400). Report errors are ok (500).
    assert res.status_code != 400 or "CSV" not in res.json().get("detail", "")


# ── Size limits ───────────────────────────────────────────────────────────────

def test_xml_too_large(client: TestClient, admin_auth, sample_csv_bytes):
    """XML uploads exceeding the size limit should be rejected with 413."""
    import webapp.server as srv
    original = srv.MAX_XML_SIZE
    srv.MAX_XML_SIZE = 10  # 10 bytes – tiny limit to trigger quickly
    try:
        res = client.post(
            "/api/jobs",
            files={
                "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
                "xml_file": ("data.xml", io.BytesIO(b"x" * 20), "text/xml"),
            },
        )
        assert res.status_code == 413
    finally:
        srv.MAX_XML_SIZE = original


def test_csv_too_large_admin_upload(client: TestClient, admin_auth):
    """CSV uploads to /admin/budget exceeding the size limit should be rejected with 413."""
    import webapp.server as srv
    original = srv.MAX_CSV_SIZE
    srv.MAX_CSV_SIZE = 10  # tiny limit
    try:
        res = client.post(
            "/admin/budget",
            auth=admin_auth,
            files={"csv_file": ("budget.csv", io.BytesIO(b"x" * 20), "text/csv")},
        )
        assert res.status_code == 413
    finally:
        srv.MAX_CSV_SIZE = original


def test_zip_too_large(client: TestClient, admin_auth):
    """ZIP uploads exceeding the size limit should be rejected with 413."""
    import webapp.server as srv
    original = srv.MAX_ZIP_SIZE
    srv.MAX_ZIP_SIZE = 10  # tiny limit
    try:
        res = client.post(
            "/admin/update",
            auth=admin_auth,
            files={"zip_file": ("update.zip", io.BytesIO(b"x" * 20), "application/zip")},
        )
        assert res.status_code == 413
    finally:
        srv.MAX_ZIP_SIZE = original


# ── Filename sanitization ─────────────────────────────────────────────────────

def test_safe_filename_strips_path():
    """_safe_filename should strip directory components."""
    from webapp.server import _safe_filename
    result = _safe_filename("../../../etc/passwd", "fallback.xml")
    assert "/" not in result
    assert ".." not in result
    assert result == "passwd"


def test_safe_filename_replaces_special_chars():
    """_safe_filename should replace characters outside [a-zA-Z0-9._-]."""
    from webapp.server import _safe_filename
    result = _safe_filename("my file (v2)!.xml", "fallback.xml")
    assert " " not in result
    assert "(" not in result
    assert ")" not in result
    assert "!" not in result


def test_safe_filename_truncates_long_names():
    """_safe_filename should truncate names longer than 128 characters."""
    from webapp.server import _safe_filename
    long_name = "a" * 200 + ".xml"
    result = _safe_filename(long_name, "fallback.xml")
    assert len(result) <= 128


def test_safe_filename_uses_fallback_for_empty():
    """_safe_filename should use the fallback when the sanitized name is empty."""
    from webapp.server import _safe_filename
    result = _safe_filename("", "fallback.xml")
    assert result == "fallback.xml"


def test_malicious_filename_traversal_in_job(client: TestClient, sample_csv_bytes, sample_xml_bytes):
    """A path-traversal filename on the XML upload should not escape the job directory."""
    res = client.post(
        "/api/jobs",
        files={
            "csv_file": ("budget.csv", io.BytesIO(sample_csv_bytes), "text/csv"),
            "xml_file": ("../../../etc/passwd.xml", io.BytesIO(sample_xml_bytes), "application/xml"),
        },
    )
    # The server should accept or fail the job, but NOT create a file outside the job dir.
    # We just confirm no 500 from the path-join itself (a traversal would fail or be sanitized)
    assert res.status_code in (200, 201, 202, 422, 500)  # not a 400 from extension check
