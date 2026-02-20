"""Tests for the OTA update endpoint: path traversal, symlinks, auth, happy path."""

from __future__ import annotations

import io
import os
import stat
import struct
import zipfile

import pytest
from fastapi.testclient import TestClient


def _make_zip(entries: dict[str, bytes]) -> bytes:
    """Helper to build a ZIP from {name: content} dict."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)
    return buf.getvalue()


def _make_zip_with_symlink(link_name: str, target: str) -> bytes:
    """Build a ZIP that contains a Unix symlink entry."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        info = zipfile.ZipInfo(link_name)
        # Set Unix symlink mode in external_attr (upper 16 bits)
        info.external_attr = (stat.S_IFLNK | 0o777) << 16
        zf.writestr(info, target)
    return buf.getvalue()


# ── Auth required ─────────────────────────────────────────────────────────────

def test_update_requires_auth(client: TestClient, valid_update_zip_bytes):
    """OTA update endpoint must require authentication."""
    res = client.post(
        "/admin/update",
        files={"zip_file": ("update.zip", io.BytesIO(valid_update_zip_bytes), "application/zip")},
    )
    assert res.status_code == 401


# ── Input validation ──────────────────────────────────────────────────────────

def test_non_zip_rejected(client: TestClient, admin_auth):
    """Uploading a non-ZIP file should be rejected with 400."""
    res = client.post(
        "/admin/update",
        auth=admin_auth,
        files={"zip_file": ("update.txt", io.BytesIO(b"not a zip"), "text/plain")},
    )
    assert res.status_code == 400


def test_bad_zip_content_rejected(client: TestClient, admin_auth):
    """Corrupted ZIP data should be rejected with 400."""
    res = client.post(
        "/admin/update",
        auth=admin_auth,
        files={"zip_file": ("update.zip", io.BytesIO(b"PK corrupt data"), "application/zip")},
    )
    assert res.status_code == 400


def test_zip_no_webapp_dir(client: TestClient, admin_auth):
    """ZIP without a webapp/ directory should be rejected with 400."""
    data = _make_zip({"other/file.txt": b"content"})
    res = client.post(
        "/admin/update",
        auth=admin_auth,
        files={"zip_file": ("update.zip", io.BytesIO(data), "application/zip")},
    )
    assert res.status_code == 400


def test_zip_path_traversal(client: TestClient, admin_auth):
    """ZIP containing path-traversal entries should be rejected with 400."""
    data = _make_zip({
        "webapp/../../etc/passwd": b"root:x:0:0:root:/root:/bin/bash",
    })
    res = client.post(
        "/admin/update",
        auth=admin_auth,
        files={"zip_file": ("update.zip", io.BytesIO(data), "application/zip")},
    )
    assert res.status_code == 400


def test_zip_with_symlink(client: TestClient, admin_auth):
    """ZIP containing a symlink entry should be rejected with 400."""
    data = _make_zip_with_symlink("webapp/evil_link", "/etc/passwd")
    res = client.post(
        "/admin/update",
        auth=admin_auth,
        files={"zip_file": ("update.zip", io.BytesIO(data), "application/zip")},
    )
    assert res.status_code == 400


# ── Happy path ────────────────────────────────────────────────────────────────

def test_valid_update_zip(client: TestClient, admin_auth, valid_update_zip_bytes):
    """A valid OTA update ZIP should be accepted and return files_updated count."""
    res = client.post(
        "/admin/update",
        auth=admin_auth,
        files={"zip_file": ("update.zip", io.BytesIO(valid_update_zip_bytes), "application/zip")},
    )
    assert res.status_code == 200
    data = res.json()
    assert data.get("files_updated", 0) >= 1
    assert data.get("status") == "reloading"
