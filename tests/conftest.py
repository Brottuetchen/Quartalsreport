"""Shared pytest fixtures for the Quartalsreport test suite."""

from __future__ import annotations

import io
import os
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

# ── Environment setup (must happen BEFORE importing server) ──────────────────
# Force test credentials regardless of pre-existing env vars
os.environ["ADMIN_USER"] = "testadmin"
os.environ["ADMIN_PASSWORD"] = "testpass"


@pytest.fixture(scope="session", autouse=True)
def _set_env(tmp_path_factory):
    """Point DATA_DIR to a temporary location so tests don't touch real data."""
    tmp = tmp_path_factory.mktemp("data")
    (tmp / "jobs").mkdir()
    os.environ["DEFAULT_CSV_PATH"] = str(tmp / "default_budget.csv")
    yield
    # Cleanup is handled by pytest tmp_path_factory


@pytest.fixture(scope="session")
def client():
    """A TestClient for the FastAPI app with admin credentials pre-configured."""
    import webapp.server as srv
    from webapp.server import app

    # Patch module-level constants so tests work regardless of container env vars
    srv._ADMIN_USER = "testadmin"
    srv._ADMIN_PASSWORD = "testpass"
    # Also clear any leftover rate-limit state from previous test runs
    srv._admin_rate.clear()

    return TestClient(app, raise_server_exceptions=True)


@pytest.fixture(scope="session")
def admin_auth():
    """HTTP Basic Auth tuple for the test admin user."""
    return ("testadmin", "testpass")


@pytest.fixture(autouse=True)
def _reset_rate_limits():
    """Clear rate-limit state before each test to prevent cross-test interference."""
    import webapp.server as srv
    srv._admin_rate.clear()
    yield


@pytest.fixture
def sample_csv_bytes() -> bytes:
    """Minimal valid budget CSV (UTF-16 LE, tab-separated)."""
    content = (
        "Projekte\tHonorarbereich\tArbeitspaket\tSollstunden Budget\n"
        "1234.01 Testprojekt\tX\t(p) 1.1 Testmeilenstein\t100,00\n"
        "1234.01 Testprojekt\t\t• 1.1.1 Unterpaket\t50,00\n"
    )
    return content.encode("utf-16")


@pytest.fixture
def sample_xml_bytes() -> bytes:
    """Minimal valid time-tracking XML."""
    xml = """\
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE xml [<!ENTITY shy "&#173;">]>
<preprocessedreport xmlns:addi="http://www.untermstrich.com/web/addi/">
  <head>
    <title id="b_timekeyword" format="A4" orientation="landscape">Zeitraum / Stichwort</title>
    <date type="date" unix="-2208992400" usdate="" xlsdate="1900-01-01T00:00:00.000">Mon, 01 Jan 1900 00:00:00 +0100</date>
    <company>
      <line1 type="text">Test GmbH</line1>
    </company>
    <columns>
      <total><field name="number" type="number" width="13">Gesamtsumme</field></total>
      <group by="staff_name" order="TRUE">
        <field name="staff_name" type="string" width="25">Mitarbeiter</field>
        <field name="work_package_name" type="string" width="40">Bereich</field>
        <field name="date" type="date" width="20">Datum</field>
        <field name="time_from" type="time" width="15">Von</field>
        <field name="time_to" type="time" width="15">Bis</field>
        <field name="number" type="number" width="17">Summe</field>
        <field name="project" type="string" width="50">Projekt</field>
        <field name="purpose" type="string" width="60">Leistung</field>
        <field name="is_client_changed" type="bool" width="5">Kundenänderung</field>
        <field name="purpose_internal" type="string" width="15">Bürointerne Informationen</field>
        <total><field name="number" type="number" width="13">Mitarbeitersumme</field></total>
      </group>
    </columns>
  </head>
  <data>
    <group value="Max Mustermann">
      <row>
        <cell name="staff_name" type="text">Max Mustermann</cell>
        <cell name="work_package_name" type="text">(p) 1.1 Testmeilenstein</cell>
        <cell name="date" type="date" unix="1727733600" usdate="01.10.2024" xlsdate="2024-10-01T00:00:00.000">Tue, 01 Oct 2024 00:00:00 +0200</cell>
        <cell name="time_from" type="time" ustime="" xlsdate="1900-01-01T00:00:00.000">Mon, 01 Jan 1900 00:00:00 +0100</cell>
        <cell name="time_to" type="time" ustime="" xlsdate="1900-01-01T00:00:00.000">Mon, 01 Jan 1900 00:00:00 +0100</cell>
        <cell name="number" type="number" usnumber="8,00">8</cell>
        <cell name="project" type="text">1234.01 Testprojekt</cell>
        <cell name="purpose" type="text">Planung</cell>
        <cell name="is_client_changed" type="bool">false</cell>
        <cell name="purpose_internal" type="text"></cell>
      </row>
      <row>
        <cell name="staff_name" type="text">Max Mustermann</cell>
        <cell name="work_package_name" type="text">(p) 1.1 Testmeilenstein</cell>
        <cell name="date" type="date" unix="1730412000" usdate="01.11.2024" xlsdate="2024-11-01T00:00:00.000">Fri, 01 Nov 2024 00:00:00 +0100</cell>
        <cell name="time_from" type="time" ustime="" xlsdate="1900-01-01T00:00:00.000">Mon, 01 Jan 1900 00:00:00 +0100</cell>
        <cell name="time_to" type="time" ustime="" xlsdate="1900-01-01T00:00:00.000">Mon, 01 Jan 1900 00:00:00 +0100</cell>
        <cell name="number" type="number" usnumber="6,00">6</cell>
        <cell name="project" type="text">1234.01 Testprojekt</cell>
        <cell name="purpose" type="text">Abstimmung</cell>
        <cell name="is_client_changed" type="bool">false</cell>
        <cell name="purpose_internal" type="text"></cell>
      </row>
    </group>
  </data>
</preprocessedreport>
"""
    return xml.encode("utf-8")


@pytest.fixture
def valid_update_zip_bytes() -> bytes:
    """A minimal valid OTA-update ZIP with a webapp/ entry."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("webapp/static/test_ota_marker.js", "/* ota test */")
    return buf.getvalue()
