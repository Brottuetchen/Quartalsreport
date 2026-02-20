/* Admin panel logic for Quartalsreport Generator */
(function () {
  'use strict';

  // ── Session storage key for credentials ──────────────────────────────────
  const CRED_KEY = 'admin_credentials';

  function getCredentials() {
    const raw = sessionStorage.getItem(CRED_KEY);
    return raw ? JSON.parse(raw) : null;
  }

  function saveCredentials(user, password) {
    sessionStorage.setItem(CRED_KEY, JSON.stringify({ user, password }));
  }

  function clearCredentials() {
    sessionStorage.removeItem(CRED_KEY);
  }

  function basicAuthHeader(creds) {
    return 'Basic ' + btoa(creds.user + ':' + creds.password);
  }

  // ── UI helpers ────────────────────────────────────────────────────────────
  function showMsg(el, text, isError) {
    el.textContent = text;
    el.className = 'admin-msg' + (isError ? ' error' : ' success');
    el.classList.remove('hidden');
  }

  function formatBytes(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function formatDate(iso) {
    try {
      return new Date(iso).toLocaleString('de-DE');
    } catch {
      return iso;
    }
  }

  // ── Budget info ───────────────────────────────────────────────────────────
  async function loadBudgetInfo(creds) {
    const infoBox = document.getElementById('budget-info');
    try {
      const res = await fetch('/admin/budget/info', {
        headers: { Authorization: basicAuthHeader(creds) },
      });
      if (!res.ok) {
        infoBox.textContent = 'Fehler beim Laden der Budget-Info.';
        return;
      }
      const data = await res.json();
      if (!data.exists) {
        infoBox.textContent = 'Noch keine Budget-CSV hinterlegt.';
      } else {
        infoBox.innerHTML =
          '<strong>Aktuelle CSV:</strong> ' + data.filename +
          '<br><strong>Größe:</strong> ' + formatBytes(data.size_bytes) +
          '<br><strong>Zuletzt geändert:</strong> ' + formatDate(data.last_modified);
      }
    } catch {
      infoBox.textContent = 'Verbindungsfehler.';
    }
  }

  // ── Show / hide admin sections ────────────────────────────────────────────
  function showAdminContent(creds) {
    document.getElementById('admin-login').classList.add('hidden');
    document.getElementById('admin-content').classList.remove('hidden');
    loadBudgetInfo(creds);
  }

  function showAdminLogin() {
    document.getElementById('admin-login').classList.remove('hidden');
    document.getElementById('admin-content').classList.add('hidden');
    document.getElementById('admin-login-error').classList.add('hidden');
  }

  // ── Init ──────────────────────────────────────────────────────────────────
  document.addEventListener('DOMContentLoaded', function () {
    const loginSection  = document.getElementById('admin-login');
    const loginBtn      = document.getElementById('admin-login-btn');
    const loginError    = document.getElementById('admin-login-error');
    const logoutBtn     = document.getElementById('admin-logout-btn');
    const csvUploadBtn  = document.getElementById('admin-csv-upload-btn');
    const updateBtn     = document.getElementById('admin-update-btn');

    if (!loginBtn) return; // admin tab not present

    // Restore session if already logged in
    const stored = getCredentials();
    if (stored) showAdminContent(stored);

    // ── Login ──
    loginBtn.addEventListener('click', async function () {
      const user     = document.getElementById('admin-user').value.trim();
      const password = document.getElementById('admin-password').value;
      const creds    = { user, password };

      try {
        const res = await fetch('/admin/budget/info', {
          headers: { Authorization: basicAuthHeader(creds) },
        });
        if (res.status === 401 || res.status === 403) {
          loginError.classList.remove('hidden');
          return;
        }
        loginError.classList.add('hidden');
        saveCredentials(user, password);
        showAdminContent(creds);
      } catch {
        loginError.textContent = 'Verbindungsfehler.';
        loginError.classList.remove('hidden');
      }
    });

    // ── Logout ──
    logoutBtn.addEventListener('click', function () {
      clearCredentials();
      showAdminLogin();
    });

    // ── Upload budget CSV ──
    csvUploadBtn.addEventListener('click', async function () {
      const creds   = getCredentials();
      const fileInput = document.getElementById('admin-csv-file');
      const msgEl   = document.getElementById('admin-csv-msg');

      if (!creds) { showAdminLogin(); return; }
      if (!fileInput.files || fileInput.files.length === 0) {
        showMsg(msgEl, 'Bitte eine CSV-Datei auswählen.', true);
        return;
      }

      const fd = new FormData();
      fd.append('csv_file', fileInput.files[0]);

      try {
        const res = await fetch('/admin/budget', {
          method: 'POST',
          headers: { Authorization: basicAuthHeader(creds) },
          body: fd,
        });
        if (res.status === 401) { clearCredentials(); showAdminLogin(); return; }
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          showMsg(msgEl, 'Fehler: ' + (err.detail || res.statusText), true);
          return;
        }
        showMsg(msgEl, 'Budget-CSV erfolgreich aktualisiert.', false);
        fileInput.value = '';
        await loadBudgetInfo(creds);
      } catch {
        showMsg(msgEl, 'Verbindungsfehler beim Upload.', true);
      }
    });

    // ── Upload OTA update ──
    updateBtn.addEventListener('click', async function () {
      const creds     = getCredentials();
      const fileInput = document.getElementById('admin-update-file');
      const msgEl     = document.getElementById('admin-update-msg');

      if (!creds) { showAdminLogin(); return; }
      if (!fileInput.files || fileInput.files.length === 0) {
        showMsg(msgEl, 'Bitte eine ZIP-Datei auswählen.', true);
        return;
      }

      showMsg(msgEl, 'Wird hochgeladen…', false);
      updateBtn.disabled = true;

      const fd = new FormData();
      fd.append('zip_file', fileInput.files[0]);

      try {
        const res = await fetch('/admin/update', {
          method: 'POST',
          headers: { Authorization: basicAuthHeader(creds) },
          body: fd,
        });
        if (res.status === 401) { clearCredentials(); showAdminLogin(); return; }
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          showMsg(msgEl, 'Fehler: ' + (err.detail || res.statusText), true);
          updateBtn.disabled = false;
          return;
        }
        const data = await res.json();
        showMsg(
          msgEl,
          'Update eingespielt (' + data.files_updated + ' Dateien). ' +
          'Seite wird in 4 Sekunden neu geladen…',
          false
        );
        fileInput.value = '';
        setTimeout(() => window.location.reload(), 4000);
      } catch {
        showMsg(msgEl, 'Verbindungsfehler beim Upload.', true);
        updateBtn.disabled = false;
      }
    });
  });
})();
