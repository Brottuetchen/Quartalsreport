(() => {
  const form = document.getElementById('upload-form');
  const submitBtn = document.getElementById('submit-btn');
  const statusCard = document.getElementById('status-card');
  const statusMessage = document.getElementById('status-message');
  const progressBar = document.getElementById('progress-bar');
  const queuePosition = document.getElementById('queue-position');
  const errorBox = document.getElementById('error-box');
  const downloadBox = document.getElementById('download-box');
  const downloadLink = document.getElementById('download-link');
  const exportPdfBtn = document.getElementById('export-pdf-btn');
  const pdfExportBox = document.getElementById('pdf-export-box');
  const pdfList = document.getElementById('pdf-list');
  const resetBtn = document.getElementById('reset-btn');

  let currentJobId = null;
  let pollTimer = null;

  function showStatusCard() {
    statusCard.classList.remove('hidden');
  }

  function hideStatusCard() {
    statusCard.classList.add('hidden');
  }

  function setLoading(isLoading) {
    submitBtn.disabled = isLoading;
  }

  function resetUI() {
    clearInterval(pollTimer);
    pollTimer = null;
    currentJobId = null;
    progressBar.style.width = '0%';
    statusMessage.textContent = 'Warte auf Start...';
    queuePosition.textContent = '-';
    errorBox.classList.add('hidden');
    errorBox.textContent = '';
    downloadBox.classList.add('hidden');
    pdfExportBox.classList.add('hidden');
    pdfList.innerHTML = '';
    hideStatusCard();
    setLoading(false);
    form.reset();
  }

  async function pollStatus() {
    if (!currentJobId) return;
    try {
      const res = await fetch(`/api/jobs/${currentJobId}`);
      if (!res.ok) {
        throw new Error('Status konnte nicht abgefragt werden');
      }
      const data = await res.json();
      statusMessage.textContent = data.message || '';
      progressBar.style.width = `${data.progress || 0}%`;
      queuePosition.textContent = data.queue_position !== null ? data.queue_position : '-';

      if (data.status === 'finished') {
        clearInterval(pollTimer);
        pollTimer = null;
        downloadLink.href = `/api/jobs/${currentJobId}/download`;
        downloadBox.classList.remove('hidden');
        statusMessage.textContent = 'Bereit zum Download';
      } else if (data.status === 'failed') {
        clearInterval(pollTimer);
        pollTimer = null;
        errorBox.textContent = data.error || 'Unbekannter Fehler';
        errorBox.classList.remove('hidden');
        setLoading(false);
      }
    } catch (err) {
      clearInterval(pollTimer);
      pollTimer = null;
      errorBox.textContent = err.message || String(err);
      errorBox.classList.remove('hidden');
      setLoading(false);
    }
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (pollTimer) {
      clearInterval(pollTimer);
      pollTimer = null;
    }

    const csvFile = document.getElementById('csv-file').files[0];
    const xmlFile = document.getElementById('xml-file').files[0];
    if (!csvFile || !xmlFile) {
      alert('Bitte CSV- und XML-Datei auswählen.');
      return;
    }

    const formData = new FormData(form);

    setLoading(true);
    showStatusCard();
    statusMessage.textContent = 'Upload läuft...';
    progressBar.style.width = '5%';
    queuePosition.textContent = '-';
    errorBox.classList.add('hidden');
    downloadBox.classList.add('hidden');

    try {
      const response = await fetch('/api/jobs', {
        method: 'POST',
        body: formData,
      });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.detail || 'Upload fehlgeschlagen');
      }
      const data = await response.json();
      currentJobId = data.job_id;
      statusMessage.textContent = data.message || 'In Warteschlange';
      queuePosition.textContent = data.queue_position !== null ? data.queue_position : '-';
      progressBar.style.width = '15%';

      pollTimer = setInterval(pollStatus, 2000);
      setLoading(true);
    } catch (err) {
      errorBox.textContent = err.message || String(err);
      errorBox.classList.remove('hidden');
      setLoading(false);
    }
  });

  exportPdfBtn.addEventListener('click', async () => {
    if (!currentJobId) {
      alert('Kein Job verfügbar.');
      return;
    }

    exportPdfBtn.disabled = true;
    exportPdfBtn.textContent = 'Exportiere...';
    pdfExportBox.classList.add('hidden');
    pdfList.innerHTML = '';

    try {
      const response = await fetch(`/api/jobs/${currentJobId}/export-pdf`, {
        method: 'POST',
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.detail || 'PDF-Export fehlgeschlagen');
      }

      const data = await response.json();

      // Display PDF list
      pdfExportBox.classList.remove('hidden');
      pdfList.innerHTML = '';

      if (data.pdfs && data.pdfs.length > 0) {
        const ul = document.createElement('ul');
        data.pdfs.forEach((pdfName) => {
          const li = document.createElement('li');
          const link = document.createElement('a');
          link.href = `/api/jobs/${currentJobId}/pdf/${encodeURIComponent(pdfName)}`;
          link.textContent = pdfName;
          link.download = pdfName;
          li.appendChild(link);
          ul.appendChild(li);
        });
        pdfList.appendChild(ul);
      }

      exportPdfBtn.textContent = 'Als PDFs exportieren';
      exportPdfBtn.disabled = false;
      alert(data.message || 'PDFs erfolgreich erstellt');
    } catch (err) {
      exportPdfBtn.textContent = 'Als PDFs exportieren';
      exportPdfBtn.disabled = false;
      errorBox.textContent = `PDF-Export: ${err.message || String(err)}`;
      errorBox.classList.remove('hidden');
    }
  });

  resetBtn.addEventListener('click', () => {
    resetUI();
  });

  window.addEventListener('beforeunload', () => {
    if (pollTimer) {
      clearInterval(pollTimer);
    }
  });
})();

