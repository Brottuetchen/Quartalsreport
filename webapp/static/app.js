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
  const resetBtn = document.getElementById('reset-btn');

  const radioQuarter = document.getElementById('radio-quarter');
  const radioCustom = document.getElementById('radio-custom');
  const quarterInput = document.getElementById('quarter-input');
  const customInput = document.getElementById('custom-input');
  const quarterField = document.getElementById('quarter');
  const customPeriodField = document.getElementById('custom-period');

  let currentJobId = null;
  let pollTimer = null;

  // Toggle between quarter and custom period inputs
  function toggleInputs() {
    if (radioQuarter.checked) {
      quarterInput.style.display = 'block';
      customInput.style.display = 'none';
      customPeriodField.value = '';
    } else {
      quarterInput.style.display = 'none';
      customInput.style.display = 'block';
      quarterField.value = '';
    }
  }

  radioQuarter.addEventListener('change', toggleInputs);
  radioCustom.addEventListener('change', toggleInputs);

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

    // Prepare form data with the correct period value
    const formData = new FormData();
    formData.append('csv_file', csvFile);
    formData.append('xml_file', xmlFile);

    // Add the appropriate period value based on selection
    if (radioCustom.checked) {
      const customValue = customPeriodField.value.trim();
      if (customValue) {
        formData.append('quarter', customValue);
      }
    } else {
      const quarterValue = quarterField.value.trim();
      if (quarterValue) {
        formData.append('quarter', quarterValue);
      }
    }

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

  resetBtn.addEventListener('click', () => {
    resetUI();
  });

  window.addEventListener('beforeunload', () => {
    if (pollTimer) {
      clearInterval(pollTimer);
    }
  });
})();

