// Extended app.js with flexible report support

(() => {
  // Tab switching
  const tabBtns = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetTab = btn.dataset.tab;

      // Update active tab button
      tabBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // Show/hide content
      tabContents.forEach(content => {
        if (content.id === `${targetTab}-form`) {
          content.classList.add('active');
        } else {
          content.classList.remove('active');
        }
      });
    });
  });

  // Filter toggles
  const filterProjectsCheckbox = document.getElementById('filter-projects');
  const projectFilterSection = document.getElementById('project-filter-section');
  const filterEmployeesCheckbox = document.getElementById('filter-employees');
  const employeeFilterSection = document.getElementById('employee-filter-section');

  if (filterProjectsCheckbox) {
    filterProjectsCheckbox.addEventListener('change', (e) => {
      if (e.target.checked) {
        projectFilterSection.classList.remove('hidden');
      } else {
        projectFilterSection.classList.add('hidden');
        document.getElementById('projects').value = '';
      }
    });
  }

  if (filterEmployeesCheckbox) {
    filterEmployeesCheckbox.addEventListener('change', (e) => {
      if (e.target.checked) {
        employeeFilterSection.classList.remove('hidden');
      } else {
        employeeFilterSection.classList.add('hidden');
        document.getElementById('employees').value = '';
      }
    });
  }

  // Set default dates (today - 1 month to today)
  const startDateInput = document.getElementById('start-date');
  const endDateInput = document.getElementById('end-date');

  if (startDateInput && endDateInput) {
    const today = new Date();
    const oneMonthAgo = new Date(today);
    oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);

    startDateInput.valueAsDate = oneMonthAgo;
    endDateInput.valueAsDate = today;
  }

  // Standard form handler (existing behavior)
  const formStandard = document.getElementById('upload-form-standard');
  const submitBtnStandard = document.getElementById('submit-btn-standard');

  if (formStandard) {
    formStandard.addEventListener('submit', async (e) => {
      e.preventDefault();
      await handleStandardFormSubmit(formStandard, submitBtnStandard);
    });
  }

  // Flexible form handler (new)
  const formFlexible = document.getElementById('upload-form-flexible');
  const submitBtnFlexible = document.getElementById('submit-btn-flexible');

  if (formFlexible) {
    formFlexible.addEventListener('submit', async (e) => {
      e.preventDefault();
      await handleFlexibleFormSubmit(formFlexible, submitBtnFlexible);
    });
  }

  // Status card elements
  const statusCard = document.getElementById('status-card');
  const statusMessage = document.getElementById('status-message');
  const progressBar = document.getElementById('progress-bar');
  const queuePosition = document.getElementById('queue-position');
  const errorBox = document.getElementById('error-box');
  const downloadBox = document.getElementById('download-box');
  const downloadLink = document.getElementById('download-link');
  const resetBtn = document.getElementById('reset-btn');

  let currentJobId = null;
  let pollTimer = null;

  function showStatusCard() {
    statusCard.classList.remove('hidden');
  }

  function hideStatusCard() {
    statusCard.classList.add('hidden');
  }

  function setLoading(btn, isLoading) {
    btn.disabled = isLoading;
  }

  function resetUI() {
    clearInterval(pollTimer);
    pollTimer = null;
    currentJobId = null;
    if (progressBar) progressBar.style.width = '0%';
    if (statusMessage) statusMessage.textContent = 'Warte auf Start...';
    if (queuePosition) queuePosition.textContent = '-';
    if (errorBox) {
      errorBox.classList.add('hidden');
      errorBox.textContent = '';
    }
    if (downloadBox) downloadBox.classList.add('hidden');
    hideStatusCard();
    if (submitBtnStandard) setLoading(submitBtnStandard, false);
    if (submitBtnFlexible) setLoading(submitBtnFlexible, false);
  }

  async function handleStandardFormSubmit(form, btn) {
    setLoading(btn, true);
    showStatusCard();

    const formData = new FormData(form);

    try {
      const res = await fetch('/api/jobs', {
        method: 'POST',
        body: formData,
      });

      if (!res.ok) {
        const error = await res.json();
        throw new Error(error.detail || 'Upload fehlgeschlagen');
      }

      const data = await res.json();
      currentJobId = data.job_id;

      // Start polling
      pollTimer = setInterval(pollStatus, 1000);
    } catch (err) {
      errorBox.textContent = err.message;
      errorBox.classList.remove('hidden');
      setLoading(btn, false);
    }
  }

  async function handleFlexibleFormSubmit(form, btn) {
    setLoading(btn, true);
    showStatusCard();

    if (statusMessage) statusMessage.textContent = 'Bereite flexiblen Report vor...';
    if (progressBar) progressBar.style.width = '10%';

    const formData = new FormData(form);

    // Only include filter fields if checkboxes are checked
    if (!filterProjectsCheckbox.checked) {
      formData.delete('projects');
    }
    if (!filterEmployeesCheckbox.checked) {
      formData.delete('employees');
    }

    try {
      const res = await fetch('/api/reports/flexible', {
        method: 'POST',
        body: formData,
      });

      if (!res.ok) {
        const error = await res.json();
        throw new Error(error.detail || 'Report-Generierung fehlgeschlagen');
      }

      if (progressBar) progressBar.style.width = '100%';
      if (statusMessage) statusMessage.textContent = 'Fertig!';

      // Download the file
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const filename = res.headers.get('content-disposition')
        ?.split('filename=')[1]
        ?.replace(/"/g, '')
        || 'report.xlsm';

      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      // Show success
      if (downloadBox) {
        downloadLink.textContent = `✓ ${filename} wurde heruntergeladen`;
        downloadLink.href = '#';
        downloadLink.onclick = (e) => e.preventDefault();
        downloadBox.classList.remove('hidden');
      }

      setLoading(btn, false);
    } catch (err) {
      if (errorBox) {
        errorBox.textContent = err.message;
        errorBox.classList.remove('hidden');
      }
      if (progressBar) progressBar.style.width = '0%';
      setLoading(btn, false);
    }
  }

  async function pollStatus() {
    if (!currentJobId) return;
    try {
      const res = await fetch(`/api/jobs/${currentJobId}`);
      if (!res.ok) {
        throw new Error('Status konnte nicht abgefragt werden');
      }
      const data = await res.json();

      statusMessage.textContent = data.message;
      progressBar.style.width = `${data.progress}%`;

      if (data.queue_position !== null && data.queue_position > 0) {
        queuePosition.textContent = data.queue_position;
      } else {
        queuePosition.textContent = 'In Bearbeitung';
      }

      if (data.status === 'finished' && data.download_available) {
        clearInterval(pollTimer);
        pollTimer = null;
        downloadLink.href = `/api/jobs/${currentJobId}/download`;
        downloadBox.classList.remove('hidden');
        setLoading(submitBtnStandard, false);
      }

      if (data.status === 'failed') {
        clearInterval(pollTimer);
        pollTimer = null;
        errorBox.textContent = data.error || 'Ein Fehler ist aufgetreten';
        errorBox.classList.remove('hidden');
        setLoading(submitBtnStandard, false);
      }
    } catch (err) {
      clearInterval(pollTimer);
      pollTimer = null;
      errorBox.textContent = err.message;
      errorBox.classList.remove('hidden');
      setLoading(submitBtnStandard, false);
    }
  }

  if (resetBtn) {
    resetBtn.addEventListener('click', resetUI);
  }

  // Theme toggle
  const themeToggle = document.getElementById('theme-toggle');
  const themeIcon = document.getElementById('theme-icon');
  const docElement = document.documentElement;

  // Function to apply theme
  const applyTheme = (theme) => {
    docElement.dataset.theme = theme;
    localStorage.setItem('theme', theme);
    if (themeIcon) {
      themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
    }
  };

  // Initial theme check
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const savedTheme = localStorage.getItem('theme');
  
  applyTheme(savedTheme || (prefersDark ? 'dark' : 'light'));

  // Event listener for the toggle button
  if (themeToggle) {
    themeToggle.addEventListener('click', () => {
      const newTheme = docElement.dataset.theme === 'dark' ? 'light' : 'dark';
      applyTheme(newTheme);
    });
  }
})();
