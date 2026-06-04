Office.onReady(() => {
  const settings = Office.context.roamingSettings;

  const enabledToggle = document.getElementById('enabled-toggle');
  const enabledStatus = document.getElementById('enabled-status');
  const sendNowToggle = document.getElementById('send-now-toggle');
  const sendNowLabel = document.getElementById('send-now-label');
  const toggleStatus = document.getElementById('toggle-status');
  const delayInput = document.getElementById('delay-seconds');
  const saveBtn = document.getElementById('save-btn');
  const saveStatus = document.getElementById('save-status');

  // Load saved values (default enabled = true)
  const isEnabled = settings.get('delaySendEnabled') !== false;
  const savedDelay = settings.get('delaySeconds') || 120;

  enabledToggle.checked = isEnabled;
  sendNowToggle.checked = !!settings.get('sendImmediately');
  delayInput.value = savedDelay;

  applyEnabledState(isEnabled);

  // Master enable/disable toggle
  enabledToggle.addEventListener('change', () => {
    settings.set('delaySendEnabled', enabledToggle.checked);
    settings.saveAsync(() => {
      enabledStatus.textContent = enabledToggle.checked
        ? 'Delay Send is active.'
        : 'Delay Send is disabled — emails will send immediately.';
      enabledStatus.style.color = enabledToggle.checked ? '#107c10' : '#605e5c';
      applyEnabledState(enabledToggle.checked);
    });
  });

  // Send Now toggle
  sendNowToggle.addEventListener('change', () => {
    settings.set('sendImmediately', sendNowToggle.checked);
    settings.saveAsync(() => {
      toggleStatus.textContent = sendNowToggle.checked
        ? 'Active — next send will bypass the delay.'
        : 'Off — delay will apply normally.';
      toggleStatus.style.color = sendNowToggle.checked ? '#107c10' : '#605e5c';
    });
  });

  // Show Save only when delay value changes
  delayInput.addEventListener('input', () => {
    const changed = parseInt(delayInput.value, 10) !== savedDelay;
    saveBtn.style.display = changed ? 'inline-block' : 'none';
    if (!changed) saveStatus.textContent = '';
  });

  // Save delay
  saveBtn.addEventListener('click', () => {
    const val = parseInt(delayInput.value, 10);
    if (isNaN(val) || val < 1) {
      saveStatus.textContent = 'Please enter a valid number of seconds (minimum 1).';
      saveStatus.style.color = '#d13438';
      return;
    }
    settings.set('delaySeconds', val);
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        saveStatus.textContent = 'Failed to save. Please try again.';
        saveStatus.style.color = '#d13438';
      } else {
        const mins = (val / 60).toFixed(1).replace(/\.0$/, '');
        saveStatus.textContent = `Saved — delay set to ${val}s (${mins} min).`;
        saveStatus.style.color = '#107c10';
        saveBtn.style.display = 'none';
        delayInput.dataset.saved = val;
      }
    });
  });

  function applyEnabledState(enabled) {
    sendNowToggle.disabled = !enabled;
    sendNowLabel.classList.toggle('disabled-label', !enabled);
    delayInput.disabled = !enabled;
    if (!enabled) saveBtn.style.display = 'none';
  }
});
