Office.onReady(() => {
  const settings = Office.context.roamingSettings;
  const toggle = document.getElementById('send-now-toggle');
  const toggleStatus = document.getElementById('toggle-status');
  const delayInput = document.getElementById('delay-seconds');
  const saveBtn = document.getElementById('save-btn');
  const saveStatus = document.getElementById('save-status');

  // Load current values
  toggle.checked = !!settings.get('sendImmediately');
  delayInput.value = settings.get('delaySeconds') || 120;

  // Send Now toggle
  toggle.addEventListener('change', () => {
    settings.set('sendImmediately', toggle.checked);
    settings.saveAsync(() => {
      toggleStatus.textContent = toggle.checked
        ? 'Active — next send will bypass the delay.'
        : 'Off — delay will apply normally.';
      toggleStatus.style.color = toggle.checked ? '#107c10' : '#605e5c';
    });
  });

  // Delay save
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
      }
    });
  });
});
