const $ = (selector) => document.querySelector(selector);
let contacts = [];
const fileInput = $('#file-input');
const dropZone = $('#drop-zone');
const panel = $('#contacts-panel');
const sendButton = $('#send-button');
const dialog = $('#confirm-dialog');

function toast(message) {
  const element = $('#toast');
  element.textContent = message;
  element.classList.add('show');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => element.classList.remove('show'), 3300);
}

async function loadFile(file) {
  if (!file) return;
  const encoded = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
  try {
    const response = await fetch('/api/contacts', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ filename: file.name, data: encoded }) });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Could not read this file.');
    contacts = result.contacts;
    renderContacts();
    toast(`Found ${contacts.length} valid email${contacts.length === 1 ? '' : 's'}.`);
  } catch (error) {
    contacts = [];
    panel.classList.add('hidden');
    refreshSend();
    toast(error.message || 'Could not read this file.');
  }
}

function renderContacts() {
  if (!contacts.length) { panel.classList.add('hidden'); refreshSend(); return; }
  const headers = Object.keys(contacts[0]);
  const shown = contacts.slice(0, 5);
  panel.innerHTML = `<div class="contacts-summary"><strong>${contacts.length} people</strong> ready to receive a personal email</div><div class="table-scroll"><table><thead><tr>${headers.slice(0, 4).map((h) => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${shown.map((person) => `<tr>${headers.slice(0, 4).map((h) => `<td>${escapeHtml(person[h])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  panel.classList.remove('hidden');
  refreshSend();
}

function escapeHtml(value) { return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]); }
function isScheduled() { return $('input[name="send-timing"]:checked').value === 'later'; }
function refreshSend() {
  const scheduleValid = !isScheduled() || ($('#schedule-datetime').value && new Date($('#schedule-datetime').value).getTime() > Date.now());
  sendButton.disabled = !(contacts.length && $('#subject').value.trim() && $('#body').value.trim() && scheduleValid);
  sendButton.querySelector('span:first-child').textContent = isScheduled() ? 'Review schedule' : 'Review & send';
}

fileInput.addEventListener('change', () => loadFile(fileInput.files[0]));
for (const eventName of ['dragenter', 'dragover']) dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.add('dragging'); });
for (const eventName of ['dragleave', 'drop']) dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.remove('dragging'); });
dropZone.addEventListener('drop', (event) => loadFile(event.dataTransfer.files[0]));

$('#template-button').addEventListener('click', () => {
  const blob = new Blob(['First name,Last name,Email,Department\nAlex,Green,alex@example.com,Design\nSam,Lee,sam@example.com,People\n'], { type: 'text/csv' });
  const anchor = document.createElement('a'); anchor.href = URL.createObjectURL(blob); anchor.download = 'letterdrop-sample.csv'; anchor.click(); URL.revokeObjectURL(anchor.href);
});

$('#setup-toggle').addEventListener('click', () => {
  const opened = $('#smtp-panel').classList.toggle('hidden');
  $('#setup-toggle').innerHTML = opened ? 'Set up SMTP <span>⌄</span>' : 'Hide setup <span>⌃</span>';
});
document.querySelectorAll('#subject, #body').forEach((element) => element.addEventListener('input', () => { $('#char-count').textContent = `${$('#body').value.length} characters`; refreshSend(); }));
document.querySelectorAll('.token-button').forEach((button) => button.addEventListener('click', () => {
  const textarea = $('#body'); const token = button.dataset.token; const start = textarea.selectionStart; const end = textarea.selectionEnd;
  textarea.setRangeText(token, start, end, 'end'); textarea.focus(); textarea.dispatchEvent(new Event('input'));
}));

function setDateMinimum() {
  const date = new Date(Date.now() + 120000);
  date.setSeconds(0, 0);
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
  $('#schedule-datetime').min = local;
}
const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'your local time';
$('#timezone-note').textContent = `Time is based on this device’s time zone (${timeZone}).`;
setDateMinimum();
document.querySelectorAll('input[name="send-timing"]').forEach((radio) => radio.addEventListener('change', () => {
  document.querySelectorAll('.timing-option').forEach((option) => option.classList.toggle('selected', option.querySelector('input').checked));
  $('#scheduled-time-fields').classList.toggle('hidden', !isScheduled());
  refreshSend();
}));
$('#schedule-datetime').addEventListener('input', refreshSend);

sendButton.addEventListener('click', () => {
  const scheduled = isScheduled();
  const when = scheduled ? new Date($('#schedule-datetime').value).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' }) : '';
  $('#confirm-title').textContent = scheduled ? 'Schedule this note?' : 'Ready to send?';
  $('#confirm-copy').textContent = scheduled
    ? `“${$('#subject').value.trim()}” will be sent to ${contacts.length} ${contacts.length === 1 ? 'person' : 'people'} on ${when} (${timeZone}).`
    : `You’re about to send “${$('#subject').value.trim()}” to ${contacts.length} ${contacts.length === 1 ? 'person' : 'people'}.`;
  $('#confirm-warning').textContent = scheduled
    ? 'The recipient list and SMTP app password will be stored encrypted until sending. You can cancel this scheduled send beforehand.'
    : 'This sends one separate email to each person on your list. You can’t undo it.';
  $('#confirm-send').innerHTML = scheduled ? 'Schedule email <span class="send-arrow">◷</span>' : 'Send emails <span class="send-arrow">↗</span>';
  dialog.showModal();
});
$('#cancel-send').addEventListener('click', () => dialog.close());
$('#confirm-send').addEventListener('click', async () => {
  const button = $('#confirm-send');
  const scheduled = isScheduled();
  button.disabled = true; button.textContent = scheduled ? 'Scheduling…' : 'Sending…';
  const payload = {
    contacts, subject: $('#subject').value.trim(), body: $('#body').value,
    host: $('#smtp-host').value, port: 587,
    sender: $('#smtp-email').value, username: $('#smtp-email').value, password: $('#smtp-password').value,
  };
  if (scheduled) payload.run_at = new Date($('#schedule-datetime').value).toISOString();
  try {
    const response = await fetch(scheduled ? '/api/schedule' : '/api/send', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Email sending failed.');
    dialog.close();
    if (scheduled) {
      $('#smtp-password').value = '';
      toast(`Scheduled for ${new Date(result.run_at).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' })}.`);
      await loadSchedules();
    } else {
      const failed = result.results.length - result.sent;
      toast(`${result.sent} sent${failed ? ` · ${failed} failed` : ' successfully'}.`);
      if (failed) toast(`${result.sent} sent · ${failed} failed. Check SMTP and try the failed addresses again.`);
    }
  } catch (error) { toast(error.message); }
  finally { button.disabled = false; button.innerHTML = isScheduled() ? 'Schedule email <span class="send-arrow">◷</span>' : 'Send emails <span class="send-arrow">↗</span>'; }
});

function statusLabel(status) { return ({ scheduled: 'Scheduled', sending: 'Sending now', sent: 'Sent', partial: 'Partially sent', failed: 'Needs attention', cancelled: 'Cancelled' })[status] || status; }
async function loadSchedules() {
  try {
    const response = await fetch('/api/schedules');
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Could not load scheduled sends.');
    const schedules = result.schedules;
    $('#schedule-board').classList.toggle('hidden', schedules.length === 0);
    $('#schedule-list').innerHTML = schedules.map((schedule) => {
      const time = new Date(schedule.run_at).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
      const message = schedule.error ? `<small class="schedule-error">${escapeHtml(schedule.error)}</small>` : '';
      const cancel = schedule.status === 'scheduled' ? `<button class="cancel-schedule" data-id="${escapeHtml(schedule.id)}">Cancel</button>` : '';
      const icon = schedule.status === 'scheduled' ? '◷' : schedule.status === 'sent' ? '✓' : '·';
      return `<article class="schedule-item"><div class="schedule-clock">${icon}</div><div class="schedule-details"><strong>${escapeHtml(schedule.recipient_count)} ${schedule.recipient_count === 1 ? 'recipient' : 'recipients'}</strong><span>${escapeHtml(time)}</span>${message}</div><span class="schedule-status status-${escapeHtml(schedule.status)}">${escapeHtml(statusLabel(schedule.status))}</span>${cancel}</article>`;
    }).join('');
  } catch (error) { toast(error.message); }
}
$('#schedule-list').addEventListener('click', async (event) => {
  const button = event.target.closest('.cancel-schedule');
  if (!button) return;
  button.disabled = true;
  try {
    const response = await fetch(`/api/schedules/${encodeURIComponent(button.dataset.id)}/cancel`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Could not cancel this scheduled send.');
    toast('Scheduled send cancelled.');
    await loadSchedules();
  } catch (error) { toast(error.message); button.disabled = false; }
});
$('#refresh-schedules').addEventListener('click', loadSchedules);
$('#history-link').addEventListener('click', () => { loadSchedules(); $('#schedule-board').scrollIntoView({ behavior: 'smooth', block: 'center' }); });
$('#char-count').textContent = `${$('#body').value.length} characters`;
loadSchedules();
setInterval(loadSchedules, 30000);
