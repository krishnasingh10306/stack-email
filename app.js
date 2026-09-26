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
function refreshSend() { sendButton.disabled = !(contacts.length && $('#subject').value.trim() && $('#body').value.trim()); }

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

sendButton.addEventListener('click', () => {
  $('#confirm-copy').textContent = `You’re about to send “${$('#subject').value.trim()}” to ${contacts.length} ${contacts.length === 1 ? 'person' : 'people'}.`;
  dialog.showModal();
});
$('#cancel-send').addEventListener('click', () => dialog.close());
$('#confirm-send').addEventListener('click', async () => {
  const button = $('#confirm-send');
  button.disabled = true; button.textContent = 'Sending…';
  const payload = {
    contacts, subject: $('#subject').value.trim(), body: $('#body').value,
    host: $('#smtp-host').value, port: 587,
    sender: $('#smtp-email').value, username: $('#smtp-email').value, password: $('#smtp-password').value,
  };
  try {
    const response = await fetch('/api/send', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Email sending failed.');
    const failed = result.results.length - result.sent;
    toast(`${result.sent} sent${failed ? ` · ${failed} failed` : ' successfully'}.`);
    if (!failed) dialog.close();
    if (failed) toast(`${result.sent} sent · ${failed} failed. Check SMTP and try the failed addresses again.`);
  } catch (error) { toast(error.message); }
  finally { button.disabled = false; button.innerHTML = 'Send emails <span class="send-arrow">↗</span>'; }
});

$('#history-link').addEventListener('click', () => toast('Email history is not stored by this private local app.'));
$('#char-count').textContent = `${$('#body').value.length} characters`;
