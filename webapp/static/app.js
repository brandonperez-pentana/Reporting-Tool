/* Reporter, in a browser.
 *
 * A thin client over /api, which is itself a thin wrapper over report_engine — the
 * same session the desktop app drives. No framework and no CDN on purpose: this has to
 * run on a managed Windows machine with no network egress and nothing installed.
 */

'use strict';

const state = {
  settings: {},
  months: { detected: [], all: [], default: '' },
  session: { has_session: false, summary: null, results: [], queue: [] },
  current: null,        // selected dealership on stage 1
  currentDetail: null,
  queueCurrent: null,   // selected dealership on stage 2
  smsLimit: 160,
  contacts: { path: '', fieldnames: [], rows: [], dealership_names: [] },
  outlookChecked: false,
  dirty: false,
  contactsDirty: false,
};

const $ = (id) => document.getElementById(id);
const el = (tag, cls, text) => {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  if (text !== undefined) node.textContent = text;
  return node;
};

/* ---------- transport ---------- */

async function api(path, options = {}) {
  const init = { method: options.method || 'GET', headers: {} };
  if (options.body !== undefined) {
    init.headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(options.body);
  }
  let response;
  try {
    response = await fetch(path, init);
  } catch (err) {
    throw new Error('Lost contact with the Reporter server. Is the window that started it still open?');
  }
  const text = await response.text();
  let payload = {};
  if (text) {
    try { payload = JSON.parse(text); } catch (err) { payload = { error: text.slice(0, 400) }; }
  }
  if (!response.ok) throw new Error(payload.error || `Request failed (${response.status}).`);
  return payload;
}

/* Poll a background job (reporting run, send batch, email build) to completion. */
async function runJob(startPath, body, onProgress) {
  const { job_id: jobId } = await api(startPath, { method: 'POST', body });
  for (;;) {
    await new Promise((resolve) => setTimeout(resolve, 600));
    const job = await api(`/api/jobs/${jobId}`);
    if (onProgress) onProgress(job);
    if (job.state === 'done') return job.result;
    if (job.state === 'failed') throw new Error(job.error || 'That job failed.');
  }
}

/* ---------- chrome ---------- */

function status(message, tone = 'info') {
  const bar = $('status-bar');
  $('status-text').textContent = message;
  bar.className = 'status' + (tone === 'good' ? ' is-good' : tone === 'bad' ? ' is-bad' : '');
  bar.hidden = !message;
}

function busy(on, text) {
  $('busy-text').textContent = text || 'Working…';
  $('busy').hidden = !on;
}

async function guard(fn, label) {
  try {
    return await fn();
  } catch (err) {
    status(err.message || String(err), 'bad');
    if (label) console.error(label, err);
    return null;
  } finally {
    busy(false);
  }
}

function openModal(id) { $(id).hidden = false; }
function closeModal(id) {
  if (id === 'contacts-modal' && state.contactsDirty
      && !window.confirm('Your contact changes have not been saved. Close without saving?')) return;
  if (id === 'contacts-modal') state.contactsDirty = false;
  $(id).hidden = true;
}

document.addEventListener('click', (event) => {
  if (event.target.matches('[data-close]')) {
    const modal = event.target.closest('.modal');
    if (modal) closeModal(modal.id);
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') document.querySelectorAll('.modal:not([hidden])').forEach((m) => closeModal(m.id));
});

document.querySelectorAll('.stage-tab').forEach((tab) => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.stage-tab').forEach((other) => other.classList.toggle('is-active', other === tab));
    const target = tab.dataset.stage;
    document.querySelectorAll('.stage').forEach((stage) => stage.classList.toggle('is-active', stage.id === `stage-${target}`));
    if (target === 'queue') { renderQueue(); checkOutlook(); }
    if (target === 'sms') loadSms();
  });
});

/* ---------- insight bullets (mirrors reporter_ui) ---------- */

const BULLET = /^\s*[•\-*]\s*/;
const bulletise = (text) => (text || '').split('\n').map((l) => l.replace(BULLET, '').trim()).filter(Boolean).map((l) => `• ${l}`).join('\n');
const stripBullets = (text) => (text || '').split('\n').map((l) => l.replace(BULLET, '').trimEnd()).filter((l) => l.trim()).join('\n');

/* ---------- bootstrap ---------- */

async function bootstrap() {
  const data = await guard(() => api('/api/bootstrap'), 'bootstrap');
  if (!data) return;
  state.settings = data.settings;
  state.months = data.months;
  state.session = data.session;
  state.smsLimit = data.sms_limit || 160;
  applyServerMode();
  renderMonths();
  renderPortfolioButton();
  renderSession();
  const missing = state.settings.missing_required || [];
  if (state.settings.first_run) {
    // Never configured on this machine: walk them through it rather than dropping
    // them into the full settings sheet with nine empty boxes.
    openSetup();
  } else if (missing.length && !canEdit(missing[0])) {
    status('Reporter is not fully set up on the server yet. Ask an administrator.', 'bad');
  } else if (missing.length) {
    status('Set the paths in Settings before running a month.', 'info');
    openModal('settings-modal');
    renderSettings();
  }
}

/* ---------- server mode ----------
 *
 * On the server, Outlook and Word belong to the server, not the person at the browser:
 * drafts and "open in app" are hidden and .eml files come back as a download.
 */

const serverMode = () => !!state.settings.server_mode;
const canEdit = (key) => !serverMode() || state.settings.is_admin || !(state.settings.shared_keys || []).includes(key);

function applyServerMode() {
  const on = serverMode();
  $('signed-in').hidden = !on;
  $('signed-in').textContent = on ? `Signed in as ${state.settings.user}${state.settings.is_admin ? ' (admin)' : ''}` : '';
  $('save-drafts').hidden = on;
  $('open-pdf').hidden = on;
  $('save-eml').textContent = on ? 'Download .eml' : 'Save as .eml';
  if (on) $('outlook-warn').hidden = true;
}

/* ---------- first run ---------- */

const SETUP_FIELDS = [
  ['accounts_folder', 'Accounts folder', 'folder', '', true,
   'The master workbooks, one per dealership. Their Sales and Service PDF folders can sit here too, or be set separately in Settings.'],
  ['output_folder', 'Output folder', 'folder', '', true,
   'Where the generated reports and insights are written.'],
  ['recipient_mapping_path', 'Contacts file', 'file', '.csv', true,
   'Who each dealership’s email goes to. Leave the suggested path and an empty contacts.csv is created there for you.'],
  ['eml_export_folder', 'Save .eml files to', 'folder', '', false,
   'Optional. Where “Save as .eml” writes emails for importing into Outlook by hand.'],
];

function openSetup() {
  const wrap = $('setup-fields');
  wrap.textContent = '';
  const values = {};
  SETUP_FIELDS.forEach(([key, label, kind, extensions, required, description]) => {
    // data/contacts.csv is the sensible starting point, so offer it rather than
    // making a new operator find a file they have never seen. It is created on demand
    // rather than shipped, so reinstalling over a copy in use cannot overwrite real
    // contacts with an empty file.
    const prefill = state.settings[key] || (key === 'recipient_mapping_path' ? (state.settings.default_contacts_path || '') : '');
    values[key] = prefill;
    const row = el('div', 'setting-row');
    const labelNode = el('label', null, required ? label : `${label} (optional)`);
    labelNode.setAttribute('for', `setup-${key}`);
    row.append(labelNode);
    row.append(el('p', 'desc', description));
    const line = el('div', 'with-check');
    const input = el('input');
    input.type = 'text';
    input.id = `setup-${key}`;
    input.value = prefill;
    input.spellcheck = false;
    input.addEventListener('input', () => { values[key] = input.value.trim(); updateSetupProgress(values); });
    line.append(input);
    const browse = el('button', 'btn btn-small', 'Browse…');
    browse.type = 'button';
    browse.addEventListener('click', () => openBrowse({
      title: `Choose the ${label.toLowerCase()}`,
      kind,
      extensions,
      start: input.value,
      onPick: (chosen) => { input.value = chosen; values[key] = chosen; updateSetupProgress(values); },
    }));
    line.append(browse);
    row.append(line);
    wrap.append(row);
  });
  setupValues = values;
  updateSetupProgress(values);
  openModal('setup-modal');
}

let setupValues = {};

function updateSetupProgress(values) {
  const required = SETUP_FIELDS.filter(([, , , , req]) => req).map(([key]) => key);
  const done = required.filter((key) => (values[key] || '').trim()).length;
  $('setup-progress').textContent = `${done} of ${required.length} required`;
  $('setup-save').disabled = done < required.length;
}

$('setup-save').addEventListener('click', () => guard(async () => {
  const payload = { ...setupValues };
  busy(true, 'Saving setup…');
  state.settings = await api('/api/settings', { method: 'POST', body: payload });
  busy(false);
  closeModal('setup-modal');
  const months = await api('/api/months');
  state.months = months;
  renderMonths();
  status('Setup saved. Pick a month and run reporting.', 'good');
}));

function renderMonths() {
  const select = $('month-select');
  select.textContent = '';
  const seen = new Set();
  const push = (label, group) => {
    if (!label || seen.has(label)) return;
    seen.add(label);
    const option = el('option', null, group ? `${label}  (${group})` : label);
    option.value = label;
    select.append(option);
  };
  state.months.detected.forEach((m) => push(m, 'found in accounts'));
  state.months.all.forEach((m) => push(m));
  select.value = state.months.default || select.options[0]?.value || '';
}

function renderSession() {
  const { has_session: has, summary, results } = state.session;
  $('session-line').textContent = has
    ? `${summary.month} · ${results.length} dealerships · ${summary.accounts_folder}`
    : 'No reporting run loaded';
  $('dealership-count').textContent = String(results.length);
  renderDealershipList();
  renderQueue();
}

/* ---------- stage 1 ---------- */

function renderDealershipList() {
  const list = $('dealership-list');
  list.textContent = '';
  state.session.results.forEach((row) => {
    const item = el('li');
    item.dataset.name = row.dealership;
    if (row.dealership === state.current) item.classList.add('is-selected');
    // Approval is the operator's call; the only automatic state is "blocked" (no real figures to send).
    let pip = 'pip';
    if (row.skipped) pip += ' skipped';
    else if (row.approved) pip += ' approved';
    else if (!row.approvable) pip += ' blocked';
    item.append(el('span', pip));
    item.append(el('span', 'name', row.dealership));
    item.addEventListener('click', () => selectDealership(row.dealership));
    list.append(item);
  });
}

async function selectDealership(name) {
  if (state.dirty && !confirm('Discard unsaved edits for the current dealership?')) return;
  state.current = name;
  state.dirty = false;
  renderDealershipList();
  busy(true, `Loading ${name}…`);
  const detail = await guard(() => api(`/api/dealerships/${encodeURIComponent(name)}`), 'dealership');
  if (!detail) return;
  state.currentDetail = detail;
  $('current-dealership').textContent = `${detail.dealership} — ${detail.month}`;
  $('insights-box').value = bulletise(detail.priority_insights);
  $('summary-box').value = detail.summary || '';
  $('sms-box').value = detail.sms_version || '';
  updateSmsCount();
  renderFlags(detail);
  $('approve-blocked').hidden = !!detail.can_approve;
  $('approve-blocked').textContent = detail.approve_blocked_reason || '';
  $('approve-next').disabled = !detail.can_approve;
  loadPdf();
  status('');
}

function renderFlags(detail) {
  const flags = $('current-flags');
  flags.textContent = '';
  const add = (text, tone) => flags.append(el('span', `flag ${tone || ''}`.trim(), text));
  if (detail.approved) add('approved', 'good');
  if (detail.skipped) add('skipped');
  if (detail.blocked) add('blocked', 'bad');
  add(detail.contact_ready ? 'contact ok' : 'no contact', detail.contact_ready ? 'good' : 'bad');
}

function updateSmsCount() {
  const length = $('sms-box').value.length;
  const over = length > state.smsLimit;
  $('sms-count').textContent = `${length}/${state.smsLimit}${over ? ' — over limit' : ''}`;
  $('sms-count').style.color = over ? 'var(--bad)' : '';
}

['insights-box', 'summary-box', 'sms-box'].forEach((id) => {
  $(id).addEventListener('input', () => { state.dirty = true; if (id === 'sms-box') updateSmsCount(); });
});

async function saveEdits({ silent = false } = {}) {
  if (!state.current) return false;
  const payload = {
    priority_insights: stripBullets($('insights-box').value),
    summary: $('summary-box').value,
    sms_version: $('sms-box').value,
  };
  busy(true, 'Saving edits…');
  const detail = await guard(() => api(`/api/dealerships/${encodeURIComponent(state.current)}/text`, { method: 'POST', body: payload }), 'save');
  if (!detail) return false;
  state.currentDetail = detail;
  state.dirty = false;
  $('insights-box').value = bulletise(detail.priority_insights);
  await refreshSession();
  if (!silent) status(`Saved edits for ${detail.dealership}.`, 'good');
  return true;
}

async function refreshSession() {
  const session = await guard(() => api('/api/session'), 'session');
  if (session) { state.session = session; renderSession(); }
}

$('save-edits').addEventListener('click', () => saveEdits());

$('skip-dealership').addEventListener('click', () => guard(async () => {
  if (!state.current) return;
  busy(true, 'Skipping…');
  await api(`/api/dealerships/${encodeURIComponent(state.current)}/skip`, { method: 'POST', body: { skipped: true } });
  await refreshSession();
  status(`${state.current} skipped.`, 'good');
  advanceToNextUnapproved();
}));

$('approve-next').addEventListener('click', () => guard(async () => {
  if (!state.current) return;
  if (state.dirty && !(await saveEdits({ silent: true }))) return;
  const name = state.current;
  busy(true, `Approving ${name}…`);
  await api(`/api/dealerships/${encodeURIComponent(name)}/approve`, { method: 'POST', body: { approved: true } });
  await refreshSession();
  status(`${name} approved. Building its email…`, 'good');
  advanceToNextUnapproved();
  // The desktop app builds the email in the background on approval; same here.
  runJob(`/api/dealerships/${encodeURIComponent(name)}/email/build`, {})
    .then(() => { refreshSession(); status(`${name} approved and its email is ready.`, 'good'); })
    .catch((err) => status(`${name} approved, but the email could not be built: ${err.message}`, 'bad'));
}));

function advanceToNextUnapproved() {
  const rows = state.session.results;
  const index = rows.findIndex((row) => row.dealership === state.current);
  const ordered = [...rows.slice(index + 1), ...rows.slice(0, Math.max(index, 0))];
  const next = ordered.find((row) => !row.approved && !row.skipped);
  if (next) selectDealership(next.dealership);
}

$('run-reporting').addEventListener('click', () => guard(async () => {
  const month = $('month-select').value;
  if (!month) { status('Choose a month first.', 'bad'); return; }
  $('run-reporting').disabled = true;
  $('run-progress').hidden = false;
  status('');
  try {
    const result = await runJob('/api/reporting/run', { month, all_dealerships: $('run-all').checked }, (job) => {
      const pct = job.total ? Math.round(((job.processed || 0) / job.total) * 100) : 0;
      $('run-progress-fill').style.width = `${pct}%`;
      $('run-progress-text').textContent = job.message || '';
    });
    state.session = result;
    state.current = null;
    renderSession();
    status(`Reporting finished: ${result.results.length} dealerships for ${result.summary.month}.`, 'good');
    if (result.results.length) selectDealership(result.results[0].dealership);
  } finally {
    $('run-reporting').disabled = false;
    $('run-progress').hidden = true;
    $('run-progress-fill').style.width = '0%';
  }
}));

/* ---------- my dealerships ---------- */

// Keyed case-insensitively, as the server matches them, but the filename is kept as written.
const portfolioState = { workbooks: [], total: 0, missing: new Set(), selected: new Map() };
const portfolioKey = (name) => name.toLowerCase();

function renderPortfolioButton() {
  const count = (state.settings.portfolio || []).length;
  $('open-portfolio').textContent = count ? `My dealerships (${count})` : 'My dealerships: all';
  $('run-all-wrap').hidden = !count;
  if (!count) $('run-all').checked = false;
}

function shownWorkbooks() {
  const needle = $('portfolio-filter').value.trim().toLowerCase();
  return portfolioState.workbooks.filter((name) => !needle || name.toLowerCase().includes(needle));
}

function updatePortfolioCount() {
  const ticked = portfolioState.selected.size;
  $('portfolio-count').textContent = ticked
    ? `${ticked} of ${portfolioState.total} ticked`
    : `None ticked — every run reads all ${portfolioState.total}`;
}

function renderPortfolio() {
  const list = $('portfolio-list');
  list.textContent = '';
  const shown = shownWorkbooks();
  if (!shown.length) {
    list.append(el('li', 'browse-empty', portfolioState.workbooks.length ? 'Nothing matches that filter.' : 'No workbooks in the accounts folder.'));
  }
  shown.forEach((name) => {
    const key = portfolioKey(name);
    const item = el('li');
    const label = el('label');
    const box = el('input');
    box.type = 'checkbox';
    box.checked = portfolioState.selected.has(key);
    box.addEventListener('change', () => {
      if (box.checked) portfolioState.selected.set(key, name);
      else portfolioState.selected.delete(key);
      updatePortfolioCount();
    });
    label.append(box, el('span', 'name', name.replace(/\.xls[xm]$/i, '')));
    if (portfolioState.missing.has(key)) label.append(el('span', 'flag bad', 'not found'));
    item.append(label);
    list.append(item);
  });
  updatePortfolioCount();
}

$('open-portfolio').addEventListener('click', () => guard(async () => {
  busy(true, 'Listing workbooks…');
  const data = await api('/api/portfolio');
  portfolioState.total = data.workbooks.length;
  portfolioState.workbooks = [...data.workbooks, ...data.missing];
  portfolioState.missing = new Set(data.missing.map(portfolioKey));
  portfolioState.selected = new Map(data.selected.map((name) => [portfolioKey(name), name]));
  $('portfolio-filter').value = '';
  $('portfolio-missing').hidden = !data.missing.length;
  $('portfolio-missing').textContent = data.missing.length
    ? `${data.missing.length} of your dealerships no longer have a workbook in the accounts folder. Untick them, or tick the renamed workbook.`
    : '';
  renderPortfolio();
  openModal('portfolio-modal');
}));

$('portfolio-filter').addEventListener('input', renderPortfolio);

$('portfolio-tick-shown').addEventListener('click', () => {
  shownWorkbooks()
    .filter((name) => !portfolioState.missing.has(portfolioKey(name)))
    .forEach((name) => portfolioState.selected.set(portfolioKey(name), name));
  renderPortfolio();
});

$('portfolio-clear').addEventListener('click', () => {
  portfolioState.selected.clear();
  renderPortfolio();
});

$('portfolio-save').addEventListener('click', () => guard(async () => {
  busy(true, 'Saving…');
  state.settings = await api('/api/settings', { method: 'POST', body: { portfolio: [...portfolioState.selected.values()] } });
  closeModal('portfolio-modal');
  renderPortfolioButton();
  busy(true, 'Refreshing months…');
  state.months = await api('/api/months');
  renderMonths();
  const count = state.settings.portfolio.length;
  status(count
    ? `Saved. Runs now cover your ${count} dealership${count === 1 ? '' : 's'}.`
    : 'Saved. Runs cover every dealership in the library.', 'good');
}));

/* ---------- source PDF ---------- */

function loadPdf() {
  const frame = $('pdf-frame');
  const empty = $('pdf-empty');
  const documents = (state.currentDetail && state.currentDetail.source_documents) || {};
  // Land on the report Reporter built. Only fall back to a source export when there is
  // no built report, so the pane never opens blank on a dealership that has one.
  if (!loadPdf.pinned && !documents.performance_pdf && $('pdf-kind').value === 'performance') {
    const fallback = documents.sales_pdf ? 'sales' : (documents.service_pdf ? 'service' : 'performance');
    $('pdf-kind').value = fallback;
  }
  const kind = $('pdf-kind').value;
  const has = !!documents[`${kind}_pdf`];
  // Chrome's embedded viewer opens with a toolbar and a thumbnail rail, which together
  // take more of a narrow pane than the page does. Hidden, the whole width goes to the
  // report. scrollbar is deliberately left on — these reports run to several pages.
  frame.src = has ? `${pdfUrl(kind)}#toolbar=0&navpanes=0&view=FitH` : 'about:blank';
  frame.style.visibility = has ? 'visible' : 'hidden';
  empty.hidden = has;
  // "Not found" reads like a fault. Since the figures now come from the workbook and
  // the PDF is only the attachment, say which month has no export and what that costs,
  // so an empty pane is not mistaken for a broken one.
  const month = (state.session && state.session.summary && state.session.summary.month) || '';
  if (kind === 'performance') {
    empty.textContent = documents.performance_note
      || `The Performance Intelligence report was not built for ${month || 'this month'}.`;
  } else {
    const label = kind.charAt(0).toUpperCase() + kind.slice(1);
    empty.textContent = has ? '' : (month
      ? `No ${label} PDF exported for ${month}. The figures come from the workbook, so this month's email will go without the attachment.`
      : `No ${label} PDF exported for this dealership.`);
  }
  if (has) empty.textContent = '';
  $('open-pdf').disabled = !has;
}

const pdfUrl = (kind) => `/api/dealerships/${encodeURIComponent(state.current)}/pdf?kind=${kind}`;

$('pdf-kind').addEventListener('change', () => { loadPdf.pinned = true; loadPdf(); });

$('pdf-expand').addEventListener('click', () => {
  const pane = document.querySelector('.three-pane');
  const wide = pane.classList.toggle('pdf-wide');
  $('pdf-expand').textContent = wide ? '\u2194 Narrower' : '\u2194 Wider';
});

$('pdf-newtab').addEventListener('click', () => {
  if (!state.current) return;
  window.open(pdfUrl($('pdf-kind').value), '_blank', 'noopener');
});
$('open-pdf').addEventListener('click', () => guard(async () => {
  await api(`/api/dealerships/${encodeURIComponent(state.current)}/open-pdf`, { method: 'POST', body: { kind: $('pdf-kind').value } });
  status('Opened in the default PDF application.', 'good');
}));

/* ---------- stage 2: email queue ---------- */

function renderQueue() {
  const body = $('queue-body');
  body.textContent = '';
  const rows = state.session.queue || [];
  $('queue-count').textContent = String(rows.length);
  rows.forEach((row) => {
    const tr = el('tr');
    tr.dataset.name = row.dealership;
    if (row.dealership === state.queueCurrent) tr.classList.add('is-selected');

    const tick = el('td', 'tick-col');
    const box = el('input');
    box.type = 'checkbox';
    box.checked = !!row.ready_to_send;
    // The desktop app lets any row that has not been sent be ticked, and decides what
    // is actually sendable at send time. Disabling the box here instead meant a month
    // whose contacts had not been matched yet could not be ticked at all.
    box.disabled = !!row.sent;
    box.title = row.sent ? 'Already sent.' : sendableReason(row) || 'Ready to send';
    box.addEventListener('click', (event) => event.stopPropagation());
    box.addEventListener('change', () => guard(async () => {
      await api(`/api/dealerships/${encodeURIComponent(row.dealership)}/ready`, { method: 'POST', body: { ready: box.checked } });
      await refreshSession();
    }));
    tick.append(box);
    tr.append(tick);

    tr.append(el('td', null, row.dealership));
    tr.append(el('td', null, row.to || '—'));

    const statusCell = el('td');
    if (row.sent) statusCell.append(el('span', 'flag good', 'sent'));
    else if (!row.contact_ready) statusCell.append(el('span', 'flag bad', 'no contact'));
    else if (!row.email_ready) statusCell.append(el('span', 'flag warn', 'no email'));
    else statusCell.append(el('span', 'flag good', 'ready'));
    tr.append(statusCell);

    tr.addEventListener('click', () => selectQueueRow(row.dealership));
    body.append(tr);
  });
  refreshSendButtons();
}

/* Mirrors reporter_ui._ticked_dealerships: ticked, not already sent, email built. */
const sendableRows = () => (state.session.queue || []).filter((row) => row.ready_to_send && !row.sent && row.email_ready);

function sendableReason(row) {
  if (row.sent) return 'Already sent.';
  if (!row.email_ready) return 'Tickable, but nothing will send until its email is built — approve it on stage 1.';
  if (!row.contact_ready) return 'Tickable, but it has no valid contact address yet.';
  return '';
}

/* Outlook drives sending and drafts. Knowing it cannot start is worth a line at the top
 * of the queue rather than a modal after every email in a batch has failed. */
async function checkOutlook() {
  if (state.outlookChecked || serverMode()) return;
  state.outlookChecked = true;
  let result;
  try {
    result = await api('/api/outlook');
  } catch (err) {
    return;
  }
  const warn = $('outlook-warn');
  warn.hidden = !!result.available;
  $('outlook-warn-text').textContent = result.available ? '' : result.reason;
}

$('outlook-recheck').addEventListener('click', () => {
  state.outlookChecked = false;
  $('outlook-warn').hidden = true;
  checkOutlook();
});

function refreshSendButtons() {
  const ticked = (state.session.queue || []).filter((row) => row.ready_to_send && !row.sent).length;
  const sendable = sendableRows().length;
  ['save-drafts', 'save-eml', 'save-pdfs'].forEach((id) => { $(id).disabled = sendable === 0; });
  $('save-pdfs').textContent = sendable ? `Save ${sendable} PDF cop${sendable === 1 ? 'y' : 'ies'}` : 'Save PDF copies';
  const skipped = ticked > sendable ? `${ticked - sendable} ticked row(s) have no email built yet and will be skipped.` : '';
  ['save-drafts', 'save-eml', 'save-pdfs'].forEach((id) => { $(id).title = skipped || $(id).dataset.title || ''; });
}

async function selectQueueRow(name) {
  state.queueCurrent = name;
  renderQueue();
  busy(true, `Loading ${name}…`);
  const detail = await guard(() => api(`/api/dealerships/${encodeURIComponent(name)}`), 'queue detail');
  if (!detail) return;
  state.currentDetail = detail;
  $('email-dealership').textContent = `${detail.dealership} — ${detail.month}`;
  $('email-to').value = detail.to || '';
  $('email-cc').value = detail.cc || '';
  $('email-bcc').value = detail.bcc || '';
  $('email-subject').value = detail.subject || '';
  $('email-body').value = detail.plain_text_body || '';
  $('email-frame').src = `/api/dealerships/${encodeURIComponent(name)}/email.html`;
  status('');
}

document.querySelectorAll('.email-tab').forEach((tab) => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.email-tab').forEach((other) => other.classList.toggle('is-active', other === tab));
    $('email-view-preview').hidden = tab.dataset.view !== 'preview';
    $('email-view-body').hidden = tab.dataset.view !== 'body';
  });
});

$('save-email').addEventListener('click', () => guard(async () => {
  if (!state.queueCurrent) return;
  busy(true, 'Saving email…');
  await api(`/api/dealerships/${encodeURIComponent(state.queueCurrent)}/email`, {
    method: 'POST',
    body: {
      to_recipients: $('email-to').value,
      cc_recipients: $('email-cc').value,
      bcc_recipients: $('email-bcc').value,
      subject: $('email-subject').value,
      body: $('email-body').value,
    },
  });
  await refreshSession();
  await selectQueueRow(state.queueCurrent);
  status('Email changes saved.', 'good');
}));

$('rebuild-email').addEventListener('click', () => guard(async () => {
  if (!state.queueCurrent) return;
  busy(true, 'Rebuilding the email…');
  await runJob(`/api/dealerships/${encodeURIComponent(state.queueCurrent)}/email/build`, {
    to: $('email-to').value || null,
    cc: $('email-cc').value || null,
    bcc: $('email-bcc').value || null,
  });
  await refreshSession();
  await selectQueueRow(state.queueCurrent);
  status('Email preview rebuilt.', 'good');
}));

$('open-html').addEventListener('click', () => {
  if (!state.queueCurrent) return;
  window.open(`/api/dealerships/${encodeURIComponent(state.queueCurrent)}/email.html`, '_blank', 'noopener');
});

$('tick-all').addEventListener('click', () => setAllReady(true));
$('untick-all').addEventListener('click', () => setAllReady(false));

async function setAllReady(ready) {
  await guard(async () => {
    busy(true, ready ? 'Ticking all…' : 'Unticking all…');
    for (const row of state.session.queue || []) {
      if (row.sent) continue;
      await api(`/api/dealerships/${encodeURIComponent(row.dealership)}/ready`, { method: 'POST', body: { ready } });
    }
    await refreshSession();
  });
}

const tickedNames = () => sendableRows().map((row) => row.dealership);

// Nothing here sends. Every action stops at something a person then acts on: a draft
// in Outlook, a file on disk, a PDF copy. Sending is done from Outlook, by a human.
function runOnTicked(path, label, body = {}) {
  return guard(async () => {
    const names = tickedNames();
    if (!names.length) return;
    const skipped = (state.session.queue || []).filter((row) => row.ready_to_send && !row.sent && !row.email_ready).length;
    if (skipped) status(`${skipped} ticked row(s) have no email built and were skipped.`, 'info');
    busy(true, `${label}…`);
    const result = await runJob(path, { ...body, dealerships: names }, (job) => busy(true, job.message || label));
    await refreshSession();
    showResults(label, result);
  });
}

$('save-drafts').addEventListener('click', () => runOnTicked('/api/drafts', 'Saving drafts'));
$('save-pdfs').addEventListener('click', () => runOnTicked('/api/pdf-copies', 'Saving PDF copies'));

// Writing files needs no Outlook, so this stays available when everything else is not.
$('save-eml').addEventListener('click', () => guard(async () => {
  const names = tickedNames();
  if (!names.length) return;
  if (!serverMode() && !state.settings.eml_export_folder) {
    status('Choose where .eml files should go first — Settings, “Save .eml files to”.', 'bad');
    openModal('settings-modal');
    renderSettings();
    return;
  }
  busy(true, 'Writing .eml files…');
  const result = await runJob('/api/export', { dealerships: names }, (job) => busy(true, job.message || 'Writing .eml files…'));
  await refreshSession();
  // An attachment response downloads in place; the page itself does not navigate.
  if (result.download) window.location.href = result.download;
  showResults(serverMode() ? 'Downloaded .eml' : 'Saved as .eml', result);
}));

function showResults(title, result) {
  $('results-title').textContent = `${title}: ${result.succeeded} ok, ${result.failed} failed`;
  const body = $('results-body');
  body.textContent = '';
  if (result.archived_note) body.append(el('p', 'modal-note', result.archived_note));
  if (result.download) {
    body.append(el('p', 'modal-note', 'Check your Downloads folder. Several emails arrive as one .zip; open a .eml and Outlook shows it ready to send.'));
  } else if (result.folder) {
    body.append(el('p', 'modal-note', `Written to ${result.folder}. Double-click one in Explorer and Outlook opens it ready to send.`));
  }
  (result.results || []).forEach((row) => {
    const line = el('div', 'result-row');
    line.append(el('span', 'who', row.dealership));
    const ok = row.status !== 'failed';
    // On an export the useful confirmation is what actually went into the file, so
    // the attachments are named rather than counted.
    let what = ok ? row.status : row.error || 'failed';
    if (ok && row.filename) {
      const atts = (row.attachments || []).length;
      what = `${row.filename} — ${atts} attachment${atts === 1 ? '' : 's'}`;
    } else if (ok && row.path) {
      // A PDF copy: the folder it landed in is the useful part, since each dealership
      // gets a different one.
      what = row.path.split(/[\\/]/).pop();
      const folder = row.path.slice(0, row.path.length - what.length - 1);
      line.append(el('span', 'who-sub', folder));
    }
    line.append(el('span', `what ${ok ? 'ok' : 'bad'}`, what));
    body.append(line);
  });
  openModal('results-modal');
  status(`${title}: ${result.succeeded} succeeded, ${result.failed} failed.`, result.failed ? 'bad' : 'good');
}

/* ---------- report reference ---------- */

$('open-reference').addEventListener('click', () => guard(async () => {
  const name = state.queueCurrent || state.current;
  if (!name) return;
  const detail = await api(`/api/dealerships/${encodeURIComponent(name)}`);
  const body = $('reference-body');
  body.textContent = '';
  const block = (heading, text) => {
    const wrap = el('div', 'ref-block');
    wrap.append(el('h3', null, heading));
    wrap.append(el('pre', null, text || '—'));
    body.append(wrap);
  };
  block('Executive summary', detail.executive_summary);
  block('Priority insights', detail.priority_insights);
  block('Summary', detail.summary);
  block(`SMS (${(detail.sms_version || '').length}/${state.smsLimit})`, detail.sms_version);
  block('Source PDFs', (detail.source_pdf_paths || []).join('\n'));
  openModal('reference-modal');
}));

/* ---------- stage 3: SMS ---------- */

async function loadSms() {
  const data = await guard(() => api('/api/sms'), 'sms');
  if (!data) return;
  $('sms-doc').value = data.document;
  $('sms-summary').textContent = `${data.dealership_count} approved · ${data.over_limit} over ${data.limit} characters`;
}

$('sms-refresh').addEventListener('click', loadSms);

$('sms-copy').addEventListener('click', async () => {
  const text = $('sms-doc').value;
  if (!text.trim()) return;
  try {
    await navigator.clipboard.writeText(text);
    status('SMS document copied to the clipboard.', 'good');
  } catch (err) {
    $('sms-doc').select();
    status('Press Ctrl+C to copy — the browser blocked automatic clipboard access.', 'info');
  }
});

$('sms-save').addEventListener('click', () => guard(async () => {
  busy(true, 'Saving…');
  const result = await api('/api/sms/save', { method: 'POST', body: { text: $('sms-doc').value } });
  status(`Saved to ${result.path}`, 'good');
}));

$('sms-export').addEventListener('click', () => guard(async () => {
  busy(true, 'Exporting…');
  const result = await api('/api/sms/export', { method: 'POST', body: {} });
  status(`Exported SMS breakdown to ${result.path || result.output_path || 'the output folder'}.`, 'good');
}));

/* ---------- settings ---------- */

const SETTING_FIELDS = [
  ['accounts_folder', 'Accounts folder', 'folder', '', 'The master workbooks, one per dealership.'],
  ['pdf_source_folder', 'PDF source folder', 'folder', '', 'Optional. Where each dealership’s folder of Sales/Service PDFs lives. Leave empty if those folders sit in the accounts folder.'],
  ['output_folder', 'Output folder', 'folder', '', 'Where generated insights and reports are written.'],
  ['recipient_mapping_path', 'Recipient mapping', 'file', '.csv', 'contacts.csv mapping dealerships to contact names and addresses.'],
  ['default_bcc', 'Always BCC', 'any', '', 'Addresses BCC’d on every email, on top of anything a contact row asks for. Separate them with semicolons.'],
  ['sent_email_pdf_folder', 'PDF copies folder', 'folder', '', 'Optional. Leave empty and each PDF copy is saved beside that dealership’s own source reports.'],
  ['eml_export_folder', 'Save .eml files to', 'folder', '', 'Where “Save as .eml” writes emails you want to import into Outlook by hand.'],
  ['sender_name', 'Sign-off name', 'any', '', 'Optional. Leave both sign-off fields empty and the email carries no signature at all.'],
  ['sender_title', 'Sign-off title', 'any', '', 'Optional. Shown under the name, e.g. “CRM Consultant”.'],
];

function renderSettings() {
  const wrap = $('settings-fields');
  wrap.textContent = '';
  SETTING_FIELDS.forEach(([key, label, kind, extensions, description]) => {
    const editable = canEdit(key);
    const row = el('div', 'setting-row');
    const labelNode = el('label', null, label);
    labelNode.setAttribute('for', `setting-${key}`);
    row.append(labelNode);
    row.append(el('p', 'desc', editable ? description : `${description} Set by an administrator.`));
    const line = el('div', 'with-check');
    const input = el('input');
    input.type = 'text';
    input.id = `setting-${key}`;
    input.value = state.settings[key] || '';
    input.spellcheck = false;
    input.readOnly = !editable;
    line.append(input);
    // Browsing lists the server's own disks, so on a server only an administrator may.
    if (kind !== 'any' && (!serverMode() || state.settings.is_admin)) {
      const browse = el('button', 'btn btn-small', 'Browse…');
      browse.type = 'button';
      browse.addEventListener('click', () => openBrowse({
        title: kind === 'folder' ? `Choose the ${label.toLowerCase()}` : `Choose the ${label.toLowerCase()}`,
        kind,
        extensions,
        start: input.value,
        onPick: (chosen) => {
          input.value = chosen;
          note.textContent = 'Found.';
          note.className = 'check-note ok';
        },
      }));
      line.append(browse);

      const check = el('button', 'btn btn-small', 'Check');
      check.type = 'button';
      check.addEventListener('click', async () => {
        const result = await api('/api/settings/check-path', { method: 'POST', body: { path: input.value, kind } });
        note.textContent = result.note;
        note.className = `check-note ${result.exists && result.note === 'Found.' ? 'ok' : 'bad'}`;
      });
      line.append(check);
    }
    row.append(line);
    const note = el('p', 'check-note', '');
    row.append(note);
    wrap.append(row);
  });
  $('insight-engine').value = state.settings.insight_engine || 'rules';
  $('performance-report').value = state.settings.performance_report || 'on';
  ['insight-engine', 'performance-report'].forEach((id) => { $(id).disabled = !canEdit('insight_engine'); });
  $('performance-engine-line').textContent = `Rendering with: ${state.settings.performance_engine || 'unknown'}`;
  $('settings-file-path').textContent = `Saved to ${state.settings.settings_path || 'user_settings.json'}`;
  renderAiStatus();
}

// The AI connection is configured in a .env file, not here: the key is never sent to
// the browser and never stored in user_settings.json. All Settings does is say whether
// the app can see one, and let the operator prove it works before month end.
function renderAiStatus(result) {
  const wrap = $('ai-status');
  const line = $('ai-status-line');
  if (!wrap || !line) return;
  const chosen = $('insight-engine').value === 'ai';
  wrap.style.display = chosen ? 'flex' : 'none';
  if (!chosen) return;

  if (result) {
    line.textContent = result.message;
    line.className = `ai-status-line ${result.ok ? 'ok' : 'bad'}`;
    return;
  }

  const ai = state.settings.ai || {};
  if (ai.api_key_set) {
    const where = (ai.env_files && ai.env_files.length) ? ` from ${ai.env_files[0]}` : '';
    line.textContent = `Key ${ai.api_key_masked} found${where} — ${ai.provider} / ${ai.model}.`;
    line.className = 'ai-status-line ok';
  } else {
    line.textContent = ai.reason || 'No API key found. Copy .env.example to .env and add one.';
    line.className = 'ai-status-line bad';
  }
}

$('insight-engine').addEventListener('change', () => renderAiStatus());

$('test-ai').addEventListener('click', () => guard(async () => {
  const line = $('ai-status-line');
  line.textContent = 'Testing…';
  line.className = 'ai-status-line idle';
  const result = await api('/api/settings/ai/test', { method: 'POST', body: {} });
  if (result.status) state.settings.ai = result.status;
  renderAiStatus(result);
}));

$('open-settings').addEventListener('click', () => { renderSettings(); openModal('settings-modal'); });

$('save-settings').addEventListener('click', () => guard(async () => {
  const payload = {
    insight_engine: $('insight-engine').value,
    performance_report: $('performance-report').value,
  };
  SETTING_FIELDS.forEach(([key]) => { payload[key] = $(`setting-${key}`).value.trim(); });
  busy(true, 'Saving settings…');
  state.settings = await api('/api/settings', { method: 'POST', body: payload });
  renderAiStatus();
  renderPortfolioButton();
  state.months = await api('/api/months');
  renderMonths();
  refreshSendButtons();
  closeModal('settings-modal');
  status('Settings saved.', 'good');
}));

/* ---------- file and folder picker ----------
 *
 * A browser cannot open a native dialog on a server's behalf, and typing a full Windows
 * path from memory is exactly the step this app should not ask for. The server runs on
 * the machine that holds the files, so it lists them and this walks the result.
 */

const browseState = { kind: 'folder', extensions: '', path: '', parent: '', selectedFile: '', onPick: null };

function openBrowse({ title, kind, extensions, start, onPick }) {
  Object.assign(browseState, { kind, extensions: extensions || '', selectedFile: '', onPick });
  $('browse-title').textContent = title;
  $('browse-choose').textContent = kind === 'folder' ? 'Use this folder' : 'Select file';
  openModal('browse-modal');
  loadBrowse(start || '');
}

async function loadBrowse(path) {
  const data = await guard(() => api('/api/browse', {
    method: 'POST',
    body: { path, kind: browseState.kind, extensions: browseState.extensions },
  }), 'browse');
  if (!data) return;
  browseState.path = data.path;
  browseState.parent = data.parent || '';
  browseState.selectedFile = '';
  $('browse-path').value = data.path;
  $('browse-up').disabled = !data.parent;
  $('browse-note').hidden = !data.truncated;
  $('browse-note').textContent = data.truncated ? 'Showing the first 400 entries. Narrow it by opening a subfolder.' : '';

  const places = $('browse-places');
  places.textContent = '';
  [...(data.places || []).map((place) => [place.label, place.path]),
   ...(data.drives || []).map((drive) => [drive, drive])].forEach(([label, target]) => {
    const chip = el('button', null, label);
    chip.type = 'button';
    chip.addEventListener('click', () => loadBrowse(target));
    places.append(chip);
  });

  const list = $('browse-list');
  list.textContent = '';
  if (!data.entries.length) {
    const empty = el('li', 'browse-empty', browseState.kind === 'folder' ? 'No subfolders here.' : 'Nothing matching here.');
    list.append(empty);
  }
  data.entries.forEach((entry) => {
    const item = el('li');
    item.append(el('span', 'glyph', entry.is_dir ? '\u{1F5C0}' : '\u{1F5CE}'));
    item.append(el('span', 'name', entry.name));
    if (!entry.is_dir) item.append(el('span', 'size', formatBytes(entry.size)));
    item.addEventListener('click', () => {
      if (entry.is_dir) { loadBrowse(entry.path); return; }
      browseState.selectedFile = entry.path;
      list.querySelectorAll('li').forEach((other) => other.classList.toggle('is-selected', other === item));
      $('browse-selected').textContent = entry.path;
    });
    item.addEventListener('dblclick', () => { if (!entry.is_dir) $('browse-choose').click(); });
    list.append(item);
  });
  $('browse-selected').textContent = browseState.kind === 'folder' ? data.path : '';
}

const formatBytes = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : n >= 1024 ? `${Math.round(n / 1024)} KB` : `${n} B`);

// The server knows what the parent is on its own filesystem; do not guess separators.
$('browse-up').addEventListener('click', () => { if (browseState.parent) loadBrowse(browseState.parent); });
$('browse-go').addEventListener('click', () => loadBrowse($('browse-path').value));
$('browse-path').addEventListener('keydown', (event) => { if (event.key === 'Enter') loadBrowse($('browse-path').value); });

$('browse-choose').addEventListener('click', () => {
  const chosen = browseState.kind === 'folder' ? browseState.path : browseState.selectedFile;
  if (!chosen) { status('Pick a file first.', 'bad'); return; }
  if (browseState.onPick) browseState.onPick(chosen);
  closeModal('browse-modal');
});

/* ---------- contacts ---------- */

$('open-contacts').addEventListener('click', () => guard(async () => {
  busy(true, 'Loading contacts…');
  state.contacts = await api('/api/contacts');
  state.contactsDirty = false;
  contactsStatus('Changes are saved to the file shown below when you click Save contacts.');
  renderContacts();
  openModal('contacts-modal');
}));

// Contacts are grouped by dealership rather than listed as flat rows: several people
// at one site is the normal case, and as a flat table those rows scatter and read as
// unrelated. Notes and active are deliberately not shown — the values already in the
// file are carried through untouched, they are just not worth the operator's attention.
const CONTACT_FIELDS = [
  ['contact_name', 'Contact name', 'Sam Taylor'],
  ['email', 'Email', 'sam@example.com'],
  ['cc', 'CC', 'optional'],
  ['bcc', 'BCC', 'optional'],
];

const today = () => new Date().toISOString().slice(0, 10);

function contactsStatus(message, failed = false) {
  const note = $('contacts-save-status');
  note.textContent = message;
  note.className = failed ? 'modal-note warn' : 'modal-note';
}

function markContactsDirty() {
  state.contactsDirty = true;
  contactsStatus('Unsaved contact changes — click Save contacts to write them to your file.');
}

function contactGroups() {
  const groups = [];
  const index = new Map();
  (state.contacts.rows || []).forEach((row) => {
    const key = (row.dealership || '').trim();
    if (!index.has(key)) {
      index.set(key, { dealership: key, rows: [] });
      groups.push(index.get(key));
    }
    index.get(key).rows.push(row);
  });
  return groups;
}

function renderContacts() {
  $('contacts-path').textContent = state.contacts.path;
  const known = state.contacts.dealership_names || [];
  const groups = contactGroups();
  updateContactsNote();

  const body = $('contacts-body');
  body.textContent = '';
  // A running count of what is actually here. "Only one saved" is invisible until
  // something says how many there are.
  const people = (state.contacts.rows || []).length;
  $('contacts-count').textContent = groups.length
    ? `${groups.length} dealership${groups.length === 1 ? '' : 's'}, ${people} contact${people === 1 ? '' : 's'}`
    : '';
  if (!groups.length) {
    body.append(el('p', 'contacts-empty', 'No contacts yet. Add a dealership to start.'));
    return;
  }
  groups.forEach((group) => body.append(contactCard(group, known)));
}

function addDealership() {
  state.contacts.rows.push(blankContactRow(''));
  markContactsDirty();
  renderContacts();
  // Put the new card in front of the operator and in the control they need next,
  // rather than appending it below the fold of a scrolling list.
  const card = $('contacts-body').lastElementChild;
  if (card) {
    card.scrollIntoView({ block: 'nearest' });
    const picker = card.querySelector('.contact-dealership');
    if (picker) picker.focus();
  }
}

function updateContactsNote() {
  const known = state.contacts.dealership_names || [];
  const note = $('contacts-note');
  if (!note) return;
  const unmatched = [...new Set(
    (state.contacts.rows || [])
      .map((row) => (row.dealership || '').trim())
      .filter((name) => name && known.length && !known.includes(name)),
  )];
  if (!known.length) {
    note.hidden = false;
    note.className = 'modal-note';
    note.textContent = 'Set the accounts folder in Settings and the dealership becomes a list to pick from.';
  } else if (unmatched.length) {
    note.hidden = false;
    note.className = 'modal-note warn';
    note.textContent = `No folder in the accounts folder for ${unmatched.slice(0, 4).join(', ')}${unmatched.length > 4 ? ' and others' : ''}. Those emails will have no recipient.`;
  } else {
    note.hidden = true;
  }
}

function contactCard(group, known) {
  const card = el('div', 'contact-card');
  const head = el('div', 'contact-card-head');
  head.append(dealershipPicker(group, known));
  if (group.rows.length > 1) {
    head.append(el('span', 'contact-count', `${group.rows.length} contacts`));
  }
  card.append(head);

  const table = el('div', 'contact-grid');
  const header = el('div', 'contact-row contact-row-head');
  CONTACT_FIELDS.forEach(([, label]) => header.append(el('span', null, label)));
  header.append(el('span', null, ''));
  table.append(header);

  group.rows.forEach((row) => {
    const line = el('div', 'contact-row');
    CONTACT_FIELDS.forEach(([field, label, placeholder]) => {
      const input = el('input');
      input.type = 'text';
      input.value = row[field] || '';
      input.placeholder = placeholder;
      input.spellcheck = false;
      input.setAttribute('aria-label', `${label} for ${group.dealership || 'this dealership'}`);
      input.addEventListener('input', () => {
        row[field] = input.value;
        row.last_updated = today();
        markContactsDirty();
      });
      line.append(input);
    });
    const remove = el('button', 'btn btn-small btn-quiet', 'Remove');
    remove.type = 'button';
    remove.addEventListener('click', () => {
      state.contacts.rows.splice(state.contacts.rows.indexOf(row), 1);
      markContactsDirty();
      renderContacts();
    });
    line.append(remove);
    table.append(line);
  });
  card.append(table);

  const add = el('button', 'btn btn-small', '+ Add another contact here');
  add.type = 'button';
  add.addEventListener('click', () => {
    state.contacts.rows.push(blankContactRow(group.dealership));
    markContactsDirty();
    renderContacts();
  });
  card.append(add);
  return card;
}

function dealershipPicker(group, known) {
  const current = (group.dealership || '').trim();

  if (!known.length) {
    const input = el('input', 'contact-dealership');
    input.type = 'text';
    input.value = current;
    input.placeholder = 'Dealership name';
    input.spellcheck = false;
    input.addEventListener('input', () => {
      setGroupDealership(group, input.value);
      updateContactsNote();
    });
    return input;
  }

  const select = el('select', 'contact-dealership');
  const blank = el('option', null, 'Select');
  blank.value = '';
  select.append(blank);
  // A row naming something with no folder keeps its value rather than being rewritten
  // to whatever sorts first; the warning above the list is where it gets explained.
  if (current && !known.includes(current)) {
    const orphan = el('option', null, current);
    orphan.value = current;
    select.append(orphan);
    select.classList.add('warn');
  }
  known.forEach((name) => {
    const option = el('option', null, name);
    option.value = name;
    select.append(option);
  });
  select.value = current;
  select.addEventListener('change', () => {
    setGroupDealership(group, select.value);
    select.classList.toggle('warn', Boolean(select.value) && !known.includes(select.value));
    updateContactsNote();
  });
  return select;
}

function setGroupDealership(group, name) {
  // Every contact in the card belongs to the same site, so the name moves together.
  group.rows.forEach((row) => { row.dealership = name; row.last_updated = today(); });
  group.dealership = name;
  markContactsDirty();
}

function blankContactRow(dealership) {
  const row = {};
  (state.contacts.fieldnames || []).forEach((field) => { row[field] = ''; });
  row.dealership = dealership || '';
  row.active = 'yes';
  row.last_updated = today();
  return row;
}

$('contacts-add-top').addEventListener('click', addDealership);

$('contacts-save').addEventListener('click', () => guard(async () => {
  busy(true, 'Saving contacts…');
  try {
    const rows = state.contacts.rows.filter((row) => (row.dealership || '').trim() || (row.email || '').trim() || (row.contact_name || '').trim());
    const result = await api('/api/contacts', { method: 'POST', body: { path: state.contacts.path, revision: state.contacts.revision, rows } });
    // The panel stays open and re-renders from what the server actually wrote.
    if (result.contacts) {
      state.contacts = result.contacts;
      renderContacts();
    }
    state.contactsDirty = false;
    const names = result.dealerships || [];
    const summary = names.length
      ? `Saved ${result.rows} contact(s) across ${names.length} dealership(s): ${names.slice(0, 6).join(', ')}${names.length > 6 ? '…' : ''}.`
      : 'Saved. The contacts file is now empty.';
    status(`${summary} ${result.note}`.trim(), 'good');
    contactsStatus(`${summary} Saved to ${result.path}. ${result.note}`.trim());
    // A session refresh failure must not make a completed disk save look unsuccessful.
    try { await refreshSession(); } catch (err) {
      contactsStatus(`Contacts saved to ${result.path}. Could not refresh the reporting session: ${err.message}`, true);
    }
  } catch (err) {
    contactsStatus(err.message || String(err), true);
    throw err;
  }
}));

/* ---------- go ---------- */

$('status-dismiss').addEventListener('click', () => status(''));
window.addEventListener('beforeunload', (event) => {
  if (state.dirty || state.contactsDirty) { event.preventDefault(); event.returnValue = ''; }
});

bootstrap();
