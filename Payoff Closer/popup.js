'use strict';

let pollTimer = null;
let currentMode = 'single';

function setStatus(msg, type) {
  const el = document.getElementById('statusMsg');
  el.textContent = msg;
  el.className = 'status-msg' + (type ? ' ' + type : '');
}

// ── Pin to a persistent window ──────────────────────────────────────────────
// A toolbar-click popup always closes the instant it loses focus — nothing in
// HTML/JS can override that. The workaround is to pop the same page out into
// its own small window instead, which behaves like any normal window and stays
// open no matter what you click, until you close it yourself.

const isPinned = new URLSearchParams(window.location.search).get('pinned') === '1';
if (isPinned) {
  document.getElementById('pinBtn').style.display = 'none';
}

document.getElementById('pinBtn').addEventListener('click', async () => {
  // Carry over whatever's currently typed in, so pinning mid-entry doesn't
  // lose anything.
  const draft = {
    mode: currentMode,
    shipmentInput: document.getElementById('shipmentInput').value,
    shipmentCount: document.getElementById('shipmentCountInput').value,
    multiShipments: [...document.querySelectorAll('.multi-shipment-input')].map(i => i.value),
    usdPaid: document.getElementById('usdInput').value,
    egpPaid: document.getElementById('egpInput').value,
    bankName: document.getElementById('bankInput').value,
  };
  await chrome.storage.local.set({ popupDraft: draft });
  await chrome.windows.create({
    url: chrome.runtime.getURL('popup.html?pinned=1'),
    type: 'popup',
    width: 440,
    height: 760,
    focused: true,
  });
  window.close();
});

// ── Mode toggle ──────────────────────────────────────────────────────────────

function setMode(mode) {
  currentMode = mode;
  document.getElementById('modeSingleBtn').classList.toggle('active', mode === 'single');
  document.getElementById('modeMultiBtn').classList.toggle('active', mode === 'multi');
  document.getElementById('singleModeSection').style.display = mode === 'single' ? 'block' : 'none';
  document.getElementById('multiModeSection').style.display = mode === 'multi' ? 'block' : 'none';
}
document.getElementById('modeSingleBtn').addEventListener('click', () => setMode('single'));
document.getElementById('modeMultiBtn').addEventListener('click', () => setMode('multi'));

function renderShipmentInputs(count) {
  const container = document.getElementById('multiShipmentInputs');
  const existing = [...container.querySelectorAll('.multi-shipment-input')].map(i => i.value);
  container.innerHTML = '';
  for (let i = 0; i < count; i++) {
    const row = document.createElement('div');
    row.className = 'field-row';
    row.innerHTML = `<label class="field-label">Shipment ${i + 1}</label>
      <input type="text" class="text-input multi-shipment-input" placeholder="e.g. NGBSOK269879874" />`;
    container.appendChild(row);
    if (existing[i]) row.querySelector('input').value = existing[i];
  }
}
document.getElementById('shipmentCountInput').addEventListener('input', (e) => {
  const n = Math.max(0, Math.min(20, parseInt(e.target.value, 10) || 0));
  renderShipmentInputs(n);
});

function getShipmentNos() {
  if (currentMode === 'single') {
    const s = document.getElementById('shipmentInput').value.trim();
    return s ? [s] : [];
  }
  return [...document.querySelectorAll('.multi-shipment-input')].map(i => i.value.trim());
}

// ── Unusual-charges BL list ──────────────────────────────────────────────────
// Occasional maintenance data (updated every couple of weeks), separate from
// the daily shipment/deposit entry above. Once loaded, it stays applied to
// every balance check regardless of whether this panel is open or collapsed.

let unusualChargesSet = new Set();

async function loadUnusualCharges() {
  const r = await chrome.storage.local.get(['unusualChargesList']);
  const list = r.unusualChargesList;
  const statusEl = document.getElementById('unusualChargesStatus');
  const summaryEl = document.getElementById('unusualChargesSummary');
  if (list && list.numbers && list.numbers.length) {
    unusualChargesSet = new Set(list.numbers.map(n => String(n).toUpperCase()));
    const dateStr = list.updatedAt ? new Date(list.updatedAt).toLocaleDateString() : '';
    statusEl.textContent = `${list.numbers.length} BL number(s) loaded from "${list.fileName || 'file'}"${dateStr ? ' — updated ' + dateStr : ''}.`;
    summaryEl.textContent = `(${list.numbers.length} loaded)`;
  } else {
    unusualChargesSet = new Set();
    statusEl.textContent = 'No list loaded yet.';
    summaryEl.textContent = '';
  }
}

document.getElementById('unusualChargesToggle').addEventListener('change', (e) => {
  document.getElementById('unusualChargesPanel').style.display = e.target.checked ? 'block' : 'none';
});

document.getElementById('unusualChargesFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById('unusualChargesStatus');
  statusEl.textContent = 'Reading file...';
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rows.length) throw new Error('The sheet appears to be empty.');

    // Look for a column whose header suggests it holds BL/shipment numbers;
    // fall back to the first column (treated as data, no header) if nothing
    // obvious matches.
    const headerRow = rows[0].map(h => String(h).trim().toLowerCase());
    let colIndex = headerRow.findIndex(h => /\bbl\b|bill of lading|shipment/.test(h));
    let dataRows;
    if (colIndex === -1) { colIndex = 0; dataRows = rows; }
    else { dataRows = rows.slice(1); }

    const numbers = dataRows.map(r => String(r[colIndex] || '').trim()).filter(Boolean);
    if (!numbers.length) throw new Error('No BL numbers found in that file.');

    await chrome.storage.local.set({
      unusualChargesList: { numbers, updatedAt: Date.now(), fileName: file.name }
    });
    await loadUnusualCharges();
  } catch (err) {
    statusEl.textContent = 'Error reading file: ' + err.message;
  }
});

loadUnusualCharges();

// ── Progress / status display ───────────────────────────────────────────────

function phaseLabel(status, state) {
  const total = (state.shipmentNos || []).length;
  const shipmentNo = (state.shipmentNos || [])[state.index] || '';
  const prefix = total > 1 ? `[${(state.index || 0) + 1}/${total} ${shipmentNo}] ` : '';
  switch (status) {
    case 'SEARCHING':
      return prefix + (state.phase === 'checking' ? 'Checking shipment...' : 'Opening shipment...');
    case 'CLICKING_RESULT':  return prefix + 'Opening shipment...';
    case 'ON_SHIPMENT_PAGE': return prefix + (state.phase === 'checking' ? 'Reading invoices...' : 'Starting...');
    case 'FINALIZING': {
      const t = (state.toFinalize || []).length;
      return prefix + `Finalizing invoice ${(state.finalizeIndex || 0) + 1}/${t}...`;
    }
    case 'PAYMENTS_SEARCH':  return prefix + 'Opening payments list...';
    case 'MARKING_PAID':     return prefix + `Recording payment ${(state.paidCount || 0) + 1}...`;
    case 'GENERATE_RECEIPT': return prefix + `Generating receipt ${(state.paidCount || 0) + 1}...`;
    default: return status;
  }
}

function updateProgress(status, state) {
  const total = (state.shipmentNos || []).length || 1;
  const idx = state.index || 0;
  let pct = state.phase === 'checking' ? (idx / total) * 50 : 50 + (idx / total) * 50;
  pct = Math.max(2, Math.min(98, Math.round(pct)));
  document.getElementById('progressBar').style.width = pct + '%';
  document.getElementById('progressLabel').textContent = phaseLabel(status, state);
}

function renderBalanceCheck(state) {
  const due = state.balance || {};
  const paidMap = { USD: state.usdPaid || 0, EGP: state.egpPaid || 0 };
  const currencies = Object.keys(due);
  const resultDiv = document.getElementById('balanceResult');
  const breakdownDiv = document.getElementById('shipmentBreakdown');
  const confirmBtn = document.getElementById('confirmBtn');

  document.getElementById('balanceSection').style.display = 'block';

  // Per-shipment breakdown, always shown, down to the individual invoice lines
  // that make up each total — so an unexpected line is visible before you
  // confirm, rather than hiding inside a combined number.
  const shipmentNos = state.shipmentNos || [];
  let bHtml = '<div class="shipment-breakdown">';
  shipmentNos.forEach(sn => {
    const lines = (state.shipmentUnpaidLines && state.shipmentUnpaidLines[sn]) || [];
    const flagged = unusualChargesSet.has(String(sn).toUpperCase());
    bHtml += `<div class="line"><strong>${sn}</strong>${flagged ? '<span class="unusual-flag">⚠ unusual charges — review manually</span>' : ''}`;
    if (lines.length === 0) {
      bHtml += ' — nothing unpaid</div>';
    } else {
      bHtml += '</div>';
      lines.forEach(l => {
        bHtml += `<div class="invoice-line">${l.currency} ${l.amount.toFixed(2)} <span class="muted">(${l.status}${l.invoiceId ? ', #' + l.invoiceId : ''})</span></div>`;
      });
    }
  });
  bHtml += '</div>';
  breakdownDiv.innerHTML = bHtml;

  if (currencies.length === 0) {
    resultDiv.innerHTML = '<div class="balance-row ok">No unpaid invoices found.</div>';
    confirmBtn.style.display = 'none';
    setStatus('Nothing to finalize here.', '');
    return;
  }

  let allOk = true;
  let html = '';
  currencies.forEach(cur => {
    const dueAmt = due[cur];
    const paidAmt = paidMap[cur] !== undefined ? paidMap[cur] : 0;
    const diff = paidAmt - dueAmt;
    const ok = diff >= -0.01;
    if (!ok) allOk = false;
    const label = shipmentNos.length > 1 ? `${cur} total` : cur;
    let note, cls;
    if (!ok) {
      note = `(short by ${Math.abs(diff).toFixed(2)})`;
      cls = 'short';
    } else if (diff > 0.01) {
      note = `(covered — overpaid by ${diff.toFixed(2)})`;
      cls = 'over';
    } else {
      note = '(covered)';
      cls = 'ok';
    }
    html += `<div class="balance-row ${cls}">${label}: due ${dueAmt.toFixed(2)} &mdash; deposited ${paidAmt.toFixed(2)} ${note}</div>`;
  });
  resultDiv.innerHTML = html;
  confirmBtn.style.display = 'block';
  confirmBtn.disabled = !allOk;
  setStatus(allOk ? 'Fully covered — ready to finalize.' : 'Not fully paid — do not proceed.', allOk ? 'done' : 'error');
}

// ── Polling ──────────────────────────────────────────────────────────────────

function startPolling() {
  stopPolling();
  pollTimer = setInterval(async () => {
    const result = await chrome.storage.local.get(['payoffAutomation']);
    const state = result.payoffAutomation;
    if (!state) return;

    if (state.status === 'AWAITING_CONFIRM') {
      document.getElementById('progressSection').style.display = 'none';
      document.getElementById('stopBtn').style.display = 'none';
      document.getElementById('checkBtn').disabled = false;
      renderBalanceCheck(state);
      stopPolling();
    } else if (state.status === 'DONE') {
      document.getElementById('progressBar').style.width = '100%';
      document.getElementById('progressLabel').textContent = 'Done';
      setStatus('Done — finalized, paid, and receipts saved. Attach the files and send when ready.', 'done');
      document.getElementById('stopBtn').style.display = 'none';
      document.getElementById('checkBtn').disabled = false;
      stopPolling();
    } else if (state.status === 'ERROR') {
      setStatus('Error: ' + state.errorMsg, 'error');
      document.getElementById('stopBtn').style.display = 'none';
      document.getElementById('checkBtn').disabled = false;
      stopPolling();
    } else {
      document.getElementById('progressSection').style.display = 'block';
      document.getElementById('stopBtn').style.display = 'block';
      updateProgress(state.status, state);
    }
  }, 1000);
}
function stopPolling() { if (pollTimer) { clearInterval(pollTimer); pollTimer = null; } }

// ── Buttons ──────────────────────────────────────────────────────────────────

document.getElementById('checkBtn').addEventListener('click', async () => {
  const shipmentNos = getShipmentNos();
  const bankName = document.getElementById('bankInput').value.trim();
  const usdPaid = parseFloat(document.getElementById('usdInput').value) || 0;
  const egpPaid = parseFloat(document.getElementById('egpInput').value) || 0;

  if (shipmentNos.length === 0 || shipmentNos.some(s => !s)) {
    setStatus(currentMode === 'single' ? 'Enter a shipment number.' : 'Fill in every shipment number field.', 'error');
    return;
  }
  if (!bankName) { setStatus('Enter the bank name.', 'error'); return; }

  document.getElementById('balanceSection').style.display = 'none';
  document.getElementById('confirmBtn').style.display = 'none';
  document.getElementById('checkBtn').disabled = true;

  const state = {
    status: 'SEARCHING',
    phase: 'checking',
    shipmentNos,
    index: 0,
    usdPaid, egpPaid, bankName,
    shipmentBalances: {},
    shipmentInvoices: {},
  };
  await chrome.storage.local.set({ payoffAutomation: state });
  setStatus(shipmentNos.length > 1 ? `Checking shipment 1/${shipmentNos.length}...` : 'Searching for shipment...', '');
  document.getElementById('progressSection').style.display = 'block';
  document.getElementById('stopBtn').style.display = 'block';
  updateProgress('SEARCHING', state);

  const tabs = await chrome.tabs.query({ url: 'https://internal-company-website/*' });
  if (tabs.length > 0) {
    await chrome.tabs.update(tabs[0].id, { active: true, url: 'https://internal-company-website/shipments' });
  } else {
    await chrome.tabs.create({ url: 'https://internal-company-website/shipments' });
  }
  startPolling();
});

document.getElementById('confirmBtn').addEventListener('click', async () => {
  const result = await chrome.storage.local.get(['payoffAutomation']);
  const state = result.payoffAutomation;
  if (!state || state.status !== 'AWAITING_CONFIRM') return;

  document.getElementById('confirmBtn').disabled = true;
  document.getElementById('progressSection').style.display = 'block';
  document.getElementById('stopBtn').style.display = 'block';
  setStatus('Finalizing...', '');

  if (state.shipmentNos.length === 1) {
    // We're already sitting on this shipment's page from the balance check —
    // no need to leave, search, and click back into it. Jump straight to
    // finalizing using the invoice list already gathered.
    const shipmentNo = state.shipmentNos[0];
    const toFinalize = (state.shipmentInvoices && state.shipmentInvoices[shipmentNo]) || [];
    const directState = { ...state, status: 'FINALIZING', phase: 'payoff', index: 0, finalizeIndex: 0, toFinalize, paidCount: 0 };
    await chrome.storage.local.set({ payoffAutomation: directState });
    updateProgress('FINALIZING', directState);
    startPolling();
    return;
  }

  const newState = { ...state, status: 'SEARCHING', phase: 'payoff', index: 0, paidCount: 0 };
  await chrome.storage.local.set({ payoffAutomation: newState });
  updateProgress('SEARCHING', newState);

  const tabs = await chrome.tabs.query({ url: 'https://internal-company-website/*' });
  if (tabs.length > 0) {
    await chrome.tabs.update(tabs[0].id, { active: true, url: 'https://internal-company-website/shipments' });
  } else {
    await chrome.tabs.create({ url: 'https://internal-company-website/shipments' });
  }
  startPolling();
});

document.getElementById('stopBtn').addEventListener('click', async () => {
  await chrome.storage.local.set({ payoffAutomation: { status: 'IDLE' } });
  document.getElementById('stopBtn').style.display = 'none';
  document.getElementById('checkBtn').disabled = false;
  setStatus('Stopped.', '');
  stopPolling();
});

// Restore any form draft carried over from pinning (only relevant if no
// automation job is already in progress — that restore below takes priority).
(async () => {
  const draftResult = await chrome.storage.local.get(['popupDraft']);
  const draft = draftResult.popupDraft;
  if (!draft) return;
  await chrome.storage.local.remove('popupDraft');

  if (draft.mode === 'multi') {
    setMode('multi');
    document.getElementById('shipmentCountInput').value = draft.shipmentCount || '';
    const n = parseInt(draft.shipmentCount, 10) || (draft.multiShipments || []).length;
    renderShipmentInputs(n);
    [...document.querySelectorAll('.multi-shipment-input')].forEach((el, i) => { el.value = (draft.multiShipments || [])[i] || ''; });
  } else {
    document.getElementById('shipmentInput').value = draft.shipmentInput || '';
  }
  document.getElementById('usdInput').value = draft.usdPaid || '';
  document.getElementById('egpInput').value = draft.egpPaid || '';
  document.getElementById('bankInput').value = draft.bankName || '';
})();

// Restore in-progress state on popup open
(async () => {
  const result = await chrome.storage.local.get(['payoffAutomation']);
  const state = result.payoffAutomation;
  if (!state || state.status === 'IDLE') return;

  if (state.shipmentNos && state.shipmentNos.length > 1) {
    setMode('multi');
    document.getElementById('shipmentCountInput').value = state.shipmentNos.length;
    renderShipmentInputs(state.shipmentNos.length);
    [...document.querySelectorAll('.multi-shipment-input')].forEach((el, i) => { el.value = state.shipmentNos[i] || ''; });
  } else if (state.shipmentNos && state.shipmentNos.length === 1) {
    document.getElementById('shipmentInput').value = state.shipmentNos[0];
  }
  if (state.bankName) document.getElementById('bankInput').value = state.bankName;
  if (state.usdPaid) document.getElementById('usdInput').value = state.usdPaid;
  if (state.egpPaid) document.getElementById('egpInput').value = state.egpPaid;

  if (state.status === 'AWAITING_CONFIRM') {
    await loadUnusualCharges();
    renderBalanceCheck(state);
  } else if (state.status === 'DONE') {
    setStatus('Done — finalized, paid, and receipts saved.', 'done');
  } else if (state.status === 'ERROR') {
    setStatus('Error: ' + state.errorMsg, 'error');
  } else {
    document.getElementById('checkBtn').disabled = true;
    document.getElementById('progressSection').style.display = 'block';
    document.getElementById('stopBtn').style.display = 'block';
    updateProgress(state.status, state);
    startPolling();
  }
})();
