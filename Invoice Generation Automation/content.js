'use strict';

// ── Helpers ───────────────────────────────────────────────────────────────────

const sleep = ms => new Promise(r => setTimeout(r, ms));

function waitFor(selector, timeout = 12000) {
  return new Promise((resolve, reject) => {
    const el = document.querySelector(selector);
    if (el) return resolve(el);
    const obs = new MutationObserver(() => {
      const el = document.querySelector(selector);
      if (el) { obs.disconnect(); resolve(el); }
    });
    obs.observe(document.body, { childList: true, subtree: true });
    setTimeout(() => { obs.disconnect(); reject(new Error('Timeout: ' + selector)); }, timeout);
  });
}

function triggerAll(el) {
  ['input', 'change', 'blur'].forEach(e => el.dispatchEvent(new Event(e, { bubbles: true })));
}

function showOverlay(msg, color = '#2f6fed') {
  let el = document.getElementById('__blOverlay__');
  if (!el) {
    el = document.createElement('div');
    el.id = '__blOverlay__';
    el.style.cssText = `position:fixed;bottom:20px;right:20px;padding:10px 16px;
      border-radius:8px;color:white;font-size:13px;font-weight:500;z-index:999999;
      box-shadow:0 4px 12px rgba(0,0,0,0.25);max-width:360px;line-height:1.5;
      font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;`;
    document.body.appendChild(el);
  }
  el.style.background = color;
  el.textContent = msg;
}

async function saveError(state, errorBL, errorMsg) {
  await chrome.storage.local.set({ blAutomation: { ...state, status: 'ERROR', errorBL, errorMsg } });
  showOverlay('❌ ' + errorBL + ': ' + errorMsg, '#f44336');
}

// ── Cost Fill Data ────────────────────────────────────────────────────────────

const DATASETS = {
  "1": { data: [120,10,5,10,25,15,30,1200,1440,1800,800], unit: "20 ft Dry" },
  "2": { data: [50,40,10,10,30,10,10,1200,1440,2600,360], unit: "40 ft High Cube" },
  "3": { data: [400,20,35,35,25,20,150], unit: "20 ft Dry" },
  "4": { data: [620,25,40,10,30,20,190], unit: "40 ft High Cube" }
};

const NAMES = {
  "1": ["DTHC Fee","Admin Fee","Delivery Order","PORT FEE","Equipment Recovery Handling","Service","Documentation Fee","Nafeza","Agency","Container fee","Cleaning Fee"],
  "2": ["DTHC Fee","Admin Fee","Delivery Order","PORT FEE","Equipment Recovery Handling","Service","Documentation Fee","Nafeza","Agency","Container fee","Cleaning Fee"],
  "3": ["DTHC Fee","Agency","Container Inspection Charges","container cleaning","Admin Fee","Delivery Order ","Nafeza"],
  "4": ["DTHC Fee","Agency","Container Inspection Charges","container cleaning","Admin Fee","Delivery Order ","Nafeza"]
};

function buildDataset(blData) {
  return DATASETS[blData.type].data.map((val, idx) => ({ val, type: blData.type, idx, qty: blData.qty }));
}

function addRowAndWait() {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => { obs.disconnect(); reject(new Error('Row add timeout')); }, 8000);
    const obs = new MutationObserver(mutations => {
      for (const m of mutations)
        for (const node of m.addedNodes)
          if (node.nodeType === 1 && node.classList.contains('nested-fields')) {
            clearTimeout(timer); obs.disconnect();
            setTimeout(() => resolve(node), 350); return;
          }
    });
    obs.observe(document.body, { childList: true, subtree: true });
    const btn = document.querySelector('.add_fields[data-billed-to="client"]') || document.querySelector('.add_fields');
    btn.click();
  });
}

async function fillCosts(blData, chargeMap) {
  for (const item of buildDataset(blData)) {
    const row = await addRowAndWait();
    const nameSelect = row.querySelectorAll('select')[0];
    const charge = chargeMap[NAMES[item.type][item.idx].toLowerCase().trim()];
    if (charge) {
      nameSelect.appendChild(new Option(charge.text, charge.id, true, true));
      nameSelect.value = charge.id;
      nameSelect.dispatchEvent(new Event('change', { bubbles: true }));
      if (window.$) $(nameSelect).trigger('change');
    }
    let unit = DATASETS[item.type].unit;
    if ((item.type === '3' || item.type === '4') && item.idx >= 4) unit = 'Bill of Lading';
    const unitSelect = row.querySelectorAll('select')[2];
    unitSelect.appendChild(new Option(unit, unit, true, true));
    if (window.$) $(unitSelect).trigger('change');
    let currency = 'USD';
    if (item.type === '3' || item.type === '4') { if (item.idx === DATASETS[item.type].data.length - 1) currency = 'EGP'; }
    else { if (item.idx >= 7) currency = 'EGP'; }
    const currencySelect = row.querySelectorAll('select')[1];
    currencySelect.appendChild(new Option(currency, currency, true, true));
    if (window.$) $(currencySelect).trigger('change');
    let q = item.qty;
    if (item.type === '1' || item.type === '2') { if ([1,2,5,7].includes(item.idx)) q = 1; }
    if (item.type === '3' || item.type === '4') { if (item.idx >= 4) q = 1; }
    const quantity = row.querySelector('.amount');
    const selling  = row.querySelector('.selling-rate');
    const cost     = row.querySelector('.cost');
    if (quantity) { quantity.value = q;        triggerAll(quantity); }
    if (selling)  { selling.value  = item.val; triggerAll(selling);  }
    if (cost)     { cost.value     = 0;         triggerAll(cost);     }
  }
}

// ── Invoice Data ──────────────────────────────────────────────────────────────

const INVOICE_PLANS = {
  ALEX: [
    { currency: 'USD', boxIndices: [0,1,2,3,4],   num: 1 },
    { currency: 'USD', boxIndices: [5,6],          num: 2 },
    { currency: 'EGP', boxIndices: [7,8,9,10],    num: 3 }
  ],
  SOKHNA: [
    { currency: 'USD', boxIndices: [0,1,2,3,4,5], num: 1 },
    { currency: 'EGP', boxIndices: [6],            num: 2 }
  ]
};

// ── Invoice Helpers ───────────────────────────────────────────────────────────

async function selectByText(selectId, optionText) {
  const select = document.querySelector(`#${selectId}`);
  if (!select) throw new Error(`Select not found: ${selectId}`);
  const option = Array.from(select.options).find(o => o.text.trim() === optionText);
  if (!option) throw new Error(`Option "${optionText}" not found in ${selectId}. Available: ${Array.from(select.options).map(o => o.text).join(', ')}`);
  select.value = option.value;
  select.dispatchEvent(new Event('change', { bubbles: true }));
  if (window.$) {
    $(select).trigger('change');
    try { $(select).trigger({ type: 'select2:select', params: { data: { id: option.value, text: option.text.trim() } } }); } catch(e) {}
  }
  await sleep(600);
}

function findPDFUrl() {
  for (const sel of ['a[href$=".pdf"]', 'a[href*="download"]', 'a[href*="print"]', '.download-pdf']) {
    const el = document.querySelector(sel);
    if (el && el.href) return el.href;
  }
  const url = window.location.href;
  if (/\/invoices\/\d+/.test(url)) return url.split('?')[0].replace(/\/(edit|show)?$/, '') + '.pdf';
  return null;
}

async function sendDownload(url, filename) {
  return new Promise(resolve => {
    chrome.runtime.sendMessage({ action: 'download', url, filename }, () => resolve());
  });
}

// ── Pre-Steps: Edit Dates + Add Container ─────────────────────────────────────

async function doPreSteps(state) {
  const blData = state.blList[state.currentIndex];

  // ── 1. Shipment Details → Edit Dates → set ETA to today ──────────────────
  showOverlay(`📅 Setting ETA for ${blData.bl}...`);

  const detailsTab = document.querySelector('a[href="#details-tab"]');
  if (detailsTab) { detailsTab.click(); await sleep(1500); }

  const editDatesBtn = await waitFor('a[href*="edit_job_file_dates"]', 8000);
  editDatesBtn.click();
  await sleep(2500);

  // Find the ETA input (not ETA POL) by its label text
  let etaInput = null;
  for (const label of document.querySelectorAll('label')) {
    if (label.textContent.trim() === 'ETA') {
      const forId = label.getAttribute('for');
      etaInput = forId ? document.getElementById(forId) : label.parentElement?.querySelector('input[type="text"], input[type="date"]');
      if (etaInput) break;
    }
  }
  // Fallback selectors
  if (!etaInput) etaInput = document.querySelector('input[id$="_eta"]:not([id*="pol"]):not([id*="etd"])');

  // Convert state.etaDate (YYYY-MM-DD) to DD/MM/YYYY for the logistaas field
  const [etaYr, etaMo, etaDy] = (state.etaDate || '').split('-');
  const formattedDate = (etaDy && etaMo && etaYr) ? `${etaDy}/${etaMo}/${etaYr}` : null;

  // Reliably set a date input value — skips calendar entirely, types directly
  async function setDateValue(input) {
    if (!formattedDate) return;
    // Close any open calendar first
    document.body.click();
    await sleep(300);
    // Focus and set value directly
    input.focus();
    // Use native setter to bypass any framework interception
    const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
    if (nativeSetter) nativeSetter.call(input, formattedDate);
    input.value = formattedDate;
    // Trigger all events so the framework registers the change
    ['input', 'change', 'blur', 'keyup'].forEach(e =>
      input.dispatchEvent(new Event(e, { bubbles: true }))
    );
    await sleep(300);
    // Close any calendar that may have opened
    document.body.click();
    await sleep(300);
  }

  // Set ETA if empty
  const etaDateInput = document.querySelector('input.eta-date') || etaInput;
  if (etaDateInput && !etaDateInput.value.trim()) {
    await setDateValue(etaDateInput);
  }

  // Set Arrival Date if empty
  const arrivalInput = document.querySelector('input.arrival-date');
  if (arrivalInput && !arrivalInput.value.trim()) {
    await setDateValue(arrivalInput);
  }

  // Save dates
  const dateSaveBtn = document.querySelector('input.job-date-save-button')
                   || document.querySelector('input[value="Save"].job-date-save-button');
  if (dateSaveBtn) { dateSaveBtn.click(); await sleep(2000); }

  // Set Destination Country to Egypt if empty (still on Shipment Details tab)
  const countryToEl = document.getElementById('select2-shipment_country_to-container');
  const countryFilled = countryToEl?.getAttribute('title')?.trim();
  if (!countryFilled) {
    showOverlay(`🌍 Setting destination country for ${blData.bl}...`);
    await selectByText('shipment_country_to', 'Egypt');
    await sleep(500);
  }

  // ── 2. MB/L tab → Shipping Line: fill AVANA if empty ────────────────────
  showOverlay(`🚢 Checking shipping line for ${blData.bl}...`);

  const mblTab = document.querySelector('a[href="#master-tab"]');
  if (mblTab) { mblTab.click(); await sleep(1500); }

  // Check if Shipping Line already filled using the underlying select value
  // (title attribute is unreliable — always null even when filled)
  const shippingLineSelect = document.querySelector(
    'select[id="shipment_master_shipment_attributes_sea_shipment_details_attributes_delivery_agent_id"]'
  );
  const shippingLineFilled = shippingLineSelect && shippingLineSelect.value !== '';

  if (!shippingLineFilled) {
    // Click the Select2 visual container to open the dropdown
    const deliveryAgentEl = document.getElementById(
      'select2-shipment_master_shipment_attributes_sea_shipment_details_attributes_delivery_agent_id-container'
    );
    const select2Selection = deliveryAgentEl?.closest('.select2-container')?.querySelector('.select2-selection');

    if (select2Selection) {
      select2Selection.click();
      await sleep(800);

      // Target the search field inside the NOW-OPEN dropdown specifically
      const sf = document.querySelector('.select2-container--open .select2-search__field')
               || document.querySelector('.select2-search__field');

      if (sf) {
        sf.value = 'AVANA';
        sf.dispatchEvent(new Event('input', { bubbles: true }));
        sf.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'a' }));
        if (window.$) $(sf).trigger('input');

        // Wait for AJAX results to load
        await sleep(2500);

        const avanaOpt = [...document.querySelectorAll('.select2-results__option')]
          .find(o => o.textContent.trim().toUpperCase().includes('AVANA'));
        if (avanaOpt) { avanaOpt.click(); await sleep(500); }
        else { document.body.click(); } // Close dropdown if no result found
      }
    }
  }

  // ── 3. Cargo & Equipment → Add Container → Save empty ────────────────────
  showOverlay(`📦 Adding container for ${blData.bl}...`);

  // Find Cargo & Equipment tab by its text content
  const cargoTab = [...document.querySelectorAll('[role="tab"]')]
    .find(el => el.textContent.includes('Cargo') && el.textContent.includes('Equipment'));
  if (cargoTab) { cargoTab.click(); await sleep(2000); }

  // Only add a container if none exist yet (Containers (0) means empty)
  const containerLegend = [...document.querySelectorAll('legend')]
    .find(l => /Containers\s*\(0\)/.test(l.textContent));

  if (containerLegend) {
    const addContainerBtn = await waitFor('a[href*="/equipment/new"]', 8000);
    addContainerBtn.click();
    await sleep(10000);
    const containerSaveBtn = document.querySelector('.reveal input[value="Save"]')
                          || document.querySelector('input[value="Save"]:not(.job-date-save-button)');
    if (!containerSaveBtn) throw new Error('Container Save button not found');
    containerSaveBtn.click();
    await sleep(2000);
  } else {
    // Customer service already added containers — skip this step
    showOverlay(`📦 Containers already present for ${blData.bl}, skipping...`);
    await sleep(500);
  }
}

// ── Cost Fill State Handlers ──────────────────────────────────────────────────

async function handleSearchPage(state) {
  const blData = state.blList[state.currentIndex];
  showOverlay(`🔍 Searching ${blData.bl} (${state.currentIndex + 1}/${state.total})...`);
  try {
    await sleep(1200);
    document.querySelectorAll('a, button').forEach(el => {
      const path = el.querySelector('svg path');
      if (path && (path.getAttribute('d') || '').startsWith('M416 208')) el.click();
    });
    await sleep(600);
    const searchInput = await waitFor('#filterrific_with_doc_no', 8000);
    searchInput.value = '';
    await sleep(200);
    searchInput.value = blData.bl;
    ['input', 'change', 'keyup'].forEach(e => searchInput.dispatchEvent(new Event(e, { bubbles: true })));
    await sleep(400);
    await chrome.storage.local.set({ blAutomation: { ...state, status: 'CLICKING_RESULT' } });
    const searchBtn = document.querySelector('input[value="Search"][type="submit"]');
    if (searchBtn) searchBtn.click();
    else if (searchInput.form) searchInput.form.submit();
  } catch(e) { await saveError(state, blData.bl, e.message); }
}

async function handleClickResult(state) {
  const blData = state.blList[state.currentIndex];
  showOverlay(`🔍 Opening result for ${blData.bl}...`);
  try {
    await sleep(1500);
    let clicked = false;
    const links = [...document.querySelectorAll('a[href*="/shipments/"]')].filter(a => /\/shipments\/\d+/.test(a.href));
    if (links.length > 0) { links[0].click(); clicked = true; }
    if (!clicked) {
      const card = document.querySelector('.copy-card-data');
      if (card) {
        let el = card;
        while (el && el !== document.body) {
          if (el.tagName === 'A' && /\/shipments\/\d+/.test(el.href)) { el.click(); clicked = true; break; }
          if (el.dataset?.href && /\/shipments\/\d+/.test(el.dataset.href)) { window.location.href = el.dataset.href; clicked = true; break; }
          el = el.parentElement;
        }
        if (!clicked) { card.click(); clicked = true; }
      }
    }
    if (!clicked) throw new Error('Could not find shipment result to click.');
    await chrome.storage.local.set({ blAutomation: { ...state, status: 'ON_SHIPMENT_PAGE' } });
  } catch(e) { await saveError(state, blData.bl, e.message); }
}

async function handleShipmentPage(state) {
  const blData = state.blList[state.currentIndex];
  try {
    await sleep(1500);

    // Run pre-steps: edit dates + add container
    await doPreSteps(state);

    // Now go to Costs & Revenues
    showOverlay(`📋 Opening Costs & Revenues for ${blData.bl}...`);
    const costsTab = await waitFor('a[href="#costs-revenues-tab"]', 10000);
    costsTab.click();
    await sleep(3000);
    await waitFor('.add_fields[data-billed-to="client"]', 15000);
    await sleep(300);

    await chrome.storage.local.set({ blAutomation: { ...state, status: 'FILLING', shipmentUrl: window.location.href } });
    await handleFilling({ ...state, status: 'FILLING', shipmentUrl: window.location.href });
  } catch(e) { await saveError(state, blData.bl, e.message); }
}

async function handleFilling(state) {
  const blData = state.blList[state.currentIndex];
  showOverlay(`✍️ Filling costs for ${blData.bl} (${state.currentIndex + 1}/${state.total})...`);
  try {
    const res = await fetch('/data/service_charge_search?service=sea_inbound');
    const charges = await res.json();
    const chargeMap = {};
    charges.forEach(c => { chargeMap[c.text.toLowerCase().trim()] = c; });
    await waitFor('.add_fields[data-billed-to="client"]', 12000);
    await sleep(400);
    await fillCosts(blData, chargeMap);
    showOverlay(`💾 Saving ${blData.bl}...`);
    await sleep(500);
    const saveBtn = document.querySelector('input[value="Save"][type="submit"]');
    if (!saveBtn) throw new Error('Save button not found');
    await chrome.storage.local.set({ blAutomation: { ...state, status: 'SAVING' } });
    saveBtn.click();
  } catch(e) { await saveError(state, blData.bl, e.message); }
}

// ── Invoice State Handlers ────────────────────────────────────────────────────

async function handleAfterSave(state) {
  const newState = { ...state, status: 'INVOICING', invoiceStep: 0, shipmentUrl: window.location.href };
  await chrome.storage.local.set({ blAutomation: newState });
  await sleep(500);
  await handleInvoicing(newState);
}

async function handleInvoicing(state) {
  const blData = state.blList[state.currentIndex];
  const port = (blData.type === '1' || blData.type === '2') ? 'ALEX' : 'SOKHNA';
  const plan = INVOICE_PLANS[port];

  for (let i = state.invoiceStep; i < plan.length; i++) {
    const inv = plan[i];
    showOverlay(`📄 Invoice ${inv.num}/${plan.length} (${inv.currency}) for ${blData.bl}...`);
    try {
      await waitFor('button[data-toggle="invoice-opts"]', 10000);
      await sleep(500);
      document.querySelector('button[data-toggle="invoice-opts"]').click();
      await sleep(700);

      const newInvLink = await waitFor('a.new-invoice[data-inv-type="client"]', 6000);
      newInvLink.click();
      await sleep(4000);

      // ── Step 1: Tax Scheme + Currency → click Next ───────────────────────
      await waitFor('#inv_tax_scheme_id', 8000);
      await sleep(400);
      await selectByText("inv_tax_scheme_id", "With VAT");
      await sleep(500);
      await selectByText("inv_currency", inv.currency);
      await sleep(500);

      const refInputStep1 = document.querySelector('#invoice_external_ref');
      if (refInputStep1 && state.reference) {
        refInputStep1.value = state.reference;
        refInputStep1.dispatchEvent(new Event('input', { bubbles: true }));
      }

      const nextBtn = document.querySelector('input[value="Next"][type="submit"]')
                   || document.querySelector('input[name="commit"]');
      if (!nextBtn) throw new Error('"Next" button not found on step 1');
      nextBtn.click();
      await sleep(3000);

      // ── Step 2: Item checkboxes appear → check right boxes ───────────────
      let boxes = [];
      const pollStart = Date.now();
      while (Date.now() - pollStart < 20000) {
        boxes = [...document.querySelectorAll('.inv-item-select')];
        if (boxes.length > 0) break;
        await sleep(500);
      }
      if (boxes.length === 0) throw new Error('Invoice item checkboxes did not appear on step 2');
      await sleep(300);

      inv.boxIndices.forEach(idx => {
        if (boxes[idx]) {
          boxes[idx].checked = true;
          ['change', 'input'].forEach(e => boxes[idx].dispatchEvent(new Event(e, { bubbles: true })));
        }
      });

      const createBtnEl = document.querySelector('.create-new-invoice');
      if (createBtnEl) { createBtnEl.disabled = false; createBtnEl.removeAttribute('disabled'); }

      const refInputStep2 = document.querySelector('#invoice_external_ref');
      if (refInputStep2 && state.reference && !refInputStep2.value) {
        refInputStep2.value = state.reference;
        refInputStep2.dispatchEvent(new Event('input', { bubbles: true }));
      }

      const draftRadio = document.querySelector('#invoice_status_draft');
      if (draftRadio) { draftRadio.checked = true; draftRadio.dispatchEvent(new Event('change', { bubbles: true })); }

      const logoCheck = document.querySelector('#invoice_extra_fields_with_logo');
      if (logoCheck) { logoCheck.checked = true; logoCheck.dispatchEvent(new Event('change', { bubbles: true })); }

      await sleep(500);

      await chrome.storage.local.set({ blAutomation: { ...state, invoiceStep: i + 1 } });

      chrome.runtime.sendMessage({
        action: 'setNextDownloadPath',
        filename: `${state.batchFolder}/${blData.bl}/invoice_${inv.num}_${inv.currency}.pdf`
      });
      await sleep(200);

      const urlBefore = window.location.href;
      const submitBtn = document.querySelector('.create-new-invoice')
                     || document.querySelector('input[value="Create"][type="submit"]')
                     || document.querySelector('input[name="commit"]');
      if (!submitBtn) throw new Error('Final create button not found on step 2');
      submitBtn.click();

      await sleep(3000);

      if (window.location.href !== urlBefore) return; // Page redirected → invoice page handler takes over

      await tryDownloadFromInvoiceList(blData, inv, state.batchFolder);

    } catch(e) {
      await saveError({ ...state, invoiceStep: i }, blData.bl, `Invoice ${inv.num}: ${e.message}`);
      return;
    }
  }

  await goToNextBL(state);
}

async function handleInvoicePageRedirect(state) {
  const blData = state.blList[state.currentIndex];
  const port = (blData.type === '1' || blData.type === '2') ? 'ALEX' : 'SOKHNA';
  const plan = INVOICE_PLANS[port];
  const completedInv = plan[state.invoiceStep - 1];
  if (completedInv) {
    showOverlay(`💾 Saving invoice ${completedInv.num} PDF for ${blData.bl}...`);
    await sleep(800);
    const pdfUrl = findPDFUrl();
    if (pdfUrl) {
      await sendDownload(pdfUrl, `${state.batchFolder}/${blData.bl}/invoice_${completedInv.num}_${completedInv.currency}.pdf`);
      await sleep(2000);
    }
  }
  if (state.shipmentUrl) window.location.href = state.shipmentUrl;
  else window.history.back();
}

async function tryDownloadFromInvoiceList(blData, inv, batchFolder) {
  await sleep(1000);
  const invoiceLinks = [...document.querySelectorAll('a[href*="/invoices/"]')]
    .filter(a => /\/invoices\/\d+/.test(a.href) && !a.href.includes('/prepare') && !a.href.includes('/new'));
  if (invoiceLinks.length > 0) {
    const pdfUrl = invoiceLinks[invoiceLinks.length - 1].href.split('?')[0] + '.pdf';
    await sendDownload(pdfUrl, `${batchFolder}/${blData.bl}/invoice_${inv.num}_${inv.currency}.pdf`);
    await sleep(1500);
  }
}

async function goToNextBL(state) {
  const nextIndex = state.currentIndex + 1;
  if (nextIndex >= state.total) {
    await chrome.storage.local.set({ blAutomation: { ...state, status: 'DONE', currentIndex: nextIndex } });
    showOverlay(`✅ All ${state.total} BLs completed!`, '#2e7d32');
    return;
  }
  const nextBL = state.blList[nextIndex];
  showOverlay(`➡️ Next: ${nextBL.bl} (${nextIndex + 1}/${state.total})...`);
  await chrome.storage.local.set({
    blAutomation: { ...state, status: 'SEARCHING', currentIndex: nextIndex, invoiceStep: 0, shipmentUrl: null }
  });
  await sleep(800);
  window.location.href = 'https://internal-company-website/shipments';
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  await sleep(500);
  const result = await chrome.storage.local.get(['blAutomation']);
  const state = result.blAutomation;
  if (!state || ['IDLE','DONE','ERROR'].includes(state.status)) return;

  const url = window.location.href;
  const onSearchPage   = url.includes('/shipments') && !url.match(/\/shipments\/\d+/);
  const onShipmentPage = !!url.match(/\/shipments\/\d+/);
  const onInvoicePage  = !!url.match(/\/invoices\/\d+/);

  if      (state.status === 'SEARCHING'        && onSearchPage)   await handleSearchPage(state);
  else if (state.status === 'CLICKING_RESULT'  && onSearchPage)   await handleClickResult(state);
  else if (state.status === 'ON_SHIPMENT_PAGE' && onShipmentPage) await handleShipmentPage(state);
  else if (state.status === 'FILLING'          && onShipmentPage) await handleShipmentPage(state);
  else if (state.status === 'SAVING'           && onShipmentPage) await handleAfterSave(state);
  else if (state.status === 'SAVING'           && onSearchPage)   await goToNextBL(state);
  else if (state.status === 'INVOICING'        && onShipmentPage) await handleInvoicing(state);
  else if (state.status === 'INVOICING'        && onInvoicePage)  await handleInvoicePageRedirect(state);
}

main();
