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
  let el = document.getElementById('__payoffOverlay_v2__');
  if (!el) {
    el = document.createElement('div');
    el.id = '__payoffOverlay_v2__';
    el.style.cssText = `position:fixed;bottom:20px;right:20px;padding:10px 16px;
      border-radius:8px;color:white;font-size:13px;font-weight:500;z-index:999999;
      box-shadow:0 4px 12px rgba(0,0,0,0.25);max-width:360px;line-height:1.5;
      font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;`;
    document.body.appendChild(el);
  }
  el.style.background = color;
  el.textContent = msg;
}

async function setState(state) {
  await chrome.storage.local.set({ payoffAutomation: state });
}

async function saveError(state, errorMsg) {
  await setState({ ...state, status: 'ERROR', errorMsg });
  showOverlay('Error: ' + errorMsg, '#f44336');
}

// state.shipmentNos is always an array (length 1 for a single shipment, N for a
// batch). state.index is the shipment currently being worked on, shared between
// the checking phase (state.phase === 'checking') and the payoff phase
// (state.phase === 'payoff') — reset to 0 when moving from one phase to the other.

// ── Step: search page → find shipment ──────────────────────────────────────────

async function handleSearchPage(state) {
  const shipmentNo = state.shipmentNos[state.index];
  const total = state.shipmentNos.length;
  const verb = state.phase === 'checking' ? 'Checking' : 'Opening';
  showOverlay(total > 1 ? `${verb} shipment ${state.index + 1}/${total}: ${shipmentNo}...` : `${verb} ${shipmentNo}...`);
  try {
    await sleep(1000);
    document.querySelectorAll('a, button').forEach(el => {
      const path = el.querySelector('svg path');
      if (path && (path.getAttribute('d') || '').startsWith('M416 208')) el.click();
    });
    await sleep(600);
    const searchInput = await waitFor('#filterrific_with_doc_no', 8000);
    searchInput.value = '';
    await sleep(200);
    searchInput.value = shipmentNo;
    ['input', 'change', 'keyup'].forEach(e => searchInput.dispatchEvent(new Event(e, { bubbles: true })));
    await sleep(400);
    await setState({ ...state, status: 'CLICKING_RESULT' });
    const searchBtn = document.querySelector('input[value="Search"][type="submit"]');
    if (searchBtn) searchBtn.click();
    else if (searchInput.form) searchInput.form.submit();
  } catch (e) { await saveError(state, e.message); }
}

async function handleClickResult(state) {
  const shipmentNo = state.shipmentNos[state.index];
  try {
    await sleep(1500);
    const links = [...document.querySelectorAll('a[href*="/shipments/"]')].filter(a => /\/shipments\/\d+/.test(a.href));
    if (links.length === 0) throw new Error(`No shipment result found for ${shipmentNo}`);

    // A single shipment card contains several links to the same shipment (serial
    // number, status icons, the card itself, its menu), so count distinct
    // shipment ids, not links.
    const idOf = a => a.href.match(/\/shipments\/(\d+)/)[1];
    const distinctIds = [...new Set(links.map(idOf))];
    if (distinctIds.length > 1) {
      throw new Error(`Search for "${shipmentNo}" matched ${distinctIds.length} different shipments, not one — stopping rather than guessing which is right. Please check this shipment number manually.`);
    }
    links[0].click();
    await setState({ ...state, status: 'ON_SHIPMENT_PAGE' });
  } catch (e) { await saveError(state, e.message); }
}

// ── Reading the Invoices / Debit Notes modal ────────────────────────────────────

// True only if the element is actually rendered and visible on screen right now
// — excludes anything sitting in a hidden tab panel, collapsed section, or
// closed modal that happens to still be present in the page's HTML.
function isVisible(el) {
  if (!el) return false;
  if (el.offsetParent === null) return false;
  let node = el;
  while (node) {
    if (node.getAttribute && node.getAttribute('aria-hidden') === 'true') return false;
    node = node.parentElement;
  }
  return true;
}

function readInvoiceRows() {
  const candidates = [...document.querySelectorAll('th')]
    .filter(th => th.textContent.trim() === 'Payment Status' && isVisible(th));
  if (candidates.length === 0) throw new Error('Invoice table not found (no visible "Payment Status" column)');
  if (candidates.length > 1) {
    throw new Error(`Found ${candidates.length} visible tables with a "Payment Status" column — expected exactly one, stopping rather than guessing which is the right one.`);
  }
  const headerCell = candidates[0];
  const table = headerCell.closest('table');
  const headers = [...table.querySelectorAll('thead th')].map(th => th.textContent.trim());
  const idx = {
    status: headers.indexOf('Status'),
    currency: headers.indexOf('Currency'),
    amount: headers.indexOf('Amount'),
    paymentStatus: headers.indexOf('Payment Status'),
  };
  return [...table.querySelectorAll('tbody tr')].map(row => {
    const cells = row.querySelectorAll('td');
    const changeToFinalLink = row.querySelector('a[href*="prepare_change_to_final"]');
    const idMatch = changeToFinalLink ? changeToFinalLink.href.match(/\/invoices\/(\d+)\//) : null;
    return {
      status: cells[idx.status]?.textContent.trim() || '',
      currency: cells[idx.currency]?.textContent.trim() || '',
      amount: parseFloat((cells[idx.amount]?.textContent || '0').replace(/,/g, '')) || 0,
      paymentStatus: cells[idx.paymentStatus]?.textContent.trim() || '',
      invoiceId: idMatch ? idMatch[1] : null,
    };
  });
}

function computeBalance(rows) {
  const balance = {};
  rows.filter(r => r.paymentStatus === 'Unpaid').forEach(r => {
    balance[r.currency] = (balance[r.currency] || 0) + r.amount;
  });
  return balance;
}

// ── On the shipment page: either read its balance (checking) or start finalizing
//    it (payoff) — which one depends on state.phase ────────────────────────────

async function handleOnShipmentPage(state) {
  if (state.phase === 'checking') {
    await handleCheckBalance(state);
    return;
  }
  const shipmentNo = state.shipmentNos[state.index];
  const toFinalize = (state.shipmentInvoices && state.shipmentInvoices[shipmentNo]) || [];
  const newState = { ...state, status: 'FINALIZING', finalizeIndex: 0, toFinalize };
  await setState(newState);
  await handleFinalizing(newState);
}

async function handleCheckBalance(state) {
  const shipmentNo = state.shipmentNos[state.index];
  const total = state.shipmentNos.length;
  try {
    await sleep(1500);
    showOverlay(total > 1 ? `Reading balance for ${shipmentNo} (${state.index + 1}/${total})...` : `Reading balance for ${shipmentNo}...`);
    const invBtn = await waitFor('button[data-toggle="invoice-opts"]', 10000);
    invBtn.click();
    await sleep(600);
    const modalLink = await waitFor('a[data-ajax-open="large-modal"][href*="/invoices"]', 6000);
    modalLink.click();
    await sleep(1500);
    const rows = readInvoiceRows();
    const balance = computeBalance(rows);
    const toFinalize = rows.filter(r => r.paymentStatus === 'Unpaid' && r.status === 'Draft');
    const unpaidLines = rows
      .filter(r => r.paymentStatus === 'Unpaid')
      .map(r => ({ currency: r.currency, amount: r.amount, status: r.status, invoiceId: r.invoiceId }));

    const shipmentBalances = { ...state.shipmentBalances, [shipmentNo]: balance };
    const shipmentInvoices = { ...state.shipmentInvoices, [shipmentNo]: toFinalize };
    const shipmentUnpaidLines = { ...state.shipmentUnpaidLines, [shipmentNo]: unpaidLines };
    const nextIndex = state.index + 1;

    if (nextIndex < total) {
      await setState({ ...state, status: 'SEARCHING', index: nextIndex, shipmentBalances, shipmentInvoices, shipmentUnpaidLines });
      window.location.href = 'https://internal-company-website/shipments';
      return;
    }

    // Last shipment checked — combine every shipment's balance into one total
    // per currency before pausing for confirmation. Nothing gets touched until
    // this combined total is confirmed as covered.
    const aggregate = {};
    Object.values(shipmentBalances).forEach(b => {
      Object.keys(b).forEach(cur => { aggregate[cur] = (aggregate[cur] || 0) + b[cur]; });
    });

    await setState({
      ...state,
      status: 'AWAITING_CONFIRM',
      shipmentBalances,
      shipmentInvoices,
      shipmentUnpaidLines,
      balance: aggregate,
    });
    showOverlay('Balance ready — check the extension popup to confirm.', '#2f6fed');
  } catch (e) { await saveError(state, e.message); }
}

// ── Step: FINALIZING → loop through the current shipment's unpaid draft invoices

async function handleFinalizing(state) {
  const shipmentNo = state.shipmentNos[state.index];
  const list = state.toFinalize || [];
  const i = state.finalizeIndex || 0;
  if (i >= list.length) {
    await setState({ ...state, status: 'PAYMENTS_SEARCH' });
    window.location.href = 'https://internal-company-website/invoices';
    return;
  }
  const inv = list[i];
  try {
    const prefix = state.shipmentNos.length > 1 ? `[${state.index + 1}/${state.shipmentNos.length} ${shipmentNo}] ` : '';
    showOverlay(`${prefix}Finalizing invoice ${i + 1}/${list.length} (${inv.currency})...`);
    let link = document.querySelector(`a[href*="/invoices/${inv.invoiceId}/prepare_change_to_final"]`);
    if (!link) {
      // Fresh page load — the modal isn't open yet, so open it first
      const invBtn = await waitFor('button[data-toggle="invoice-opts"]', 10000);
      invBtn.click();
      await sleep(500);
      const modalLink = await waitFor('a[data-ajax-open="large-modal"][href*="/invoices"]', 6000);
      modalLink.click();
      await sleep(1200);
      link = await waitFor(`a[href*="/invoices/${inv.invoiceId}/prepare_change_to_final"]`, 6000);
    }
    link.click();
    await sleep(1000);
    // The row link opens a small modal with date preferences (left at their
    // default — today's date — same as your normal process). The actual
    // generation is triggered by the "Generate" button inside it, which keeps
    // a stable id regardless of which invoice is open.
    const generateBtn = await waitFor('#link-change-to-final', 6000);
    generateBtn.click();
    await sleep(2500);

    await setState({ ...state, status: 'FINALIZING', finalizeIndex: i + 1 });
    await sleep(500);
    // Reopening the invoice modal in place on the same page load wasn't
    // reliable — a fresh page load before the next invoice is what actually
    // works, so we reload here rather than trying to reuse the current page.
    window.location.reload();
  } catch (e) { await saveError(state, e.message); }
}

// ── Step: PAYMENTS_SEARCH → search by Shipment Master / House No. ──────────────

async function handlePaymentsSearch(state) {
  const shipmentNo = state.shipmentNos[state.index];
  try {
    await sleep(1500);
    let searchInput = document.querySelector('#filterrific_with_shipment_doc_no');
    if (!searchInput) {
      document.querySelectorAll('a, button').forEach(el => {
        const path = el.querySelector('svg path');
        if (path && (path.getAttribute('d') || '').startsWith('M416 208')) el.click();
      });
      await sleep(600);
      const searchByContainer = document.querySelector('[id^="select2-search-filter-select-container"]');
      const select2Selection = searchByContainer?.closest('.select2-container')?.querySelector('.select2-selection');
      if (select2Selection) {
        select2Selection.click();
        await sleep(600);
        // Matching by visible text since the option id includes a session-specific
        // random fragment (e.g. "...-uj3k-shipment_doc_no") that isn't stable.
        const option = [...document.querySelectorAll('.select2-results__option')]
          .find(o => o.textContent.trim() === 'Shipment Master / House No.');
        if (option) { option.click(); await sleep(500); }
      }
      searchInput = await waitFor('#filterrific_with_shipment_doc_no', 8000);
    }
    searchInput.value = '';
    await sleep(200);
    searchInput.value = shipmentNo;
    ['input', 'change', 'keyup'].forEach(e => searchInput.dispatchEvent(new Event(e, { bubbles: true })));
    await sleep(400);
    const searchBtn = document.querySelector('input[value="Search"][type="submit"]');
    await setState({ ...state, status: 'MARKING_PAID' });
    if (searchBtn) searchBtn.click();
    else if (searchInput.form) searchInput.form.submit();
  } catch (e) { await saveError(state, e.message); }
}

// ── Step: MARKING_PAID → click money-bill icon, then fill the form once it loads ─

async function handleMarkingPaid(state) {
  try {
    await sleep(1200);

    // We might land here either on the payment form itself (after clicking a
    // money-bill icon on a previous run navigated us here) or on the payments
    // list (need to find the next unpaid row). Check which one we're actually
    // looking at rather than assuming.
    if (document.querySelector('#payment_payment_method')) {
      await fillPaymentForm(state);
      return;
    }

    let icon = document.querySelector('i.fa-money-bill');
    if (!icon) {
      try {
        await waitFor('i.fa-money-bill', 5000);
        icon = document.querySelector('i.fa-money-bill');
      } catch (_) {
        // Nothing left to pay for this shipment — move to the next one, if any
        await advanceToNextShipmentOrFinish(state);
        return;
      }
    }

    const paidCount = state.paidCount || 0;
    showOverlay(`Opening payment form ${paidCount + 1}...`);
    const link = icon.closest('a') || icon.parentElement;
    link.click();
    // This click navigates to the payment form on its own page — the current
    // script instance ends here, and main() re-enters this same function once
    // that new page loads, where the check above will route to fillPaymentForm.
  } catch (e) { await saveError(state, e.message); }
}

async function advanceToNextShipmentOrFinish(state) {
  const nextIndex = (state.index || 0) + 1;
  if (nextIndex < state.shipmentNos.length) {
    await setState({ ...state, status: 'SEARCHING', phase: 'payoff', index: nextIndex });
    window.location.href = 'https://internal-company-website/shipments';
  } else {
    await setState({ ...state, status: 'DONE' });
    showOverlay('All done — every shipment finalized, paid, and receipted.', '#2e7d32');
  }
}

async function fillPaymentForm(state) {
  try {
    showOverlay(`Recording payment ${(state.paidCount || 0) + 1}...`);
    await sleep(400);
    const methodSelect = document.querySelector('#payment_payment_method');
    methodSelect.value = 'cheque';
    methodSelect.dispatchEvent(new Event('change', { bubbles: true }));
    await sleep(300);

    const bankInput = document.querySelector('#payment_cash_bank_reference, #payment_bank_name');
    if (bankInput) { bankInput.value = state.bankName || ''; triggerAll(bankInput); }

    // Confirmed: "Amount" on this form is a read-only <td>, not an input — the
    // payment always covers the full invoice total, nothing to fill in here.

    await sleep(300);
    const saveBtn = document.querySelector('input[name="save_finalize"][value="Save & Finalize"]');
    if (!saveBtn) throw new Error('"Save & Finalize" button not found');
    saveBtn.click();
    // This also navigates to another page (the one with the Documents dropdown)
    // — main() picks up GENERATE_RECEIPT once that page loads.
    await setState({ ...state, status: 'GENERATE_RECEIPT' });
  } catch (e) { await saveError(state, e.message); }
}

// ── Step: GENERATE_RECEIPT → Documents > RV, check logo, Next, download ───────

async function handleGenerateReceipt(state) {
  try {
    showOverlay('Generating receipt...');
    const docsBtn = await waitFor('button[data-toggle="document-templates"]', 8000);
    docsBtn.click();
    await sleep(600);
    // Matching by href*="model_name=Payment" since the template id and model id
    // are specific to each payment record and will differ every time.
    const rvLink = await waitFor('a[data-ajax-open="small-modal"][href*="model_name=Payment"]', 6000);
    rvLink.click();
    await sleep(1200);

    const logoCheck = await waitFor('#doc_form_with_logo', 6000);
    if (!logoCheck.checked) { logoCheck.checked = true; logoCheck.dispatchEvent(new Event('change', { bubbles: true })); }
    await sleep(300);

    const nextBtn = document.querySelector('input[type="submit"][name="commit"][value="Next"]');
    if (!nextBtn) throw new Error('"Next" button not found on receipt form');
    nextBtn.click();
    await sleep(2000);

    await setState({ ...state, status: 'PAYMENTS_SEARCH', paidCount: (state.paidCount || 0) + 1 });
    await sleep(500);
    // Back to a clean, freshly-searched list before looking for the next unpaid
    // invoice within this same shipment — same reasoning as the reload between
    // finalize steps.
    window.location.href = 'https://internal-company-website/invoices';
  } catch (e) { await saveError(state, e.message); }
}

// ── Main ──────────────────────────────────────────────────────────────────────

let __busy = false;

async function main() {
  if (__busy) return;
  __busy = true;
  try {
    await sleep(500);
    const result = await chrome.storage.local.get(['payoffAutomation']);
    const state = result.payoffAutomation;
    if (!state || ['IDLE', 'DONE', 'ERROR', 'AWAITING_CONFIRM'].includes(state.status)) return;

    const url = window.location.href;
    const onSearchPage = url.includes('/shipments') && !url.match(/\/shipments\/\d+/);
    const onShipmentPage = !!url.match(/\/shipments\/\d+/);
    // Loose on purpose: the payment form's own URL isn't confirmed, so this
    // matches "/invoices" (the list) and "payment" (whatever the form's path is).
    const onPaymentsArea = url.includes('/invoices') || url.includes('payment');

    if      (state.status === 'SEARCHING'        && onSearchPage)   await handleSearchPage(state);
    else if (state.status === 'CLICKING_RESULT'  && onSearchPage)   await handleClickResult(state);
    else if (state.status === 'ON_SHIPMENT_PAGE' && onShipmentPage) await handleOnShipmentPage(state);
    else if (state.status === 'FINALIZING'       && onShipmentPage) await handleFinalizing(state);
    else if (state.status === 'PAYMENTS_SEARCH'  && onPaymentsArea) await handlePaymentsSearch(state);
    else if (state.status === 'MARKING_PAID'     && onPaymentsArea) await handleMarkingPaid(state);
    else if (state.status === 'GENERATE_RECEIPT' && onPaymentsArea) await handleGenerateReceipt(state);
  } finally {
    __busy = false;
  }
}

// Runs once when the page loads (handles every step that involves navigation —
// search results, shipment page, the payments list, etc).
main();

// Also runs whenever the popup changes state without a page navigation (e.g. the
// initial "Confirm & finalize" click, which only updates storage, doesn't reload
// the tab itself).
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'local' && changes.payoffAutomation) {
    main();
  }
});
