/**
 * RTLH Admin — Google Sheet code (orders, expenses, events, stock, quotations & invoices).
 *
 * HOW TO UPDATE (2 minutes, same link as before):
 * 1. Open the "RTLH" Google Sheet (the one with Orders + Expenses) → Extensions → Apps Script.
 * 2. Select everything in Code.gs, delete it, paste this whole file, press Save.
 * 3. Deploy → Manage deployments → ✏️ (pencil) → Version: "New version" → Deploy.
 *    The /exec link stays the same, so the Admin app needs no changes.
 *
 * The first time the app loads, this code adds the tabs it needs (Events, Stock,
 * Moves, Settings) and a few columns at the end of Orders. Nothing existing is removed.
 * An order is "done" when its row is highlighted green — same as you do by hand.
 */
const DONE_GREEN = '#b7e1cd';
const COLS = {
  Orders: ['id', 'date', 'customer', 'product', 'color', 'type', 'text', 'price', 'notes', 'qty', 'kind', 'ref', 'channel'],
  Expenses: ['id', 'date', 'category', 'description', 'amount', 'paidBy', 'settled'],
  Events: ['id', 'name', 'client', 'phone', 'email', 'date', 'emirate', 'venue', 'status', 'lines', 'fees', 'discount', 'notes', 'terms',
    'quoteNo', 'quoteDate', 'invoiceNo', 'invoiceDate', 'paidDate', 'stockDone', 'updated'],
  Stock: ['product', 'color', 'physical', 'reserved', 'incoming', 'low', 'updated'],
  Moves: ['id', 'time', 'kind', 'product', 'color', 'qty', 'physical', 'reserved', 'incoming', 'note', 'ref', 'by'],
  Settings: ['key', 'value'],
};
const TZ = 'Asia/Dubai';

// ---------------------------------------------------------------- helpers
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

/** Returns the tab, creating it and any missing header columns. */
function tab_(name) {
  let sh = ss_().getSheetByName(name);
  if (!sh) { sh = ss_().insertSheet(name); sh.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]).setFontWeight('bold'); sh.setFrozenRows(1); }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  COLS[name].forEach(function (c) {
    if (head.indexOf(c) < 0) { head.push(c); sh.getRange(1, head.length).setValue(c).setFontWeight('bold'); }
  });
  return sh;
}
function head_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }

/** All rows as objects (display values, so dates read "2026-11-22"). */
function rows_(name, withDone) {
  const sh = tab_(name), n = sh.getLastRow();
  if (n < 2) return [];
  const h = head_(sh), w = h.length;
  const vals = sh.getRange(2, 1, n - 1, w).getDisplayValues();
  const bg = withDone ? sh.getRange(2, 1, n - 1, 1).getBackgrounds() : null;
  const out = [];
  vals.forEach(function (r, i) {
    if (r.join('') === '') return;
    const o = {};
    h.forEach(function (k, j) { if (k) o[k] = r[j]; });
    if (withDone) o.done = isGreen_(bg[i][0]);
    out.push(o);
  });
  return out;
}
function findRow_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function writeObj_(sh, row, obj) {
  const h = head_(sh);
  const cur = row > 0 ? sh.getRange(row, 1, 1, h.length).getValues()[0] : h.map(function () { return ''; });
  h.forEach(function (k, j) { if (obj[k] !== undefined) cur[j] = obj[k]; });
  if (row < 0) row = sh.getLastRow() + 1;
  sh.getRange(row, 2).setNumberFormat('@'); // keep dates as typed text
  sh.getRange(row, 1, 1, h.length).setValues([cur]);
  return row;
}
function isGreen_(hex) {
  hex = String(hex || '').toLowerCase();
  if (hex === DONE_GREEN) return true;
  const m = hex.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/);
  if (!m) return false;
  const r = parseInt(m[1], 16), g = parseInt(m[2], 16), b = parseInt(m[3], 16);
  return g > 150 && g - r > 25 && g >= b;
}
function paint_(sh, row, done) {
  sh.getRange(row, 1, 1, sh.getLastColumn()).setBackground(done ? DONE_GREEN : null);
}
function settings_() {
  const o = {};
  rows_('Settings').forEach(function (r) { if (r.key) o[r.key] = r.value; });
  return o;
}

// ---------------------------------------------------------------- read
function doGet(e) {
  // Tiny "has anything changed?" check used by the Admin app every few seconds
  if (e && e.parameter && e.parameter.rev) {
    return json_({ ok: true, rev: PropertiesService.getScriptProperties().getProperty('rev') || '0' });
  }
  try {
    const moves = rows_('Moves');
    return json_({
      ok: true, v: 2,
      orders: rows_('Orders', true),
      expenses: rows_('Expenses'),
      events: rows_('Events'),
      stock: rows_('Stock'),
      moves: moves.slice(-400),
      settings: settings_(),
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------------------------------------------------------------- write
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const out = handle_(d) || {};
    PropertiesService.getScriptProperties().setProperty('rev', String(Date.now()));
    out.ok = true;
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function handle_(d) {
  const a = d.action;
  // ----- orders
  if (a === 'order') { writeObj_(tab_('Orders'), -1, pick_(d, COLS.Orders)); return; }
  if (a === 'orders') { const sh = tab_('Orders'); (d.rows || []).forEach(function (r) { writeObj_(sh, -1, pick_(r, COLS.Orders)); }); return; }
  if (a === 'update_order') return update_('Orders', d);
  if (a === 'delete_order') return remove_('Orders', d.id);
  if (a === 'fulfill' || a === 'reopen') {
    const sh = tab_('Orders'), row = findRow_(sh, d.id);
    if (row < 0) throw new Error('Order not found');
    paint_(sh, row, a === 'fulfill');
    // the other lines of the same multi-item order follow
    const ref = sh.getRange(row, head_(sh).indexOf('ref') + 1).getDisplayValue();
    if (ref) rowsWhere_(sh, 'ref', ref).forEach(function (r) { paint_(sh, r, a === 'fulfill'); });
    return;
  }
  // ----- expenses
  if (a === 'expense') { writeObj_(tab_('Expenses'), -1, pick_(d, COLS.Expenses)); return; }
  if (a === 'update_expense') return update_('Expenses', d);
  if (a === 'delete_expense') return remove_('Expenses', d.id);
  if (a === 'settle' || a === 'unsettle') return update_('Expenses', { id: d.id, settled: a === 'settle' });
  // ----- events, quotations, invoices
  if (a === 'event_save') return saveEvent_(d.event || {});
  if (a === 'event_delete') return remove_('Events', d.id);
  // ----- stock
  if (a === 'stock_moves') return { stock: applyMoves_(d.moves || [], d.by || '') };
  if (a === 'stock_seed') return seed_(d.rows || [], d.by || '');
  if (a === 'stock_low') { const sh = tab_('Stock'), r = stockRow_(sh, d.product, d.color, true); writeObj_(sh, r, { low: d.low }); return; }
  // ----- settings
  if (a === 'settings_save') {
    const sh = tab_('Settings');
    Object.keys(d.settings || {}).forEach(function (k) {
      const row = findRow_(sh, k);
      writeObj_(sh, row, { key: k, value: d.settings[k] });
    });
    return { settings: settings_() };
  }
  throw new Error('Unknown action: ' + a);
}

function pick_(o, keys) { const r = {}; keys.forEach(function (k) { if (o[k] !== undefined) r[k] = o[k]; }); return r; }
function update_(name, d) {
  const sh = tab_(name), row = findRow_(sh, d.id);
  if (row < 0) throw new Error('Not found');
  writeObj_(sh, row, pick_(d, COLS[name]));
}
function remove_(name, id) {
  const sh = tab_(name), row = findRow_(sh, id);
  if (row > 0) sh.deleteRow(row);
}
function rowsWhere_(sh, col, val) {
  const j = head_(sh).indexOf(col), n = sh.getLastRow(), out = [];
  if (j < 0 || n < 2) return out;
  sh.getRange(2, j + 1, n - 1, 1).getDisplayValues().forEach(function (r, i) { if (String(r[0]) === String(val)) out.push(i + 2); });
  return out;
}

function saveEvent_(ev) {
  const sh = tab_('Events');
  if (!ev.id) ev.id = String(Date.now());
  const props = PropertiesService.getScriptProperties();
  const year = Utilities.formatDate(new Date(), TZ, 'yyyy');
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (ev.issue === 'quote' && !ev.quoteNo) {
    const n = Number(props.getProperty('quoteNo') || 0) + 1; props.setProperty('quoteNo', String(n));
    ev.quoteNo = 'Q-' + year + '-' + ('000' + n).slice(-3); ev.quoteDate = today;
  }
  if (ev.issue === 'invoice' && !ev.invoiceNo) {
    const n = Number(props.getProperty('invoiceNo') || 0) + 1; props.setProperty('invoiceNo', String(n));
    ev.invoiceNo = 'INV-' + year + '-' + ('000' + n).slice(-3); ev.invoiceDate = today;
  }
  ['lines', 'fees'].forEach(function (k) { if (ev[k] && typeof ev[k] !== 'string') ev[k] = JSON.stringify(ev[k]); });
  ev.updated = now_();
  writeObj_(sh, findRow_(sh, ev.id), pick_(ev, COLS.Events));
  const saved = rows_('Events').filter(function (r) { return String(r.id) === String(ev.id); })[0];
  return { event: saved };
}

// ---------------------------------------------------------------- stock
function stockRow_(sh, product, color, create) {
  const n = sh.getLastRow();
  if (n >= 2) {
    const v = sh.getRange(2, 1, n - 1, 2).getDisplayValues();
    for (let i = 0; i < v.length; i++) if (v[i][0] === product && v[i][1] === (color || '')) return i + 2;
  }
  if (!create) return -1;
  return writeObj_(sh, -1, { product: product, color: color || '', physical: 0, reserved: 0, incoming: 0, low: '', updated: now_() });
}

/**
 * kinds: sale / event_sale (−physical), restock (+physical), adjust (± physical),
 * count (set physical), incoming (+incoming), receive (incoming → physical),
 * cancel_incoming (−incoming), reserve (+reserved), release (−reserved)
 */
function applyMoves_(moves, by) {
  const sh = tab_('Stock'), mv = tab_('Moves'), h = head_(sh);
  const idx = function (k) { return h.indexOf(k); };
  const props = PropertiesService.getScriptProperties();
  let seen = props.getProperty('seenMoves') || '';
  moves.forEach(function (m) {
    if (m.sub && seen.indexOf(m.sub) >= 0) return; // already applied (double tap / retry)
    const row = stockRow_(sh, m.product, m.color, true);
    const cur = sh.getRange(row, 1, 1, h.length).getValues()[0];
    let p = Number(cur[idx('physical')]) || 0, r = Number(cur[idx('reserved')]) || 0, inc = Number(cur[idx('incoming')]) || 0;
    const q = Number(m.qty) || 0;
    switch (m.kind) {
      case 'sale': case 'event_sale': p -= q; break;
      case 'restock': p += q; break;
      case 'adjust': p += q; break;
      case 'count': p = q; break;
      case 'incoming': inc += q; break;
      case 'receive': inc = Math.max(0, inc - q); p += q; break;
      case 'cancel_incoming': inc = Math.max(0, inc - q); break;
      case 'reserve': r += q; break;
      case 'release': r = Math.max(0, r - q); break;
      default: throw new Error('Unknown stock move: ' + m.kind);
    }
    cur[idx('physical')] = p; cur[idx('reserved')] = r; cur[idx('incoming')] = inc; cur[idx('updated')] = now_();
    sh.getRange(row, 1, 1, h.length).setValues([cur]);
    writeObj_(mv, -1, { id: m.sub || String(Date.now()), time: m.time || now_(), kind: m.kind, product: m.product, color: m.color || '',
      qty: q, physical: p, reserved: r, incoming: inc, note: m.note || '', ref: m.ref || '', by: by });
    if (m.sub) seen = (m.sub + ',' + seen).slice(0, 8500);
  });
  props.setProperty('seenMoves', seen);
  return rows_('Stock');
}

/** Opening stock — only runs while the Stock tab is empty. */
function seed_(rows, by) {
  if (rows_('Stock').length) throw new Error('Stock already has items');
  const sh = tab_('Stock'), mv = tab_('Moves'), t = now_();
  const out = rows.map(function (r) { return [r.product, r.color || '', Number(r.physical) || 0, 0, Number(r.incoming) || 0, r.low === undefined ? '' : r.low, t]; });
  if (out.length) sh.getRange(2, 1, out.length, 7).setValues(out);
  const mh = head_(mv);
  const mrows = rows.map(function (r) {
    const o = { id: 'open-' + r.product + '-' + r.color, time: t, kind: 'opening', product: r.product, color: r.color || '',
      qty: Number(r.physical) || 0, physical: Number(r.physical) || 0, reserved: 0, incoming: Number(r.incoming) || 0, note: 'Opening stock', ref: '', by: by };
    return mh.map(function (k) { return o[k] === undefined ? '' : o[k]; });
  });
  if (mrows.length) mv.getRange(mv.getLastRow() + 1, 1, mrows.length, mh.length).setValues(mrows);
  return { stock: rows_('Stock') };
}
