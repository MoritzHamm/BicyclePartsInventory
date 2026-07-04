/**
 * Bike Parts Inventory — Apps Script backend
 * ------------------------------------------------------------------
 * Data lives in the Google Sheet this script is bound to.
 * The scanning page (hosted elsewhere) calls this deployment's /exec URL.
 * POST uses Content-Type: text/plain so it counts as a "simple" request
 * and skips the CORS preflight that Apps Script web apps can't answer.
 *
 * SETUP (once):
 *   1. Paste this into Extensions > Apps Script of your Sheet.
 *   2. Change SECRET below to a long random string.
 *   3. Run > setup   (grant permissions when asked).
 *   4. Deploy > New deployment > Web app,
 *        Execute as: Me,  Who has access: Anyone.
 *      Copy the /exec URL into index.html.
 */

const SECRET = 'CHANGE-ME-to-a-long-random-string';

const TAB = { products: 'Products', items: 'Items', bikes: 'Bikes' };

/* ------------------------------------------------------------------ setup */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory')
    .addItem('Run setup', 'setup')
    .addToUi();
}

/** Run once from the editor to build tabs, headers and reference formulas. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const products = getOrCreate_(ss, TAB.products);
  setHeaders_(products, ['barcode', 'brand', 'product_name', 'category', 'notes']);

  const bikes = getOrCreate_(ss, TAB.bikes);
  setHeaders_(bikes, ['bike_name', 'odometer_km', 'notes']);

  const items = getOrCreate_(ss, TAB.items);
  setHeaders_(items, ['item_id', 'barcode', 'state', 'current_bike',
                      'install_odo', 'accumulated_km', 'date_added', 'notes']);

  // Reference columns: describe a product once in Products, it shows up here.
  items.getRange('I1').setFormula(
    '=ARRAYFORMULA(IF(ROW(B:B)=1,"brand",IF(B:B="","",' +
    'IFERROR(VLOOKUP(B:B,' + TAB.products + '!A:E,2,FALSE),""))))');
  items.getRange('J1').setFormula(
    '=ARRAYFORMULA(IF(ROW(B:B)=1,"product_name",IF(B:B="","",' +
    'IFERROR(VLOOKUP(B:B,' + TAB.products + '!A:E,3,FALSE),""))))');
  items.getRange('K1').setFormula(
    '=ARRAYFORMULA(IF(ROW(A:A)=1,"current_km",IF(A:A="","",' +
    'F:F+IF(C:C="on_bike",' +
    'IFERROR(VLOOKUP(D:D,' + TAB.bikes + '!A:B,2,FALSE),0)-IF(E:E="",0,E:E),0))))');

  // A starter bike so "Install" works immediately.
  if (lastRowIn_(bikes, 1) < 2) bikes.getRange(2, 1, 1, 3).setValues([['Trail bike', 0, '']]);

  try { SpreadsheetApp.getUi().alert('Setup complete — Products, Items, Bikes are ready.'); }
  catch (e) { /* running without UI is fine */ }
}

/* ----------------------------------------------------------------- routing */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    guard_(p.key);
    switch (p.action) {
      case 'bikes':   return json_({ ok: true, bikes: listBikes_() });
      case 'items':    return json_({ ok: true, items: listItems_() });
      case 'products': return json_({ ok: true, products: listProducts_() });
      case 'product':  return json_({ ok: true, product: lookupProduct_(p.barcode) });
      default:        return json_({ ok: true, msg: 'Bike parts inventory API is live.' });
    }
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (_) {}
  try {
    guard_(body.key);
    switch (body.action) {
      case 'intake':     return json_(intake_(body));
      case 'newProduct': return json_(newProduct_(body));
      case 'install':  return json_(install_(body));
      case 'remove':   return json_(remove_(body));
      case 'retire':   return json_(retire_(body));
      case 'describe': return json_(describe_(body));
      default:         return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function guard_(key) { if (SECRET && key !== SECRET) throw 'unauthorized'; }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ----------------------------------------------------------------- actions */

function intake_(body) {
  const barcode = String(body.barcode || '').trim();   // an EAN, or a synthetic NB-#### key
  if (!barcode) return { ok: false, error: 'no barcode' };
  const count = Math.max(1, Math.min(50, parseInt(body.count, 10) || 1));

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const products = ss.getSheetByName(TAB.products);
    const items = ss.getSheetByName(TAB.items);

    const known = ensureProduct_(products, barcode);   // stub row if new
    const startRow = lastRowIn_(items, 1) + 1;
    let n = nextItemNum_(items);
    const now = new Date(), ids = [], rows = [];
    for (let k = 0; k < count; k++) {
      const id = String(n++).padStart(4, '0');
      ids.push(id);
      rows.push([id, barcode, 'storage', '', '', 0, now, '']);
    }
    items.getRange(startRow, 1, rows.length, 8).setValues(rows);
    return { ok: true, item_id: ids[0], item_ids: ids, productKnown: known.known, product: known.product };
  } finally { lock.releaseLock(); }
}

function newProduct_(body) {
  const products = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.products);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const key = nextManualKey_(products);              // NB-#### synthetic key
    const r = lastRowIn_(products, 1) + 1;
    products.getRange(r, 1, 1, 5).setValues([[
      key, body.brand || '', body.product_name || '', body.category || '', body.notes || ''
    ]]);
    return { ok: true, key: key };
  } finally { lock.releaseLock(); }
}

function install_(body) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.items);
    const r = findRow_(items, 1, body.item_id);
    if (r < 0) return { ok: false, error: 'item not found' };
    const bike = String(body.bike || '').trim();
    if (!bike) return { ok: false, error: 'no bike' };
    const odo = numOr_(body.odometer, getBikeOdo_(bike));
    setBikeOdo_(bike, odo);
    items.getRange(r, 3, 1, 3).setValues([['on_bike', bike, odo]]); // C,D,E
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function remove_(body) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.items);
    const r = findRow_(items, 1, body.item_id);
    if (r < 0) return { ok: false, error: 'item not found' };
    const row = items.getRange(r, 1, 1, 8).getValues()[0];
    const state = row[2], bike = row[3], installOdo = Number(row[4] || 0), accum = Number(row[5] || 0);
    if (state === 'on_bike') {
      const odo = numOr_(body.odometer, getBikeOdo_(bike));
      setBikeOdo_(bike, odo);
      const stint = Math.max(0, odo - installOdo);
      items.getRange(r, 3, 1, 4).setValues([['storage', '', '', accum + stint]]); // C,D,E,F
    } else {
      items.getRange(r, 3).setValue('storage');
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function retire_(body) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.items);
    const r = findRow_(items, 1, body.item_id);
    if (r < 0) return { ok: false, error: 'item not found' };
    const row = items.getRange(r, 1, 1, 8).getValues()[0];
    const state = row[2], bike = row[3], installOdo = Number(row[4] || 0), accum = Number(row[5] || 0);
    let newAccum = accum;
    if (state === 'on_bike') {
      const odo = numOr_(body.odometer, getBikeOdo_(bike));
      setBikeOdo_(bike, odo);
      newAccum = accum + Math.max(0, odo - installOdo);
    }
    items.getRange(r, 3, 1, 4).setValues([['retired', '', '', newAccum]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function describe_(body) {
  const products = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.products);
  const barcode = String(body.barcode || '').trim();
  if (!barcode) return { ok: false, error: 'no barcode' };
  let r = findRow_(products, 1, barcode);
  if (r < 0) r = lastRowIn_(products, 1) + 1;
  products.getRange(r, 1, 1, 5).setValues([[
    barcode, body.brand || '', body.product_name || '', body.category || '', body.notes || ''
  ]]);
  return { ok: true };
}

/* -------------------------------------------------------------- read views */

function listBikes_() {
  const b = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.bikes);
  const last = lastRowIn_(b, 1);
  if (last < 2) return [];
  return b.getRange(2, 1, last - 1, 2).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ name: r[0], odometer: Number(r[1] || 0) }));
}

function listItems_() {
  const items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.items);
  const last = lastRowIn_(items, 1);
  if (last < 2) return [];
  const rows = items.getRange(2, 1, last - 1, 8).getValues();
  const prod = productMap_(), bikeOdo = bikeOdoMap_();
  return rows.filter(r => r[0] !== '').map(r => {
    const [id, barcode, state, bike, installOdo, accum] = r;
    const cur = Number(accum || 0) +
      (state === 'on_bike' ? Math.max(0, Number(bikeOdo[bike] || 0) - Number(installOdo || 0)) : 0);
    const p = prod[String(barcode)] || {};
    return {
      item_id: id, barcode: String(barcode), state, current_bike: bike,
      current_km: Math.round(cur), brand: p.brand || '', product_name: p.name || ''
    };
  });
}

function lookupProduct_(barcode) {
  const p = productMap_()[String(barcode).trim()];
  return p ? { barcode: String(barcode).trim(), brand: p.brand, name: p.name, category: p.category } : null;
}

/* ---------------------------------------------------------------- helpers */

function ensureProduct_(products, barcode) {
  const last = lastRowIn_(products, 1);
  if (last >= 2) {
    const data = products.getRange(2, 1, last - 1, 3).getValues();
    for (const r of data) {
      if (String(r[0]) === barcode) {
        const named = r[1] || r[2];
        return { known: !!named, product: named ? { brand: r[1], name: r[2] } : null };
      }
    }
  }
  products.getRange((last < 1 ? 1 : last) + 1, 1).setValue(barcode); // append stub
  return { known: false, product: null };
}

function nextItemNum_(items) {
  const last = lastRowIn_(items, 1);
  let max = 0;
  if (last >= 2) {
    const ids = items.getRange(2, 1, last - 1, 1).getValues();
    for (const r of ids) {
      const n = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return max + 1;
}

function nextManualKey_(products) {
  const last = lastRowIn_(products, 1);
  let max = 0;
  if (last >= 2) {
    const vals = products.getRange(2, 1, last - 1, 1).getValues();
    for (const r of vals) {
      const m = String(r[0]).match(/^NB-(\d+)$/);
      if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
    }
  }
  return 'NB-' + String(max + 1).padStart(4, '0');
}

function listProducts_() {
  const p = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.products);
  const last = lastRowIn_(p, 1);
  if (last < 2) return [];
  return p.getRange(2, 1, last - 1, 4).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ key: String(r[0]), brand: r[1], name: r[2], category: r[3] }));
}

function productMap_() {
  const p = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.products);
  const last = lastRowIn_(p, 1);
  const map = {};
  if (last >= 2) {
    p.getRange(2, 1, last - 1, 4).getValues().forEach(r => {
      if (r[0] !== '') map[String(r[0])] = { brand: r[1], name: r[2], category: r[3] };
    });
  }
  return map;
}

function bikeOdoMap_() {
  const map = {};
  listBikes_().forEach(b => { map[b.name] = b.odometer; });
  return map;
}

function getBikeOdo_(name) {
  const b = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.bikes);
  const r = findRow_(b, 1, name);
  return r < 0 ? 0 : Number(b.getRange(r, 2).getValue() || 0);
}

function setBikeOdo_(name, odo) {
  const b = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.bikes);
  let r = findRow_(b, 1, name);
  if (r < 0) { r = lastRowIn_(b, 1) + 1; b.getRange(r, 1, 1, 3).setValues([[name, odo, '']]); }
  else if (odo !== '' && odo != null) b.getRange(r, 2).setValue(odo);
}

function getOrCreate_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function setHeaders_(sh, headers) {
  if (sh.getRange(1, 1).getValue() === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function lastRowIn_(sh, col) {
  const vals = sh.getRange(1, col, sh.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (vals[i][0] !== '') return i + 1;
  return 0;
}

function findRow_(sh, col, value) {
  const last = lastRowIn_(sh, col);
  if (last < 2) return -1;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(value)) return i + 2;
  return -1;
}

function numOr_(v, fallback) {
  return (v === '' || v == null || isNaN(Number(v))) ? fallback : Number(v);
}
