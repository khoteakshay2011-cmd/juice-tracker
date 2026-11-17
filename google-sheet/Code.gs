/**
 * Juice Stall Tracker — Google Sheet backend.
 * Paste this whole file into Extensions → Apps Script (replace everything),
 * then Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 */

const DEFAULT_MENU = [
  ['amla',   'Amla juice',     25, '#7a9a1e', true],
  ['karela', 'Karela juice',   25, '#2f6b3a', true],
  ['beet',   'Beetroot juice', 25, '#a3214f', true],
  ['carrot', 'Carrot juice',   25, '#e0702a', true],
  ['bhel',   'Sprouts bhel',   35, '#b07b3a', true],
];
const PALETTE = ['#7a9a1e','#2f6b3a','#a3214f','#e0702a','#b07b3a','#3d6fb6','#7b4fa8','#c49a12','#0f8a8a','#d0485e'];

const SALES_HEAD = ['Date','Time','Item','Price ₹','Item ID','Sale ID','Epoch ms','Device'];
const ING_HEAD   = ['Date','Time','Ingredient','Quantity','Unit','Cost ₹','Purchase ID','Epoch ms','Device'];
const MENU_HEAD  = ['ID (leave blank for new)','Name','Price ₹','Colour (optional)','Show? (TRUE/FALSE)'];

/** Run once from the editor (optional — the web app also does this on first use). */
function setup() { ensureSheets_(); }

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('Sales');
  if (!s) {
    s = ss.insertSheet('Sales');
    s.appendRow(SALES_HEAD); s.setFrozenRows(1);
    s.getRange('A:B').setNumberFormat('@');
    s.getRange(1,1,1,SALES_HEAD.length).setFontWeight('bold');
  }
  let i = ss.getSheetByName('Ingredients');
  if (!i) {
    i = ss.insertSheet('Ingredients');
    i.appendRow(ING_HEAD); i.setFrozenRows(1);
    i.getRange('A:B').setNumberFormat('@');
    i.getRange(1,1,1,ING_HEAD.length).setFontWeight('bold');
  }
  let m = ss.getSheetByName('Menu');
  if (!m) {
    m = ss.insertSheet('Menu');
    m.appendRow(MENU_HEAD); m.setFrozenRows(1);
    m.getRange(2,1,DEFAULT_MENU.length,5).setValues(DEFAULT_MENU);
    m.getRange(1,1,1,MENU_HEAD.length).setFontWeight('bold');
  }
  let sum = ss.getSheetByName('Summary');
  if (!sum) {
    sum = ss.insertSheet('Summary');
    sum.getRange('A1').setFormula(`=QUERY(Sales!A2:D, "select A, count(D), sum(D) where A is not null group by A order by A desc label A 'Date', count(D) 'Items sold', sum(D) 'Sales ₹'", 0)`);
    sum.getRange('E1').setFormula(`=QUERY(Ingredients!A2:F, "select A, sum(F) where A is not null group by A order by A desc label A 'Date', sum(F) 'Ingredients ₹'", 0)`);
    sum.getRange('H1').setFormula(`=QUERY(Sales!A2:D, "select A, count(D) where A is not null group by A pivot C order by A desc label A 'Date (qty by item)'", 0)`);
  }
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  return { s, i, m };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function readMenu_(m) {
  const n = m.getLastRow() - 1;
  if (n < 1) return [];
  const rows = m.getRange(2,1,n,5).getValues();
  const out = []; const fix = [];
  rows.forEach((r, ix) => {
    const name = String(r[1] || '').trim();
    if (!name) return;
    let id = String(r[0] || '').trim();
    if (!id) {
      id = name.toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'') || ('item' + (ix+1));
      fix.push([ix + 2, id]);
    }
    const show = r[4] === '' ? true : (r[4] === true || String(r[4]).toUpperCase() === 'TRUE');
    if (!show) return;
    out.push({ id, name, price: Number(r[2]) || 0, color: String(r[3] || '').trim() || PALETTE[ix % PALETTE.length] });
  });
  fix.forEach(([row, id]) => m.getRange(row, 1).setValue(id)); // give new rows a permanent ID
  return out;
}

function readDay_(sh, date, kind) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const w = kind === 'sale' ? 8 : 9;
  const rows = sh.getRange(2,1,n,w).getValues();
  const out = [];
  rows.forEach(r => {
    if (String(r[0]) !== date) return;
    if (kind === 'sale') out.push({ id: String(r[5]), i: String(r[4]), n: String(r[2]), p: Number(r[3]) || 0, t: Number(r[6]) || 0 });
    else out.push({ id: String(r[6]), name: String(r[2]), qty: r[3], unit: String(r[4]), cost: Number(r[5]) || 0, t: Number(r[7]) || 0 });
  });
  return out;
}

function recentIngredientNames_(i) {
  const n = i.getLastRow() - 1;
  if (n < 1) return [];
  const start = Math.max(2, i.getLastRow() - 300);
  const names = i.getRange(start, 3, i.getLastRow() - start + 1, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean).reverse();
  return [...new Set(names)].slice(0, 30);
}

/** GET ?action=state&dates=2026-11-17,2026-11-16 */
function doGet(e) {
  try {
    const { s, i, m } = ensureSheets_();
    const dates = String((e && e.parameter && e.parameter.dates) || '').split(',').filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d)).slice(0, 3);
    const days = {};
    dates.forEach(d => { days[d] = { sales: readDay_(s, d, 'sale'), purchases: readDay_(i, d, 'purchase') }; });
    return json_({ ok: true, menu: readMenu_(m), ingredients: recentIngredientNames_(i), days });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** POST body: {ops:[{op:'sale'|'unsale'|'purchase'|'unpurchase', ...}]} */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const { s, i } = ensureSheets_();
    const body = JSON.parse(e.postData.contents || '{}');
    const ops = Array.isArray(body.ops) ? body.ops.slice(0, 200) : [];
    const dev = String(body.device || '').slice(0, 12);

    const idsIn = (sh, col) => {
      const n = sh.getLastRow() - 1;
      return n < 1 ? [] : sh.getRange(2, col, n, 1).getValues().map(r => String(r[0]));
    };
    let saleIds = idsIn(s, 6), ingIds = idsIn(i, 7);
    const newSales = [], newIng = [], delSale = new Set(), delIng = new Set();

    ops.forEach(o => {
      if (o.op === 'sale' && o.id && !saleIds.includes(String(o.id))) {
        newSales.push([String(o.date), String(o.time), String(o.n).slice(0,60), Number(o.p) || 0, String(o.i).slice(0,40), String(o.id), Number(o.t) || Date.now(), dev]);
        saleIds.push(String(o.id));
      } else if (o.op === 'unsale' && o.id) {
        delSale.add(String(o.id));
      } else if (o.op === 'purchase' && o.id && !ingIds.includes(String(o.id))) {
        newIng.push([String(o.date), String(o.time), String(o.name).slice(0,60), o.qty === '' ? '' : Number(o.qty), String(o.unit).slice(0,12), Number(o.cost) || 0, String(o.id), Number(o.t) || Date.now(), dev]);
        ingIds.push(String(o.id));
      } else if (o.op === 'unpurchase' && o.id) {
        delIng.add(String(o.id));
      }
    });

    if (newSales.length) s.getRange(s.getLastRow() + 1, 1, newSales.length, 8).setValues(newSales);
    if (newIng.length)   i.getRange(i.getLastRow() + 1, 1, newIng.length, 9).setValues(newIng);

    const delRows = (sh, col, set) => {
      if (!set.size) return;
      const ids = idsIn(sh, col);
      for (let r = ids.length - 1; r >= 0; r--) if (set.has(ids[r])) sh.deleteRow(r + 2);
    };
    delRows(s, 6, delSale);
    delRows(i, 7, delIng);

    return json_({ ok: true, applied: ops.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
