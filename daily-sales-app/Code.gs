/**
 * DAILY SHOP SALES  -  Google Apps Script (container-bound to the Google Sheet)
 *
 * Sheets created by setup():
 *   UAE / OMAN : one row per shop per day (rows are pre-created in shop order, blank until the shop posts)
 *   Targets    : admin types the daily target per shop (one row per date, one column per shop code)
 *   Shops      : country / code / name, in the order rows are created
 *   Users      : username / password / shop code / active  (admin managed)
 *   Config     : edit window, back-date days, history days, POS labels, currency
 */

var TZ = 'Asia/Dubai';            // UAE and Oman are both UTC+4
var SHEET_ROWS_TAIL = 600;        // how many latest rows to scan when building a shop's history

// Column numbers (1-based) in UAE / OMAN sheets
var C = { DATE:1, COUNTRY:2, CODE:3, SHOP:4, TOTAL:5, POS1:6, POS2:7, OTHERS:8, QTY:9, INV:10,
          CASH:11, CARD:12, CASHP:13, CARDP:14, POS1P:15, POS2P:16, OTHP:17,
          IPC:18, ATV:19, ASP:20, TARGET:21, ACH:22, SUBMITTED:23, UPDATED:24 };

var SHOP_LIST = [
  ['UAE','TO','TAHERI OKADH'],
  ['UAE','NS','NEW SARA PLAZA'],
  ['UAE','SA','SARA ARABIAN'],
  ['UAE','HS','HUSSAINI SHARJAH'],
  ['UAE','MM','SARA MANAR MALL'],
  ['OMAN','NF','AVENUES MALL'],
  ['OMAN','AK','AL KHOUD'],
  ['OMAN','SB','SEEB MALL'],
  ['OMAN','QCC','QURUM CITY CENTER'],
  ['OMAN','CCM','CITY CENTER MUSCAT'],
  ['OMAN','SCC','SOHAR CITY CENTER'],
  ['OMAN','NQ','NAQIYA FASHION']
];

var CONFIG_DEFAULTS = [
  ['EDIT_MINUTES', 30,       'Minutes a shop can edit after FIRST submission of a day'],
  ['BACKDATE_DAYS', 1,       'How many days back a shop may post (1 = today + yesterday, for the "forgot at night" case)'],
  ['HISTORY_DAYS', 30,       'Days shown in the shop history view'],
  ['UAE_POS1', 'ADIB',       'Label for POS-1 in UAE'],
  ['UAE_POS2', 'PAYMOB',     'Label for POS-2 in UAE'],
  ['OMAN_POS1', 'BOM',       'Label for POS-1 in Oman'],
  ['OMAN_POS2', 'AB',        'Label for POS-2 in Oman'],
  ['UAE_CURRENCY', 'AED',    ''],
  ['OMAN_CURRENCY', 'OMR',   '']
];

/* ------------------------------------------------------------------ */
/*  Web app entry  (JSON API - the screen lives in the Sara Sales site) */
/* ------------------------------------------------------------------ */
// Spreadsheet "Daily Sales". Used if this script is NOT bound to the sheet.
var SHEET_ID = '1wdTqQxkDQqiLVdLDAagM6PzV3VJ3p8fy6Ub-SDO4hlo';

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function api_(p) {
  try {
    switch (String(p.action || '')) {
      case 'ping':   return { ok: true, data: 'Daily Sales API is running' };
      case 'login':  return { ok: true, data: login(p.username, p.password) };
      case 'state':  return { ok: true, data: getState(p.token) };
      case 'save':   return { ok: true, data: saveEntry(p.token, p.entry) };
      case 'logout': return { ok: true, data: logout(p.token) };
      default: throw new Error('Unknown action');
    }
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

// GET is health-check only, so passwords never travel in a URL.
function doGet() { return json_(api_({ action: 'ping' })); }

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) { /* empty body */ }
  return json_(api_(p));
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sales App')
    .addItem('Setup / repair sheets', 'setup')
    .addItem("Create today's rows now", 'dailyCreateRows')
    .addItem('Extend Targets dates (+31 days)', 'extendTargetDates')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */
function ss_() {
  var a = null;
  try { a = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* standalone script */ }
  return a || SpreadsheetApp.openById(SHEET_ID);
}

function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function fmtDate_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }

function addDays_(ds, n) {
  var p = ds.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function dateObj_(ds) { var p = ds.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }

function labelDate_(ds) { return Utilities.formatDate(dateObj_(ds), TZ, 'dd-MMM-yy'); }

function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

function config_() {
  var cfg = {};
  CONFIG_DEFAULTS.forEach(function (r) { cfg[r[0]] = r[1]; });
  var sh = ss_().getSheetByName('Config');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0] !== '' && r[1] !== '') cfg[String(r[0]).trim()] = r[1];
    });
  }
  return {
    editMinutes: Number(cfg.EDIT_MINUTES) || 30,
    backdateDays: cfg.BACKDATE_DAYS === '' ? 1 : Number(cfg.BACKDATE_DAYS),
    historyDays: Number(cfg.HISTORY_DAYS) || 30,
    raw: cfg
  };
}

function shops_() {
  var sh = ss_().getSheetByName('Shops');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return r[0] && r[1] && r[2]; })
    .map(function (r) {
      return { country: String(r[0]).trim().toUpperCase(), code: String(r[1]).trim(), name: String(r[2]).trim() };
    });
}

function shopByCode_(code) {
  var l = shops_();
  for (var i = 0; i < l.length; i++) if (l[i].code === code) return l[i];
  return null;
}

/* ------------------------------------------------------------------ */
/*  One-time setup (run from the "Sales App" menu or the editor)       */
/* ------------------------------------------------------------------ */
function setup() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);

  var order = ['UAE', 'OMAN', 'Targets', 'Shops', 'Users', 'Config'];
  order.forEach(function (n) { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  seedConfig_();
  seedShops_();
  seedUsers_();
  seedTargets_();
  buildDataSheet_('UAE');
  buildDataSheet_('OMAN');
  installTrigger_();

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { ensureRows_(todayStr_()); } finally { lock.release(); }

  try {
    SpreadsheetApp.getUi().alert('Setup complete.\n\nUsers sheet has the generated passwords.\nNow deploy: Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).');
  } catch (e) { /* run from editor without UI */ }
}

function seedConfig_() {
  var sh = ss_().getSheetByName('Config');
  if (sh.getLastRow() > 1) return;
  sh.getRange(1, 1, 1, 3).setValues([['KEY', 'VALUE', 'NOTE']]).setFontWeight('bold').setBackground('#d9d9d9');
  sh.getRange(2, 1, CONFIG_DEFAULTS.length, 3).setValues(CONFIG_DEFAULTS);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(3, 520);
  sh.setFrozenRows(1);
}

function seedShops_() {
  var sh = ss_().getSheetByName('Shops');
  if (sh.getLastRow() > 1) return;
  sh.getRange(1, 1, 1, 3).setValues([['COUNTRY', 'CODE', 'SHOP NAME']]).setFontWeight('bold').setBackground('#d9d9d9');
  sh.getRange(2, 1, SHOP_LIST.length, 3).setValues(SHOP_LIST);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);
}

function seedUsers_() {
  var sh = ss_().getSheetByName('Users');
  if (sh.getLastRow() > 1) return;
  sh.getRange(1, 1, 1, 4).setValues([['USERNAME', 'PASSWORD', 'SHOP CODE', 'ACTIVE (YES/NO)']])
    .setFontWeight('bold').setBackground('#d9d9d9');
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  var rows = SHOP_LIST.map(function (s) {
    var pw = ''; for (var i = 0; i < 6; i++) pw += chars.charAt(Math.floor(Math.random() * chars.length));
    return [s[1].toLowerCase(), pw, s[1], 'YES'];
  });
  sh.getRange(2, 2, 500, 1).setNumberFormat('@');          // keep passwords as text
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 4);
}

function seedTargets_() {
  var sh = ss_().getSheetByName('Targets');
  var codes = shops_().map(function (s) { return s.code; });
  sh.getRange(1, 1, 1, codes.length + 1).setValues([['DATE'].concat(codes)])
    .setFontWeight('bold').setBackground('#fff2cc');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  if (sh.getLastRow() > 1) return;
  var today = todayStr_();
  var rows = [];
  for (var i = 0; i < 31; i++) rows.push([dateObj_(addDays_(today, i))]);
  sh.getRange(2, 1, rows.length, 1).setValues(rows).setNumberFormat('dd-mmm-yy');
  sh.getRange(1, 1).setNote('Admin: type each shop\'s target next to the date. Exact date match is used.');
}

function extendTargetDates() {
  var sh = ss_().getSheetByName('Targets');
  var last = sh.getLastRow();
  var lastDate = last > 1 ? sh.getRange(last, 1).getValue() : null;
  var start = (lastDate instanceof Date) ? addDays_(fmtDate_(lastDate), 1) : todayStr_();
  var rows = [];
  for (var i = 0; i < 31; i++) rows.push([dateObj_(addDays_(start, i))]);
  sh.getRange(last + 1, 1, rows.length, 1).setValues(rows).setNumberFormat('dd-mmm-yy');
}

function buildDataSheet_(country) {
  var sh = ss_().getSheetByName(country);
  var cfg = config_().raw;
  var p1 = cfg[country + '_POS1'], p2 = cfg[country + '_POS2'];
  var headers = ['DATE', 'COUNTRY', 'CODE', 'SHOP',
    'TOTAL SALES VALUE', 'CARD (' + p1 + ') POS', 'CARD (' + p2 + ') POS', 'OTHERS', 'TOTAL QTY', 'TOTAL INVOICES',
    'CASH', 'TOTAL CARD', 'CASH %', 'CARD %', 'POS-1 % OF CARD', 'POS-2 % OF CARD', 'OTHERS %',
    'IPC', 'ATV', 'ASP', 'TARGET', 'ACHVD %', 'FIRST SUBMITTED', 'LAST UPDATED'];
  var h = sh.getRange(1, 1, 1, headers.length);
  h.setValues([headers]).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(1, 1, 1, 4).setBackground('#d9d9d9');
  sh.getRange(1, C.TOTAL, 1, 6).setBackground('#fff2cc');     // inputs entered by the shop
  sh.getRange(1, C.CASH, 1, 12).setBackground('#d9ead3');     // calculated by sheet
  sh.getRange(1, C.SUBMITTED, 1, 2).setBackground('#d9d9d9');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  sh.setRowHeight(1, 42);
  sh.getRange(1, C.SUBMITTED).setNote('Set by the app on first save. Edit window counts from here. Admin: clear this cell to re-open the day for another edit window.');
}

function installTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyCreateRows') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyCreateRows').timeBased().everyDays(1).atHour(0).nearMinute(5).create();
}

/** Time-driven: pre-creates the blank rows for today, in shop order. */
function dailyCreateRows() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { ensureRows_(todayStr_()); } finally { lock.release(); }
}

/* ------------------------------------------------------------------ */
/*  Row creation                                                       */
/* ------------------------------------------------------------------ */
function formulasFor_(r) {
  var E = '$E' + r, F = '$F' + r, G = '$G' + r, H = '$H' + r, I = '$I' + r, J = '$J' + r;
  var K = '$K' + r, L = '$L' + r, U = '$U' + r;
  var tgt = 'INDEX(Targets!$B$2:$Z$2000,MATCH($A' + r + ',Targets!$A$2:$A$2000,0),MATCH($C' + r + ',Targets!$B$1:$Z$1,0))';
  return [
    '=IF(' + E + '="","",' + E + '-' + L + ')',                                          // K cash
    '=IF(' + E + '="","",N(' + F + ')+N(' + G + ')+N(' + H + '))',                        // L total card
    '=IF(' + E + '="","",IF(' + E + '=0,0,' + K + '/' + E + '))',                         // M cash %
    '=IF(' + E + '="","",IF(' + E + '=0,0,' + L + '/' + E + '))',                         // N card %
    '=IF(' + E + '="","",IF(' + L + '=0,0,N(' + F + ')/' + L + '))',                      // O pos1 % of card
    '=IF(' + E + '="","",IF(' + L + '=0,0,N(' + G + ')/' + L + '))',                      // P pos2 % of card
    '=IF(' + E + '="","",IF(' + E + '=0,0,N(' + H + ')/' + E + '))',                      // Q others % of total
    '=IF(OR(' + E + '="",N(' + J + ')=0),"",' + I + '/' + J + ')',                        // R IPC
    '=IF(OR(' + E + '="",N(' + J + ')=0),"",' + E + '/' + J + ')',                        // S ATV
    '=IF(OR(' + E + '="",N(' + I + ')=0),"",' + E + '/' + I + ')',                        // T ASP
    '=IFERROR(IF(' + tgt + '="","",' + tgt + '),"")',                                     // U target
    '=IF(OR(' + E + '="",' + U + '=""),"",IF(' + U + '=0,"",' + E + '/' + U + '))'        // V achvd %
  ];
}

/** Creates the missing blank rows for a date, in Shops-sheet order. Caller must hold the lock. */
function ensureRows_(ds) {
  var ss = ss_();
  var shops = shops_();
  var added = false;
  ['UAE', 'OMAN'].forEach(function (country) {
    var sh = ss.getSheetByName(country);
    if (!sh) return;
    var list = shops.filter(function (s) { return s.country === country; });
    var last = sh.getLastRow();
    var existing = {};
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) {
        if (r[0] instanceof Date && fmtDate_(r[0]) === ds) existing[r[2]] = true;
      });
    }
    var missing = list.filter(function (s) { return !existing[s.code]; });
    if (!missing.length) return;

    var start = last + 1, n = missing.length;
    var need = start + n - 1;
    if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 200);

    sh.getRange(start, 1, n, 4).setValues(missing.map(function (s) {
      return [dateObj_(ds), country, s.code, s.name];
    }));
    sh.getRange(start, C.CASH, n, 12).setFormulas(missing.map(function (s, i) { return formulasFor_(start + i); }));

    sh.getRange(start, C.DATE, n, 1).setNumberFormat('dd-mmm-yy');
    sh.getRange(start, C.TOTAL, n, 4).setNumberFormat('#,##0.000');
    sh.getRange(start, C.QTY, n, 2).setNumberFormat('#,##0');
    sh.getRange(start, C.CASH, n, 2).setNumberFormat('#,##0.000');
    sh.getRange(start, C.CASHP, n, 5).setNumberFormat('0%');
    sh.getRange(start, C.IPC, n, 1).setNumberFormat('0.0');
    sh.getRange(start, C.ATV, n, 3).setNumberFormat('#,##0.000');
    sh.getRange(start, C.ACH, n, 1).setNumberFormat('0%');
    sh.getRange(start, C.SUBMITTED, n, 2).setNumberFormat('dd-mmm-yy hh:mm');
    added = true;
  });
  if (added) SpreadsheetApp.flush();
}

function findRow_(sh, ds, code) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var v = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (v[i][0] instanceof Date && fmtDate_(v[i][0]) === ds && v[i][2] === code) return i + 2;
  }
  return 0;
}

/* ------------------------------------------------------------------ */
/*  Login / session                                                    */
/* ------------------------------------------------------------------ */
function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  var cache = CacheService.getScriptCache();
  var fk = 'f_' + username;
  var fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('Too many wrong attempts. Try again in 10 minutes.');

  var sh = ss_().getSheetByName('Users');
  var code = null;
  if (sh && sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[0]).trim().toLowerCase() === username && String(r[1]) === password &&
          String(r[3]).trim().toUpperCase() !== 'NO') {
        code = String(r[2]).trim();
        break;
      }
    }
  }
  if (!code || !shopByCode_(code)) {
    cache.put(fk, String(fails + 1), 600);
    throw new Error('Wrong username or password');
  }
  cache.remove(fk);
  var token = Utilities.getUuid();
  cache.put('s_' + token, JSON.stringify({ u: username, code: code }), 21600);
  return { token: token, state: getState(token) };
}

function auth_(token) {
  var v = token ? CacheService.getScriptCache().get('s_' + token) : null;
  if (!v) throw new Error('SESSION_EXPIRED');
  return JSON.parse(v);
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
  return true;
}

/* ------------------------------------------------------------------ */
/*  State for the app                                                  */
/* ------------------------------------------------------------------ */
function getState(token) {
  var s = auth_(token);
  return buildState_(s.code, false);
}

function buildState_(code, alreadyLocked) {
  var shop = shopByCode_(code);
  if (!shop) throw new Error('Shop not found: ' + code);
  var cfg = config_();
  var today = todayStr_();

  var allowed = [];
  for (var i = 0; i <= cfg.backdateDays; i++) allowed.push(addDays_(today, -i));

  if (alreadyLocked) {
    for (var a = allowed.length - 1; a >= 0; a--) ensureRows_(allowed[a]);
  } else {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try { for (var b = allowed.length - 1; b >= 0; b--) ensureRows_(allowed[b]); } finally { lock.release(); }
  }

  var sh = ss_().getSheetByName(shop.country);
  var last = sh.getLastRow();
  var map = {};
  if (last > 1) {
    var startRow = Math.max(2, last - SHEET_ROWS_TAIL + 1);
    sh.getRange(startRow, 1, last - startRow + 1, 24).getValues().forEach(function (r) {
      if (r[2] === shop.code && r[0] instanceof Date) map[fmtDate_(r[0])] = r;
    });
  }

  var nowMs = new Date().getTime();
  var days = [];
  for (var d = 0; d < cfg.historyDays; d++) {
    var ds = addDays_(today, -d);
    days.push(dayObj_(ds, map[ds], cfg, nowMs, allowed.indexOf(ds) >= 0));
  }

  var raw = cfg.raw;
  return {
    shop: {
      code: shop.code, name: shop.name, country: shop.country,
      currency: raw[shop.country + '_CURRENCY'] || '',
      pos1Label: String(raw[shop.country + '_POS1']), pos2Label: String(raw[shop.country + '_POS2']),
      dp: shop.country === 'UAE' ? 0 : 3
    },
    serverNow: nowMs,
    today: today,
    editMinutes: cfg.editMinutes,
    dates: allowed.map(function (ds, i) {
      return { date: ds, label: labelDate_(ds), tag: i === 0 ? 'Today' : (i === 1 ? 'Yesterday' : i + ' days ago') };
    }),
    days: days
  };
}

function dayObj_(ds, r, cfg, nowMs, inWindow) {
  var o = {
    date: ds, label: labelDate_(ds), posted: false, inWindow: inWindow, lockAt: null,
    total: null, pos1: null, pos2: null, others: null, qty: null, inv: null,
    cash: null, card: null, cashPct: null, cardPct: null, pos1Pct: null, pos2Pct: null, othersPct: null,
    ipc: null, atv: null, asp: null, target: null, ach: null, submittedAt: null, updatedAt: null, updatedLabel: ''
  };
  if (!r) return o;
  o.posted = !(r[C.TOTAL - 1] === '' || r[C.TOTAL - 1] === null);
  o.target = num_(r[C.TARGET - 1]);
  if (o.posted) {
    o.total = num_(r[C.TOTAL - 1]); o.pos1 = num_(r[C.POS1 - 1]) || 0; o.pos2 = num_(r[C.POS2 - 1]) || 0;
    o.others = num_(r[C.OTHERS - 1]) || 0; o.qty = num_(r[C.QTY - 1]); o.inv = num_(r[C.INV - 1]);
    o.cash = num_(r[C.CASH - 1]); o.card = num_(r[C.CARD - 1]);
    o.cashPct = num_(r[C.CASHP - 1]); o.cardPct = num_(r[C.CARDP - 1]);
    o.pos1Pct = num_(r[C.POS1P - 1]); o.pos2Pct = num_(r[C.POS2P - 1]); o.othersPct = num_(r[C.OTHP - 1]);
    o.ipc = num_(r[C.IPC - 1]); o.atv = num_(r[C.ATV - 1]); o.asp = num_(r[C.ASP - 1]);
    o.ach = num_(r[C.ACH - 1]);
  }
  var sub = r[C.SUBMITTED - 1], upd = r[C.UPDATED - 1];
  if (sub instanceof Date) { o.submittedAt = sub.getTime(); o.lockAt = o.submittedAt + cfg.editMinutes * 60000; }
  if (upd instanceof Date) { o.updatedAt = upd.getTime(); o.updatedLabel = Utilities.formatDate(upd, TZ, 'h:mm a'); }
  return o;
}

/* ------------------------------------------------------------------ */
/*  Save                                                               */
/* ------------------------------------------------------------------ */
function parseNum_(v, label, required) {
  if (v === null || v === undefined || String(v).trim() === '') {
    if (required) throw new Error(label + ' is required');
    return 0;
  }
  var n = Number(v);
  if (!isFinite(n) || n < 0) throw new Error(label + ' must be a number (0 or more)');
  return Math.round(n * 1000) / 1000;
}

function saveEntry(token, e) {
  var s = auth_(token);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var cfg = config_();
    var shop = shopByCode_(s.code);
    var today = todayStr_();
    var ds = String((e && e.date) || '');

    var okDate = false;
    for (var i = 0; i <= cfg.backdateDays; i++) if (addDays_(today, -i) === ds) okDate = true;
    if (!okDate) throw new Error('This date cannot be posted or edited.');

    var total = parseNum_(e.total, 'Total sales', true);
    var pos1 = parseNum_(e.pos1, 'POS-1 sale', false);
    var pos2 = parseNum_(e.pos2, 'POS-2 sale', false);
    var others = parseNum_(e.others, 'Others', false);
    var qty = parseNum_(e.qty, 'Total quantity', true);
    var inv = parseNum_(e.inv, 'Total invoices', true);
    if (qty % 1 !== 0 || inv % 1 !== 0) throw new Error('Quantity and invoices must be whole numbers.');
    if (pos1 + pos2 + others > total + 0.0005) throw new Error('POS-1 + POS-2 + Others cannot be more than Total sales.');
    if (total > 0 && (qty < 1 || inv < 1)) throw new Error('Quantity and invoices must be at least 1 when there are sales.');

    ensureRows_(ds);
    var sh = ss_().getSheetByName(shop.country);
    var row = findRow_(sh, ds, shop.code);
    if (!row) throw new Error('Row not found. Please contact admin.');

    var now = new Date();
    var sub = sh.getRange(row, C.SUBMITTED).getValue();
    if (sub instanceof Date && now.getTime() > sub.getTime() + cfg.editMinutes * 60000) {
      throw new Error('Locked: the ' + cfg.editMinutes + '-minute edit window is over. Please contact admin.');
    }

    sh.getRange(row, C.TOTAL, 1, 6).setValues([[total, pos1, pos2, others, qty, inv]]);
    if (!(sub instanceof Date)) sh.getRange(row, C.SUBMITTED).setValue(now);
    sh.getRange(row, C.UPDATED).setValue(now);
    SpreadsheetApp.flush();

    return buildState_(s.code, true);
  } finally {
    lock.release();
  }
}
