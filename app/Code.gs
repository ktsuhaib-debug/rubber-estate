/**
 * Rubber Estate Manager — Google Sheets + Apps Script web app
 *
 * The app runs as the Sheet owner, so the Sheet is never shared.
 * Family members sign in inside the app with a username + PIN.
 */
const CURRENCY = '₹';          // currency symbol shown in the app
const APP_NAME = 'Rubber Estate';
const DEFAULT_LANG = 'en';     // 'en' or 'ml' for new users
const SESSION_DAYS = 60;       // how long a phone stays logged in


const SCHEMA = {
  Users:    ['Username', 'Name', 'Role', 'PIN (hashed)', 'Language', 'Active'],
  Estates:  ['Estate', 'Total trees', 'Tapper', 'Rate per tree'],
  Tapping:  ['ID', 'Date', 'Estate', 'Tapper', 'Trees tapped', 'Rate', 'Wage', 'Sheets', 'Notes', 'Entered by', 'Entered at'],
  Work:     ['ID', 'Date', 'Estate', 'Category', 'Description', 'Workers', 'Labour cost', 'Material cost', 'Total', 'Entered by', 'Entered at'],
  Sales:    ['ID', 'Date', 'Estate', 'Sheets sold', 'Weight (kg)', 'Price per kg', 'Amount', 'Buyer', 'Entered by', 'Entered at'],
  Payments: ['ID', 'Date', 'Paid to', 'Amount', 'For', 'Entered by', 'Entered at']
};
const ENTRY_SHEETS = ['Tapping', 'Work', 'Sales', 'Payments'];

/* ---------- One-time setup (safe to run again after updates) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  // older versions used an Email-based Users tab: replace it
  const oldUsers = ss.getSheetByName('Users');
  if (oldUsers && oldUsers.getRange(1, 1).getValue() === 'Email') oldUsers.clear();

  Object.entries(SCHEMA).forEach(([name, head]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0 || sh.getLastColumn() < head.length) {
      sh.getRange(1, 1, 1, head.length).setValues([head])
        .setFontWeight('bold').setBackground('#e6efe8');
      sh.setFrozenRows(1);
    }
  });

  const users = ss.getSheetByName('Users');
  users.getRange('A:A').setNumberFormat('@');
  if (users.getLastRow() === 1) {
    users.appendRow(['admin', 'Owner', 'Admin', hashPin_('1234'), DEFAULT_LANG, true]);
    Logger.log('Created login → username: admin  PIN: 1234  (change it in the app: Settings → Account)');
  }
  const est = ss.getSheetByName('Estates');
  if (est.getLastRow() === 1) {
    est.getRange(2, 1, 2, 4).setValues([
      ['Estate 1', 1000, 'Tapper 1', 2],
      ['Estate 2', 1000, 'Tapper 2', 2]
    ]);
  }
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  reportsFolder_();                     // creates the "Reports" folder next to the Sheet
  UrlFetchApp.getRequest;               // so Google asks for the permissions PDF reports need
  DocumentApp.getActiveDocument;
  alignTimeZone_(ss);
}

// The Sheet and the script must use the same time zone, or dates show a day early in the Sheet.
// Existing entries are re-written so each keeps its real date.
function alignTimeZone_(ss) {
  const tz = Session.getScriptTimeZone();
  if (ss.getSpreadsheetTimeZone() !== tz) {
    const saved = ENTRY_SHEETS.map(n => {
      const sh = ss.getSheetByName(n), last = sh.getLastRow();
      if (last < 2) return null;
      const w = sh.getLastColumn();
      return { sh, date: sh.getRange(2, 2, last - 1, 1).getValues(), at: sh.getRange(2, w, last - 1, 1).getValues(), w, rows: last - 1 };
    });
    ss.setSpreadsheetTimeZone(tz);
    saved.forEach(s => {
      if (!s) return;
      s.sh.getRange(2, 2, s.rows, 1).setValues(s.date);
      s.sh.getRange(2, s.w, s.rows, 1).setValues(s.at);
    });
    Logger.log('Sheet time zone set to ' + tz);
  }
  ENTRY_SHEETS.forEach(n => {
    const sh = ss.getSheetByName(n);
    sh.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, SCHEMA[n].length, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('dd/mm/yyyy hh:mm');
  });
}

/* ---------- Web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Login ---------- */
function login(username, pin) {
  username = norm_(username);
  const cache = CacheService.getScriptCache();
  const key = 'fail_' + username;
  const fails = Number(cache.get(key) || 0);
  if (fails >= 5) err_('locked');
  const users = rows_('Users');
  const u = users.find(r => norm_(r[0]) === username);
  if (!u || !active_(u) || !checkPin_(pin, u[3])) {
    cache.put(key, String(fails + 1), 600);
    err_('bad_login');
  }
  cache.remove(key);
  pruneSessions_();
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('s_' + token,
    JSON.stringify({ u: username, exp: Date.now() + SESSION_DAYS * 864e5 }));
  return token;
}

function logout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('s_' + token);
  return true;
}

function changePin(token, oldPin, newPin) {
  const me = me_(token);
  const row = rows_('Users')[me.row - 2];
  if (!checkPin_(oldPin, row[3])) err_('bad_pin');
  pinFormat_(newPin);
  sh_('Users').getRange(me.row, 4).setValue(hashPin_(newPin));
  return true;
}

/* ---------- App data ---------- */
function init(token) {
  const u = me_(token);
  const estates = rows_('Estates').filter(r => r[0]).map(r => ({
    name: String(r[0]), trees: num_(r[1]), tapper: String(r[2] || ''), rate: num_(r[3])
  }));
  const people = [...new Set([
    ...estates.map(e => e.tapper),
    ...rows_('Tapping').map(r => String(r[3] || ''))
  ].filter(Boolean))];
  return {
    user: { username: u.username, name: u.name, role: u.role, lang: u.lang },
    estates, people, currency: CURRENCY, appName: APP_NAME, today: ymd_(new Date())
  };
}

function setLanguage(token, lang) {
  if (!['en', 'ml'].includes(lang)) err_('bad_lang');
  const u = me_(token);
  sh_('Users').getRange(u.row, 5).setValue(lang);
  return lang;
}

/* ---------- Add entries ---------- */
function addTapping(token, f) {
  const u = me_(token);
  const trees = num_(f.trees);
  let rate = num_(f.rate), tapper = f.tapper;
  if (u.role !== 'Admin') {              // members always use the admin-set tapper & rate
    const e = rows_('Estates').find(r => String(r[0]) === f.estate);
    if (!e) err_('unknown');
    tapper = String(e[2] || ''); rate = num_(e[3]);
  }
  if (trees <= 0) err_('no_trees');
  const date = toDate_(f.date);
  const dup = rows_('Tapping').some(r => r[1] instanceof Date && ymd_(r[1]) === f.date && r[2] === f.estate);
  if (dup) err_('dup_tap', f.estate + ' · ' + f.date);
  const wage = trees * rate;
  add_(u, 'Tapping', [date, f.estate, tapper, trees, rate, wage, num_(f.sheets), f.notes || '']);
  return wage;
}

function addWork(token, f) {
  const u = me_(token);
  const labour = num_(f.labour), material = num_(f.material);
  if (labour + material <= 0) err_('no_cost');
  add_(u, 'Work', [toDate_(f.date), f.estate, f.category, f.description || '', f.workers || '', labour, material, labour + material]);
  return true;
}

function addSale(token, f) {
  const u = me_(token);
  const amount = num_(f.amount) || num_(f.kg) * num_(f.price);
  if (amount <= 0) err_('no_amount');
  add_(u, 'Sales', [toDate_(f.date), f.estate, num_(f.sheets), num_(f.kg), num_(f.price), amount, f.buyer || '']);
  return true;
}

function addPayment(token, f) {
  const u = me_(token);
  if (num_(f.amount) <= 0) err_('no_pay');
  add_(u, 'Payments', [toDate_(f.date), f.to, num_(f.amount), f.note || '']);
  return true;
}

/* ---------- Lists ---------- */
function getRecent(token, sheet) {
  me_(token);
  if (!ENTRY_SHEETS.includes(sheet)) err_('unknown');
  return rows_(sheet).slice(-20).reverse()
    .map(r => r.map(v => v instanceof Date ? ymd_(v) : v));
}

function deleteEntry(token, sheet, id) {
  admin_(token);
  if (!ENTRY_SHEETS.includes(sheet)) err_('unknown');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sh_(sheet);
    const last = sh.getLastRow();
    if (last < 2) err_('not_found');
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); return true; }
    }
    err_('not_found');
  } finally { lock.releaseLock(); }
}

/* ---------- Admin: estates ---------- */
function saveEstates(token, list) {
  admin_(token);
  if (!Array.isArray(list) || !list.length) err_('est_need');
  const clean = list.map(e => ({
    name: String(e.name || '').trim(), old: String(e.old || '').trim(),
    trees: num_(e.trees), tapper: String(e.tapper || '').trim(), rate: num_(e.rate)
  }));
  if (clean.some(e => !e.name)) err_('est_name');
  if (clean.some(e => e.name.toLowerCase() === 'both')) err_('est_both');
  const lower = clean.map(e => e.name.toLowerCase());
  if (new Set(lower).size !== lower.length) err_('est_dup');
  if (clean.some(e => e.trees < 0 || e.rate < 0)) err_('est_neg');
  const renamed = clean.filter(e => e.old && e.old !== e.name);
  const oldNames = renamed.map(e => e.old);
  if (renamed.some(e => oldNames.includes(e.name))) err_('est_swap');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sh_('Estates');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
    sh.getRange(2, 1, clean.length, 4).setValues(clean.map(e => [e.name, e.trees, e.tapper, e.rate]));
    renamed.forEach(e => ['Tapping', 'Work', 'Sales'].forEach(s => renameIn_(s, 3, e.old, e.name)));
  } finally { lock.releaseLock(); }
  return true;
}

function renameIn_(sheet, col, from, to) {
  const sh = sh_(sheet), last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, col, last - 1, 1), v = rng.getValues();
  let changed = false;
  v.forEach(x => { if (String(x[0]) === from) { x[0] = to; changed = true; } });
  if (changed) rng.setValues(v);
}

/* ---------- Admin: users ---------- */
function getUsers(token) {
  admin_(token);
  return rows_('Users').filter(r => r[0]).map(r => ({
    username: norm_(r[0]), name: String(r[1] || ''),
    role: String(r[2]).trim() === 'Admin' ? 'Admin' : 'Member', active: active_(r)
  }));
}

function saveUser(token, f) {
  const me = admin_(token);
  const username = norm_(f.username);
  const name = String(f.name || '').trim();
  const role = f.role === 'Admin' ? 'Admin' : 'Member';
  if (!/^[a-z0-9._-]{3,20}$/.test(username)) err_('bad_user');
  if (!name) err_('no_name');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sh_('Users');
    const i = rows_('Users').findIndex(r => norm_(r[0]) === username);
    if (f.isNew) {
      if (i >= 0) err_('user_taken');
      pinFormat_(f.pin);
      sh.appendRow([username, name, role, hashPin_(f.pin), DEFAULT_LANG, true]);
    } else {
      if (i < 0) err_('not_found');
      if (username === me.username && role !== 'Admin') err_('self_admin');
      sh.getRange(i + 2, 2, 1, 2).setValues([[name, role]]);
      if (f.pin) { pinFormat_(f.pin); sh.getRange(i + 2, 4).setValue(hashPin_(f.pin)); }
    }
  } finally { lock.releaseLock(); }
  return true;
}

function removeUser(token, username) {
  const me = admin_(token);
  username = norm_(username);
  if (username === me.username) err_('self_remove');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const i = rows_('Users').findIndex(r => norm_(r[0]) === username);
    if (i < 0) err_('not_found');
    sh_('Users').deleteRow(i + 2);      // their sessions stop working on the next request
  } finally { lock.releaseLock(); }
  return true;
}

/* ---------- Dashboard ---------- */
function getDashboard(token, period) {
  me_(token);
  const tz = tz_();
  const isYear = period.length === 4;
  const year = period.slice(0, 4);
  const fmtP = isYear ? 'yyyy' : 'yyyy-MM';
  const inP = d => d instanceof Date && Utilities.formatDate(d, tz, fmtP) === period;
  const inY = d => d instanceof Date && Utilities.formatDate(d, tz, 'yyyy') === year;

  const est = {};
  const E = n => est[n] || (est[n] = { name: String(n), trees: 0, sheets: 0, sold: 0, kg: 0, income: 0, wages: 0, labour: 0, material: 0 });
  rows_('Estates').forEach(r => r[0] && E(r[0]));

  const months = {};
  const M = d => {
    const k = Utilities.formatDate(d, tz, 'yyyy-MM');
    return months[k] || (months[k] = { income: 0, cost: 0 });
  };

  let made = 0, soldAll = 0;
  const earned = {}, paid = {};

  rows_('Tapping').forEach(r => {
    made += num_(r[7]);
    earned[r[3]] = (earned[r[3]] || 0) + num_(r[6]);
    if (inY(r[1])) M(r[1]).cost += num_(r[6]);
    if (!inP(r[1])) return;
    const x = E(r[2]); x.trees += num_(r[4]); x.wages += num_(r[6]); x.sheets += num_(r[7]);
  });
  rows_('Work').forEach(r => {
    if (inY(r[1])) M(r[1]).cost += num_(r[8]);
    if (!inP(r[1])) return;
    const x = E(r[2]); x.labour += num_(r[6]); x.material += num_(r[7]);
  });
  rows_('Sales').forEach(r => {
    soldAll += num_(r[3]);
    if (inY(r[1])) M(r[1]).income += num_(r[6]);
    if (!inP(r[1])) return;
    const x = E(r[2]); x.sold += num_(r[3]); x.kg += num_(r[4]); x.income += num_(r[6]);
  });
  rows_('Payments').forEach(r => { paid[r[2]] = (paid[r[2]] || 0) + num_(r[3]); });

  const list = Object.values(est);
  const keys = ['trees', 'sheets', 'sold', 'kg', 'income', 'wages', 'labour', 'material'];
  const total = { name: 'All' };
  keys.forEach(k => total[k] = list.reduce((s, x) => s + x[k], 0));
  [...list, total].forEach(x => { x.cost = x.wages + x.labour + x.material; x.profit = x.income - x.cost; });

  const due = Object.keys(earned).map(n => ({
    name: n, earned: earned[n], paid: paid[n] || 0, due: earned[n] - (paid[n] || 0)
  })).filter(x => Math.abs(x.due) > 0.009);

  const trend = [];
  for (let m = 1; m <= 12; m++) {
    const k = year + '-' + String(m).padStart(2, '0');
    const v = months[k] || { income: 0, cost: 0 };
    trend.push({ month: k, income: v.income, cost: v.cost, profit: v.income - v.cost });
  }

  return { total, estates: list.filter(x => x.cost || x.income || x.trees), stock: made - soldAll, due, trend };
}

/* ---------- PDF reports ---------- */
const RT = {
  en: {
    tap_title: 'Tapper statement', biz_title: 'Business statement', period: 'Period', tapper: 'Tapper',
    date: 'Date', estate: 'Estate', trees: 'Trees', rate: 'Rate', wage: 'Wage', sheets: 'Sheets',
    opening: 'Balance due at start', earned: 'Wages earned', paid: 'Paid', closing: 'Balance due',
    daily: 'Daily tapping', payments: 'Payments', note: 'For', total: 'Total', none: 'No entries',
    summary: 'Summary', income: 'Sales income', tap_wages: 'Tapping wages', labour: 'Work labour',
    material: 'Materials', costs: 'Total costs', profit: 'Profit', loss: 'Loss',
    production: 'Production', trees_tapped: 'Trees tapped', sheets_made: 'Sheets made', sheets_sold: 'Sheets sold',
    kg: 'Weight sold (kg)', stock: 'Sheets in stock', work: 'Estate work & expenses', sales: 'Sales',
    buyer: 'Buyer', type: 'Type', desc: 'Details', amount: 'Amount', due: 'Wages due', generated: 'Generated on',
    shared: 'Shared', price: 'Price/kg', days: 'Days tapped',
    'c:Grass cutting': 'Grass cutting', 'c:Fertiliser': 'Fertiliser', 'c:Weeding': 'Weeding',
    'c:Spraying': 'Spraying', 'c:Rain guarding': 'Rain guarding', 'c:Other': 'Other'
  },
  ml: {
    tap_title: 'ടാപ്പർ സ്റ്റേറ്റ്മെന്റ്', biz_title: 'ബിസിനസ് സ്റ്റേറ്റ്മെന്റ്', period: 'കാലയളവ്', tapper: 'ടാപ്പർ',
    date: 'തീയതി', estate: 'എസ്റ്റേറ്റ്', trees: 'മരങ്ങൾ', rate: 'നിരക്ക്', wage: 'കൂലി', sheets: 'ഷീറ്റ്',
    opening: 'തുടക്കത്തിലെ ബാക്കി', earned: 'ലഭിക്കേണ്ട കൂലി', paid: 'നൽകിയത്', closing: 'കൊടുക്കാനുള്ള ബാക്കി',
    daily: 'ദിവസേനയുള്ള ടാപ്പിംഗ്', payments: 'നൽകിയ തുകകൾ', note: 'കാലയളവ്', total: 'ആകെ', none: 'എൻട്രികൾ ഇല്ല',
    summary: 'സംഗ്രഹം', income: 'വിൽപ്പന വരുമാനം', tap_wages: 'ടാപ്പിംഗ് കൂലി', labour: 'പണിക്കൂലി',
    material: 'സാധനങ്ങൾ', costs: 'ആകെ ചെലവ്', profit: 'ലാഭം', loss: 'നഷ്ടം',
    production: 'ഉത്പാദനം', trees_tapped: 'ടാപ്പ് ചെയ്ത മരങ്ങൾ', sheets_made: 'ഉണ്ടാക്കിയ ഷീറ്റ്', sheets_sold: 'വിറ്റ ഷീറ്റ്',
    kg: 'വിറ്റ തൂക്കം (കി.ഗ്രാം)', stock: 'സ്റ്റോക്കിലുള്ള ഷീറ്റ്', work: 'എസ്റ്റേറ്റ് പണികളും ചെലവുകളും', sales: 'വിൽപ്പന',
    buyer: 'വാങ്ങിയയാൾ', type: 'തരം', desc: 'വിവരം', amount: 'തുക', due: 'കൊടുക്കാനുള്ള കൂലി', generated: 'തയ്യാറാക്കിയത്',
    shared: 'പൊതുവായത്', price: 'വില/കി.ഗ്രാം', days: 'ടാപ്പ് ചെയ്ത ദിവസങ്ങൾ',
    'c:Grass cutting': 'പുല്ലുവെട്ട്', 'c:Fertiliser': 'വളമിടൽ', 'c:Weeding': 'കളപറിക്കൽ',
    'c:Spraying': 'മരുന്നടി', 'c:Rain guarding': 'റെയിൻ ഗാർഡ്', 'c:Other': 'മറ്റുള്ളവ'
  }
};

function makeReport(token, f) {
  me_(token);
  const from = String(f.from || ''), to = String(f.to || '');
  toDate_(from); toDate_(to);
  if (from > to) err_('bad_range');
  const lang = f.lang === 'ml' ? 'ml' : 'en';
  const T = k => RT[lang][k] || RT.en[k] || k;
  let title, body, fileName;
  if (f.kind === 'tapper') {
    if (!f.tapper) err_('no_tapper');
    title = T('tap_title');
    body = tapperReport_(String(f.tapper), from, to, T);
    fileName = 'Statement - ' + f.tapper + ' - ' + dmy_(from) + ' to ' + dmy_(to);
  } else {
    title = T('biz_title');
    body = businessReport_(from, to, T);
    fileName = 'Business statement - ' + dmy_(from) + ' to ' + dmy_(to);
  }
  const html = '<html><head><meta charset="utf-8"></head><body style="font-family:Arial;font-size:10pt;color:#16201a">' +
    '<p style="font-size:9pt;color:#2d7a50;margin:0"><b>' + h_(APP_NAME) + '</b></p>' +
    '<h1 style="font-size:18pt;margin:2pt 0 2pt 0">' + h_(title) + '</h1>' +
    '<p style="color:#6b746d;margin:0 0 10pt 0">' + h_(T('period')) + ': ' + dmy_(from) + ' – ' + dmy_(to) + '</p>' +
    body +
    '<p style="color:#6b746d;font-size:8pt;margin-top:14pt">' + h_(T('generated')) + ' ' +
    Utilities.formatDate(new Date(), tz_(), 'dd-MM-yyyy HH:mm') + '</p></body></html>';
  return savePdf_(fileName, html);
}

function tapperReport_(tapper, from, to, T) {
  const inR = d => d instanceof Date && ymd_(d) >= from && ymd_(d) <= to;
  const before = d => d instanceof Date && ymd_(d) < from;
  let opening = 0, earned = 0, paid = 0, trees = 0, sheets = 0;
  const days = [], pays = [];
  rows_('Tapping').forEach(r => {
    if (String(r[3]) !== tapper) return;
    if (before(r[1])) opening += num_(r[6]);
    if (inR(r[1])) { days.push(r); earned += num_(r[6]); trees += num_(r[4]); sheets += num_(r[7]); }
  });
  rows_('Payments').forEach(r => {
    if (String(r[2]) !== tapper) return;
    if (before(r[1])) opening -= num_(r[3]);
    if (inR(r[1])) { pays.push(r); paid += num_(r[3]); }
  });
  days.sort((a, b) => a[1] - b[1]); pays.sort((a, b) => a[1] - b[1]);
  const closing = opening + earned - paid;

  let out = '<p style="font-size:12pt;margin:0 0 8pt 0"><b>' + h_(T('tapper')) + ': ' + h_(tapper) + '</b></p>';
  out += table_([], [
    [T('opening'), money_(opening)],
    [T('earned') + ' (' + days.length + ' ' + T('days').toLowerCase() + ')', money_(earned)],
    [T('paid'), money_(-paid)],
    [{ v: T('closing'), b: true }, { v: money_(closing), b: true }]
  ], [70, 30]);
  out += h2_(T('daily'));
  out += days.length ? table_([T('date'), T('estate'), T('trees'), T('rate'), T('sheets'), T('wage')],
    days.map(r => [dmy_(r[1]), r[2], int_(r[4]), money_(r[5], 2), int_(r[7]), money_(r[6])])
      .concat([[{ v: T('total'), b: true }, '', { v: int_(trees), b: true }, '', { v: int_(sheets), b: true }, { v: money_(earned), b: true }]]),
    [16, 22, 14, 14, 14, 20], [2, 3, 4, 5]) : none_(T);
  out += h2_(T('payments'));
  out += pays.length ? table_([T('date'), T('note'), T('amount')],
    pays.map(r => [dmy_(r[1]), r[4] || '', money_(r[3])])
      .concat([[{ v: T('total'), b: true }, '', { v: money_(paid), b: true }]]), [20, 55, 25], [2]) : none_(T);
  return out;
}

function businessReport_(from, to, T) {
  const inR = d => d instanceof Date && ymd_(d) >= from && ymd_(d) <= to;
  const upTo = d => d instanceof Date && ymd_(d) <= to;
  const names = rows_('Estates').filter(r => r[0]).map(r => String(r[0]));
  const est = {};
  const E = n => est[n] || (est[n] = { trees: 0, sheets: 0, sold: 0, kg: 0, income: 0, wages: 0, labour: 0, material: 0 });
  names.forEach(E);
  const work = [], sales = [], earned = {}, paidTo = {};
  let made = 0, soldAll = 0;
  rows_('Tapping').forEach(r => {
    if (upTo(r[1])) { made += num_(r[7]); earned[r[3]] = (earned[r[3]] || 0) + num_(r[6]); }
    if (!inR(r[1])) return;
    const x = E(String(r[2])); x.trees += num_(r[4]); x.sheets += num_(r[7]); x.wages += num_(r[6]);
  });
  rows_('Work').forEach(r => {
    if (!inR(r[1])) return;
    work.push(r); const x = E(String(r[2])); x.labour += num_(r[6]); x.material += num_(r[7]);
  });
  rows_('Sales').forEach(r => {
    if (upTo(r[1])) soldAll += num_(r[3]);
    if (!inR(r[1])) return;
    sales.push(r); const x = E(String(r[2])); x.sold += num_(r[3]); x.kg += num_(r[4]); x.income += num_(r[6]);
  });
  rows_('Payments').forEach(r => { if (upTo(r[1])) paidTo[r[2]] = (paidTo[r[2]] || 0) + num_(r[3]); });

  const cols = Object.keys(est).filter(n => names.includes(n) || Object.values(est[n]).some(v => v));
  const tot = {}; Object.keys(E('__t')).forEach(k => tot[k] = cols.reduce((s, n) => s + est[n][k], 0));
  delete est.__t;
  const all = cols.map(n => est[n]).concat([tot]);
  all.forEach(x => { x.cost = x.wages + x.labour + x.material; x.profit = x.income - x.cost; });
  const head = [''].concat(cols.map(n => n === 'Both' ? T('shared') : n), [T('total')]);
  const w = Math.floor(60 / (cols.length + 1));
  const widths = [40].concat(cols.map(() => w), [w]);
  const numCols = head.map((_, i) => i).slice(1);
  const line = (label, k, fmt, bold) => [bold ? { v: label, b: true } : label]
    .concat(all.map(x => bold ? { v: fmt(x[k]), b: true } : fmt(x[k])));

  let out = h2_(T('summary'));
  out += table_(head, [
    line(T('income'), 'income', money_),
    line(T('tap_wages'), 'wages', v => money_(-v)),
    line(T('labour'), 'labour', v => money_(-v)),
    line(T('material'), 'material', v => money_(-v)),
    line(T('costs'), 'cost', v => money_(-v)),
    line(tot.profit < 0 ? T('loss') : T('profit'), 'profit', money_, true)
  ], widths, numCols);
  out += h2_(T('production'));
  out += table_(head, [
    line(T('trees_tapped'), 'trees', int_),
    line(T('sheets_made'), 'sheets', int_),
    line(T('sheets_sold'), 'sold', int_),
    line(T('kg'), 'kg', v => int_(v, 2))
  ], widths, numCols);
  out += '<p style="margin:4pt 0 0 0">' + h_(T('stock')) + ': <b>' + int_(made - soldAll) + '</b></p>';

  work.sort((a, b) => a[1] - b[1]); sales.sort((a, b) => a[1] - b[1]);
  out += h2_(T('work'));
  out += work.length ? table_([T('date'), T('estate'), T('type'), T('desc'), T('amount')],
    work.map(r => [dmy_(r[1]), r[2] === 'Both' ? T('shared') : r[2], T('c:' + r[3]), [r[4], r[5]].filter(String).join(' · '), money_(r[8])])
      .concat([[{ v: T('total'), b: true }, '', '', '', { v: money_(tot.labour + tot.material), b: true }]]),
    [14, 16, 18, 32, 20], [4]) : none_(T);
  out += h2_(T('sales'));
  out += sales.length ? table_([T('date'), T('estate'), T('sheets'), 'kg', T('price'), T('buyer'), T('amount')],
    sales.map(r => [dmy_(r[1]), r[2] === 'Both' ? T('shared') : r[2], int_(r[3]), int_(r[4], 2), money_(r[5], 2), r[7] || '', money_(r[6])])
      .concat([[{ v: T('total'), b: true }, '', { v: int_(tot.sold), b: true }, { v: int_(tot.kg, 2), b: true }, '', '', { v: money_(tot.income), b: true }]]),
    [13, 15, 11, 11, 13, 18, 19], [2, 3, 4, 6]) : none_(T);
  const due = Object.keys(earned).map(n => [n, earned[n] - (paidTo[n] || 0)]).filter(x => Math.abs(x[1]) > 0.009);
  out += h2_(T('due') + ' (' + dmy_(to) + ')');
  out += due.length ? table_([T('tapper'), T('amount')], due.map(x => [x[0], money_(x[1])]), [70, 30], [1]) : none_(T);
  return out;
}

// Build the PDF through a temporary Google Doc so Malayalam text renders correctly
function savePdf_(name, html) {
  const folder = reportsFolder_();
  const boundary = 'rep' + Date.now();
  const meta = { name: name, mimeType: 'application/vnd.google-apps.document', parents: [folder.getId()] };
  const body = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
    '\r\n--' + boundary + '\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n' + html + '\r\n--' + boundary + '--';
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
    method: 'post', contentType: 'multipart/related; boundary=' + boundary,
    payload: Utilities.newBlob(body).getBytes(),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) { console.error(res.getContentText()); err_('pdf_fail'); }
  const docId = JSON.parse(res.getContentText()).id;
  try {
    const doc = DocumentApp.openById(docId);
    doc.getBody().setPageWidth(595.28).setPageHeight(841.89)          // A4
      .setMarginTop(36).setMarginBottom(36).setMarginLeft(40).setMarginRight(40);
    doc.saveAndClose();
    const pdf = folder.createFile(DriveApp.getFileById(docId).getAs('application/pdf').setName(name + '.pdf'));
    pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: pdf.getUrl(), name: pdf.getName() };
  } finally {
    DriveApp.getFileById(docId).setTrashed(true);
  }
}

function reportsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('REPORTS_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const parents = DriveApp.getFileById(ss_().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const it = parent.getFoldersByName('Reports');
  const f = it.hasNext() ? it.next() : parent.createFolder('Reports');
  props.setProperty('REPORTS_FOLDER', f.getId());
  return f;
}

function h_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function h2_(s) { return '<h2 style="font-size:12pt;color:#1e5b3a;margin:14pt 0 4pt 0">' + h_(s) + '</h2>'; }
function none_(T) { return '<p style="color:#6b746d;margin:0">' + h_(T('none')) + '</p>'; }
function dmy_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, tz_(), 'dd-MM-yyyy');
  const p = String(d).split('-'); return p.length === 3 ? p[2] + '-' + p[1] + '-' + p[0] : String(d);
}
// Indian digit grouping (1,23,456.50)
function int_(v, dec) {
  const n = Number(v) || 0, neg = n < 0, fixed = Math.abs(n).toFixed(dec || 0);
  let [i, f] = fixed.split('.');
  if (dec && /^0+$/.test(f)) f = '';
  const last3 = i.slice(-3), rest = i.slice(0, -3);
  i = (rest ? rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' : '') + last3;
  return (neg ? '-' : '') + i + (f ? '.' + f : '');
}
function money_(v, dec) { const n = Number(v) || 0; return (n < 0 ? '-' : '') + CURRENCY + int_(Math.abs(n), dec); }
// cells: string or {v, b}; numCols: right-aligned column indexes
function table_(head, rows, widths, numCols) {
  numCols = numCols || [];
  const cell = (c, i, th) => {
    const o = (c && typeof c === 'object') ? c : { v: c };
    const st = 'border:1px solid #d5dbd5;padding:3pt 5pt;' + (widths && widths[i] ? 'width:' + widths[i] + '%;' : '') +
      (numCols.includes(i) || (!head.length && i === 1) ? 'text-align:right;' : '') +
      (th ? 'background-color:#e9f2ec;color:#1e5b3a;' : '');
    const v = h_(o.v);
    return '<td style="' + st + '">' + (th || o.b ? '<b>' + v + '</b>' : v) + '</td>';
  };
  return '<table style="border-collapse:collapse;width:100%;font-size:9.5pt">' +
    (head.length ? '<tr>' + head.map((c, i) => cell(c, i, true)).join('') + '</tr>' : '') +
    rows.map(r => '<tr>' + r.map((c, i) => cell(c, i, false)).join('') + '</tr>').join('') + '</table>';
}

/* ---------- Helpers ---------- */
// errors are sent as codes so the app can show them in English or Malayalam
function err_(code, detail) { throw new Error('ERR:' + code + (detail ? '|' + detail : '')); }

function me_(token) {
  if (!token) err_('auth');
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('s_' + token);
  if (!raw) err_('auth');
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { props.deleteProperty('s_' + token); err_('auth'); }
  const all = rows_('Users');
  const i = all.findIndex(r => norm_(r[0]) === s.u);
  if (i < 0 || !active_(all[i])) { props.deleteProperty('s_' + token); err_('auth'); }
  const u = all[i];
  const lang = String(u[4] || '').trim().toLowerCase();
  return {
    username: s.u, row: i + 2,
    name: String(u[1] || s.u),
    role: String(u[2]).trim() === 'Admin' ? 'Admin' : 'Member',
    lang: ['en', 'ml'].includes(lang) ? lang : DEFAULT_LANG
  };
}
function admin_(token) { const u = me_(token); if (u.role !== 'Admin') err_('admin_only'); return u; }

function active_(r) { return !(r[5] === false || String(r[5]).trim().toUpperCase() === 'FALSE'); }
function norm_(s) { return String(s || '').trim().toLowerCase(); }

function pinFormat_(pin) { if (!/^\d{4,8}$/.test(String(pin || ''))) err_('pin_format'); }
function hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function hashPin_(pin) { const salt = Utilities.getUuid().slice(0, 8); return salt + ':' + hex_(salt + String(pin)); }
function checkPin_(pin, stored) {
  const [salt, h] = String(stored || '').split(':');
  return !!salt && !!h && hex_(salt + String(pin || '')) === h;
}
function pruneSessions_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(k => {
    if (k.indexOf('s_') !== 0) return;
    try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

function add_(u, sheet, values) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sh_(sheet).appendRow([Utilities.getUuid().slice(0, 8), ...values, u.name, new Date()]);
  } finally { lock.releaseLock(); }
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sh_(n) { return ss_().getSheetByName(n); }
function rows_(n) { const v = sh_(n).getDataRange().getValues(); v.shift(); return v; }
function num_(x) { return Number(x) || 0; }
function tz_() { return Session.getScriptTimeZone(); }
function ymd_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function toDate_(s) {
  const [y, m, d] = String(s).split('-').map(Number);
  if (!y || !m || !d) err_('bad_date');
  return new Date(y, m - 1, d);
}
