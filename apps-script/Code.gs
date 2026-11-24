/**
 * Finances app backend (Google Apps Script).
 *
 * The app on GitHub Pages sends requests here; this script reads and writes
 * your Google Sheet and saves receipts to Google Drive.
 *
 * Setup (once):
 *   1. Paste this file into Extensions > Apps Script of your sheet.
 *   2. Set SHEET_ID below to the sheet the app should use.
 *   3. Run the function "setup" once (Run button). Allow the permissions.
 *      The key is shown in the Execution log and in
 *      Project Settings > Script Properties > APP_KEY.
 *   4. Deploy > New deployment > Web app.
 *        Execute as: Me      Who has access: Anyone
 *   5. In the app's More screen, paste the URL and the key.
 *
 * After changing this file: Deploy > Manage deployments > Edit >
 * Version: New version > Deploy. The URL stays the same.
 */

// ---------- Configuration ----------
const SHEET_ID = '1zuRAD9ZuQ7eUTuEknaQ2J02mMITWfgYPTdnjHbzt7Vk'; // "Finances (app test)"
const SALARY_DAY = 20;            // salary arrives on this day of the month
const RECEIPTS_FOLDER = 'Finances Receipts';
const TZ = 'Europe/Zurich';
const ACCOUNT_TYPES = ['Current', 'Credit card', 'Savings', 'Cash'];

// Where each kind of entry lives. inputs = columns the app writes (1-based);
// formulas in the other columns are already filled down in your tabs.
const KINDS = {
  expense:  { sheet: 'Log',       width: 13, inputs: [1, 2, 3, 4, 5, 6, 7, 8, 12, 13], fStart: 9, fCount: 3 },
  income:   { sheet: 'Income',    width: 11, inputs: [1, 2, 3, 4, 5, 6, 7, 8],         fStart: 9, fCount: 3 },
  transfer: { sheet: 'Transfers', width: 12, inputs: [1, 2, 3, 4, 5, 6],               fStart: 7, fCount: 6 },
};

// ---------- Entry point ----------
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'Bad request' });
  }
  const key = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (!key || req.key !== key) return out_({ ok: false, error: 'Wrong app key' });

  ensureSchema_();
  const reads = {
    bootstrap: () => bootstrap_(req.month),
    summary: () => summary_(req.month),
    plan: () => plan_(),
    search: () => search_(req.query || {}),
    categoryDetail: () => categoryDetail_(req.name, req.month),
    transferCosts: () => transferCosts_(),
    reconcileCheck: () => reconcile_(req.account, req.bankBalance, false),
    reminderSync: () => reminderSync_(req.existing, req.days),
  };

  const writes = {
    addExpense:     () => addEntry_('expense', req.entry),
    addIncome:      () => addEntry_('income', req.entry),
    addTransfer:    () => addEntry_('transfer', req.entry),
    updateEntry:    () => updateEntry_(req.kind, req.row, req.match, req.entry),
    deleteEntry:    () => deleteEntry_(req.kind, req.row, req.match),
    saveAccount:    () => saveAccount_(req.account),
    deleteAccount:  () => deleteAccount_(req.name),
    saveCategory:   () => saveCategory_(req.category),
    deleteCategory: () => deleteCategory_(req.name),
    saveRecurring:  () => saveRecurring_(req.item),
    deleteRecurring: () => deleteRecurring_(req.row, req.match),
    markPaid:       () => markPaid_(req.payment),
    skipPayment:    () => skipPayment_(req.row, req.due, req.match),
    reconcile:      () => reconcile_(req.account, req.bankBalance, true),
    saveEnvelope:   () => saveEnvelope_(req.envelope),
    deleteEnvelope: () => deleteEnvelope_(req.row, req.match),
    moveEnvelope:   () => moveEnvelope_(req.row, req.match, req.amount, req.reason),
    saveSavingsTarget: () => saveSavingsTarget_(req.amount),
  };
  try {
    if (reads[req.action]) return out_({ ok: true, data: reads[req.action]() });
    if (!writes[req.action]) return out_({ ok: false, error: 'Unknown action' });
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const data = writes[req.action]();
      SpreadsheetApp.flush();
      return out_({ ok: true, data });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// Opening the URL in a browser just shows that it works.
function doGet() {
  return out_({ ok: true, message: 'Finances backend is running.' });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run once from the editor to create the secret key.
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('APP_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('APP_KEY', key);
  }
  console.log('Your app key: ' + key);
  console.log('You can also find it any time in Project Settings > Script Properties > APP_KEY');
  try {
    receiptsFolder_(); // create the Drive folder now so permissions are asked once
    console.log('Receipts folder is ready in your Google Drive.');
  } catch (err) {
    console.log('Receipts folder could not be created yet: ' + err.message);
  }
}

// Run this if you ever need a new key (the old one stops working).
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty('APP_KEY');
  setup();
}

// ---------- Reading ----------
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function readTable_(sheet, firstRow, lastCol) {
  const last = sheet.getLastRow();
  if (last < firstRow) return [];
  return sheet.getRange(firstRow, 1, last - firstRow + 1, lastCol).getValues();
}

// Settings: A currency, D rate used | F category, G monthly budget | I item, J usual category
function settings_(ss) {
  const sh = ss.getSheetByName('Settings');
  const rows = sh.getRange(4, 1, 100, 10).getValues();
  const currencies = [], categories = [], items = [];
  rows.forEach((r, i) => {
    if (r[0] !== '') currencies.push({ code: String(r[0]), rate: Number(r[3]) || 1 });
    if (r[5] !== '') categories.push({ name: String(r[5]), budget: Number(r[6]) || 0, row: i + 4 });
    if (r[8] !== '') items.push({ name: String(r[8]), category: String(r[9] || '') });
  });
  return { currencies, categories, items };
}

// Accounts rows 5-30: A name, B currency, C type, D starting balance, I balance, K balance CHF
function accounts_(ss) {
  const rows = ss.getSheetByName('Accounts').getRange(5, 1, 26, 12).getValues();
  return rows.map((r, i) => ({ r, row: i + 5 })).filter(x => x.r[0] !== '').map(x => ({
    name: String(x.r[0]), currency: String(x.r[1]), type: String(x.r[2]),
    start: num_(x.r[3]), balance: num_(x.r[8]), rate: num_(x.r[9]) || 1, balanceCHF: num_(x.r[10]),
    cushion: x.r[11] === '' ? null : num_(x.r[11]), row: x.row
  }));
}

// Log: A Date, B Merchant, C Item, D Amount, E Currency, F Category, G Note,
// H Manual rate, I Rate used, J Amount CHF, K Month, L Paid from, M Receipt
function logRows_(ss) {
  const sh = ss.getSheetByName('Log');
  const rows = readTable_(sh, 2, 13);
  // The Receipt cell holds =HYPERLINK(url,"View"); read the formula to get the link.
  const links = rows.length ? sh.getRange(2, 13, rows.length, 1).getFormulas() : [];
  return rows
    .map((r, i) => ({ row: i + 2, r, link: links[i][0] || r[12] }))
    .filter(x => x.r[0] instanceof Date && x.r[3] !== '')
    .map(x => ({
      kind: 'expense', row: x.row, date: dayOf_(x.r[0]), merchant: String(x.r[1]), item: String(x.r[2]),
      amount: num_(x.r[3]), currency: String(x.r[4] || 'CHF'), category: String(x.r[5]),
      note: String(x.r[6]), manualRate: x.r[7] === '' ? '' : num_(x.r[7]), chf: num_(x.r[9]),
      account: String(x.r[11]), receipt: receiptUrl_(x.link)
    }));
}

// Income: A Date, B Source, C Type, D Amount, E Currency, F Received into,
// G Note, H Manual rate, I Rate used, J Amount CHF, K Month
function incomeRows_(ss) {
  return readTable_(ss.getSheetByName('Income'), 2, 11)
    .map((r, i) => ({ row: i + 2, r }))
    .filter(x => x.r[0] instanceof Date && x.r[3] !== '')
    .map(x => ({
      kind: 'income', row: x.row, date: dayOf_(x.r[0]), source: String(x.r[1]), type: String(x.r[2]),
      amount: num_(x.r[3]), currency: String(x.r[4] || 'CHF'), account: String(x.r[5]),
      note: String(x.r[6]), manualRate: x.r[7] === '' ? '' : num_(x.r[7]), chf: num_(x.r[9])
    }));
}

// Transfers: A Date, B From, C To, D Amount sent, E Amount received, F Note,
// G From currency, H To currency, I Received used, J Value CHF
function transferRows_(ss) {
  return readTable_(ss.getSheetByName('Transfers'), 2, 12)
    .map((r, i) => ({ row: i + 2, r }))
    .filter(x => x.r[0] instanceof Date && x.r[3] !== '')
    .map(x => ({
      kind: 'transfer', row: x.row, date: dayOf_(x.r[0]), from: String(x.r[1]), to: String(x.r[2]),
      amount: num_(x.r[3]), received: x.r[4] === '' ? '' : num_(x.r[4]), note: String(x.r[5]),
      currency: String(x.r[6] || 'CHF'), toCurrency: String(x.r[7] || 'CHF'),
      receivedUsed: num_(x.r[8]), chf: num_(x.r[9])
    }));
}

function bootstrap_(month) {
  const ss = ss_();
  const s = settings_(ss);
  const log = logRows_(ss);

  // Last-used values per merchant, newest entry wins.
  const sorted = log.slice().sort((a, b) => a.date - b.date || a.row - b.row);
  const merchants = {};
  sorted.forEach(e => {
    if (!e.merchant) return;
    merchants[e.merchant] = { item: e.item, category: e.category, account: e.account, currency: e.currency };
  });
  const lastAccount = sorted.length ? sorted[sorted.length - 1].account : '';

  return {
    currencies: s.currencies,
    categories: s.categories.map(c => c.name),
    budgets: s.categories.map(c => ({ name: c.name, budget: c.budget })),
    items: s.items,
    accounts: accounts_(ss).map(a => ({ name: a.name, currency: a.currency, type: a.type, start: a.start, cushion: a.cushion })),
    accountTypes: ACCOUNT_TYPES,
    lastReconciled: lastReconciled_(accounts_(ss).map(a => a.name)),
    merchants, lastAccount,
    summary: summaryFrom_(ss, month, s, log)
  };
}

function summary_(month) {
  const ss = ss_();
  return summaryFrom_(ss, month, settings_(ss), logRows_(ss));
}

function summaryFrom_(ss, month, s, log) {
  const today = dayOf_(new Date());
  const ym = parseMonth_(month) || { y: today.getFullYear(), m: today.getMonth() };
  const first = new Date(ym.y, ym.m, 1);
  const daysInMonth = new Date(ym.y, ym.m + 1, 0).getDate();
  const inMonth = d => d.getFullYear() === ym.y && d.getMonth() === ym.m;
  const isCurrent = ym.y === today.getFullYear() && ym.m === today.getMonth();

  const monthAll = log.filter(e => inMonth(e.date));
  const monthLog = monthAll.filter(e => e.category !== CORRECTION);   // corrections don't count as spending
  const monthIncomeAll = incomeRows_(ss).filter(e => inMonth(e.date));
  const monthIncome = monthIncomeAll.filter(e => e.type !== CORRECTION);
  const monthTransfers = transferRows_(ss).filter(e => inMonth(e.date));

  // Spending per category
  const spentBy = {};
  monthLog.forEach(e => { spentBy[e.category] = (spentBy[e.category] || 0) + e.chf; });
  const categories = s.categories.map(c => ({ name: c.name, budget: c.budget, spent: round2_(spentBy[c.name] || 0) }));
  Object.keys(spentBy).forEach(name => {           // categories not in Settings
    if (!s.categories.some(c => c.name === name)) categories.push({ name: name || 'Uncategorised', budget: 0, spent: round2_(spentBy[name]) });
  });

  const budgetTotal = s.categories.reduce((t, c) => t + c.budget, 0);
  const spentTotal = monthLog.reduce((t, e) => t + e.chf, 0);
  const incomeTotal = monthIncome.reduce((t, e) => t + e.chf, 0);

  // Spending per day (CHF) for the chart
  const daily = new Array(daysInMonth).fill(0);
  monthLog.forEach(e => { daily[e.date.getDate() - 1] += e.chf; });

  // Days left (including today) and daily allowance, current month only
  let daysLeft = null, dailyAllowance = null;
  if (isCurrent) {
    daysLeft = daysInMonth - today.getDate() + 1;
    dailyAllowance = round2_(Math.max(0, budgetTotal - spentTotal) / daysLeft);
  }

  const salary = nextSalaryDate_(today);
  const accounts = accounts_(ss);

  // All entries of the month, newest first, with what the app needs to edit them.
  const costs = transferCostMap_(ss);
  monthTransfers.forEach(t => { const c = costs[t.row]; if (c) { t.costPct = c.pct; t.costCHF = c.costCHF; } });
  const entries = monthAll.concat(monthIncomeAll, monthTransfers)
    .sort((a, b) => b.date - a.date || (a.kind === b.kind ? b.row - a.row : a.kind < b.kind ? -1 : 1))
    .map(e => {
      const o = Object.assign({}, e);
      o.date = ymd_(e.date);
      o.chf = round2_(e.chf);
      return o;
    });

  return {
    month: ymd_(first).slice(0, 7),
    monthLabel: MONTHS_[first.getMonth()] + ' ' + first.getFullYear(),
    isCurrent, today: ymd_(today), daysInMonth,
    budgetTotal: round2_(budgetTotal), spentTotal: round2_(spentTotal), incomeTotal: round2_(incomeTotal),
    daysLeft, dailyAllowance,
    nextSalary: ymd_(salary),
    daysToSalary: Math.round((salary - today) / 86400000),
    categories, daily: daily.map(round2_), entries,
    accounts: accounts.map(a => ({ name: a.name, currency: a.currency, type: a.type, balance: a.balance, balanceCHF: a.balanceCHF })),
    netWorthCHF: round2_(accounts.reduce((t, a) => t + a.balanceCHF, 0)),
    reservedCHF: round2_(envelopeRows_(ss).reduce((t, e) => t + Math.max(0, e.saved), 0)),
    spendableCHF: round2_(accounts.filter(a => a.type !== 'Credit card').reduce((t, a) => t + a.balanceCHF, 0))
  };
}

// Salary on SALARY_DAY; if that is a weekend, the Friday before.
function nextSalaryDate_(today) {
  const pay = (y, m) => {
    const d = new Date(y, m, SALARY_DAY);
    if (d.getDay() === 6) d.setDate(d.getDate() - 1);
    if (d.getDay() === 0) d.setDate(d.getDate() - 2);
    return d;
  };
  let d = pay(today.getFullYear(), today.getMonth());
  if (d < today) d = pay(today.getFullYear(), today.getMonth() + 1);
  return d;
}

// ---------- Entries: add, edit, delete ----------
function kind_(kind) {
  const k = KINDS[kind];
  if (!k) throw new Error('Unknown entry type');
  return k;
}

function validate_(kind, e) {
  if (!e) throw new Error('Missing entry');
  const need = { expense: ['date', 'amount', 'currency', 'category', 'account'],
                 income: ['date', 'amount', 'currency', 'account'],
                 transfer: ['date', 'from', 'to', 'amount'] }[kind];
  need.forEach(f => { if (e[f] === undefined || e[f] === '' || e[f] === null) throw new Error('Missing ' + f); });
  if (!(Number(e.amount) > 0)) throw new Error('Amount must be more than 0');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(e.date)) throw new Error('Date must be YYYY-MM-DD');
  if (kind === 'transfer' && e.from === e.to) throw new Error('From and To are the same account');
}

// Writes the input cells of one row. Receipt: a new upload replaces it,
// removeReceipt clears it, otherwise the existing one is kept.
function writeEntry_(ss, kind, row, e) {
  const sh = ss.getSheetByName(kind_(kind).sheet);
  const rate = e.manualRate === '' || e.manualRate === undefined || e.manualRate === null ? '' : Number(e.manualRate);
  if (kind === 'expense') {
    sh.getRange(row, 1, 1, 8).setValues([[
      date_(e.date), e.merchant || '', e.item || '', Number(e.amount),
      e.currency, e.category, e.note || '', rate
    ]]);
    sh.getRange(row, 12).setValue(e.account);
    if (e.receipt && e.receipt.data) {
      sh.getRange(row, 13).setFormula('=HYPERLINK("' + saveReceipt_(e) + '","View")');
    } else if (e.removeReceipt) {
      sh.getRange(row, 13).clearContent();
    }
  } else if (kind === 'income') {
    sh.getRange(row, 1, 1, 8).setValues([[
      date_(e.date), e.source || '', e.type || 'Salary', Number(e.amount),
      e.currency, e.account, e.note || '', rate
    ]]);
  } else {
    sh.getRange(row, 1, 1, 6).setValues([[
      date_(e.date), e.from, e.to, Number(e.amount),
      e.received ? Number(e.received) : '', e.note || ''
    ]]);
  }
}

function addEntry_(kind, e) {
  validate_(kind, e);
  const ss = ss_();
  const k = kind_(kind);
  const row = firstEmptyRow_(ss.getSheetByName(k.sheet), [1, 4], k.fStart, k.fCount);
  writeEntry_(ss, kind, row, e);
  SpreadsheetApp.flush();
  return { row, summary: summary_(e.date.slice(0, 7)) };
}

// Before changing a row, check it still holds the entry the app showed.
// This protects you if rows were sorted or edited in the sheet meanwhile.
function checkRow_(sh, row, match) {
  if (!(row >= 2) || row > sh.getMaxRows()) throw new Error('Entry not found. Refresh and try again.');
  const v = sh.getRange(row, 1, 1, 4).getValues()[0];
  const ok = v[0] instanceof Date && match && ymd_(dayOf_(v[0])) === match.date &&
             Math.abs(num_(v[3]) - Number(match.amount)) < 0.005;
  if (!ok) throw new Error('This entry was changed in the sheet. Refresh and try again.');
}

function updateEntry_(kind, row, match, e) {
  validate_(kind, e);
  const ss = ss_();
  const sh = ss.getSheetByName(kind_(kind).sheet);
  checkRow_(sh, row, match);
  writeEntry_(ss, kind, row, e);
  SpreadsheetApp.flush();
  return { row, summary: summary_(e.date.slice(0, 7)) };
}

// Clears only the input cells, so the row's formulas stay and the row is reused.
// A receipt file stays in your Drive folder.
function deleteEntry_(kind, row, match) {
  const ss = ss_();
  const k = kind_(kind);
  const sh = ss.getSheetByName(k.sheet);
  checkRow_(sh, row, match);
  k.inputs.forEach(c => sh.getRange(row, c).clearContent());
  SpreadsheetApp.flush();
  return { summary: summary_(match.date.slice(0, 7)) };
}

/**
 * First row (from row 2) where all the given input columns are empty.
 * If the tab is full, add a row and copy the formulas down.
 */
function firstEmptyRow_(sh, inputCols, formulaStartCol, formulaCount) {
  const last = sh.getMaxRows();
  const width = Math.max.apply(null, inputCols);
  const vals = sh.getRange(2, 1, last - 1, width).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (inputCols.every(c => vals[i][c - 1] === '')) return i + 2;
  }
  sh.insertRowAfter(last);
  const src = sh.getRange(last, formulaStartCol, 1, formulaCount);
  src.copyTo(sh.getRange(last + 1, formulaStartCol, 1, formulaCount));
  return last + 1;
}

// ---------- Accounts ----------
// Where account names are used, so a rename updates everything.
const ACCOUNT_REFS = [['Log', 12], ['Income', 6], ['Transfers', 2], ['Transfers', 3], ['Planned', 7], ['Planned', 17]];
const CATEGORY_REFS = [['Log', 6], ['Planned', 5]];

function saveAccount_(a) {
  if (!a) throw new Error('Missing account');
  const name = String(a.name || '').trim();
  const original = String(a.original || '').trim();
  if (!name) throw new Error('Enter an account name');
  if (ACCOUNT_TYPES.indexOf(a.type) < 0) throw new Error('Choose an account type');
  const ss = ss_();
  const s = settings_(ss);
  if (!s.currencies.some(c => c.code === a.currency)) throw new Error('Choose a currency');
  const start = a.start === '' || a.start === undefined ? 0 : Number(a.start);
  if (!isFinite(start)) throw new Error('Starting balance must be a number');

  const list = accounts_(ss);
  const sh = ss.getSheetByName('Accounts');
  if (list.some(x => x.name.toLowerCase() === name.toLowerCase() && x.name !== original)) {
    throw new Error('An account with this name already exists');
  }

  let row;
  if (original) {
    const cur = list.find(x => x.name === original);
    if (!cur) throw new Error('Account not found. Refresh and try again.');
    row = cur.row;
    if (cur.currency !== a.currency && countRefs_(ss, ACCOUNT_REFS, original) > 0) {
      throw new Error('The currency can\'t change because entries already use this account');
    }
    if (name !== original) renameRefs_(ss, ACCOUNT_REFS, original, name);
  } else {
    const rows = sh.getRange(5, 1, 26, 1).getValues();
    const i = rows.findIndex(r => r[0] === '');
    if (i < 0) throw new Error('The Accounts tab is full (26 accounts)');
    row = i + 5;
  }
  sh.getRange(row, 1, 1, 4).setValues([[name, a.currency, a.type, start]]);
  const cushion = a.cushion === '' || a.cushion === undefined || a.cushion === null ? '' : Number(a.cushion);
  if (cushion !== '' && !isFinite(cushion)) throw new Error('Safety cushion must be a number');
  sh.getRange(row, 12).setValue(cushion);
  return { accounts: accounts_(ss).map(x => ({ name: x.name, currency: x.currency, type: x.type, start: x.start, cushion: x.cushion })) };
}

function deleteAccount_(name) {
  const ss = ss_();
  const cur = accounts_(ss).find(x => x.name === name);
  if (!cur) throw new Error('Account not found');
  const used = countRefs_(ss, ACCOUNT_REFS, name);
  if (used > 0) throw new Error('This account is used by ' + used + ' entries, so it can\'t be deleted. You can rename it instead.');
  ss.getSheetByName('Accounts').getRange(cur.row, 1, 1, 4).clearContent();
  ss.getSheetByName('Accounts').getRange(cur.row, 12).clearContent();
  return { accounts: accounts_(ss).map(x => ({ name: x.name, currency: x.currency, type: x.type, start: x.start, cushion: x.cushion })) };
}

// ---------- Categories and budgets ----------
function saveCategory_(c) {
  if (!c) throw new Error('Missing category');
  const name = String(c.name || '').trim();
  const original = String(c.original || '').trim();
  const budget = c.budget === '' || c.budget === undefined ? 0 : Number(c.budget);
  if (!name) throw new Error('Enter a category name');
  if (!(budget >= 0)) throw new Error('Budget must be 0 or more');
  const ss = ss_();
  const sh = ss.getSheetByName('Settings');
  const cats = settings_(ss).categories;
  if (cats.some(x => x.name.toLowerCase() === name.toLowerCase() && x.name !== original)) {
    throw new Error('A category with this name already exists');
  }
  let row;
  if (original) {
    const cur = cats.find(x => x.name === original);
    if (!cur) throw new Error('Category not found. Refresh and try again.');
    row = cur.row;
    if (name !== original) {
      renameRefs_(ss, CATEGORY_REFS, original, name);
      renameInRange_(sh, 4, 10, original, name);   // Items: usual category
    }
  } else {
    const vals = sh.getRange(4, 6, 200, 1).getValues();
    row = vals.findIndex(r => r[0] === '') + 4;
  }
  sh.getRange(row, 6, 1, 2).setValues([[name, budget]]);
  return { budgets: settings_(ss).categories.map(x => ({ name: x.name, budget: x.budget })) };
}

function deleteCategory_(name) {
  const ss = ss_();
  const cur = settings_(ss).categories.find(x => x.name === name);
  if (!cur) throw new Error('Category not found');
  const used = countRefs_(ss, CATEGORY_REFS, name);
  if (used > 0) throw new Error('This category is used by ' + used + ' entries, so it can\'t be deleted. You can rename it instead.');
  ss.getSheetByName('Settings').getRange(cur.row, 6, 1, 2).clearContent();
  return { budgets: settings_(ss).categories.map(x => ({ name: x.name, budget: x.budget })) };
}

// ---------- Rename helpers ----------
function countRefs_(ss, refs, value) {
  let n = 0;
  refs.forEach(([tab, col]) => {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0] === value) n++; });
  });
  return n;
}

function renameRefs_(ss, refs, from, to) {
  refs.forEach(([tab, col]) => {
    const sh = ss.getSheetByName(tab);
    if (sh) renameInRange_(sh, 2, col, from, to);
  });
}

function renameInRange_(sh, firstRow, col, from, to) {
  const last = sh.getLastRow();
  if (last < firstRow) return;
  const rng = sh.getRange(firstRow, col, last - firstRow + 1, 1);
  const vals = rng.getValues();
  let changed = false;
  vals.forEach(r => { if (r[0] === from) { r[0] = to; changed = true; } });
  if (changed) rng.setValues(vals);
}

// ---------- Receipts ----------
function saveReceipt_(e) {
  const r = e.receipt;
  const bytes = Utilities.base64Decode(r.data);
  if (bytes.length > 10 * 1024 * 1024) throw new Error('Receipt is larger than 10 MB');
  const ext = r.mimeType === 'application/pdf' ? 'pdf' : 'jpg';
  const label = [e.date, e.merchant || e.category, e.amount + ' ' + e.currency]
    .join(' ').replace(/[\\/:*?"<>|]/g, '-');
  const blob = Utilities.newBlob(bytes, r.mimeType || 'image/jpeg', label + '.' + ext);
  return receiptsFolder_().createFile(blob).getUrl();
}

function receiptsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('RECEIPTS_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (err) { /* deleted: make a new one */ }
  }
  const folder = DriveApp.createFolder(RECEIPTS_FOLDER);
  props.setProperty('RECEIPTS_FOLDER_ID', folder.getId());
  return folder;
}

// ---------- Helpers ----------
// Noon avoids any timezone shift moving the date to the day before.
function date_(s) { const p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12); }
// All day calculations use the Zurich calendar day, whatever time zone the
// Apps Script project is set to. dayOf_ turns any moment into that day
// (as midnight in the script's own time zone, so getDate()/getMonth() work).
function dayOf_(d) {
  const s = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return new Date(Number(s.slice(0, 4)), Number(s.slice(5, 7)) - 1, Number(s.slice(8, 10)));
}
// Only for days already made by dayOf_ (sheet values and "now" go through dayOf_ first).
function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
// For days made by dayOf_/addDays_/addMonths_ (not raw sheet values).
function ymd_(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
const MONTHS_ = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS_ = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
function longDay_(d) { return DAYS_[d.getDay()] + ' ' + d.getDate() + ' ' + MONTHS_[d.getMonth()].slice(0, 3) + ' ' + d.getFullYear(); }
function parseMonth_(s) {
  if (!s || !/^\d{4}-\d{2}$/.test(s)) return null;
  return { y: Number(s.slice(0, 4)), m: Number(s.slice(5, 7)) - 1 };
}
function num_(v) { return typeof v === 'number' ? v : (Number(v) || 0); }
function round2_(n) { return Math.round(n * 100) / 100; }
function receiptUrl_(v) {
  if (!v) return '';
  const m = String(v).match(/https?:\/\/[^"\s]+/);
  return m ? m[0] : '';
}

// =====================================================================
// Phase 2: recurring payments, forecast, gaps, search, category detail
// =====================================================================

// Planned tab: A First due, B Description, C Amount (minus = money in), D Currency,
// E Category, F Repeats, G Pay from, H Status, I Note, J-M formulas,
// N Last paid, O Ends on, P If weekend, Q To account
const PLANNED = { sheet: 'Planned', inputs: [1, 2, 3, 4, 5, 6, 7, 8, 9, 14, 15, 16, 17, 18], fStart: 10, fCount: 4 };
const REPEATS = { 'Once': 0, 'Monthly': 1, 'Every 2 months': 2, 'Quarterly': 3, 'Every 6 months': 6, 'Yearly': 12 };
const WEEKEND = ['No change', 'Earlier', 'Later'];
const FORECAST_MONTHS = 24;

// Adds the new columns the first time the app runs on a sheet. Safe to run again.
function ensureSchema_() {
  ensureSchema2_();
  ensureSchema3_();
}

// Column R in Planned: "Pay full balance" for transfers that pay off a card.
function ensureSchema3_() {
  const props = PropertiesService.getScriptProperties();
  const flag = 'SCHEMA3_' + SHEET_ID;
  if (props.getProperty(flag)) return;
  const pl = ss_().getSheetByName('Planned');
  if (pl.getMaxColumns() < 18) pl.insertColumnsAfter(pl.getMaxColumns(), 18 - pl.getMaxColumns());
  if (pl.getRange(1, 18).getValue() !== 'Pay full balance') {
    try { pl.getRange(1, 17).copyTo(pl.getRange(1, 18), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (err) { /* keep default format */ }
    pl.getRange(1, 18).setValue('Pay full balance');
  }
  props.setProperty(flag, '1');
}

function ensureSchema2_() {
  const props = PropertiesService.getScriptProperties();
  const flag = 'SCHEMA2_' + SHEET_ID;
  if (props.getProperty(flag)) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (props.getProperty(flag)) return;
    const ss = ss_();
    const pl = ss.getSheetByName('Planned');
    if (pl.getMaxColumns() < 17) pl.insertColumnsAfter(pl.getMaxColumns(), 17 - pl.getMaxColumns());
    if (pl.getRange(1, 14).getValue() !== 'Last paid') {
      pl.getRange(1, 9).copyTo(pl.getRange(1, 14, 1, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      pl.getRange(1, 14, 1, 4).setValues([['Last paid', 'Ends on', 'If weekend', 'To account']]);
      // Rows marked Paid: the first due date counts as paid.
      const last = pl.getLastRow();
      if (last >= 2) {
        const v = pl.getRange(2, 1, last - 1, 8).getValues();
        const lp = v.map(r => [r[0] instanceof Date && String(r[7]) === 'Paid' ? r[0] : '']);
        pl.getRange(2, 14, lp.length, 1).setValues(lp);
      }
      addSalaryRow_(ss, pl);
    }
    const acc = ss.getSheetByName('Accounts');
    if (acc.getRange(4, 12).getValue() !== 'Safety cushion') {
      acc.getRange(4, 11).copyTo(acc.getRange(4, 12), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      acc.getRange(4, 12).setValue('Safety cushion');
    }
    SpreadsheetApp.flush();
    props.setProperty(flag, '1');
  } finally {
    lock.releaseLock();
  }
}

// Adds your salary as a monthly "money in" item, from the latest Salary in the Income tab.
function addSalaryRow_(ss, pl) {
  const salary = incomeRows_(ss).filter(e => e.type === 'Salary').sort((a, b) => b.date - a.date)[0];
  if (!salary) return;
  const rows = plannedRows_(ss);
  if (rows.some(r => r.amount < 0 && r.repeats === 'Monthly')) return;
  const row = firstEmptyRow_(pl, [1, 3], PLANNED.fStart, PLANNED.fCount);
  pl.getRange(row, 1, 1, 9).setValues([[date_(ymd_(salary.date)), 'Salary ' + salary.source, -salary.amount, salary.currency,
    'Salary', 'Monthly', salary.account, 'Planned', '']]);
  pl.getRange(row, 14, 1, 3).setValues([[date_(ymd_(salary.date)), '', 'Earlier']]);
}

function plannedRows_(ss) {
  const sh = ss.getSheetByName('Planned');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, 18).getValues();
  return v.map((r, i) => ({ r, row: i + 2 }))
    .filter(x => x.r[0] instanceof Date && x.r[2] !== '' && x.r[1] !== '')
    .map(x => ({
      row: x.row, first: dayOf_(x.r[0]), description: String(x.r[1]), amount: num_(x.r[2]),
      currency: String(x.r[3] || 'CHF'), category: String(x.r[4]), repeats: REPEATS[x.r[5]] !== undefined ? String(x.r[5]) : 'Once',
      account: String(x.r[6]), status: String(x.r[7]), note: String(x.r[8]),
      lastPaid: x.r[13] instanceof Date ? dayOf_(x.r[13]) : null, endsOn: x.r[14] instanceof Date ? dayOf_(x.r[14]) : null,
      weekend: WEEKEND.indexOf(String(x.r[15])) >= 0 ? String(x.r[15]) : 'No change',
      toAccount: String(x.r[16] || ''),
      fullBalance: !!x.r[16] && (x.r[17] === true || /^(yes|true|1)$/i.test(String(x.r[17]).trim()))
    }));
}

function kindOfPlanned_(p) { return p.toAccount ? 'transfer' : p.amount < 0 ? 'income' : 'expense'; }

function shiftWeekend_(d, rule) {
  const day = d.getDay();
  if (rule === 'Earlier') { if (day === 6) return addDays_(d, -1); if (day === 0) return addDays_(d, -2); }
  if (rule === 'Later') { if (day === 6) return addDays_(d, 2); if (day === 0) return addDays_(d, 1); }
  return d;
}

// Scheduled dates of one item between from and to (inclusive), before weekend shifting.
function occurrences_(p, from, to) {
  const out = [];
  const step = REPEATS[p.repeats];
  const first = startOfDay_(p.first);
  if (!step) { if (first >= from && first <= to) out.push(first); return out; }
  for (let i = 0; i < 1000; i++) {
    const d = addMonths_(first, i * step);
    if (d > to || (p.endsOn && d > startOfDay_(p.endsOn))) break;
    if (d >= from) out.push(d);
  }
  return out;
}

// Payments still open: not yet marked paid, from up to 90 days ago until `to`.
function openOccurrences_(p, today, to) {
  if (p.status === 'Paid' && p.repeats === 'Once') return [];
  // Cancelled (your sheet's own status) and Stopped items are left out of the forecast.
  if (p.status === 'Skipped' || p.status === 'Stopped' || p.status === 'Cancelled') return [];
  const lookBack = addDays_(today, -90);
  const after = p.lastPaid ? addDays_(startOfDay_(p.lastPaid), 1) : lookBack;
  const from = after > lookBack ? after : lookBack;
  return occurrences_(p, from, to).map(due => ({ due, date: shiftWeekend_(due, p.weekend) }));
}

function plan_() {
  const ss = ss_();
  const today = dayOf_(new Date());
  const end = addMonths_(new Date(today.getFullYear(), today.getMonth(), 1), FORECAST_MONTHS);
  end.setDate(end.getDate() - 1);
  const s = settings_(ss);
  const rates = {}; s.currencies.forEach(c => { rates[c.code] = c.rate || 1; });
  const rate = c => rates[c] || 1;
  const accounts = accounts_(ss);
  const accBy = {}; accounts.forEach(a => { accBy[a.name] = a; });
  const planned = plannedRows_(ss);

  // ---- Coming up: the next 45 days (and unpaid overdue items) ----
  const upcoming = [];
  const in45 = addDays_(today, 45);
  planned.forEach(p => {
    const kind = kindOfPlanned_(p);
    openOccurrences_(p, today, in45).forEach(o => {
      const u = { row: p.row, description: p.description, kind, due: ymd_(o.due), date: ymd_(o.date), fullBalance: p.fullBalance,
        overdue: o.date < today, amount: Math.abs(p.amount), currency: p.currency, category: p.category,
        account: p.account, toAccount: p.toAccount, chf: round2_(Math.abs(p.amount) * rate(p.currency)), repeats: p.repeats };
      if (p.fullBalance) {
        // What the card owes today, in the paying account's currency.
        const to = accBy[p.toAccount], from = accBy[p.account];
        u.amount = to && from ? round2_(Math.max(0, -to.balance) * to.rate / from.rate) : 0;
        u.currency = from ? from.currency : p.currency;
        u.chf = round2_(u.amount * (from ? from.rate : 1)); u.estimated = true;
      }
      upcoming.push(u);
    });
  });
  upcoming.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : a.description < b.description ? -1 : 1);

  const monthly = monthlyPlan_(ss, { today, end, planned, rate, settings: s, accounts });
  return {
    today: ymd_(today),
    upcoming,
    monthly,
    budget: monthly.budget,
    recurring: planned.map(p => ({
      row: p.row, first: ymd_(p.first), description: p.description, amount: p.amount, currency: p.currency,
      category: p.category, repeats: p.repeats, account: p.account, status: p.status, note: p.note,
      lastPaid: p.lastPaid ? ymd_(p.lastPaid) : '', endsOn: p.endsOn ? ymd_(p.endsOn) : '',
      weekend: p.weekend, toAccount: p.toAccount, kind: kindOfPlanned_(p), fullBalance: p.fullBalance,
      next: (openOccurrences_(p, today, addMonths_(today, 25))[0] || {}).date ? ymd_(openOccurrences_(p, today, addMonths_(today, 25))[0].date) : ''
    })),
    repeatsOptions: Object.keys(REPEATS), weekendOptions: WEEKEND,
    envelopes: envelopes_(ss)
  };
}

// Cushion: the number you set; if empty, 0 for normal accounts and none for credit cards.
function cushionOf_(a) {
  if (a.cushion !== null && a.cushion !== undefined) return a.cushion;
  return a.type === 'Credit card' ? null : 0;
}

// ---- Recurring items: add / edit / delete ----
function saveRecurring_(it) {
  if (!it) throw new Error('Missing item');
  const desc = String(it.description || '').trim();
  if (!desc) throw new Error('Enter a description');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(it.first || '')) throw new Error('Choose the first due date');
  const kind = it.kind;
  const full = kind === 'transfer' && !!it.fullBalance;
  const amount = full ? (Number(it.amount) > 0 ? Number(it.amount) : 0) : Number(it.amount);
  if (!full && !(amount > 0)) throw new Error('Amount must be more than 0');
  if (REPEATS[it.repeats] === undefined) throw new Error('Choose how often it repeats');
  if (['expense', 'income', 'transfer'].indexOf(kind) < 0) throw new Error('Choose money out, money in or transfer');
  if (!it.account) throw new Error('Choose an account');
  if (kind === 'transfer' && (!it.toAccount || it.toAccount === it.account)) throw new Error('Choose a different To account');
  if (kind === 'expense' && !it.category) throw new Error('Choose a category');
  const ss = ss_();
  const sh = ss.getSheetByName('Planned');
  let row = it.row;
  if (row) {
    const cur = plannedRows_(ss).find(p => p.row === row);
    if (!cur || cur.description !== it.original) throw new Error('This item was changed in the sheet. Refresh and try again.');
  } else {
    row = firstEmptyRow_(sh, [1, 2, 3], PLANNED.fStart, PLANNED.fCount);
  }
  const signed = kind === 'income' ? -amount : amount;
  sh.getRange(row, 1, 1, 9).setValues([[date_(it.first), desc, signed, it.currency || 'CHF',
    kind === 'transfer' ? '' : (it.category || ''), it.repeats, it.account, it.status || 'Planned', it.note || '']]);
  sh.getRange(row, 14, 1, 5).setValues([[it.lastPaid ? date_(it.lastPaid) : '', it.endsOn ? date_(it.endsOn) : '',
    WEEKEND.indexOf(it.weekend) >= 0 ? it.weekend : 'No change', kind === 'transfer' ? it.toAccount : '', full ? 'Yes' : '']]);
  SpreadsheetApp.flush();
  return plan_();
}

function deleteRecurring_(row, match) {
  const ss = ss_();
  const cur = plannedRows_(ss).find(p => p.row === row);
  if (!cur || !match || cur.description !== match.description) throw new Error('This item was changed in the sheet. Refresh and try again.');
  const sh = ss.getSheetByName('Planned');
  PLANNED.inputs.forEach(c => sh.getRange(row, c).clearContent());
  SpreadsheetApp.flush();
  return plan_();
}

// ---- Mark a planned payment as paid: records it and moves the item on ----
function markPaid_(pay) {
  if (!pay) throw new Error('Missing payment');
  const ss = ss_();
  const p = plannedRows_(ss).find(x => x.row === pay.row);
  if (!p || p.description !== pay.description) throw new Error('This item was changed in the sheet. Refresh and try again.');
  const kind = kindOfPlanned_(p);
  const date = pay.date || ymd_(dayOf_(new Date()));
  let amount = Number(pay.amount || Math.abs(p.amount));
  if (!(amount > 0) && p.fullBalance) {
    const accs = accounts_(ss), to = accs.find(a => a.name === p.toAccount), from = accs.find(a => a.name === p.account);
    amount = to && from ? round2_(Math.max(0, -to.balance) * to.rate / from.rate) : 0;
  }
  if (!(amount > 0)) throw new Error('Enter the amount you paid');
  const currency = pay.currency || p.currency;
  const account = pay.account || p.account;
  let entry;
  if (kind === 'expense') {
    entry = { date, amount, currency, category: p.category || 'Other', merchant: p.description, item: '', account, note: pay.note || '' };
  } else if (kind === 'income') {
    entry = { date, amount, currency, source: p.description, type: p.category || 'Other', account, note: pay.note || '' };
  } else {
    entry = { date, amount, from: account, to: p.toAccount, received: pay.received || '', note: pay.note || p.description };
  }
  const added = addEntry_(kind, entry);
  setPaidThrough_(ss, p, pay.due, 'Paid');
  if (kind === 'expense') useLinkedEnvelope_(ss, p, amount * currencyRate_(ss, currency));
  SpreadsheetApp.flush();
  return { summary: added.summary, plan: plan_() };
}

function skipPayment_(row, due, match) {
  const ss = ss_();
  const p = plannedRows_(ss).find(x => x.row === row);
  if (!p || !match || p.description !== match.description) throw new Error('This item was changed in the sheet. Refresh and try again.');
  setPaidThrough_(ss, p, due, 'Skipped');
  SpreadsheetApp.flush();
  return plan_();
}

function setPaidThrough_(ss, p, due, onceStatus) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(due || '')) throw new Error('Missing due date');
  const sh = ss.getSheetByName('Planned');
  const d = date_(due);
  if (!p.lastPaid || d > p.lastPaid) {
    sh.getRange(p.row, 14).setValue(d);   // the table's Date column formats it
  }
  if (p.repeats === 'Once') sh.getRange(p.row, 8).setValue(onceStatus);
}

// ---- Search across all years ----
function search_(q) {
  const ss = ss_();
  const text = String(q.text || '').trim().toLowerCase();
  const from = q.from ? date_(q.from) : null, to = q.to ? addDays_(date_(q.to), 1) : null;
  const min = q.min === '' || q.min == null ? null : Number(q.min);
  const max = q.max === '' || q.max == null ? null : Number(q.max);
  const all = logRows_(ss).concat(incomeRows_(ss), transferRows_(ss));
  const hits = all.filter(e => {
    if (q.kind && e.kind !== q.kind) return false;
    if (q.category && e.category !== q.category) return false;
    if (q.account && e.account !== q.account && e.from !== q.account && e.to !== q.account) return false;
    if (from && e.date < from) return false;
    if (to && e.date >= to) return false;
    if (min !== null && e.amount < min) return false;
    if (max !== null && e.amount > max) return false;
    if (text) {
      const hay = [e.merchant, e.item, e.note, e.source, e.type, e.category, e.account, e.from, e.to, String(e.amount)]
        .filter(Boolean).join(' ').toLowerCase();
      if (text.split(/\s+/).some(w => hay.indexOf(w) < 0)) return false;
    }
    return true;
  }).sort((a, b) => b.date - a.date);
  const totalCHF = hits.filter(e => e.kind === 'expense').reduce((t, e) => t + e.chf, 0);
  return {
    count: hits.length, totalCHF: round2_(totalCHF),
    entries: hits.slice(0, 200).map(e => { const o = Object.assign({}, e); o.date = ymd_(e.date); o.chf = round2_(e.chf); return o; })
  };
}

// ---- One category: last 12 months and this month's entries ----
function categoryDetail_(name, month) {
  const ss = ss_();
  const s = settings_(ss);
  const today = dayOf_(new Date());
  const ym = parseMonth_(month) || { y: today.getFullYear(), m: today.getMonth() };
  const cat = s.categories.find(c => c.name === name);
  const log = logRows_(ss).filter(e => e.category === name);
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(ym.y, ym.m - i, 1);
    const total = log.filter(e => e.date.getFullYear() === d.getFullYear() && e.date.getMonth() === d.getMonth()).reduce((t, e) => t + e.chf, 0);
    months.push({ month: ymd_(d).slice(0, 7), total: round2_(total) });
  }
  const withData = months.filter(m => m.total > 0);
  const entries = log.filter(e => e.date.getFullYear() === ym.y && e.date.getMonth() === ym.m)
    .sort((a, b) => b.date - a.date)
    .map(e => { const o = Object.assign({}, e); o.date = ymd_(e.date); o.chf = round2_(e.chf); return o; });
  const merchants = {};
  log.filter(e => e.date >= new Date(ym.y, ym.m - 11, 1)).forEach(e => {
    const k = e.merchant || e.item || '(no merchant)'; merchants[k] = (merchants[k] || 0) + e.chf;
  });
  return {
    name, budget: cat ? cat.budget : 0, months, entries,
    average: round2_(withData.length ? withData.reduce((t, m) => t + m.total, 0) / withData.length : 0),
    topMerchants: Object.keys(merchants).map(k => ({ name: k, chf: round2_(merchants[k]) })).sort((a, b) => b.chf - a.chf).slice(0, 5)
  };
}

function addDays_(d, n) { const x = new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); return x; }
// Adds months, keeping the day (the 31st becomes the last day of a shorter month).
function addMonths_(d, n) {
  const y = d.getFullYear(), m = d.getMonth() + n, day = d.getDate();
  const last = new Date(y, m + 1, 0).getDate();
  return new Date(y, m, Math.min(day, last));
}


// =====================================================================
// Part B: reconcile, transfer cost, sinking funds and savings goals
// =====================================================================
const CORRECTION = 'Balance correction';

function currencyRate_(ss, code) {
  const c = settings_(ss).currencies.find(x => x.code === code);
  return c ? c.rate : 1;
}

// ---- Reconcile: compare the app's balance with what the bank shows ----
function reconcile_(name, bankBalance, apply) {
  const ss = ss_();
  const a = accounts_(ss).find(x => x.name === name);
  if (!a) throw new Error('Account not found');
  const bank = Number(bankBalance);
  if (!isFinite(bank) || bankBalance === '' || bankBalance === null) throw new Error('Enter the balance your bank shows');
  const diff = round2_(bank - a.balance);
  const props = PropertiesService.getScriptProperties();
  const key = 'RECON_' + SHEET_ID + '_' + name;
  const today = ymd_(dayOf_(new Date()));
  if (apply) {
    if (Math.abs(diff) >= 0.005) {
      if (diff < 0) {
        addEntry_('expense', { date: today, amount: -diff, currency: a.currency, category: CORRECTION,
          merchant: CORRECTION, item: '', account: name, note: 'Bank showed ' + bank });
      } else {
        addEntry_('income', { date: today, amount: diff, currency: a.currency, source: CORRECTION,
          type: CORRECTION, account: name, note: 'Bank showed ' + bank });
      }
    }
    props.setProperty(key, JSON.stringify({ date: today, bank }));
    SpreadsheetApp.flush();
  }
  const last = JSON.parse(props.getProperty(key) || 'null');
  return { account: name, currency: a.currency, appBalance: round2_(a.balance), bankBalance: bank, difference: diff,
    lastChecked: last ? last.date : '', applied: !!apply };
}

function lastReconciled_(names) {
  const props = PropertiesService.getScriptProperties();
  const out = {};
  names.forEach(n => { const v = props.getProperty('RECON_' + SHEET_ID + '_' + n); if (v) out[n] = JSON.parse(v).date; });
  return out;
}

// ---- Transfer cost: what arrived compared with the market rate that day ----
function ratesTable_(ss) {
  const sh = ss.getSheetByName('Rates');
  const out = { EGP: [], USD: [] };
  if (!sh || sh.getLastRow() < 3) return out;
  const v = sh.getRange(3, 1, sh.getLastRow() - 2, 5).getValues();
  v.forEach(r => {
    if (r[0] instanceof Date && typeof r[1] === 'number') out.EGP.push([ymd_(dayOf_(r[0])), r[1]]);
    if (r[3] instanceof Date && typeof r[4] === 'number') out.USD.push([ymd_(dayOf_(r[3])), r[4]]);
  });
  return out;
}

// Rate to CHF on a date: the last closing rate on or before it (like the sheet's VLOOKUP).
function marketRate_(table, live, code, ymd) {
  if (code === 'CHF') return { rate: 1, exact: true };
  const t = table[code] || [];
  let lo = 0, hi = t.length - 1, found = -1;
  while (lo <= hi) { const mid = (lo + hi) >> 1; if (t[mid][0] <= ymd) { found = mid; lo = mid + 1; } else hi = mid - 1; }
  if (found >= 0) return { rate: t[found][1], exact: true };
  return { rate: live[code] || 1, exact: false };
}

function transferCostMap_(ss) {
  const table = ratesTable_(ss);
  const live = {}; settings_(ss).currencies.forEach(c => { live[c.code] = c.rate; });
  const map = {};
  transferRows_(ss).forEach(t => {
    if (t.received === '' || !(t.received > 0)) return;
    const d = ymd_(t.date);
    const rf = marketRate_(table, live, t.currency, d), rt = marketRate_(table, live, t.toCurrency, d);
    const expected = t.amount * rf.rate / rt.rate;          // what should have arrived, in the To currency
    const cost = expected - t.received;
    map[t.row] = {
      row: t.row, date: d, from: t.from, to: t.to, sent: t.amount, sentCurrency: t.currency,
      received: t.received, receivedCurrency: t.toCurrency, expected: round2_(expected),
      bankRate: round6_(t.received / t.amount), marketRate: round6_(rf.rate / rt.rate),
      costCHF: round2_(cost * rt.rate), pct: round2_(expected ? cost / expected * 100 : 0),
      approximate: !(rf.exact && rt.exact), note: t.note
    };
  });
  return map;
}

function transferCosts_() {
  const ss = ss_();
  const map = transferCostMap_(ss);
  const list = Object.keys(map).map(k => map[k]);
  return transferCostSummary_(list);
}
function transferCostSummary_(list) {
  list.sort((a, b) => a.date < b.date ? 1 : -1);
  const year = String(new Date().getFullYear());
  const thisYear = list.filter(x => x.date.slice(0, 4) === year);
  const routes = {};
  list.forEach(x => { routes[x.from + ' → ' + x.to] = true; });
  const routeList = Object.keys(routes).map(k => {
    const items = list.filter(x => x.from + ' → ' + x.to === k);
    const avg = items.reduce((t, x) => t + x.pct, 0) / items.length;
    return { route: k, count: items.length, costCHF: round2_(items.reduce((t, x) => t + x.costCHF, 0)), avgPct: round2_(avg) };
  }).sort((a, b) => a.avgPct - b.avgPct);
  return {
    transfers: list.slice(0, 50),
    yearCostCHF: round2_(thisYear.reduce((t, x) => t + x.costCHF, 0)),
    yearCount: thisYear.length,
    avgPct: round2_(list.length ? list.reduce((t, x) => t + x.pct, 0) / list.length : 0),
    routes: routeList
  };
}

// ---- Envelopes: sinking funds and savings goals (money stays in your accounts) ----
// Envelopes tab: A Name, B Type, C Target (CHF), D Target date, E Monthly amount (optional),
// F Saved so far (CHF), G Linked planned item, H Last set aside (month), I Note
const ENV_HEADERS = ['Name', 'Type', 'Target CHF', 'Target date', 'Monthly amount CHF (optional)', 'Saved so far CHF', 'Linked planned item', 'Last set aside', 'Note'];
const ENV_TYPES = ['Sinking fund', 'Goal'];

function envSheet_(ss) {
  let sh = ss.getSheetByName('Envelopes');
  if (!sh) {
    sh = ss.insertSheet('Envelopes');
    sh.getRange(1, 1, 1, ENV_HEADERS.length).setValues([ENV_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    format_(sh.getRange('D:D'), 'yyyy-mm-dd');
    format_(sh.getRange('C:C'), '#,##0.00');
    format_(sh.getRange('E:F'), '#,##0.00');
  }
  return sh;
}

function envelopeRows_(ss) {
  const sh = ss.getSheetByName('Envelopes');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues()
    .map((r, i) => ({ r, row: i + 2 })).filter(x => x.r[0] !== '')
    .map(x => ({
      row: x.row, name: String(x.r[0]), type: ENV_TYPES.indexOf(String(x.r[1])) >= 0 ? String(x.r[1]) : 'Goal',
      target: num_(x.r[2]), targetDate: x.r[3] instanceof Date ? dayOf_(x.r[3]) : null,
      monthly: x.r[4] === '' ? null : num_(x.r[4]), saved: num_(x.r[5]), linked: String(x.r[6] || ''),
      lastSetAside: x.r[7] instanceof Date ? ymd_(dayOf_(x.r[7])).slice(0, 7) : String(x.r[7] || '').replace(/^'/, ''),
      note: String(x.r[8] || '')
    }));
}

function monthsUntil_(from, to) {
  return (to.getFullYear() - from.getFullYear()) * 12 + (to.getMonth() - from.getMonth());
}

function envelopes_(ss) {
  const today = dayOf_(new Date());
  const thisMonth = ymd_(today).slice(0, 7);
  const planned = plannedRows_(ss);
  const list = envelopeRows_(ss).map(e => {
    const left = Math.max(0, e.target - e.saved);
    // months to save, counting this month if nothing was set aside yet this month
    let months = e.targetDate ? monthsUntil_(today, e.targetDate) + (e.lastSetAside === thisMonth ? 0 : 1) : null;
    if (months !== null) months = Math.max(1, months);
    const needed = left <= 0 ? 0 : months ? left / months : null;
    return {
      row: e.row, name: e.name, type: e.type, target: round2_(e.target), saved: round2_(e.saved),
      targetDate: e.targetDate ? ymd_(e.targetDate) : '', monthly: e.monthly, linked: e.linked, note: e.note,
      lastSetAside: e.lastSetAside, setAsideThisMonth: e.lastSetAside === thisMonth,
      left: round2_(left), monthsLeft: months, neededMonthly: needed === null ? null : round2_(needed),
      suggested: round2_(e.monthly !== null ? e.monthly : needed || 0),
      pct: e.target ? Math.min(100, Math.round(e.saved / e.target * 100)) : 0,
      onTrack: needed === null || e.monthly === null || e.monthly >= needed - 0.005
    };
  });
  // Bills that repeat less often than monthly and have no fund yet.
  const linked = {}; list.forEach(e => { if (e.linked) linked[e.linked] = true; });
  const suggestions = planned.filter(p => kindOfPlanned_(p) === 'expense' && REPEATS[p.repeats] >= 3 && !linked[p.description])
    .map(p => {
      const next = (openOccurrences_(p, today, addMonths_(today, 25))[0] || {}).date;
      return next ? { description: p.description, amountCHF: round2_(p.amount * currencyRate_(ss, p.currency)), next: ymd_(next), repeats: p.repeats } : null;
    }).filter(Boolean);
  return {
    envelopes: list, suggestions, types: ENV_TYPES,
    reservedCHF: round2_(list.reduce((t, e) => t + Math.max(0, e.saved), 0)),
    monthlyCHF: round2_(list.reduce((t, e) => t + (e.left > 0 ? e.suggested : 0), 0))
  };
}

function saveEnvelope_(v) {
  if (!v) throw new Error('Missing envelope');
  const name = String(v.name || '').trim();
  if (!name) throw new Error('Enter a name');
  if (ENV_TYPES.indexOf(v.type) < 0) throw new Error('Choose sinking fund or goal');
  const target = Number(v.target);
  if (!(target > 0)) throw new Error('Target must be more than 0');
  if (v.targetDate && !/^\d{4}-\d{2}-\d{2}$/.test(v.targetDate)) throw new Error('Choose a target date');
  const saved = v.saved === '' || v.saved == null ? 0 : Number(v.saved);
  if (!isFinite(saved) || saved < 0) throw new Error('Saved so far must be 0 or more');
  const monthly = v.monthly === '' || v.monthly == null ? '' : Number(v.monthly);
  if (monthly !== '' && !(monthly >= 0)) throw new Error('Monthly amount must be 0 or more');
  const ss = ss_();
  const sh = envSheet_(ss);
  const rows = envelopeRows_(ss);
  if (rows.some(e => e.name.toLowerCase() === name.toLowerCase() && e.row !== v.row)) throw new Error('An envelope with this name already exists');
  let row = v.row;
  if (row) {
    const cur = rows.find(e => e.row === row);
    if (!cur || cur.name !== v.original) throw new Error('This envelope was changed in the sheet. Refresh and try again.');
  } else {
    const vals = sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), 1).getValues();
    const i = vals.findIndex(r => r[0] === '');
    row = i >= 0 ? i + 2 : sh.getMaxRows() + 1;
  }
  const cur = rows.find(e => e.row === row);
  sh.getRange(row, 1, 1, 9).setValues([[name, v.type, target, v.targetDate ? date_(v.targetDate) : '', monthly, saved,
    v.linked || '', cur ? cur.lastSetAside : '', v.note || '']]);
  SpreadsheetApp.flush();
  return envelopes_(ss);
}

function deleteEnvelope_(row, match) {
  const ss = ss_();
  const cur = envelopeRows_(ss).find(e => e.row === row);
  if (!cur || !match || cur.name !== match.name) throw new Error('This envelope was changed in the sheet. Refresh and try again.');
  ss.getSheetByName('Envelopes').getRange(row, 1, 1, 9).clearContent();
  SpreadsheetApp.flush();
  return envelopes_(ss);
}

// Set money aside (positive) or take it out (negative). Nothing moves between accounts.
function moveEnvelope_(row, match, amount, reason) {
  const ss = ss_();
  const cur = envelopeRows_(ss).find(e => e.row === row);
  if (!cur || !match || cur.name !== match.name) throw new Error('This envelope was changed in the sheet. Refresh and try again.');
  const n = Number(amount);
  if (!isFinite(n) || n === 0) throw new Error('Enter an amount');
  const sh = ss.getSheetByName('Envelopes');
  sh.getRange(row, 6).setValue(round2_(Math.max(0, cur.saved + n)));
  if (n > 0 && reason === 'monthly') sh.getRange(row, 8).setValue("'" + ymd_(dayOf_(new Date())).slice(0, 7));
  SpreadsheetApp.flush();
  return envelopes_(ss);
}

// When a planned bill linked to a sinking fund is paid, the fund pays for it
// and its target moves to the next due date.
function useLinkedEnvelope_(ss, p, paidCHF) {
  const env = envelopeRows_(ss).find(e => e.linked === p.description);
  if (!env) return;
  const sh = ss.getSheetByName('Envelopes');
  sh.getRange(env.row, 6).setValue(round2_(Math.max(0, env.saved - paidCHF)));
  if (env.type === 'Sinking fund') {
    const fresh = plannedRows_(ss).find(x => x.row === p.row);
    const next = fresh ? (openOccurrences_(fresh, dayOf_(new Date()), addMonths_(dayOf_(new Date()), 25))[0] || {}).date : null;
    if (next) sh.getRange(env.row, 4).setValue(date_(ymd_(next)));
  }
}

function round6_(n) { return Math.round(n * 1e6) / 1e6; }

// Number formats are set one column at a time: tabs formatted as a Sheets table
// refuse a format change across several columns. A format problem never stops a save.
function format_(range, pattern) {
  for (let c = 0; c < range.getNumColumns(); c++) {
    try { range.offset(0, c, range.getNumRows(), 1).setNumberFormat(pattern); } catch (err) { /* keep the existing format */ }
  }
}
function dateFormat_(range) { format_(range, 'yyyy-mm-dd'); }


// =====================================================================
// iPhone Reminders ("To Pay" list) sync, used by the iPhone shortcut
// =====================================================================
// Each reminder carries an ID in its notes: FIN-<row>-<yyyymmdd>-<code>.
// <code> comes from the description, so a reused row never matches an old reminder.
function reminderCode_(text) {
  let h = 0;
  for (let i = 0; i < text.length; i++) h = (h * 31 + text.charCodeAt(i)) % 1679616; // 36^4
  return ('0000' + h.toString(36)).slice(-4);
}
function reminderId_(p, due) { return 'FIN-' + p.row + '-' + ymd_(due).replace(/-/g, '') + '-' + reminderCode_(p.description); }

function reminderSync_(existing, days) {
  const ss = ss_();
  const today = dayOf_(new Date());
  const until = addDays_(today, Math.min(60, Math.max(1, Number(days) || 7)));
  const ids = (Array.isArray(existing) ? existing : String(existing || '').split(/[\s,;]+/))
    .map(x => String(x).trim()).filter(x => /^FIN-\d+-\d{8}-[0-9a-z]{4}$/.test(x));
  const have = {}; ids.forEach(x => { have[x] = true; });
  const planned = plannedRows_(ss);
  const accBy = {}; accounts_(ss).forEach(a => { accBy[a.name] = a; });
  const nf = n => n.toFixed(2).replace(/\.00$/, '').replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  // Payments and transfers due up to `until` (and unpaid overdue ones). Income is left out.
  const open = {};
  planned.forEach(p => {
    const kind = kindOfPlanned_(p);
    if (kind === 'income') return;
    openOccurrences_(p, today, until).forEach(o => {
      const id = reminderId_(p, o.due);
      let amount = nf(Math.abs(p.amount)) + ' ' + p.currency;
      if (p.fullBalance) {
        const to = accBy[p.toAccount];
        amount = to ? 'full balance, now ' + nf(Math.max(0, -to.balance)) + ' ' + to.currency : 'full balance';
      }
      const title = p.fullBalance ? p.description + ' · full balance (' + amount.replace('full balance, ', '') + ')'
        : kind === 'transfer' ? 'Transfer ' + amount + ': ' + p.account + ' → ' + p.toAccount : p.description + ' · ' + amount;
      const when = o.date < today ? today : o.date;
      open[id] = {
        id, title,
        due: ymd_(when) + ' 09:00',
        notes: (kind === 'transfer' ? p.description + '\n' : 'From ' + p.account + '\n') +
               'Due ' + longDay_(o.date) + (o.date < today ? ' (overdue)' : '') +
               '\nMark it paid in the Finances app.\n' + id
      };
    });
  });

  const create = [], update = [], complete = [], remove = [];
  Object.keys(open).forEach(id => (have[id] ? update : create).push(open[id]));
  ids.forEach(id => {
    if (open[id]) return;
    const m = id.match(/^FIN-(\d+)-(\d{4})(\d{2})(\d{2})-([0-9a-z]{4})$/);
    const p = planned.find(x => x.row === Number(m[1]) && reminderCode_(x.description) === m[5]);
    const due = new Date(Number(m[2]), Number(m[3]) - 1, Number(m[4]));
    const paid = p && ((p.lastPaid && startOfDay_(p.lastPaid) >= due) || (p.repeats === 'Once' && (p.status === 'Paid' || p.status === 'Skipped')));
    if (paid) complete.push(id);
    else if (p && due > until) return;        // still planned, just further away: leave it
    else remove.push(id);                     // deleted, changed or moved: the new date gets its own reminder
  });
  create.sort((a, b) => a.due < b.due ? -1 : 1);
  return { create, update, complete, remove, checked: ymd_(today) + ' ' + Utilities.formatDate(new Date(), TZ, 'HH:mm') };
}


// =====================================================================
// Monthly plan: income − planned payments − day-to-day − set aside − savings
// Transfers between your own accounts are left out: they don't change the total.
// =====================================================================
function savingsTarget_(ss) {
  const sh = ss.getSheetByName('Cash flow');
  if (!sh) return 0;
  const v = sh.getRange(2, 2).getValue();
  return typeof v === 'number' && v > 0 ? v : 0;
}

function saveSavingsTarget_(amount) {
  const n = Number(amount);
  if (!(n >= 0)) throw new Error('Savings target must be 0 or more');
  const ss = ss_();
  const sh = ss.getSheetByName('Cash flow');
  if (!sh) throw new Error('The Cash flow tab is missing');
  sh.getRange(2, 2).setValue(n);
  SpreadsheetApp.flush();
  return plan_();
}

// Planned payments per category and month (CHF), paid or not; skipped/cancelled left out.
function plannedByCategory_(planned, rate, from, to) {
  const out = {};
  planned.forEach(p => {
    if (kindOfPlanned_(p) !== 'expense' || ['Cancelled', 'Stopped', 'Skipped'].indexOf(p.status) >= 0) return;
    occurrences_(p, from, to).forEach(due => {
      const d = shiftWeekend_(due, p.weekend), mk = ymd_(d).slice(0, 7);
      const m = out[mk] || (out[mk] = {});
      m[p.category] = (m[p.category] || 0) + p.amount * rate(p.currency);
      (m['@dates'] || (m['@dates'] = [])).push({ cat: p.category, date: d, chf: p.amount * rate(p.currency) });
    });
  });
  return out;
}

function monthlyPlan_(ss, c) {
  const { today, end, planned, rate, settings, accounts } = c;
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const monthKeys = [];
  for (let i = 0; i < FORECAST_MONTHS; i++) monthKeys.push(ymd_(addMonths_(monthStart, i)).slice(0, 7));
  const rows = monthKeys.map(mk => ({ month: mk, income: 0, planned: 0, budget: 0, pace: 0, setAside: 0, fromFunds: 0, savings: 0 }));
  const idx = {}; monthKeys.forEach((mk, i) => { idx[mk] = i; });
  const billsByMonth = {};

  // Income and planned payments still open (overdue ones count this month)
  planned.forEach(p => {
    const kind = kindOfPlanned_(p);
    if (kind === 'transfer') return;
    openOccurrences_(p, today, end).forEach(o => {
      const mk = ymd_(o.date < today ? today : o.date).slice(0, 7);
      if (idx[mk] === undefined) return;
      const chf = Math.abs(p.amount) * rate(p.currency);
      if (kind === 'income') rows[idx[mk]].income += chf;
      else {
        rows[idx[mk]].planned += chf;
        const b = billsByMonth[mk] || (billsByMonth[mk] = {});
        b[p.description] = (b[p.description] || 0) + chf;
      }
    });
  });

  // ---- Spending outside planned payments, per category ----
  // Budget: each category's budget minus what is planned in it that month (never below 0).
  // Real pace: what you really spend outside planned payments.
  const log = logRows_(ss).filter(e => e.category !== CORRECTION);
  const histStart = new Date(today.getFullYear(), today.getMonth() - 3, 1);
  const plannedCat = plannedByCategory_(planned, rate, histStart, end);
  const spentIn = (mk, cat, until) => log.filter(e => ymd_(e.date).slice(0, 7) === mk && e.category === cat && (!until || e.date <= until)).reduce((t, e) => t + e.chf, 0);
  const plannedUpTo = (mk, cat, until) => ((plannedCat[mk] || {})['@dates'] || []).filter(x => x.cat === cat && x.date <= until).reduce((t, x) => t + x.chf, 0);
  const curMk = monthKeys[0];
  const dim = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const restShare = (dim - today.getDate()) / dim;

  // Full months with records before this one (up to 3) decide the pace; otherwise this month's pace.
  const oldest = log.reduce((m, e) => e.date < m ? e.date : m, today);
  const fullMonths = [];
  for (let i = 1; i <= 3; i++) {
    const m0 = new Date(today.getFullYear(), today.getMonth() - i, 1);
    if (oldest <= m0) fullMonths.push(ymd_(m0).slice(0, 7));   // records cover the whole month
  }
  // This month's pace only counts the days you have been recording.
  const recFrom = oldest > monthStart ? oldest : monthStart;
  const recDays = Math.round((today - recFrom) / 86400000) + 1;
  const paceMethod = fullMonths.length ? 'average' : recDays >= 7 ? 'thisMonth' : 'budget';

  const perCategory = settings.categories.map(cat => {
    const name = cat.name;
    const nextMk = monthKeys[1];
    const budgetNext = Math.max(0, cat.budget - ((plannedCat[nextMk] || {})[name] || 0));
    const actualNow = Math.max(0, spentIn(curMk, name, today) - plannedUpTo(curMk, name, today));
    let pace;
    if (paceMethod === 'average') {
      pace = fullMonths.reduce((t, mk) => t + Math.max(0, spentIn(mk, name) - ((plannedCat[mk] || {})[name] || 0)), 0) / fullMonths.length;
    } else if (paceMethod === 'thisMonth') {
      pace = actualNow / recDays * dim;
    } else {
      pace = Math.max(0, cat.budget - ((plannedCat[curMk] || {})[name] || 0));
    }
    return { category: name, budget: cat.budget, planned: round2_((plannedCat[nextMk] || {})[name] || 0), budgetLeftMonthly: round2_(budgetNext),
      budgetThisMonth: Math.max(0, cat.budget - ((plannedCat[curMk] || {})[name] || 0)), spentThisMonth: round2_(actualNow), pace: round2_(pace) };
  });

  rows.forEach((r, i) => {
    if (i === 0) {
      // Rest of this month: what is still left in each budget / what your pace still adds
      r.budget = perCategory.reduce((t, v) => t + Math.max(0, v.budgetThisMonth - v.spentThisMonth), 0);
      r.pace = perCategory.reduce((t, v) => t + v.pace * restShare, 0);
    } else {
      r.budget = settings.categories.reduce((t, cat) => t + Math.max(0, cat.budget - ((plannedCat[r.month] || {})[cat.name] || 0)), 0);
      r.pace = perCategory.reduce((t, v) => t + v.pace, 0);
    }
  });

  // Funds and goals: monthly amount until the target; a linked bill is paid from its fund
  const E = envelopes_(ss).envelopes;
  E.forEach(e => {
    let saved = e.saved, targetDate = e.targetDate ? date_(e.targetDate) : null;
    rows.forEach((r, i) => {
      const mStart = date_(r.month + '-01');
      if (!(i === 0 && e.setAsideThisMonth) && saved < e.target - 0.005) {
        let amount;
        if (e.monthly !== null && e.monthly !== undefined) amount = e.monthly;
        else if (targetDate) amount = (e.target - saved) / Math.max(1, monthsUntil_(mStart, targetDate) + 1);
        else amount = 0;
        amount = Math.min(amount, e.target - saved);
        r.setAside += amount; saved += amount;
      }
      if (e.type === 'Sinking fund' && e.linked && billsByMonth[r.month] && billsByMonth[r.month][e.linked]) {
        const use = Math.min(saved, billsByMonth[r.month][e.linked]);
        r.fromFunds += use; saved -= use;
        const p = planned.find(x => x.description === e.linked);
        const nxt = p ? occurrences_(p, new Date(mStart.getFullYear(), mStart.getMonth() + 1, 1), addMonths_(mStart, 30))[0] : null;
        targetDate = nxt || targetDate;
      }
    });
  });

  const target = savingsTarget_(ss);
  const netWorth = accounts.reduce((t, a) => t + a.balanceCHF, 0);
  const reserved = E.reduce((t, e) => t + Math.max(0, e.saved), 0);
  const start = round2_(netWorth - reserved);
  let free = start, freePace = start;
  rows.forEach(r => {
    r.savings = target;
    ['income', 'planned', 'budget', 'pace', 'setAside', 'fromFunds', 'savings'].forEach(k => { r[k] = round2_(r[k]); });
    const common = r.income - r.planned - r.setAside + r.fromFunds - r.savings;
    r.left = round2_(common - r.budget);
    r.leftPace = round2_(common - r.pace);
    free = round2_(free + r.left); freePace = round2_(freePace + r.leftPace);
    r.free = free; r.freePace = freePace;
  });
  return {
    start, savingsTarget: target, paceMethod, paceMonths: fullMonths, paceDays: recDays, rows,
    budget: {
      month: monthKeys[1], paceMethod, paceDays: recDays,
      monthlyBudget: round2_(perCategory.reduce((t, v) => t + v.budgetLeftMonthly, 0)),
      monthlyPace: round2_(perCategory.reduce((t, v) => t + v.pace, 0)),
      perCategory: perCategory.map(v => ({ category: v.category, budget: v.budget, planned: v.planned, budgetLeft: v.budgetLeftMonthly, spentThisMonth: v.spentThisMonth, pace: v.pace }))
    }
  };
}
