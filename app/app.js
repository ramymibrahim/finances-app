/* Finances app - talks to the Google Apps Script backend. */
(() => {
  'use strict';

  // ---------- Storage ----------
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
  };
  const cfg = () => store.get('cfg', { url: '', key: '' });

  // ---------- State ----------
  const today = () => toYMD(new Date());
  const state = {
    view: 'home',
    month: today().slice(0, 7),
    boot: store.get('boot', null),          // last bootstrap data (lists)
    summaries: store.get('summaries', {}),  // month -> summary
    type: 'expense',
    currency: 'CHF',
    category: '',
    receipt: null,                          // {data, mimeType, name, preview}
    editing: null,                          // {kind, row, match, receiptUrl, removeReceipt}
    activityItems: [],
    searchItems: [],
    plan: store.get('plan', null),
    fcAccount: 'all',
  };

  // ---------- Helpers ----------
  const $ = (s) => document.querySelector(s);
  const $$ = (s) => Array.from(document.querySelectorAll(s));
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf2 = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const nf0 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
  const money = (n, cur) => (n < 0 ? '−' : '') + nf2.format(Math.abs(n)) + (cur ? ' ' + cur : '');
  const money0 = (n) => (n < 0 ? '−' : '') + nf0.format(Math.abs(n));
  function toYMD(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
  function shiftMonth(ym, n) { const [y, m] = ym.split('-').map(Number); const d = new Date(y, m - 1 + n, 1); return toYMD(d).slice(0, 7); }
  function monthName(ym) { const [y, m] = ym.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('en-GB', { month: 'long', year: 'numeric' }); }
  function dayLabel(ymd) {
    const d = new Date(ymd + 'T12:00:00');
    if (ymd === today()) return 'Today';
    const y = new Date(); y.setDate(y.getDate() - 1);
    if (ymd === toYMD(y)) return 'Yesterday';
    return d.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' });
  }
  const parseSigned = (s) => { const t = String(s).trim().replace(/^[−–]/, '-'); const n = Number(t.replace(/[’',\s]/g, '')); return n; };
  const parseNum = (s) => { const n = Number(String(s).replace(/[’',\s]/g, '').replace(/[^\d.\-]/g, '')); return isFinite(n) ? n : NaN; };

  let toastTimer;
  function toast(msg, ms = 2600) {
    const t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, ms);
  }

  // ---------- API ----------
  class NetworkError extends Error {}
  async function api(action, payload = {}) {
    const { url, key } = cfg();
    if (!url || !key) throw new Error('Open Settings and add the Web app URL and key.');
    let res;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },  // avoids a CORS preflight
        body: JSON.stringify(Object.assign({ key, action }, payload)),
        redirect: 'follow',
      });
    } catch (e) {
      throw new NetworkError('No connection');
    }
    let body;
    try { body = await res.json(); } catch (e) { throw new Error('Unexpected reply from the server (check the URL).'); }
    if (!body.ok) throw new Error(body.error || 'Server error');
    return body.data;
  }

  function saveSummary(s) {
    state.summaries[s.month] = s;
    // keep the six most recent months on the phone
    const keys = Object.keys(state.summaries).sort().reverse();
    keys.slice(6).forEach(k => delete state.summaries[k]);
    store.set('summaries', state.summaries);
  }

  function setSync(text) {
    const el = $('#syncState');
    if (!text) { el.hidden = true; return; }
    el.textContent = text; el.hidden = false;
  }

  async function loadBoot() {
    setSync('Updating…');
    try {
      const data = await api('bootstrap', { month: state.month });
      const { summary, ...lists } = data;
      state.boot = lists; store.set('boot', lists);
      saveSummary(summary);
      setSync('');
      renderAll();
      fillFormLists();
      fillSearchLists();
      if (state.view === 'settings') renderManage();
    } catch (e) {
      setSync(e instanceof NetworkError ? 'Offline' : '');
      if (!(e instanceof NetworkError)) toast(e.message, 4000);
    }
  }

  async function loadMonth(ym) {
    state.month = ym;
    renderAll();
    setSync('Updating…');
    try {
      saveSummary(await api('summary', { month: ym }));
      setSync('');
      renderAll();
    } catch (e) {
      setSync(e instanceof NetworkError ? 'Offline' : '');
      if (!(e instanceof NetworkError)) toast(e.message, 4000);
    }
  }

  // ---------- Offline queue ----------
  const queue = () => store.get('queue', []);
  function setQueue(q) {
    const ok = store.set('queue', q);
    $('#queueInfo').textContent = q.length ? q.length + ' entr' + (q.length === 1 ? 'y' : 'ies') + ' saved on this phone, waiting to upload.' : 'Nothing waiting.';
    return ok;
  }
  let flushing = false;
  async function flushQueue() {
    if (flushing) return;
    let q = queue();
    if (!q.length) return;
    flushing = true;
    setSync('Uploading ' + q.length + '…');
    try {
      while (q.length) {
        const job = q[0];
        try {
          const r = await api(job.action, { entry: job.entry });
          saveSummary(r.summary);
        } catch (e) {
          if (e instanceof NetworkError) break;
          toast('Could not upload an entry: ' + e.message, 5000);
        }
        q = q.slice(1); setQueue(q);
      }
    } finally {
      flushing = false;
      setSync(queue().length ? 'Offline · ' + queue().length + ' waiting' : '');
      renderAll();
    }
  }

  // ---------- Views ----------
  function show(view) {
    state.view = view;
    $$('.view').forEach(v => { v.hidden = v.id !== 'view-' + view; });
    $$('.tabbar button').forEach(b => b.classList.toggle('active', b.dataset.view === view));
    $('#title').textContent = { home: 'This month', activity: 'Activity', add: state.editing ? 'Edit' : 'Add', plan: 'Plan', settings: 'More' }[view];
    if (view === 'home' && state.month !== today().slice(0, 7)) $('#title').textContent = 'Overview';
    window.scrollTo(0, 0);
    if (view === 'add') prepareForm();
    if (view === 'settings') { renderManage(); loadCosts(); }
    if (view === 'plan') { renderPlan(); loadPlan(); }
  }

  function renderAll() {
    const s = state.summaries[state.month];
    $('#monthLabel').textContent = monthName(state.month);
    $$('.monthLabel2').forEach(el => { el.textContent = monthName(state.month); });
    $('#nextMonth').disabled = false;
    if (state.view === 'home') $('#title').textContent = state.month === today().slice(0, 7) ? 'This month' : 'Overview';
    renderHome(s);
    renderActivity(s);
  }

  function renderHome(s) {
    if (!s) {
      $('#heroValue').textContent = '–';
      $('#heroSub').textContent = cfg().url ? 'Loading…' : 'Open Settings to connect your sheet.';
      ['#statSpent', '#statIncome', '#statSalary'].forEach(id => { $(id).textContent = '–'; });
      $('#chart').innerHTML = ''; $('#categories').innerHTML = ''; $('#accounts').innerHTML = '';
      return;
    }
    const left = s.budgetTotal - s.spentTotal;
    const over = left < 0;
    $('#heroLabel').textContent = over ? 'Over budget by' : 'Left to spend';
    $('#heroValue').textContent = money(Math.abs(left), 'CHF');
    $('#heroValue').classList.toggle('over', over);
    let sub = 'of ' + money0(s.budgetTotal) + ' CHF budget';
    if (s.isCurrent && !over) sub = money0(s.dailyAllowance) + ' CHF a day for the next ' + s.daysLeft + ' day' + (s.daysLeft === 1 ? '' : 's');
    $('#heroSub').textContent = sub;
    const pct = s.budgetTotal ? Math.min(100, s.spentTotal / s.budgetTotal * 100) : 0;
    $('#heroMeter').style.width = pct + '%';
    $('#heroMeter').classList.toggle('over', over);

    $('#statSpent').textContent = money0(s.spentTotal);
    $('#statIncome').textContent = money0(s.incomeTotal);
    const sd = new Date(s.nextSalary + 'T12:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
    $('#statSalary').textContent = s.daysToSalary === 0 ? 'Today' : s.daysToSalary + ' d · ' + sd;

    renderChart(s);
    renderCategories(s);
    renderAccounts(s);
  }

  function renderCategories(s) {
    const cats = s.categories.slice().filter(c => c.budget > 0 || c.spent > 0)
      .sort((a, b) => ratio(b) - ratio(a));
    function ratio(c) { return c.budget ? c.spent / c.budget : (c.spent ? 9 : 0); }
    $('#categories').innerHTML = cats.map(c => {
      const r = ratio(c);
      const over = c.spent > c.budget;
      const width = c.budget ? Math.min(100, r * 100) : 100;
      let flag = '';
      if (over) flag = '<div class="cat-flag over">▲ Over by ' + money(c.spent - c.budget) + '</div>';
      else if (r >= 0.8) flag = '<div class="cat-flag near">● ' + Math.round(r * 100) + '% used</div>';
      return '<div class="cat" data-catname="' + esc(c.name) + '"><div class="cat-top"><span class="cat-name">' + esc(c.name) + '</span>' +
        '<span class="cat-nums">' + money0(c.spent) + ' / ' + money0(c.budget) + '</span></div>' +
        '<div class="bar"><span class="' + (over ? 'over' : '') + '" style="width:' + width + '%"></span></div>' + flag + '</div>';
    }).join('') || '<div class="muted small">No budget set.</div>';
  }

  function renderAccounts(s) {
    $('#netWorth').textContent = 'Total ' + money(s.netWorthCHF, 'CHF');
    if (s.reservedCHF > 0) {
      $('#freeToSpend').innerHTML = 'Set aside in funds and goals <b>' + money(s.reservedCHF, 'CHF') + '</b><br>Free after that (accounts without credit cards) <b>' + money(s.spendableCHF - s.reservedCHF, 'CHF') + '</b>';
      $('#freeToSpend').hidden = false;
    } else $('#freeToSpend').hidden = true;
    $('#accounts').innerHTML = s.accounts.map(a =>
      '<div class="acc"><div><div>' + esc(a.name) + '</div><div class="acc-type">' + esc(a.type) + '</div></div>' +
      '<div class="acc-bal"><div class="' + (a.balance < 0 ? 'neg' : '') + '">' + money(a.balance, a.currency) + '</div>' +
      (a.currency !== 'CHF' ? '<div class="acc-type">' + money(a.balanceCHF, 'CHF') + '</div>' : '') + '</div></div>'
    ).join('');
  }

  // Cumulative spending line with a dashed "budget pace" reference line.
  function renderChart(s) {
    const el = $('#chart');
    const W = el.clientWidth || 340, H = el.clientHeight || 190;
    const m = { t: 12, r: 12, b: 22, l: 40 };
    const n = s.daysInMonth;
    const lastDay = s.isCurrent ? Number(s.today.slice(8, 10)) : n;
    const cum = []; let run = 0;
    for (let i = 0; i < lastDay; i++) { run += s.daily[i] || 0; cum.push(run); }
    const maxY = niceMax(Math.max(s.budgetTotal, run, 1));
    const x = d => m.l + (d - 1) / Math.max(1, n - 1) * (W - m.l - m.r);
    const y = v => H - m.b - v / maxY * (H - m.t - m.b);

    const ticks = [0, maxY / 2, maxY];
    let svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" aria-hidden="true"><g class="axis">';
    ticks.forEach(t => {
      svg += '<line class="' + (t === 0 ? 'baseline' : 'gridline') + '" x1="' + m.l + '" x2="' + (W - m.r) + '" y1="' + y(t) + '" y2="' + y(t) + '"/>';
      svg += '<text x="' + (m.l - 6) + '" y="' + (y(t) + 4) + '" text-anchor="end">' + money0(t) + '</text>';
    });
    [1, 10, 20, n].forEach(d => { svg += '<text x="' + x(d) + '" y="' + (H - 4) + '" text-anchor="middle">' + d + '</text>'; });
    svg += '</g>';
    // budget pace
    svg += '<line class="pace" x1="' + x(1) + '" y1="' + y(0) + '" x2="' + x(n) + '" y2="' + y(s.budgetTotal) + '"/>';
    svg += '<text class="label" x="' + (x(n) - 2) + '" y="' + (y(s.budgetTotal) - 6) + '" text-anchor="end">Budget pace</text>';
    if (cum.length) {
      const pts = cum.map((v, i) => x(i + 1).toFixed(1) + ',' + y(v).toFixed(1));
      svg += '<path class="spent-area" d="M' + x(1) + ',' + y(0) + ' L' + pts.join(' L') + ' L' + x(cum.length) + ',' + y(0) + 'Z"/>';
      svg += '<polyline class="spent" points="' + pts.join(' ') + '"/>';
      const lx = x(cum.length), ly = y(run);
      svg += '<circle class="dot" cx="' + lx + '" cy="' + ly + '" r="4"/>';
      const anchor = cum.length > n * 0.7 ? 'end' : 'start';
      svg += '<text class="label" x="' + (lx + (anchor === 'end' ? -8 : 8)) + '" y="' + (ly - 8) + '" text-anchor="' + anchor + '">Spent ' + money0(run) + '</text>';
    }
    svg += '<line class="cross" id="cross" y1="' + m.t + '" y2="' + (H - m.b) + '" visibility="hidden"/>';
    svg += '<circle class="dot" id="crossDot" r="4" visibility="hidden"/></svg>';
    el.innerHTML = svg;

    // table view
    let rows = '';
    for (let d = 1; d <= cum.length; d++) {
      if (!s.daily[d - 1]) continue;
      rows += '<tr><td>' + d + '</td><td>' + money(s.daily[d - 1]) + '</td><td>' + money(cum[d - 1]) + '</td><td>' + money(s.budgetTotal * (d - 1) / Math.max(1, n - 1)) + '</td></tr>';
    }
    $('#chartTable').innerHTML = rows ? '<table><tr><th>Day</th><th>Spent</th><th>Total</th><th>Budget pace</th></tr>' + rows + '</table>' : '<p>No spending yet.</p>';

    // hover / touch
    const tip = $('#tooltip');
    const move = (clientX) => {
      const r = el.getBoundingClientRect();
      const px = (clientX - r.left) / r.width * W;
      const d = Math.max(1, Math.min(n, Math.round((px - m.l) / (W - m.l - m.r) * (n - 1) + 1)));
      const cross = el.querySelector('#cross'), dot = el.querySelector('#crossDot');
      cross.setAttribute('x1', x(d)); cross.setAttribute('x2', x(d)); cross.setAttribute('visibility', 'visible');
      const pace = s.budgetTotal * (d - 1) / Math.max(1, n - 1);
      let html = '<b>' + d + ' ' + monthName(s.month).split(' ')[0] + '</b>';
      if (d <= cum.length) {
        dot.setAttribute('cx', x(d)); dot.setAttribute('cy', y(cum[d - 1])); dot.setAttribute('visibility', 'visible');
        html += '<br>Spent that day ' + money(s.daily[d - 1]) + '<br>Total so far ' + money(cum[d - 1]);
      } else dot.setAttribute('visibility', 'hidden');
      html += '<br><span style="color:var(--muted)">Budget pace ' + money(pace) + '</span>';
      tip.innerHTML = html; tip.hidden = false;
      const tx = Math.min(window.innerWidth - tip.offsetWidth - 8, Math.max(8, r.left + x(d) / W * r.width - tip.offsetWidth / 2));
      tip.style.left = tx + 'px'; tip.style.top = (r.top - tip.offsetHeight - 6) + 'px';
    };
    const hide = () => { tip.hidden = true; const c = el.querySelector('#cross'), dd = el.querySelector('#crossDot'); if (c) c.setAttribute('visibility', 'hidden'); if (dd) dd.setAttribute('visibility', 'hidden'); };
    el.onpointermove = (e) => move(e.clientX);
    el.onpointerdown = (e) => move(e.clientX);
    el.onpointerleave = hide;
    el.onpointerup = (e) => { if (e.pointerType !== 'mouse') setTimeout(hide, 1500); };
  }
  function niceMax(v) {
    const p = Math.pow(10, Math.floor(Math.log10(v)));
    for (const k of [1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) if (k * p >= v) return k * p;
    return 10 * p;
  }

  function renderActivity(s) {
    const box = $('#activity');
    const pending = queue().filter(j => j.entry.date.slice(0, 7) === state.month).map(j => ({
      pending: true, kind: { addExpense: 'expense', addIncome: 'income', addTransfer: 'transfer' }[j.action],
      date: j.entry.date, merchant: j.entry.merchant, item: j.entry.item, source: j.entry.source,
      from: j.entry.from, to: j.entry.to, amount: Number(j.entry.amount),
      currency: j.entry.currency || accountCurrency(j.entry.from), category: j.entry.category, type: j.entry.type, chf: 0 }));
    const items = pending.concat(s ? (s.entries || []) : []);
    state.activityItems = items;
    if (!items.length) { box.innerHTML = '<div class="empty">' + (s ? 'Nothing recorded this month.' : 'Nothing to show yet.') + '</div>'; return; }
    const byDay = {};
    items.forEach((e, i) => { (byDay[e.date] = byDay[e.date] || []).push(i); });
    box.innerHTML = Object.keys(byDay).sort().reverse().map(d => {
      const spent = byDay[d].reduce((t, i) => t + (items[i].kind === 'expense' ? (items[i].chf || 0) : 0), 0);
      return '<div class="day-head"><span>' + dayLabel(d) + '</span><span>' + (spent ? money(spent, 'CHF') + ' spent' : '') + '</span></div>' +
        byDay[d].map(i => entryHtml(items[i], i)).join('');
    }).join('');
  }
  function entryHtml(e, i, attr) {
    let title, sub, ic, icClass = '', amt, amtClass = '';
    if (e.kind === 'income') {
      title = e.source || e.type || 'Income'; sub = (e.type || 'Income') + (e.account ? ' · ' + e.account : '');
      ic = '+'; icClass = 'in'; amt = '+' + money(e.amount, e.currency); amtClass = 'in';
    } else if (e.kind === 'transfer') {
      title = e.from + ' → ' + e.to; sub = 'Transfer' + (e.costPct != null ? ' · cost ' + e.costPct.toFixed(1) + '% (' + money(e.costCHF) + ' CHF)' : '') + (e.note ? ' · ' + e.note : '');
      ic = '⇄'; icClass = 'tr'; amt = money(e.amount, e.currency);
    } else {
      title = e.merchant || e.item || e.category; sub = e.category + (e.item && e.merchant && e.item !== e.category ? ' · ' + e.item : '');
      ic = (e.category || '?').slice(0, 1); amt = money(e.amount, e.currency);
    }
    if (e.pending) sub = '⏳ Waiting to upload';
    const chf = e.kind !== 'transfer' && e.currency && e.currency !== 'CHF' && e.chf ? '<span class="small">' + money(e.chf, 'CHF') + '</span>' : '';
    return '<div class="entry" data-' + (attr || 'i') + '="' + i + '"><div class="entry-ic ' + icClass + '">' + esc(ic) + '</div>' +
      '<div class="entry-main"><div class="entry-title">' + esc(title) + (e.receipt ? ' 📎' : '') + '</div>' +
      '<div class="entry-sub">' + esc(sub) + '</div></div>' +
      '<div class="entry-amt ' + amtClass + '">' + esc(amt) + chf + '</div></div>';
  }

  // ---------- Add form ----------
  function fillFormLists() {
    const b = state.boot; if (!b) return;
    $('#merchantList').innerHTML = Object.keys(b.merchants || {}).sort().map(n => '<option value="' + esc(n) + '">').join('');
    $('#itemList').innerHTML = (b.items || []).map(i => '<option value="' + esc(i.name) + '">').join('');
    const opts = b.accounts.map(a => '<option value="' + esc(a.name) + '">' + esc(a.name) + ' (' + esc(a.currency) + ')</option>').join('');
    ['#fAccount', '#fIncomeAccount', '#fFrom', '#fTo'].forEach(id => { const v = $(id).value; $(id).innerHTML = opts; if (v) $(id).value = v; });
    $('#fCurrency').innerHTML = b.currencies.map(c => '<button type="button" class="chip" role="radio" data-cur="' + esc(c.code) + '">' + esc(c.code) + '</button>').join('');
    $('#fCategory').innerHTML = b.categories.map(c => '<button type="button" class="chip" role="radio" data-cat="' + esc(c) + '">' + esc(c) + '</button>').join('');
    setCurrency(state.currency); setCategory(state.category);
  }

  function setCurrency(c) {
    state.currency = c;
    $$('#fCurrency .chip').forEach(b => { const on = b.dataset.cur === c; b.classList.toggle('active', on); b.setAttribute('aria-checked', on); });
  }
  function setCategory(c) {
    state.category = c;
    $$('#fCategory .chip').forEach(b => { const on = b.dataset.cat === c; b.classList.toggle('active', on); b.setAttribute('aria-checked', on); });
  }
  function setType(t) {
    state.type = t;
    $$('.segmented button').forEach(b => b.classList.toggle('active', b.dataset.type === t));
    $$('.only-expense').forEach(el => { el.hidden = t !== 'expense'; });
    $$('.only-income').forEach(el => { el.hidden = t !== 'income'; });
    $$('.only-transfer').forEach(el => { el.hidden = t !== 'transfer'; });
    $$('.only-notransfer').forEach(el => { el.hidden = t === 'transfer'; });
    $('#fCurrency').hidden = t === 'transfer';
    $('#formError').hidden = true;
    if (t === 'transfer') updateTransferCurrencies();
  }
  function accountCurrency(name) { const a = (state.boot && state.boot.accounts || []).find(x => x.name === name); return a ? a.currency : 'CHF'; }
  function updateTransferCurrencies() {
    const fc = accountCurrency($('#fFrom').value), tc = accountCurrency($('#fTo').value);
    $('#receivedWrap').hidden = fc === tc;
    $('#receivedCur').textContent = '(' + tc + ')';
    $('#fAmount').placeholder = '0.00 ' + fc;
  }

  function prepareForm() {
    if (!$('#fDate').value) $('#fDate').value = today();
    if (state.editing) return;
    if (state.boot && state.boot.lastAccount && !$('#fAccount').dataset.touched) $('#fAccount').value = state.boot.lastAccount;
    setType(state.type);
    setTimeout(() => $('#fAmount').focus(), 50);
  }

  function resetForm() {
    ['#fAmount', '#fMerchant', '#fItem', '#fNote', '#fRate', '#fSource', '#fReceived'].forEach(id => { $(id).value = ''; });
    $('#fDate').value = today();
    setCategory('');
    setCurrency('CHF');
    clearReceipt();
    state.editing = null;
    $('#editBar').hidden = true; $('#deleteBtn').hidden = true; $('#duplicateBtn').hidden = true;
    $('.segmented').hidden = false;
    $('#receiptExisting').hidden = true;
    $('#saveBtn').textContent = 'Save';
  }

  // Open an existing entry in the form.
  function openEdit(e) {
    resetForm();
    state.editing = { kind: e.kind, row: e.row, match: { date: e.date, amount: e.amount }, receiptUrl: e.receipt || '', removeReceipt: false };
    setType(e.kind);
    $('.segmented').hidden = true;
    $('#editBar').hidden = false;
    $('#editLabel').textContent = { expense: 'Expense', income: 'Income', transfer: 'Transfer' }[e.kind];
    $('#deleteBtn').hidden = false; $('#duplicateBtn').hidden = false;
    $('#saveBtn').textContent = 'Save changes';
    $('#fAmount').value = Number(e.amount).toFixed(2);
    $('#fDate').value = e.date;
    $('#fNote').value = e.note || '';
    $('#fRate').value = e.manualRate === '' || e.manualRate == null ? '' : e.manualRate;
    if (e.kind === 'expense') {
      setCurrency(e.currency); setCategory(e.category);
      $('#fMerchant').value = e.merchant || ''; $('#fItem').value = e.item || '';
      $('#fAccount').value = e.account;
      if (e.receipt) { $('#receiptLink').href = e.receipt; $('#receiptExisting').hidden = false; $('#receiptBtn').textContent = '📷 Replace receipt'; }
    } else if (e.kind === 'income') {
      setCurrency(e.currency);
      $('#fSource').value = e.source || '';
      const sel = $('#fIncomeType');
      if (e.type && !Array.from(sel.options).some(o => o.value === e.type)) sel.add(new Option(e.type, e.type));
      sel.value = e.type || 'Salary';
      $('#fIncomeAccount').value = e.account;
    } else {
      $('#fFrom').value = e.from; $('#fTo').value = e.to;
      updateTransferCurrencies();
      $('#fReceived').value = e.received === '' || e.received == null ? '' : e.received;
    }
    show('add');
  }

  // Merchant memory: fill in what you used last time at this merchant.
  function onMerchant() {
    const m = state.boot && state.boot.merchants && state.boot.merchants[$('#fMerchant').value.trim()];
    if (!m) return;
    if (m.category) setCategory(m.category);
    if (m.currency) setCurrency(m.currency);
    if (m.account) $('#fAccount').value = m.account;
    if (m.item && !$('#fItem').value) $('#fItem').value = m.item;
  }
  function onItem() {
    const it = (state.boot && state.boot.items || []).find(i => i.name === $('#fItem').value.trim());
    if (it && it.category) setCategory(it.category);
  }

  // ---------- Receipt ----------
  async function onReceiptChosen(file) {
    if (!file) return;
    try {
      if (file.type === 'application/pdf') {
        if (file.size > 8 * 1024 * 1024) throw new Error('PDF is larger than 8 MB.');
        const data = await readBase64(file);
        state.receipt = { data, mimeType: 'application/pdf', name: file.name, preview: '' };
      } else {
        const { data, preview } = await compressImage(file, 1600, 0.72);
        state.receipt = { data, mimeType: 'image/jpeg', name: file.name || 'Photo', preview };
      }
      $('#receiptImg').src = state.receipt.preview || 'icons/pdf.svg';
      $('#receiptName').textContent = state.receipt.name + ' · ' + Math.round(state.receipt.data.length * 0.75 / 1024) + ' KB';
      $('#receiptPreview').hidden = false;
      $('#receiptBtn').textContent = '📷 Change receipt';
    } catch (e) {
      toast('Could not read that file: ' + e.message, 4000);
    }
  }
  function clearReceipt() {
    state.receipt = null; $('#fReceipt').value = '';
    $('#receiptPreview').hidden = true; $('#receiptBtn').textContent = '📷 Add receipt';
  }
  function readBase64(file) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(String(r.result).split(',')[1]);
      r.onerror = () => rej(r.error);
      r.readAsDataURL(file);
    });
  }
  function compressImage(file, maxSide, quality) {
    return new Promise((res, rej) => {
      const url = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxSide / Math.max(img.naturalWidth, img.naturalHeight));
        const c = document.createElement('canvas');
        c.width = Math.round(img.naturalWidth * scale); c.height = Math.round(img.naturalHeight * scale);
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
        ctx.drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        const dataUrl = c.toDataURL('image/jpeg', quality);
        res({ data: dataUrl.split(',')[1], preview: dataUrl });
      };
      img.onerror = () => { URL.revokeObjectURL(url); rej(new Error('Unsupported image')); };
      img.src = url;
    });
  }

  // ---------- Save ----------
  function buildEntry() {
    const amount = parseNum($('#fAmount').value);
    const date = $('#fDate').value;
    const note = $('#fNote').value.trim();
    const rate = $('#fRate').value.trim() ? parseNum($('#fRate').value) : '';
    if (!(amount > 0)) throw new Error('Enter an amount.');
    if (!date) throw new Error('Choose a date.');
    if (rate !== '' && !(rate > 0)) throw new Error('Bank rate must be a number.');
    if (state.type === 'expense') {
      if (!state.category) throw new Error('Choose a category.');
      return { action: 'addExpense', label: 'Expense', entry: {
        date, amount, currency: state.currency, category: state.category,
        merchant: $('#fMerchant').value.trim(), item: $('#fItem').value.trim(),
        account: $('#fAccount').value, note, manualRate: rate,
        receipt: state.receipt ? { data: state.receipt.data, mimeType: state.receipt.mimeType } : null,
        removeReceipt: !!(state.editing && state.editing.removeReceipt && !state.receipt),
      } };
    }
    if (state.type === 'income') {
      return { action: 'addIncome', label: 'Income', entry: {
        date, amount, currency: state.currency, source: $('#fSource').value.trim(),
        type: $('#fIncomeType').value, account: $('#fIncomeAccount').value, note, manualRate: rate,
      } };
    }
    const from = $('#fFrom').value, to = $('#fTo').value;
    if (from === to) throw new Error('Choose two different accounts.');
    const received = $('#receivedWrap').hidden ? '' : parseNum($('#fReceived').value);
    if (!$('#receivedWrap').hidden && $('#fReceived').value.trim() && !(received > 0)) throw new Error('Amount received must be a number.');
    return { action: 'addTransfer', label: 'Transfer', entry: { date, amount, from, to, received: received > 0 ? received : '', note } };
  }

  async function onSave(ev) {
    ev.preventDefault();
    const err = $('#formError'); err.hidden = true;
    let job;
    try { job = buildEntry(); } catch (e) { err.textContent = e.message; err.hidden = false; return; }
    const btn = $('#saveBtn'); const label = btn.textContent; btn.disabled = true; btn.textContent = 'Saving…';
    if (state.editing) {
      const ed = state.editing;
      try {
        const r = await api('updateEntry', { kind: ed.kind, row: ed.row, match: ed.match, entry: job.entry });
        saveSummary(r.summary);
        toast('Changes saved ✓');
        afterSave(job, 'activity');
      } catch (e) {
        err.textContent = e instanceof NetworkError ? 'Editing needs a connection. Try again when you are online.' : e.message;
        err.hidden = false;
      } finally { btn.disabled = false; btn.textContent = label; }
      return;
    }
    try {
      const r = await api(job.action, { entry: job.entry });
      saveSummary(r.summary);
      toast(job.label + ' saved ✓');
      state.planStale = true;
      afterSave(job);
    } catch (e) {
      if (e instanceof NetworkError) {
        const q = queue(); q.push(job);
        if (setQueue(q)) { toast('No connection. Saved on this phone, will upload later.', 3500); afterSave(job); }
        else { err.textContent = 'No connection, and the phone storage is full (receipt too large). Try again when online.'; err.hidden = false; }
      } else {
        err.textContent = e.message; err.hidden = false;
      }
    } finally {
      btn.disabled = false; btn.textContent = 'Save';
    }
  }
  function afterSave(job, view) {
    const ym = job.entry.date.slice(0, 7);
    resetForm();
    state.month = ym;
    renderAll();
    show(view || 'home');
  }
  async function onDelete() {
    const ed = state.editing; if (!ed) return;
    if (!window.confirm('Delete this entry? This can\'t be undone in the app.')) return;
    const btn = $('#deleteBtn'); btn.disabled = true;
    try {
      const r = await api('deleteEntry', { kind: ed.kind, row: ed.row, match: ed.match });
      saveSummary(r.summary);
      toast('Entry deleted');
      const ym = ed.match.date.slice(0, 7);
      resetForm(); state.month = ym; renderAll(); show('activity');
    } catch (e) {
      $('#formError').textContent = e instanceof NetworkError ? 'Deleting needs a connection.' : e.message;
      $('#formError').hidden = false;
    } finally { btn.disabled = false; }
  }

  // ---------- More: accounts and categories ----------
  function renderManage() {
    const b = state.boot;
    const s = state.summaries[today().slice(0, 7)] || state.summaries[state.month];
    if (!b) { $('#accList').innerHTML = '<div class="muted small">Connect first.</div>'; $('#catList').innerHTML = ''; return; }
    const bal = {}; (s ? s.accounts : []).forEach(a => { bal[a.name] = a; });
    $('#accList').innerHTML = b.accounts.map((a, i) =>
      '<div class="li" data-acc="' + i + '"><div><div>' + esc(a.name) + '</div><div class="li-sub">' + esc(a.type) + ' · ' + esc(a.currency) + (b.lastReconciled && b.lastReconciled[a.name] ? ' · checked ' + dateShort(b.lastReconciled[a.name]).slice(0, -5) : '') + '</div></div>' +
      '<div class="li-right">' + (bal[a.name] ? money(bal[a.name].balance, a.currency) : '') + '</div></div>').join('') || '<div class="muted small">No accounts yet.</div>';
    const budgets = b.budgets || [];
    $('#catList').innerHTML = budgets.map((c, i) =>
      '<div class="li" data-cat="' + i + '"><div>' + esc(c.name) + '</div><div class="li-right">' + money0(c.budget) + ' CHF</div></div>').join('');
    $('#catTotal').textContent = money0(budgets.reduce((t, c) => t + c.budget, 0)) + ' CHF';
  }

  let sheetMode = null;
  function field(label, html) { return '<label>' + label + html + '</label>'; }
  function opts(list, sel) { return list.map(v => '<option' + (v === sel ? ' selected' : '') + '>' + esc(v) + '</option>').join(''); }
  function openSheet(mode, item) {
    sheetMode = { mode, item };
    const b = state.boot;
    $('#sheetError').hidden = true;
    $('#sheetDelete').hidden = !item;
    $('#sheetSave').hidden = false; $('#sheetCancel').textContent = 'Cancel';
    if (mode === 'recurring') return openRecurringSheet(item);
    if (mode === 'paid') return openPaidSheet(item);
    if (mode === 'catdetail') return openCategoryDetail(item);
    if (mode === 'envelope') return openEnvelopeSheet(item);
    let html = '';
    if (mode === 'account') {
      $('#sheetTitle').textContent = item ? 'Edit account' : 'New account';
      html += field('Name', '<input id="mName" value="' + esc(item ? item.name : '') + '" placeholder="e.g. PostFinance CHF">');
      html += field('Currency', '<select id="mCur">' + b.currencies.map(c => '<option' + (item && item.currency === c.code ? ' selected' : '') + '>' + esc(c.code) + '</option>').join('') + '</select>');
      html += field('Type', '<select id="mType">' + (b.accountTypes || ['Current', 'Credit card', 'Savings', 'Cash']).map(t => '<option' + (item && item.type === t ? ' selected' : '') + '>' + esc(t) + '</option>').join('') + '</select>');
      html += field('Starting balance <span class="muted">(use minus for debt, e.g. -1200)</span>', '<input id="mStart" inputmode="decimal" value="' + (item ? item.start : '') + '" placeholder="0.00">');
      html += field('Safety cushion <span class="muted">(warn me if the balance goes below)</span>', '<input id="mCushion" inputmode="decimal" value="' + (item && item.cushion != null ? item.cushion : '') + '" placeholder="Empty = 0; no warning for credit cards">');
      if (item) html += '<p class="muted small">Renaming updates every entry that uses this account.</p>';
      if (item) html += '<div class="recon"><h3>Check against your bank</h3><p class="muted small">Type the balance your bank shows right now. Debt on a card is a minus number.</p>' +
        '<label>Bank balance (' + esc(item.currency) + ')<input id="rcBank" inputmode="decimal" placeholder="e.g. 2540.15"></label>' +
        '<button type="button" class="btn-secondary" id="rcCheck">Check</button><div id="rcResult"></div></div>';
    } else {
      $('#sheetTitle').textContent = item ? 'Edit category' : 'New category';
      html += field('Name', '<input id="mName" value="' + esc(item ? item.name : '') + '" placeholder="e.g. Gifts">');
      html += field('Monthly budget (CHF)', '<input id="mBudget" inputmode="decimal" value="' + (item ? item.budget : '') + '" placeholder="0">');
      if (item) html += '<p class="muted small">Renaming updates every entry that uses this category.</p>';
    }
    $('#sheetBody').innerHTML = html;
    if ($('#rcCheck')) $('#rcCheck').onclick = () => reconcileCheck(item);
    $('#sheet').showModal();
  }
  async function reconcileCheck(item) {
    const out = $('#rcResult');
    const bank = parseSigned($('#rcBank').value);
    if ($('#rcBank').value.trim() === '' || !isFinite(bank)) { out.innerHTML = '<p class="form-error">Enter a number.</p>'; return; }
    out.innerHTML = '<p class="muted small">Checking…</p>';
    try {
      const r = await api('reconcileCheck', { account: item.name, bankBalance: bank });
      const d = r.difference, cur = r.currency;
      let html = '<div class="kv"><span>The app shows</span><b class="num">' + money(r.appBalance, cur) + '</b></div>' +
        '<div class="kv"><span>Your bank shows</span><b class="num">' + money(r.bankBalance, cur) + '</b></div>' +
        '<div class="kv"><span>Difference</span><b class="num ' + (d < 0 ? 'neg' : d > 0 ? 'pos' : '') + '">' + (d > 0 ? '+' : '') + money(d, cur) + '</b></div>';
      if (Math.abs(d) < 0.005) html += '<p class="small" style="color:var(--good)">✓ They match.</p><button type="button" class="btn-secondary" id="rcApply">Mark as checked today</button>';
      else html += '<p class="muted small">' + (d < 0 ? 'Money left the account that the app doesn\'t know about. Add the missing entries, or add a correction.' : 'More money arrived than the app knows about. Add the missing income, or add a correction.') +
        ' Corrections don\'t count as spending or income in your budget.</p><button type="button" class="btn-primary" id="rcApply">Add correction of ' + (d > 0 ? '+' : '') + money(d, cur) + '</button>';
      out.innerHTML = html;
      $('#rcApply').onclick = async () => {
        $('#rcApply').disabled = true;
        try {
          await api('reconcile', { account: item.name, bankBalance: bank });
          $('#sheet').close(); toast(Math.abs(d) < 0.005 ? 'Checked ✓' : 'Correction added ✓'); loadBoot(); state.planStale = true;
        } catch (e) { out.innerHTML += '<p class="form-error">' + esc(e instanceof NetworkError ? 'This needs a connection.' : e.message) + '</p>'; }
      };
    } catch (e) { out.innerHTML = '<p class="form-error">' + esc(e instanceof NetworkError ? 'This needs a connection.' : e.message) + '</p>'; }
  }
  async function onSheetSave(ev) {
    ev.preventDefault();
    const { mode, item } = sheetMode;
    if (mode === 'recurring') return saveRecurringSheet();
    if (mode === 'paid') return savePaidSheet();
    if (mode === 'envelope') return saveEnvelopeSheet();
    if (mode === 'catdetail') { $('#sheet').close(); return; }
    const errEl = $('#sheetError'); errEl.hidden = true;
    const btn = $('#sheetSave'); btn.disabled = true;
    try {
      if (mode === 'account') {
        const start = $('#mStart').value.trim() === '' ? 0 : parseSigned($('#mStart').value);
        if (!isFinite(start)) throw new Error('Starting balance must be a number.');
        const cushionTxt = $('#mCushion').value.trim();
        const cushion = cushionTxt === '' ? '' : parseSigned(cushionTxt);
        if (cushion !== '' && !isFinite(cushion)) throw new Error('Safety cushion must be a number.');
        const r = await api('saveAccount', { account: { original: item ? item.name : '', name: $('#mName').value.trim(), currency: $('#mCur').value, type: $('#mType').value, start, cushion } });
        state.boot.accounts = r.accounts;
      } else {
        const budget = $('#mBudget').value.trim() === '' ? 0 : parseNum($('#mBudget').value);
        if (!(budget >= 0)) throw new Error('Budget must be a number.');
        const r = await api('saveCategory', { category: { original: item ? item.name : '', name: $('#mName').value.trim(), budget } });
        state.boot.budgets = r.budgets; state.boot.categories = r.budgets.map(c => c.name);
      }
      store.set('boot', state.boot);
      $('#sheet').close();
      toast('Saved ✓');
      renderManage(); fillFormLists();
      loadBoot();
    } catch (e) {
      errEl.textContent = e instanceof NetworkError ? 'This needs a connection.' : e.message; errEl.hidden = false;
    } finally { btn.disabled = false; }
  }
  async function onSheetDelete() {
    const { mode, item } = sheetMode;
    if (mode === 'recurring') return deleteRecurringSheet();
    if (mode === 'envelope') return deleteEnvelopeSheet();
    if (!window.confirm('Delete "' + item.name + '"?')) return;
    try {
      if (mode === 'account') { const r = await api('deleteAccount', { name: item.name }); state.boot.accounts = r.accounts; }
      else { const r = await api('deleteCategory', { name: item.name }); state.boot.budgets = r.budgets; state.boot.categories = r.budgets.map(c => c.name); }
      store.set('boot', state.boot);
      $('#sheet').close(); toast('Deleted'); renderManage(); fillFormLists(); loadBoot();
    } catch (e) {
      $('#sheetError').textContent = e instanceof NetworkError ? 'This needs a connection.' : e.message; $('#sheetError').hidden = false;
    }
  }


  // ---------- Plan: forecast, gaps, coming up, recurring ----------
  async function loadPlan() {
    if (!cfg().url) return;
    setSync('Updating forecast…');
    try {
      state.plan = await api('plan');
      store.set('plan', state.plan);
      setSync('');
      if (state.view === 'plan') renderPlan();
    } catch (e) {
      setSync(e instanceof NetworkError ? 'Offline' : '');
      if (!(e instanceof NetworkError)) toast(e.message, 4000);
    }
  }
  const monthShort = (ym) => { const [y, m] = ym.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('en-GB', { month: 'short' }) + ' ' + String(y).slice(2); };
  const dateShort = (ymd) => new Date(ymd + 'T12:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });

  function renderPlan() {
    const p = state.plan;
    if (!p) { $('#gapCard').innerHTML = '<div class="muted">' + (cfg().url ? 'Loading forecast…' : 'Connect first in More.') + '</div>'; return; }
    // Gaps
    $('#gapCard').innerHTML = '<div class="card-head"><h2>Warnings for the next 24 months</h2></div>' + (p.gaps.length ? p.gaps.map(g =>
      '<div class="gap"><div class="gap-ic bad" aria-hidden="true">!</div><div><div class="gap-title">' + esc(g.account) + ' goes below ' + (g.cushion ? 'its cushion of ' + money0(g.cushion) + ' ' + esc(g.currency) : 'zero') + ' on ' + dateShort(g.firstDate) + '</div>' +
      '<div class="gap-sub">Lowest point ' + money0(g.lowest) + ' ' + esc(g.currency) + ' on ' + dateShort(g.lowestDate) + ', about ' + money0(g.shortCHF) + ' CHF short.' +
      (g.suggestFrom ? ' You could plan a transfer from ' + esc(g.suggestFrom) + '.' : ' Plan a transfer into this account, or reduce what is paid from it.') + '</div></div></div>').join('')
      : '<div class="gap"><div class="gap-ic okay" aria-hidden="true">✓</div><div><div class="gap-title">No gaps found</div><div class="gap-sub">Every account stays above its safety cushion.</div></div></div>');

    // Account chips
    const chips = [{ k: 'all', label: 'All (CHF)' }].concat(p.accounts.map(a => ({ k: a.name, label: a.name })));
    if (!chips.some(c => c.k === state.fcAccount)) state.fcAccount = 'all';
    $('#fcAccounts').innerHTML = chips.map(c => '<button type="button" class="chip' + (c.k === state.fcAccount ? ' active' : '') + '" data-fc="' + esc(c.k) + '">' + esc(c.label) + '</button>').join('');
    renderForecastChart(p);
    renderCashflowChart(p);
    $('#varNote').textContent = 'Money out includes planned items and expected day-to-day spending of about ' + money0(p.variable.monthlyCHF) + ' CHF a month, ' +
      (p.variable.method === 'budget' ? 'taken from your budgets (the app switches to your real 3-month average once it has 3 full months).' : 'based on your last 3 months.');

    // Coming up
    $('#upcoming').innerHTML = p.upcoming.length ? p.upcoming.map((u, i) => {
      const d = new Date(u.date + 'T12:00:00');
      const sub = (u.kind === 'transfer' ? u.account + ' → ' + u.toAccount : u.account) + (u.overdue ? ' · overdue' : u.date !== u.due ? ' · moved from ' + dateShort(u.due).slice(0, -5) : '');
      return '<div class="up"><div class="up-date' + (u.overdue ? ' overdue' : '') + '"><b>' + d.getDate() + '</b>' + d.toLocaleDateString('en-GB', { month: 'short' }) + '</div>' +
        '<div class="up-main"><div class="up-title">' + esc(u.description) + '</div><div class="up-sub">' + esc(sub) + '</div></div>' +
        '<div><div class="up-amt' + (u.kind === 'income' ? ' in' : '') + '">' + (u.kind === 'income' ? '+' : '') + money(u.amount, u.currency) + '</div>' +
        '<div class="up-btns"><button class="mini" data-skip="' + i + '">Skip</button><button class="mini primary" data-paid="' + i + '">' + (u.kind === 'income' ? 'Received' : 'Paid') + '</button></div></div></div>';
    }).join('') : '<div class="muted small">Nothing due in the next 45 days.</div>';

    renderEnvelopes(p.envelopes);

    // Recurring list
    const rec = p.recurring.slice().sort((a, b) => (a.next || '9') < (b.next || '9') ? -1 : 1);
    $('#recList').innerHTML = rec.map(r =>
      '<div class="li" data-rec="' + r.row + '"><div><div>' + esc(r.description) + '</div><div class="li-sub">' + esc(r.repeats) + ' · ' +
      (r.kind === 'transfer' ? esc(r.account) + ' → ' + esc(r.toAccount) : esc(r.account)) + (r.next ? ' · next ' + dateShort(r.next) : ' · done') + '</div></div>' +
      '<div class="li-right">' + (r.kind === 'income' ? '+' : '') + money0(Math.abs(r.amount)) + ' ' + esc(r.currency) + '</div></div>').join('') || '<div class="muted small">Nothing planned yet.</div>';
  }

  function niceRange(lo, hi) {
    if (hi === lo) { hi = lo + 1; }
    const raw = (hi - lo) / 4;
    const p10 = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map(k => k * p10).find(v => v >= raw);
    return { lo: Math.floor(lo / step) * step, hi: Math.ceil(hi / step) * step, step };
  }
  function hover(el, W, n, xOf, html, onMove) {
    const tip = $('#tooltip');
    const move = (clientX) => {
      const r = el.getBoundingClientRect();
      const px = (clientX - r.left) / r.width * W;
      let best = 0, bd = Infinity;
      for (let i = 0; i < n; i++) { const dd = Math.abs(xOf(i) - px); if (dd < bd) { bd = dd; best = i; } }
      tip.innerHTML = html(best); tip.hidden = false;
      if (onMove) onMove(best);
      const tx = Math.min(window.innerWidth - tip.offsetWidth - 8, Math.max(8, r.left + xOf(best) / W * r.width - tip.offsetWidth / 2));
      tip.style.left = tx + 'px'; tip.style.top = Math.max(8, r.top - tip.offsetHeight - 6) + 'px';
    };
    const hide = () => { tip.hidden = true; if (onMove) onMove(-1); };
    el.onpointermove = e => move(e.clientX);
    el.onpointerdown = e => move(e.clientX);
    el.onpointerleave = hide;
    el.onpointerup = e => { if (e.pointerType !== 'mouse') setTimeout(hide, 1800); };
  }
  function yAxis(W, m, y, rg) {
    let g = '<g class="axis">';
    for (let v = rg.lo; v <= rg.hi + rg.step / 2; v += rg.step) {
      g += '<line class="' + (Math.abs(v) < rg.step / 2 ? 'baseline' : 'gridline') + '" x1="' + m.l + '" x2="' + (W - m.r) + '" y1="' + y(v) + '" y2="' + y(v) + '"/>';
      g += '<text x="' + (m.l - 6) + '" y="' + (y(v) + 4) + '" text-anchor="end">' + shortNum(v) + '</text>';
    }
    return g + '</g>';
  }
  // Show fewer month labels when space is tight (about 46px per label).
  function labelEvery(n, plotW) { return [1, 2, 3, 6, 12].find(k => Math.ceil(n / k) * 46 <= plotW) || 12; }
  function shortNum(v) {
    const a = Math.abs(v), sign = v < 0 ? '−' : '';
    if (a >= 1e6) return sign + (a / 1e6).toFixed(a >= 1e7 ? 0 : 1) + 'M';
    if (a >= 1e4) return sign + Math.round(a / 1e3) + 'k';
    if (a >= 1e3) return sign + (a / 1e3).toFixed(1).replace('.0', '') + 'k';
    return sign + Math.round(a);
  }

  function renderForecastChart(p) {
    const el = $('#fcChart');
    const W = el.clientWidth || 340, H = el.clientHeight || 190;
    const m = { t: 14, r: 12, b: 22, l: 46 };
    let rows, cushion, cur;
    if (state.fcAccount === 'all') {
      rows = p.months.map(x => ({ month: x.month, min: x.min, minDate: x.minDate, end: x.end }));
      cushion = p.totalCushionCHF; cur = 'CHF';
    } else {
      const a = p.accounts.find(x => x.name === state.fcAccount);
      rows = a.months; cushion = a.cushion; cur = a.currency;
    }
    $('#fcUnit').textContent = cur;
    const vals = rows.map(r => r.min).concat(rows.map(r => r.end));
    if (cushion != null) vals.push(cushion);
    vals.push(0);
    const rg = niceRange(Math.min.apply(null, vals), Math.max.apply(null, vals));
    const n = rows.length;
    const x = i => m.l + (n === 1 ? 0 : i / (n - 1)) * (W - m.l - m.r);
    const y = v => m.t + (rg.hi - v) / (rg.hi - rg.lo) * (H - m.t - m.b);
    let svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" aria-hidden="true">' + yAxis(W, m, y, rg) + '<g class="axis">';
    const evF = labelEvery(n, W - m.l - m.r);
    rows.forEach((r, i) => { if (i % evF === 0) svg += '<text x="' + x(i) + '" y="' + (H - 4) + '" text-anchor="middle">' + monthShort(r.month) + '</text>'; });
    svg += '</g>';
    if (cushion != null) {
      svg += '<line class="cushion" x1="' + m.l + '" x2="' + (W - m.r) + '" y1="' + y(cushion) + '" y2="' + y(cushion) + '"/>';
      svg += '<text class="label" x="' + (W - m.r) + '" y="' + (y(cushion) - 5) + '" text-anchor="end">Cushion</text>';
    }
    const pts = rows.map((r, i) => x(i).toFixed(1) + ',' + y(r.min).toFixed(1));
    svg += '<polyline class="spent" points="' + pts.join(' ') + '"/>';
    rows.forEach((r, i) => { if (cushion != null && r.min < cushion) svg += '<circle class="below" cx="' + x(i) + '" cy="' + y(r.min) + '" r="4"/>'; });
    svg += '<line class="cross" id="fcCross" y1="' + m.t + '" y2="' + (H - m.b) + '" visibility="hidden"/></svg>';
    el.innerHTML = svg;
    const cross = el.querySelector('#fcCross');
    hover(el, W, n, x, i => {
      const r = rows[i];
      const flag = cushion != null && r.min < cushion ? '<br><span style="color:var(--critical)">▲ Below cushion</span>' : '';
      return '<b>' + monthName(r.month) + '</b><br>Lowest ' + money(r.min, cur) + ' on ' + dateShort(r.minDate).slice(0, -5) + '<br>Month end ' + money(r.end, cur) + flag;
    }, i => { if (i < 0) cross.setAttribute('visibility', 'hidden'); else { cross.setAttribute('x1', x(i)); cross.setAttribute('x2', x(i)); cross.setAttribute('visibility', 'visible'); } });
    $('#fcTable').innerHTML = '<table><tr><th>Month</th><th>Lowest</th><th>Month end</th></tr>' + rows.map(r =>
      '<tr><td>' + monthShort(r.month) + '</td><td' + (cushion != null && r.min < cushion ? ' style="color:var(--critical)"' : '') + '>' + money0(r.min) + '</td><td>' + money0(r.end) + '</td></tr>').join('') + '</table>';
  }

  function renderCashflowChart(p) {
    const el = $('#cfChart');
    const W = el.clientWidth || 340, H = el.clientHeight || 190;
    const m = { t: 10, r: 8, b: 22, l: 46 };
    const rows = p.months;
    const rg = niceRange(0, Math.max.apply(null, rows.map(r => Math.max(r.in, r.out)).concat([1])));
    const n = rows.length;
    const slot = (W - m.l - m.r) / n;
    const bw = Math.max(2, Math.min(10, slot / 2 - 2));
    const xc = i => m.l + slot * (i + 0.5);
    const y = v => m.t + (rg.hi - v) / (rg.hi - rg.lo) * (H - m.t - m.b);
    const bar = (cls, x0, v) => { const h = Math.max(0, y(0) - y(v)); if (h < 0.5) return ''; const r = Math.min(3, bw / 2, h);
      return '<path class="' + cls + '" d="M' + x0 + ',' + y(0) + 'V' + (y(v) + r) + 'Q' + x0 + ',' + y(v) + ' ' + (x0 + r) + ',' + y(v) + 'H' + (x0 + bw - r) + 'Q' + (x0 + bw) + ',' + y(v) + ' ' + (x0 + bw) + ',' + (y(v) + r) + 'V' + y(0) + 'Z"/>'; };
    let svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" aria-hidden="true">' + yAxis(W, m, y, rg) + '<g class="axis">';
    const evC = labelEvery(n, W - m.l - m.r);
    rows.forEach((r, i) => { if (i % evC === 0) svg += '<text x="' + xc(i) + '" y="' + (H - 4) + '" text-anchor="middle">' + monthShort(r.month) + '</text>'; });
    svg += '</g>';
    rows.forEach((r, i) => { svg += bar('bar-in', xc(i) - bw - 1, r.in) + bar('bar-out', xc(i) + 1, r.out); });
    svg += '<rect id="cfHi" y="' + m.t + '" height="' + (H - m.t - m.b) + '" width="' + slot + '" fill="currentColor" opacity="0.06" visibility="hidden"/></svg>';
    el.innerHTML = svg;
    const hi = el.querySelector('#cfHi');
    hover(el, W, n, xc, i => {
      const r = rows[i];
      return '<b>' + monthName(r.month) + '</b><br>Money in ' + money0(r.in) + '<br>Money out ' + money0(r.out) +
        '<br><span style="color:var(--muted)">planned ' + money0(r.planned) + ' · day-to-day ' + money0(r.variable) + '</span><br>Net ' + money0(r.in - r.out) + ' CHF';
    }, i => { if (i < 0) hi.setAttribute('visibility', 'hidden'); else { hi.setAttribute('x', xc(i) - slot / 2); hi.setAttribute('visibility', 'visible'); } });
  }

  // Mark paid / skip
  function openPaidSheet(u) {
    const b = state.boot;
    const isToday = u.date <= today();
    $('#sheetTitle').textContent = u.kind === 'income' ? 'Mark as received' : 'Mark as paid';
    $('#sheetDelete').hidden = true;
    $('#sheetSave').textContent = 'Save';
    const accOpts = b.accounts.map(a => '<option value="' + esc(a.name) + '"' + (a.name === u.account ? ' selected' : '') + '>' + esc(a.name) + '</option>').join('');
    let html = '<p class="muted small">' + esc(u.description) + ' · due ' + dateShort(u.due) + '. This adds the ' + (u.kind === 'transfer' ? 'transfer' : u.kind === 'income' ? 'income' : 'expense') + ' to your records.</p>';
    html += field('Amount (' + esc(u.currency) + ')', '<input id="pAmount" inputmode="decimal" value="' + Number(u.amount).toFixed(2) + '">');
    html += field('Date', '<input id="pDate" type="date" value="' + (isToday ? today() : u.date) + '">');
    html += field(u.kind === 'income' ? 'Received into' : u.kind === 'transfer' ? 'From' : 'Paid from', '<select id="pAcc">' + accOpts + '</select>');
    if (u.kind === 'transfer' && accountCurrency(u.toAccount) !== u.currency) {
      html += field('Amount received in ' + esc(u.toAccount) + ' (' + esc(accountCurrency(u.toAccount)) + ')', '<input id="pReceived" inputmode="decimal" placeholder="As shown by the bank">');
    }
    html += field('Note', '<input id="pNote" placeholder="Optional">');
    $('#sheetBody').innerHTML = html;
    $('#sheet').showModal();
  }
  async function savePaidSheet() {
    const u = sheetMode.item;
    const errEl = $('#sheetError'); errEl.hidden = true;
    const amount = parseNum($('#pAmount').value);
    if (!(amount > 0)) { errEl.textContent = 'Enter an amount.'; errEl.hidden = false; return; }
    const btn = $('#sheetSave'); btn.disabled = true;
    try {
      const rec = $('#pReceived') ? parseNum($('#pReceived').value) : '';
      const r = await api('markPaid', { payment: { row: u.row, description: u.description, due: u.due, date: $('#pDate').value, amount,
        currency: u.currency, account: $('#pAcc').value, received: rec > 0 ? rec : '', note: $('#pNote').value.trim() } });
      state.plan = r.plan; store.set('plan', r.plan); saveSummary(r.summary);
      $('#sheet').close(); toast((u.kind === 'income' ? 'Received' : 'Paid') + ' ✓ added to your records'); renderPlan(); renderAll();
    } catch (e) {
      errEl.textContent = e instanceof NetworkError ? 'This needs a connection.' : e.message; errEl.hidden = false;
    } finally { btn.disabled = false; }
  }
  async function skipUpcoming(u) {
    if (!window.confirm('Skip "' + u.description + '" due ' + dateShort(u.due) + '? Nothing is recorded.')) return;
    try {
      state.plan = await api('skipPayment', { row: u.row, due: u.due, match: { description: u.description } });
      store.set('plan', state.plan); renderPlan(); toast('Skipped');
    } catch (e) { toast(e instanceof NetworkError ? 'This needs a connection.' : e.message, 4000); }
  }

  // Recurring editor
  function openRecurringSheet(r) {
    const b = state.boot, p = state.plan;
    const kind = r ? r.kind : 'expense';
    $('#sheetTitle').textContent = r ? 'Edit planned item' : 'New planned item';
    $('#sheetSave').textContent = 'Save';
    const accList = b.accounts.map(a => a.name);
    let html = '<div class="seg2" id="rKind">' + [['expense', 'Money out'], ['income', 'Money in'], ['transfer', 'Transfer']].map(k =>
      '<button type="button" data-k="' + k[0] + '" class="' + (k[0] === kind ? 'active' : '') + '">' + k[1] + '</button>').join('') + '</div>';
    html += field('Description', '<input id="rDesc" value="' + esc(r ? r.description : '') + '" placeholder="e.g. Rent">');
    html += '<div class="row2">' + field('Amount', '<input id="rAmount" inputmode="decimal" value="' + (r ? Math.abs(r.amount) : '') + '" placeholder="0.00">') +
      field('Currency', '<select id="rCur">' + opts(b.currencies.map(c => c.code), r ? r.currency : 'CHF') + '</select>') + '</div>';
    html += '<div id="rCatWrap">' + field('Category', '<select id="rCat"><option value=""></option>' + opts(b.categories, r && r.kind === 'expense' ? r.category : '') + '</select>') + '</div>';
    html += '<div id="rTypeWrap">' + field('Income type', '<input id="rType" list="rTypeList" value="' + esc(r && r.kind === 'income' ? r.category : 'Salary') + '"><datalist id="rTypeList"><option>Salary</option><option>Bonus</option><option>Refund</option><option>Repayment</option><option>Other</option></datalist>') + '</div>';
    html += field('<span id="rAccLabel">Account</span>', '<select id="rAcc">' + opts(accList, r ? r.account : (b.lastAccount || accList[0])) + '</select>');
    html += '<div id="rToWrap">' + field('To account', '<select id="rTo">' + opts(accList, r ? r.toAccount : '') + '</select>') + '</div>';
    html += '<div class="row2">' + field('Repeats', '<select id="rRep">' + opts(p.repeatsOptions, r ? r.repeats : 'Monthly') + '</select>') +
      field('First due', '<input id="rFirst" type="date" value="' + (r ? r.first : today()) + '">') + '</div>';
    html += '<div class="row2">' + field('Ends on <span class="muted">(optional)</span>', '<input id="rEnds" type="date" value="' + (r ? r.endsOn : '') + '">') +
      field('If on a weekend', '<select id="rWeekend">' + opts(p.weekendOptions, r ? r.weekend : 'No change') + '</select>') + '</div>';
    html += field('Note', '<input id="rNote" value="' + esc(r ? r.note : '') + '" placeholder="Optional">');
    if (r && r.lastPaid) html += '<p class="muted small">Last paid: ' + dateShort(r.lastPaid) + '</p>';
    $('#sheetBody').innerHTML = html;
    const setKind = k => {
      $$('#rKind button').forEach(x => x.classList.toggle('active', x.dataset.k === k));
      $('#rCatWrap').hidden = k !== 'expense'; $('#rTypeWrap').hidden = k !== 'income'; $('#rToWrap').hidden = k !== 'transfer';
      $('#rAccLabel').textContent = k === 'income' ? 'Received into' : k === 'transfer' ? 'From account' : 'Paid from';
      sheetMode.kind = k;
    };
    $('#rKind').onclick = e => { const x = e.target.closest('button'); if (x) setKind(x.dataset.k); };
    setKind(kind);
    $('#sheet').showModal();
  }
  async function saveRecurringSheet() {
    const r = sheetMode.item, k = sheetMode.kind;
    const errEl = $('#sheetError'); errEl.hidden = true;
    const amount = parseNum($('#rAmount').value);
    const item = {
      row: r ? r.row : null, original: r ? r.description : '', kind: k,
      description: $('#rDesc').value.trim(), amount, currency: $('#rCur').value,
      category: k === 'expense' ? $('#rCat').value : k === 'income' ? $('#rType').value.trim() : '',
      account: $('#rAcc').value, toAccount: k === 'transfer' ? $('#rTo').value : '',
      repeats: $('#rRep').value, first: $('#rFirst').value, endsOn: $('#rEnds').value, weekend: $('#rWeekend').value,
      note: $('#rNote').value.trim(), lastPaid: r ? r.lastPaid : '', status: r ? r.status : 'Planned'
    };
    if (r && item.first !== r.first && r.repeats === 'Once' && r.status === 'Paid') item.status = 'Planned';
    const btn = $('#sheetSave'); btn.disabled = true;
    try {
      state.plan = await api('saveRecurring', { item }); store.set('plan', state.plan);
      $('#sheet').close(); toast('Saved ✓'); renderPlan();
    } catch (e) { errEl.textContent = e instanceof NetworkError ? 'This needs a connection.' : e.message; errEl.hidden = false; }
    finally { btn.disabled = false; }
  }
  async function deleteRecurringSheet() {
    const r = sheetMode.item;
    if (!window.confirm('Delete "' + r.description + '" from your plan? Past records stay.')) return;
    try {
      state.plan = await api('deleteRecurring', { row: r.row, match: { description: r.description } }); store.set('plan', state.plan);
      $('#sheet').close(); toast('Deleted'); renderPlan();
    } catch (e) { $('#sheetError').textContent = e instanceof NetworkError ? 'This needs a connection.' : e.message; $('#sheetError').hidden = false; }
  }

  // ---------- Category detail ----------
  async function openCategoryDetail(name) {
    $('#sheetTitle').textContent = name;
    $('#sheetSave').hidden = true; $('#sheetDelete').hidden = true; $('#sheetCancel').textContent = 'Close';
    $('#sheetBody').innerHTML = '<p class="muted">Loading…</p>';
    $('#sheet').showModal();
    try {
      const d = await api('categoryDetail', { name, month: state.month });
      if (!$('#sheet').open || sheetMode.item !== name) return;
      let html = '<div class="kv"><span>Monthly budget</span><span class="num">' + money0(d.budget) + ' CHF</span></div>' +
        '<div class="kv"><span>Average (months with spending)</span><span class="num">' + money0(d.average) + ' CHF</span></div>';
      html += '<h3>Last 12 months</h3><div id="catChart" class="chart"></div>';
      if (d.topMerchants.length) html += '<h3>Where it went (12 months)</h3>' + d.topMerchants.map(t => '<div class="kv"><span>' + esc(t.name) + '</span><span class="num">' + money0(t.chf) + ' CHF</span></div>').join('');
      html += '<h3>' + monthName(state.month) + '</h3><div id="catEntries">' + (d.entries.length ? d.entries.map((e, i) => entryHtml(e, i, 'c')).join('') : '<div class="muted small">Nothing this month.</div>') + '</div>';
      $('#sheetBody').innerHTML = html;
      sheetMode.entries = d.entries;
      $('#catEntries').onclick = ev => { const en = ev.target.closest('[data-c]'); if (!en) return; $('#sheet').close(); openEdit(d.entries[Number(en.dataset.c)]); };
      requestAnimationFrame(() => renderCategoryChart(d));
    } catch (e) { $('#sheetBody').innerHTML = '<p class="form-error">' + esc(e instanceof NetworkError ? 'This needs a connection.' : e.message) + '</p>'; }
  }
  function renderCategoryChart(d) {
    const el = $('#catChart'); if (!el) return;
    const W = el.clientWidth || 320, H = el.clientHeight || 170;
    const m = { t: 12, r: 8, b: 22, l: 40 };
    const rows = d.months, n = rows.length;
    const rg = niceRange(0, Math.max.apply(null, rows.map(r => r.total).concat([d.budget, 1])));
    const slot = (W - m.l - m.r) / n, bw = Math.max(4, Math.min(18, slot - 6));
    const xc = i => m.l + slot * (i + 0.5);
    const y = v => m.t + (rg.hi - v) / (rg.hi - rg.lo) * (H - m.t - m.b);
    let svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" aria-hidden="true">' + yAxis(W, m, y, rg) + '<g class="axis">';
    rows.forEach((r, i) => { if (i % 2 === 1 || n <= 6) svg += '<text x="' + xc(i) + '" y="' + (H - 4) + '" text-anchor="middle">' + monthShort(r.month).split(' ')[0] + '</text>'; });
    svg += '</g>';
    rows.forEach((r, i) => {
      const h = y(0) - y(r.total); if (h < 0.5) return;
      const x0 = xc(i) - bw / 2, rr = Math.min(4, bw / 2, h);
      svg += '<path class="' + (d.budget && r.total > d.budget ? 'below' : 'bar-in') + '" d="M' + x0 + ',' + y(0) + 'V' + (y(r.total) + rr) + 'Q' + x0 + ',' + y(r.total) + ' ' + (x0 + rr) + ',' + y(r.total) + 'H' + (x0 + bw - rr) + 'Q' + (x0 + bw) + ',' + y(r.total) + ' ' + (x0 + bw) + ',' + (y(r.total) + rr) + 'V' + y(0) + 'Z"/>';
    });
    if (d.budget) svg += '<line class="pace" x1="' + m.l + '" x2="' + (W - m.r) + '" y1="' + y(d.budget) + '" y2="' + y(d.budget) + '"/><text class="label" x="' + (W - m.r) + '" y="' + (y(d.budget) - 5) + '" text-anchor="end">Budget</text>';
    svg += '</svg>';
    el.innerHTML = svg;
    hover(el, W, n, xc, i => '<b>' + monthName(rows[i].month) + '</b><br>' + money(rows[i].total, 'CHF') + (d.budget && rows[i].total > d.budget ? '<br><span style="color:var(--critical)">▲ Over budget</span>' : ''));
  }

  // ---------- Search ----------
  let searchTimer;
  function searchQuery() {
    return { text: $('#qText').value.trim(), kind: $('#qKind').value, category: $('#qCat').value, account: $('#qAcc').value,
      from: $('#qFrom').value, to: $('#qTo').value, min: $('#qMin').value.trim() ? parseNum($('#qMin').value) : '', max: $('#qMax').value.trim() ? parseNum($('#qMax').value) : '' };
  }
  function searchActive(q) { return Object.keys(q).some(k => q[k] !== '' && q[k] !== null); }
  async function runSearch() {
    const q = searchQuery();
    const active = searchActive(q);
    $('#searchResults').hidden = !active; $('#monthActivity').hidden = active;
    if (!active) return;
    $('#searchResults').innerHTML = '<div class="muted small">Searching…</div>';
    try {
      const r = await api('search', { query: q });
      state.searchItems = r.entries;
      $('#searchResults').innerHTML = '<div class="result-head"><span>' + r.count + ' result' + (r.count === 1 ? '' : 's') + (r.count > 200 ? ' (showing newest 200)' : '') + '</span>' +
        (r.totalCHF ? '<span>Spent ' + money(r.totalCHF, 'CHF') + '</span>' : '') + '</div>' +
        (r.entries.map((e, i) => '<div class="day-head"><span>' + dateShort(e.date) + '</span></div>' + entryHtml(e, i, 's')).join('') || '<div class="empty">Nothing found.</div>');
    } catch (e) { $('#searchResults').innerHTML = '<div class="form-error">' + esc(e instanceof NetworkError ? 'Search needs a connection.' : e.message) + '</div>'; }
  }
  function fillSearchLists() {
    const b = state.boot; if (!b) return;
    const keep = (id, list) => { const v = $(id).value; $(id).innerHTML = '<option value="">All</option>' + list.map(x => '<option>' + esc(x) + '</option>').join(''); $(id).value = v; };
    keep('#qCat', b.categories); keep('#qAcc', b.accounts.map(a => a.name));
  }

  // ---------- Duplicate ----------
  function duplicateEntry() {
    const ed = state.editing; if (!ed) return;
    state.editing = null;
    $('#editBar').hidden = true; $('#deleteBtn').hidden = true; $('#duplicateBtn').hidden = true;
    $('.segmented').hidden = false; $('#receiptExisting').hidden = true; $('#receiptBtn').textContent = '📷 Add receipt';
    $('#saveBtn').textContent = 'Save';
    $('#fDate').value = today();
    setType(ed.kind);
    $('#title').textContent = 'Add';
    toast('Copied. Change what you need and save.');
    window.scrollTo(0, 0);
  }


  // ---------- Sinking funds and goals ----------
  function renderEnvelopes(E) {
    if (!E) { $('#envList').innerHTML = ''; return; }
    $('#envSummary').textContent = E.envelopes.length ? 'Set aside so far ' + money(E.reservedCHF, 'CHF') + '. To stay on track, set aside about ' + money0(E.monthlyCHF) + ' CHF a month. The money stays in your accounts; the app only keeps it aside.'
      : 'Keep money aside for yearly bills (sinking funds) or for something you are saving towards (goals). The money stays in your accounts.';
    $('#envList').innerHTML = E.envelopes.map((e, i) => {
      const reached = e.left <= 0;
      const pace = reached ? '<span class="done">✓ Reached</span>' : e.neededMonthly != null ? money0(e.neededMonthly) + ' CHF a month until ' + dateShort(e.targetDate) : 'No target date';
      const btn = reached ? '' : e.setAsideThisMonth ? '<span class="done">✓ Set aside this month</span>' : '<button class="mini primary" data-setaside="' + i + '">Set aside ' + money0(e.suggested) + '</button>';
      return '<div class="env" data-env="' + i + '"><div class="env-top"><div><div class="env-name">' + esc(e.name) + '</div><div class="env-type">' + esc(e.type) + (e.linked ? ' · pays ' + esc(e.linked) : '') + '</div></div>' +
        '<div class="env-nums">' + money0(e.saved) + ' / ' + money0(e.target) + ' CHF</div></div>' +
        '<div class="bar"><span class="' + (reached ? 'good' : '') + '" style="width:' + e.pct + '%"></span></div>' +
        '<div class="env-foot"><span>' + pace + (!reached && e.monthly != null && !e.onTrack ? ' · <span style="color:var(--critical)">▲ your ' + money0(e.monthly) + ' a month is not enough</span>' : '') + '</span>' + btn + '</div></div>';
    }).join('');
    $('#envSuggest').innerHTML = E.suggestions.length ? '<div class="sugg"><div class="sugg-title">Bills that come less than monthly, without a fund</div>' + E.suggestions.map((x, i) =>
      '<div class="sugg-row"><span>' + esc(x.description) + ' · ' + money0(x.amountCHF) + ' CHF on ' + dateShort(x.next) + '</span><button class="mini" data-sugg="' + i + '">Create fund</button></div>').join('') + '</div>' : '';
  }
  function openEnvelopeSheet(e) {
    const p = state.plan;
    const fromSugg = e && e.fromSuggestion;
    const editing = e && !fromSugg;
    $('#sheetTitle').textContent = editing ? e.name : 'New fund or goal';
    $('#sheetSave').textContent = 'Save';
    $('#sheetDelete').hidden = !editing;
    const plannedNames = p.recurring.filter(r => r.kind === 'expense').map(r => r.description);
    const v = e || {};
    let html = '<div class="seg2" id="eType" style="grid-template-columns:1fr 1fr">' + ['Sinking fund', 'Goal'].map(t =>
      '<button type="button" data-t="' + t + '" class="' + ((v.type || 'Goal') === t ? 'active' : '') + '">' + t + '</button>').join('') + '</div>';
    html += '<p class="muted small" id="eHelp"></p>';
    html += field('Name', '<input id="eName" value="' + esc(v.name || '') + '" placeholder="e.g. Emergency fund">');
    html += '<div class="row2">' + field('Target (CHF)', '<input id="eTarget" inputmode="decimal" value="' + (v.target || '') + '">') +
      field('By', '<input id="eDate" type="date" value="' + (v.targetDate || '') + '">') + '</div>';
    html += '<div class="row2">' + field('Saved so far (CHF)', '<input id="eSaved" inputmode="decimal" value="' + (v.saved != null ? v.saved : 0) + '">') +
      field('Monthly <span class="muted">(optional)</span>', '<input id="eMonthly" inputmode="decimal" value="' + (v.monthly != null ? v.monthly : '') + '" placeholder="Auto">') + '</div>';
    html += '<div id="eLinkWrap">' + field('Pays for planned bill', '<select id="eLinked"><option value="">None</option>' + opts(plannedNames, v.linked || '') + '</select>') + '</div>';
    html += field('Note', '<input id="eNote" value="' + esc(v.note || '') + '" placeholder="Optional">');
    if (editing) html += '<div class="recon"><h3>Take money out</h3><p class="muted small">For example when you used part of it. Nothing moves between accounts.</p>' +
      '<div class="row2"><input id="eOut" inputmode="decimal" placeholder="Amount CHF"><button type="button" class="btn-secondary" id="eOutBtn" style="margin-top:5px">Take out</button></div></div>';
    $('#sheetBody').innerHTML = html;
    const setT = t => {
      $$('#eType button').forEach(x => x.classList.toggle('active', x.dataset.t === t));
      $('#eHelp').textContent = t === 'Sinking fund' ? 'For a bill that comes once a year or a few times a year. When you mark the linked bill as paid, it is paid from this fund and the target moves to the next due date.' : 'For something you are saving towards, like an emergency fund or a trip.';
      $('#eLinkWrap').hidden = t !== 'Sinking fund';
      sheetMode.type = t;
    };
    $('#eType').onclick = ev => { const x = ev.target.closest('button'); if (x) setT(x.dataset.t); };
    setT(v.type || 'Goal');
    if ($('#eOutBtn')) $('#eOutBtn').onclick = async () => {
      const n = parseNum($('#eOut').value);
      if (!(n > 0)) { $('#sheetError').textContent = 'Enter an amount.'; $('#sheetError').hidden = false; return; }
      try { setEnvelopes(await api('moveEnvelope', { row: e.row, match: { name: e.name }, amount: -n, reason: 'use' })); $('#sheet').close(); toast('Taken out'); }
      catch (err) { $('#sheetError').textContent = err instanceof NetworkError ? 'This needs a connection.' : err.message; $('#sheetError').hidden = false; }
    };
    $('#sheet').showModal();
  }
  function setEnvelopes(E) { state.plan.envelopes = E; store.set('plan', state.plan); renderEnvelopes(E); }
  async function saveEnvelopeSheet() {
    const e = sheetMode.item, editing = e && !e.fromSuggestion;
    const errEl = $('#sheetError'); errEl.hidden = true;
    const t = sheetMode.type;
    const env = { row: editing ? e.row : null, original: editing ? e.name : '', name: $('#eName').value.trim(), type: t,
      target: parseNum($('#eTarget').value), targetDate: $('#eDate').value, saved: $('#eSaved').value.trim() === '' ? 0 : parseNum($('#eSaved').value),
      monthly: $('#eMonthly').value.trim() === '' ? '' : parseNum($('#eMonthly').value), linked: t === 'Sinking fund' ? $('#eLinked').value : '', note: $('#eNote').value.trim() };
    const btn = $('#sheetSave'); btn.disabled = true;
    try { setEnvelopes(await api('saveEnvelope', { envelope: env })); $('#sheet').close(); toast('Saved ✓'); loadMonth(state.month); }
    catch (err) { errEl.textContent = err instanceof NetworkError ? 'This needs a connection.' : err.message; errEl.hidden = false; }
    finally { btn.disabled = false; }
  }
  async function deleteEnvelopeSheet() {
    const e = sheetMode.item;
    if (!window.confirm('Delete "' + e.name + '"? The money stays in your accounts; only the reservation is removed.')) return;
    try { setEnvelopes(await api('deleteEnvelope', { row: e.row, match: { name: e.name } })); $('#sheet').close(); toast('Deleted'); loadMonth(state.month); }
    catch (err) { $('#sheetError').textContent = err instanceof NetworkError ? 'This needs a connection.' : err.message; $('#sheetError').hidden = false; }
  }
  async function setAside(e) {
    try { setEnvelopes(await api('moveEnvelope', { row: e.row, match: { name: e.name }, amount: e.suggested, reason: 'monthly' })); toast('Set aside ' + money0(e.suggested) + ' CHF ✓'); loadMonth(state.month); }
    catch (err) { toast(err instanceof NetworkError ? 'This needs a connection.' : err.message, 4000); }
  }

  // ---------- Transfer costs ----------
  const fmtRate = (r) => Number(r).toFixed(r >= 10 ? 2 : r >= 1 ? 4 : 6);
  async function loadCosts() {
    if (!cfg().url) return;
    try {
      const c = await api('transferCosts');
      if (!c.transfers.length) { $('#costBox').innerHTML = 'No transfers between currencies with an "Amount received" yet. Fill in the amount that arrived when you add a transfer, and its cost shows here.'; return; }
      let html = '<div class="big-stat"><div>This year<b>' + money(c.yearCostCHF, 'CHF') + '</b></div><div>Average cost<b>' + c.avgPct.toFixed(2) + '%</b></div><div>Transfers<b>' + c.yearCount + '</b></div></div>';
      if (c.routes.length > 1) html += '<h3 class="small muted" style="margin:10px 0 2px">By route (cheapest first)</h3>' + c.routes.map(r => '<div class="cost-row"><span>' + esc(r.route) + ' · ' + r.count + '×</span><span class="num">' + r.avgPct.toFixed(2) + '%</span></div>').join('');
      html += '<h3 class="small muted" style="margin:10px 0 2px">Recent</h3>' + c.transfers.slice(0, 8).map(t =>
        '<div class="cost-row"><span>' + dateShort(t.date).slice(0, -5) + ' · ' + money0(t.sent) + ' ' + esc(t.sentCurrency) + ' → ' + money0(t.received) + ' ' + esc(t.receivedCurrency) +
        '<br><span class="muted small">bank rate ' + fmtRate(t.bankRate) + ' · market ' + fmtRate(t.marketRate) + (t.approximate ? ' (approx.)' : '') + '</span></span><span class="num">' + t.pct.toFixed(2) + '%<br><span class="muted small">' + money(t.costCHF) + ' CHF</span></span></div>').join('');
      $('#costBox').innerHTML = html; $('#costBox').classList.remove('muted', 'small');
    } catch (e) { if (!(e instanceof NetworkError)) $('#costBox').textContent = e.message; }
  }

  // ---------- Settings ----------
  async function onSettingsSave() {
    const url = $('#sUrl').value.trim(), key = $('#sKey').value.trim();
    const out = $('#sResult');
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) { out.textContent = 'The URL should start with https://script.google.com/ and end with /exec'; return; }
    store.set('cfg', { url, key });
    out.textContent = 'Testing…';
    try {
      await api('summary', { month: state.month });
      out.textContent = '✓ Connected';
      loadBoot();
    } catch (e) { out.textContent = '✕ ' + e.message; }
  }

  // ---------- Wire up ----------
  function init() {
    $$('.tabbar button').forEach(b => b.addEventListener('click', () => { if (state.editing) resetForm(); show(b.dataset.view); }));
    $('#prevMonth').addEventListener('click', () => loadMonth(shiftMonth(state.month, -1)));
    $('#nextMonth').addEventListener('click', () => loadMonth(shiftMonth(state.month, 1)));
    $$('[data-month]').forEach(b => b.addEventListener('click', () => loadMonth(shiftMonth(state.month, Number(b.dataset.month)))));
    $$('.segmented button').forEach(b => b.addEventListener('click', () => setType(b.dataset.type)));
    $('#fCurrency').addEventListener('click', e => { const b = e.target.closest('.chip'); if (b) setCurrency(b.dataset.cur); });
    $('#fCategory').addEventListener('click', e => { const b = e.target.closest('.chip'); if (b) setCategory(b.dataset.cat); });
    $('#fMerchant').addEventListener('change', onMerchant);
    $('#fItem').addEventListener('change', onItem);
    $('#fAccount').addEventListener('change', () => { $('#fAccount').dataset.touched = '1'; });
    $('#fFrom').addEventListener('change', updateTransferCurrencies);
    $('#fTo').addEventListener('change', updateTransferCurrencies);
    $('#receiptBtn').addEventListener('click', () => $('#fReceipt').click());
    $('#fReceipt').addEventListener('change', e => onReceiptChosen(e.target.files[0]));
    $('#receiptClear').addEventListener('click', clearReceipt);
    $('#addForm').addEventListener('submit', onSave);
    $('#sSave').addEventListener('click', onSettingsSave);
    $('#syncNow').addEventListener('click', flushQueue);
    $('#activity').addEventListener('click', e => {
      const en = e.target.closest('.entry'); if (!en) return;
      const item = state.activityItems[Number(en.dataset.i)];
      if (!item) return;
      if (item.pending) { toast('This entry is still waiting to upload.'); return; }
      openEdit(item);
    });
    $('#cancelEdit').addEventListener('click', () => { resetForm(); show('activity'); });
    $('#deleteBtn').addEventListener('click', onDelete);
    $('#receiptRemove').addEventListener('click', () => {
      if (state.editing) state.editing.removeReceipt = true;
      $('#receiptExisting').hidden = true; $('#receiptBtn').textContent = '📷 Add receipt';
    });
    $('#addAccount').addEventListener('click', () => openSheet('account', null));
    $('#addCategory').addEventListener('click', () => openSheet('category', null));
    $('#accList').addEventListener('click', e => { const li = e.target.closest('[data-acc]'); if (li) openSheet('account', state.boot.accounts[Number(li.dataset.acc)]); });
    $('#catList').addEventListener('click', e => { const li = e.target.closest('[data-cat]'); if (li) openSheet('category', state.boot.budgets[Number(li.dataset.cat)]); });
    $('#sheetForm').addEventListener('submit', onSheetSave);
    $('#sheetCancel').addEventListener('click', () => $('#sheet').close());
    $('#sheetDelete').addEventListener('click', onSheetDelete);
    $('#duplicateBtn').addEventListener('click', duplicateEntry);
    $('#categories').addEventListener('click', e => { const c = e.target.closest('[data-catname]'); if (c) openSheet('catdetail', c.dataset.catname); });
    $('#fcAccounts').addEventListener('click', e => { const c = e.target.closest('[data-fc]'); if (!c) return; state.fcAccount = c.dataset.fc; renderPlan(); });
    $('#upcoming').addEventListener('click', e => {
      const p = e.target.closest('[data-paid]'), k = e.target.closest('[data-skip]');
      if (p) openSheet('paid', state.plan.upcoming[Number(p.dataset.paid)]);
      if (k) skipUpcoming(state.plan.upcoming[Number(k.dataset.skip)]);
    });
    $('#addEnvelope').addEventListener('click', () => { if (state.plan) openSheet('envelope', null); });
    $('#envList').addEventListener('click', ev => {
      const sa = ev.target.closest('[data-setaside]');
      if (sa) { ev.stopPropagation(); setAside(state.plan.envelopes.envelopes[Number(sa.dataset.setaside)]); return; }
      const en = ev.target.closest('[data-env]'); if (en) openSheet('envelope', state.plan.envelopes.envelopes[Number(en.dataset.env)]);
    });
    $('#envSuggest').addEventListener('click', ev => {
      const b = ev.target.closest('[data-sugg]'); if (!b) return;
      const x = state.plan.envelopes.suggestions[Number(b.dataset.sugg)];
      openSheet('envelope', { fromSuggestion: true, type: 'Sinking fund', name: x.description, target: x.amountCHF, targetDate: x.next, saved: 0, linked: x.description });
    });
    $('#addRecurring').addEventListener('click', () => { if (state.plan && state.boot) openSheet('recurring', null); });
    $('#recList').addEventListener('click', e => { const li = e.target.closest('[data-rec]'); if (li) openSheet('recurring', state.plan.recurring.find(r => r.row === Number(li.dataset.rec))); });
    $('#qText').addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(runSearch, 400); });
    $('#qFilterBtn').addEventListener('click', () => { $('#qFilters').hidden = !$('#qFilters').hidden; fillSearchLists(); });
    $('#qGo').addEventListener('click', runSearch);
    $('#qClear').addEventListener('click', () => { ['#qText', '#qFrom', '#qTo', '#qMin', '#qMax'].forEach(id => { $(id).value = ''; }); ['#qKind', '#qCat', '#qAcc'].forEach(id => { $(id).value = ''; }); runSearch(); });
    $('#searchResults').addEventListener('click', e => { const en = e.target.closest('[data-s]'); if (en) openEdit(state.searchItems[Number(en.dataset.s)]); });
    window.addEventListener('online', flushQueue);
    window.addEventListener('resize', () => { const s = state.summaries[state.month]; if (s && state.view === 'home') renderChart(s); if (state.view === 'plan' && state.plan) renderPlan(); });
    document.addEventListener('visibilitychange', () => { if (!document.hidden && cfg().url) { flushQueue(); loadMonth(state.month); } });

    const c = cfg(); $('#sUrl').value = c.url; $('#sKey').value = c.key;
    setQueue(queue());
    fillFormLists();
    renderAll();
    if (c.url && c.key) { loadBoot().then(flushQueue); } else { show('settings'); }

    if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  init();
})();
