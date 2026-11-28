/* DPR Assistant — conversational front end. Everything runs in this browser. */
(function () {
  'use strict';
  const M = window.DPRModel, X = window.XlsxFill;
  const AI_URL = 'https://dpr-assistant-ai.cakashoza.workers.dev';
  const STORE = 'dpr-assistant-v1';
  function norm(a) { const f = fresh(); a = a || f; for (const k of Object.keys(f)) if (a[k] == null) a[k] = f[k]; for (const k of Object.keys(f.facilities)) if (!Array.isArray(a.facilities[k])) a.facilities[k] = []; if (!a.project.items) a.project.items = []; return a; }
  let A = norm(load());
  let idx = 0, aiOK = false, pending = new Set(), aiConsent = localStorage.getItem(STORE + '-consent') === 'yes', narrative = null;
  const $ = s => document.querySelector(s);
  function fresh() { return { business: { state: 'Gujarat', promoters: [{}] }, project: { items: [] }, funds: {}, facilities: { existing: [], proposedTL: [], wc: [], nonfund: [] }, collateral: [], ops: {}, hist: {}, settings: {}, narrative: {} }; }
  function load() { try { return JSON.parse(localStorage.getItem(STORE)); } catch { return null; } }
  function save() { localStorage.setItem(STORE, JSON.stringify(A)); }
  const get = (p, o = A) => p.split('.').reduce((x, k) => (x == null ? undefined : x[k]), o);
  function set(p, v, o = A) { const ks = p.split('.'); let x = o; ks.slice(0, -1).forEach((k, i) => { if (x[k] == null) x[k] = /^\d+$/.test(ks[i + 1]) ? [] : {}; x = x[k]; }); x[ks[ks.length - 1]] = v; }
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const num = s => { if (s === '' || s == null) return null; const v = Number(String(s).replace(/[,₹\s]/g, '')); return isFinite(v) ? v : null; };
  const f2 = v => v == null || !isFinite(v) ? '—' : (Math.round(v * 100) / 100).toLocaleString('en-IN', { maximumFractionDigits: 2 });
  const pctShow = v => v == null ? '' : String(Math.round(v * 1e6) / 1e4);
  const isRev = () => M.revenueDirect(A.activity);
  const hasHist = () => A.stage === 'expansion' && A.history && A.history !== 'none';
  const yearsN = () => A.baseFY ? M.projYears(A) : 0;
  const FY = j => M.fyLabel(A.baseFY + 1 + j);

  // ---------------- questions ----------------
  const ACT = [['manufacturing', 'We make products'], ['jobwork', 'Job work for others'], ['trading', 'We buy and sell (trading)'], ['contract', 'Contract-based work'], ['service', 'A service'], ['mixed', 'A mix (make + trade / job work)']];
  const CONST = [['Proprietorship', 'Proprietorship'], ['Partnership Firm', 'Partnership firm'], ['LLP', 'LLP'], ['Private Limited Company', 'Private limited company'], ['Public Limited Company', 'Public limited company']];
  const rmWord = () => ({ trading: 'Purchases of goods', jobwork: 'Consumables & stores used' }[A.activity] || 'Raw material used');
  const kindOpts = () => {
    const all = Object.entries(M.KINDS).map(([k, d]) => [k, d.label]);
    return A.activity === 'trading' || A.activity === 'service' ? all.filter(([k]) => !['machinery', 'electrification'].includes(k)) : all;
  };
  const CARDS = [
    { id: 'welcome', sec: 'Start', title: 'Let\'s prepare your bank loan file', custom: welcome },
    { id: 'name', sec: 'About you', title: 'What is your business called?', ai: true, fields: [
      { t: 'text', p: 'business.name', label: 'Business name', ph: 'e.g. Shree Ganesh Industries' },
      { t: 'choice', p: 'business.constitution', label: 'How is it set up?', opts: CONST }] },
    { id: 'promoters', sec: 'About you', title: 'Who runs the business?', help: 'Banks look closely at the people behind the business. Add each owner / partner / director.', ai: true, fields: [
      { t: 'list', p: 'business.promoters', add: 'Add another person', item: [
        { k: 'name', t: 'text', label: 'Name' }, { k: 'role', t: 'text', label: 'Role / share', ph: 'e.g. Partner, 50%' },
        { k: 'years', t: 'num', label: 'Years of experience', unit: 'years' }, { k: 'experienceLine', t: 'text', label: 'In which line?', ph: 'e.g. CNC machining, yarn trading' },
        { k: 'qualification', t: 'text', label: 'Qualification', ph: 'e.g. B.E. Mechanical' }] }] },
    { id: 'activity', sec: 'Your business', title: 'What kind of work do you do?', ai: true, fields: [
      { t: 'choice', p: 'activity', opts: ACT },
      { t: 'text', p: 'business.industry', label: 'Industry', ph: 'e.g. Textiles - embroidery, Engineering, Chemicals' },
      { t: 'text', p: 'business.activityDesc', label: 'In one line, what do you make or do?', ph: 'e.g. CNC turned auto components' }] },
    { id: 'where', sec: 'Your business', title: 'Where is the unit?', fields: [
      { t: 'text', p: 'business.area', label: 'Area / GIDC estate', ph: 'e.g. GIDC Sachin' },
      { t: 'text', p: 'business.district', label: 'District', ph: 'e.g. Surat' },
      { t: 'text', p: 'business.state', label: 'State' }] },
    { id: 'jobwork', sec: 'Your business', when: () => A.activity === 'jobwork', title: 'Who gives you job work?', ai: true, fields: [
      { t: 'textarea', p: 'jobwork.principals', label: 'Main principals you work for (types of companies are fine)' },
      { t: 'text', p: 'jobwork.basis', label: 'How do you charge?', ph: 'e.g. per metre, per piece, per kg' }] },
    { id: 'orderbook', sec: 'Your business', when: () => A.activity === 'contract' || A.activity === 'service', title: 'What work do you have in hand?', ai: true, fields: [
      { t: 'num', p: 'orderBook.value', label: 'Value of orders / contracts in hand', unit: '₹ lakh' },
      { t: 'textarea', p: 'orderBook.customers', label: 'Main customers or contracts' },
      { t: 'num', p: 'orderBook.months', label: 'Months to complete them', unit: 'months' }] },
    { id: 'stage', sec: 'Your business', title: 'Is this a new unit or are you expanding?', fields: [
      { t: 'choice', p: 'stage', opts: [['new', 'Starting a new unit'], ['expansion', 'Expanding a running business']] },
      { t: 'choice', p: 'history', label: 'Do you have accounts for the last 3 years?', when: () => A.stage === 'expansion', opts: [['audited', 'Yes, audited'], ['provisional', 'Yes, some provisional'], ['none', 'No']] }] },
    { id: 'basefy', sec: 'Your business', title: () => A.stage === 'new' ? 'When does the unit start?' : 'Which year are your latest accounts for?', fields: [
      { t: 'year', p: 'baseFY', label: () => A.stage === 'new' ? 'The financial year that ends just before you start (year ending 31 March)' : 'Latest accounts are for the year ending 31 March of', required: true },
      { t: 'info', html: () => A.baseFY ? `Your projections will start from <b>FY${String(A.baseFY + 1).slice(2)}</b> (April ${A.baseFY} – March ${A.baseFY + 1}).` : '' },
      { t: 'text', p: 'business.auditor', label: 'Your CA / auditor firm', when: hasHist }] },
    // past accounts
    { id: 'histhow', sec: 'Past accounts', when: hasHist, title: 'Your last 3 years\' accounts', custom: histHow },
    { id: 'hsales', sec: 'Past accounts', when: hasHist, title: 'What were your sales?', help: 'All amounts in ₹ lakh (₹1,00,000 = 1 lakh).', fields: [{ t: 'hist', rows: [8, 9, 11], cls: true }] },
    { id: 'hcost', sec: 'Past accounts', when: hasHist, title: 'What did it cost to run?', fields: [{ t: 'hist', rows: [14, 13, 15, 16, 17, 18, 20, 21] }] },
    { id: 'hother', sec: 'Past accounts', when: hasHist, title: 'Office costs, interest and tax', fields: [{ t: 'hist', rows: [23, 25, 26, 27, 29] }] },
    { id: 'howe', sec: 'Past accounts', when: hasHist, title: 'What did the business owe at each year end?', fields: [{ t: 'hist', rows: [33, 34, 35, 36, 37, 39, 40] }] },
    { id: 'hown', sec: 'Past accounts', when: hasHist, title: 'Owners\' money in the business', fields: [{ t: 'hist', rows: [41, 42] }] },
    { id: 'hhave', sec: 'Past accounts', when: hasHist, title: 'What did the business own?', fields: [{ t: 'hist', rows: [44, 45, 46, 47, 48, 50, 51, 53] }, { t: 'info', html: tieInfo }] },
    { id: 'assets', sec: 'Past accounts', when: () => A.stage === 'expansion', title: 'Book value of the assets you already have', help: 'Written-down value at the end of your latest year, from the depreciation schedule. Leave a type blank if you have none.', fields: M.CLASSES.map(c => ({ t: 'num', p: 'existingAssets.' + c, label: c, unit: '₹ lakh' })) },
    // project
    { id: 'project', sec: 'The project', title: () => A.activity === 'trading' || A.activity === 'service' ? 'Anything you are buying for the business? (shop, godown, vehicle…)' : 'What are you buying or building?', help: 'Add land, building, each machine (with supplier), electrification, pre-operative expenses and contingency. Cost before GST, in ₹ lakh.', ai: true, fields: [
      { t: 'list', p: 'project.items', add: 'Add another item', item: [
        { k: 'kind', t: 'select', label: 'Type', opts: kindOpts },
        { k: 'desc', t: 'text', label: 'What is it?', ph: 'e.g. 2 CNC turning centres' },
        { k: 'supplier', t: 'text', label: 'Supplier / quotation from' },
        { k: 'cost', t: 'num', label: 'Cost before GST', unit: '₹ lakh' },
        { k: 'gst', t: 'pct', label: 'GST rate', defFrom: 'kind' },
        { k: 'months', t: 'num', label: 'Months it will run in the first year', unit: 'months' },
        { k: 'loan', t: 'select', label: 'Fund with bank loan?', opts: [['Y', 'Yes'], ['N', 'No']] },
        { k: 'cls', t: 'select', label: 'Asset type for depreciation', opts: () => M.CLASSES.map(c => [c, c]), defFrom: 'kind', adv: true }] },
      { t: 'num', p: 'project.marginMoneyWC', label: 'Margin money for working capital (if you include it in project cost)', unit: '₹ lakh' },
      { t: 'info', html: costInfo }] },
    { id: 'funds', sec: 'Funding', title: 'How much are you putting in yourself?', ai: true, fields: [
      { t: 'num', p: 'funds.promoterCapital', label: 'New capital you are bringing in', unit: '₹ lakh' },
      { t: 'num', p: 'funds.unsecuredLoan', label: 'New loans from family / friends / directors (unsecured)', unit: '₹ lakh' },
      { t: 'num', p: 'funds.subsidy', label: 'Subsidy you expect, if any', unit: '₹ lakh' },
      { t: 'text', p: 'funds.subsidyScheme', label: 'Which scheme?', ph: 'e.g. Gujarat Atmanirbhar capital subsidy' },
      { t: 'info', html: fundInfo }] },
    { id: 'existing', sec: 'Funding', when: () => A.stage === 'expansion', title: 'Any term loans still running?', help: 'Enter what is still outstanding at the end of your latest year. Skip if none.', ai: true, fields: [
      { t: 'list', p: 'facilities.existing', add: 'Add a running loan', min: 0, item: [
        { k: 'name', t: 'text', label: 'Loan name', ph: 'e.g. Machinery loan - XYZ Bank' },
        { k: 'type', t: 'select', label: 'Type', opts: [['Term Loan', 'Term loan'], ['WCTL / GECL', 'WCTL / GECL']] },
        { k: 'outstanding', t: 'num', label: 'Amount outstanding', unit: '₹ lakh' }, { k: 'rate', t: 'pct', label: 'Interest rate (p.a.)' },
        { k: 'remainingMonths', t: 'num', label: 'Months left to repay', unit: 'months' },
        { k: 'basis', t: 'select', label: 'How is it repaid?', opts: [['Principal Run-Down', 'Equal principal each month'], ['Level EMI', 'Fixed EMI']] }] }] },
    { id: 'newtl', sec: 'Funding', title: 'What term loan are you asking for?', ai: true, fields: [
      { t: 'list', p: 'facilities.proposedTL', add: 'Add another term loan', min: 0, item: [
        { k: 'name', t: 'text', label: 'Loan name', ph: 'e.g. Machinery term loan' },
        { k: 'amount', t: 'num', label: 'Amount', unit: '₹ lakh' }, { k: 'rate', t: 'pct', label: 'Interest rate (p.a.)' },
        { k: 'start', t: 'date', label: 'Expected disbursement date' },
        { k: 'moratorium', t: 'num', label: 'Months before repayment starts (moratorium)', unit: 'months' },
        { k: 'tenor', t: 'num', label: 'Months to repay (after moratorium)', unit: 'months' },
        { k: 'basis', t: 'select', label: 'Repayment', opts: [['Level EMI', 'Fixed EMI'], ['Principal Run-Down', 'Equal principal each month']] }] },
      { t: 'info', html: tlInfo }] },
    { id: 'wc', sec: 'Funding', title: 'Working-capital limit (cash credit / overdraft)', help: 'Existing or proposed. Skip if none.', ai: true, fields: [
      { t: 'list', p: 'facilities.wc', add: 'Add another limit', min: 0, item: [
        { k: 'type', t: 'select', label: 'Type', opts: [['Cash Credit', 'Cash credit'], ['Overdraft', 'Overdraft'], ['Bill Discounting', 'Bill discounting']] },
        { k: 'name', t: 'text', label: 'Name', ph: 'e.g. Proposed CC limit' }, { k: 'amount', t: 'num', label: 'Limit', unit: '₹ lakh' },
        { k: 'rate', t: 'pct', label: 'Interest rate (p.a.)' }, { k: 'start', t: 'date', label: 'From date' },
        { k: 'tenor', t: 'num', label: 'Tenor', unit: 'months' }] }] },
    { id: 'nonfund', sec: 'Funding', title: 'Bank guarantee or letter of credit?', help: 'Skip if none.', fields: [
      { t: 'list', p: 'facilities.nonfund', add: 'Add another', min: 0, item: [
        { k: 'type', t: 'select', label: 'Type', opts: [['Bank Guarantee', 'Bank guarantee'], ['Letter of Credit', 'Letter of credit']] },
        { k: 'name', t: 'text', label: 'Purpose', ph: 'e.g. Power security BG' }, { k: 'amount', t: 'num', label: 'Limit', unit: '₹ lakh' },
        { k: 'rate', t: 'pct', label: 'Commission (p.a.)' }, { k: 'start', t: 'date', label: 'From date' }, { k: 'tenor', t: 'num', label: 'Tenor', unit: 'months' }] }] },
    { id: 'collateral', sec: 'Security', title: 'What security can you offer?', help: 'Property, FDs, machinery etc. Value as per latest valuation if available.', ai: true, fields: [
      { t: 'list', p: 'collateral', add: 'Add another security', min: 0, item: [
        { k: 'type', t: 'text', label: 'Type', ph: 'e.g. Factory land & building' }, { k: 'description', t: 'text', label: 'Description / address area' },
        { k: 'value', t: 'num', label: 'Value', unit: '₹ lakh' }, { k: 'owner', t: 'text', label: 'Owner', ph: 'e.g. Firm, Partner' },
        { k: 'basis', t: 'text', label: 'Basis of value', ph: 'e.g. Valuer report Mar-2026' }] }] },
    // operations
    { id: 'capacity', sec: 'Running it', when: () => !isRev(), title: () => A.activity === 'jobwork' ? 'How much job work can you do, and at what rate?' : 'How much can you make, and at what price?', ai: true, fields: [
      { t: 'text', p: 'ops.unitName', label: 'Count it in', ph: 'e.g. pieces, metres, kg, components' },
      { t: 'series', p: 'ops.capacity', label: () => `Full capacity per year (${A.ops.unitName || 'units'})`, kind: 'num' },
      { t: 'series', p: 'ops.util', label: 'How much of that capacity you expect to use', kind: 'pct' },
      { t: 'series', p: 'ops.price', label: () => A.activity === 'jobwork' ? 'Job charge per unit (₹)' : 'Average selling price per unit (₹)', kind: 'num' },
      { t: 'series', p: 'ops.exportShare', label: 'Share of sales exported', kind: 'pct', opt: true },
      { t: 'info', html: salesInfo }] },
    { id: 'salesplan', sec: 'Running it', when: isRev, title: 'What sales do you expect each year?', ai: true, fields: [
      { t: 'series', p: 'ops.salesPlan', label: 'Expected sales / revenue (₹ lakh)', kind: 'num' },
      { t: 'series', p: 'ops.exportShare', label: 'Share of sales exported', kind: 'pct', opt: true }] },
    { id: 'costs', sec: 'Running it', title: 'What does it cost to run, as a share of sales?', ai: true, fields: [
      { t: 'series', p: 'ops.rmPct', label: () => rmWord() + ' (% of sales)', kind: 'pct', bm: 'rmPct' },
      { t: 'series', p: 'ops.importShare', label: 'Share of raw material imported', kind: 'pct', opt: true, when: () => A.activity === 'manufacturing' || A.activity === 'mixed' },
      { t: 'series', p: 'ops.powerPct', label: 'Power & fuel (% of sales)', kind: 'pct', bm: 'powerPct' },
      { t: 'series', p: 'ops.otherMfgPct', label: () => A.activity === 'trading' ? 'Other direct expenses (% of sales)' : 'Other factory expenses (% of sales)', kind: 'pct', bm: 'otherMfgPct' },
      { t: 'series', p: 'ops.sgaPct', label: 'Office, selling & admin (% of sales)', kind: 'pct', bm: 'sgaPct' }] },
    { id: 'team', sec: 'Running it', title: 'Your team', ai: true, fields: [
      { t: 'series', p: 'ops.heads', label: 'Number of workers & staff', kind: 'num' },
      { t: 'series', p: 'ops.days', label: 'Working days in a year', kind: 'num' },
      { t: 'series', p: 'ops.wage', label: 'Average daily wage per person (₹)', kind: 'num', grow: true }] },
    { id: 'wcdays', sec: 'Running it', title: 'How long does money stay stuck?', ai: true, fields: [
      { t: 'series', p: 'ops.rmDays', label: () => A.activity === 'trading' ? 'Days of stock you keep' : 'Days of raw material you keep', kind: 'num', bm: 'rmDays' },
      { t: 'series', p: 'ops.fgDays', label: 'Days of finished goods you keep', kind: 'num', when: () => !isRev() },
      { t: 'series', p: 'ops.debtorDays', label: 'Days customers take to pay you', kind: 'num', bm: 'debtorDays' },
      { t: 'series', p: 'ops.creditorDays', label: 'Days you take to pay suppliers', kind: 'num', bm: 'creditorDays' },
      { t: 'series', p: 'ops.otherNCA', label: 'Deposits / BG margin / other long-term assets (₹ lakh)', kind: 'num', opt: true }] },
    { id: 'settings', sec: 'Bank settings', title: 'A few bank settings (pre-filled, change if needed)', custom: settingsCard },
    { id: 'story', sec: 'Your story', title: 'Tell us about the business in your own words (optional)', help: 'Used only for the written part of the report. Please do not type PAN, Aadhaar, account numbers or phone numbers.', fields: [
      { t: 'textarea', p: 'narrative.project', label: 'Why this project? What will it let you do?' },
      { t: 'textarea', p: 'narrative.process', label: 'How is the product made / service delivered? (main steps)' },
      { t: 'textarea', p: 'narrative.market', label: 'Who buys from you, and how do you sell?' },
      { t: 'textarea', p: 'narrative.swot', label: 'Your strengths, and the main risks you see' }] },
    { id: 'review', sec: 'Done', title: 'Your file is ready', custom: review }
  ];
  const visible = () => CARDS.filter(c => !c.when || c.when());
  const val = (x) => typeof x === 'function' ? x() : x;

  // ---------------- rendering ----------------
  function render() {
    const cards = visible(); if (idx >= cards.length) idx = cards.length - 1; if (idx < 0) idx = 0;
    const c = cards[idx];
    const secs = [...new Set(cards.map(x => x.sec))];
    $('#progress').innerHTML = `<div class="bar"><i style="width:${Math.round(idx / (cards.length - 1) * 100)}%"></i></div><div class="sec">${esc(c.sec)} · ${idx + 1} of ${cards.length}</div>`;
    const body = $('#card');
    body.innerHTML = `<h2>${esc(val(c.title))}</h2>${c.help ? `<p class="help">${esc(val(c.help))}</p>` : ''}<div id="fields"></div>`;
    const box = $('#fields');
    if (c.custom) c.custom(box);
    else {
      (c.fields || []).forEach(f => { if (!f.when || f.when()) box.appendChild(field(f)); });
      if (c.ai && aiOK) box.prepend(aiBox(c));
    }
    $('#back').style.visibility = idx === 0 ? 'hidden' : 'visible';
    $('#next').style.display = c.id === 'review' || c.id === 'welcome' ? 'none' : '';
    $('#next').textContent = 'Next';
    window.scrollTo(0, 0);
  }
  function wrapLabel(label, el, extra) {
    const d = document.createElement('label'); d.className = 'f';
    if (label) d.innerHTML = `<span class="lbl">${esc(val(label))}</span>`;
    d.appendChild(el); if (extra) d.insertAdjacentHTML('beforeend', extra); return d;
  }
  function markPending(el, p) { if (pending.has(p)) el.classList.add('pending'); }
  function field(f) {
    const p = f.p;
    if (f.t === 'info') { const d = document.createElement('div'); d.className = 'info'; d.dataset.info = '1'; d._fn = f.html; d.innerHTML = f.html() || ''; return d; }
    if (f.t === 'text' || f.t === 'year' || f.t === 'num' || f.t === 'pct' || f.t === 'date') {
      const i = document.createElement('input');
      i.type = f.t === 'date' ? 'date' : 'text'; if (f.t !== 'text' && f.t !== 'date') i.inputMode = 'decimal';
      const v = get(p); i.value = f.t === 'pct' ? pctShow(v) : (v == null ? '' : v); if (f.ph) i.placeholder = f.ph;
      if (f.t === 'year') i.placeholder = 'e.g. 2026';
      i.oninput = () => { set(p, f.t === 'text' || f.t === 'date' ? i.value : f.t === 'pct' ? (num(i.value) == null ? null : num(i.value) / 100) : num(i.value)); pending.delete(p); i.classList.remove('pending'); save(); refreshInfo(); };
      markPending(i, p);
      const unit = f.t === 'pct' ? '%' : f.unit; return wrapLabel(f.label, unit ? withUnit(i, unit) : i);
    }
    if (f.t === 'textarea') { const t = document.createElement('textarea'); t.rows = 3; t.value = get(p) || ''; t.oninput = () => { set(p, t.value); save(); }; return wrapLabel(f.label, t); }
    if (f.t === 'choice') {
      const d = document.createElement('div'); d.className = 'choices';
      val(f.opts).forEach(([v, l]) => { const b = document.createElement('button'); b.type = 'button'; b.textContent = l; if (get(p) === v) b.className = 'on'; b.onclick = () => { set(p, v); pending.delete(p); save(); render(); }; d.appendChild(b); });
      markPending(d, p); return wrapLabel(f.label, d);
    }
    if (f.t === 'list') return listField(f);
    if (f.t === 'series') return seriesField(f);
    if (f.t === 'hist') return histField(f);
    return document.createTextNode('');
  }
  function withUnit(i, unit) { const s = document.createElement('span'); s.className = 'unit'; s.appendChild(i); s.insertAdjacentHTML('beforeend', `<em>${esc(unit)}</em>`); return s; }
  function listField(f) {
    let arr = get(f.p); if (!Array.isArray(arr)) { arr = []; set(f.p, arr); }
    if (!arr.length && f.min !== 0) arr.push({});
    const wrap = document.createElement('div'); wrap.className = 'list';
    arr.forEach((it, n) => {
      const card = document.createElement('div'); card.className = 'item';
      card.innerHTML = `<div class="ihead">#${n + 1}<button type="button" class="x" title="Remove">Remove</button></div>`;
      card.querySelector('.x').onclick = () => { arr.splice(n, 1); save(); render(); };
      f.item.forEach(sf => {
        const path = `${f.p}.${n}.${sf.k}`;
        if (sf.t === 'select') {
          const s = document.createElement('select'); const opts = val(sf.opts);
          s.innerHTML = `<option value="">Choose…</option>` + opts.map(([v, l]) => `<option value="${esc(v)}">${esc(l)}</option>`).join('');
          let cur = it[sf.k]; if (cur == null && sf.defFrom === 'kind' && it.kind) cur = M.KINDS[it.kind].cls;
          if (cur == null && sf.k === 'loan') cur = 'Y';
          if (cur == null && opts.length && sf.k !== 'kind' && sf.k !== 'cls') cur = opts[0][0];
          if (cur != null) { s.value = cur; it[sf.k] = cur; }
          s.onchange = () => {
            it[sf.k] = s.value || null; pending.delete(path);
            if (sf.k === 'kind' && s.value) { const k = M.KINDS[s.value]; it.cls = k.cls; if (it.gst == null) it.gst = k.gst; if (!it.desc) it.desc = k.label; }
            save(); render();
          };
          markPending(s, path);
          const l = wrapLabel(sf.label, s); if (sf.adv) l.classList.add('adv'); card.appendChild(l);
        } else card.appendChild(field({ ...sf, p: path }));
      });
      if (f.p === 'facilities.proposedTL' && it.amount && it.tenor) {
        const e = M.emi(it.amount, it.rate || 0, it.tenor); const k = M.activeMonthsY1(it.start, A.baseFY);
        card.insertAdjacentHTML('beforeend', `<div class="calc">${it.basis === 'Principal Run-Down' ? 'Monthly principal' : 'EMI'} ≈ ₹${f2(it.basis === 'Principal Run-Down' ? it.amount / it.tenor : e)} lakh/month${k != null ? ` · active ${k} months in FY${String(A.baseFY + 1).slice(2)}` : ''} (worked out from your answers)</div>`);
      }
      wrap.appendChild(card);
    });
    const add = document.createElement('button'); add.type = 'button'; add.className = 'add'; add.textContent = '+ ' + (f.add || 'Add');
    add.onclick = () => { arr.push({}); save(); render(); };
    wrap.appendChild(add);
    return wrap;
  }
  function seriesField(f) {
    const n = yearsN() || 1; let arr = get(f.p); if (!Array.isArray(arr)) { arr = []; set(f.p, arr); }
    const d = document.createElement('div'); d.className = 'series';
    const b = f.bm ? (M.benchmarks(A) || {})[f.bm] : null;
    const bmTxt = b != null && isFinite(b) ? `<button type="button" class="mini usebm">Use my 3-yr average: ${f.kind === 'pct' ? f2(b * 100) + '%' : f2(b)}</button>` : '';
    d.innerHTML = `<div class="lbl">${esc(val(f.label))}${f.opt ? ' <small>(optional)</small>' : ''}</div><div class="yrs"></div><div class="tools"><button type="button" class="mini same">Same for all years</button>${f.grow ? '<button type="button" class="mini grow">Grow each year by…</button>' : ''}${bmTxt}</div>`;
    const yrs = d.querySelector('.yrs');
    const inputs = [];
    for (let j = 0; j < n; j++) {
      const i = document.createElement('input'); i.inputMode = 'decimal'; const v = arr[j];
      i.value = f.kind === 'pct' ? pctShow(v) : (v == null ? '' : v);
      const pth = `${f.p}.${j}`; markPending(i, pth); markPending(i, f.p);
      i.oninput = () => { const x = num(i.value); arr[j] = x == null ? null : (f.kind === 'pct' ? x / 100 : x); pending.delete(pth); pending.delete(f.p); i.classList.remove('pending'); save(); refreshInfo(); };
      const cell = document.createElement('label'); cell.className = 'yr'; cell.innerHTML = `<span>${A.baseFY ? FY(j) : 'Year ' + (j + 1)}</span>`;
      cell.appendChild(f.kind === 'pct' ? withUnit(i, '%') : i); yrs.appendChild(cell); inputs.push(i);
    }
    d.querySelector('.same').onclick = () => { const v = arr[0]; if (v == null) return alert('Fill the first year first.'); for (let j = 1; j < n; j++) arr[j] = v; save(); render(); };
    if (f.grow) d.querySelector('.grow').onclick = () => { const v = arr[0]; if (v == null) return alert('Fill the first year first.'); const g = num(prompt('Grow by what % each year?', '')); if (g == null) return; for (let j = 1; j < n; j++) arr[j] = Math.round(v * Math.pow(1 + g / 100, j) * 100) / 100; save(); render(); };
    if (bmTxt) d.querySelector('.usebm').onclick = () => { const v = f.kind === 'pct' ? Math.round(b * 10000) / 10000 : Math.round(b); for (let j = 0; j < n; j++) arr[j] = v; save(); render(); };
    if (!A.baseFY) d.insertAdjacentHTML('beforeend', '<p class="warn">Tell us the latest financial year first (Your business section).</p>');
    return d;
  }
  function histField(f) {
    const d = document.createElement('div'); d.className = 'hist';
    const yrs = [A.baseFY - 2, A.baseFY - 1, A.baseFY].map(M.fyLabel);
    let h = `<table><thead><tr><th></th>${yrs.map(y => `<th>${y}</th>`).join('')}</tr></thead><tbody>`;
    if (f.cls) h += `<tr><td>Audited or provisional?</td>${[0, 1, 2].map(j => `<td><select data-cls="${j}"><option value=""></option><option${(A.histClass || [])[j] === 'Audited' ? ' selected' : ''}>Audited</option><option${(A.histClass || [])[j] === 'Provisional' ? ' selected' : ''}>Provisional</option></select></td>`).join('')}</tr>`;
    f.rows.forEach(r => {
      const label = r === 14 ? rmWord() : r === 13 ? 'Imported raw material used' : M.HIST_LABEL[r];
      if (r === 13 && A.activity === 'trading') return;
      h += `<tr><td>${esc(label)}</td>${[0, 1, 2].map(j => { const v = ((A.hist || {})[r] || [])[j]; return `<td><input inputmode="decimal" data-r="${r}" data-j="${j}" value="${v == null ? '' : v}"></td>`; }).join('')}</tr>`;
    });
    d.innerHTML = h + '</tbody></table><p class="help small">₹ lakh. Use a minus sign for an expense in "other income".</p>';
    d.querySelectorAll('input').forEach(i => i.oninput = () => { const r = i.dataset.r, j = +i.dataset.j; A.hist[r] = A.hist[r] || [null, null, null]; A.hist[r][j] = num(i.value); save(); refreshInfo(); });
    d.querySelectorAll('select').forEach(s => s.onchange = () => { A.histClass = A.histClass || []; A.histClass[+s.dataset.cls] = s.value; save(); });
    return d;
  }
  function refreshInfo() { document.querySelectorAll('[data-info]').forEach(d => d.innerHTML = d._fn() || ''); }
  // ---------- live info panels (computed only from the user's answers) ----------
  function tieInfo() {
    const h = M.histCalc(A); return '<b>Check:</b> ' + h.map((y, j) => `${M.fyLabel(A.baseFY - 2 + j)}: assets ${f2(y.assets)} vs liabilities ${f2(y.liab)} ${Math.abs(y.assets - y.liab) < 0.05 ? '✓' : '<span class="bad">≠</span>'}`).join(' · ');
  }
  function projCost() { return (A.project.items || []).reduce((s, i) => s + (i.cost || 0) * (1 + (i.gst || 0)), 0); }
  function costInfo() { const c = projCost(); return c ? `Total project cost incl. GST: <b>₹${f2(c)} lakh</b>` : ''; }
  function fundInfo() {
    const c = projCost(), tl = (A.facilities.proposedTL || []).reduce((s, t) => s + (t.amount || 0), 0), own = (A.funds.promoterCapital || 0) + (A.funds.unsecuredLoan || 0) + (A.funds.subsidy || 0);
    if (!c) return ''; return `Project cost ₹${f2(c)} L · your money + subsidy ₹${f2(own)} L${tl ? ` · term loan ₹${f2(tl)} L · balance ₹${f2(c - own - tl)} L` : ''}`;
  }
  function tlInfo() { if (!A.baseFY) return ''; const h = M.horizon(A); return `Last term-loan instalment falls in <b>FY${String(h).slice(2)}</b>, so the report will project ${M.projYears(A)} years (FY${String(A.baseFY + 1).slice(2)}–FY${String(h).slice(2)}).`; }
  function salesInfo() { const s = M.projSales(A); return s.some(isFinite) ? 'Sales this gives (₹ lakh): ' + s.map((v, j) => `${FY(j)} ${f2(v)}`).join(' · ') : ''; }

  // ---------- custom cards ----------
  function welcome(box) {
    box.innerHTML = `<p>Answer simple questions about your business. At the end you get a bank-ready <b>Detailed Project Report + CMA</b> in Excel, with every formula working.</p>
    <ul class="ticks"><li>Your answers stay on this device. Nothing is uploaded when you make the Excel.</li><li>We never fill in numbers for you. Anything you skip is listed as <i>"collect from customer"</i>.</li><li>Takes about 20–30 minutes. You can stop and come back; answers are saved in this browser.</li></ul>
    <div class="row"><button class="primary" id="go">${Object.keys(A.business || {}).length > 2 ? 'Continue' : 'Start'}</button><button id="demo" class="ghost">Try the fictional demo case</button><button id="clear" class="ghost">Clear my answers</button></div>
    <p class="help small" id="aistat"></p>`;
    $('#go').onclick = () => { idx = 1; render(); };
    $('#demo').onclick = async () => { if (!confirm('Load the fictional "Apex Precision Components" practice case? This replaces your answers on this device.')) return; A = norm(await (await fetch('demo_apex.json')).json()); save(); idx = visible().length - 1; render(); };
    $('#clear').onclick = () => { if (confirm('Delete all answers saved on this device?')) { localStorage.removeItem(STORE); A = fresh(); render(); } };
    $('#aistat').textContent = aiOK ? 'Optional AI help is available (you will be asked before anything is sent).' : 'Working without AI help — everything still works.';
  }
  function histHow(box) {
    box.innerHTML = `<p>You can type the numbers on the next screens, or fill our one-page sheet in Excel and upload it.</p>
    <div class="row"><a class="ghost btn" href="tab4_upload_template.xlsx" download>Download blank sheet</a><label class="ghost btn">Upload filled sheet<input type="file" id="up" accept=".xlsx,.csv" hidden></label></div><p id="upmsg" class="help"></p>
    <p class="help small">The upload is read inside your browser; the file is not sent anywhere.</p>`;
    $('#up').onchange = async e => {
      const file = e.target.files[0]; if (!file) return; let cells = {};
      try {
        if (/\.csv$/i.test(file.name)) { (await file.text()).split(/\r?\n/).forEach((line, r) => line.split(',').forEach((v, c) => { cells[X.colName(c + 1) + (r + 1)] = isNaN(Number(v)) || v === '' ? v : Number(v); })); }
        else cells = await X.readFirstSheet(JSZip, await file.arrayBuffer());
        let n = 0; const rowsSeen = [];
        Object.entries(cells).forEach(([ref, v]) => { const m = /^A(\d+)$/.exec(ref); if (m && /^R\d+$/.test(String(v).trim())) rowsSeen.push([String(v).trim().slice(1), m[1]]); });
        rowsSeen.forEach(([r, line]) => { if (!M.HIST_ROWS.includes(+r)) return; A.hist[r] = ['C', 'D', 'E'].map(c => { const v = cells[c + line]; if (typeof v === 'number') { n++; return v; } return null; }); });
        save(); $('#upmsg').textContent = `Read ${n} numbers from ${file.name}. Please check them on the next screens.`;
      } catch (err) { $('#upmsg').textContent = 'Could not read that file: ' + err.message; }
    };
  }
  function settingsCard(box) {
    const td = M.taxDefault((A.business || {}).constitution); const S = A.settings;
    const D = M.DEFAULTS;
    const row = (key, label, kind, why) => { const v = S[key] != null ? S[key] : key === 'taxRate' ? td.v : (D[key] || {}).v; return `<label class="f"><span class="lbl">${label}</span>${kind === 'pct' ? `<span class="unit"><input data-k="${key}" data-kind="pct" inputmode="decimal" value="${pctShow(v)}"><em>%</em></span>` : `<input data-k="${key}" value="${esc(v)}">`}<small class="why">${esc(why)}</small></label>`; };
    box.innerHTML = row('taxRate', 'Income-tax rate used in projections', 'pct', td.why) +
      `<details><summary>More bank settings (template defaults)</summary>
      ${row('mpbf', 'MPBF method (1 or 2)', 'txt', D.mpbf.why)}${row('tlInCL', 'Show next-year TL instalments as current liability (Y/N)', 'txt', D.tlInCL.why)}${row('exclLC', 'Exclude usance LC creditors (Y/N)', 'txt', D.exclLC.why)}
      ${row('mDom', 'Bank margin on domestic debtors', 'pct', D.mDom.why)}${row('mExp', 'Bank margin on export debtors', 'pct', D.mExp.why)}${row('mRM', 'Bank margin on raw-material stock', 'pct', D.mRM.why)}
      ${['14', '15', '16', '17', '18'].map(r => row('bep' + r, 'Break-even: variable share, row ' + r, 'pct', D['bep' + r].why)).join('')}
      ${row('shock5', 'Stress test: sales drop', 'pct', D.shock5.why)}${row('shock6', 'Stress test: raw-material cost rise', 'pct', D.shock6.why)}${row('shock7', 'Stress test: other expenses rise', 'pct', D.shock7.why)}
      <p class="help small">Depreciation rates (WDV): ${M.CLASSES.map((c, i) => `${c} ${Math.round((S.rates || D.rates.v)[i] * 100)}%`).join(' · ')} — edit in the Excel (Tab 3, C22:C28) if your CA uses different rates.</p></details>`;
    box.querySelectorAll('input').forEach(i => i.oninput = () => { const k = i.dataset.k; let v = i.dataset.kind === 'pct' ? (num(i.value) == null ? null : num(i.value) / 100) : (k === 'mpbf' ? num(i.value) : i.value.toUpperCase()); S[k] = v; save(); });
  }
  function review(box) {
    const res = M.build(A), W = M.warnings(A); const c = projCost(); const s = M.projSales(A);
    box.innerHTML = `<div class="sum"><div><span>Project cost</span><b>₹${f2(c)} L</b></div><div><span>Term loan asked</span><b>₹${f2((A.facilities.proposedTL || []).reduce((t, x) => t + (x.amount || 0), 0))} L</b></div><div><span>Your money</span><b>₹${f2((A.funds.promoterCapital || 0) + (A.funds.unsecuredLoan || 0))} L</b></div><div><span>Year-1 sales</span><b>₹${f2(s[0])} L</b></div></div>
    <p class="help small">These are simple totals of your answers. DSCR, balance sheet and MPBF are calculated by the Excel formulas when you open the file.</p>
    <h3>Points a banker will ask about (${W.length})</h3>${W.length ? '<ul class="warns">' + W.map(w => `<li>${esc(w.msg)}</li>`).join('') + '</ul>' : '<p>None found.</p>'}
    <h3>Still to collect (${res.unknown.length})</h3>${res.unknown.length ? '<details><summary>Show list</summary><ul class="unk">' + res.unknown.map(u => `<li>${esc(u.label)}</li>`).join('') + '</ul></details><p class="help small">These stay blank in the Excel and are listed on the "19_Unknown_Collect" sheet.</p>' : '<p>Nothing missing.</p>'}
    <div class="row"><button class="primary" id="dl">Download Excel (DPR + CMA)</button></div>
    <div id="aiwrap"></div>
    <p class="help small">Need to change something? Use Back, or tap a section: ${[...new Set(visible().map(x => x.sec))].slice(1, -1).map(sx => `<a href="#" data-sec="${esc(sx)}">${esc(sx)}</a>`).join(' · ')}</p>`;
    box.querySelectorAll('[data-sec]').forEach(a => a.onclick = e => { e.preventDefault(); idx = visible().findIndex(x => x.sec === a.dataset.sec); render(); });
    $('#dl').onclick = () => download(res, W);
    if (aiOK) {
      $('#aiwrap').innerHTML = `<div class="aibox"><h3>Optional: write the report text with AI</h3><p class="help small">Adds a "20_Report_Narrative" sheet (project, promoters, process, market, SWOT, conclusion). Every figure in it is a live link to the workbook; the AI is not allowed to write numbers, and any sentence with a number that is not yours is removed.</p>
      <label class="chk"><input type="checkbox" id="cons"> I agree to send my <b>text answers</b> (no figures) to DeepSeek, on servers outside India. Business and people names are replaced before sending. Don't include PAN/Aadhaar/account numbers.</label>
      <button id="narr" class="ghost">Write report text</button><div id="narrout"></div></div>`;
      $('#narr').onclick = writeNarrative;
    }
    if (narrative) showNarrative();
  }
  async function download(res, W) {
    const btn = $('#dl'); btn.disabled = true; btn.textContent = 'Preparing…';
    try {
      const tpl = await (await fetch('template_public.xlsx')).arrayBuffer();
      const extra = [M.profileSheet(A), M.unknownSheet(res, W)];
      const hasNotes = Object.values(A.narrative || {}).some(Boolean);
      if (narrative || hasNotes) extra.push(M.reportSheet(A, narrative));
      const blob = await X.fill(JSZip, tpl, res.cells, extra, 'blob');
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
      a.download = `DPR_CMA_${(A.business.name || 'draft').replace(/[^\w]+/g, '_').slice(0, 40)}.xlsx`; a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      btn.textContent = 'Downloaded ✓ (download again)';
    } catch (e) { alert('Could not build the file: ' + e.message); btn.textContent = 'Download Excel (DPR + CMA)'; }
    btn.disabled = false;
  }
  // ---------- AI (optional; via proxy, never required) ----------
  const names = () => [A.business.name, A.business.auditor, ...(A.business.promoters || []).map(p => p.name), ...(A.project.items || []).map(i => i.supplier)].filter(Boolean);
  async function callAI(path, body) {
    const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), 45000);
    try { const r = await fetch(AI_URL + path, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body), signal: ctl.signal }); const d = await r.json(); if (!r.ok) throw new Error(d.error || ('error ' + r.status)); return d; }
    finally { clearTimeout(t); }
  }
  function aiBox(c) {
    const d = document.createElement('div'); d.className = 'aibox';
    d.innerHTML = `<details><summary>✍️ Or just tell me in your own words</summary><textarea rows="3" placeholder="e.g. We are buying two CNC machines from LMW for 1.1 crore plus 18% GST…"></textarea>
    <label class="chk small"><input type="checkbox" class="cons" ${aiConsent ? 'checked' : ''}> Send this text to DeepSeek (servers outside India) to fill the boxes. Names, PAN, Aadhaar, account numbers, phones and emails are removed first. You check every value before it is used.</label>
    <button type="button" class="ghost go">Fill the boxes from my words</button><p class="aimsg help small"></p></details>`;
    d.querySelector('.go').onclick = async () => {
      if (!d.querySelector('.cons').checked) return alert('Please tick the box to agree first, or fill the boxes yourself.');
      aiConsent = true; localStorage.setItem(STORE + '-consent', 'yes');
      const raw = d.querySelector('textarea').value; const text = M.redact(raw, names());
      const msg = d.querySelector('.aimsg'); msg.textContent = 'Reading…';
      const fields = []; (c.fields || []).forEach(f => {
        if (f.when && !f.when()) return;
        if (['text', 'num', 'pct', 'year', 'date', 'textarea'].includes(f.t)) fields.push({ key: f.p, label: val(f.label) || f.p, type: f.t === 'pct' ? 'percent' : f.t === 'num' ? 'number' : f.t });
        if (f.t === 'choice') fields.push({ key: f.p, label: val(f.label) || 'choice', type: 'choice', options: val(f.opts).map(o => o[0]) });
        if (f.t === 'list') fields.push({ key: f.p, label: 'list', type: 'list', sub: f.item.map(s => s.k + (s.t === 'pct' ? ':percent' : s.t === 'num' ? ':number' : s.t === 'date' ? ':date' : '')) });
        if (f.t === 'series') fields.push({ key: f.p, label: (val(f.label) || '') + ` - one value per year for ${yearsN() || 1} years (${f.kind === 'pct' ? 'fractions' : 'numbers'}); repeat a value only if the user says it stays the same`, type: 'list' });
      });
      try {
        const out = await callAI('/extract', { card: c.id, context: `${A.activity || ''} ${A.stage || ''} base FY ${A.baseFY || ''}`, fields, text });
        const found = []; const notLiteral = [];
        const literal = new Set((raw.match(/\d[\d,.]*/g) || []).map(s => Number(s.replace(/,/g, ''))));
        const chk = v => { if (typeof v === 'number' && v !== 0 && !literal.has(v) && !literal.has(Math.round(v * 100 * 1e6) / 1e6)) notLiteral.push(v); };
        for (const [p, v] of Object.entries(out.values || {})) {
          if (Array.isArray(v) && Array.isArray(get(p) || [])) {
            if (v.length && typeof v[0] === 'object') { const arr = get(p) || []; const startN = arr.filter(x => x && Object.keys(x).length).length; const base = arr.slice(0, startN); v.forEach((o, n) => { base.push(o); Object.entries(o).forEach(([k, x]) => { chk(x); pending.add(`${p}.${startN + n}.${k}`); }); }); set(p, base); }
            else { set(p, v.map(x => typeof x === 'number' ? x : num(x))); v.forEach((x, j) => { chk(x); pending.add(`${p}.${j}`); }); }
          } else { set(p, v); chk(v); pending.add(p); }
          found.push(p);
        }
        save(); render();
        const m2 = document.querySelector('.aimsg');
        if (m2) m2.innerHTML = found.length ? `Filled ${found.length} item(s) — <b>highlighted in yellow: please check each one</b>.${notLiteral.length ? ' Some numbers were converted (e.g. crore → lakh): ' + notLiteral.map(f2).join(', ') + '.' : ''}${out.followup ? '<br>Next: ' + esc(out.followup) : ''}` : 'I could not find answers in that text. Please fill the boxes.';
        const det = document.querySelector('.aibox details'); if (det) det.open = true;
      } catch (e) { msg.textContent = 'AI help is not available right now (' + e.message + '). Please fill the boxes yourself.'; }
    };
    return d;
  }
  async function writeNarrative() {
    if (!$('#cons').checked) return alert('Please tick the consent box first.');
    const out = $('#narrout'); out.textContent = 'Writing…';
    const B = A.business, nm = names();
    const facts = {
      constitution: B.constitution, activity_type: A.activity, industry: B.industry, activity: B.activityDesc, location: [B.area, B.district, B.state].filter(Boolean).join(', '),
      stage: A.stage === 'new' ? 'New unit' : 'Expansion of a running business',
      project_items: (A.project.items || []).map(i => `${(M.KINDS[i.kind] || {}).label || ''}: ${i.desc || ''}`).join('; '),
      promoters: (B.promoters || []).map((p, n) => `Promoter ${n + 1}: ${p.experienceLine ? 'experience in ' + p.experienceLine : ''}${p.years ? ' (' + p.years + ' years)' : ''}${p.qualification ? '; ' + p.qualification : ''}`).join(' | '),
      why_project: A.narrative.project, process: A.narrative.process, market: A.narrative.market, strengths_and_risks: A.narrative.swot,
      job_work_principals: (A.jobwork || {}).principals, order_book_customers: (A.orderBook || {}).customers,
      collateral_types: (A.collateral || []).map(c => c.type).filter(Boolean).join(', ')
    };
    Object.keys(facts).forEach(k => { if (!facts[k]) delete facts[k]; else facts[k] = M.redact(String(facts[k]), nm); });
    try {
      const r = await callAI('/narrative', { facts });
      const chk = M.checkNarrative(r.sections || [], Object.values(facts).join(' '));
      narrative = { sections: chk.sections, dropped: chk.dropped }; showNarrative();
    } catch (e) { out.textContent = 'AI report text is not available right now (' + e.message + '). Your Excel works without it.'; }
  }
  function showNarrative() {
    const out = $('#narrout'); if (!out || !narrative) return;
    out.innerHTML = narrative.sections.map(s => `<h4>${esc(s.title)}</h4><p>${esc(s.text).replace(/\{\{([A-Z0-9_]+)\}\}/g, (m, k) => `<span class="ph">${k.toLowerCase().replace(/_/g, ' ')}</span>`).replace(/\n/g, '<br>')}</p>`).join('') +
      (narrative.dropped.length ? `<p class="help small">${narrative.dropped.length} sentence(s) removed because they contained numbers not in your answers.</p>` : '') +
      '<p class="help small">Grey chips are filled from the Excel formulas. Download the Excel again to include this text.</p>';
  }
  // ---------- navigation ----------
  function validate() {
    const c = visible()[idx];
    for (const f of c.fields || []) if (f.required && (get(f.p) == null || get(f.p) === '') && (!f.when || f.when())) { alert('Please answer: ' + val(f.label)); return false; }
    if (c.id === 'basefy' && A.baseFY && (A.baseFY < 2000 || A.baseFY > 2100)) { alert('Please enter a 4-digit year, e.g. 2026.'); return false; }
    const pend = [...pending].filter(p => (c.fields || []).some(f => p.startsWith(f.p || '~')));
    if (pend.length && !confirm('Some values were filled from your words (yellow). Have you checked them?')) return false;
    pend.forEach(p => pending.delete(p));
    return true;
  }
  $('#next').onclick = () => { if (!validate()) return; idx++; render(); };
  $('#back').onclick = () => { idx--; render(); };
  // AI health check (short timeout; the app never waits on it)
  (async () => {
    try { const ctl = new AbortController(); setTimeout(() => ctl.abort(), 4000); const r = await fetch(AI_URL + '/health', { signal: ctl.signal }); const d = await r.json(); aiOK = !!d.ai; } catch { aiOK = false; }
    if (aiOK) render();
  })();
  render();
})();
