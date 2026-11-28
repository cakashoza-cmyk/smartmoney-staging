/* model.js — turns plain answers into workbook cells, unknown list, warnings and extra sheets.
   No numbers are invented here: every value comes from an answer, a documented default the user saw,
   or a deterministic calculation on the user's own answers (shown to them). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DPRModel = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const S3 = '3_Project_&_Debt_Setup', S4 = '4_Historical_&_Benchmarks', S5 = '5_Projection_Assumptions', S17 = '17_Sensitivity_Analysis';
  const P3 = "'3_Project_&_Debt_Setup'";
  const YC = 'DEFGHIJKLM'.split('');
  const CLASSES = ['Plant & Machinery', 'Construction & Elec.', 'Misc. Fixed Assets', 'Computers & Software', 'Furniture & Fixtures', 'Vehicles & Transport', 'Factory Building / Shed'];
  const KINDS = {
    land: { label: 'Land', cls: 'Misc. Fixed Assets', gst: 0 },
    building: { label: 'Building / shed', cls: 'Factory Building / Shed', gst: 0.18 },
    machinery: { label: 'Machinery', cls: 'Plant & Machinery', gst: 0.18 },
    electrification: { label: 'Electrification', cls: 'Construction & Elec.', gst: 0.18 },
    computers: { label: 'Computers / software', cls: 'Computers & Software', gst: 0.18 },
    furniture: { label: 'Furniture & fixtures', cls: 'Furniture & Fixtures', gst: 0.18 },
    vehicle: { label: 'Vehicle', cls: 'Vehicles & Transport', gst: 0.28 },
    preop: { label: 'Pre-operative expenses', cls: 'Misc. Fixed Assets', gst: 0 },
    contingency: { label: 'Contingency', cls: 'Misc. Fixed Assets', gst: 0 },
    other: { label: 'Other', cls: 'Plant & Machinery', gst: 0.18 }
  };
  // Documented defaults — always shown to the user and editable.
  const DEFAULTS = {
    mpbf: { v: 1, why: 'Tandon Method 1 (template default). Use 2 if your bank asks for Method 2.' },
    tlInCL: { v: 'Y', why: 'Next-12-month term loan instalments shown as current liabilities (template default).' },
    exclLC: { v: 'Y', why: 'Usance LC creditors excluded from trade creditors (template default).' },
    mExp: { v: 0.25, why: 'Bank margin on export receivables (template default).' },
    mDom: { v: 0.30, why: 'Bank margin on domestic receivables (template default).' },
    mRM: { v: 0.25, why: 'Bank margin on raw-material stock (template default).' },
    bep14: { v: 1, why: 'Raw material treated as 100% variable for break-even (template default).' },
    bep15: { v: 0.95, why: 'Power, labour & mfg. expenses 95% variable (template default).' },
    bep16: { v: 0.95, why: 'SG&A 95% variable (template default).' },
    bep17: { v: 0.95, why: 'WC interest & LC/BG commission 95% variable (template default).' },
    bep18: { v: 0, why: 'Term-loan interest treated as fixed (template default).' },
    shock5: { v: 0.10, why: 'Stress test: sales drop (template default).' },
    shock6: { v: 0.05, why: 'Stress test: raw-material cost rise (template default).' },
    shock7: { v: 0.10, why: 'Stress test: other operating expenses rise (template default).' },
    rates: { v: [0.15, 0.35, 0, 0.40, 0.10, 0.15, 0.10], why: 'WDV rates per Income-tax Act blocks as set in the template.' }
  };
  // Tax default by constitution (statutory; editable). Proprietor: slab rates -> no default.
  function taxDefault(constitution) {
    if (/Private Limited|Public Limited/i.test(constitution || '')) return { v: 0.2517, why: 'Company under s.115BAA: 22% + 10% surcharge + 4% cess = 25.17%. Change if the company has not opted for it.' };
    if (/Partnership|LLP/i.test(constitution || '')) return { v: 0.312, why: 'Firm / LLP: 30% + 4% cess = 31.2% (34.94% if income exceeds Rs 1 crore, due to 12% surcharge).' };
    return { v: null, why: 'Proprietor: tax depends on slab rates, so please enter the effective rate your CA uses.' };
  }
  const HIST_ROWS = [8, 9, 11, 13, 14, 15, 16, 17, 18, 20, 21, 23, 25, 26, 27, 29, 33, 34, 35, 36, 37, 39, 40, 41, 42, 44, 45, 46, 47, 48, 50, 51, 53];
  const HIST_LABEL = {
    8: 'Sales in India', 9: 'Export sales', 11: 'Other operating income', 13: 'Imported raw material used', 14: 'Raw material used (Indian)',
    15: 'Power & fuel', 16: 'Wages & salaries (factory)', 17: 'Other factory expenses', 18: 'Depreciation', 20: 'Opening stock of finished goods',
    21: 'Closing stock of finished goods', 23: 'Office, selling & admin expenses', 25: 'Interest on CC/OD & bank charges', 26: 'Interest on term loans',
    27: 'Other income / (expense), net', 29: 'Income tax provided', 33: 'Bank CC / OD outstanding', 34: 'Creditors for purchases', 35: 'Provision for tax',
    36: 'Other current liabilities', 37: 'Term-loan instalments due within 12 months', 39: 'Term loans (beyond 12 months)', 40: 'Unsecured loans (promoters/relatives)',
    41: 'Capital', 42: 'Reserves & surplus / accumulated profit', 44: 'Cash & bank', 45: 'Debtors', 46: 'Raw-material stock', 47: 'Finished-goods stock',
    48: 'Advance tax & other current assets', 50: 'Gross block (fixed assets at cost)', 51: 'Accumulated depreciation', 53: 'Deposits, BG margin & other non-current assets'
  };
  const fyLabel = y => 'FY' + String(y).slice(2);
  const isNum = v => typeof v === 'number' && isFinite(v);
  const blank = v => v === null || v === undefined || v === '' || (typeof v === 'number' && !isFinite(v));
  function eomonth(iso, add) { const [y, m] = iso.split('-').map(Number); const d = new Date(Date.UTC(y, m - 1 + add + 1, 0)); return d; }
  function fyOf(d) { return d.getUTCMonth() + 1 <= 3 ? d.getUTCFullYear() : d.getUTCFullYear() + 1; }
  function activeMonthsY1(startIso, baseFY) {
    if (!startIso || !baseFY) return null;
    const [y, m] = startIso.split('-').map(Number);
    const idx = (y - baseFY) * 12 + (m - 4); // months after 1 Apr of FY1
    if (idx <= 0) return 12; if (idx >= 12) return 0; return 12 - idx;
  }
  function emi(P, r, n) { if (!(P > 0) || !(n > 0)) return 0; if (!r) return P / n; const i = r / 12; return P * i * Math.pow(1 + i, n) / (Math.pow(1 + i, n) - 1); }
  function facilities(A) {
    const F = A.facilities || {}; const rows = [];
    for (const e of F.existing || []) rows.push({ kind: 'existingTL', type: e.type || 'Term Loan', name: e.name || 'Existing term loan', basis: e.basis || 'Principal Run-Down', amount: e.outstanding, rate: e.rate, start: A.baseFY ? `${A.baseFY}-04-01` : null, mor: 0, tenor: e.remainingMonths, k: 12 });
    for (const t of F.proposedTL || []) rows.push({ kind: 'newTL', type: t.type || 'Term Loan', name: t.name || 'Proposed term loan', basis: t.basis || 'Level EMI', amount: t.amount, rate: t.rate, start: t.start, mor: t.moratorium, tenor: t.tenor, k: t.k != null ? t.k : activeMonthsY1(t.start, A.baseFY) });
    for (const w of F.wc || []) rows.push({ kind: 'wc', type: w.type || 'Cash Credit', name: w.name || w.type || 'Cash credit', basis: 'WC Revolving', amount: w.amount, rate: w.rate, start: w.start, mor: 0, tenor: w.tenor != null ? w.tenor : 12, k: w.k != null ? w.k : activeMonthsY1(w.start, A.baseFY) });
    for (const n of F.nonfund || []) rows.push({ kind: 'nf', type: n.type || 'Bank Guarantee', name: n.name || n.type, basis: 'Non-Fund Commission', amount: n.amount, rate: n.rate, start: n.start, mor: 0, tenor: n.tenor, k: n.k != null ? n.k : activeMonthsY1(n.start, A.baseFY) });
    return rows;
  }
  function horizon(A) {
    let mx = (A.baseFY || 0) + 1;
    for (const f of facilities(A)) {
      if (!/Term Loan|WCTL/.test(f.type) || !(f.amount > 0) || !(f.tenor > 0) || !f.start) continue;
      mx = Math.max(mx, fyOf(eomonth(f.start, (f.mor || 0) + f.tenor - 1)));
    }
    return mx;
  }
  function projYears(A) { return Math.max(1, Math.min(10, horizon(A) - A.baseFY)); }
  const revenueDirect = a => a === 'trading' || a === 'contract' || a === 'service';

  function build(A) {
    const cells = { [S3]: {}, [S4]: {}, [S5]: {}, [S17]: {} };
    const unknown = []; const defaultsUsed = []; const notes = [];
    const put = (sheet, ref, v, label, opt) => {
      if (blank(v)) { if (!(opt && opt.optional)) unknown.push({ label, sheet, ref }); return; }
      cells[sheet][ref] = v;
    };
    const S = A.settings || {};
    const useDef = (key, val, label, ref, sheet) => {
      const d = DEFAULTS[key]; const v = S[key] != null ? S[key] : d.v;
      if (S[key] == null || S[key] === d.v) defaultsUsed.push({ label, value: v, why: d.why, ref: `${sheet || S3}!${ref}` });
      return v;
    };
    const B = A.business || {};
    // 1A
    put(S3, 'C5', B.name, 'Name of the business');
    put(S3, 'G5', B.constitution, 'Constitution (proprietorship / firm / company)');
    put(S3, 'C6', B.activityDesc || B.industry, 'What the business does');
    const loc = [B.area, B.district, B.state].filter(Boolean).join(', ');
    put(S3, 'G6', loc, 'Unit location (area, district)');
    put(S3, 'C7', B.auditor, 'Auditor / CA firm', { optional: A.stage === 'new' || A.history === 'none' });
    put(S3, 'G7', A.baseFY, 'Latest financial year (base FY)');
    cells[S3].C8 = useDef('mpbf', 0, 'MPBF method', 'C8');
    cells[S3].G8 = useDef('tlInCL', 0, 'Include TL instalments in current liabilities', 'G8');
    cells[S3].C9 = useDef('exclLC', 0, 'Exclude usance LC creditors', 'C9');
    const td = taxDefault(B.constitution);
    const tax = S.taxRate != null ? S.taxRate : td.v;
    if (S.taxRate == null && td.v != null) defaultsUsed.push({ label: 'Projected income-tax rate', value: td.v, why: td.why, ref: S3 + '!G9' });
    put(S3, 'G9', tax, 'Projected income-tax rate');
    cells[S3].C10 = useDef('mExp', 0, 'Margin on export receivables', 'C10');
    cells[S3].G10 = useDef('mDom', 0, 'Margin on domestic receivables', 'G10');
    cells[S3].C11 = useDef('mRM', 0, 'Margin on raw-material stock', 'C11');
    [14, 15, 16, 17, 18].forEach(r => { cells[S3]['B' + r] = useDef('bep' + r, 0, 'Break-even variable share row ' + r, 'B' + r); });
    // 1B existing assets
    const rates = S.rates || DEFAULTS.rates.v;
    if (!S.rates) defaultsUsed.push({ label: 'Depreciation rates (WDV)', value: rates.map(x => Math.round(x * 100) + '%').join(' / '), why: DEFAULTS.rates.why, ref: S3 + '!C22:C28' });
    CLASSES.forEach((c, i) => {
      const r = 22 + i; cells[S3]['A' + r] = c; cells[S3]['C' + r] = rates[i];
      const wdv = (A.existingAssets || {})[c];
      if (A.stage === 'new') cells[S3]['B' + r] = 0;
      else put(S3, 'B' + r, wdv, `Book value (WDV) of existing ${c}`, { optional: true });
    });
    if (A.stage !== 'new' && !Object.values(A.existingAssets || {}).some(isNum)) unknown.push({ label: 'Book value (WDV) of existing fixed assets by type', sheet: S3, ref: 'B22:B28' });
    // 1C capex
    const items = ((A.project || {}).items || []).filter(it => it && (it.desc || it.cost));
    if (items.length > 20) notes.push('Only the first 20 project items fit the template; the rest were not written.');
    if (!items.length) unknown.push({ label: 'Project items (what is being bought and its cost)', sheet: S3, ref: 'A33:J52' });
    items.slice(0, 20).forEach((it, i) => {
      const r = 33 + i, k = KINDS[it.kind] || KINDS.other, L = `Project item ${i + 1} (${it.desc || k.label})`;
      put(S3, 'A' + r, it.supplier, L + ' - supplier', { optional: true });
      put(S3, 'B' + r, it.desc || k.label, L + ' - description');
      put(S3, 'C' + r, it.cls || k.cls, L + ' - asset type');
      put(S3, 'D' + r, it.cost, L + ' - cost before GST (Rs lakh)');
      cells[S3]['E' + r] = ['INR', k.label, it.note].filter(Boolean).join(' | ');
      put(S3, 'F' + r, it.gst, L + ' - GST rate');
      put(S3, 'I' + r, it.months, L + ' - months in use in first projected year');
      put(S3, 'J' + r, it.loan || 'Y', L + ' - to be funded by bank loan?');
    });
    // 1D
    const Fd = A.funds || {};
    put(S3, 'C56', Fd.promoterCapital, 'Fresh capital the promoters are bringing in');
    put(S3, 'G56', Fd.unsecuredLoan, 'Fresh unsecured loans from promoters / relatives');
    // 1E
    const fac = facilities(A);
    if (fac.length > 10) notes.push('Only 10 credit facilities fit the template; the rest were not written.');
    fac.slice(0, 10).forEach((f, i) => {
      const r = 62 + i, L = `Facility ${i + 1} (${f.name})`;
      cells[S3]['A' + r] = f.type; cells[S3]['B' + r] = f.name; cells[S3]['D' + r] = f.basis;
      put(S3, 'E' + r, f.amount, L + (f.kind === 'existingTL' ? ' - amount outstanding' : ' - limit / amount'));
      put(S3, 'F' + r, f.rate, L + ' - interest / commission rate');
      put(S3, 'G' + r, f.start ? { date: f.start } : null, L + ' - start / disbursement date');
      cells[S3]['H' + r] = f.mor || 0;
      put(S3, 'I' + r, f.tenor, L + (f.kind === 'existingTL' ? ' - remaining months' : ' - repayment / tenor months'));
      put(S3, 'K' + r, f.k, L + ' - months active in first projected year');
    });
    // Tab 4
    if (A.stage === 'new' || A.history === 'none') {
      notes.push(A.stage === 'new' ? 'New unit: past-year accounts not applicable, Tab 4 left blank.' : 'No past accounts available: Tab 4 left blank.');
      if (A.stage !== 'new') unknown.push({ label: 'Past 3 years accounts (needed for an expansion)', sheet: S4, ref: 'B8:D53' });
    } else {
      const H = A.hist || {};
      const yrs = [A.baseFY - 2, A.baseFY - 1, A.baseFY].map(fyLabel);
      HIST_ROWS.forEach(r => {
        const arr = H[r] || [];
        ['B', 'C', 'D'].forEach((c, j) => put(S4, c + r, arr[j], `${HIST_LABEL[r]} (${yrs[j]})`, { optional: [9, 11, 13, 20, 21, 27, 36, 37, 46].includes(r) }));
      });
      ['B', 'C', 'D'].forEach((c, j) => { const v = (A.histClass || [])[j]; if (v) cells[S4][c + '6'] = v; });
    }
    // Tab 5
    const n = A.baseFY ? projYears(A) : 10; const O = A.ops || {};
    const series = (arr, row, label, opt) => {
      let last = null;
      YC.forEach((c, j) => {
        let v = arr ? arr[j] : null;
        if (j >= n) { if (blank(v)) v = last; if (!blank(v)) cells[S5][c + row] = v; return; }
        put(S5, c + row, v, `${label} (${A.baseFY ? fyLabel(A.baseFY + 1 + j) : 'year ' + (j + 1)})`, opt);
        if (!blank(v)) last = v;
      });
    };
    if (revenueDirect(A.activity)) {
      const plan = O.salesPlan || [];
      series(plan, 8, 'Expected sales (Rs lakh)');
      YC.forEach((c, j) => { if (!blank(cells[S5][c + '8'])) { cells[S5][c + '9'] = 1; cells[S5][c + '10'] = 100000; } });
      notes.push('Sales entered directly in Rs lakh: Tab 5 capacity row = sales in Rs lakh, utilisation 100%, price Rs 1,00,000 per Rs 1 lakh unit, so Tab 10 sales equal the sales you entered.');
    } else {
      series(O.capacity, 8, `Installed capacity (${O.unitName || 'units'} per year)`);
      series(O.util, 9, 'Capacity utilisation');
      series(O.price, 10, A.activity === 'jobwork' ? 'Job charge per unit (Rs)' : 'Average selling price per unit (Rs)');
    }
    series(O.exportShare, 11, 'Export share of sales', { optional: true });
    series(O.rmPct, 14, A.activity === 'trading' ? 'Purchase cost (% of sales)' : A.activity === 'jobwork' ? 'Consumables & stores (% of job income)' : 'Raw material (% of sales)');
    series(O.importShare, 15, 'Imported share of raw material', { optional: true });
    series(O.powerPct, 16, 'Power & fuel (% of sales)');
    series(O.heads, 17, 'Number of workers & staff');
    series(O.days, 18, 'Working days per year');
    series(O.wage, 19, 'Average daily wage per person (Rs)');
    series(O.otherMfgPct, 20, 'Other factory expenses (% of sales)');
    series(O.sgaPct, 21, 'Office, selling & admin (% of sales)');
    series(O.rmDays, 24, A.activity === 'trading' ? 'Stock holding (days)' : 'Raw-material stock (days)');
    series(O.fgDays, 25, 'Finished-goods stock (days)', { optional: A.activity === 'trading' || revenueDirect(A.activity) });
    series(O.debtorDays, 26, 'Days customers take to pay');
    series(O.creditorDays, 27, 'Days you take to pay suppliers');
    series(O.otherNCA, 28, 'Deposits / BG margin / other non-current assets (Rs lakh)', { optional: true });
    // 17 shocks
    cells[S17].B5 = useDef('shock5', 0, 'Stress: sales drop', 'B5', S17);
    cells[S17].B6 = useDef('shock6', 0, 'Stress: raw-material cost rise', 'B6', S17);
    cells[S17].B7 = useDef('shock7', 0, 'Stress: other expenses rise', 'B7', S17);
    // profile-sheet items (no blue cell in the template; collected on 18_Profile_&_Collateral)
    const PS = '18_Profile_&_Collateral';
    const prs = (B.promoters || []).filter(p => p && Object.values(p).some(v => !blank(v)));
    if (!prs.length) unknown.push({ label: 'Promoter names, experience and qualification', sheet: PS, ref: 'Section 1' });
    prs.forEach((p, i) => ['name', 'years', 'experienceLine', 'qualification'].forEach(k => { if (blank(p[k])) unknown.push({ label: `Promoter ${i + 1}: ${{ name: 'name', years: 'years of experience', experienceLine: 'line of experience', qualification: 'qualification' }[k]}`, sheet: PS, ref: 'Section 1' }); }));
    const cols = (A.collateral || []).filter(c => c && Object.values(c).some(v => !blank(v)));
    if (!cols.length) unknown.push({ label: 'Collateral / security offered (type, value, owner)', sheet: PS, ref: 'Section 5' });
    cols.forEach((c, i) => ['type', 'value', 'owner'].forEach(k => { if (blank(c[k])) unknown.push({ label: `Collateral ${i + 1}: ${k}`, sheet: PS, ref: 'Section 5' }); }));
    return { cells, unknown, defaultsUsed, notes, years: n, horizon: A.baseFY ? horizon(A) : null, facilities: fac };
  }

  // ---------- user-side figures (deterministic, from the user's own numbers) ----------
  function histCalc(A) {
    const H = A.hist || {}; const g = (r, j) => { const v = (H[r] || [])[j]; return isNum(v) ? v : 0; };
    const has = r => (H[r] || []).some(isNum);
    return [0, 1, 2].map(j => {
      const sales = g(8, j) + g(9, j), opInc = sales + g(11, j);
      const cop = [13, 14, 15, 16, 17, 18].reduce((s, r) => s + g(r, j), 0);
      const pbt = opInc - (cop + g(20, j) - g(21, j)) - g(23, j) - g(25, j) - g(26, j) + g(27, j);
      const pat = pbt - g(29, j);
      const liab = [33, 34, 35, 36, 37, 39, 40, 41, 42].reduce((s, r) => s + g(r, j), 0);
      const assets = [44, 45, 46, 47, 48].reduce((s, r) => s + g(r, j), 0) + g(50, j) - g(51, j) + g(53, j);
      return { sales, pbt, pat, liab, assets, rm: g(13, j) + g(14, j), power: g(15, j), sga: g(23, j), omfg: g(17, j), dep: g(18, j), accdep: g(51, j), res: g(42, j), cap: g(41, j), tax: g(29, j),
        debt: g(45, j), cred: g(34, j), rmInv: g(46, j), fgInv: g(47, j), hasBS: has(44) || has(45) || has(50) };
    });
  }
  const avg = a => { const b = a.filter(x => isFinite(x)); return b.length ? b.reduce((s, x) => s + x, 0) / b.length : NaN; };
  function benchmarks(A) {
    if (A.stage === 'new' || A.history === 'none') return null;
    const h = histCalc(A); const pct = (k) => avg(h.map(y => y.sales > 0 ? y[k] / y.sales : NaN));
    return {
      rmPct: pct('rm'), powerPct: pct('power'), sgaPct: pct('sga'), otherMfgPct: pct('omfg'),
      debtorDays: avg(h.map(y => y.sales > 0 ? y.debt / y.sales * 365 : NaN)),
      creditorDays: avg(h.map(y => y.rm > 0 ? y.cred / y.rm * 365 : NaN)),
      rmDays: avg(h.map(y => y.rm > 0 ? y.rmInv / y.rm * 365 : NaN)),
      taxRate: avg(h.map(y => y.pbt > 0 ? y.tax / y.pbt : NaN)), lastSales: h[2].sales, h
    };
  }
  function projSales(A) {
    const O = A.ops || {}; const n = A.baseFY ? projYears(A) : 10; const out = [];
    for (let j = 0; j < n; j++) {
      if (revenueDirect(A.activity)) out.push(isNum((O.salesPlan || [])[j]) ? O.salesPlan[j] : NaN);
      else { const c = (O.capacity || [])[j], u = (O.util || [])[j], p = (O.price || [])[j]; out.push(isNum(c) && isNum(u) && isNum(p) ? c * u * p / 1e5 : NaN); }
    }
    return out;
  }
  function warnings(A) {
    const W = []; const w = (msg, where) => W.push({ msg, where });
    const b = benchmarks(A); const O = A.ops || {}; const f1 = v => (Math.round(v * 10) / 10).toLocaleString('en-IN');
    const yr = j => A.baseFY ? fyLabel(A.baseFY - 2 + j) : 'year ' + (j + 1);
    if (b) {
      b.h.forEach((y, j) => { if (y.hasBS && Math.abs(y.assets - y.liab) > 0.05) w(`${yr(j)} balance sheet does not tie: assets ${f1(y.assets)} vs liabilities ${f1(y.liab)} (difference ${f1(y.assets - y.liab)} L).`, 'Past accounts'); });
      for (let j = 1; j < 3; j++) {
        const y = b.h[j], p = b.h[j - 1];
        const dRes = y.res - p.res;
        if (Math.abs(dRes - y.pat) > 1) w(`${yr(j)}: reserves rose by ${f1(dRes)} L but profit after tax works out to ${f1(y.pat)} L (gap ${f1(y.pat - dRes)} L). Ask about dividends, drawings or adjustments.`, 'Past accounts');
        const dAcc = y.accdep - p.accdep;
        if (y.dep > 0 && Math.abs(dAcc - y.dep) > 0.5) w(`${yr(j)}: depreciation in P&L is ${f1(y.dep)} L but accumulated depreciation rose by ${f1(dAcc)} L. Ask for the fixed-asset schedule (sale / write-off?).`, 'Past accounts');
      }
      const tax = A.settings && A.settings.taxRate != null ? A.settings.taxRate : taxDefault((A.business || {}).constitution).v;
      const lastTax = b.h[2].pbt > 0 ? b.h[2].tax / b.h[2].pbt : NaN;
      if (isFinite(lastTax) && tax != null && Math.abs(lastTax - tax) > 0.05) w(`Last year's tax was ${(lastTax * 100).toFixed(1)}% of profit before tax, but projections use ${(tax * 100).toFixed(2)}%. Confirm the rate with the CA.`, 'Tax');
      const cmp = (key, label, unit, tol, scale) => {
        const arr = O[key] || []; const bm = b[key]; if (!isFinite(bm)) return;
        const off = []; arr.slice(0, projYears(A)).forEach((v, j) => { if (isNum(v) && Math.abs(v - bm) > tol) off.push([fyLabel(A.baseFY + 1 + j), v]); });
        if (off.length) { const vals = [...new Set(off.map(o => f1(o[1] * scale)))]; w(`${label} is ${vals.join(' / ')}${unit} in ${off[0][0]}${off.length > 1 ? '-' + off[off.length - 1][0] : ''}, against your own 3-year average of ${f1(bm * scale)}${unit}.`, 'Projections'); }
      };
      cmp('rmPct', 'Raw material / purchase cost', '% of sales', 0.05, 100);
      cmp('powerPct', 'Power & fuel', '% of sales', 0.02, 100);
      cmp('sgaPct', 'Office, selling & admin', '% of sales', 0.03, 100);
      cmp('debtorDays', 'Debtor days', ' days', 15, 1);
      cmp('creditorDays', 'Creditor days', ' days', 15, 1);
      cmp('rmDays', 'Raw-material stock days', ' days', 15, 1);
    }
    const ps = projSales(A);
    if (b && isNum(ps[0]) && b.lastSales > 0 && Math.round((ps[0] / b.lastSales - 1) * 1000) > 300) w(`First projected year sales are ${f1(ps[0])} L, up ${f1((ps[0] / b.lastSales - 1) * 100)}% on last year's ${f1(b.lastSales)} L. Banks usually ask how this is reached (new capacity runs only part of the year?).`, 'Projections');
    for (let j = 1; j < ps.length; j++) if (isNum(ps[j]) && isNum(ps[j - 1]) && ps[j - 1] > 0 && Math.round((ps[j] / ps[j - 1] - 1) * 1000) > 300) w(`Sales jump ${f1((ps[j] / ps[j - 1] - 1) * 100)}% in ${fyLabel(A.baseFY + 1 + j)}.`, 'Projections');
    if (A.stage === 'new' && !revenueDirect(A.activity) && isNum((O.util || [])[0]) && O.util[0] > 0.70) w(`First-year utilisation of ${f1(O.util[0] * 100)}% is high for a new unit; banks often expect a ramp-up.`, 'Projections');
    // means of finance (deterministic from answers)
    const items = ((A.project || {}).items || []).filter(i => isNum(i.cost));
    const cost = items.reduce((s, i) => s + i.cost * (1 + (isNum(i.gst) ? i.gst : 0)), 0);
    const elig = items.filter(i => (i.loan || 'Y') === 'Y').reduce((s, i) => s + i.cost * (1 + (isNum(i.gst) ? i.gst : 0)), 0);
    const tl = ((A.facilities || {}).proposedTL || []).reduce((s, t) => s + (isNum(t.amount) ? t.amount : 0), 0);
    if (tl > 0 && elig > 0 && tl > elig + 0.01) w(`Proposed term loan ${f1(tl)} L is more than the loan-eligible project cost ${f1(elig)} L.`, 'Funding');
    const fd = A.funds || {}; const own = (fd.promoterCapital || 0) + (fd.unsecuredLoan || 0) + (fd.subsidy || 0);
    if (cost > 0 && tl + own < cost - 0.5) w(`Project cost ${f1(cost)} L, but term loan + your money + subsidy add up to ${f1(tl + own)} L. The gap of ${f1(cost - tl - own)} L must come from internal accruals; the model treats it as a promoter plug.`, 'Funding');
    if ((fd.subsidy || 0) > 0) w('Subsidy is recorded on the profile sheet only: Tab 6 "Subsidy / Capital Grant" is fixed at 0 in the template.', 'Funding');
    if ((A.project || {}).marginMoneyWC > 0) w('Margin money for working capital is recorded on the profile sheet; the template\'s project cost (Tab 6) covers fixed assets only.', 'Funding');
    return W;
  }

  // ---------- extra sheets ----------
  function profileSheet(A) {
    const B = A.business || {}; const rows = []; const R = (...c) => rows.push(c);
    R({ v: 'BORROWER PROFILE, PROMOTERS & COLLATERAL (entered in the DPR Assistant) [Rs. in Lakhs]', s: 'title' });
    R({ v: 'Separate input sheet added by the DPR Assistant. No existing formula reads from it; figures in Tabs 1-17 are unchanged.', s: 'wrap' });
    R();
    R({ v: '1. PROMOTERS', s: 'section' });
    R({ v: 'Name', s: 'label' }, { v: 'Role / % holding', s: 'label' }, { v: 'Experience (years)', s: 'label' }, { v: 'Line of experience', s: 'label' }, { v: 'Qualification', s: 'label' });
    const pr = (B.promoters || []).filter(p => p && (p.name || p.experienceLine));
    if (!pr.length) R({ v: 'Unknown - collect from customer', s: 'wrap' });
    pr.forEach(p => R(p.name || 'Unknown', p.role || '', isNum(p.years) ? p.years : 'Unknown', p.experienceLine || 'Unknown', p.qualification || 'Unknown'));
    R();
    R({ v: '2. BUSINESS', s: 'section' });
    const kv = (k, v) => R({ v: k, s: 'label' }, { v: blank(v) ? 'Unknown - collect from customer' : v, s: 'wrap' });
    kv('Industry', B.industry); kv('Activity type', { manufacturing: 'Manufacturing', jobwork: 'Job work', trading: 'Trading', contract: 'Contract-based', service: 'Service', mixed: 'Mixed (manufacturing + trading/job work)' }[A.activity]);
    kv('What the business does', B.activityDesc); kv('Area / GIDC estate', B.area); kv('District', B.district); kv('New unit or expansion', A.stage === 'new' ? 'New unit' : 'Expansion of running unit');
    if (A.activity === 'jobwork') { kv('Principals (for whom job work is done)', (A.jobwork || {}).principals); kv('Basis of job charge', (A.jobwork || {}).basis); }
    if (A.activity === 'contract' || A.activity === 'service') { const ob = A.orderBook || {}; kv('Order book in hand (Rs lakh)', ob.value); kv('Main customers / contracts', ob.customers); kv('Period to execute (months)', ob.months); }
    R();
    R({ v: '3. PROJECT COST BY TYPE (live from Tab 3 capex rows, incl. GST)', s: 'section' });
    Object.entries(KINDS).forEach(([k, d]) => R({ v: d.label, s: 'label' }, { f: `SUMIF(${P3}!$E$33:$E$52,"*| ${d.label}*",${P3}!$H$33:$H$52)` }));
    R({ v: 'Total project cost (Tab 6)', s: 'label' }, { f: "'6_Capex_&_Means'!B38" });
    const P = A.project || {}, Fd = A.funds || {};
    R({ v: 'Margin money for working capital (memo, not in Tab 6)', s: 'label' }, blank(P.marginMoneyWC) ? 'Not given' : P.marginMoneyWC);
    R();
    R({ v: '4. MEANS OF FINANCE', s: 'section' });
    R({ v: 'Proposed term loans (Tab 6)', s: 'label' }, { f: "'6_Capex_&_Means'!F33" });
    R({ v: 'Fresh promoter capital (Tab 3 C56)', s: 'label' }, { f: `${P3}!C56` });
    R({ v: 'Fresh unsecured loans (Tab 3 G56)', s: 'label' }, { f: `${P3}!G56` });
    R({ v: 'Subsidy expected (memo; Tab 6 F34 is fixed at 0)', s: 'label' }, blank(Fd.subsidy) ? 'None stated' : Fd.subsidy, { v: Fd.subsidyScheme || '', s: 'wrap' });
    R({ v: 'Term loan as % of project cost', s: 'label' }, { f: "IFERROR('6_Capex_&_Means'!F33/'6_Capex_&_Means'!B38,0)" });
    R();
    R({ v: '5. COLLATERAL / SECURITY OFFERED', s: 'section' });
    R({ v: 'Type', s: 'label' }, { v: 'Description', s: 'label' }, { v: 'Value (Rs lakh)', s: 'label' }, { v: 'Owner', s: 'label' }, { v: 'Basis of value', s: 'label' });
    const col = (A.collateral || []).filter(c => c && (c.type || c.value));
    const c0 = rows.length + 1;
    if (!col.length) R({ v: 'Unknown - collect from customer', s: 'wrap' });
    col.forEach(c => R(c.type || 'Unknown', c.description || '', isNum(c.value) ? c.value : 'Unknown', c.owner || 'Unknown', c.basis || ''));
    const c1 = rows.length;
    R({ v: 'Total collateral value', s: 'label' }, null, { f: col.length ? `SUM(C${c0}:C${c1})` : '0' });
    R({ v: 'Collateral / total limits (Tab 3 E73)', s: 'label' }, null, { f: `IFERROR(C${rows.length}/${P3}!E73,0)` });
    return { name: '18_Profile_&_Collateral', cols: [44, 34, 18, 26, 26], rows };
  }
  function unknownSheet(res, warns) {
    const rows = []; const R = (...c) => rows.push(c);
    R({ v: 'UNKNOWN — COLLECT FROM CUSTOMER', s: 'title' });
    R({ v: 'These inputs were not answered. They are BLANK in the workbook (treated as zero by formulas). Fill them on the sheet/cell shown before relying on the output.', s: 'wrap' });
    R({ v: 'Question', s: 'label' }, { v: 'Sheet', s: 'label' }, { v: 'Cell', s: 'label' });
    if (!res.unknown.length) R('None — every mapped question was answered.');
    res.unknown.forEach(u => R({ v: u.label, s: 'wrap' }, u.sheet, u.ref));
    R(); R({ v: 'POINTS TO CHECK (soft warnings; no value was changed)', s: 'section' });
    if (!warns.length) R('None');
    warns.forEach(x => R({ v: x.msg, s: 'wrap' }, x.where));
    R(); R({ v: 'DEFAULTS USED (shown to the user and accepted; edit the cell to change)', s: 'section' });
    res.defaultsUsed.forEach(d => R({ v: `${d.label}: ${typeof d.value === 'number' ? (d.value <= 1 && d.value > 0 ? (Math.round(d.value * 10000) / 100) + '%' : d.value) : d.value} — ${d.why}`, s: 'wrap' }, d.ref.split('!')[0], d.ref.split('!')[1]));
    R(); R({ v: 'NOTES', s: 'section' });
    res.notes.forEach(n => R({ v: n, s: 'wrap' }));
    R({ v: 'Template corrections in this copy: see README sheet notes in the DPR Assistant (BS "==" formula, existing-TL double count, hard-coded names, % of total assets denominator, FACR residue, stress-average, yield factor).', s: 'wrap' });
    return { name: '19_Unknown_Collect', cols: [90, 26, 12], rows };
  }
  // placeholders -> live workbook references (figures never come from the AI)
  const PH = {
    BUSINESS_NAME: { f: `${P3}!C5`, t: 'text' }, CONSTITUTION: { f: `${P3}!G5`, t: 'text' }, ACTIVITY: { f: `${P3}!C6`, t: 'text' }, LOCATION: { f: `${P3}!G6`, t: 'text' },
    BASE_FY: { f: `${P3}!G7`, t: 'int' }, FIRST_PROJ_FY: { f: "'5_Projection_Assumptions'!D4", t: 'int' }, LAST_TL_FY: { f: `${P3}!G11`, t: 'int' },
    PROJECT_COST: { f: "'6_Capex_&_Means'!B38", t: 'lakh' }, TERM_LOAN: { f: "'6_Capex_&_Means'!F33", t: 'lakh' }, PROMOTER_CONTRIBUTION: { f: "'6_Capex_&_Means'!F31", t: 'lakh' },
    UNSECURED_LOANS: { f: "'6_Capex_&_Means'!F32", t: 'lakh' }, WC_LIMITS: { f: "'2_Executive_Summary'!G7", t: 'lakh' }, NONFUND_LIMITS: { f: "'2_Executive_Summary'!G8", t: 'lakh' },
    SALES_Y1: { f: "'11_Profit_&_Loss'!E9", t: 'lakh' }, SALES_Y3: { f: "'11_Profit_&_Loss'!G9", t: 'lakh' }, PAT_Y1: { f: "'11_Profit_&_Loss'!E32", t: 'lakh' },
    CASH_ACCRUALS_Y1: { f: "'11_Profit_&_Loss'!E34", t: 'lakh' }, AVG_DSCR: { f: "'15_DSCR_&_BEP'!L13", t: 'x' }, MIN_DSCR: { f: "MINIFS('15_DSCR_&_BEP'!B13:K13,'15_DSCR_&_BEP'!B12:K12,\">0\")", t: 'x' },
    BEP_SALES_Y1: { f: "'15_DSCR_&_BEP'!B23", t: 'lakh' }, INSTALLED_CAPACITY: { f: "'5_Projection_Assumptions'!D8", t: 'num' }, UTILISATION_Y1: { f: "'5_Projection_Assumptions'!D9", t: 'pct' },
    TOTAL_COLLATERAL: { f: null, t: 'lakh' }
  };
  const fmt = (t, ref) => t === 'text' ? ref : t === 'int' ? `TEXT(${ref},"0")` : t === 'lakh' ? `"Rs "&TEXT(${ref},"#,##0.00")&" lakh"` : t === 'x' ? `TEXT(${ref},"0.00")&"x"` : t === 'pct' ? `TEXT(${ref},"0%")` : `TEXT(${ref},"#,##0")`;
  function textToFormula(text) {
    // split into literal chunks (<=240 chars, Excel's 255 limit) and placeholder refs
    const parts = []; let last = 0;
    for (const m of text.matchAll(/\{\{([A-Z0-9_]+)\}\}/g)) {
      if (m.index > last) parts.push({ lit: text.slice(last, m.index) });
      const p = PH[m[1]]; parts.push(p && p.f ? { ref: fmt(p.t, p.f) } : { lit: '[' + m[1].toLowerCase().replace(/_/g, ' ') + ': see workbook]' });
      last = m.index + m[0].length;
    }
    if (last < text.length) parts.push({ lit: text.slice(last) });
    const out = [];
    for (const p of parts) {
      if (p.ref) { out.push(p.ref); continue; }
      for (let i = 0; i < p.lit.length; i += 240) out.push('"' + p.lit.slice(i, i + 240).replace(/"/g, '""') + '"');
    }
    return out.length ? out.join('&') : '""';
  }
  function reportSheet(A, narrative) {
    const rows = []; const R = (...c) => rows.push(c);
    R({ v: 'DETAILED PROJECT REPORT — NARRATIVE', s: 'title' });
    R({ v: narrative && narrative.sections ? 'Draft text written with AI help from the answers given. Every figure below is a live link to the workbook (not typed by the AI). Read and edit before use.' : 'Notes in the borrower\'s own words (no AI used).', s: 'wrap' });
    R();
    const secs = narrative && narrative.sections ? narrative.sections : Object.entries(A.narrative || {}).filter(([k, v]) => v).map(([k, v]) => ({ title: { promoters: 'Promoter background', process: 'Manufacturing / service process', market: 'Market, buyers and selling', swot: 'Strengths and risks', project: 'About the project' }[k] || k, text: v }));
    secs.forEach(s => {
      R({ v: s.title, s: 'section' });
      String(s.text).split(/\n+/).filter(Boolean).forEach(p => R({ f: textToFormula(p), s: 'wrap' }));
      R();
    });
    return { name: '20_Report_Narrative', cols: [120], rows };
  }
  // ---------- AI-safety helpers ----------
  const NUMWORDS = /\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|twenty|thirty|forty|fifty|hundred|thousand|lakh|lakhs|crore|crores|million|billion|percent|half|double|triple|dozen)\b/i;
  function checkNarrative(sections, allowedText) {
    // drop any sentence containing a digit or number-word that is not inside {{PLACEHOLDER}} and not in the user's own words
    const allowed = new Set((String(allowedText || '').match(/\d[\d,.]*/g) || []));
    const dropped = []; const clean = sections.map(s => {
      const sentences = String(s.text || '').split(/(?<=[.!?])\s+/);
      const keep = sentences.filter(sen => {
        const bare = sen.replace(/\{\{[A-Z0-9_]+\}\}/g, '');
        const nums = bare.match(/\d[\d,.]*/g) || [];
        const bad = nums.some(x => !allowed.has(x)) || NUMWORDS.test(bare) || /\{\{(?![A-Z0-9_]+\}\})/.test(sen) || (sen.match(/\{\{([A-Z0-9_]+)\}\}/g) || []).some(p => !PH[p.slice(2, -2)]);
        if (bad) dropped.push(sen); return !bad;
      });
      return { title: String(s.title || '').replace(/\d/g, '').slice(0, 80), text: keep.join(' ') };
    });
    return { sections: clean, dropped };
  }
  const PII = [
    [/\b[A-Z]{5}\d{4}[A-Z]\b/g, '[PAN removed]'], [/\b\d{4}\s?\d{4}\s?\d{4}\b/g, '[Aadhaar removed]'], [/\b\d{9,18}\b/g, '[number removed]'],
    [/[\w.+-]+@[\w-]+\.[\w.]+/g, '[email removed]'], [/(\+91[\s-]?)?\b[6-9]\d{9}\b/g, '[phone removed]'], [/\b[A-Z]{4}0[A-Z0-9]{6}\b/g, '[IFSC removed]'], [/\b\d{2}[A-Z]{5}\d{4}[A-Z]\d[Z][A-Z\d]\b/g, '[GSTIN removed]']
  ];
  function redact(text, names) {
    let t = String(text || '');
    for (const [re, rep] of PII) t = t.replace(re, rep);
    (names || []).filter(n => n && n.trim().length > 2).sort((a, b) => b.length - a.length).forEach((n, i) => {
      const re = new RegExp(n.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'); t = t.replace(re, i === 0 ? 'the unit' : 'the promoter');
    });
    return t;
  }
  return { build, warnings, benchmarks, histCalc, projSales, profileSheet, unknownSheet, reportSheet, checkNarrative, redact, textToFormula, taxDefault,
    activeMonthsY1, horizon, projYears, emi, facilities, CLASSES, KINDS, DEFAULTS, HIST_ROWS, HIST_LABEL, PH, S3, S4, S5, fyLabel, revenueDirect };
});
