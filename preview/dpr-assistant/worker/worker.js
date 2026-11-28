// DPR Assistant AI proxy (Cloudflare Worker).
// Holds the DeepSeek key as a secret (DEEPSEEK_API_KEY). Only two narrow, server-built prompts are allowed:
//   POST /extract   -> turn the user's own words into structured fields (user confirms every value in the app)
//   POST /narrative -> write DPR narrative text with {{PLACEHOLDERS}} instead of figures
// Origin check, per-IP + daily rate limits (KV "RL"), size limits. Request bodies are never logged or stored.
const DS = 'https://api.deepseek.com/chat/completions';
const LIM = { extract: { ip: 40, day: 600 }, narrative: { ip: 6, day: 80 } };
const PHD = { BUSINESS_NAME: 'name of the business', CONSTITUTION: 'legal form', ACTIVITY: 'line of business', LOCATION: 'unit location', BASE_FY: 'latest audited/base financial year',
  FIRST_PROJ_FY: 'first projected financial year', LAST_TL_FY: 'financial year in which the last term-loan instalment is paid (NOT a completion date)', PROJECT_COST: 'total project cost',
  TERM_LOAN: 'proposed term loan amount', PROMOTER_CONTRIBUTION: "promoters' contribution to project cost", UNSECURED_LOANS: 'unsecured loans in means of finance',
  WC_LIMITS: 'fund-based working-capital limits', NONFUND_LIMITS: 'LC/BG limits', SALES_Y1: 'first projected year sales', SALES_Y3: 'third projected year sales',
  PAT_Y1: 'first projected year net profit', CASH_ACCRUALS_Y1: 'first projected year cash accruals', AVG_DSCR: 'average DSCR over the loan period', MIN_DSCR: 'lowest annual DSCR',
  BEP_SALES_Y1: 'break-even sales in the first projected year', INSTALLED_CAPACITY: 'installed capacity per year (units)', UTILISATION_Y1: 'capacity utilisation in the first projected year' };
const PH = Object.keys(PHD);
const SYS_EXTRACT = `You convert a small-business owner's own words into form fields for an Indian bank project report.
Rules: (1) Return ONLY JSON: {"values":{<field key>: value}, "followup": "<one short, friendly next question in plain English, or empty>"}.
(2) Fill a field ONLY when the user's text states it. Never guess, estimate, compute averages or fill typical values. If unsure, leave the field out.
(3) Money fields are in Rs lakh: convert only explicit units (1 crore = 100 lakh; Rs 5,00,000 = 5 lakh). Percent fields as fractions (18% -> 0.18). Dates as YYYY-MM-DD.
(4) For choice fields use exactly one of the given options. (5) Lists (type "list") are arrays of objects with the given sub-keys.
(6) The followup asks about the most important missing field for THIS card only; do not ask for PAN, Aadhaar or account numbers.`;
const SYS_NARR = `You write the narrative sections of a Detailed Project Report (DPR) for an Indian MSME bank loan, in plain, professional English.
HARD RULES: (1) Never write any number, amount, percentage, year, ratio, count or number-word (one, two, lakh, crore, percent, double...). Where a figure belongs, insert one of these placeholders exactly (meaning in brackets): ${PH.map(p => '{{' + p + '}} (' + PHD[p] + ')').join('; ')}. Use only these placeholders, only where the meaning fits.
(2) Use only facts given in the input. If something is unknown, write "to be confirmed with the borrower". Do not invent customers, certifications, awards, market sizes or growth rates.
(3) Refer to the business as {{BUSINESS_NAME}} and to people by role (the promoter), never by name.
(4) Return ONLY JSON: {"sections":[{"title":"...","text":"..."}]} with these sections in order: Project description, Promoter background, Manufacturing / service process, Market and selling arrangement, Strengths, weaknesses, opportunities and threats, Conclusion. 80-160 words each; SWOT may use short lines separated by newlines.`;
function cors(env, origin) {
  return { 'Access-Control-Allow-Origin': origin === env.ALLOW_ORIGIN ? origin : env.ALLOW_ORIGIN, 'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'content-type', 'Access-Control-Max-Age': '86400', 'Vary': 'Origin', 'Cache-Control': 'no-store' };
}
const json = (h, code, obj) => new Response(JSON.stringify(obj), { status: code, headers: { ...h, 'content-type': 'application/json' } });
async function sha(s) { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return [...new Uint8Array(b)].slice(0, 12).map(x => x.toString(16).padStart(2, '0')).join(''); }
async function limit(env, kind, ip) {
  const now = new Date(); const day = now.toISOString().slice(0, 10), hour = now.toISOString().slice(0, 13);
  const kIp = `ip:${kind}:${hour}:${await sha(ip + (env.SALT || 'dpr'))}`, kDay = `day:${kind}:${day}`;
  const [a, b] = await Promise.all([env.RL.get(kIp), env.RL.get(kDay)]);
  const ni = parseInt(a || '0', 10), nd = parseInt(b || '0', 10);
  if (ni >= LIM[kind].ip) return 'Too many requests from this device this hour. The form still works without AI.';
  if (nd >= LIM[kind].day) return 'Daily AI limit reached. The form still works without AI.';
  await Promise.all([env.RL.put(kIp, String(ni + 1), { expirationTtl: 3700 }), env.RL.put(kDay, String(nd + 1), { expirationTtl: 2 * 86400 })]);
  return null;
}
async function deepseek(env, system, user, maxTokens) {
  const r = await fetch(DS, { method: 'POST', headers: { 'content-type': 'application/json', authorization: 'Bearer ' + env.DEEPSEEK_API_KEY },
    body: JSON.stringify({ model: 'deepseek-chat', temperature: 0.2, max_tokens: maxTokens, response_format: { type: 'json_object' },
      messages: [{ role: 'system', content: system }, { role: 'user', content: user }] }) });
  if (!r.ok) throw new Error('AI service error ' + r.status);
  const d = await r.json(); return JSON.parse(d.choices[0].message.content);
}
const str = (v, n) => String(v == null ? '' : v).slice(0, n);
export default {
  async fetch(req, env) {
    const origin = req.headers.get('Origin') || '';
    const h = cors(env, origin); const url = new URL(req.url);
    if (req.method === 'OPTIONS') return new Response(null, { status: 204, headers: h });
    if (url.pathname === '/health') return json(h, 200, { ok: true, ai: env.AI_ENABLED === 'true' && !!env.DEEPSEEK_API_KEY });
    if (origin !== env.ALLOW_ORIGIN) return json(h, 403, { error: 'Origin not allowed' });
    if (env.AI_ENABLED !== 'true' || !env.DEEPSEEK_API_KEY) return json(h, 503, { error: 'AI is switched off.' });
    if (req.method !== 'POST') return json(h, 405, { error: 'POST only' });
    const len = parseInt(req.headers.get('content-length') || '0', 10); if (len > 12000) return json(h, 413, { error: 'Too large' });
    let b; try { b = await req.json(); } catch { return json(h, 400, { error: 'Bad JSON' }); }
    const ip = req.headers.get('CF-Connecting-IP') || 'x';
    try {
      if (url.pathname === '/extract') {
        const fields = (Array.isArray(b.fields) ? b.fields : []).slice(0, 40).map(f => ({ key: str(f.key, 40), label: str(f.label, 120), type: str(f.type, 12), options: Array.isArray(f.options) ? f.options.slice(0, 12).map(o => str(o, 40)) : undefined, sub: Array.isArray(f.sub) ? f.sub.slice(0, 12).map(o => str(o, 40)) : undefined }));
        const text = str(b.text, 2000); if (!text.trim() || !fields.length) return json(h, 400, { error: 'Nothing to read' });
        const why = await limit(env, 'extract', ip); if (why) return json(h, 429, { error: why });
        const out = await deepseek(env, SYS_EXTRACT, JSON.stringify({ card: str(b.card, 60), context: str(b.context, 300), fields, user_text: text }), 700);
        const allowed = new Set(fields.map(f => f.key)); const values = {};
        for (const [k, v] of Object.entries(out.values || {})) if (allowed.has(k)) values[k] = v;
        return json(h, 200, { values, followup: str(out.followup, 200) });
      }
      if (url.pathname === '/narrative') {
        const facts = {}; for (const [k, v] of Object.entries(b.facts || {}).slice(0, 30)) facts[str(k, 40)] = str(v, 1200);
        const why = await limit(env, 'narrative', ip); if (why) return json(h, 429, { error: why });
        const out = await deepseek(env, SYS_NARR, JSON.stringify({ facts }), 2200);
        const sections = (out.sections || []).slice(0, 8).map(s => ({ title: str(s.title, 80), text: str(s.text, 2500) }));
        return json(h, 200, { sections });
      }
      return json(h, 404, { error: 'Not found' });
    } catch (e) { return json(h, 502, { error: 'AI request failed. The form still works without AI.' }); }
  }
};
