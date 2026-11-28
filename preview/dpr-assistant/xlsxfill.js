/* xlsxfill.js — writes values into an existing .xlsx by patching sheet XML only.
   Keeps every formula, style, chart, validation untouched. Works in browser and Node (pass JSZip). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsxFill = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const STYLE = { title: 26, section: 3, label: 6, input: 29, wrap: 76 };
  function serial(iso) { const [y, m, d] = iso.split('-').map(Number); return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); }
  function colNum(c) { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function colName(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function cellXml(ref, v, s) {
    const sa = s != null ? ` s="${s}"` : '';
    if (v === null || v === undefined || v === '') return `<c r="${ref}"${sa}/>`;
    if (typeof v === 'object' && v.f) return `<c r="${ref}"${sa} t="str"><f>${esc(v.f.replace(/^=/, ''))}</f><v></v></c>`;
    if (typeof v === 'object' && v.date) return `<c r="${ref}"${sa}><v>${serial(v.date)}</v></c>`;
    if (typeof v === 'number') { if (!isFinite(v)) throw new Error('bad number at ' + ref); return `<c r="${ref}"${sa}><v>${v}</v></c>`; }
    return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  }
  async function sheetPaths(zip) {
    const wb = await zip.file('xl/workbook.xml').async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const map = {};
    const relMap = {};
    for (const m of rels.matchAll(/<Relationship [^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])[1], t = /Target="([^"]+)"/.exec(m[0])[1];
      relMap[id] = t.replace(/^\//, '').replace(/^(?!xl\/)/, 'xl/');
    }
    for (const m of wb.matchAll(/<sheet [^>]*>/g)) {
      const name = /name="([^"]+)"/.exec(m[0])[1].replace(/&amp;/g, '&');
      const rid = /r:id="([^"]+)"/.exec(m[0])[1];
      map[name] = relMap[rid];
    }
    return map;
  }
  function patchCell(xml, ref, v) {
    const re = new RegExp(`<c r="${ref}"( s="\\d+")?( t="\\w+")?(/>|>[\\s\\S]*?</c>)`);
    const m = re.exec(xml);
    if (!m) throw new Error('Input cell not found in template: ' + ref);
    if (/<f>/.test(m[0])) throw new Error('Refusing to overwrite a formula cell: ' + ref);
    const s = m[1] ? Number(m[1].match(/\d+/)[0]) : null;
    return xml.slice(0, m.index) + cellXml(ref, v, s) + xml.slice(m.index + m[0].length);
  }
  function sheetXml(sheet) {
    // sheet: {cols:[widths], rows:[[cell|null]]}; cell = {v, s} or plain value
    const cols = (sheet.cols || []).map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
    let rows = '';
    sheet.rows.forEach((r, i) => {
      if (!r || !r.length) return;
      const cells = r.map((c, j) => {
        if (c === null || c === undefined) return '';
        const o = (typeof c === 'object' && !c.f && !c.date && 'v' in c) ? c : { v: c };
        const s = o.s != null ? (typeof o.s === 'string' ? STYLE[o.s] : o.s) : null;
        return cellXml(colName(j + 1) + (i + 1), o.v, s);
      }).join('');
      rows += `<row r="${i + 1}">${cells}</row>`;
    });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_R}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${rows}</sheetData><pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/></worksheet>`;
  }
  async function addSheet(zip, name, sheet) {
    if (name.length > 31 || /[\\\/?*\[\]:]/.test(name)) throw new Error('bad sheet name ' + name);
    let wb = await zip.file('xl/workbook.xml').async('string');
    let rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    let ct = await zip.file('[Content_Types].xml').async('string');
    const ids = [...wb.matchAll(/sheetId="(\d+)"/g)].map(m => +m[1]);
    const sid = Math.max(...ids) + 1;
    const rnums = [...rels.matchAll(/Id="rId(\d+)"/g)].map(m => +m[1]);
    const rid = 'rId' + (Math.max(...rnums) + 1);
    let n = 1; while (zip.file(`xl/worksheets/sheet${n}.xml`)) n++;
    const path = `xl/worksheets/sheet${n}.xml`;
    zip.file(path, sheetXml(sheet));
    wb = wb.replace('</sheets>', `<sheet xmlns:r="${NS_R}" name="${esc(name)}" sheetId="${sid}" state="visible" r:id="${rid}"/></sheets>`);
    rels = rels.replace('</Relationships>', `<Relationship Type="${NS_R}/worksheet" Target="/${path}" Id="${rid}"/></Relationships>`);
    ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
    zip.file('xl/workbook.xml', wb); zip.file('xl/_rels/workbook.xml.rels', rels); zip.file('[Content_Types].xml', ct);
  }
  /** cells: {sheetName: {A1: value}}; extraSheets: [{name, cols, rows}] */
  async function fill(JSZip, templateBuf, cells, extraSheets, outType) {
    const zip = await JSZip.loadAsync(templateBuf);
    const paths = await sheetPaths(zip);
    for (const [sheet, vals] of Object.entries(cells)) {
      const p = paths[sheet]; if (!p) throw new Error('No sheet ' + sheet);
      let xml = await zip.file(p).async('string');
      for (const [ref, v] of Object.entries(vals)) xml = patchCell(xml, ref, v);
      zip.file(p, xml);
    }
    for (const s of extraSheets || []) await addSheet(zip, s.name, s);
    return zip.generateAsync({ type: outType || 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  /* minimal reader for the uploaded Tab-4 sheet: returns {ref: value} of first sheet */
  async function readFirstSheet(JSZip, buf) {
    const zip = await JSZip.loadAsync(buf);
    const paths = await sheetPaths(zip);
    const first = Object.values(paths)[0];
    const xml = await zip.file(first).async('string');
    let shared = [];
    const ss = zip.file('xl/sharedStrings.xml');
    if (ss) { const sx = await ss.async('string'); shared = [...sx.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => t[1]).join('')); }
    const un = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
    const out = {};
    for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = m[1], attrs = m[2], inner = m[3] || '';
      const t = (/t="(\w+)"/.exec(attrs) || [])[1];
      const v = (/<v>([\s\S]*?)<\/v>/.exec(inner) || [])[1];
      if (t === 's' && v != null) out[ref] = un(shared[+v] || '');
      else if (t === 'inlineStr') out[ref] = un([...inner.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join(''));
      else if (t === 'str') out[ref] = v != null ? un(v) : '';
      else if (v != null && v !== '') out[ref] = Number(v);
    }
    return out;
  }
  return { fill, readFirstSheet, serial, colName, colNum, STYLE };
});
