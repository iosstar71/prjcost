// =============================================================================
// 엑셀(xlsx) 만들기 · 읽기 — 외부 라이브러리 없이 (브라우저 · Node 공용)
//  만들기: 압축하지 않은 zip(STORE) + 꾸밈(머리 · 항목 · 금액 · 합계) · 열 너비 · 첫 줄 고정 · 숨은 시트
//  읽기: 압축된 zip 도 읽음(DecompressionStream) — Excel 에서 열어 다시 저장한 파일 · 원본(v4) 파일 모두
//  JSON 은 칸 하나에 32,767자까지만 들어가므로 여러 줄로 나눠 저장 (jsonRows / readJson)
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const enc = new TextEncoder();
  const JSON_CHUNK = 30000;

  // ---------- zip ----------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };

  function zip(files) {
    const parts = [], central = [];
    let offset = 0;
    files.forEach(f => {
      const name = enc.encode(f.name), data = typeof f.data === "string" ? enc.encode(f.data) : f.data, crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);   // UTF-8 이름
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + data.length;
    });
    const size = central.reduce((a, b) => a + b.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, size, true); e.setUint32(16, offset, true);
    return new Blob([...parts, ...central, new Uint8Array(e.buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  async function inflateRaw(data) {
    const ds = new DecompressionStream("deflate-raw");
    const out = await new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer();
    return new Uint8Array(out);
  }

  async function unzip(buf) {
    const b = new Uint8Array(buf), v = new DataView(b.buffer, b.byteOffset, b.byteLength), out = {};
    let eocd = -1;
    for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (v.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error("엑셀(xlsx) 파일이 아니거나 손상되었습니다.");
    let p = v.getUint32(eocd + 16, true);
    const count = v.getUint16(eocd + 10, true);
    for (let k = 0; k < count; k++) {
      const method = v.getUint16(p + 10, true), csize = v.getUint32(p + 20, true), nlen = v.getUint16(p + 28, true);
      const elen = v.getUint16(p + 30, true), clen = v.getUint16(p + 32, true), loc = v.getUint32(p + 42, true);
      const name = new TextDecoder().decode(b.subarray(p + 46, p + 46 + nlen));
      const start = loc + 30 + v.getUint16(loc + 26, true) + v.getUint16(loc + 28, true);
      const data = b.subarray(start, start + csize);
      if (method === 0) out[name] = data;
      else if (method === 8) out[name] = await inflateRaw(data);
      p += 46 + nlen + elen + clen;
    }
    return out;
  }

  // ---------- XML ----------
  const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  const xmlUnesc = s => String(s).replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (m, e) =>
    e[0] === "#" ? String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }[e]);
  const attr = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? xmlUnesc(m[1]) : null; };
  const colName = i => { let s = ""; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const colIndex = ref => { const L = /^[A-Z]+/.exec(ref)[0]; let n = 0; for (const ch of L) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };

  // ---------- 꾸밈 ----------
  //  칸 꾸밈 = 이름(아래 NAMED) 또는 객체 { font: { b, i, sz, color }, fill: "D9D9D9", border: "thin" | "medium" | "light" | { l, r, t, b },
  //            fmt: 기본 번호 | "#,##0" 같은 서식 글, h: left|center|right, v: top|center|bottom, wrap, indent }
  //  build 할 때 쓴 꾸밈만 모아 styles.xml 을 만듦
  const NAMED = {
    plain: {},
    head: { font: { b: 1, color: "FFFFFF" }, fill: "334155", border: "light", h: "center", v: "center" },
    label: { font: { b: 1 }, fill: "F1F5F9", border: "light" },
    body: { border: "light" },
    money: { fmt: 3, border: "light" },
    dec: { fmt: 4, border: "light" },
    total: { fmt: 3, font: { b: 1 }, fill: "E2E8F0", border: "light" },
    wrap: { border: "light", wrap: 1, v: "top" },
    ok: { fmt: 4, font: { b: 1 }, fill: "DCFCE7", border: "light" },
    warn: { fmt: 4, font: { b: 1 }, fill: "FEE2E2", border: "light" },
    title: { font: { b: 1, sz: 13, color: "0F172A" } }
  };
  const FONT_NAME = "맑은 고딕";
  const argb = c => (String(c).length === 6 ? `FF${c}` : String(c));
  const LINE = { thin: ["thin", "FF808080"], medium: ["medium", "FF000000"], light: ["thin", "FFCBD5E1"], black: ["thin", "FF000000"] };
  function Styles() {
    const fonts = [JSON.stringify({})], fills = ["none", "gray125"], borders = [JSON.stringify({})], fmts = [], xfs = [JSON.stringify({})], ids = new Map();
    const idx = (list, key) => { let i = list.indexOf(key); if (i < 0) { list.push(key); i = list.length - 1; } return i; };
    return {
      id(st) {
        const o = typeof st === "string" ? NAMED[st] || {} : st || {};
        const key = JSON.stringify(o);
        if (ids.has(key)) return ids.get(key);
        const x = {
          font: o.font ? idx(fonts, JSON.stringify(o.font)) : 0,
          fill: o.fill ? idx(fills, argb(o.fill)) : 0,
          border: o.border ? idx(borders, JSON.stringify(typeof o.border === "string" ? { l: o.border, r: o.border, t: o.border, b: o.border } : o.border)) : 0,
          fmt: o.fmt === undefined ? 0 : typeof o.fmt === "number" ? o.fmt : 164 + idx(fmts, o.fmt),
          al: o.h || o.v || o.wrap || o.indent ? { h: o.h, v: o.v, wrap: o.wrap, indent: o.indent } : null
        };
        const i = xfs.length;
        xfs.push(x); ids.set(key, i);
        return i;
      },
      xml() {
        const fontXml = f => { const o = JSON.parse(f); return `<font>${o.b ? "<b/>" : ""}${o.i ? "<i/>" : ""}<sz val="${o.sz || 10}"/>${o.color ? `<color rgb="${argb(o.color)}"/>` : ""}<name val="${FONT_NAME}"/></font>`; };
        const fillXml = f => (f === "none" || f === "gray125" ? `<fill><patternFill patternType="${f}"/></fill>` : `<fill><patternFill patternType="solid"><fgColor rgb="${f}"/></patternFill></fill>`);
        const side = (tag, k) => { const L = LINE[k]; return L ? `<${tag} style="${L[0]}"><color rgb="${L[1]}"/></${tag}>` : `<${tag}/>`; };
        const borderXml = bd => { const o = JSON.parse(bd); return `<border>${side("left", o.l)}${side("right", o.r)}${side("top", o.t)}${side("bottom", o.b)}<diagonal/></border>`; };
        const xfXml = x => {
          if (typeof x === "string") return "<xf/>";
          const al = x.al ? `<alignment${x.al.h ? ` horizontal="${x.al.h}"` : ""}${x.al.v ? ` vertical="${x.al.v}"` : ""}${x.al.wrap ? ' wrapText="1"' : ""}${x.al.indent ? ` indent="${x.al.indent}"` : ""}/>` : "";
          return `<xf numFmtId="${x.fmt}" fontId="${x.font}" fillId="${x.fill}" borderId="${x.border}" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${al ? ` applyAlignment="1">${al}</xf>` : "/>"}`;
        };
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${fmts.length ? `<numFmts count="${fmts.length}">${fmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xmlEsc(f)}"/>`).join("")}</numFmts>` : ""}
<fonts count="${fonts.length}">${fonts.map(fontXml).join("")}</fonts>
<fills count="${fills.length}">${fills.map(fillXml).join("")}</fills>
<borders count="${borders.length}">${borders.map(borderXml).join("")}</borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="${xfs.length}">${xfs.map(xfXml).join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
      }
    };
  }

  // 시트: { name, rows: [{ cells:[값 | {v, s}], style, merge: [첫 열, 끝 열](0부터, 그 줄 안에서 합침), ht: 줄 높이 }], cols: [너비 | { w, hidden }],
  //        hidden, freeze(false | 없으면 첫 줄 고정 | { x: 고정할 열 수, y: 고정할 줄 수 }), merges: ["B2:D2" …], gridLines(false 면 눈금선 숨김),
  //        page: { landscape, fitWidth(한 장 너비에 맞춤) }, rowHeight(기본 줄 높이) }
  //  숫자 칸은 꾸밈이 없으면 금액(#,##0), 글자 칸은 테두리만
  function sheetXml(sh, styles) {
    const merges = (sh.merges || []).slice();
    const rows = (sh.rows || []).map((row, ri) => {
      const r = Array.isArray(row) ? { cells: row } : row || { cells: [] };
      const raw = (r.cells || []).slice();
      if (r.merge) {   // 합친 칸의 나머지도 같은 꾸밈 (테두리가 끊기지 않게)
        const sty = (raw[r.merge[0]] && raw[r.merge[0]].s) || r.style || "body";
        for (let k = r.merge[0] + 1; k <= r.merge[1]; k++) if (raw[k] === undefined || raw[k] === "") raw[k] = { v: "", s: sty };
        merges.push(`${colName(r.merge[0])}${ri + 1}:${colName(r.merge[1])}${ri + 1}`);
      }
      const cells = raw.map((c, ci) => {
        const cell = c !== null && typeof c === "object" && "v" in c ? c : { v: c };
        const v = cell.v;
        if (v === null || v === undefined || v === "") return r.style || cell.s ? `<c r="${colName(ci)}${ri + 1}" s="${styles.id(cell.s || r.style)}"/>` : "";
        const isNum = typeof v === "number" && Number.isFinite(v);
        const s = styles.id(cell.s || r.style || (isNum ? "money" : "body"));
        return isNum ? `<c r="${colName(ci)}${ri + 1}" s="${s}"><v>${v}</v></c>`
          : `<c r="${colName(ci)}${ri + 1}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
      }).join("");
      return `<row r="${ri + 1}"${r.ht ? ` ht="${r.ht}" customHeight="1"` : ""}>${cells}</row>`;
    }).join("");
    const cols = (sh.cols || []).length ? `<cols>${sh.cols.map((w, i) => { const o = typeof w === "object" ? w : { w }; return `<col min="${i + 1}" max="${i + 1}" width="${o.w || 10}" customWidth="1"${o.hidden ? ' hidden="1"' : ""}/>`; }).join("")}</cols>` : "";
    const grid = sh.gridLines === false ? ' showGridLines="0"' : "";
    let pane = "";
    if (sh.freeze && typeof sh.freeze === "object") {
      const x = sh.freeze.x || 0, y = sh.freeze.y || 0;
      if (x || y) pane = `<pane${x ? ` xSplit="${x}"` : ""}${y ? ` ySplit="${y}"` : ""} topLeftCell="${colName(x)}${y + 1}" activePane="${x && y ? "bottomRight" : x ? "topRight" : "bottomLeft"}" state="frozen"/>`;
    } else if (sh.freeze !== false && !sh.hidden) pane = `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>`;
    const view = `<sheetViews><sheetView workbookViewId="0"${grid}>${pane}</sheetView></sheetViews>`;
    const pg = sh.page || null;
    const pr = pg && pg.fitWidth ? `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` : "";
    const fmt = sh.rowHeight ? `<sheetFormatPr defaultRowHeight="${sh.rowHeight}" customHeight="1"/>` : "";
    const setup = pg ? `<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="${pg.landscape ? "landscape" : "portrait"}"${pg.fitWidth ? ' fitToWidth="1" fitToHeight="0"' : ""}/>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${pr}${view}${fmt}${cols}<sheetData>${rows}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : ""}${setup}</worksheet>`;
  }

  function build(sheets) {
    const n = sheets.length;
    const files = [
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>` },
      { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
      { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xmlEsc(String(s.name).slice(0, 31))}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>` },
      { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    ];
    const styles = Styles();
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s, styles) }));
    files.push({ name: "xl/styles.xml", data: styles.xml() });
    return zip(files);
  }

  // ---------- 읽기 ----------
  // 결과: { sheets: [{ name, hidden, rows: [[값…]…] }], sheet(name) }
  async function read(src) {
    const buf = src instanceof ArrayBuffer ? src : ArrayBuffer.isView(src) ? src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength) : await src.arrayBuffer();
    const files = await unzip(buf);
    const text = name => { const k = Object.keys(files).find(f => f.toLowerCase() === name.toLowerCase()); return k ? new TextDecoder().decode(files[k]) : ""; };
    const shared = [];
    (text("xl/sharedStrings.xml").match(/<si\b[\s\S]*?<\/si>/g) || []).forEach(si => {
      shared.push(xmlUnesc((si.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/^<t\b[^>]*>|<\/t>$/g, "")).join("")));
    });
    const rels = {};
    (text("xl/_rels/workbook.xml.rels").match(/<Relationship\b[^>]*>/g) || []).forEach(r => { rels[attr(r, "Id")] = attr(r, "Target"); });
    const sheets = (text("xl/workbook.xml").match(/<sheet\b[^>]*>/g) || []).map(tag => {
      const target = String(rels[attr(tag, "r:id")] || "").replace(/^\/?xl\//, "");
      return { name: attr(tag, "name"), hidden: /^(hidden|veryHidden)$/.test(attr(tag, "state") || ""), rows: parseSheet(text(`xl/${target}`), shared) };
    });
    return { sheets, sheet: name => sheets.find(s => s.name === name) || null };
  }

  function parseSheet(xml, shared) {
    const rows = [];
    (xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []).forEach((rowXml, i) => {
      const rTag = /^<row\b[^>]*>/.exec(rowXml)[0];
      const rIdx = Number(attr(rTag, "r") || i + 1) - 1;
      const row = [];
      (rowXml.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || []).forEach((cXml, ci) => {
        const tag = /^<c\b[^>]*?\/?>/.exec(cXml)[0];
        const ref = attr(tag, "r"), type = attr(tag, "t");
        const col = ref ? colIndex(ref) : ci;
        let val = "";
        if (type === "inlineStr") val = xmlUnesc((cXml.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/^<t\b[^>]*>|<\/t>$/g, "")).join(""));
        else {
          const m = /<v>([\s\S]*?)<\/v>/.exec(cXml);
          const raw = m ? m[1] : "";
          if (type === "s") val = shared[Number(raw)] || "";
          else if (type === "str" || type === "e") val = xmlUnesc(raw);
          else if (type === "b") val = raw === "1";
          else val = raw === "" ? "" : Number(raw);
        }
        row[col] = val;
      });
      for (let k = 0; k < row.length; k++) if (row[k] === undefined) row[k] = "";
      rows[rIdx] = row;
    });
    for (let k = 0; k < rows.length; k++) if (!rows[k]) rows[k] = [];
    return rows;
  }

  // JSON 을 A열 여러 줄로 (원본 v4 는 A1 한 칸 — 읽을 때 둘 다 받음)
  const jsonRows = obj => { const s = JSON.stringify(obj); const out = []; for (let i = 0; i < s.length || i === 0; i += JSON_CHUNK) out.push([s.slice(i, i + JSON_CHUNK)]); return out; };
  const readJson = sheet => { if (!sheet) return null; const s = sheet.rows.map(r => (r && r[0] !== undefined ? String(r[0]) : "")).join(""); return s.trim() ? JSON.parse(s) : null; };

  PC.Xlsx = { build, read, jsonRows, readJson, colName, zip, unzip, xmlEsc, xmlUnesc };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Xlsx;
})(typeof window !== "undefined" ? window : globalThis);
