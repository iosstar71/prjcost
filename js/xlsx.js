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

  // 꾸밈 번호 (styles.xml 의 cellXfs 순서)
  const STYLE = { plain: 0, head: 1, label: 2, body: 3, money: 4, dec: 5, total: 6, wrap: 7, ok: 8, warn: 9, title: 10 };
  const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="10"/><name val="맑은 고딕"/></font><font><b/><sz val="10"/><name val="맑은 고딕"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="맑은 고딕"/></font><font><b/><sz val="13"/><color rgb="FF0F172A"/><name val="맑은 고딕"/></font></fonts>
<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF334155"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE2E8F0"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFCBD5E1"/></left><right style="thin"><color rgb="FFCBD5E1"/></right><top style="thin"><color rgb="FFCBD5E1"/></top><bottom style="thin"><color rgb="FFCBD5E1"/></bottom></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="11">
<xf/>
<xf fontId="2" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf fontId="1" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf fontId="0" borderId="1" applyBorder="1"/>
<xf numFmtId="3" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="4" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="3" fontId="1" fillId="4" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf borderId="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="4" fontId="1" fillId="5" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="4" fontId="1" fillId="6" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf fontId="3" applyFont="1"/>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

  // 시트: { name, rows: [{ cells:[값 | {v, s}], style, merge: [첫 열, 끝 열](0부터, 그 줄 안에서 합침), ht: 줄 높이 }], cols: [너비…], hidden, freeze }
  //  숫자 칸은 줄 꾸밈이 없으면 금액(#,##0), 글자 칸은 테두리만
  function sheetXml(sh) {
    const merges = [];
    const rows = (sh.rows || []).map((row, ri) => {
      const r = Array.isArray(row) ? { cells: row } : row;
      const raw = (r.cells || []).slice();
      if (r.merge) {   // 합친 칸의 나머지도 같은 꾸밈 (테두리가 끊기지 않게)
        const sty = (raw[r.merge[0]] && raw[r.merge[0]].s) || r.style || "body";
        for (let k = r.merge[0] + 1; k <= r.merge[1]; k++) if (raw[k] === undefined || raw[k] === "") raw[k] = { v: "", s: sty };
        merges.push(`${colName(r.merge[0])}${ri + 1}:${colName(r.merge[1])}${ri + 1}`);
      }
      const cells = raw.map((c, ci) => {
        const cell = c !== null && typeof c === "object" && "v" in c ? c : { v: c };
        const v = cell.v;
        if (v === null || v === undefined || v === "") return r.style || cell.s ? `<c r="${colName(ci)}${ri + 1}" s="${STYLE[cell.s || r.style] || 0}"/>` : "";
        const isNum = typeof v === "number" && Number.isFinite(v);
        const s = STYLE[cell.s || r.style || (isNum ? "money" : "body")] || 0;
        return isNum ? `<c r="${colName(ci)}${ri + 1}" s="${s}"><v>${v}</v></c>`
          : `<c r="${colName(ci)}${ri + 1}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
      }).join("");
      return `<row r="${ri + 1}"${r.ht ? ` ht="${r.ht}" customHeight="1"` : ""}>${cells}</row>`;
    }).join("");
    const cols = (sh.cols || []).length ? `<cols>${sh.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
    const view = sh.freeze === false || sh.hidden ? `<sheetViews><sheetView workbookViewId="0"/></sheetViews>`
      : `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${view}${cols}<sheetData>${rows}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : ""}</worksheet>`;
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
      { name: "xl/styles.xml", data: STYLES_XML }
    ];
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s) }));
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
