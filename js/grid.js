// =============================================================================
// 엑셀처럼 쓰는 표 (table.grid — 내부 · 외주 인력) — 키보드 이동 · 붙여넣기 · 새 줄 · 오른쪽 채우기
//  표의 열 정의는 views.js 가 data-cols 에 넣어 둠 [{ b: 행 안의 자료 경로, t: 값 읽는 법, l: 머리글 이름 }]
//  자료를 바꾸는 일은 app.js 가 넘겨준 함수(h)로만: pushUndo · setValue · newRow · render
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC;
  const { num } = PC;

  const MONTH = /^(\d{4})\s*[-./년]\s*(\d{1,2})/;
  const TRUE = /^(y|yes|true|1|o|v|✓|✔|예|포함)$/i;
  const parse = (t, raw) => {
    const v = String(raw === undefined || raw === null ? "" : raw).trim();
    if (t === "num") return v === "" ? 0 : num(v);
    if (t === "blank-num") return v === "" ? "" : num(v);
    if (t === "bool") return TRUE.test(v);
    if (t === "month") { const m = MONTH.exec(v); return m ? `${m[1]}-${m[2].padStart(2, "0")}` : ""; }
    return v;
  };
  // 엑셀이 복사한 글: 탭 = 칸, 줄바꿈 = 줄. 따옴표로 묶인 칸 안의 줄바꿈 · 탭도 처리
  function parseTsv(text) {
    const rows = [[]];
    let cur = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') q = false;
        else cur += ch;
      } else if (ch === '"' && cur === "") q = true;
      else if (ch === "\t") { rows[rows.length - 1].push(cur); cur = ""; }
      else if (ch === "\n" || ch === "\r") {
        if (ch === "\r" && text[i + 1] === "\n") i++;
        rows[rows.length - 1].push(cur); cur = ""; rows.push([]);
      } else cur += ch;
    }
    rows[rows.length - 1].push(cur);
    while (rows.length && rows[rows.length - 1].every(c => c === "")) rows.pop();   // 끝 빈 줄
    return rows;
  }
  // 머리글째 복사했는지: 첫 줄의 칸이 표의 머리글 이름(성명 · 등급 · 2026-03 …)과 맞으면 그 열로
  function headerMap(first, cols) {
    const norm = s => String(s || "").replace(/\s|\(.*?\)/g, "").replace(/MM$/i, "").toLowerCase();
    const monthKey = s => { const m = MONTH.exec(String(s)); return m ? `${m[1]}-${m[2].padStart(2, "0")}` : null; };
    const idx = first.map(h => {
      const mk = monthKey(h);
      if (mk) return cols.findIndex(c => c.l === mk);
      const n = norm(h);
      return n ? cols.findIndex(c => norm(c.l) === n || (n.length >= 2 && norm(c.l).startsWith(n))) : -1;
    });
    return idx.filter(i => i >= 0).length >= Math.max(2, Math.ceil(first.filter(Boolean).length / 2)) ? idx : null;
  }

  const Grid = {
    h: null,
    init(h) {
      this.h = h;
      document.addEventListener("keydown", e => this.onKey(e), true);
      document.addEventListener("paste", e => this.onPaste(e));
      // 칸에 들어가면 값 전체를 골라 바로 덮어쓸 수 있게
      document.addEventListener("focusin", e => {
        const el = e.target;
        if (el.closest && el.closest("table.grid") && el.tagName === "INPUT" && /^(text|number)$/.test(el.type)) setTimeout(() => { try { el.select(); } catch (er) { /* 이미 다시 그려짐 */ } }, 0);
      });
      document.addEventListener("change", e => this.onNewRow(e), true);
    },
    table: el => (el && el.closest ? el.closest("table.grid") : null),
    rows: tbl => [...tbl.tBodies[0].rows].filter(r => r.dataset.row !== undefined || r.classList.contains("new-row")),
    ctrls: tr => [...tr.querySelectorAll("input[data-bind], select[data-bind], input[data-new]")],
    cols: tbl => JSON.parse(tbl.dataset.cols),
    pos(el) {
      const tbl = this.table(el), tr = el.closest("tr");
      return { tbl, r: this.rows(tbl).indexOf(tr), c: this.ctrls(tr).indexOf(el) };
    },
    focusAt(tbl, r, c) {
      const rows = this.rows(tbl);
      if (r < 0 || r >= rows.length) return false;
      const cs = this.ctrls(rows[r]);
      if (!cs.length) return false;
      cs[Math.max(0, Math.min(c, cs.length - 1))].focus();
      return true;
    },

    onKey(e) {
      const el = e.target, tbl = this.table(el);
      if (!tbl || !(el.matches("input, select"))) return;
      const { r, c } = this.pos(el);
      if (r < 0 || c < 0) return;
      const caretAt = (end) => { try { return el.selectionStart === el.selectionEnd && el.selectionStart === (end ? el.value.length : 0); } catch (er) { return true; } };
      const textLike = el.tagName === "INPUT" && /^(text|number)$/.test(el.type);
      let to = null;
      if (e.key === "ArrowDown" || (e.key === "Enter" && !e.shiftKey)) to = [r + 1, c];
      else if (e.key === "ArrowUp" || (e.key === "Enter" && e.shiftKey)) to = [r - 1, c];
      else if (e.key === "ArrowLeft" && (el.type === "checkbox" || el.tagName === "SELECT" || (textLike && (el.type === "number" || caretAt(false))))) to = [r, c - 1];
      else if (e.key === "ArrowRight" && (el.type === "checkbox" || el.tagName === "SELECT" || (textLike && (el.type === "number" || caretAt(true))))) to = [r, c + 1];
      else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "r") { e.preventDefault(); e.stopPropagation(); return this.fillRight(el); }
      if (!to) return;
      e.preventDefault();
      e.stopPropagation();
      if (e.key === "Enter" && el.dataset.new) { el.blur(); return; }   // 새 줄: 값이 있으면 change 로 추가됨
      if (!this.focusAt(tbl, to[0], to[1]) && e.key === "Enter") el.blur();   // 마지막 줄에서 Enter = 확정만
    },

    // 새 줄의 이름 칸: 값을 넣으면 한 줄 추가하고 그 줄의 다음 칸으로
    onNewRow(e) {
      const el = e.target;
      if (!el.dataset || !el.dataset.new) return;
      e.stopPropagation();
      const name = el.value.trim();
      if (!name) return;
      const list = el.dataset.new;
      this.h.pushUndo();
      const i = this.h.newRow(list);
      this.h.setValue(`${list}.${i}.name`, name);
      this.h.render(`${list}.${i}.role`);
    },

    // 오른쪽 채우기: 이 달 MM 을 뒤의 모든 달에 (Ctrl+R)
    fillRight(el) {
      const tbl = this.table(el), { r, c } = this.pos(el), cols = this.cols(tbl), list = tbl.dataset.list;
      if (!cols[c] || !/^mm\./.test(cols[c].b) || r >= this.h.proj()[list].length) return;
      const v = parse("num", el.value);
      this.h.pushUndo();
      for (let k = c; k < cols.length; k++) if (/^mm\./.test(cols[k].b)) this.h.setValue(`${list}.${r}.${cols[k].b}`, v);
      this.h.render(el.dataset.bind);
      PC.UI.toast(`${cols[c].l} 부터 끝 달까지 ${v} MM 으로 채웠습니다.`);
    },

    // 엑셀에서 복사한 영역 붙여넣기: 고른 칸부터 (머리글째 복사했으면 머리글 이름으로 열을 맞춤), 모자라는 줄은 추가
    onPaste(e) {
      const el = e.target, tbl = this.table(el);
      if (!tbl) return;
      const text = (e.clipboardData || window.clipboardData).getData("text/plain") || "";
      if (!/[\t\n\r]/.test(text.replace(/[\r\n]+$/, ""))) return;   // 한 칸 값은 브라우저가 그대로 붙여넣음
      e.preventDefault();
      this.paste(tbl, el, parseTsv(text));
    },
    paste(tbl, el, data) {
      const cols = this.cols(tbl), list = tbl.dataset.list, { r, c } = this.pos(el);
      let colIdx = null, lines = data;
      const hm = headerMap(data[0] || [], cols);
      if (hm) { colIdx = hm; lines = data.slice(1); }
      if (!lines.length) return;
      this.h.pushUndo();
      const proj = this.h.proj();
      let cells = 0, added = 0, unknown = 0;
      lines.forEach((line, i) => {
        const row = r + i;
        if (row >= proj[list].length) { this.h.newRow(list); added++; }
        line.forEach((raw, j) => {
          const ci = colIdx ? colIdx[j] : c + j;
          const col = cols[ci];
          if (ci === undefined || ci < 0 || !col) return;
          const v = parse(col.t, raw);
          this.h.setValue(`${list}.${row}.${col.b}`, v);
          cells++;
          if (col.b === "grade" || col.b === "vendor") unknown += this.h.known(list, col.b, v) ? 0 : 1;
        });
      });
      this.h.render(`${list}.${r}.${cols[colIdx ? Math.max(0, colIdx.find(x => x >= 0)) : c].b}`);
      PC.UI.toast(`${lines.length}줄 · ${cells}칸을 붙여 넣었습니다${added ? ` (새 인력 ${added}명)` : ""}${unknown ? ` — 기준정보에 없는 등급 · 업체 ${unknown}칸은 빨갛게 표시` : ""}. Ctrl+Z 로 되돌릴 수 있습니다.`, unknown ? "error" : "ok");
    },
    _test: { parseTsv, headerMap, parse }
  };

  PC.Grid = Grid;
})(window);
