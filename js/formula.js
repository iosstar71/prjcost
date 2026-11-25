// =============================================================================
// 엑셀 수식 — 읽기 · 칸 참조 옮기기 · 계산 (사업검토서 양식용, 외부 라이브러리 없음)
//  양식의 합계 칸은 수식(=SUM(O5:O6), =원가상세!O7 …)이라 출력 파일을 Excel 에서 고치면 합계가 따라 바뀜.
//  반복 줄로 줄이 늘면 수식의 칸 참조도 맞춰 옮기고(shift), 출력 전에 수식을 직접 계산해 화면 값과 맞춰 봄(evaluate).
//  계산하는 함수는 양식에 쓰는 정도만 (FUNCS). 모르는 함수 · 이름은 "계산 못 함"으로 알림 — Excel 에서는 그대로 계산됨
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");

  // ---------------- 낱말 나누기 ----------------
  const CELLREF = /(\$?)([A-Z]{1,3})(\$?)(\d+)/y;
  const SHEET = /(?:'((?:[^']|'')+)'|([^\s'!(),;:+\-*/^&=<>"{}%]+))!/y;
  const IDENT = /[A-Za-z_\u0080-￿][A-Za-z0-9_.\u0080-￿]*/y;
  const NUM = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
  const ERR = /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)/y;
  const at = (re, s, i) => { re.lastIndex = i; return re.exec(s); };
  const colNo = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const colName = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
  const cell = m => ({ ca: !!m[1], c: colNo(m[2]), ra: !!m[3], r: Number(m[4]) });
  const isIdentCh = ch => !!ch && /[A-Za-z0-9_.(\u0080-￿]/.test(ch);

  // 칸 참조 하나를 읽음 (시트!A1, A1:B2) → { sheet, a, b, end } | null
  function readRef(s, i) {
    let sheet = null, j = i;
    const sm = at(SHEET, s, i);
    if (sm) { sheet = sm[1] !== undefined ? sm[1].replace(/''/g, "'") : sm[2]; j = SHEET.lastIndex; }
    const a = at(CELLREF, s, j);
    if (!a) return null;
    j = CELLREF.lastIndex;
    let b = null;
    if (s[j] === ":") { const m = at(CELLREF, s, j + 1); if (m) { b = m; j = CELLREF.lastIndex; } }
    if (isIdentCh(s[j])) return null;   // LOG10( · A1B 같은 이름
    return { sheet, sheetText: sm ? sm[0] : "", a: cell(a), b: b ? cell(b) : null, end: j };
  }

  function tokenize(s) {
    const out = [];
    let i = 0;
    while (i < s.length) {
      const ch = s[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        let j = i + 1, v = "";
        for (; j < s.length; j++) { if (s[j] === '"') { if (s[j + 1] === '"') { v += '"'; j++; } else break; } else v += s[j]; }
        out.push({ t: "str", v, s: i, e: j + 1 }); i = j + 1; continue;
      }
      if (ch === "#") { const m = at(ERR, s, i); if (m) { out.push({ t: "err", v: m[0], s: i, e: ERR.lastIndex }); i = ERR.lastIndex; continue; } }
      const prev = out[out.length - 1];
      if (!prev || prev.t === "op" || prev.t === "(" || prev.t === ",") {
        const rf = readRef(s, i);
        if (rf) { out.push(Object.assign({ t: "ref", s: i, e: rf.end }, rf)); i = rf.end; continue; }
      }
      if (/[0-9.]/.test(ch)) { const m = at(NUM, s, i); if (m) { out.push({ t: "num", v: Number(m[0]), s: i, e: NUM.lastIndex }); i = NUM.lastIndex; continue; } }
      const id = at(IDENT, s, i);
      if (id) {
        const e = IDENT.lastIndex, name = id[0].toUpperCase().replace(/^_XLFN\.|^_XLWS\./, "");
        if (s[e] === "(") out.push({ t: "fn", v: name, s: i, e });
        else if (name === "TRUE" || name === "FALSE") out.push({ t: "bool", v: name === "TRUE", s: i, e });
        else out.push({ t: "name", v: id[0], s: i, e });
        i = e; continue;
      }
      const two = s.slice(i, i + 2);
      if (two === "<=" || two === ">=" || two === "<>") { out.push({ t: "op", v: two, s: i, e: i + 2 }); i += 2; continue; }
      if ("+-*/^&=<>%".includes(ch)) { out.push({ t: "op", v: ch, s: i, e: i + 1 }); i++; continue; }
      if (ch === "(" || ch === ")" || ch === ",") { out.push({ t: ch, s: i, e: i + 1 }); i++; continue; }
      out.push({ t: "bad", v: ch, s: i, e: i + 1 }); i++;   // { 배열 상수 등 — 계산 못 함 (옮기기는 그대로)
    }
    return out;
  }

  // ---------------- 칸 참조 바꾸기 (글자는 그대로, 참조만) ----------------
  const cellText = x => `${x.ca ? "$" : ""}${colName(x.c)}${x.ra ? "$" : ""}${x.r}`;
  // fn(ref) → { a, b } (바뀐 참조) | null (그대로)
  function mapRefs(f, fn) {
    let out = "", last = 0;
    tokenize(f).forEach(t => {
      if (t.t !== "ref") return;
      const n = fn(t);
      if (!n) return;
      out += f.slice(last, t.s) + t.sheetText + cellText(n.a) + (n.b ? ":" + cellText(n.b) : "");
      last = t.e;
    });
    return out + f.slice(last);
  }
  // 공유 수식(Excel 이 같은 수식을 끌어 채운 칸)을 칸마다 풀 때: 상대 참조를 (dr, dc) 만큼
  const offset = (f, dr, dc) => mapRefs(f, t => {
    const mv = x => x && { ca: x.ca, ra: x.ra, c: x.ca ? x.c : x.c + dc, r: x.ra ? x.r : x.r + dr };
    return { a: mv(t.a), b: mv(t.b) };
  });

  // ---------------- 읽기 (식 → 나무) ----------------
  function parse(f) {
    const ts = tokenize(f);
    let i = 0;
    const peek = () => ts[i], next = () => ts[i++];
    const isOp = (...v) => peek() && peek().t === "op" && v.includes(peek().v);
    const fail = msg => { throw new Unsupported(msg); };
    const bin = (sub, ops) => () => { let l = sub(); while (isOp(...ops)) { const op = next().v; l = { k: "bin", op, l, r: sub() }; } return l; };
    const primary = () => {
      const t = next();
      if (!t) fail("수식이 끝남");
      if (t.t === "num") return { k: "v", v: t.v };
      if (t.t === "str") return { k: "v", v: t.v };
      if (t.t === "bool") return { k: "v", v: t.v };
      if (t.t === "err") return { k: "v", v: new XErr(t.v) };
      if (t.t === "ref") return { k: "ref", sheet: t.sheet, a: t.a, b: t.b };
      if (t.t === "(") { const x = expr(); if (!peek() || next().t !== ")") fail("괄호"); return x; }
      if (t.t === "fn") {
        next();   // (
        const args = [];
        if (peek() && peek().t === ")") { next(); return { k: "fn", name: t.v, args }; }
        for (;;) {
          args.push(peek() && (peek().t === "," || peek().t === ")") ? { k: "v", v: null } : expr());
          const sep = next();
          if (!sep) fail("괄호");
          if (sep.t === ")") break;
          if (sep.t !== ",") fail("함수 인수");
        }
        return { k: "fn", name: t.v, args };
      }
      if (t.t === "op" && (t.v === "-" || t.v === "+")) { const x = postfix();   // -2^2 = 4 (Excel: 부호가 거듭제곱보다 먼저)
        return t.v === "-" ? { k: "neg", x } : x; }
      fail(t.t === "name" ? `이름 ${t.v}` : `기호 ${t.v || t.t}`);
    };
    const postfix = () => { let x = primary(); while (isOp("%")) { next(); x = { k: "pct", x }; } return x; };
    const power = bin(postfix, ["^"]);
    const mul = bin(power, ["*", "/"]);
    const add = bin(mul, ["+", "-"]);
    const cat = bin(add, ["&"]);
    const cmp = bin(cat, ["=", "<>", "<", ">", "<=", ">="]);
    const expr = cmp;
    const x = expr();
    if (i < ts.length) fail(`기호 ${ts[i].v || ts[i].t}`);
    return x;
  }

  // ---------------- 계산 ----------------
  class XErr { constructor(v) { this.v = v; } toString() { return this.v; } }
  class Unsupported extends Error {}
  const isErr = v => v instanceof XErr;
  const blank = v => v === null || v === undefined;
  function toNum(v) {
    if (isErr(v)) throw v;
    if (blank(v)) return 0;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (typeof v === "number") return v;
    const n = Number(String(v).replace(/,/g, ""));
    if (String(v).trim() === "" || !Number.isFinite(n)) throw new XErr("#VALUE!");   // ="" + 1 → #VALUE! (Excel 과 같게)
    return n;
  }
  const toStr = v => { if (isErr(v)) throw v; return blank(v) ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v); };
  const toBool = v => { if (isErr(v)) throw v; if (typeof v === "string") { if (/^true$/i.test(v)) return true; if (/^false$/i.test(v)) return false; throw new XErr("#VALUE!"); } return !!toNum(v); };
  // 비교: 숫자 < 글자 < 참거짓 (Excel 과 같은 순서), 글자는 대소문자 무시
  const rank = v => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  function compare(a, b) {
    if (blank(a)) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
    if (blank(b)) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === "string") { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }
  // COUNTIF · SUMIF 조건: ">0", "<>", "AI예비비", 5 …
  function criterion(c) {
    if (typeof c === "number" || typeof c === "boolean") return v => !blank(v) && compare(v, c) === 0;
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(toStr(c)), op = m[1] || "=", raw = m[2];
    const n = raw.trim() !== "" && Number.isFinite(Number(raw)) ? Number(raw) : null;
    const target = n !== null ? n : raw;
    return v => {
      if (isErr(v)) return false;
      if (raw === "") return op === "=" ? blank(v) || v === "" : op === "<>" ? !(blank(v) || v === "") : false;
      if (blank(v)) return op === "<>";
      if (n !== null && typeof v !== "number") return op === "<>";
      if (n === null && typeof v === "number") return op === "<>";
      const d = compare(v, target);
      return { "=": d === 0, "<>": d !== 0, "<": d < 0, ">": d > 0, "<=": d <= 0, ">=": d >= 0 }[op];
    };
  }
  const round = (x, d, mode) => {
    const p = Math.pow(10, d), y = Math.abs(x) * p, s = Math.sign(x);
    const r = mode === "up" ? Math.ceil(y - 1e-9) : mode === "down" ? Math.floor(y + 1e-9) : Math.round(y + 1e-9);
    return (s * r) / p;
  };
  // 범위 안의 숫자만 (글자 · 빈칸은 건너뜀), 직접 적은 값은 숫자로 바꿈 — SUM · MIN · MAX · AVERAGE
  const nums = (args, ev) => {
    const out = [];
    args.forEach(a => {
      if (a.k === "ref") ev.range(a).forEach(v => { if (isErr(v)) throw v; if (typeof v === "number") out.push(v); });
      else out.push(toNum(ev.node(a)));
    });
    return out;
  };
  const FUNCS = {
    SUM: (args, ev) => nums(args, ev).reduce((a, b) => a + b, 0),
    MIN: (args, ev) => { const n = nums(args, ev); return n.length ? Math.min(...n) : 0; },
    MAX: (args, ev) => { const n = nums(args, ev); return n.length ? Math.max(...n) : 0; },
    AVERAGE: (args, ev) => { const n = nums(args, ev); if (!n.length) throw new XErr("#DIV/0!"); return n.reduce((a, b) => a + b, 0) / n.length; },
    COUNT: (args, ev) => args.reduce((a, x) => a + (x.k === "ref" ? ev.range(x).filter(v => typeof v === "number").length : typeof ev.node(x) === "number" ? 1 : 0), 0),
    COUNTA: (args, ev) => args.reduce((a, x) => a + (x.k === "ref" ? ev.range(x).filter(v => !blank(v)).length : blank(ev.node(x)) ? 0 : 1), 0),
    COUNTIF: (args, ev) => { const ok = criterion(ev.node(args[1])); return ev.range(args[0]).filter(ok).length; },
    SUMIF: (args, ev) => {
      const ok = criterion(ev.node(args[1])), keys = ev.range(args[0]), vals = args[2] ? ev.range(args[2]) : keys;
      return keys.reduce((a, v, i) => a + (ok(v) && typeof vals[i] === "number" ? vals[i] : 0), 0);
    },
    IF: (args, ev) => (toBool(ev.node(args[0])) ? (args[1] ? ev.node(args[1]) : true) : args[2] ? ev.node(args[2]) : false),
    IFERROR: (args, ev) => { try { const v = ev.node(args[0]); return isErr(v) ? ev.node(args[1]) : v; } catch (e) { if (isErr(e)) return ev.node(args[1]); throw e; } },
    ROUND: (args, ev) => round(toNum(ev.node(args[0])), args[1] ? toNum(ev.node(args[1])) : 0),
    ROUNDUP: (args, ev) => round(toNum(ev.node(args[0])), args[1] ? toNum(ev.node(args[1])) : 0, "up"),
    ROUNDDOWN: (args, ev) => round(toNum(ev.node(args[0])), args[1] ? toNum(ev.node(args[1])) : 0, "down"),
    INT: (args, ev) => Math.floor(toNum(ev.node(args[0]))),
    ABS: (args, ev) => Math.abs(toNum(ev.node(args[0]))),
    AND: (args, ev) => args.every(a => toBool(ev.node(a))),
    OR: (args, ev) => args.some(a => toBool(ev.node(a))),
    NOT: (args, ev) => !toBool(ev.node(args[0]))
  };

  // 통합 문서 계산. sheets: { 시트 이름: { "A1": { v, f } } } (v: 숫자 · 글자 · 참거짓 · null, f: 수식 글자 "=" 없이)
  // → { values: { 시트: { 칸: 값 } } (수식 칸만), problems: [{ sheet, ref, f, reason }] (계산 못 한 칸) }
  function evaluate(sheets) {
    const memo = new Map(), busy = new Set(), problems = [], values = {};
    const lower = {};
    Object.keys(sheets).forEach(n => { lower[n.toLowerCase()] = n; });
    const sheetOf = (name, here) => (name === null ? here : lower[String(name).toLowerCase()]);
    function cellValue(sheet, ref) {
      const c = (sheets[sheet] || {})[ref];
      if (!c) return null;
      if (c.f === undefined) return c.v;
      const key = `${sheet}!${ref}`;
      if (memo.has(key)) { const m = memo.get(key); if (m.problem) throw new Unsupported(m.problem); return m.v; }
      if (busy.has(key)) throw new XErr("#CIRC");
      busy.add(key);
      let v;
      try {
        const ast = parse(c.f);
        v = node(ast, sheet);
        if (v && v.k === "ref") v = single(v, sheet);
      } catch (e) {
        busy.delete(key);
        if (isErr(e)) v = e;
        else { const reason = e instanceof Unsupported ? e.message : String(e.message || e); memo.set(key, { problem: reason }); throw new Unsupported(reason); }
      }
      busy.delete(key);
      memo.set(key, { v });
      return v;
    }
    const cellsOf = (r, here) => {
      const sh = sheetOf(r.sheet, here);
      if (!sh) throw new XErr("#REF!");
      const b = r.b || r.a, out = [];
      for (let row = Math.min(r.a.r, b.r); row <= Math.max(r.a.r, b.r); row++)
        for (let col = Math.min(r.a.c, b.c); col <= Math.max(r.a.c, b.c); col++) out.push(cellValue(sh, colName(col) + row));
      return out;
    };
    const single = (r, here) => { if (r.b && (r.b.r !== r.a.r || r.b.c !== r.a.c)) throw new XErr("#VALUE!"); return cellsOf(r, here)[0]; };
    function node(n, here) {
      const ev = { node: x => { const v = node(x, here); return v && v.k === "ref" ? single(v, here) : v; }, range: x => (x.k === "ref" ? cellsOf(x, here) : [ev.node(x)]) };
      switch (n.k) {
        case "v": return n.v;
        case "ref": return n;   // 함수 인수에서 범위로 쓰일 수 있어 여기서는 풀지 않음
        case "neg": return -toNum(ev.node(n.x));
        case "pct": return toNum(ev.node(n.x)) / 100;
        case "bin": {
          const a = ev.node(n.l), b = ev.node(n.r);
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          switch (n.op) {
            case "+": return toNum(a) + toNum(b);
            case "-": return toNum(a) - toNum(b);
            case "*": return toNum(a) * toNum(b);
            case "/": { const d = toNum(b); if (d === 0) return new XErr("#DIV/0!"); return toNum(a) / d; }
            case "^": return Math.pow(toNum(a), toNum(b));
            case "&": return toStr(a) + toStr(b);
            default: { const d = compare(a, b); return { "=": d === 0, "<>": d !== 0, "<": d < 0, ">": d > 0, "<=": d <= 0, ">=": d >= 0 }[n.op]; }
          }
        }
        case "fn": {
          const fn = FUNCS[n.name];
          if (!fn) throw new Unsupported(`함수 ${n.name}`);
          try { return fn(n.args, ev); } catch (e) { if (isErr(e)) return e; throw e; }
        }
      }
      throw new Unsupported("알 수 없는 식");
    }
    Object.keys(sheets).forEach(sh => Object.keys(sheets[sh]).forEach(ref => {
      const c = sheets[sh][ref];
      if (c.f === undefined) return;
      try { (values[sh] = values[sh] || {})[ref] = cellValue(sh, ref); }
      catch (e) { problems.push({ sheet: sh, ref, f: c.f, reason: e instanceof Unsupported ? e.message : String(e.message || e) }); }
    }));
    return { values, problems };
  }

  PC.Formula = { tokenize, parse, mapRefs, offset, evaluate, isErr, colNo, colName, FUNCS };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Formula;
})(typeof window !== "undefined" ? window : globalThis);
