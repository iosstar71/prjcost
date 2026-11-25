// =============================================================================
// 공통 도구 — 숫자 · 반올림 · 월 계산 · 화면 문자열 이스케이프
//  브라우저(<script>)와 Node 시험(require) 양쪽에서 쓰도록 전역 PC 에 붙임 (file:// 에서 ES 모듈은 막히므로 일반 스크립트)
// =============================================================================
(function (root) {
  "use strict";

  // 숫자로: 빈 값 · 숫자가 아닌 값은 0 (원본 파이썬의 float(x or 0) 와 같은 결과)
  function num(v) {
    if (v === null || v === undefined || v === "" || v === false) return 0;
    const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : 0;
  }

  // 값이 비었는지 (원본의 `x is not None and str(x).strip() != ""` 의 반대)
  function isBlank(v) { return v === null || v === undefined || String(v).trim() === ""; }

  // 파이썬 round(x, d) 와 같은 반올림 — 정확히 한가운데(…5)일 때만 짝수 쪽으로 (은행가 반올림)
  //  금액 계산 결과가 파이썬 원본과 한 원 단위까지 같도록 (시험: tests/run.js 가 원본 계산 결과와 비교)
  function pyRound(x, d) {
    d = d === undefined ? 2 : d;
    if (!Number.isFinite(x)) return x;
    if (x === 0) return 0;
    const exact = Math.abs(x).toFixed(Math.min(100, d + 60));   // 2진 값의 정확한 10진 전개 (금액 규모에서는 60자리면 끝남)
    const tail = exact.slice(exact.indexOf(".") + 1 + d);
    if (/^50*$/.test(tail)) {   // 정확히 한가운데 → 짝수 쪽
      const base = Math.abs(x).toFixed(d + 1).slice(0, -1);       // 버림한 값 (문자열)
      const lastDigit = Number(base.replace(".", "").slice(-1));
      const down = Number(base);
      const up = Number((down + Math.pow(10, -d)).toFixed(d));
      const r = lastDigit % 2 === 0 ? down : up;
      return x < 0 ? -r : r;
    }
    const r = Number(Math.abs(x).toFixed(d));
    return x < 0 ? -r : r;
  }

  // 'YYYY-MM-DD' · 'YYYY-MM' → 'YYYY-MM' (아니면 fallback)
  function normalizeYm(v, fallback) {
    const s = String(v === null || v === undefined ? "" : v).trim();
    return s.length >= 7 && s[4] === "-" ? s.slice(0, 7) : fallback;
  }

  // 시작 ~ 종료 월 목록 (최대 121개 — 원본과 같음)
  function monthRange(startVal, endVal) {
    const s = normalizeYm(startVal, "2026-01"), e = normalizeYm(endVal, "2026-06");
    const sy = Number(s.slice(0, 4)), sm = Number(s.slice(5, 7)), ey = Number(e.slice(0, 4)), em = Number(e.slice(5, 7));
    if (![sy, sm, ey, em].every(Number.isInteger) || sm < 1 || sm > 12 || em < 1 || em > 12) return ["2026-01", "2026-02", "2026-03"];
    if (sy * 12 + sm > ey * 12 + em) return [s];
    const out = [];
    let y = sy, m = sm;
    while (y * 12 + m <= ey * 12 + em) {
      out.push(`${String(y).padStart(4, "0")}-${String(m).padStart(2, "0")}`);
      if (m === 12) { y += 1; m = 1; } else m += 1;
      if (out.length > 120) break;
    }
    return out;
  }

  // 화면 표시
  const esc = s => String(s === null || s === undefined ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const won = v => Math.round(num(v)).toLocaleString("ko-KR");
  const man = v => (Math.round(num(v) / 10000)).toLocaleString("ko-KR");   // 만 원 단위
  const pct = v => `${num(v).toFixed(1)}%`;
  const clone = o => JSON.parse(JSON.stringify(o));
  const stamp = (d = new Date()) => {
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  };
  const nowText = (d = new Date()) => {
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  };
  // 파일 이름에 쓸 수 있는 ID (원본과 같은 규칙: 영문 · 숫자 · 한글 · - _ 만)
  const safeId = id => String(id || "PRJ-NEW").split("").filter(c => /[\p{L}\p{N}_-]/u.test(c)).join("") || "PRJ-NEW";

  const PC = root.PC = Object.assign(root.PC || {}, { num, isBlank, pyRound, normalizeYm, monthRange, esc, won, man, pct, clone, stamp, nowText, safeId });
  if (typeof module !== "undefined" && module.exports) module.exports = PC;
})(typeof window !== "undefined" ? window : globalThis);
