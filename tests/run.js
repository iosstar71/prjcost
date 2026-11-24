// =============================================================================
// 단위 시험 (Node, 1초) — node tests/run.js
//  1) 계산 회귀: 저장해 둔 기준 결과(tests/golden.json, 42가지 경우)와 1원 단위까지 같은지 — 규칙을 일부러 바꾸면 node tools/make_golden.js
//  2) 반올림: 한가운데는 짝수 쪽 (v4 파일과 같은 금액이 나오도록)
//  3) 엑셀: 만든 xlsx 를 다시 읽으면 같은 자료인지 (큰 JSON 은 여러 칸으로 나눠 저장)
// =============================================================================
"use strict";
const fs = require("fs"), path = require("path");
const PC = require("../js/util.js");
require("../js/calc.js");
require("../js/xlsx.js");
require("../js/model.js");

let pass = 0, fail = 0;
const eq = (name, got, want) => {
  const a = JSON.stringify(got), b = JSON.stringify(want);
  if (a === b) pass++;
  else { fail++; console.log(`  [FAIL] ${name}\n     got  ${a.slice(0, 300)}\n     want ${b.slice(0, 300)}`); }
};

// 숫자는 값으로 비교
function deepDiff(a, b, p = "") {
  if (typeof a === "number" && typeof b === "number") return a === b ? null : `${p}: ${a} != ${b}`;
  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) return `${p}: length ${a.length} != ${b.length}`;
    for (let i = 0; i < a.length; i++) { const d = deepDiff(a[i], b[i], `${p}[${i}]`); if (d) return d; }
    return null;
  }
  if (a && b && typeof a === "object" && typeof b === "object") {
    const ka = Object.keys(a).sort(), kb = Object.keys(b).sort();
    if (ka.join() !== kb.join()) return `${p}: keys ${ka} != ${kb}`;
    for (const k of ka) { const d = deepDiff(a[k], b[k], `${p}.${k}`); if (d) return d; }
    return null;
  }
  return a === b ? null : `${p}: ${JSON.stringify(a)} != ${JSON.stringify(b)}`;
}

// 1) 계산
const golden = JSON.parse(fs.readFileSync(path.join(__dirname, "golden.json"), "utf8"));
golden.forEach((g, i) => {
  const got = PC.Calc.calculateProjectCost(JSON.parse(JSON.stringify(g.input)), g.fallback);
  delete got.computed_at;
  const d = deepDiff(got, g.expected);
  eq(`calc case ${i} (${g.input.project_id}) matches the saved baseline${d ? " — " + d : ""}`, !d, true);
});

// 2) 반올림 (파이썬 round 값)
[[0.125, 0.12], [0.375, 0.38], [2.675, 2.67], [1.005, 1.0], [-0.125, -0.12], [12345.675, 12345.67], [0.5, 0.5], [1234567.885, 1234567.89], [7.5, 7.5]]
  .forEach(([x, want]) => eq(`pyRound(${x}, 2)`, PC.pyRound(x, 2), want));
eq("pyRound(2.5, 0) → 2 (even)", PC.pyRound(2.5, 0), 2);
eq("pyRound(3.5, 0) → 4 (even)", PC.pyRound(3.5, 0), 4);
eq("monthRange crosses the year", PC.monthRange("2026-11-15", "2027-02"), ["2026-11", "2026-12", "2027-01", "2027-02"]);
eq("monthRange reversed → start month only", PC.monthRange("2026-05", "2026-02"), ["2026-05"]);

// 3) 엑셀 쓰기 → 읽기
(async () => {
  const X = PC.Xlsx;
  const big = { text: "가".repeat(80000), n: 1.5 };
  const blob = X.build([
    { name: "01_정보", cols: [20, 30], rows: [{ cells: ["항목", "값"], style: "head" }, { cells: ["금액", 1234567.5] }, { cells: ["설명", "줄1\n줄2 <&>\"'"] }] },
    { name: "State_JSON", hidden: true, rows: X.jsonRows(big) }
  ]);
  const wb = await X.read(blob);
  eq("xlsx: sheet names", wb.sheets.map(s => s.name), ["01_정보", "State_JSON"]);
  eq("xlsx: cells come back (number · text with line break and XML characters)", wb.sheet("01_정보").rows, [["항목", "값"], ["금액", 1234567.5], ["설명", "줄1\n줄2 <&>\"'"]]);
  eq("xlsx: JSON longer than one Excel cell (32,767) is split and joined back", X.readJson(wb.sheet("State_JSON")), big);
  eq("xlsx: hidden sheet flag", wb.sheet("State_JSON").hidden, true);

  // 기본 프로젝트 · 데모: 계산이 돌고 이익률이 나옴
  const M = PC.Model;
  const demo = M.demoProject(M.defaultSettings());
  const c = PC.Calc.calculateProjectCost(demo, M.defaultSettings());
  eq("demo project calculates (6 months, revenue 240,000,000)", [c.months.length, c.summary.revenue], [6, 240000000]);
  eq("normalize fills missing parts of an old project", Object.keys(M.normalize({ project_id: "A" }, M.defaultSettings())).sort(), Object.keys(M.newProject(M.defaultSettings())).sort());

  // 투입 현황: 영역별 · 등급별, 분할계약 기간만
  const st = PC.Calc.staffStats(c, { by: "area" });
  eq("staff stats by area: groups, MM and people (demo)", st.rows.map(r => [r.key, r.total_mm, r.people]), [["PMO", 6, 1], ["분석 · 설계", 4.5, 1], ["개발", 4, 1]]);
  eq("staff stats total = labor cost", st.total.cost, c.summary.breakdown.internal_labor + c.summary.breakdown.external_labor);
  const ph1 = PC.Calc.staffStats(c, { by: "kind", months: ["2026-01", "2026-02", "2026-03"] });
  eq("staff stats limited to a phase's months (internal / external MM)", ph1.rows.map(r => [r.key, r.total_mm]), [["내부", 5.5], ["외주", 2]]);
  eq("staff stats internal only by grade", PC.Calc.staffStats(c, { by: "grade", scope: "internal" }).rows.map(r => r.key), ["수석", "선임"]);

  // 단가 인상 · 승진: 그 달부터 끝까지 새 단가. 내부는 인상과 승진을 따로 (같은 달이면 승진), 외주는 협상 단가
  const rateOf = (row, m) => row.calc[m].rate;
  const kim = c.details.internal_labor[0], ext = c.details.external_labor[0];
  eq("internal: raise from 03, promotion from 05", ["2026-02", "2026-03", "2026-05"].map(m => rateOf(kim, m)), [8500000, 8800000, 9500000]);
  eq("external: negotiated rate from 04", ["2026-03", "2026-04"].map(m => rateOf(ext, m)), [7800000, 8000000]);
  const st2 = M.defaultSettings();
  const r = (rc, pr) => PC.Calc.resolveInternalRate({ grade: "선임", base_rate: 6000000, rate_changes: rc, promotions: pr }, "2026-06", st2);
  eq("promotion with grade only uses the grade rate", r([], [{ effective_ym: "2026-06", grade: "책임", rate: "" }]), 7200000);
  eq("same month: promotion wins over raise", r([{ effective_ym: "2026-06", rate: 6500000 }], [{ effective_ym: "2026-06", rate: 7000000 }]), 7000000);
  eq("later raise after promotion applies", r([{ effective_ym: "2026-06", rate: 7300000 }], [{ effective_ym: "2026-04", rate: 7000000 }]), 7300000);
  const nz = M.normalize({ internal_labor: [{ name: "a", promotions: [{ effective_ym: "2026-03", rate: "" }] }], external_labor: [{ name: "b" }] }, st2);
  eq("normalize drops empty promotions and adds rate_changes", [nz.internal_labor[0].promotions, nz.internal_labor[0].rate_changes, nz.external_labor[0].rate_changes], [[], [], []]);
  console.log(`\nUnit: ${pass} / ${pass + fail} passed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
