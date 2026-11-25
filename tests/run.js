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

  // 단가 인상(기준정보: 내부 등급 인상단가 + 연봉 인상월, 외주업체 인상월 · 인상단가) · 승진(사람마다)
  const rateOf = (row, m) => row.calc[m].rate;
  const kim = c.details.internal_labor[0], lee = c.details.internal_labor[1], ext = c.details.external_labor[0];
  eq("demo: raise month 03 within the project", PC.Calc.raiseMonth(demo.settings, c.months), "2026-03");
  eq("internal: grade raise from 03 (also over a personal base rate)", ["2026-02", "2026-03", "2026-05"].map(m => rateOf(kim, m)), [8500000, 8800000, 8800000]);
  eq("internal: other grade raise from 03", ["2026-02", "2026-03"].map(m => rateOf(lee, m)), [6000000, 6200000]);
  eq("external: vendor rate raised from the negotiated month", ["2026-03", "2026-04"].map(m => rateOf(ext, m)), [7800000, 8000000]);
  const st2 = M.defaultSettings();
  st2.internal_grades.find(g => g.grade === "책임").raise_rate = 7500000;
  const r = (g, ym, raise) => PC.Calc.resolveInternalRate({ grade: g, base_rate: "" }, ym, st2, raise);
  eq("grade rate, raised from the raise month", [r("책임", "2026-02", "2026-03"), r("책임", "2026-04", "2026-03")], [7200000, 7500000]);
  eq("grade without a raise rate keeps its rate", r("선임", "2026-06", "2026-03"), 6000000);
  eq("no raise month in the project period = no raise", PC.Calc.raiseMonth({ raise_month: "11" }, ["2026-01", "2026-02"]), "");
  const nz = M.normalize({ settings: M.defaultSettings(), internal_labor: [{ name: "a", promotions: [{ effective_ym: "2026-03", rate: "" }], rate_changes: [{ effective_ym: "2026-04", rate: 100 }] }],
    external_labor: [{ name: "b", vendor: "파트너A", grade: "고급", rate_changes: [{ effective_ym: "2026-05", rate: 200 }] }] }, st2);
  const va = nz.settings.vendor_rates.find(v => v.vendor === "파트너A" && v.grade === "고급");
  eq("normalize: promotions and per-person raises dropped (noted on open); external raise moves to the vendor rate", [nz.internal_labor[0].promotions, nz.internal_labor[0].rate_changes, nz.external_labor[0].rate_changes, va.raise_ym, va.raise_rate], [undefined, undefined, undefined, "2026-05", 200]);
  eq("legacy note mentions the dropped promotion", /승진 단가\(a 2026-04 100원\)/.test(M.legacyNote({ internal_labor: [{ name: "a", promotions: [{ effective_ym: "2026-04", rate: 100 }] }] })), true);
  eq("normalize: total MM = sum of the monthly plan", M.normalize({ start_date: "2026-01", end_date: "2026-03", internal_labor: [{ name: "x", target_mm: 9, mm: { "2026-01": 1, "2026-02": 0.25, "2026-09": 5 } }] }, st2).internal_labor[0].target_mm, 1.25);

  // 기준정보 순서: 투입 현황 · 원가 구성 항목 · 줄 세우기
  const ost = M.defaultSettings(); ost.areas = ["개발", "PMO", "분석 · 설계"];
  eq("staff stats follow the settings order (areas)", PC.Calc.staffStats(c, { by: "area", settings: ost }).rows.map(x => x.key), ["개발", "PMO", "분석 · 설계"]);
  eq("staff stats by grade: internal grades then external grades", PC.Calc.staffStats(c, { by: "grade", settings: demo.settings }).rows.map(x => x.key), ["수석", "선임", "고급"]);
  const rows = [{ name: "a", area: "PMO", grade: "선임" }, { name: "b", area: "개발", grade: "수석" }, { name: "c", area: "PMO", grade: "수석" }, { name: "d", area: "", grade: "수석" }];
  eq("orderRows: area then grade, blank area last", PC.Calc.orderRows("internal_labor", rows, ost).map(x => x.name), ["b", "c", "a", "d"]);
  eq("expense items in account order", c.summary.composition.find(g => g.key === "expense").items.map(x => x.name), ["출장비", "복리후생비", "외주 관리비"].filter(n => c.summary.composition.find(g => g.key === "expense").items.some(x => x.name === n)));

  // 분할계약 원가: 내부 · 외주 · 경비 합 = 회차 원가
  eq("phase cost = internal + external + expense", c.phases.map(ph => PC.pyRound(ph.internal_labor + ph.external_labor + ph.expense, 2) === ph.cost), c.phases.map(() => true));

  // 버전(기준선): 남긴 자료와 지금을 비교, 불러오기는 ID · 버전 목록을 지킴
  const vp = M.demoProject(M.defaultSettings());
  vp.versions = [{ no: 1, name: "초기 계획", data: M.snapshot(vp) }];
  eq("snapshot leaves out the version list", vp.versions[0].data.versions, undefined);
  vp.external_labor = []; vp.contract_revenue = 250000000;
  const cmpr = PC.Calc.compareVersions(PC.Calc.calculateProjectCost(vp.versions[0].data, vp.settings), PC.Calc.calculateProjectCost(vp, vp.settings));
  const byKey = k => cmpr.summary.find(x => x.key === k);
  eq("compare: revenue +10,000,000, external labor removed, staff row shows -4 MM", [byKey("revenue").diff, byKey("external_labor").b, cmpr.staff.find(x => x.kind === "외주").d_mm], [10000000, 0, -4]);
  eq("next version name follows the list", M.nextVersionName(vp), "실계약");
  vp.project_id = "KEEP";
  const rv = M.restoreVersion(vp, vp.versions[0], M.defaultSettings());
  eq("restore brings back the data, keeps ID and versions", [rv.external_labor.length, rv.contract_revenue, rv.project_id, rv.versions.length], [1, 240000000, "KEEP", 1]);
  console.log(`\nUnit: ${pass} / ${pass + fail} passed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
