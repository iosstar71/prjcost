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
require("../js/formula.js");
require("../js/template.js");

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
  const demo = M.sampleProject(M.defaultSettings());
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
  const vp = M.sampleProject(M.defaultSettings());
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

  // 부서별 손익: 인력 · 인비례는 사람의 부서, 경비는 항목의 부서(공통은 기준대로), 매출은 입력 — 합계 = 통합
  const dp = M.sampleProject(M.defaultSettings());
  eq("one department or none = no department screens", [M.multiDept(dp), M.multiDept(Object.assign({}, dp, { settings: Object.assign({}, dp.settings, { departments: ["SI"] }) }))], [false, false]);
  dp.settings.departments = ["SI", "AI"];
  dp.internal_labor[0].dept = "SI"; dp.internal_labor[1].dept = "AI"; dp.external_labor[0].dept = "SI";
  dp.custom_costs[0].dept = "AI";
  dp.dept_revenue = { SI: 140000000, AI: 100000000 };
  const dc = PC.Calc.calculateProjectCost(dp, dp.settings), ds = PC.Calc.deptSummary(dp, dc);
  const bk = dc.summary.breakdown;
  eq("dept totals = project totals (cost, revenue, overhead)", [ds.total.cost, ds.total.revenue, ds.total.overhead, ds.revenue_gap], [dc.summary.total_cost, 240000000, bk.internal_overhead + bk.external_overhead, 0]);
  eq("dept: external staff counted in their department", ds.rows.map(g => [g.dept, g.external_labor > 0]), [["SI", true], ["AI", false]]);
  eq("dept: common expense split by direct cost (sums to the common total)", [PC.pyRound(ds.rows[0].common + ds.rows[1].common, 2), ds.common, ds.rows[0].common > ds.rows[1].common], [bk.general_expense, bk.general_expense, true]);
  dp.settings.dept_common_basis = "equal";
  const de = PC.Calc.deptSummary(dp, dc);
  eq("dept: common split equally", de.rows[0].common, de.rows[1].common);
  dp.internal_labor[1].dept = "";
  eq("dept: unassigned staff shown as (미지정)", PC.Calc.deptSummary(dp, PC.Calc.calculateProjectCost(dp, dp.settings)).rows.map(g => g.dept), ["SI", "AI", "(미지정)"]);

  // 분할계약 회차 × 부서: 기간은 회차 그대로, 원가 합 = 회차 원가, 매출은 회차 × 부서 입력, 별도 원가는 부서 비율대로
  dp.internal_labor[1].dept = "AI";
  dp.settings.dept_common_basis = "cost";
  dp.phases[0].dept_revenue = { SI: 70000000, AI: 40000000 };
  const pc = PC.Calc.calculateProjectCost(dp, dp.settings), pds = PC.Calc.phaseDeptSummary(dp, pc);
  eq("phase x dept: each phase's dept cost sums to the phase cost", pds.map(r => r.total.cost), pc.phases.map(ph => ph.cost));
  eq("phase x dept: revenue entered per phase (gap vs phase revenue)", [pds[0].total.revenue, pds[0].revenue_gap, pds[1].revenue_gap], [110000000, 0, 130000000]);
  eq("phase x dept: sum over phases = project dept cost (per dept)", ["SI", "AI"].map(d => PC.pyRound(pds.reduce((a, r) => a + r.rows.find(g => g.dept === d).cost, 0), 2)), ["SI", "AI"].map(d => PC.Calc.deptSummary(dp, pc).rows.find(g => g.dept === d).cost));
  dp.phases[1].override_cost = 50000000;
  const po = PC.Calc.phaseDeptSummary(dp, PC.Calc.calculateProjectCost(dp, dp.settings))[1];
  eq("phase x dept: override cost scaled to the dept ratio", [po.overridden, po.total.cost], [true, 50000000]);

  // 보고서 양식: 자리표시 칸에만 값 (숫자는 숫자로 → 양식 서식), 섞인 글은 글자로, 모르는 이름은 알림
  const T = PC.Template, tp = M.sampleProject(M.defaultSettings());
  tp.project_type = "SI-AD"; tp.settings.approvers = ["김철수", "", "", "", "정본부"]; tp.approvers = ["옛이름"]; tp.settings.departments = ["SI사업부", "AI사업부"];
  tp.custom_costs.push({ name: "AI예비비", dist_mode: "equal_total", total_input: 3000000, unit_input: 0, monthly: {}, memo: "" });
  tp.versions = [{ no: 1, name: "수주협 3차", saved_at: "2026-01-02 10:00:00", data: Object.assign(M.snapshot(tp), { contract_revenue: 200000000 }) },
    { no: 2, name: "수주협 4차", saved_at: "2026-02-02 10:00:00", data: Object.assign(M.snapshot(tp), { contract_revenue: 230000000 }) }];
  const tc = PC.Calc.calculateProjectCost(tp, tp.settings), tv = T.values(tp, tc, { lastSaved: new Date(2026, 10, 25) });
  eq("template values: cost split (AI reserve apart), margin, hurdle, gap, latest version", [tv.노무비, tv.외주비, tv.경비, tv.AI예비비, tv.총원가, PC.pyRound(tv.수주마진율, 4), tv.허들, PC.pyRound(tv.Gap, 4), tv.최근버전이름, tv.최근버전매출, tv.시작일, tv.종료일, tv.귀속부서],
    [79800000, 31600000, 13600000, 3000000, 128000000, 0.4667, 0.18, 0.2867, "수주협 4차", 230000000, "2026.01.01", "2026.06.30", "SI사업부"]);
  const tf = fs.readFileSync(path.join(__dirname, "..", "data", "templates", "Offline사업검토서_단가고정.xlsx"));
  const filled = await T.fill(tf, tv);
  const fwb = await X.read(Buffer.from(await filled.blob.arrayBuffer())), sh = fwb.sheet("사업개요").rows;
  const at = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); const col = m[1].charCodeAt(0) - 65; return (sh[Number(m[2]) - 1] || [])[col]; };
  eq("template fill: numbers stay numbers, mixed text replaced, approvers, empty value cleared", [at("D8"), at("I8"), at("D9"), at("F4"), at("G4"), at("D13"), PC.pyRound(at("F13"), 4), at("E38"), String(at("D7")).split("\n")[0], filled.unknown.length],
    [240000000, "SI-AD", "2026.01.01 ~ 2026.06.30 (6개월)", "김철수", "", 79800000, 0.6234, 230000000, "1. 과제명 : 차세대 원가·공정 시스템 구축", 0]);
  eq("template fill: empty values reported", filled.empty.includes("결재2"), true);
  // 원가상세: 사람마다 한 줄(반복), OT비용은 프로젝트 전체 한 줄, 소계 · 인원수 · 근무지 · 내부 기술등급
  const dsh = fwb.sheet("원가상세").rows, cellAt = (sh, r, c) => (sh[r - 1] || [])[c - 1];
  eq("detail sheet: one row per person, project OT row, subtotals with head count, external row shifted", [cellAt(dsh, 5, 8), cellAt(dsh, 6, 8), cellAt(dsh, 5, 7), cellAt(dsh, 5, 10), cellAt(dsh, 7, 5), cellAt(dsh, 8, 8), cellAt(dsh, 8, 13), cellAt(dsh, 11, 8), cellAt(dsh, 11, 11), cellAt(dsh, 13, 8)],
    ["김철수", "이민지", "본사", "특급", "OT비용", 2, 10.5, "외주개발1", "개발", 3]);
  const t5 = fwb.sheet("인력투입계획 점검(5Tier)").rows;
  eq("5Tier sheet: 자사 ○ for internal rows, BP ○ for external, totals by month", [cellAt(t5, 5, 23), cellAt(t5, 6, 23), cellAt(t5, 7, 24), cellAt(t5, 7, 23), cellAt(t5, 8, 7), cellAt(t5, 8, 9)], ["○", "○", "○", "", 14.5, 3]);
  const odd = await X.read(Buffer.from(await (await T.fill(X.build([{ name: "S", rows: [["{{없는값}}", "합계 {{매출액}}원"]] }]), tv)).blob.arrayBuffer()));
  eq("template fill: our own xlsx (inline strings) and unknown keys", odd.sheet("S").rows[0], ["", "합계 240,000,000원"]);
  eq("output name: 양식_ID_이름_최종작성일", T.outName("Offline사업검토서.xlsx", { project_id: "PRJ-1", project_name: "A/B 구축" }, new Date(2026, 10, 25)), "Offline사업검토서_PRJ-1_A_B 구축_20261125.xlsx");
  const mv = M.normalize({ approvers: ["홍", "", "", "", ""], settings: M.defaultSettings() }, M.defaultSettings());
  eq("approvers live in settings; old per-project names move there", [mv.settings.approvers[0], mv.approvers], ["홍", undefined]);
  // 보고서: 4개 시트(부서 2개 이상), 제목 · 회색 이름 칸 · 원가 구성 총 원가 = 계산 결과
  require("../js/report.js");
  const rs = PC.Report.sheets(dp, PC.Calc.calculateProjectCost(dp, dp.settings));
  const flat = rs[0].rows.flatMap(r => r.cells.map(x => (x ? x.v : "")));
  eq("report: sheets and the total cost shown on the summary", [rs.map(x => x.name).join(), flat.includes("총 원가"), flat.includes(PC.Calc.calculateProjectCost(dp, dp.settings).summary.total_cost)], ["01_통합보고,02_항목별상세,03_투입현황,04_부서별손익", true, true]);

  // 제안 직무 · 등급 → 견적 인건비, 마스터 (전사 기준정보) 반영 · 비교
  const ep = M.sampleProject(M.defaultSettings()), ec = PC.Calc.calculateProjectCost(ep, ep.settings).estimate;
  eq("estimate: labor lines = MM x proposal rate (role x level), plus other items", [ec.labor.map(l => l.amount), ec.labor_total, ec.items_total, ec.supply_total], [[78000000, 60000000, 64000000], 202000000, 30000000, 232000000]);
  ep.estimate.labor[0].unit_price = 12000000; ep.estimate.labor.push({ role: "", level: "Lvl1", mm: 1 });
  const ec2 = PC.Calc.calculateProjectCost(ep, ep.settings).estimate;
  eq("estimate: typed unit price wins; a line without role is flagged", [ec2.labor[0].amount, ec2.labor[3].missing], [72000000, true]);
  const fp = PC.Calc.estimateFromPlan(PC.Calc.calculateProjectCost(ep, ep.settings));
  eq("estimate from plan: MM grouped by proposal role · level", fp.labor.map(l => [l.role, l.level, l.mm]), [["프로젝트관리", "Lvl1", 6], ["컨설팅", "Lvl2", 4.5], ["개발", "Lvl3", 4]]);
  eq("master leaves out project-only settings (vendor rates, systems)", Object.keys(M.masterOf(ep.settings)).filter(k => ["vendor_rates", "systems", "dept_common_basis"].includes(k)), []);
  const mst = M.masterOf(M.defaultSettings());
  eq("master has no amounts (grade rates, proposal rates, overhead base amounts, raise month)", [mst.internal_grades[0], mst.proposal_roles[0], "unit_amount" in mst.expense_accounts[5], "raise_month" in mst], [{ grade: "수석" }, { role: "프로젝트관리" }, false, false]);
  mst.areas = ["컨설팅", "개발"]; mst.internal_grades.push({ grade: "전문위원" }); mst.expense_accounts[0].desc = "출장 (전사)";
  const prj = ep.settings; prj.internal_grades.push({ grade: "인턴", rate: 3000000 });
  const ch = M.masterChanges(prj, mst), merged = M.mergeMaster(prj, mst);
  eq("master changes listed (names added, account description)", [ch.includes("영역 추가: 컨설팅"), ch.includes("직급 추가: 전문위원"), ch.some(x => x.startsWith("경비 계정 출장비:")), ch.some(x => x.includes("8,500,000"))], [true, true, true, false]);
  eq("merge master: names added, project amounts and extras kept", [merged.internal_grades.find(g => g.grade === "수석").rate, merged.internal_grades.some(g => g.grade === "인턴"), merged.internal_grades.some(g => g.grade === "전문위원"), merged.areas.includes("컨설팅") && merged.areas.includes("PMO"), merged.vendor_rates.length, merged.proposal_roles[2].rates.Lvl3], [8500000, true, true, true, 6, 8000000]);
  eq("new project from a master starts with empty amounts and project-only lists", (() => { const np = M.newProject(mst); return [np.settings.internal_grades[0].rate, np.settings.vendor_rates.length, np.settings.systems.length]; })(), [undefined, 0, 0]);

  // OT비용(프로젝트 전체 · 내부 인건비) · 재료비(총 원가): 합계 · 원가 구성 · 부서 합계 = 통합
  const op = M.sampleProject(M.defaultSettings()), oc0 = PC.Calc.calculateProjectCost(op, op.settings);
  op.ot_cost = 6000000; op.materials = [{ category: "HW", content: "서버", amount: 12000000 }];
  const oc = PC.Calc.calculateProjectCost(op, op.settings), ob = oc.summary.breakdown;
  eq("OT is internal labor, materials add to total cost; month and composition follow", [ob.internal_labor - oc0.summary.breakdown.internal_labor, ob.ot_cost, ob.material, oc.summary.total_cost - oc0.summary.total_cost, oc.summary.composition.map(g => g.key).join(), oc.summary.composition[0].items.map(x => x.name).join()],
    [6000000, 6000000, 12000000, 18000000, "internal_labor,external_labor,material,expense,custom_cost", "OT비용"]);
  op.settings.departments = ["SI", "AI"]; op.internal_labor[0].dept = "SI"; op.internal_labor[1].dept = "AI"; op.external_labor[0].dept = "AI";
  const od = PC.Calc.deptSummary(op, PC.Calc.calculateProjectCost(op, op.settings));
  eq("departments: OT split by internal labor, materials as common — totals reconcile", [od.total.cost, od.total.internal_labor], [oc.summary.total_cost, ob.internal_labor]);

  // 견적: 합계 → 할인(%) → 절사(단위 아래 버림) → 견적 금액 (부가세 없음)
  const xp = M.sampleProject(M.defaultSettings());
  xp.estimate.discount_rate = 3.5; xp.estimate.round_unit = 100000;
  const xe = PC.Calc.calculateProjectCost(xp, xp.settings).estimate;
  eq("estimate discount and truncation (232,000,000 → −8,120,000 → 223,880,000 → 223,800,000)", [xe.subtotal, xe.discount, xe.truncation, xe.supply_total, "vat" in xe], [232000000, 8120000, 80000, 223800000, false]);
  xp.estimate.discount_rate = 0; xp.estimate.round_unit = 100; xp.estimate.items[0].unit_price = 30009999;
  eq("truncation by 100 won (…9,999 → …9,900)", PC.Calc.calculateProjectCost(xp, xp.settings).estimate.supply_total, 232009900);
  // 양식: 기간 밖의 달 열은 숨김 (샘플 6개월 → 5Tier 의 7~12월 열 N~S 숨김, H~M 보임)
  const t5x = new TextDecoder().decode((await X.unzip(await filled.blob.arrayBuffer()))[Object.keys(await X.unzip(await filled.blob.arrayBuffer())).find(n => /sheet3\.xml$/.test(n))]);
  const hidden = [...t5x.matchAll(/<col\b[^>]*hidden="1"[^>]*\/>/g)].map(m => Number(/min="(\d+)"/.exec(m[0])[1]));
  eq("template: month columns outside the period are hidden", [8, 13, 14, 19].map(c => hidden.includes(c)), [false, false, true, true]);
  // 수식: 계산기 (사업검토서 합계 수식 검증용) — Excel 과 같은 규칙
  const Fm = PC.Formula;
  const fv = Fm.evaluate({ S: { A1: { v: 1 }, A2: { v: 2 }, A3: { v: "AI예비비" }, B1: { f: "SUM(A1:A3)" }, B2: { f: 'COUNTIF(A1:A3,">1")' }, B3: { f: "IF(B1=0,0,A2/B1)" }, B4: { f: "'T 1'!A1+1" }, B5: { f: "VLOOKUP(1,A1:A2,1)" }, B6: { f: "1/0" }, B7: { f: "-2^2+50%" }, B8: { f: 'SUMIF(A3:A3,"AI예비비",A1:A1)+COUNTA(A1:A9)' }, B9: { f: "IFERROR(B6,7)&\"원\"" } }, "T 1": { A1: { f: "S!B1*2" } } });
  eq("formula: SUM · COUNTIF · IF · other sheet · -2^2 · % · SUMIF · COUNTA · IFERROR", [fv.values.S.B1, fv.values.S.B2, PC.pyRound(fv.values.S.B3, 4), fv.values.S.B4, fv.values.S.B7, fv.values.S.B8, fv.values.S.B9], [3, 1, 0.6667, 7, 4.5, 4, "7원"]);
  eq("formula: #DIV/0! is a value, unknown function reported", [Fm.isErr(fv.values.S.B6) && fv.values.S.B6.v, fv.problems.map(x => [x.ref, x.reason])], ["#DIV/0!", [["B5", "함수 VLOOKUP"]]]);
  eq("formula: move references only (text and strings kept), shared formula offset", [Fm.mapRefs('SUM(O5:O6)+원가상세!O7&"A1"', t => ({ a: Object.assign({}, t.a, { r: t.a.r + 2 }), b: t.b && Object.assign({}, t.b, { r: t.b.r + 2 }) })), Fm.offset("SUM(H5:S5)+$A$1+B$2", 2, 1)],
    ['SUM(O7:O8)+원가상세!O9&"A1"', "SUM(I7:T7)+$A$1+C$2"]);
  // 사업검토서: 여러 명이면 합계 범위가 늘고, 아래 칸 · 다른 시트 참조가 밀리고, 반복 줄 안의 수식은 줄마다 — 검증 시트로 프로그램 값과 비교
  const fbig = M.sampleProject(M.defaultSettings());
  fbig.internal_labor = Array.from({ length: 5 }, (_, i) => Object.assign(PC.clone(tp.internal_labor[i % 2]), { name: `내부${i + 1}` }));
  fbig.external_labor = Array.from({ length: 3 }, (_, i) => Object.assign(PC.clone(tp.external_labor[0]), { name: `외주${i + 1}` }));
  fbig.ot_cost = 3000000; fbig.materials = [{ category: "HW", content: "서버", amount: 5000000 }, { category: "SW", content: "DB", amount: 1000000 }];
  const bf = await T.fill(tf, T.values(fbig, PC.Calc.calculateProjectCost(fbig, fbig.settings), {}));
  const bfiles = await X.unzip(await bf.blob.arrayBuffer()), sx = n => new TextDecoder().decode(bfiles[Object.keys(bfiles).find(k => k.endsWith(`/${n}`))]);
  const fOf = (xml, ref) => X.xmlUnesc((new RegExp(`<c r="${ref}"[^>]*><f>([^<]*)</f>`).exec(xml) || [])[1] || "");
  const s1 = sx("sheet1.xml"), s2 = sx("sheet2.xml"), s3 = sx("sheet3.xml");
  eq("template formulas: ranges grow with rows, cells below and other-sheet refs move, row formulas per copy",
    [fOf(s2, "O11"), fOf(s2, "H11"), fOf(s2, "O17"), fOf(s2, "O18"), fOf(s1, "D13"), fOf(s1, "D14"), fOf(s1, "D17"), fOf(s3, "G7"), fOf(s3, "G12"), fOf(s3, "H13")],
    ["SUM(O5:O10)", 'COUNTIF(M5:M9,">0")', "SUM(O14:O16)", "O11+O17", "원가상세!O11", "원가상세!O17", 'SUMIF(원가상세!D28:D31,"AI예비비",원가상세!O28:O31)', "SUM(H7:S7)", "SUM(H12:S12)", "SUM(H5:H12)"]);
  eq("5Tier: project type merge runs to the last external row; external rows keep a bottom border", [/<mergeCell ref="B5:B12"\/>/.test(s3), ["C10", "C11", "C12", "D12", "F12", "G12"].map(c => new RegExp(`<c r="${c}" s="(\\d+)"`).exec(s3)[1]).every((v, i, l) => v === l[0] || i > 2)], [true, true]);
  eq("template check: every check row matches the program (cost, ratios, margin, gap, subtotals, 5Tier)", [bf.check.items.length, bf.check.failed.map(x => x.name), bf.check.problems.length], [40, [], 0]);
  eq("template: formula cells carry the computed value (for viewers that do not recalc)", /<c r="D19"[^>]*><f>[^<]*<\/f><v>([\d.]+)<\/v>/.exec(s1)[1], String(PC.Calc.calculateProjectCost(fbig, fbig.settings).summary.total_cost));
  // 양식 수식이 틀리면 검증에서 잡힘 (내부 소계가 OT 줄을 빠뜨림)
  const tfiles = await X.unzip(tf), tname = Object.keys(tfiles).find(k => k.endsWith("/sheet2.xml"));
  tfiles[tname] = new TextEncoder().encode(new TextDecoder().decode(tfiles[tname]).replace("<f>SUM(O5:O6)</f>", "<f>SUM(O5:O5)</f>"));
  const wr = await T.fill(await X.zip(Object.keys(tfiles).map(n => ({ name: n, data: tfiles[n] }))).arrayBuffer(), T.values(fbig, PC.Calc.calculateProjectCost(fbig, fbig.settings), {}));
  eq("template check: a subtotal that misses the OT row is reported", wr.check.failed.map(x => x.name).slice(0, 2), ["사업개요 노무비", "사업개요 원가 합계"]);
  // 사용자 매뉴얼: 머리줄 링크 · 그림 파일이 모두 있는지
  const man = fs.readFileSync(path.join(__dirname, "..", "manual", "index.html"), "utf8"), imgs = [...man.matchAll(/<img src="([^"]+)"/g)].map(m => m[1]);
  eq("manual: linked from the header, every picture exists", [/href="manual\/index.html"/.test(fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8")), imgs.length > 20, imgs.filter(f => !fs.existsSync(path.join(__dirname, "..", "manual", f)))], [true, true, []]);
  // 샘플 = 이 PC 의 마스터 목록 + 예시 금액 (숫자로 끝나는 등급은 같은 금액, 그 밖은 위가 높음), 결재자는 가짜
  const cm = M.masterOf(M.defaultSettings());
  Object.assign(cm, { internal_grades: [{ grade: "부장" }, { grade: "차장" }, { grade: "과장" }], external_grades: ["S", "A", "B"], proposal_levels: ["1등급", "2등급", "3등급"], proposal_roles: [{ role: "PMO" }, { role: "설계" }],
    areas: ["기획"], departments: ["DX본부", "SI본부"], approvers: ["실명1", "실명2", "", "", ""] });
  const sm = M.sampleProject(Object.assign(M.defaultSettings(), cm)), sc = PC.Calc.calculateProjectCost(sm, sm.settings), ss = sm.settings;
  eq("sample from the company master: names from the master lists, example rates (top = highest, numbered = same), fake approvers, departments split",
    [sm.internal_labor.map(r => r.grade).join(), ss.internal_grades.map(g => g.rate).join(), new Set(Object.values(ss.proposal_roles[0].rates)).size, sm.external_labor[0].grade,
      sm.estimate.labor.every(l => ss.proposal_roles.some(r => r.role === l.role && r.rates[l.level] > 0)), sc.estimate.labor_total > 0, sc.summary.breakdown.external_labor > 0,
      ss.approvers.includes("실명1"), sm.internal_labor.map(r => r.dept).join(), sm.dept_revenue["DX본부"], sm.internal_labor[0].area],
    ["부장,과장", "9000000,8000000,7000000", 1, "A", true, true, true, false, "DX본부,SI본부", 140000000, "기획"]);
  // 사업검토서 _단가변경: 최초 · 변경 단가와 MM, 통합 MM(수식), 근무지는 G, 재료비 · 경비 금액은 투입원가와 같은 S 열
  const cv = M.sampleProject(M.defaultSettings());
  cv.materials = [{ category: "HW", content: "서버", amount: 5000000 }];
  const cvf = await T.fill(fs.readFileSync(path.join(__dirname, "..", "data", "templates", "Offline사업검토서_단가변경.xlsx")), T.values(cv, PC.Calc.calculateProjectCost(cv, cv.settings), {}));
  const cvb = await X.read(Buffer.from(await cvf.blob.arrayBuffer())), cd = cvb.sheet("원가상세").rows;
  eq("단가변경 template: initial / changed rate and month, MM per rate, site in G, amounts in S, all checks pass",
    [cellAt(cd, 5, 8), cellAt(cd, 5, 13), cellAt(cd, 5, 14), cellAt(cd, 5, 15), cellAt(cd, 5, 16), cellAt(cd, 5, 17), cellAt(cd, 5, 18), cellAt(cd, 5, 7), cellAt(cd, 5, 19), cellAt(cd, 13, 19), cellAt(cd, 16, 19), cvf.check.failed.length, cvf.check.items.length],
    ["김철수", 2, 4, 6, 8500000, 8800000, "26.03", "본사", 52200000, 111400000, "금액", 0, 40]);
  eq("rate change: none when the rate stays (first rate, all MM at it)", (() => { const r = PC.Calc.rateChange({ calc: { a: { mm: 1, rate: 5 }, b: { mm: 2, rate: 5 } } }, ["a", "b"]); return [r.first, r.changed, r.mmFirst, r.mmChanged]; })(), [5, null, 3, 0]);
  console.log(`\nUnit: ${pass} / ${pass + fail} passed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
