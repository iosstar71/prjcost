// =============================================================================
// 엑셀 파일의 시트 구성 — 원본 v4 와 같은 시트 이름 · 열 (원본 프로그램 · 사람이 Excel 로 열어 봐도 같게)
//  프로젝트: 01_Project_Info … 07_Estimate, 99_Audit_Log, State_JSON(숨김 — 다시 불러올 때 쓰는 원본 자료)
//  기준정보: Overhead_Config · Internal_Grades · External_Grades · External_Rates · Expense_Accounts, Master_JSON(숨김)
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const X = PC.Xlsx || require("./xlsx.js");
  const { num, nowText } = PC;
  const AUDIT_KEEP = 300;   // 저장 이력은 최근 300줄만 (파일이 끝없이 커지지 않게)

  const head = cells => ({ cells, style: "head" });
  const widths = (n, w = 16, first) => Array.from({ length: n }, (_, i) => (i === 0 && first ? first : w));

  function projectSheets(project, calc, auditRows) {
    const s = calc.summary, b = s.breakdown, months = calc.months || [], st = project.settings || {};
    const info = [
      ["프로젝트 ID", project.project_id], ["프로젝트명", project.project_name], ["고객사", project.client_name], ["프로젝트 PM", project.pm_name],
      ["영업대표", project.sales_rep], ["시작일/월", project.start_date || project.start_ym], ["종료일/월", project.end_date || project.end_ym],
      ["프로젝트 설명", { v: project.description || "", s: "wrap" }],
      ["계약 매출(원)", s.revenue], ["총 원가(원)", s.total_cost], ["예상 이익(원)", s.profit], ["예상 이익률(%)", { v: s.margin_rate, s: "dec" }],
      ["목표 허들(%)", { v: s.hurdle_rate, s: "dec" }], ["허들 대비(%p)", { v: s.hurdle_gap, s: s.hurdle_met ? "ok" : "warn" }],
      ["내부 인건비(원)", b.internal_labor], ["외주 인건비(원)", b.external_labor], ["일반 경비(원)", b.general_expense],
      ["내부 인비례경비(원)", b.internal_overhead], ["외주 인비례경비(원)", b.external_overhead], ["기타 비용(원)", b.custom_cost]
    ].map(([k, v]) => ({ cells: [{ v: k, s: "label" }, v === undefined || v === null ? "" : v] }));

    const settingsRows = [
      ["인비례설정", "internal_overhead_ratio", "비율", num(st.internal_overhead_ratio)],
      ["인비례설정", "external_overhead_ratio", "비율", num(st.external_overhead_ratio)],
      ["승진설정", "default_promotion_month", "기본승진월", String(st.default_promotion_month || "03")],
      ...(st.internal_grades || []).map(g => ["내부등급", g.grade || "", "기준단가", num(g.rate)]),
      ...(st.external_grades || []).map(g => ["외주등급", String(g), "등급목록", ""]),
      ...(st.vendor_rates || []).map(r => ["외주업체단가", r.vendor || "", r.grade || "", num(r.rate)]),
      ...(st.expense_accounts || []).map(a => ["경비계정", a.account || "", a.category || "general", a.desc || ""])
    ].map(r => ({ cells: r.map(v => (typeof v === "number" && v < 1 && v > 0 ? { v, s: "dec" } : v)) }));

    const d = calc.details;
    const internal = d.internal_labor.map(m => {
      const p0 = (m.promotions || [])[0] || {};
      return { cells: [m.name || "", m.role || "", m.grade || "", num(m.base_rate), m.overhead_eligible ? "Y" : "N", p0.effective_ym || "", p0.rate === undefined || p0.rate === "" ? "" : num(p0.rate), { v: m.total_mm, s: "dec" },
        ...months.map(mo => ({ v: m.calc[mo].mm, s: "dec" })), m.total_cost] };
    });
    const external = d.external_labor.map(e => ({ cells: [e.name || "", e.role || "", e.vendor || "", e.grade || "", e.default_rate, e.applied_rate, { v: e.total_mm, s: "dec" },
      ...months.map(mo => ({ v: e.calc[mo].mm, s: "dec" })), e.total_cost] }));
    const expIn = e => (e.dist_mode === "fill_all" ? num(e.unit_input) : num(e.total_input));
    const ratioText = r => `자동(${(num(r) * 100).toFixed(1)}%)`;
    const expenses = [
      ...d.general_expenses.map(e => ["일반경비", e.name || "", e.dist_mode || "", expIn(e), ...months.map(m => e.calc_monthly[m]), e.total_cost]),
      ["내부인비례(자동)", "내부 인비례성 경비", ratioText(st.internal_overhead_ratio), "-", ...months.map(m => d.internal_overhead_monthly[m]), b.internal_overhead],
      ["외주인비례(자동)", "외주 인비례성 경비", ratioText(st.external_overhead_ratio), "-", ...months.map(m => d.external_overhead_monthly[m]), b.external_overhead],
      ...d.custom_costs.map(e => ["기타비용", e.name || "", e.dist_mode || "", expIn(e), ...months.map(m => e.calc_monthly[m]), e.total_cost])
    ].map(cells => ({ cells }));
    const phases = calc.phases.map(p => ({ cells: [p.phase_no, p.title, p.start_ym, p.end_ym, p.revenue, p.cost, { v: p.margin_rate, s: "dec" }, p.cum_revenue, p.cum_cost, p.rem_revenue, p.rem_cost] }));
    const est = calc.estimate;
    const estimate = [
      { cells: [{ v: "견적번호", s: "label" }, est.estimate_no, { v: "견적일자", s: "label" }, est.estimate_date, { v: "매출적용여부", s: "label" }, est.applied_to_revenue ? "Y" : "N"] },
      head(["항목명", "규격/등급", "수량", "단가", "금액", "비고"]),
      ...est.items.map(it => ({ cells: [it.item_name || "", it.spec || "", { v: it.qty, s: "dec" }, it.unit_price, it.amount, it.memo || ""] })),
      { cells: ["공급가액 합계", "", "", "", est.supply_total, ""], style: "total" },
      { cells: ["부가세(10%)", "", "", "", est.vat, ""], style: "total" },
      { cells: ["견적 총액(VAT포함)", "", "", "", est.grand_total, ""], style: "total" }
    ];
    const audit = [head(["일시", "작업자", "프로젝트ID", "계약매출", "총원가", "이익률(%)"]),
      ...(auditRows || []).slice(-(AUDIT_KEEP - 1)).map(r => ({ cells: r })),
      { cells: [nowText(), "local_user", project.project_id || "", s.revenue, s.total_cost, { v: s.margin_rate, s: "dec" }] }];

    const mHead = months.map(m => m);
    return [
      { name: "01_Project_Info", cols: [22, 48], rows: [head(["항목", "내용"]), ...info] },
      { name: "02_Project_Settings", cols: widths(4, 22), rows: [head(["구분", "항목1(업체/등급/키)", "항목2(등급/유형)", "값(단가/비율/설명)"]), ...settingsRows] },
      { name: "03_Internal_Labor", cols: widths(9 + months.length), rows: [head(["성명", "역할", "등급", "기준단가", "인비례포함", "승진월", "승진단가", "총MM", ...mHead, "합계금액"]), ...internal] },
      { name: "04_External_Labor", cols: widths(8 + months.length), rows: [head(["성명/구분", "역할", "외주업체", "등급", "기준단가", "적용단가", "총MM", ...mHead, "합계금액"]), ...external] },
      { name: "05_Expenses", cols: widths(5 + months.length), rows: [head(["구분", "경비계정/항목", "분배방식", "입력금액", ...mHead, "합계금액"]), ...expenses] },
      { name: "06_Phases", cols: widths(11), rows: [head(["회차", "구분", "시작월", "종료월", "회차매출", "회차원가", "회차이익률(%)", "누적매출", "누적원가", "잔여매출", "잔여원가"]), ...phases] },
      { name: "07_Estimate", cols: widths(6, 18), freeze: false, rows: estimate },
      { name: "99_Audit_Log", cols: [20, 12, 16, 16, 16, 12], rows: audit },
      { name: "State_JSON", hidden: true, rows: X.jsonRows(project) }
    ];
  }

  // 저장한 파일의 이전 이력 (머리 줄 제외)
  const auditRowsOf = wb => { const sh = wb && wb.sheet("99_Audit_Log"); return sh ? sh.rows.slice(1).filter(r => r && r.length && r[0] !== "") : []; };

  function masterSheets(settings) {
    const st = settings || {};
    return [
      { name: "Overhead_Config", cols: [28, 14, 24], rows: [head(["config_key", "value", "description"]),
        { cells: ["internal_overhead_ratio", { v: num(st.internal_overhead_ratio), s: "dec" }, "내부 인비례 경비율"] },
        { cells: ["external_overhead_ratio", { v: num(st.external_overhead_ratio), s: "dec" }, "외주 인비례 경비율"] },
        { cells: ["default_promotion_month", String(st.default_promotion_month || "03"), "내부 기본 승진월(MM)"] }] },
      { name: "Internal_Grades", cols: [16, 16], rows: [head(["grade", "monthly_rate"]), ...(st.internal_grades || []).map(g => ({ cells: [g.grade || "", num(g.rate)] }))] },
      { name: "External_Grades", cols: [16], rows: [head(["grade"]), ...(st.external_grades || []).map(g => ({ cells: [String(g)] }))] },
      { name: "External_Rates", cols: [16, 12, 16], rows: [head(["vendor", "grade", "monthly_rate"]), ...(st.vendor_rates || []).map(r => ({ cells: [r.vendor || "", r.grade || "", num(r.rate)] }))] },
      { name: "Expense_Accounts", cols: [18, 12, 32], rows: [head(["account", "category", "desc"]), ...(st.expense_accounts || []).map(a => ({ cells: [a.account || "", a.category || "general", a.desc || ""] }))] },
      { name: "Master_JSON", hidden: true, rows: X.jsonRows(st) }
    ];
  }

  // 기준정보 읽기: Master_JSON 이 있으면 그것, 없으면 예전 시트(Overhead_Config · External_Rates)
  function settingsFrom(wb, defaults) {
    const json = X.readJson(wb.sheet("Master_JSON"));
    if (json) return Object.assign({}, defaults, json);
    const st = Object.assign({}, defaults);
    const cfg = wb.sheet("Overhead_Config");
    if (cfg) cfg.rows.slice(1).forEach(r => {
      const k = String(r[0] || "").trim();
      if (k === "internal_overhead_ratio" && r[1] !== "") st.internal_overhead_ratio = num(r[1]);
      else if (k === "external_overhead_ratio" && r[1] !== "") st.external_overhead_ratio = num(r[1]);
      else if (k === "default_promotion_month" && r[1] !== "") st.default_promotion_month = String(r[1]).padStart(2, "0");
    });
    const vr = wb.sheet("External_Rates");
    if (vr) {
      const rows = vr.rows.slice(1).filter(r => r[0] && r[1] && r[2] !== "" && r[2] !== undefined).map(r => ({ vendor: String(r[0]).trim(), grade: String(r[1]).trim(), rate: num(r[2]) }));
      if (rows.length) st.vendor_rates = rows;
    }
    return st;
  }

  PC.Workbook = { projectSheets, masterSheets, settingsFrom, auditRowsOf };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Workbook;
})(typeof window !== "undefined" ? window : globalThis);
