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
  const catName = c => ((PC.Model.CATEGORIES.find(x => x.v === c) || {}).t || c || "일반경비");
  // 글이 들어갈 줄 높이: 칸 너비(글자 수) 기준으로 줄 수를 어림 (한글은 한 글자에 두 칸)
  function textHeight(text, width) {
    const lines = String(text || "").split(/\r?\n/).reduce((n, line) => n + Math.max(1, Math.ceil([...line].reduce((w, ch) => w + (ch.charCodeAt(0) > 255 ? 2 : 1), 0) / Math.max(10, width - 2))), 0);
    return Math.max(20, Math.min(400, lines * 16 + 6));
  }
  const widths = (n, w = 16, first) => Array.from({ length: n }, (_, i) => (i === 0 && first ? first : w));

  function projectSheets(project, calc, auditRows) {
    const s = calc.summary, b = s.breakdown, months = calc.months || [], st = project.settings || {};
    const info = [
      ["프로젝트 ID", project.project_id], ["프로젝트명", project.project_name], ["고객사", project.client_name], ["프로젝트 PM", project.pm_name],
      ["영업대표", project.sales_rep], ["시작일/월", project.start_date || project.start_ym], ["종료일/월", project.end_date || project.end_ym],
      ["프로젝트 설명", { v: project.description || "", s: "wrap" }, textHeight(project.description, 46)],
      ["계약 매출(원)", s.revenue], ["총 원가(원)", s.total_cost], ["예상 이익(원)", s.profit], ["예상 이익률(%)", { v: s.margin_rate, s: "dec" }],
      ["목표 허들(%)", { v: s.hurdle_rate, s: "dec" }], ["허들 대비(%p)", { v: s.hurdle_gap, s: s.hurdle_met ? "ok" : "warn" }],
      ["내부 인건비(원)", b.internal_labor], ["외주 인건비(원)", b.external_labor], ["일반 경비(원)", b.general_expense],
      ["내부 인비례경비(원)", b.internal_overhead], ["외주 인비례경비(원)", b.external_overhead], ["기타 비용(원)", b.custom_cost]
    ].map(([k, v, ht]) => ({ cells: [{ v: k, s: "label" }, v === undefined || v === null ? "" : v], ht }));

    const settingsRows = [
      ["연봉인상", "raise_month", "연봉인상월", String(st.raise_month || "")],
      ...(st.internal_grades || []).map(g => ["내부등급", g.grade || "", "기준단가", num(g.rate)]),
      ...(st.internal_grades || []).filter(g => g.raise_rate !== undefined && g.raise_rate !== "").map(g => ["내부등급인상", g.grade || "", "인상단가", num(g.raise_rate)]),
      ...(st.external_grades || []).map(g => ["외주등급", String(g), "등급목록", ""]),
      ...(st.vendor_rates || []).map(r => ["외주업체단가", r.vendor || "", r.grade || "", num(r.rate)]),
      ...(st.vendor_rates || []).filter(r => r.raise_rate !== undefined && r.raise_rate !== "").map(r => ["외주업체인상", `${r.vendor || ""} ${r.grade || ""}`, r.raise_ym || "", num(r.raise_rate)]),
      ...(st.areas || []).map(a => ["영역", String(a), "", ""]),
      ...(st.systems || []).map(a => ["대표시스템", String(a), "", ""]),
      ...(st.proposal_roles || []).flatMap(r => (st.proposal_levels || []).map(l => ["제안단가", r.role || "", l, num((r.rates || {})[l])])),
      ...(st.approvers || []).map((a, i) => ["결재자", ["기안(PM)", "검토", "사업팀장", "영업팀장", "본부장"][i] || `결재${i + 1}`, "", String(a || "")]),
      ...(st.departments || []).map(a => ["부서", String(a), "부서매출", num((project.dept_revenue || {})[a])]),
      ...(project.phases || []).flatMap((ph, i) => (st.departments || []).filter(a => num((ph.dept_revenue || {})[a])).map(a => ["회차부서매출", `${i + 1}차 ${ph.title || ""}`, String(a), num(ph.dept_revenue[a])])),
      ...(st.expense_accounts || []).map(a => ["경비계정", a.account || "", catName(a.category), PC.Model.isMmCategory(a.category) ? num(a.unit_amount) : a.desc || ""])
    ].map(cells => ({ cells }));

    const d = calc.details;
    const internal = d.internal_labor.map(m => {
      return { cells: [m.name || "", m.role || "", m.area || "", m.system || "", m.dept || "", m.prop_role || "", m.prop_level || "", m.grade || "", num(m.base_rate), m.overhead_eligible ? "Y" : "N", { v: m.total_mm, s: "dec" },
        ...months.map(mo => ({ v: m.calc[mo].mm, s: "dec" })), m.total_cost] };
    });
    const external = d.external_labor.map(e => ({ cells: [e.name || "", e.role || "", e.area || "", e.system || "", e.dept || "", e.prop_role || "", e.prop_level || "", e.vendor || "", e.grade || "", e.default_rate, e.applied_rate, { v: e.total_mm, s: "dec" },
      ...months.map(mo => ({ v: e.calc[mo].mm, s: "dec" })), e.total_cost] }));
    const expIn = e => (e.dist_mode === "fill_all" ? num(e.unit_input) : num(e.total_input));
    const expenses = [
      ...d.general_expenses.map(e => ["일반경비", e.name || "", e.dist_mode || "", expIn(e), ...months.map(m => e.calc_monthly[m]), e.total_cost, e.dept || ""]),
      ...(d.mm_overheads || []).map(o => [o.category === "mm_internal" ? "인비례(내부)" : "인비례(외주)", o.account, "MM×기준금액", o.unit_amount, ...months.map(m => o.calc_monthly[m]), o.total_cost]),
      ...d.custom_costs.map(e => ["기타경비", e.name || "", e.dist_mode || "", expIn(e), ...months.map(m => e.calc_monthly[m]), e.total_cost, e.dept || ""])
    ].map(cells => ({ cells }));
    const phases = calc.phases.map(p => ({ cells: [p.phase_no, p.title, p.start_ym, p.end_ym, p.revenue, p.cost, { v: p.margin_rate, s: "dec" }, p.cum_revenue, p.cum_cost, p.rem_revenue, p.rem_cost, p.internal_labor, p.external_labor, p.expense] }));
    const est = calc.estimate;
    const estimate = [
      { cells: [{ v: "견적번호", s: "label" }, est.estimate_no, { v: "견적일자", s: "label" }, est.estimate_date, { v: "매출적용여부", s: "label" }, est.applied_to_revenue ? "Y" : "N"] },
      head(["항목명", "규격/등급", "수량(MM)", "단가", "금액", "비고"]),
      ...est.labor.map(l => ({ cells: [`인건비 · ${l.role || "?"}`, l.level || "", { v: l.mm, s: "dec" }, l.applied_rate, l.amount, l.memo || ""] })),
      ...est.items.map(it => ({ cells: [it.item_name || "", it.spec || "", { v: it.qty, s: "dec" }, it.unit_price, it.amount, it.memo || ""] })),
      { cells: ["공급가액 합계", "", "", "", est.supply_total, ""], style: "total" },
      { cells: ["부가세(10%)", "", "", "", est.vat, ""], style: "total" },
      { cells: ["견적 총액(VAT포함)", "", "", "", est.grand_total, ""], style: "total" }
    ];
    const audit = [head(["일시", "작업자", "프로젝트ID", "계약매출", "총원가", "이익률(%)"]),
      ...(auditRows || []).slice(-(AUDIT_KEEP - 1)).map(r => ({ cells: r })),
      { cells: [nowText(), "local_user", project.project_id || "", s.revenue, s.total_cost, { v: s.margin_rate, s: "dec" }] }];

    // 버전(기준선) 요약 — 자료 전체는 State_JSON 의 versions 에
    const versions = [head(["번호", "이름", "저장일시", "메모", "계약매출", "총원가", "이익률(%)"]),
      ...(project.versions || []).map(v => { const vs = PC.Calc.calculateProjectCost(v.data || {}, st).summary; return { cells: [v.no, v.name || "", v.saved_at || "", { v: v.memo || "", s: "wrap" }, vs.revenue, vs.total_cost, { v: vs.margin_rate, s: "dec" }] }; })];
    const mHead = months.map(m => m);
    return [
      { name: "01_Project_Info", cols: [22, 48], rows: [head(["항목", "내용"]), ...info] },
      { name: "02_Project_Settings", cols: widths(4, 22), rows: [head(["구분", "항목1(업체/등급/키)", "항목2(등급/유형)", "값(단가/비율/설명)"]), ...settingsRows] },
      { name: "03_Internal_Labor", cols: widths(12 + months.length), rows: [head(["성명", "역할", "영역", "대표시스템", "부서", "제안직무", "제안등급", "직급", "기준단가", "인비례포함", "총MM", ...mHead, "합계금액"]), ...internal] },
      { name: "04_External_Labor", cols: widths(13 + months.length), rows: [head(["성명/구분", "역할", "영역", "대표시스템", "부서", "제안직무", "제안등급", "외주업체", "등급", "기준단가", "적용단가", "총MM", ...mHead, "합계금액"]), ...external] },
      { name: "05_Expenses", cols: widths(6 + months.length), rows: [head(["구분", "경비계정/항목", "분배방식", "입력금액", ...mHead, "합계금액", "부서(비우면 공통)"]), ...expenses] },
      { name: "06_Phases", cols: widths(14), rows: [head(["회차", "구분", "시작월", "종료월", "회차매출", "회차원가", "회차이익률(%)", "누적매출", "누적원가", "잔여매출", "잔여원가", "내부인건비", "외주인건비", "경비"]), ...phases] },
      { name: "07_Estimate", cols: widths(6, 18), freeze: false, rows: estimate },
      { name: "08_Versions", cols: [8, 16, 20, 40, 16, 16, 12], freeze: false, rows: versions },
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
        { cells: ["raise_month", String(st.raise_month || ""), "연봉 인상월(MM) — 등급의 인상단가가 이 달부터"] }] },
      { name: "Internal_Grades", cols: [16, 16, 16], rows: [head(["grade", "monthly_rate", "raise_rate"]), ...(st.internal_grades || []).map(g => ({ cells: [g.grade || "", num(g.rate), g.raise_rate === undefined || g.raise_rate === "" ? "" : num(g.raise_rate)] }))] },
      { name: "External_Grades", cols: [16], rows: [head(["grade"]), ...(st.external_grades || []).map(g => ({ cells: [String(g)] }))] },
      { name: "External_Rates", cols: [16, 12, 16, 12, 16], rows: [head(["vendor", "grade", "monthly_rate", "raise_ym", "raise_rate"]), ...(st.vendor_rates || []).map(r => ({ cells: [r.vendor || "", r.grade || "", num(r.rate), r.raise_ym || "", r.raise_rate === undefined || r.raise_rate === "" ? "" : num(r.raise_rate)] }))] },
      { name: "Proposal_Rates", cols: [16, 10, 16], rows: [head(["role", "level", "monthly_rate"]), ...(st.proposal_roles || []).flatMap(r => (st.proposal_levels || []).map(l => ({ cells: [r.role || "", l, num((r.rates || {})[l])] })))] },
      { name: "Areas", cols: [20], rows: [head(["area"]), ...(st.areas || []).map(a => ({ cells: [String(a)] }))] },
      { name: "Expense_Accounts", cols: [18, 14, 16, 32], rows: [head(["account", "category", "unit_amount(원/MM)", "desc"]), ...(st.expense_accounts || []).map(a => ({ cells: [a.account || "", a.category || "general", PC.Model.isMmCategory(a.category) ? num(a.unit_amount) : "", a.desc || ""] }))] },
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
