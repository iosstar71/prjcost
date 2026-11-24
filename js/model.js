// =============================================================================
// 자료 모양 — 기본 기준정보 · 새 프로젝트 · 데모 · 예전 자료 채우기 · 업로드 양식 읽기
//  자료 정의는 docs/DATA_MODEL.md. 필드를 바꾸면 그 문서와 normalize() 를 함께 고친다
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const { clone, num, isBlank } = PC;

  // 경비 계정 구분: 일반경비 · 기타경비 (금액을 직접 넣음) / 인비례 (인력 MM × 기준금액으로 자동)
  const CATEGORIES = [
    { v: "general", t: "일반경비" }, { v: "custom", t: "기타경비" },
    { v: "mm_internal", t: "인비례 · 내부" }, { v: "mm_external", t: "인비례 · 외주" }
  ];
  const isMmCategory = c => c === "mm_internal" || c === "mm_external";

  // 기준정보가 하나도 없을 때(새 폴더) 쓰는 값. 인비례 기준금액(원/MM)은 회사 기준으로 넣어 쓰도록 0
  function defaultSettings() {
    return {
      default_promotion_month: "03",
      areas: ["PMO", "분석 · 설계", "개발", "테스트", "인프라"],   // 인력의 담당 영역(시스템) — 투입 현황을 영역별로 볼 때 씀
      internal_grades: [{ grade: "수석", rate: 8500000 }, { grade: "책임", rate: 7200000 }, { grade: "선임", rate: 6000000 }, { grade: "전임", rate: 4800000 }],
      external_grades: ["특급", "고급", "중급", "초급"],
      vendor_rates: [
        { vendor: "파트너A", grade: "특급", rate: 9500000 }, { vendor: "파트너A", grade: "고급", rate: 7800000 }, { vendor: "파트너A", grade: "중급", rate: 6200000 },
        { vendor: "파트너B", grade: "특급", rate: 9200000 }, { vendor: "파트너B", grade: "고급", rate: 7500000 }, { vendor: "파트너B", grade: "중급", rate: 6000000 }
      ],
      expense_accounts: [
        { account: "출장비", category: "general", desc: "국내/외 출장 경비" },
        { account: "회의·식대", category: "general", desc: "프로젝트 회의 및 업무추진비" },
        { account: "장비·SW임차", category: "general", desc: "개발장비 및 라이선스" },
        { account: "사전심의·인증비", category: "custom", desc: "외부 심의 및 보안인증 수수료" },
        { account: "예비비", category: "custom", desc: "리스크 대응 예비 비용" },
        { account: "복리후생비", category: "mm_internal", unit_amount: 0, desc: "인비례 대상 내부 인력 MM × 기준금액" },
        { account: "외주 관리비", category: "mm_external", unit_amount: 0, desc: "외주 인력 MM × 기준금액" }
      ]
    };
  }

  // 인비례 계정이 없는 기준정보(예전 v4 파일 등)에 기본 인비례 계정(기준금액 0)을 넣음 — 화면에서 바로 채울 수 있게
  function ensureMmAccounts(settings) {
    const accs = settings.expense_accounts = Array.isArray(settings.expense_accounts) ? settings.expense_accounts : [];
    defaultSettings().expense_accounts.filter(a => isMmCategory(a.category)).forEach(a => {
      if (!accs.some(x => x.category === a.category)) accs.push(Object.assign({}, a));
    });
    return settings;
  }

  const emptyEstimate = () => ({ estimate_no: "", estimate_date: "", note: "", applied_to_revenue: false, items: [] });

  function newProject(settings, id) {
    const y = new Date().getFullYear();
    return {
      project_id: id || `PRJ-${y}-NEW`, project_name: "신규 프로젝트", client_name: "", pm_name: "", sales_rep: "", description: "",
      date_mode: "month", start_date: `${y}-01`, end_date: `${y}-06`, contract_revenue: 0, target_hurdle_rate: 18.0,
      settings: ensureMmAccounts(clone(settings || defaultSettings())),
      internal_labor: [], external_labor: [], general_expenses: [], custom_costs: [], phases: [], estimate: emptyEstimate()
    };
  }

  // 화면 확인 · 교육용 예시 (원본의 [데모 채우기]와 같은 자료)
  function demoProject(settings) {
    const p = newProject(settings, "PRJ-2026-01");
    // 데모는 인비례 기준금액을 채워 자동 계산이 보이게
    p.settings.expense_accounts.forEach(a => { if (a.category === "mm_internal" && !num(a.unit_amount)) a.unit_amount = 300000; if (a.category === "mm_external" && !num(a.unit_amount)) a.unit_amount = 100000; });
    Object.assign(p, {
      project_name: "차세대 원가·공정 시스템 구축", client_name: "대한정밀(주)", pm_name: "김철수 PM", sales_rep: "박영희 매니저",
      description: "1) 착수 전 원가 산정 및 수익성 검토\n2) 1차 선행 분석/설계 후 2차 본 구축 진행",
      start_date: "2026-01", end_date: "2026-06", contract_revenue: 240000000, target_hurdle_rate: 18.0,
      internal_labor: [
        { name: "김철수", role: "PM", area: "PMO", grade: "수석", base_rate: 8500000, overhead_eligible: true, target_mm: 6, rate_changes: [{ effective_ym: "2026-03", rate: 8800000 }], promotions: [{ effective_ym: "2026-05", grade: "", rate: 9500000 }],
          mm: { "2026-01": 1, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 1 } },
        { name: "이민지", role: "PL/아키텍트", area: "분석 · 설계", grade: "선임", base_rate: 6000000, overhead_eligible: false, target_mm: 4.5, promotions: [],
          mm: { "2026-01": 0.5, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 0 } }
      ],
      external_labor: [{ name: "외주개발1", role: "백엔드 개발", area: "개발", vendor: "파트너A", grade: "고급", custom_rate: "", target_mm: 4, rate_changes: [{ effective_ym: "2026-04", rate: 8000000 }], mm: { "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1 } }],
      general_expenses: [{ name: "출장비", dist_mode: "equal_total", total_input: 9000000, unit_input: 0, monthly: {}, memo: "지방 사업장 방문" }],
      custom_costs: [{ name: "사전심의·인증비", dist_mode: "fill_all", total_input: 0, unit_input: 400000, monthly: {}, memo: "보안 점검" }],
      phases: [
        { title: "1차 선행계약", start_ym: "2026-01", end_ym: "2026-03", revenue: 110000000, override_cost: "" },
        { title: "2차 본계약", start_ym: "2026-04", end_ym: "2026-06", revenue: 130000000, override_cost: "" }
      ],
      estimate: { estimate_no: "EST-2026-0101", estimate_date: "2026-01-05", note: "공급가액 기준", applied_to_revenue: false,
        items: [{ item_name: "컨설팅 및 설계", spec: "수석/선임", qty: 1, unit_price: 95000000, memo: "1차 범위" }, { item_name: "시스템 개발 및 이행", spec: "구축 일괄", qty: 1, unit_price: 155000000, memo: "2차 범위" }] }
    });
    return p;
  }

  // 파일에서 읽은 프로젝트의 빠진 부분을 채움 (원본 v4 파일 · 손으로 고친 파일)
  function normalize(p, master) {
    const base = newProject(master);
    const out = Object.assign(base, p || {});
    if (!out.settings || !Object.keys(out.settings).length) out.settings = clone(master || defaultSettings());
    out.settings = Object.assign(defaultSettings(), out.settings);
    ensureMmAccounts(out.settings);
    if (!Array.isArray(out.settings.areas)) out.settings.areas = defaultSettings().areas;
    ["internal_labor", "external_labor", "general_expenses", "custom_costs", "phases"].forEach(k => { if (!Array.isArray(out[k])) out[k] = []; });
    // 단가 인상 rate_changes(내부 연봉 인상 · 외주 협상)와 내부 승진 promotions(v4 이름 그대로)는 따로. 단가도 등급도 없는 빈 줄(예전 새 인력 기본값)은 뺌
    const list = v => (Array.isArray(v) ? v : []).filter(r => r && (!isBlank(r.rate) || r.grade));
    out.internal_labor.forEach(m => { m.promotions = list(m.promotions); m.rate_changes = list(m.rate_changes); });
    out.external_labor.forEach(e => { e.rate_changes = list(e.rate_changes); });
    out.estimate = Object.assign(emptyEstimate(), out.estimate || {});
    if (!Array.isArray(out.estimate.items)) out.estimate.items = [];
    return out;
  }

  // 예전(v4) 비율 방식 인비례 경비가 있던 자료인지 — 열 때 한 번 알림 (이제는 계정별 MM × 기준금액)
  function legacyRatioNote(p) {
    const st = (p && p.settings) || {};
    const ratios = [["내부", st.internal_overhead_ratio], ["외주", st.external_overhead_ratio]].filter(([, r]) => num(r) > 0);
    const hasUnit = (st.expense_accounts || []).some(a => isMmCategory(a.category) && num(a.unit_amount) > 0);
    return ratios.length && !hasUnit ? `이 프로젝트는 예전 방식(인건비의 ${ratios.map(([k, r]) => `${k} ${(num(r) * 100).toFixed(1)}%`).join(" · ")})으로 인비례 경비를 계산했습니다. 이제는 [기준정보 > 경비 계정]의 인비례 계정에 기준금액(원/MM)을 넣어야 계산됩니다.` : "";
  }

  // 업로드 양식(Project_Upload_Template): 첫 시트의 A열 = 키, B열 = 값
  function fromTemplateRows(rows, master) {
    const kv = {};
    (rows || []).forEach(r => { if (r && r[0] !== undefined && r[0] !== "") kv[String(r[0]).trim()] = r[1]; });
    const p = newProject(master, String(kv.project_id || "PRJ-UPLOAD"));
    p.project_name = String(kv.project_name || `${p.project_id} 프로젝트`);
    ["client_name", "pm_name", "sales_rep", "description"].forEach(k => { p[k] = String(kv[k] || ""); });
    p.start_date = String(kv.start_date || kv.start_ym || p.start_date);
    p.end_date = String(kv.end_date || kv.end_ym || p.end_date);
    if (p.start_date.length > 7 || p.end_date.length > 7) p.date_mode = "date";
    p.contract_revenue = num(kv.contract_revenue);
    p.target_hurdle_rate = kv.target_hurdle_rate === undefined || kv.target_hurdle_rate === "" ? 18.0 : num(kv.target_hurdle_rate);
    return p;
  }

  PC.Model = { CATEGORIES, isMmCategory, defaultSettings, newProject, demoProject, normalize, fromTemplateRows, emptyEstimate, legacyRatioNote };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Model;
})(typeof window !== "undefined" ? window : globalThis);
