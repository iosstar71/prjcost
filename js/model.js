// =============================================================================
// 자료 모양 — 기본 기준정보 · 새 프로젝트 · 데모 · 예전 자료 채우기 · 업로드 양식 읽기
//  자료 정의는 docs/DATA_MODEL.md. 필드를 바꾸면 그 문서와 normalize() 를 함께 고친다
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const { clone, num, isBlank, pyRound, monthRange } = PC;

  // 경비 계정 구분: 일반경비 · 기타경비 (금액을 직접 넣음) / 인비례 (인력 MM × 기준금액으로 자동)
  const CATEGORIES = [
    { v: "general", t: "일반경비" }, { v: "custom", t: "기타경비" },
    { v: "mm_internal", t: "인비례 · 내부" }, { v: "mm_external", t: "인비례 · 외주" }
  ];
  const isMmCategory = c => c === "mm_internal" || c === "mm_external";

  // 기준정보가 하나도 없을 때(새 폴더) 쓰는 값. 인비례 기준금액(원/MM)은 회사 기준으로 넣어 쓰도록 0
  function defaultSettings() {
    return {
      raise_month: "",
      departments: [],           // 부서 — 2개 이상이면 부서 칸 · [부서별] 탭이 나타남
      dept_common_basis: "cost", // 공통경비(부서를 정하지 않은 경비)를 나누는 기준   // 연봉 인상월(MM) — 등급의 인상단가가 이 달부터 (비우면 인상 없음)
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
      internal_labor: [], external_labor: [], general_expenses: [], custom_costs: [], phases: [], estimate: emptyEstimate(), versions: [], dept_revenue: {}
    };
  }

  // 화면 확인 · 교육용 예시 (원본의 [데모 채우기]와 같은 자료)
  function demoProject(settings) {
    const p = newProject(settings, "PRJ-2026-01");
    // 데모: 3월 연봉 인상(수석 · 선임), 외주 파트너A 고급은 4월부터 협상 단가
    p.settings.raise_month = "03";
    p.settings.internal_grades.forEach(g => { if (g.grade === "수석") g.raise_rate = 8800000; if (g.grade === "선임") g.raise_rate = 6200000; });
    p.settings.vendor_rates.forEach(v => { if (v.vendor === "파트너A" && v.grade === "고급") Object.assign(v, { raise_ym: "2026-04", raise_rate: 8000000 }); });
    // 데모는 인비례 기준금액을 채워 자동 계산이 보이게
    p.settings.expense_accounts.forEach(a => { if (a.category === "mm_internal" && !num(a.unit_amount)) a.unit_amount = 300000; if (a.category === "mm_external" && !num(a.unit_amount)) a.unit_amount = 100000; });
    Object.assign(p, {
      project_name: "차세대 원가·공정 시스템 구축", client_name: "대한정밀(주)", pm_name: "김철수 PM", sales_rep: "박영희 매니저",
      description: "1) 착수 전 원가 산정 및 수익성 검토\n2) 1차 선행 분석/설계 후 2차 본 구축 진행",
      start_date: "2026-01", end_date: "2026-06", contract_revenue: 240000000, target_hurdle_rate: 18.0,
      internal_labor: [
        { name: "김철수", role: "PM", area: "PMO", grade: "수석", base_rate: 8500000, overhead_eligible: true, target_mm: 6,
          mm: { "2026-01": 1, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 1 } },
        { name: "이민지", role: "PL/아키텍트", area: "분석 · 설계", grade: "선임", base_rate: 6000000, overhead_eligible: false, target_mm: 4.5,
          mm: { "2026-01": 0.5, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 0 } }
      ],
      external_labor: [{ name: "외주개발1", role: "백엔드 개발", area: "개발", vendor: "파트너A", grade: "고급", custom_rate: "", target_mm: 4, mm: { "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1 } }],
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
    // 단가 인상은 기준정보(내부 등급의 인상단가 + 연봉 인상월, 외주업체 단가의 인상월 · 인상단가)뿐
    //  예전 사람별 승진(promotions, v4) · 인상(rate_changes, v0.9.1)은 더 쓰지 않음 — 열 때 알리고(legacyNote) 지움
    const list = v => (Array.isArray(v) ? v : []).filter(r => r && (!isBlank(r.rate) || r.grade));
    out.internal_labor.forEach(m => { delete m.promotions; delete m.rate_changes; });
    delete out.settings.default_promotion_month;
    // 총 MM: 달별 계획의 합 (예전 파일에는 따로 넣은 값이 있을 수 있음)
    const months = monthRange(String(out.start_date || "2026-01"), String(out.end_date || "2026-06"));
    out.internal_labor.concat(out.external_labor).forEach(r => { r.target_mm = pyRound(months.reduce((a, m) => a + num((r.mm || {})[m]), 0), 2); });
    out.external_labor.forEach(e => {   // v0.9.1 의 사람별 협상 단가 → 그 업체 · 등급 단가의 인상(비어 있을 때)
      const r0 = list(e.rate_changes)[0], vr = (out.settings.vendor_rates || []).find(v => v.vendor === e.vendor && v.grade === e.grade);
      if (r0 && vr && isBlank(vr.raise_rate)) Object.assign(vr, { raise_ym: r0.effective_ym || "", raise_rate: num(r0.rate) });
      delete e.rate_changes;
    });
    if (out.settings.raise_month === undefined) out.settings.raise_month = "";
    if (!Array.isArray(out.versions)) out.versions = [];
    if (!Array.isArray(out.settings.departments)) out.settings.departments = [];
    if (!out.settings.dept_common_basis) out.settings.dept_common_basis = "cost";
    if (!out.dept_revenue || typeof out.dept_revenue !== "object") out.dept_revenue = {};
    out.estimate = Object.assign(emptyEstimate(), out.estimate || {});
    if (!Array.isArray(out.estimate.items)) out.estimate.items = [];
    return out;
  }

  // 부서가 2개 이상일 때만 부서 칸 · 탭을 보여 줌 (하나 이하면 지금과 같은 화면)
  const multiDept = p => ((p && p.settings && p.settings.departments) || []).filter(d => String(d).trim()).length >= 2;

  // 버전(기준선): 저장할 때 이름 · 메모를 붙여 그때의 자료를 통째로 남김 (초기 계획 · 실계약 · 수정계약 …) — project.versions[]
  const VERSION_NAMES = ["초기 계획", "실계약", "수정계약", "2차 수정계약", "3차 수정계약"];
  function snapshot(p) { const d = clone(p); delete d.versions; return d; }
  function nextVersionName(p) {
    const used = new Set((p.versions || []).map(v => v.name));
    return VERSION_NAMES.find(n => !used.has(n)) || `${(p.versions || []).length + 1}차 수정계약`;
  }
  // 버전 자료를 지금 프로젝트로: 자료는 그 버전, 버전 목록 · 파일 이름(ID)은 지금 것
  function restoreVersion(p, v, master) {
    const out = normalize(clone(v.data), master);
    out.versions = p.versions;
    out.project_id = p.project_id;
    return out;
  }

  // 파일의 원래 자료(normalize 전)에 더 쓰지 않는 계산이 있었는지 — 열 때 한 번 알림
  //  ① 예전(v4) 비율 방식 인비례 경비 (이제는 계정별 MM × 기준금액) ② 사람별 승진 · 인상 단가 (이제는 기준정보의 등급 인상만)
  function legacyNote(raw) {
    const st = (raw && raw.settings) || {}, notes = [];
    const ratios = [["내부", st.internal_overhead_ratio], ["외주", st.external_overhead_ratio]].filter(([, r]) => num(r) > 0);
    const hasUnit = (st.expense_accounts || []).some(a => isMmCategory(a.category) && num(a.unit_amount) > 0);
    if (ratios.length && !hasUnit) notes.push(`인비례 경비: 예전 방식(인건비의 ${ratios.map(([k, r]) => `${k} ${(num(r) * 100).toFixed(1)}%`).join(" · ")})은 더 쓰지 않습니다. [기준정보 > 경비 계정]의 인비례 계정에 기준금액(원/MM)을 넣으세요.`);
    const promos = ((raw && raw.internal_labor) || []).flatMap(m => (m.promotions || []).concat(m.rate_changes || []).filter(r => r && (!isBlank(r.rate) || r.grade)).map(r => `${m.name || "?"} ${r.effective_ym || ""} ${isBlank(r.rate) ? r.grade : `${num(r.rate).toLocaleString("ko-KR")}원`}`));
    if (promos.length) notes.push(`승진 단가(${promos.join(", ")})는 더 쓰지 않아 기준단가로 계산합니다. 단가가 바뀌어야 하면 [기준정보 > 내부 등급 · 단가]의 연봉 인상으로 넣으세요. 저장하면 파일에서도 빠집니다.`);
    return notes.join("\n\n");
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

  PC.Model = { multiDept, VERSION_NAMES, snapshot, nextVersionName, restoreVersion, CATEGORIES, isMmCategory, defaultSettings, newProject, demoProject, normalize, fromTemplateRows, emptyEstimate, legacyNote };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Model;
})(typeof window !== "undefined" ? window : globalThis);
