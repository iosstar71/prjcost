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
      dept_common_basis: "cost", // 공통경비(부서를 정하지 않은 경비)를 나누는 기준
      approvers: ["", "", "", "", ""],   // 보고서 결재자 (기안(PM) · 검토 · 사업팀장 · 영업팀장 · 본부장) — 인사이동이 있을 때만 바꿈
      areas: ["PMO", "분석 · 설계", "개발", "테스트", "인프라"],   // 인력의 영역 (마스터)
      systems: [],                // 대표 시스템 (프로젝트마다) — 인력마다 하나
      // 고객 제안 직무 × 등급 → 제안 단가 (견적서). 등급 체계는 모든 직무가 같음
      proposal_levels: ["Lvl1", "Lvl2", "Lvl3", "Lvl4", "Lvl5"],
      proposal_roles: [
        { role: "프로젝트관리", rates: { Lvl1: 13000000, Lvl2: 11500000, Lvl3: 10000000, Lvl4: 8500000, Lvl5: 7000000 } },
        { role: "컨설팅", rates: { Lvl1: 14000000, Lvl2: 12000000, Lvl3: 10500000, Lvl4: 9000000, Lvl5: 7500000 } },
        { role: "개발", rates: { Lvl1: 11000000, Lvl2: 9500000, Lvl3: 8000000, Lvl4: 6500000, Lvl5: 5500000 } },
        { role: "테스트", rates: { Lvl1: 9000000, Lvl2: 8000000, Lvl3: 7000000, Lvl4: 6000000, Lvl5: 5000000 } }
      ],
      master_version: { name: "", date: "" },   // 마스터를 저장 · 배포할 때 붙이는 이름 · 날짜
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

  // ---------- 마스터 (전사에서 같이 쓰는 기준정보 — master/master_rates.xlsx 로 배포) ----------
  //  마스터 = 전사 공통 목록(이름)뿐 — 금액(직급 단가 · 인상단가 · 제안 단가 · 인비례 기준금액)과 프로젝트마다 다른 것(외주업체 단가 · 대표 시스템 · 연봉 인상월 · 공통경비 기준)은 넣지 않음
  const PROJECT_ONLY = ["vendor_rates", "external_rates", "systems", "dept_common_basis", "raise_month"];
  function masterOf(settings) {
    const m = clone(settings || {});
    PROJECT_ONLY.forEach(k => { delete m[k]; });
    m.internal_grades = (m.internal_grades || []).map(g => ({ grade: g.grade }));
    m.proposal_roles = (m.proposal_roles || []).map(r => ({ role: r.role }));
    m.expense_accounts = (m.expense_accounts || []).map(a => ({ account: a.account, category: a.category, desc: a.desc || "" }));
    return m;
  }
  const RATE_LISTS = [
    ["internal_grades", "grade", "직급", []],
    ["proposal_roles", "role", "제안 직무", []],
    ["expense_accounts", "account", "경비 계정", ["category", "desc"]]
  ];
  const NAME_LISTS = [["external_grades", "외주 기술등급"], ["proposal_levels", "제안 등급"], ["areas", "영역"], ["departments", "부서"]];
  const same = (a, b) => JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b);
  const show = v => (v && typeof v === "object" ? Object.entries(v).map(([k, x]) => `${k} ${num(x).toLocaleString("ko-KR")}`).join(", ") : typeof v === "number" ? v.toLocaleString("ko-KR") : String(v === undefined || v === "" ? "-" : v));
  // 마스터 b 를 a 에 맞춰 넣었을 때 바뀌는 것 (목록은 더하고 · 같은 이름은 값을 바꿈 — a 에만 있는 것은 그대로 둠, removed: b 에 없는 것)
  function masterChanges(a, b) {
    const out = [];
    RATE_LISTS.forEach(([k, key, label, fields]) => {
      const al = a[k] || [], bl = b[k] || [];
      bl.forEach(x => {
        const y = al.find(z => z[key] === x[key]);
        if (!y) out.push(`${label} 추가: ${x[key]}`);
        else fields.forEach(f => { if (!same(y[f], x[f]) && !(f === "rates" && same(Object.assign({}, y.rates, x.rates), y.rates))) out.push(`${label} ${x[key]}: ${show(y[f])} → ${show(x[f])}`); });
      });
      al.filter(y => !bl.some(x => x[key] === y[key])).forEach(y => out.push(`(마스터에 없음) ${label}: ${y[key]}`));
    });
    NAME_LISTS.forEach(([k, label]) => {
      const al = a[k] || [], bl = b[k] || [];
      bl.filter(x => !al.includes(x)).forEach(x => out.push(`${label} 추가: ${x}`));
      al.filter(x => !bl.includes(x)).forEach(x => out.push(`(마스터에 없음) ${label}: ${x}`));
    });
    if ((b.approvers || []).some(Boolean) && !same(a.approvers, b.approvers)) out.push(`결재자: ${(a.approvers || []).filter(Boolean).join(" · ") || "-"} → ${(b.approvers || []).filter(Boolean).join(" · ")}`);
    return out;
  }
  // 프로젝트 기준정보에 마스터를 반영: 마스터에만 있는 항목을 더하고 (경비 계정은 구분 · 설명도 마스터대로), 프로젝트의 항목 · 금액 · 프로젝트 전용 정보는 그대로
  function mergeMaster(settings, master) {
    const out = clone(settings || {}), m = master || {};
    RATE_LISTS.forEach(([k, key, , fields]) => {
      const list = out[k] = Array.isArray(out[k]) ? out[k] : [];
      (m[k] || []).forEach(x => {
        const y = list.find(z => z[key] === x[key]);
        if (!y) list.push(Object.assign(k === "proposal_roles" ? { rates: {} } : {}, clone(x)));
        else fields.forEach(f => { if (x[f] !== undefined) y[f] = clone(x[f]); });
      });
    });
    NAME_LISTS.forEach(([k]) => { const list = out[k] = Array.isArray(out[k]) ? out[k] : []; (m[k] || []).forEach(x => { if (!list.includes(x)) list.push(x); }); });
    if ((m.approvers || []).some(Boolean)) out.approvers = clone(m.approvers);
    if (m.master_version) out.master_version = clone(m.master_version);
    return out;
  }

  // 인비례 계정이 없는 기준정보(예전 v4 파일 등)에 기본 인비례 계정(기준금액 0)을 넣음 — 화면에서 바로 채울 수 있게
  function ensureMmAccounts(settings) {
    const accs = settings.expense_accounts = Array.isArray(settings.expense_accounts) ? settings.expense_accounts : [];
    defaultSettings().expense_accounts.filter(a => isMmCategory(a.category)).forEach(a => {
      if (!accs.some(x => x.category === a.category)) accs.push(Object.assign({}, a));
    });
    return settings;
  }

  const emptyEstimate = () => ({ estimate_no: "", estimate_date: "", note: "", applied_to_revenue: false, labor: [], items: [] });

  function newProject(settings, id) {
    const y = new Date().getFullYear();
    const st = ensureMmAccounts(clone(settings || defaultSettings()));
    ["vendor_rates", "systems"].forEach(k => { if (!Array.isArray(st[k])) st[k] = []; });
    if (st.raise_month === undefined) st.raise_month = "";
    return {
      project_id: id || `PRJ-${y}-NEW`, project_name: "신규 프로젝트", client_name: "", pm_name: "", sales_rep: "", description: "",
      date_mode: "month", start_date: `${y}-01`, end_date: `${y}-06`, contract_revenue: 0, target_hurdle_rate: 18.0,
      settings: st,
      internal_labor: [], external_labor: [], general_expenses: [], custom_costs: [], phases: [], estimate: emptyEstimate(), versions: [], dept_revenue: {},
      project_type: ""
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
    p.settings.systems = ["원가관리", "공정관리"];
    Object.assign(p, {
      project_name: "차세대 원가·공정 시스템 구축", client_name: "대한정밀(주)", pm_name: "김철수 PM", sales_rep: "박영희 매니저",
      description: "1) 착수 전 원가 산정 및 수익성 검토\n2) 1차 선행 분석/설계 후 2차 본 구축 진행",
      start_date: "2026-01", end_date: "2026-06", contract_revenue: 240000000, target_hurdle_rate: 18.0,
      internal_labor: [
        { name: "김철수", role: "PM", area: "PMO", system: "원가관리", prop_role: "프로젝트관리", prop_level: "Lvl1", grade: "수석", base_rate: 8500000, overhead_eligible: true, target_mm: 6,
          mm: { "2026-01": 1, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 1 } },
        { name: "이민지", role: "PL/아키텍트", area: "분석 · 설계", system: "공정관리", prop_role: "컨설팅", prop_level: "Lvl2", grade: "선임", base_rate: 6000000, overhead_eligible: false, target_mm: 4.5,
          mm: { "2026-01": 0.5, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 0 } }
      ],
      external_labor: [{ name: "외주개발1", role: "백엔드 개발", area: "개발", system: "공정관리", prop_role: "개발", prop_level: "Lvl3", vendor: "파트너A", grade: "고급", custom_rate: "", target_mm: 4, mm: { "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1 } }],
      general_expenses: [{ name: "출장비", dist_mode: "equal_total", total_input: 9000000, unit_input: 0, monthly: {}, memo: "지방 사업장 방문" }],
      custom_costs: [{ name: "사전심의·인증비", dist_mode: "fill_all", total_input: 0, unit_input: 400000, monthly: {}, memo: "보안 점검" }],
      phases: [
        { title: "1차 선행계약", start_ym: "2026-01", end_ym: "2026-03", revenue: 110000000, override_cost: "" },
        { title: "2차 본계약", start_ym: "2026-04", end_ym: "2026-06", revenue: 130000000, override_cost: "" }
      ],
      estimate: { estimate_no: "EST-2026-0101", estimate_date: "2026-01-05", note: "공급가액 기준", applied_to_revenue: false,
        labor: [{ role: "프로젝트관리", level: "Lvl1", mm: 6, unit_price: "", memo: "" }, { role: "컨설팅", level: "Lvl2", mm: 5, unit_price: "", memo: "분석 · 설계" }, { role: "개발", level: "Lvl3", mm: 8, unit_price: "", memo: "" }],
        items: [{ item_name: "원가관리 SW 라이선스", spec: "1식", qty: 1, unit_price: 30000000, memo: "" }] }
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
    ["systems", "proposal_levels", "proposal_roles"].forEach(k => { if (!Array.isArray(out.settings[k])) out.settings[k] = k === "systems" ? [] : clone(defaultSettings()[k]); });
    out.settings.proposal_roles.forEach(r => { if (!r.rates || typeof r.rates !== "object") r.rates = {}; });
    if (!out.settings.master_version) out.settings.master_version = { name: "", date: "" };
    out.internal_labor.concat(out.external_labor).forEach(r => { ["system", "prop_role", "prop_level"].forEach(f => { if (r[f] === undefined) r[f] = ""; }); });
    if (!Array.isArray(out.estimate.labor)) out.estimate.labor = [];   // 견적 인건비 (제안 직무 · 등급 · 견적 MM)
    // 결재자는 기준정보(settings.approvers). v0.9.7 에서 프로젝트에 넣었던 이름은 기준정보가 비어 있으면 옮김
    const ap = i => String(((Array.isArray(out.settings.approvers) ? out.settings.approvers : [])[i]) || "");
    const old = Array.isArray(out.approvers) ? out.approvers : [];
    out.settings.approvers = [0, 1, 2, 3, 4].map(i => ap(i) || String(old[i] || ""));
    delete out.approvers;
    if (typeof out.project_type !== "string") out.project_type = "";
    if (!Array.isArray(out.settings.departments)) out.settings.departments = [];
    if (!out.settings.dept_common_basis) out.settings.dept_common_basis = "cost";
    if (!out.dept_revenue || typeof out.dept_revenue !== "object") out.dept_revenue = {};
    out.phases.forEach(ph => { if (!ph.dept_revenue || typeof ph.dept_revenue !== "object") ph.dept_revenue = {}; });   // 회차 × 부서 매출
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

  PC.Model = { PROJECT_ONLY, masterOf, masterChanges, mergeMaster, multiDept, VERSION_NAMES, snapshot, nextVersionName, restoreVersion, CATEGORIES, isMmCategory, defaultSettings, newProject, demoProject, normalize, fromTemplateRows, emptyEstimate, legacyNote };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Model;
})(typeof window !== "undefined" ? window : globalThis);
