// =============================================================================
// 프로그램 흐름 — 지금 프로젝트 · 입력 반영 · 저장 / 열기 / 지우기 · 폴더 연결
//  자료가 바뀔 때마다 그 자리에서 다시 계산하고 다시 그림 (서버 없음). 저장은 [저장] 또는 Ctrl+S 를 눌러야 기록
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC;
  const { Calc, Model: M, Store, UI, Views, Xlsx: X, Workbook: W, clone, won, esc, num, safeId } = PC;
  const APP_NAME = "PM Hub Cost", VERSION = "v1.0.2";

  const S = {
    master: M.defaultSettings(),   // 폴더의 기준정보 (새 프로젝트 기본값)
    proj: null,
    file: null,                    // projects 폴더 안의 파일 이름 (저장 전이면 null)
    mtime: null,                   // 불러온 때의 파일 수정 시각 (다른 곳에서 바뀌었는지 확인)
    saved: "",                     // 마지막으로 저장 · 불러온 자료 (바뀌었는지 비교)
    calc: null,
    tab: "info",
    list: [],
    past: [], future: [],
    noted: new Set(),
    settingsTab: "rates",          // 기준정보 안의 작은 탭
    stats: { by: "area", scope: "all", phase: 0 },
    deptPhase: 0,                   // 부서별 탭의 기간 (0 = 전체, n = n 회차)
    cmp: { a: null, b: 0 }          // 버전 비교: 기준 버전 번호 · 비교 대상 (0 = 지금 화면)   // 투입 현황 보기 조건 (phase: 분할계약 회차, 0 = 전체 기간)               // 이번에 이미 알린 '예전 비율 방식' 파일           // 되돌리기 · 다시 하기 (자료 JSON, 최근 UNDO_KEEP 개)
  };
  const UNDO_KEEP = 100;
  const $ = id => document.getElementById(id);
  const isDirty = () => JSON.stringify(S.proj) !== S.saved;
  const markSaved = () => { S.saved = JSON.stringify(S.proj); };
  // 자료를 바꾸기 직전에 부름 (같은 상태를 두 번 쌓지 않음)
  function pushUndo() {
    const cur = JSON.stringify(S.proj);
    if (S.past[S.past.length - 1] !== cur) { S.past.push(cur); if (S.past.length > UNDO_KEEP) S.past.shift(); }
    S.future = [];
  }
  function undo() { if (!S.past.length) return; S.future.push(JSON.stringify(S.proj)); S.proj = JSON.parse(S.past.pop()); render(); }
  function redo() { if (!S.future.length) return; S.past.push(JSON.stringify(S.proj)); S.proj = JSON.parse(S.future.pop()); render(); }

  // ---------------- 그리기 ----------------
  const renderTimes = [];
  function render(focusBind) {
    const now = Date.now();
    while (renderTimes.length && now - renderTimes[0] > 1000) renderTimes.shift();
    renderTimes.push(now);
    const storm = renderTimes.length > 40;
    if (storm && !render.warned) { render.warned = true; console.error("[App] render storm — focus restore paused"); UI.toast("화면이 너무 자주 다시 그려져 커서 되돌리기를 잠시 멈췄습니다. 칸을 다시 눌러 주세요.", "error"); setTimeout(() => { render.warned = false; }, 3000); }
    S.calc = Calc.calculateProjectCost(S.proj, S.master);
    const focus = storm ? null : focusBind ? { bind: focusBind } : captureFocus();
    renderKpis();
    renderTabs();
    const tab = shownTabs().find(t => t.id === S.tab) || Views.TABS[0];
    $("panel").innerHTML = tab.view(S.proj, S.calc, { file: S.file, sub: S.settingsTab, stats: S.stats, cmp: S.cmp, deptPhase: S.deptPhase, master: S.master });
    renderSaveState();
    renderUndo();
    restoreFocus(focus);
  }
  function renderUndo() { $("btnUndo").disabled = !S.past.length; $("btnRedo").disabled = !S.future.length; }
  function renderKpis() {
    const s = S.calc.summary, b = s.breakdown, est = S.calc.estimate;
    $("kRev").textContent = `${won(s.revenue)}원`;
    $("kRevSub").textContent = est.supply_total > 0 ? `견적 공급가 ${won(est.supply_total)}원${est.applied_to_revenue ? " · 매출 적용" : ""}` : "견적 미등록";
    $("kCost").textContent = `${won(s.total_cost)}원`;
    const labor = b.internal_labor + b.external_labor, exp = b.general_expense + b.internal_overhead + b.external_overhead + b.custom_cost;
    const nonLabor = exp + (b.material || 0);
    const m = v => (v / 1e6).toFixed(1);
    $("kCostSub").textContent = `내부 ${m(b.internal_labor)} · 외주 ${m(b.external_labor)}${b.material ? ` · 재료 ${m(b.material)}` : ""} · 경비 ${m(exp)} (백만 원)`;
    const mix = $("kMix");
    mix.innerHTML = `<i class="m1"></i><i class="m2"></i><i class="m3"></i>`;
    const tot = s.total_cost || 1;
    [b.internal_labor, b.external_labor, nonLabor].forEach((v, i) => { mix.children[i].style.width = `${Math.max(0, (v / tot) * 100)}%`; });
    mix.title = `인건비 ${won(labor)}원 · 재료비 · 경비 ${won(nonLabor)}원`;
    $("kMargin").textContent = `${s.margin_rate.toFixed(1)}%`;
    $("kProfit").textContent = `이익 ${won(s.profit)}원`;
    $("kMarginCard").className = `kpi ${s.revenue > 0 ? (s.hurdle_met ? "ok" : "bad") : ""}`;
    $("kHurdle").textContent = `${s.hurdle_gap >= 0 ? "+" : ""}${s.hurdle_gap.toFixed(1)}%p`;
    $("kHurdleSub").textContent = `목표 ${s.hurdle_rate}% · ${s.hurdle_met ? "달성" : "미달"}`;
    $("kHurdleCard").className = `kpi ${s.revenue > 0 ? (s.hurdle_met ? "ok" : "bad") : ""}`;
  }
  const shownTabs = () => Views.TABS.filter(t => !t.show || t.show(S.proj));   // 부서별 탭은 부서가 2개 이상일 때만
  function renderTabs() {
    $("tabs").innerHTML = shownTabs().map(t => {
      const n = t.count ? t.count(S.proj) : 0;
      return `<button class="tab${t.id === S.tab ? " active" : ""}" role="tab" aria-selected="${t.id === S.tab}" data-act="showTab" data-args="${esc(JSON.stringify([t.id]))}">${PC.icon(t.icon)}${esc(t.label)}${n ? `<span class="cnt">${n}</span>` : ""}</button>`;
    }).join("");
  }
  function renderSaveState() {
    const el = $("saveState"), dirty = isDirty();
    el.textContent = dirty ? "● 저장 안 됨" : S.file ? "저장됨" : "";
    el.className = `save-state${dirty ? " dirty" : ""}`;
    document.title = `${dirty ? "● " : ""}${S.proj.project_name || S.proj.project_id} — ${APP_NAME}`;
  }
  function renderFolder() {
    const chip = $("folderChip");
    chip.classList.toggle("off", !Store.dir);
    chip.querySelector(".name").textContent = Store.dir ? Store.label : "폴더 연결";
    chip.hidden = !Store.supported;
  }
  function renderList() {
    const sel = $("projectSelect");
    const opts = S.list.map(x => `<option value="${esc(x.file)}">${esc(x.id)} · ${esc(new Date(x.mtime).toLocaleDateString("ko-KR"))}</option>`);
    if (!S.file) opts.unshift(`<option value="">${esc(S.proj.project_id)} (저장 전)</option>`);
    if (!opts.length) opts.push(`<option value="">저장된 프로젝트 없음</option>`);
    sel.innerHTML = opts.join("");
    sel.value = S.file || "";
    sel.disabled = !Store.dir;
  }
  // 지금 테마의 반대로 바꾸는 버튼: 밝을 때 [Dark], 어두울 때 [Light]
  const renderTheme = () => { const dark = document.documentElement.dataset.theme === "dark"; $("btnTheme").innerHTML = `${PC.icon(dark ? "sun" : "moon")}<span class="lbl">${dark ? "Light" : "Dark"}</span>`; };
  function notice(html) { const n = $("noticeBox"); n.hidden = !html; n.innerHTML = html || ""; }

  // 다시 그려도 입력 중이던 칸을 잃지 않게
  function captureFocus() {
    const a = document.activeElement;
    return a && a.dataset && a.dataset.bind ? { bind: a.dataset.bind, start: a.selectionStart, end: a.selectionEnd } : null;
  }
  function restoreFocus(f) {
    if (!f) return;
    const el = [...document.querySelectorAll("[data-bind]")].find(e => e.dataset.bind === f.bind);
    if (!el) return;
    el.focus();
    try { if (f.start !== null && f.start !== undefined) el.setSelectionRange(f.start, f.end); } catch (e) { /* number · month 칸은 커서 위치가 없음 */ }
  }

  // ---------------- 입력 반영 ----------------
  function setPath(obj, path, value) {
    const keys = path.split(".");
    let o = obj;
    keys.slice(0, -1).forEach((k, i) => {
      if (o[k] === undefined || o[k] === null || typeof o[k] !== "object") o[k] = /^\d+$/.test(keys[i + 1]) ? [] : {};
      o = o[k];
    });
    o[keys[keys.length - 1]] = value;
  }
  function valueOf(el) {
    const t = el.dataset.type;
    if (t === "bool") return el.checked;
    const v = el.value;
    if (t === "num") return v === "" ? 0 : num(v);
    if (t === "mm") return v === "" ? 0 : PC.pyRound(num(v), 2);   // MM 은 소수 둘째 자리까지
    if (t === "blank-num") return v === "" ? "" : num(v);
    if (t === "pct") return num(v) / 100;
    if (t === "month2b") return /\d/.test(v) ? String(v).replace(/\D/g, "").padStart(2, "0").slice(-2) : "";   // 비울 수 있는 월(MM)
    if (t === "month2") return String(v || "03").replace(/\D/g, "").padStart(2, "0").slice(-2);
    return el.tagName === "TEXTAREA" ? v : v.trim();
  }
  // 값을 넣은 뒤 함께 바꿀 것 (원본과 같은 동작)
  function afterBind(path, value) {
    const p = S.proj;
    let m;
    // 부서 이름은 자료 경로(dept_revenue.<부서>)에 쓰여서 '.' 대신 '·'
    if (/^settings\.proposal_levels\.\d+$/.test(path) && /\./.test(value)) { const v2 = String(value).replace(/\./g, "·"); setPath(p, path, v2); renameRefs(path, value, v2); }   // 등급 이름도 자료 경로(rates.<등급>)에 쓰임
    if (/^settings\.departments\.\d+$/.test(path) && /\./.test(value)) { const v2 = String(value).replace(/\./g, "·"); setPath(p, path, v2); renameRefs(path, value, v2); }
    // 부서가 2개가 되는 순간: 부서가 비어 있는 인력은 첫 부서로 (그전까지 모두 한 부서였으므로). 경비의 빈 부서는 '공통' 그대로
    if (/^settings\.departments\.\d+$/.test(path) && M.multiDept(p)) p.internal_labor.concat(p.external_labor).forEach(r => { if (!r.dept) r.dept = p.settings.departments[0]; });
    if ((m = /^internal_labor\.(\d+)\.grade$/.exec(path))) {   // 등급을 바꾸면 기준단가도 그 등급 단가로
      const g = (p.settings.internal_grades || []).find(x => x.grade === value);
      if (g) p.internal_labor[+m[1]].base_rate = num(g.rate);
    } else if ((m = /^(internal_labor|external_labor)\.(\d+)\.mm\./.exec(path))) {   // 달별 MM 을 고치면 총 MM 은 그 합
      syncTotal(p[m[1]][+m[2]]);
    } else if ((m = /^(general_expenses|custom_costs)\.(\d+)\.dist_mode$/.exec(path)) && value === "manual") {   // 월별 직접 입력으로: 지금 달별 금액에서 시작
      const r = p[m[1]][+m[2]], prev = S.calc && S.calc.details[m[1]][+m[2]];
      if (prev && !Object.values(r.monthly || {}).some(v => num(v))) r.monthly = Object.assign({}, prev.calc_monthly);
    } else if (path === "date_mode") {   // 월 ↔ 일자: 값 모양만 바꿈
      const toDate = (v, end) => (String(v || "").length === 7 ? `${v}-${end ? "28" : "01"}` : v);
      p.start_date = value === "date" ? toDate(p.start_date, false) : String(p.start_date || "").slice(0, 7);
      p.end_date = value === "date" ? toDate(p.end_date, true) : String(p.end_date || "").slice(0, 7);
    }
  }
  // 기준정보의 담당 영역 · 부서 이름을 바꾸면 그 이름을 쓰던 인력 · 경비 · 부서 매출도 함께
  function renameRefs(path, oldV, newV) {
    const p = S.proj;
    if (!oldV || oldV === newV) return;
    const people = p.internal_labor.concat(p.external_labor);
    if (/^settings\.systems\.\d+$/.test(path)) return people.forEach(r => { if (r.system === oldV) r.system = newV; });
    if (/^settings\.work_sites\.\d+$/.test(path)) return people.forEach(r => { if (r.site === oldV) r.site = newV; });
    if (/^settings\.proposal_levels\.\d+$/.test(path)) {
      (p.settings.proposal_roles || []).forEach(r => { if (r.rates && oldV in r.rates) { r.rates[newV] = r.rates[oldV]; delete r.rates[oldV]; } });
      people.forEach(r => { if (r.prop_level === oldV) r.prop_level = newV; });
      return (p.estimate.labor || []).forEach(l => { if (l.level === oldV) l.level = newV; });
    }
    if (/^settings\.proposal_roles\.\d+\.role$/.test(path)) {
      people.forEach(r => { if (r.prop_role === oldV) r.prop_role = newV; });
      return (p.estimate.labor || []).forEach(l => { if (l.role === oldV) l.role = newV; });
    }
    if (/^settings\.internal_grades\.\d+\.grade$/.test(path)) return p.internal_labor.forEach(r => { if (r.grade === oldV) r.grade = newV; });
    if (/^settings\.external_grades\.\d+$/.test(path)) { p.external_labor.forEach(r => { if (r.grade === oldV) r.grade = newV; }); p.internal_labor.forEach(r => { if (r.tech_grade === oldV) r.tech_grade = newV; }); return (p.settings.vendor_rates || []).forEach(v => { if (v.grade === oldV) v.grade = newV; }); }
    if (/^settings\.expense_accounts\.\d+\.account$/.test(path)) return p.general_expenses.concat(p.custom_costs).forEach(e => { if (e.name === oldV) e.name = newV; });
    const m = /^settings\.(areas|departments)\.\d+$/.exec(path);
    if (!m) return;
    const f = m[1] === "areas" ? "area" : "dept";
    const lists = f === "area" ? [p.internal_labor, p.external_labor] : [p.internal_labor, p.external_labor, p.general_expenses, p.custom_costs];
    lists.forEach(l => l.forEach(r => { if (r[f] === oldV) r[f] = newV; }));
    if (f === "dept") [p.dept_revenue].concat(p.phases.map(ph => ph.dept_revenue)).forEach(m => { if (m && oldV in m) { m[newV] = m[oldV]; delete m[oldV]; } });
  }
  // 총 MM = 프로젝트 기간 달별 MM 의 합
  function syncTotal(row) {
    const months = (S.calc && S.calc.months) || [];
    row.target_mm = PC.pyRound(months.reduce((a, mo) => a + num((row.mm || {})[mo]), 0), 2);
  }
  // 총 MM 을 고른 달들에 나눔: 0.01 단위로 내리고 남는 것은 마지막 달에
  function spread(total, months) {
    const per = Math.floor((total / months.length) * 100) / 100;
    const rest = Math.round((total - per * months.length) * 100) / 100;
    return Object.fromEntries(months.map((m, k) => [m, k === months.length - 1 ? Math.round((per + rest) * 100) / 100 : per]));
  }
  function onInput(e) {
    const el = e.target.closest("[data-bind]");
    if (!el) return;
    const v = valueOf(el);
    const cur = el.dataset.bind.split(".").reduce((o, k) => (o === undefined || o === null ? undefined : o[k]), S.proj);
    if (cur === v || (cur === undefined && (v === "" || v === 0 || v === false))) return;   // 값이 그대로 (다시 그리며 사라진 칸의 늦은 change 등)
    pushUndo();
    setPath(S.proj, el.dataset.bind, v);
    renameRefs(el.dataset.bind, cur, v);
    afterBind(el.dataset.bind, v);
    // Tab · 클릭으로 다음 칸에 갈 때: 포커스가 옮겨진 뒤에 다시 그려야 그 칸을 잃지 않음
    clearTimeout(onInput.t);
    onInput.t = setTimeout(render, 0);
  }

  // ---------------- 동작 ----------------
  const ROW = {
    internal_labor: p => {
      const g0 = (p.settings.internal_grades || [])[0] || { grade: "선임", rate: 6000000 };
      return { name: "신규인력", role: "개발", area: "", system: "", prop_role: "", prop_level: "", site: "", tech_grade: "", dept: firstDept(p), grade: g0.grade, base_rate: num(g0.rate), overhead_eligible: true, target_mm: 0,
        mm: {} };
    },
    external_labor: p => {
      const v0 = (p.settings.vendor_rates || [])[0] || { vendor: "파트너A", grade: "고급" };
      return { name: "외주인력", role: "개발", area: "", system: "", prop_role: "", prop_level: "", site: "", dept: firstDept(p), vendor: v0.vendor, grade: v0.grade, custom_rate: "", target_mm: 0, mm: {} };
    },
    general_expenses: p => expenseRow(p, "general", "일반경비"),
    custom_costs: p => expenseRow(p, "custom", "기타비용"),
    phases: p => ({ title: `${p.phases.length + 1}차 계약`, start_ym: String(p.start_date || "").slice(0, 7), end_ym: String(p.end_date || "").slice(0, 7), revenue: 0, override_cost: "" }),
    "estimate.items": () => ({ item_name: "신규 품목", spec: "", qty: 1, unit_price: 0, memo: "" }),
    "estimate.labor": () => ({ role: "", level: "", mm: 0, unit_price: "", memo: "" }),
    "settings.proposal_levels": () => "새 등급",
    "settings.proposal_roles": () => ({ role: "새 직무", rates: {} }),
    "settings.systems": () => "새 시스템",
    "settings.work_sites": () => "새 근무지",
    materials: () => ({ category: "재료", content: "", amount: 0 }),
    "settings.internal_grades": () => ({ grade: "신규등급", rate: 5000000 }),
    "settings.external_grades": () => "신규등급",
    "settings.vendor_rates": p => ({ vendor: "신규업체", grade: (p.settings.external_grades || [])[0] || "고급", rate: 7500000 }),
    "settings.expense_accounts": () => ({ account: "신규계정", category: "general", desc: "" }),
    "settings.areas": () => "새 영역",
    "settings.departments": () => "새 부서"
  };
  const firstDept = p => (M.multiDept(p) ? p.settings.departments[0] : "");
  function expenseRow(p, cat, fallback) {
    const acc = (p.settings.expense_accounts || []).find(a => a.category === cat);
    return { name: acc ? acc.account : fallback, dist_mode: "equal_total", total_input: 0, unit_input: 0, monthly: {}, memo: "" };
  }
  const listAt = path => path.split(".").reduce((o, k) => o[k], S.proj);

  const A = {
    showTab(id) { S.tab = id; render(); },
    showSettings(id) { S.settingsTab = id; render(); },
    setStats(key, v) { S.stats[key] = key === "phase" ? Number(v) || 0 : v; render(); },
    setDeptPhase(key, v) { S.deptPhase = Number(v) || 0; render(); },
    setCompare(key, v) { S.cmp[key] = Number(v) || 0; render(); },
    // 버전 저장: 이름 · 메모를 받아 지금 자료를 버전으로 남기고 파일에 저장 (저장을 못 하면 버전도 되돌림)
    async saveVersion() {
      const p = S.proj;
      const names = (p.versions || []).map(v => v.name);
      const f = await UI.form("버전 저장", "지금 자료를 이 이름으로 남기고 저장합니다. 이미 있는 버전 이름을 고르면 그 버전을 지금 자료로 덮어씁니다.",
        [{ key: "name", label: "버전 이름", value: M.nextVersionName(p), options: [...new Set(names.concat(M.VERSION_NAMES))] }, { key: "memo", label: "메모", ph: "예: 고객 요청으로 2차 범위 축소, 외주 1명 감축 (덮어쓸 때 비우면 원래 메모)", multiline: true }], "버전 저장");
      if (!f) return;
      if (!f.name) return UI.toast("버전 이름을 넣으세요.", "error");
      const i = (p.versions || []).findIndex(v => v.name === f.name);
      if (i >= 0) return A.overwriteVersion(i, f.memo);
      pushUndo();
      const no = (p.versions || []).reduce((m, v) => Math.max(m, v.no || 0), 0) + 1;
      p.versions = (p.versions || []).concat([{ no, name: f.name, memo: f.memo, saved_at: PC.nowText(), data: M.snapshot(p) }]);
      S.cmp = { a: no, b: 0 };
      if (await save(false)) UI.toast(`버전 ${no}. ${f.name} 을(를) 남기고 저장했습니다.`, "ok");
      else { p.versions = p.versions.filter(v => v.no !== no); S.past.pop(); render(); }
    },
    // 버전 덮어쓰기: 번호 · 이름은 그대로, 자료 · 저장 일시(· 메모)를 지금 것으로 바꾸고 파일에 저장
    async overwriteVersion(i, memo) {
      const p = S.proj, v = p.versions[i];
      if (!(await UI.confirm("버전 덮어쓰기", `${v.no}. ${v.name} (${v.saved_at}) 을(를) 지금 자료로 덮어쓰고 저장합니다.\n예전 자료는 backups 폴더의 이전 파일에 남습니다.`, "덮어쓰기", true))) return;
      pushUndo();
      const before = PC.clone(v);
      Object.assign(v, { saved_at: PC.nowText(), data: M.snapshot(p) }, memo ? { memo } : {});
      S.cmp = { a: v.no, b: 0 };
      if (await save(false)) UI.toast(`버전 ${v.no}. ${v.name} 을(를) 지금 자료로 덮어썼습니다.`, "ok");
      else { p.versions[i] = before; S.past.pop(); render(); }
    },
    async restoreVersion(i) {
      const v = S.proj.versions[i];
      if (!(await UI.confirm("버전 불러오기", `${v.no}. ${v.name} (${v.saved_at}) 의 자료로 지금 화면을 바꿉니다.\n버전 목록은 그대로이고, [저장]을 눌러야 파일에 반영됩니다. Ctrl+Z 로 되돌릴 수 있습니다.`, "불러오기"))) return;
      pushUndo();
      S.proj = M.restoreVersion(S.proj, v, S.master);
      render();
      UI.toast(`${v.no}. ${v.name} 을(를) 불러왔습니다 — 저장해야 파일에 반영됩니다.`);
    },
    async removeVersion(i) {
      const v = S.proj.versions[i];
      if (!(await UI.confirm("버전 지우기", `${v.no}. ${v.name} 을(를) 지울까요? [저장]을 눌러야 파일에서도 빠집니다.`, "지우기", true))) return;
      pushUndo();
      S.proj.versions.splice(i, 1);
      render();
    },
    addRow(list) {
      if (list === "phases" && S.proj.phases.length >= 5) return UI.toast("분할계약은 5회차까지 넣을 수 있습니다.", "error");
      pushUndo();
      listAt(list).push(ROW[list](S.proj));
      render();
    },
    removeRow(list, i) { pushUndo(); listAt(list).splice(i, 1); render(); },
    undo, redo,
    // 총 MM 을 기간에 고르게 (끝자리는 마지막 달에)
    distribute(list, i) {   // 총 MM 을 전체 기간에 고르게
      const row = S.proj[list][i], months = S.calc.months, total = num(row.target_mm);
      if (!months.length) return;
      pushUndo();
      row.mm = spread(total, months);
      syncTotal(row);
      render();
    },
    // 기간을 골라 그 기간에만 나눔 (기간 밖의 달: 0 으로 / 그대로)
    async distributeRange(list, i) {
      const row = S.proj[list][i], months = S.calc.months;
      if (!months.length) return;
      const ch = months.map(m => ({ v: m, t: m }));
      const f = await UI.form("기간 배분", `${row.name || ""} — 고른 기간의 달에 MM 을 고르게 나눕니다.`, [
        { key: "from", label: "시작월", choices: ch, value: months[0] }, { key: "to", label: "종료월", choices: ch, value: months[months.length - 1] },
        { key: "mm", label: "이 기간에 넣을 MM", value: String(num(row.target_mm) || "") },
        { key: "rest", label: "기간 밖의 달", choices: [{ v: "zero", t: "0 으로 비움" }, { v: "keep", t: "지금 값 그대로" }], value: "zero" }], "배분");
      if (!f) return;
      const sel = months.filter(m => m >= (f.from < f.to ? f.from : f.to) && m <= (f.from < f.to ? f.to : f.from)), total = PC.pyRound(num(f.mm), 2);
      if (!sel.length || !(total > 0)) return UI.toast("기간과 MM 을 확인하세요.", "error");
      pushUndo();
      row.mm = Object.assign(f.rest === "keep" ? Object.assign({}, row.mm) : {}, spread(total, sel));
      syncTotal(row);
      render();
      UI.toast(`${sel[0]} ~ ${sel[sel.length - 1]} (${sel.length}개월)에 ${total} MM 을 나눴습니다 — 총 MM ${row.target_mm}`, "ok");
    },
    sortRows(list) {   // 기준정보 순서로 줄 세우기
      pushUndo();
      S.proj[list] = Calc.orderRows(list, S.proj[list], S.proj.settings);
      render();
      UI.toast("기준정보에 등록한 순서로 줄을 세웠습니다. Ctrl+Z 로 되돌릴 수 있습니다.");
    },
    insertRow(list, i) {   // 이 줄 아래에 새 줄 (위아래로 여러 번 옮기지 않아도 되게)
      const l = listAt(list), row = ROW[list](S.proj);
      pushUndo();
      l.splice(i + 1, 0, row);
      render(typeof row === "object" && row ? `${list}.${i + 1}.${Object.keys(row)[0]}` : `${list}.${i + 1}`);
    },
    moveRow(list, i, dir) {   // 기준정보 목록의 순서 바꾸기
      const l = listAt(list), j = i + dir;
      if (j < 0 || j >= l.length) return;
      pushUndo();
      [l[i], l[j]] = [l[j], l[i]];
      render();
    },
    applyEstimate() {
      const sup = S.calc.estimate.supply_total;
      if (sup <= 0) return UI.toast("견적 공급가액이 0원입니다. 품목을 먼저 넣으세요.", "error");
      pushUndo();
      S.proj.contract_revenue = sup;
      S.proj.estimate.applied_to_revenue = true;
      render();
      UI.toast(`견적 공급가액 ${won(sup)}원을 계약 매출로 넣었습니다.`, "ok");
    },
    // 샘플 프로젝트 추가: 예시 자료로 새 프로젝트(저장 전)를 엶 — 지금 프로젝트 파일은 그대로
    async addSample() {
      if (!(await UI.confirm("샘플 프로젝트 추가", "예시 자료가 든 새 프로젝트를 엽니다. 지금 프로젝트 파일은 그대로이고, 샘플은 [저장]해야 파일로 남습니다.", "추가"))) return;
      if (!(await guardDirty())) return;
      setProject(M.sampleProject(S.master), null, null);
      UI.toast("샘플 프로젝트를 열었습니다 — 저장 전입니다.");
    },
    // 마스터를 이 프로젝트에 반영: 무엇이 바뀌는지 보여 주고 (항목 더하기 · 같은 이름은 마스터 값으로)
    async applyMaster() {
      const changes = M.masterChanges(S.proj.settings, S.master).filter(x => !x.startsWith("(마스터에 없음)"));
      if (!changes.length) return UI.toast("이 프로젝트의 기준정보가 이미 마스터와 같습니다.");
      if (!(await UI.confirm("마스터 반영", `${changes.length}가지가 바뀝니다 (이미 있는 항목의 금액 · 프로젝트에서 더한 항목 · 외주업체 단가 · 시스템은 그대로):\n\n${changes.slice(0, 15).join("\n")}${changes.length > 15 ? `\n… 외 ${changes.length - 15}가지` : ""}`, "반영"))) return;
      pushUndo();
      S.proj.settings = M.mergeMaster(S.proj.settings, S.master);
      render();
      UI.toast("마스터를 반영했습니다. Ctrl+Z 로 되돌릴 수 있습니다.", "ok");
    },
    // 이 프로젝트의 기준정보(프로젝트 전용 제외)를 폴더의 마스터로 — 이름 · 날짜를 붙여 (이전 마스터는 backups 에)
    async saveDefaults() {
      if (!(await needFolder())) return;
      const today = new Date().toISOString().slice(0, 10);
      const f = await UI.form("마스터로 저장", "이 프로젝트의 마스터 목록(직급 · 기술등급 · 제안 직무/등급 · 영역 · 경비 계정 · 부서 · 결재자)을 폴더의 마스터(master_rates.xlsx)로 저장합니다. 금액 · 외주업체 · 시스템은 넣지 않습니다. 이전 마스터는 backups 폴더에 남습니다.",
        [{ key: "name", label: "마스터 이름 (배포 구분용)", value: ((S.master || {}).master_version || {}).name || `${today.slice(0, 7)} 전사 기준정보` }], "저장");
      if (!f) return;
      const st = M.masterOf(S.proj.settings);
      st.master_version = { name: f.name, date: today };
      await Store.saveMaster(st);
      S.master = clone(st);
      pushUndo();
      S.proj.settings.master_version = clone(st.master_version);
      render();
      UI.toast(`마스터로 저장했습니다 — ${f.name} (master/master_rates.xlsx)`, "ok");
    },
    // 배포용 마스터 파일 내려받기
    exportMaster() {
      const m = S.master || M.masterOf(S.proj.settings);
      UI.download(X.build(W.masterSheets(M.masterOf(m))), `master_rates_${(m.master_version && m.master_version.date) || new Date().toISOString().slice(0, 10)}.xlsx`);
    },
    // 본사에서 받은 마스터 파일로 폴더의 마스터를 바꿈 (바뀌는 것을 먼저 보여 줌)
    async uploadMaster(fileObj) {
      if (!(await needFolder())) return;
      let next;
      try { next = W.settingsFrom(await X.read(fileObj), M.defaultSettings()); } catch (e) { return UI.toast(`마스터 파일을 읽지 못했습니다: ${e.message}`, "error"); }
      next = M.masterOf(next);
      const changes = M.masterChanges(M.masterOf(S.master || {}), next);
      const v = next.master_version || {};
      if (!(await UI.confirm("마스터 파일 올리기", `${fileObj.name}${v.name ? ` — ${v.name}${v.date ? ` (${v.date})` : ""}` : ""}\n폴더의 마스터를 이 파일로 바꿉니다 (지금 마스터는 backups 에 남음).\n\n${changes.length ? changes.slice(0, 15).join("\n") + (changes.length > 15 ? `\n… 외 ${changes.length - 15}가지` : "") : "바뀌는 것이 없습니다."}\n\n이미 만든 프로젝트는 그대로이고, 필요하면 [마스터를 이 프로젝트에 반영]을 누르세요.`, "바꾸기"))) return;
      await Store.saveMaster(next);
      S.master = clone(next);
      render();
      UI.toast("폴더의 마스터를 바꿨습니다.", "ok");
    },
    // 인력 계획의 제안 직무 · 등급별 MM 으로 견적 인건비 줄 만들기 (견적 MM 은 뒤에 고쳐도 됨)
    async estimateFromPlan() {
      const r = Calc.estimateFromPlan(S.calc);
      if (!r.labor.length) return UI.toast("제안 직무 · 등급을 넣은 인력이 없습니다. 내부 · 외주 인력 표의 [제안 직무] · [제안 등급]을 먼저 고르세요.", "error");
      if ((S.proj.estimate.labor || []).length && !(await UI.confirm("인력 계획에서 가져오기", "지금 인건비 줄을 인력 계획으로 바꿉니다.", "바꾸기"))) return;
      pushUndo();
      S.proj.estimate.labor = r.labor;
      render();
      UI.toast(`${r.labor.length}줄을 만들었습니다${r.missing.length ? ` — 제안 직무 · 등급이 없는 인력 ${r.missing.length}명(${r.missing.slice(0, 3).join(", ")}${r.missing.length > 3 ? " …" : ""})은 빠짐` : ""}.`, r.missing.length ? "error" : "ok");
    },
    toggleTheme() {
      const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
      document.documentElement.dataset.theme = next;
      renderTheme();
      try { localStorage.setItem("prjcost.theme", next); } catch (e) { /* 저장 못 해도 이번 화면은 바뀜 */ }
    },
    async connectFolder() {
      if (!(await guardDirty())) return;
      try { await Store.connect(); } catch (e) { if (e.name !== "AbortError") UI.toast(`폴더를 열지 못했습니다: ${e.message}`, "error"); return; }
      await afterConnect(true);
    },
    async reconnect() {
      const r = await Store.restore(true);
      if (r.ok) await afterConnect(false); else UI.toast("폴더 사용을 허락하지 않았습니다.", "error");
    },
    async newProject() {
      if (!(await guardDirty())) return;
      setProject(M.newProject(S.master), null, null);
    },
    async openSelected() {
      const sel = $("projectSelect"), file = sel.value;
      if (!file || file === S.file) return;
      if (!(await guardDirty())) { sel.value = S.file || ""; return; }
      await openFile(file);
    },
    async deleteProject() {
      if (!S.file) return;
      if (!(await UI.confirm("프로젝트 지우기", `${S.file} 을(를) 지웁니다.\n파일은 backups/deleted 폴더로 옮겨지므로 되살릴 수 있습니다.`, "지우기", true))) return;
      try {
        const where = await Store.deleteProject(S.file);
        UI.toast(`지웠습니다 (보관: ${where})`, "ok");
        setProject(M.newProject(S.master), null, null);
        await refreshList();
      } catch (e) { UI.toast(e.message, "error"); }
    },
    save: () => save(false),
    // 사업검토서 출력: templates 폴더의 회사 양식에 값을 채워 새 파일로 (양식 파일은 그대로)
    async exportTemplate() {
      let names = [];
      try { names = Store.dir ? await Store.listTemplates() : Object.keys(PC.BUILTIN_TEMPLATES || {}); } catch (e) { console.warn("[App] templates", e); }
      if (!names.length) return UI.toast("templates 폴더에 보고서 양식(xlsx)이 없습니다.", "error");
      const f = await UI.form("사업검토서 출력", `고른 양식의 {{자리표시}} 칸에 이 프로젝트의 값을 넣어 새 파일로 만듭니다. 양식은 ${Store.dir ? `${Store.label}/templates` : "기본 양식"} 에 있습니다.`,
        [{ key: "name", label: "보고서 양식", choices: names.map(n => ({ v: n, t: n.replace(/\.xlsx$/i, "") })), value: S.lastTemplate && names.includes(S.lastTemplate) ? S.lastTemplate : names[0] }], "출력");
      if (!f) return;
      S.lastTemplate = f.name;
      try {
        let src = Store.dir ? await Store.templateFile(f.name) : null;
        if (!src) { const b = atob((PC.BUILTIN_TEMPLATES || {})[f.name] || ""); src = Uint8Array.from(b, ch => ch.charCodeAt(0)); }
        const lastSaved = S.file && S.mtime && !isDirty() ? new Date(S.mtime) : new Date();
        const r = await PC.Template.fill(src, PC.Template.values(S.proj, Calc.calculateProjectCost(S.proj, S.master), { lastSaved }));
        const name = PC.Template.outName(f.name, S.proj, lastSaved);
        UI.download(r.blob, name);
        if (Store.dir) { try { await Store.saveReport(r.blob, name); } catch (e) { console.warn("[App] report copy", e); } }
        const notes = [], ck = r.check;
        if (r.unknown.length) notes.push(`모르는 자리표시 ${r.unknown.map(k => `{{${k}}}`).join(", ")} — 빈칸으로 둠`);
        if (r.empty.length) notes.push(`값이 없는 칸: ${r.empty.join(", ")}`);
        if (ck.items.length && !ck.failed.length) notes.push(`합계 수식 검증 ${ck.items.length}개 항목 일치`);
        if (ck.problems.length) notes.push(`계산해 보지 못한 수식 ${ck.problems.length}개 (${ck.problems.slice(0, 3).map(x => `${x.sheet}!${x.ref} ${x.reason}`).join(", ")}) — Excel 에서는 계산됨`);
        UI.toast(`${name} 을(를) 만들었습니다${Store.dir ? " (reports 폴더에도 저장)" : ""}.${notes.length ? " " + notes.join(" · ") : ""}`, r.unknown.length ? "error" : "ok");
        // 양식 수식 값이 화면 값과 다르면 (양식의 합계 범위 · 참조가 틀린 경우) 알려 줌
        if (ck.failed.length) {
          const show = v => (typeof v === "number" ? (Math.abs(v) >= 1000 ? PC.pyRound(v, 2).toLocaleString("ko-KR") : String(PC.pyRound(v, 4))) : v === null || v === "" ? "빈칸" : String(v));
          await UI.choose("사업검토서 수식 검증", `파일은 만들었지만, 양식의 수식으로 계산한 값이 프로그램 값과 다른 항목이 있습니다. 양식(templates 폴더)의 합계 수식 범위 · 참조를 확인해 주세요.\n\n${ck.failed.slice(0, 10).map(x => `· ${x.name}: 양식 ${show(x.sheet)} / 프로그램 ${show(x.program)}`).join("\n")}${ck.failed.length > 10 ? `\n… 외 ${ck.failed.length - 10}개` : ""}`, [{ id: "ok", label: "확인", cls: "primary" }]);
        }
      } catch (e) { console.error(e); UI.toast(`양식을 채우지 못했습니다: ${e.message}`, "error"); }
    },
    // 보고서: 통합보고 · 항목별 상세 2시트 엑셀 — 내려받고, 폴더가 연결돼 있으면 reports 폴더에도 남김
    async exportReport() {
      const calc = Calc.calculateProjectCost(S.proj, S.master);
      const blob = X.build(PC.Report.sheets(S.proj, calc));
      const name = `${safeId(S.proj.project_id)}_원가보고서_${PC.stamp()}.xlsx`;
      UI.download(blob, name);
      if (Store.dir) { try { await Store.saveReport(blob, name); } catch (e) { console.warn("[App] report copy", e); } }
      UI.toast(`보고서를 내려받았습니다 — ${name}${Store.dir ? " (reports 폴더에도 저장)" : ""}`, "ok");
    }
  };

  // ---------------- 열기 · 저장 ----------------
  function setProject(p, file, mtime) {
    S.proj = p; S.file = file; S.mtime = mtime;
    S.past = []; S.future = [];
    markSaved();
    render();
    renderList();
  }
  async function openFile(file) {
    try {
      const r = await Store.loadProject(file, S.master);
      setProject(r.project, file, r.mtime);
      const note = S.noted.has(file) ? "" : M.legacyNote(r.raw);
      S.noted.add(file);
      if (note) UI.choose("계산 방식이 바뀐 자료가 있습니다", note, [{ id: "ok", label: "확인", cls: "primary" }]);
      else UI.toast(`${r.project.project_id} 을(를) 열었습니다.`);
    } catch (e) { UI.toast(e.message, "error"); await refreshList(); }
  }
  async function refreshList() {
    S.list = Store.dir ? await Store.listProjects() : [];
    renderList();
  }
  // 바뀐 내용이 있으면 먼저 묻기: 저장 / 버리기 / 취소 → 계속해도 되면 true
  async function guardDirty() {
    if (!isDirty()) return true;
    const c = await UI.choose("저장하지 않은 변경", `${S.proj.project_id} 에 저장하지 않은 변경이 있습니다.`,
      [{ id: "cancel", label: "취소" }, { id: "discard", label: "버리기", cls: "danger" }, { id: "save", label: "저장하고 계속", cls: "primary" }]);
    if (c === "save") return save(false);
    return c === "discard";
  }
  async function needFolder() {
    if (Store.dir) return true;
    if (!Store.supported) { UI.toast("이 브라우저는 폴더 저장을 지원하지 않습니다. 크롬이나 엣지로 여세요.", "error"); return false; }
    const ok = await UI.confirm("폴더 연결", "프로젝트 엑셀을 저장할 폴더를 먼저 고르세요.\n예전에 쓰던 data 폴더를 고르면 그 프로젝트들을 그대로 씁니다.", "폴더 고르기");
    if (!ok) return false;
    try { await Store.connect(); } catch (e) { return false; }
    await afterConnect(false);
    return true;
  }
  async function save(force) {
    if (!Store.dir && !Store.supported) {   // 폴더를 못 쓰는 브라우저: 파일로 내려받기
      UI.download(X.build(W.projectSheets(S.proj, Calc.calculateProjectCost(S.proj, S.master), [])), `${safeId(S.proj.project_id)}.xlsx`);
      markSaved(); renderSaveState();
      return true;
    }
    if (!(await needFolder())) return false;
    const file = S.file || Store.fileFor(S.proj);
    if (!S.file && await Store.exists(file)) {
      const ok = await UI.confirm("같은 ID 의 파일이 있습니다", `projects/${file} 이(가) 이미 있습니다. 덮어쓸까요?\n지금 파일은 backups 폴더에 남습니다. 다른 프로젝트라면 [취소] 후 프로젝트 ID 를 바꾸세요.`, "덮어쓰기", true);
      if (!ok) return false;
    }
    // 더 쓰지 않는 예전 비율 방식 인비례 값은 저장할 때 지움 (열 때 이미 알림)
    ["internal_overhead_ratio", "external_overhead_ratio"].forEach(k => { delete S.proj.settings[k]; });
    try {
      S.mtime = await Store.saveProject(S.proj, Calc.calculateProjectCost(S.proj, S.master), file, force ? null : S.mtime);
    } catch (e) {
      if (e.code !== "CHANGED") { UI.toast(`저장하지 못했습니다: ${e.message}`, "error"); return false; }
      const c = await UI.choose("파일이 바뀌었습니다", `${e.message}\n\n[덮어쓰기]를 누르면 바뀐 파일은 backups 폴더에 남기고 지금 내용으로 저장합니다.`,
        [{ id: "cancel", label: "취소" }, { id: "reload", label: "다시 불러오기" }, { id: "force", label: "덮어쓰기", cls: "danger" }]);
      if (c === "reload") { await openFile(file); return false; }
      return c === "force" ? save(true) : false;
    }
    S.file = file;
    markSaved();
    await refreshList();
    render();
    UI.toast(`저장했습니다 — projects/${file} (이전 파일은 backups 에)`, "ok");
    return true;
  }
  async function afterConnect(announce) {
    S.master = await Store.loadMaster();
    const fresh = !S.file && !isDirty();
    if (fresh) setProject(M.newProject(S.master), null, null);
    await refreshList();
    renderFolder();
    notice("");
    if (announce) UI.toast(`${Store.label} 폴더를 연결했습니다 — 프로젝트 ${S.list.length}개`, "ok");
  }
  async function upload(fileObj) {
    if (!fileObj) return;
    if (!(await guardDirty())) return;
    try {
      const p = await Store.readUpload(fileObj, S.master);
      setProject(p, null, null);
      S.saved = "";   // 아직 폴더에 없음 → 저장 안 됨
      renderSaveState();
      UI.toast(`${fileObj.name} 을(를) 열었습니다. [저장]하면 projects 폴더에 새 파일로 기록합니다.`);
    } catch (e) { UI.toast(`열지 못했습니다: ${e.message}`, "error"); }
  }

  // ---------------- 시작 ----------------
  async function start() {
    $("verBadge").textContent = VERSION;
    PC.fillIcons();
    S.proj = M.newProject(S.master);
    markSaved();
    document.addEventListener("change", e => {
      if (e.target.id === "uploadInput") { upload(e.target.files[0]); e.target.value = ""; return; }
      if (e.target.dataset.act === "openSelected") return A.openSelected();
      if (e.target.dataset.changeAct) return A[e.target.dataset.changeAct](e.target.dataset.key, e.target.value);
      if (e.target.id === "masterInput") { const f = e.target.files && e.target.files[0]; e.target.value = ""; if (f) A.uploadMaster(f); return; }
      onInput(e);
    });
    document.addEventListener("click", e => {
      const b = e.target.closest("button[data-act]");
      if (!b || b.disabled) return;
      const fn = A[b.dataset.act];
      if (fn) fn.apply(null, b.dataset.args ? JSON.parse(b.dataset.args) : []);
    });
    // Enter = 칸 확정 (다음 계산), Ctrl+S = 저장
    document.addEventListener("keydown", e => {
      if (e.isComposing || e.keyCode === 229) return;   // 한글 조합 중 키는 입력기 몫
      const k = e.key.toLowerCase(), mod = e.ctrlKey || e.metaKey;
      // 되돌리기: 칸에 치던 중(아직 확정 전)이면 브라우저의 글자 되돌리기, 아니면 자료 되돌리기
      if (mod && (k === "z" || k === "y")) {
        const a = document.activeElement, typing = a && a.matches && a.matches("input[data-bind], textarea[data-bind], input[data-new]") && a.value !== a.defaultValue;
        if (!typing) { e.preventDefault(); if (k === "y" || e.shiftKey) redo(); else undo(); }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") { e.preventDefault(); const a = document.activeElement; if (a && a.dataset && a.dataset.bind) a.blur(); save(false); }
      else if (e.key === "Enter" && e.target.matches("input[data-bind]")) e.target.dispatchEvent(new Event("change", { bubbles: true }));
    });
    PC.Grid.init({
      proj: () => S.proj, pushUndo, render,
      listAt,
      newRow: list => { const l = listAt(list); l.push(ROW[list](S.proj)); return l.length - 1; },
      setValue: (path, v) => { setPath(S.proj, path, v); afterBind(path, v); },
      // 표의 한 칸에 값: 경비의 입력 금액(@amount)은 분배 방식에 따라 총액 · 한 달 금액으로, 달별 금액을 넣으면 '월별 직접 입력'으로
      setCell: (list, row, b, v) => {
        if (!b) {   // 글자 목록 (외주 등급 · 담당 영역 · 부서)
          const old = listAt(list)[row];
          setPath(S.proj, `${list}.${row}`, v);
          renameRefs(`${list}.${row}`, old, v);
          afterBind(`${list}.${row}`, v);
          return;
        }
        const r = listAt(list)[row];
        if (b === "@amount") { if (r.dist_mode === "manual") return; b = r.dist_mode === "fill_all" ? "unit_input" : "total_input"; }
        if (/^monthly\./.test(b) && r.dist_mode !== "manual") {
          const calc = Calc.calculateProjectCost(S.proj, S.master).details[list][row].calc_monthly;
          r.monthly = Object.assign({}, calc);
          r.dist_mode = "manual";
        }
        setPath(S.proj, `${list}.${row}.${b}`, v);
        afterBind(`${list}.${row}.${b}`, v);
      },
      known: (list, key, v) => !v || (key === "site" ? (S.proj.settings.work_sites || []).includes(v) : key === "tech_grade" ? (S.proj.settings.external_grades || []).includes(v) : key === "dept" ? (S.proj.settings.departments || []).includes(v) : key === "area" ? (S.proj.settings.areas || []).includes(v) : key === "grade"
        ? (list === "internal_labor" ? (S.proj.settings.internal_grades || []).some(g => g.grade === v) : (S.proj.settings.external_grades || []).includes(v))
        : (S.proj.settings.vendor_rates || []).some(r => r.vendor === v))
    });
    const oops = msg => { console.error(msg); if (!oops.shown) { oops.shown = true; UI.toast(`프로그램 오류가 났습니다: ${String(msg).slice(0, 120)} — 저장 후 새로 고침(F5)을 권합니다.`, "error"); setTimeout(() => { oops.shown = false; }, 5000); } };
    window.addEventListener("error", e => oops(e.message || e.error));
    window.addEventListener("unhandledrejection", e => oops((e.reason && e.reason.message) || e.reason));
    window.addEventListener("beforeunload", e => { if (isDirty()) { e.preventDefault(); e.returnValue = ""; } });
    render();
    renderFolder();
    renderList();
    renderTheme();
    if (!Store.supported) {
      notice(`<b>이 브라우저는 폴더 저장을 지원하지 않습니다.</b> 크롬 또는 엣지로 여세요. 지금은 [저장]을 누르면 엑셀 파일을 내려받습니다.`);
      return;
    }
    const r = await Store.restore(false);
    if (r.ok) return afterConnect(false);
    notice(r.needsClick
      ? `지난번 폴더 <b>${esc(r.name)}</b> 에 다시 연결하려면 눌러 주세요. <button class="btn sm primary" data-act="reconnect">📁 폴더 다시 연결</button>`
      : `프로젝트 엑셀을 저장할 <b>폴더를 연결</b>하세요. 예전 버전에서 쓰던 <b>data</b> 폴더를 고르면 그 프로젝트들을 그대로 씁니다. <button class="btn sm primary" data-act="connectFolder">📁 폴더 연결</button>`);
  }

  PC.App = { S, A, render, save, openFile, afterConnect, setProject, undo, redo };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start); else start();
})(window);
