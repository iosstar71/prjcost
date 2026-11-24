// =============================================================================
// 프로그램 흐름 — 지금 프로젝트 · 입력 반영 · 저장 / 열기 / 지우기 · 폴더 연결
//  자료가 바뀔 때마다 그 자리에서 다시 계산하고 다시 그림 (서버 없음). 저장은 [저장] 또는 Ctrl+S 를 눌러야 기록
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC;
  const { Calc, Model: M, Store, UI, Views, Xlsx: X, Workbook: W, clone, won, esc, num, safeId } = PC;
  const APP_NAME = "PM Hub Cost", VERSION = "v0.9";

  const S = {
    master: M.defaultSettings(),   // 폴더의 기준정보 (새 프로젝트 기본값)
    proj: null,
    file: null,                    // projects 폴더 안의 파일 이름 (저장 전이면 null)
    mtime: null,                   // 불러온 때의 파일 수정 시각 (다른 곳에서 바뀌었는지 확인)
    saved: "",                     // 마지막으로 저장 · 불러온 자료 (바뀌었는지 비교)
    calc: null,
    tab: "info",
    list: []
  };
  const $ = id => document.getElementById(id);
  const isDirty = () => JSON.stringify(S.proj) !== S.saved;
  const markSaved = () => { S.saved = JSON.stringify(S.proj); };

  // ---------------- 그리기 ----------------
  function render() {
    S.calc = Calc.calculateProjectCost(S.proj, S.master);
    const focus = captureFocus();
    renderKpis();
    renderTabs();
    const tab = Views.TABS.find(t => t.id === S.tab) || Views.TABS[0];
    $("panel").innerHTML = tab.view(S.proj, S.calc, { file: S.file });
    renderSaveState();
    restoreFocus(focus);
  }
  function renderKpis() {
    const s = S.calc.summary, b = s.breakdown, est = S.calc.estimate;
    $("kRev").textContent = `${won(s.revenue)}원`;
    $("kRevSub").textContent = est.supply_total > 0 ? `견적 공급가 ${won(est.supply_total)}원${est.applied_to_revenue ? " · 매출 적용" : ""}` : "견적 미등록";
    $("kCost").textContent = `${won(s.total_cost)}원`;
    const labor = b.internal_labor + b.external_labor, exp = b.general_expense + b.internal_overhead + b.external_overhead + b.custom_cost;
    const m = v => (v / 1e6).toFixed(1);
    $("kCostSub").textContent = `내부 ${m(b.internal_labor)} · 외주 ${m(b.external_labor)} · 경비 ${m(exp)} (백만 원)`;
    const mix = $("kMix");
    mix.innerHTML = `<i class="m1"></i><i class="m2"></i><i class="m3"></i>`;
    const tot = s.total_cost || 1;
    [b.internal_labor, b.external_labor, exp].forEach((v, i) => { mix.children[i].style.width = `${Math.max(0, (v / tot) * 100)}%`; });
    mix.title = `인건비 ${won(labor)}원 · 경비 ${won(exp)}원`;
    $("kMargin").textContent = `${s.margin_rate.toFixed(1)}%`;
    $("kProfit").textContent = `이익 ${won(s.profit)}원`;
    $("kMarginCard").className = `kpi ${s.revenue > 0 ? (s.hurdle_met ? "ok" : "bad") : ""}`;
    $("kHurdle").textContent = `${s.hurdle_gap >= 0 ? "+" : ""}${s.hurdle_gap.toFixed(1)}%p`;
    $("kHurdleSub").textContent = `목표 ${s.hurdle_rate}% · ${s.hurdle_met ? "달성" : "미달"}`;
    $("kHurdleCard").className = `kpi ${s.revenue > 0 ? (s.hurdle_met ? "ok" : "bad") : ""}`;
  }
  function renderTabs() {
    $("tabs").innerHTML = Views.TABS.map(t => {
      const n = t.count ? t.count(S.proj) : 0;
      return `<button class="tab${t.id === S.tab ? " active" : ""}" role="tab" aria-selected="${t.id === S.tab}" data-act="showTab" data-args="${esc(JSON.stringify([t.id]))}">${esc(t.label)}${n ? `<span class="cnt">${n}</span>` : ""}</button>`;
    }).join("");
  }
  function renderSaveState() {
    const el = $("saveState"), dirty = isDirty();
    el.textContent = dirty ? "● 저장 안 됨" : S.file ? "저장됨" : "";
    el.className = `save-state${dirty ? " dirty" : ""}`;
    document.title = `${dirty ? "● " : ""}${S.proj.project_name || S.proj.project_id} — ${APP_NAME}`;
  }
  function renderFolder() {
    const chip = $("folderChip"), btn = $("btnFolder");
    chip.classList.toggle("off", !Store.dir);
    chip.querySelector(".name").textContent = Store.dir ? Store.label : "폴더 연결 전";
    btn.textContent = Store.dir ? "📁 폴더 바꾸기" : "📁 폴더 연결";
    btn.hidden = !Store.supported;
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
  const renderTheme = () => { $("btnTheme").textContent = document.documentElement.dataset.theme === "dark" ? "밝게" : "어둡게"; };
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
    if (t === "blank-num") return v === "" ? "" : num(v);
    if (t === "pct") return num(v) / 100;
    if (t === "month2") return String(v || "03").replace(/\D/g, "").padStart(2, "0").slice(-2);
    return el.tagName === "TEXTAREA" ? v : v.trim();
  }
  // 값을 넣은 뒤 함께 바꿀 것 (원본과 같은 동작)
  function afterBind(path, value) {
    const p = S.proj;
    let m;
    if ((m = /^internal_labor\.(\d+)\.grade$/.exec(path))) {   // 등급을 바꾸면 기준단가도 그 등급 단가로
      const g = (p.settings.internal_grades || []).find(x => x.grade === value);
      if (g) p.internal_labor[+m[1]].base_rate = num(g.rate);
    } else if ((m = /^internal_labor\.(\d+)\.promotions\.0\./.exec(path))) {   // 승진월 · 단가가 모두 비면 승진 없음
      const mem = p.internal_labor[+m[1]], p0 = mem.promotions[0] || {};
      if (!p0.effective_ym && (p0.rate === "" || p0.rate === undefined)) mem.promotions = [];
    } else if (path === "date_mode") {   // 월 ↔ 일자: 값 모양만 바꿈
      const toDate = (v, end) => (String(v || "").length === 7 ? `${v}-${end ? "28" : "01"}` : v);
      p.start_date = value === "date" ? toDate(p.start_date, false) : String(p.start_date || "").slice(0, 7);
      p.end_date = value === "date" ? toDate(p.end_date, true) : String(p.end_date || "").slice(0, 7);
    }
  }
  function onInput(e) {
    const el = e.target.closest("[data-bind]");
    if (!el) return;
    const v = valueOf(el);
    setPath(S.proj, el.dataset.bind, v);
    afterBind(el.dataset.bind, v);
    // Tab · 클릭으로 다음 칸에 갈 때: 포커스가 옮겨진 뒤에 다시 그려야 그 칸을 잃지 않음
    clearTimeout(onInput.t);
    onInput.t = setTimeout(render, 0);
  }

  // ---------------- 동작 ----------------
  const ROW = {
    internal_labor: p => {
      const g0 = (p.settings.internal_grades || [])[0] || { grade: "선임", rate: 6000000 };
      return { name: "신규인력", role: "개발", grade: g0.grade, base_rate: num(g0.rate), overhead_eligible: true, target_mm: 0,
        promotions: [{ effective_ym: `${String(p.start_date || "2026").slice(0, 4)}-${p.settings.default_promotion_month || "03"}`, rate: "" }], mm: {} };
    },
    external_labor: p => {
      const v0 = (p.settings.vendor_rates || [])[0] || { vendor: "파트너A", grade: "고급" };
      return { name: "외주인력", role: "개발", vendor: v0.vendor, grade: v0.grade, custom_rate: "", target_mm: 0, mm: {} };
    },
    general_expenses: p => expenseRow(p, "general", "일반경비"),
    custom_costs: p => expenseRow(p, "custom", "기타비용"),
    phases: p => ({ title: `${p.phases.length + 1}차 계약`, start_ym: String(p.start_date || "").slice(0, 7), end_ym: String(p.end_date || "").slice(0, 7), revenue: 0, override_cost: "" }),
    "estimate.items": () => ({ item_name: "신규 품목", spec: "", qty: 1, unit_price: 0, memo: "" }),
    "settings.internal_grades": () => ({ grade: "신규등급", rate: 5000000 }),
    "settings.external_grades": () => "신규등급",
    "settings.vendor_rates": p => ({ vendor: "신규업체", grade: (p.settings.external_grades || [])[0] || "고급", rate: 7500000 }),
    "settings.expense_accounts": () => ({ account: "신규계정", category: "general", desc: "" })
  };
  function expenseRow(p, cat, fallback) {
    const acc = (p.settings.expense_accounts || []).find(a => a.category === cat);
    return { name: acc ? acc.account : fallback, dist_mode: "equal_total", total_input: 0, unit_input: 0, monthly: {}, memo: "" };
  }
  const listAt = path => path.split(".").reduce((o, k) => o[k], S.proj);

  const A = {
    showTab(id) { S.tab = id; render(); },
    addRow(list) {
      if (list === "phases" && S.proj.phases.length >= 5) return UI.toast("분할계약은 5회차까지 넣을 수 있습니다.", "error");
      listAt(list).push(ROW[list](S.proj));
      render();
    },
    removeRow(list, i) { listAt(list).splice(i, 1); render(); },
    // 총 MM 을 기간에 고르게 (끝자리는 마지막 달에)
    distribute(list, i) {
      const row = S.proj[list][i], months = S.calc.months, total = num(row.target_mm);
      if (!months.length) return;
      const per = Math.floor((total / months.length) * 100) / 100;
      const rest = Math.round((total - per * months.length) * 100) / 100;
      row.mm = Object.fromEntries(months.map((m, k) => [m, k === months.length - 1 ? Math.round((per + rest) * 100) / 100 : per]));
      render();
    },
    applyEstimate() {
      const sup = S.calc.estimate.supply_total;
      if (sup <= 0) return UI.toast("견적 공급가액이 0원입니다. 품목을 먼저 넣으세요.", "error");
      S.proj.contract_revenue = sup;
      S.proj.estimate.applied_to_revenue = true;
      render();
      UI.toast(`견적 공급가액 ${won(sup)}원을 계약 매출로 넣었습니다.`, "ok");
    },
    async loadDemo() {
      if (!(await UI.confirm("데모 채우기", "지금 화면의 내용을 예시 자료로 바꿉니다. 계속할까요?", "채우기"))) return;
      const keepFile = S.file;
      S.proj = M.demoProject(S.master);
      S.file = keepFile;
      render();
    },
    async resetSettings() {
      if (!(await UI.confirm("기준정보 바꾸기", "이 프로젝트의 기준정보를 폴더의 기본값(master_rates.xlsx)으로 바꿉니다.", "바꾸기"))) return;
      S.proj.settings = clone(S.master);
      render();
    },
    async saveDefaults() {
      if (!(await needFolder())) return;
      await Store.saveMaster(S.proj.settings);
      S.master = clone(S.proj.settings);
      UI.toast("새 프로젝트 기본값(master/master_rates.xlsx)으로 저장했습니다.", "ok");
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
    save: () => save(false)
  };

  // ---------------- 열기 · 저장 ----------------
  function setProject(p, file, mtime) {
    S.proj = p; S.file = file; S.mtime = mtime;
    markSaved();
    render();
    renderList();
  }
  async function openFile(file) {
    try {
      const r = await Store.loadProject(file, S.master);
      setProject(r.project, file, r.mtime);
      UI.toast(`${r.project.project_id} 을(를) 열었습니다.`);
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
    S.proj = M.newProject(S.master);
    markSaved();
    document.addEventListener("change", e => {
      if (e.target.id === "uploadInput") { upload(e.target.files[0]); e.target.value = ""; return; }
      if (e.target.dataset.act === "openSelected") return A.openSelected();
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
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") { e.preventDefault(); const a = document.activeElement; if (a && a.dataset && a.dataset.bind) a.blur(); save(false); }
      else if (e.key === "Enter" && e.target.matches("input[data-bind]")) e.target.dispatchEvent(new Event("change", { bubbles: true }));
    });
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

  PC.App = { S, A, render, save, openFile, afterConnect, setProject };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start); else start();
})(window);
