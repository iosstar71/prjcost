// =============================================================================
// 탭 화면 그리기 — 프로젝트 자료(p)와 계산 결과(c)로 HTML 을 만듦. 자료를 바꾸지 않음
//  입력칸: data-bind="자료 경로" data-type="text|num|mm|blank-num|bool|month2" (app.js 가 값을 넣고 다시 계산)
//  버튼:   data-act="동작 이름" data-args='JSON 배열'
//  엑셀처럼 쓰는 표(gridTable): 열 정의 한 곳에서 머리글 · 칸 · 붙여넣기(data-cols)를 함께 만듦 → js/grid.js
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC;
  const { esc, won, man, num, icon } = PC;
  const M = PC.Model;

  const args = (...a) => esc(JSON.stringify(a));
  const btn = (label, act, a, cls, title, ic) => `<button class="btn ${cls || ""}" data-act="${act}" data-args="${args(...(a || []))}"${title ? ` title="${esc(title)}"` : ""}>${ic ? icon(ic) : ""}${label}</button>`;
  const del = (list, i) => `<button class="btn ghost-danger icon-only" data-act="removeRow" data-args="${args(list, i)}" title="이 줄 지우기" aria-label="이 줄 지우기">${icon("trash")}</button>`;
  // 금액은 쉼표를 넣어 보여 줌 (숫자 칸은 쉼표를 못 보여서 글자 칸 + 숫자 키패드). 빈 값은 그대로 빈 칸
  const money = v => {
    if (v === "" || v === null || v === undefined) return "";
    const n = Number(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n.toLocaleString("ko-KR", { maximumFractionDigits: 2 }) : v;
  };
  function inp(path, value, o = {}) {
    const type = o.money ? "text" : (o.type || "text");
    const kind = o.kind || (o.money || type === "number" ? "num" : "text");
    const v = o.money ? money(value) : value === null || value === undefined ? "" : value;
    const cls = [o.money ? "money" : "", o.cls || ""].join(" ").trim();
    return `<input type="${type}"${o.money ? ' inputmode="decimal"' : ""} data-bind="${esc(path)}" data-type="${kind}" value="${esc(v)}"${cls ? ` class="${cls}"` : ""}${o.ph !== undefined ? ` placeholder="${esc(o.ph)}"` : ""}${o.step ? ` step="${o.step}"` : ""}${o.list ? ` list="${o.list}"` : ""}${o.label ? ` aria-label="${esc(o.label)}"` : ""}>`;
  }
  const mmInp = (path, value, label, cls) => inp(path, value, { type: "number", step: "0.01", kind: "mm", cls: cls || "mm", ph: "0", label });
  function sel(path, value, options, o = {}) {
    const opts = options.map(x => (typeof x === "object" ? x : { v: x, t: x }));
    const unknown = value !== undefined && value !== "" && !opts.some(x => String(x.v) === String(value));
    if (unknown) opts.unshift({ v: value, t: `${value} (목록에 없음)` });
    const cls = [o.cls || "", unknown ? "invalid" : ""].join(" ").trim();
    return `<select data-bind="${esc(path)}" data-type="${o.kind || "text"}"${cls ? ` class="${cls}"` : ""}${unknown ? ` title="기준정보에 없는 값입니다"` : ""}${o.label ? ` aria-label="${esc(o.label)}"` : ""}>${opts.map(x => `<option value="${esc(x.v)}"${String(x.v) === String(value) ? " selected" : ""}>${esc(x.t)}</option>`).join("")}</select>`;
  }
  const chk = (path, on, label) => `<input type="checkbox" data-bind="${esc(path)}" data-type="bool"${on ? " checked" : ""}${label ? ` aria-label="${esc(label)}"` : ""}>`;
  const field = (label, control, o = {}) => `<label class="field${o.full ? " full" : ""}"><span>${esc(label)}</span>${control}${o.help ? `<small class="help">${esc(o.help)}</small>` : ""}</label>`;
  const head = (ic, title, desc, right) => `<div class="panel-head"><div class="ph-title">${icon(ic)}<h2>${esc(title)}</h2>${desc ? `<span class="ph-desc">${esc(desc)}</span>` : ""}</div><div class="bar-group">${right || ""}</div></div>`;
  const subHead = (title, desc, right) => `<div class="sec-head"><h3>${esc(title)}</h3>${desc ? `<span class="ph-desc">${esc(desc)}</span>` : ""}<div class="bar-group">${right || ""}</div></div>`;
  const empty = (cols, text) => `<tr><td class="empty" colspan="${cols}">${esc(text)}</td></tr>`;
  const sumBy = (rows, f) => rows.reduce((a, r) => a + num(f(r)), 0);
  const GRID_HINT = "화살표 · Enter 이동 · 엑셀 붙여넣기 · Ctrl+R 오른쪽 채우기 · Ctrl+Z 되돌리기";

  // ---------- 엑셀처럼 쓰는 표 ----------
  //  cols: [{ b: 행 안의 자료 경로, t: 붙여넣을 때 값 읽는 법, l: 머리글 이름, th, cell(row, base, i), fill: 오른쪽 채우기 대상 }]
  //  칸 하나 = 열 하나 (첫 칸은 지우기) → grid.js 는 칸 위치로 열을 찾음
  function gridTable(o) {
    const meta = o.cols.map(x => ({ b: x.b, t: x.t, l: x.l, fill: !!x.fill }));
    const mv = i => (o.order ? `<span class="mv">${btn("", "moveRow", [o.list, i, -1], "icon-only ghost", "위로", "up")}${btn("", "moveRow", [o.list, i, 1], "icon-only ghost", "아래로", "down")}</span>` : "");
    const ins = i => btn("", "insertRow", [o.list, i], "icon-only ghost ins", "이 줄 아래에 새 줄 넣기", "plus");
    const body = o.rows.map((r, i) => `<tr data-row="${i}"><td class="c nowrap">${o.fixed ? "" : mv(i) + ins(i) + del(o.list, i)}</td>${o.cols.map(x => x.cell(r, `${o.list}.${i}`, i)).join("")}${o.tail ? o.tail(r) : ""}</tr>${o.after ? o.after(r, i, o.cols.length + 1 + (o.tail ? 1 : 0)) : ""}`).join("");
    const newRow = `<tr class="new-row"><td></td><td class="l"><input type="text" data-new="${o.list}"${o.newKey ? ` data-new-key="${o.newKey}"` : ""} placeholder="＋ ${esc(o.newLabel)}" aria-label="${esc(o.newLabel)}"></td>${o.cols.length - 1 + (o.tail ? 1 : 0) > 0 ? `<td class="l hint-cell" colspan="${o.cols.length - 1 + (o.tail ? 1 : 0)}">값을 넣으면 한 줄이 추가됩니다 · 엑셀에서 복사한 영역을 붙여 넣어도 됩니다</td>` : ""}</tr>`;
    return `<div class="tbl-wrap"><table class="grid" data-list="${o.list}" data-cols="${esc(JSON.stringify(meta))}"><thead><tr><th class="w-del"></th>${o.cols.map(x => x.th).join("")}${o.tailHead || ""}</tr></thead>
      <tbody>${body}${o.fixed ? "" : newRow}${o.foot || ""}</tbody></table></div>`;
  }
  const monthCols = (months, path, cellFn) => months.map((m, k) => ({ b: `${path}.${m}`, t: "mm", l: m, fill: true, th: `<th class="m">${esc(m.slice(2).replace("-", "."))}</th>`, cell: cellFn(m, k) }));

  const APPROVER_TITLES = ["기안(PM)", "검토", "사업팀장", "영업팀장", "본부장"];

  // ---------------- 프로젝트 정보 + 원가 구성 ----------------
  function info(p, c, ctx) {
    const dateType = p.date_mode === "date" ? "date" : "month";
    const toMode = v => (dateType === "date" ? (String(v || "").length === 7 ? `${v}-01` : v) : String(v || "").slice(0, 7));
    const fileText = ctx.file ? `projects/${ctx.file}` : "저장 전 — 처음 저장할 때 프로젝트 ID 로 파일 이름이 정해집니다";
    const s = c.summary;
    const comp = s.composition.map(g => `
      <tr class="cg"><td class="l"><b>${esc(g.label)}</b></td><td class="num"><b>${won(g.amount)}</b></td><td class="bar-cell"><span class="bar"><i class="k-${g.key}" style="width:${Math.min(100, g.share)}%"></i></span></td><td class="num"><b>${g.share.toFixed(1)}%</b></td></tr>
      ${g.items.map(it => `<tr class="ci"><td class="l">${esc(it.name)}</td><td class="num">${won(it.amount)}</td><td class="bar-cell"><span class="bar thin"><i class="k-${g.key}" style="width:${Math.min(100, it.share)}%"></i></span></td><td class="num">${it.share.toFixed(1)}%</td></tr>`).join("")}`).join("");
    return head("info", "프로젝트 개요", fileText,
      ctx.file ? btn("프로젝트 지우기", "deleteProject", [], "sm danger", "파일은 backups/deleted 로 옮겨집니다", "trash") : "")
      + `<div class="info-layout"><div class="form-grid">
      ${field("프로젝트 ID", inp("project_id", p.project_id), { help: ctx.file ? "이미 저장한 프로젝트는 ID 를 바꿔도 같은 파일에 저장됩니다" : "" })}
      ${field("프로젝트명", inp("project_name", p.project_name))}
      ${field("고객사", inp("client_name", p.client_name))}
      ${field("프로젝트 PM", inp("pm_name", p.pm_name, { ph: "예: 김철수 PM" }))}
      ${field("영업대표", inp("sales_rep", p.sales_rep, { ph: "예: 박영희 매니저" }))}
      ${field("기간 입력 단위", sel("date_mode", p.date_mode || "month", [{ v: "month", t: "월 (YYYY-MM)" }, { v: "date", t: "일자 (YYYY-MM-DD)" }]))}
      ${field("시작", inp("start_date", toMode(p.start_date), { type: dateType }))}
      ${field("종료", inp("end_date", toMode(p.end_date), { type: dateType }), { help: `${c.months.length}개월 · ${c.months[0] || ""} ~ ${c.months[c.months.length - 1] || ""}` })}
      ${field("계약 매출 (원)", inp("contract_revenue", p.contract_revenue, { money: true }))}
      ${field("목표 허들 이익률 (%)", inp("target_hurdle_rate", p.target_hurdle_rate, { type: "number", step: "0.1" }))}
      ${field("프로젝트 Type", inp("project_type", p.project_type, { ph: "예: SI-AD" }), { help: "보고서 양식의 {{프로젝트Type}}" })}
      ${field("프로젝트 설명", `<textarea rows="4" data-bind="description" data-type="text" placeholder="배경 · 범위 · 특이사항">${esc(p.description)}</textarea>`, { full: true })}
    </div>
    <section class="sub-card comp-card">${subHead("원가 구성", `총 원가 ${won(s.total_cost)}원 대비 비중`)}
      <table class="comp"><tbody>${s.total_cost > 0 ? comp : empty(4, "인력 · 경비를 넣으면 항목별 금액과 비중이 보입니다.")}</tbody>
      <tfoot><tr><td class="l">총 원가</td><td class="num">${won(s.total_cost)}</td><td></td><td class="num">${s.total_cost > 0 ? "100%" : "-"}</td></tr>
      <tr class="${s.revenue > 0 ? (s.hurdle_met ? "ok" : "bad") : ""}"><td class="l">예상 이익 (이익률)</td><td class="num">${won(s.profit)}</td><td></td><td class="num">${s.margin_rate.toFixed(1)}%</td></tr></tfoot></table>
    </section></div>`;
  }

  // ---------------- 인력 표 (내부 · 외주 공통) ----------------
  // 시스템(프로젝트 기준정보) · 고객 제안 직무 · 등급(마스터 — 견적서에 씀, 인력 표에서는 부가정보)
  const systemCol = p => ({ b: "system", t: "text", l: "시스템", th: `<th>시스템</th>`, cell: (r, base) => `<td>${sel(`${base}.system`, r.system || "", [{ v: "", t: "-" }].concat(p.settings.systems || []), { label: "시스템", cls: "s-md" })}</td>` });
  const siteCol = p => ({ b: "site", t: "text", l: "근무지", th: `<th>근무지</th>`, cell: (r, base) => `<td>${sel(`${base}.site`, r.site || "", [{ v: "", t: "-" }].concat(p.settings.work_sites || []), { label: "근무지", cls: "s-sm" })}</td>` });
  const propCols = p => [
    { b: "prop_role", t: "text", l: "제안직무", th: `<th class="prop" title="고객 제안 직무 — 견적서의 인건비를 만들 때 씀">제안 직무</th>`, cell: (r, base) => `<td class="prop">${sel(`${base}.prop_role`, r.prop_role || "", [{ v: "", t: "-" }].concat((p.settings.proposal_roles || []).map(x => x.role)), { label: "제안 직무", cls: "s-md" })}</td>` },
    { b: "prop_level", t: "text", l: "제안등급", th: `<th class="prop" title="고객 제안 등급">제안 등급</th>`, cell: (r, base) => `<td class="prop">${sel(`${base}.prop_level`, r.prop_level || "", [{ v: "", t: "-" }].concat(p.settings.proposal_levels || []), { label: "제안 등급", cls: "s-sm" })}</td>` }
  ];

  // 부서 칸: 부서가 2개 이상일 때만 (하나 이하면 지금과 같은 화면)
  const deptChoices = p => (p.settings.departments || []).filter(d => String(d).trim());
  const deptColumn = p => (M.multiDept(p) ? [{ b: "dept", t: "text", l: "부서", th: `<th>부서</th>`, cell: (r, base) => `<td>${sel(`${base}.dept`, r.dept || "", [{ v: "", t: "-" }].concat(deptChoices(p)), { label: "부서", cls: "s-md" })}</td>` }] : []);

  // 기준정보 순서로 줄을 다시 세움 (인력: 영역 → 등급(외주는 업체 → 등급), 경비: 계정)
  const sortBtn = list => btn("기준정보 순서로", "sortRows", [list], "sm", "담당 영역 · 등급 · 업체 · 계정을 기준정보에 등록한 순서대로 줄 세우기", "sort");

  // 매입단가 칸: 단가가 바뀌는 사람이 있으면 제목을 '최초 / 변경'으로 하고 칸 아래에 변경 단가 · 달 (없으면 지금 모양 그대로)
  const rateChanges = (c, list) => { const out = (c.details[list] || []).map(r => PC.Calc.rateChange(r, c.months)); return out.some(x => x.changed !== null) ? out : null; };
  const rateTitle = (label, chg, title) => `<th title="${esc(title)}">${chg ? `${esc(label)}<small class="th-sub">최초 / 변경</small>` : esc(label)}</th>`;
  const rateNote = (chg, i) => (chg && chg[i] && chg[i].changed !== null ? `<span class="cell-cost rate-change">→ ${won(chg[i].changed)} (${chg[i].changeYm.slice(2).replace("-", ".")}~)</span>` : "");

  function laborTab(list, p, c, cols, o) {
    const months = c.months, rows = c.details[list];
    // 투입 합계(금액 · MM)는 달 칸 왼쪽에
    cols = cols.concat([{ b: "@total", t: "skip", l: "합계", th: `<th>합계(원)</th>`, cell: r => `<td class="num"><b>${won(r.total_cost)}</b><span class="cell-cost">${r.total_mm} MM</span></td>` }]);
    const all = cols.concat(monthCols(months, "mm", (m, k) => (r, base) => { const st = k > 0 && r.calc[m].rate !== r.calc[months[k - 1]].rate; return `<td${st ? ` class="rate-step" title="${esc(m)}부터 단가 ${won(r.calc[m].rate)}원 (인상)"` : ""}>${mmInp(`${base}.mm.${m}`, r.calc[m].mm || "", `${m} MM`)}${r.calc[m].cost ? `<span class="cell-cost">${man(r.calc[m].cost)}만</span>` : ""}</td>`; }));
    const foot = rows.length ? `<tr class="total"><td></td><td colspan="${cols.length - 1}" class="l">합계 · ${sumBy(rows, r => r.total_mm).toFixed(2)} MM${list === "internal_labor" && num(p.ot_cost) ? ` (OT비용 ${won(p.ot_cost)}원 포함)` : ""}</td><td class="num">${won(o.total)}</td>${months.map(m => `<td class="num">${man(sumBy(rows, r => r.calc[m].cost))}만</td>`).join("")}</tr>` : "";
    return head(o.icon, o.title, GRID_HINT, sortBtn(list) + btn(o.addLabel, "addRow", [list], "primary sm", "", "plus")) + (o.top || "")
      + gridTable({ list, rows, cols: all, newLabel: "이름 입력", newKey: "name", foot, after: o.after });
  }
  // 총 MM: 직접 넣고 [고르게]로 전체 기간에 나누거나 [기간 배분]으로 고른 기간에만. 달별 MM 을 고치면 총 MM 은 그 합으로 바뀜
  const mmTarget = (list, r, base, i) => `<td class="c nowrap">${mmInp(`${base}.target_mm`, r.target_mm !== undefined ? r.target_mm : r.total_mm, "총 MM")}<button class="btn sm icon-only" data-act="distribute" data-args="${args(list, i)}" title="총 MM 을 전체 기간에 고르게 나눔" aria-label="고르게 배분">${icon("distribute")}</button><button class="btn sm icon-only" data-act="distributeRange" data-args="${args(list, i)}" title="기간을 골라 그 기간에만 나눔" aria-label="기간 배분">${icon("range")}</button></td>`;

  function internal(p, c, ctx) {
    const grades = (p.settings.internal_grades || []).map(g => g.grade), areas = p.settings.areas || [];
    const chg = rateChanges(c, "internal_labor");
    return laborTab("internal_labor", p, c, [
      { b: "area", t: "text", l: "영역", th: `<th>영역</th>`, cell: (r, base) => `<td>${sel(`${base}.area`, r.area || "", [{ v: "", t: "-" }].concat(areas), { label: "영역", cls: "s-md" })}</td>` },
      systemCol(p),
      { b: "role", t: "text", l: "역할", th: `<th class="l">역할</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.role`, r.role, { cls: "w-sm", ph: "PM/개발", label: "역할" })}</td>` },
      ...deptColumn(p),
      { b: "name", t: "text", l: "성명", th: `<th class="l">성명</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.name`, r.name, { cls: "w-sm", label: "성명" })}</td>` },
      { b: "grade", t: "text", l: "직급", th: `<th>직급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", grades, { label: "직급", cls: "s-sm" })}</td>` },
      { b: "tech_grade", t: "text", l: "기술등급", th: `<th>기술등급</th>`, cell: (r, base) => `<td>${sel(`${base}.tech_grade`, r.tech_grade || "", [{ v: "", t: "-" }].concat(p.settings.external_grades || []), { label: "기술등급", cls: "s-sm" })}</td>` },
      ...propCols(p),
      { b: "target_mm", t: "mm", l: "총MM", th: `<th>총 MM</th>`, cell: (r, base, i) => mmTarget("internal_labor", r, base, i) },
      { b: "base_rate", t: "blank-num", l: "매입단가", th: rateTitle("매입단가", chg, "비우면 직급 단가 · 바뀌는 단가는 [기준정보 > 인건비 단가]의 연봉 인상월 · 인상단가"), cell: (r, base, i) => `<td>${inp(`${base}.base_rate`, r.base_rate, { money: true, kind: "blank-num", cls: "w-md", label: "기준단가" })}${rateNote(chg, i)}</td>` },
      siteCol(p),
      { b: "overhead_eligible", t: "bool", l: "인비례", th: `<th title="내부 인비례 경비(MM × 기준금액)를 매길 인력">인비례</th>`, cell: (r, base) => `<td class="c">${chk(`${base}.overhead_eligible`, r.overhead_eligible, "인비례")}</td>` }
    ], { icon: "users", title: "내부 인력 투입 계획", addLabel: "내부 인력", total: c.summary.breakdown.internal_labor,
      top: `<div class="form-grid compact raise-bar">${field("OT비용 (프로젝트 전체, 원)", inp("ot_cost", p.ot_cost || "", { money: true, ph: "0", label: "OT비용" }), { help: "내부 인건비(노무비)에 들어가고 달에는 고르게 나눔" })}</div>` });
  }

  function external(p, c, ctx) {
    const vendors = [...new Set((p.settings.vendor_rates || []).map(v => v.vendor))];
    const grades = p.settings.external_grades || [], areas = p.settings.areas || [];
    const chg = rateChanges(c, "external_labor");
    return laborTab("external_labor", p, c, [
      { b: "area", t: "text", l: "영역", th: `<th>영역</th>`, cell: (r, base) => `<td>${sel(`${base}.area`, r.area || "", [{ v: "", t: "-" }].concat(areas), { label: "영역", cls: "s-md" })}</td>` },
      systemCol(p),
      { b: "role", t: "text", l: "역할", th: `<th class="l">역할</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.role`, r.role, { cls: "w-sm", ph: "개발/디자인", label: "역할" })}</td>` },
      ...deptColumn(p),
      { b: "vendor", t: "text", l: "협력사", th: `<th>협력사</th>`, cell: (r, base) => `<td>${sel(`${base}.vendor`, r.vendor || "", vendors, { label: "외주업체", cls: "s-md" })}</td>` },
      { b: "name", t: "text", l: "성명", th: `<th class="l">성명</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.name`, r.name, { cls: "w-sm", label: "성명" })}</td>` },
      { b: "grade", t: "text", l: "기술등급", th: `<th>기술등급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", grades, { label: "기술등급", cls: "s-sm" })}</td>` },
      ...propCols(p),
      { b: "target_mm", t: "mm", l: "총MM", th: `<th>총 MM</th>`, cell: (r, base, i) => mmTarget("external_labor", r, base, i) },
      { b: "custom_rate", t: "blank-num", l: "매입단가", th: rateTitle("매입단가", chg, "비워 두면 협력사 · 기술등급 단가표의 단가 (인상월 · 인상단가도 단가표대로)"), cell: (r, base, i) => `<td>${inp(`${base}.custom_rate`, r.custom_rate, { money: true, kind: "blank-num", cls: "w-md", ph: won(r.default_rate), label: "적용단가" })}${r.is_overridden ? `<span class="cell-cost">기준 ${won(r.default_rate)}</span>` : ""}${rateNote(chg, i)}</td>` },
      siteCol(p)
    ], { icon: "partner", title: "외주 인력 투입 계획", addLabel: "외주 인력", total: c.summary.breakdown.external_labor });
  }

  // ---------------- 경비 (일반 · 인비례 · 기타) ----------------
  const MODES = [{ v: "equal_total", t: "총액 균등 분배" }, { v: "fill_all", t: "매달 같은 금액" }, { v: "manual", t: "월별 직접 입력" }];
  function expenseGrid(list, p, c, label) {
    const months = c.months, rows = c.details[list];
    const cat = list === "general_expenses" ? "general" : "custom";
    const accounts = (p.settings.expense_accounts || []).filter(a => a.category === cat).map(a => `<option value="${esc(a.account)}"></option>`).join("");
    const cols = [
      { b: "name", t: "text", l: "계정", th: `<th class="l">계정 / 항목</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.name`, r.name, { cls: "w-lg", list: `acc_${list}`, label: "계정" })}</td>` },
      ...(M.multiDept(p) ? [{ b: "dept", t: "dept", l: "부서", th: `<th title="비우면 공통 — 부서별 손익에서 기준에 따라 나눔">부서</th>`, cell: (r, base) => `<td>${sel(`${base}.dept`, r.dept || "", [{ v: "", t: "공통" }].concat(deptChoices(p)), { label: "부서", cls: "s-md" })}</td>` }] : []),
      { b: "dist_mode", t: "mode", l: "분배", th: `<th>분배 방식</th>`, cell: (r, base) => `<td>${sel(`${base}.dist_mode`, r.dist_mode || "equal_total", MODES, { label: "분배 방식", cls: "s-lg" })}</td>` },
      { b: "@amount", t: "num", l: "금액", th: `<th title="총액 균등: 전체 금액 / 매달 같은 금액: 한 달 금액">입력 금액</th>`, cell: (r, base) => r.dist_mode === "manual" ? `<td class="c"><span class="tag">달마다</span></td>`
        : `<td>${inp(r.dist_mode === "fill_all" ? `${base}.unit_input` : `${base}.total_input`, r.dist_mode === "fill_all" ? r.unit_input : r.total_input, { money: true, cls: "w-md", label: "입력 금액" })}</td>` },
      ...months.map(m => ({ b: `monthly.${m}`, t: "num", l: m, fill: true, th: `<th class="m">${esc(m.slice(2).replace("-", "."))}</th>`,
        cell: (r, base) => (r.dist_mode === "manual" ? `<td>${inp(`${base}.monthly.${m}`, r.calc_monthly[m] || "", { money: true, cls: "w-sm", ph: "0", label: `${m} 금액` })}</td>` : `<td class="num muted">${won(r.calc_monthly[m])}</td>`) }))
    ];
    const foot = rows.length ? `<tr class="total"><td></td><td colspan="${cols.length - months.length}" class="l">${label} 합계</td>${months.map(m => `<td class="num">${won(sumBy(rows, r => r.calc_monthly[m]))}</td>`).join("")}<td class="num">${won(sumBy(rows, r => r.total_cost))}</td></tr>` : "";
    return `<datalist id="acc_${list}">${accounts}</datalist>` + gridTable({ list, rows, cols, newLabel: "계정 입력", foot, tailHead: `<th>합계(원)</th>`, tail: r => `<td class="num"><b>${won(r.total_cost)}</b></td>` });
  }
  function expense(p, c) {
    const months = c.months, d = c.details, b = c.summary.breakdown;
    const accs = p.settings.expense_accounts || [];
    const mmRows = (d.mm_overheads || []).map(o => {
      const i = accs.findIndex(a => a.account === o.account && a.category === o.category);
      return `<tr><td class="l"><b>${esc(o.account)}</b></td><td class="c"><span class="tag">${o.category === "mm_internal" ? "내부 · 인비례 대상" : "외주"}</span></td>
        <td>${i >= 0 ? inp(`settings.expense_accounts.${i}.unit_amount`, o.unit_amount, { money: true, cls: "w-md", label: "기준금액(원/MM)" }) : won(o.unit_amount)}</td>
        ${months.map(m => `<td class="num">${won(o.calc_monthly[m])}<span class="cell-cost">${o.mm_monthly[m]} MM</span></td>`).join("")}<td class="num"><b>${won(o.total_cost)}</b></td></tr>`;
    }).join("");
    const mmHead = `<tr><th class="l">계정</th><th>대상 인력</th><th>기준금액 (원/MM)</th>${months.map(m => `<th class="m">${esc(m.slice(2).replace("-", "."))}</th>`).join("")}<th>합계(원)</th></tr>`;
    const totalAll = b.general_expense + b.internal_overhead + b.external_overhead + b.custom_cost;
    return head("wallet", "경비 · 재료비", `${b.material ? `재료 ${won(b.material)} · ` : ""}일반 ${won(b.general_expense)} · 인비례 ${won(b.internal_overhead + b.external_overhead)} · 기타 ${won(b.custom_cost)} = ${won(totalAll + b.material)}원`, "")
      + `<div class="stack">
      <section>${subHead("일반경비", "금액을 넣고 달에 나누는 방식을 고릅니다 · " + GRID_HINT, sortBtn("general_expenses") + btn("일반경비", "addRow", ["general_expenses"], "primary sm", "", "plus"))}${expenseGrid("general_expenses", p, c, "일반경비")}</section>
      <section>${subHead("재료비", "구분 · 내용 · 금액 — 총 원가에 들어가고 달에는 고르게 나눔 · " + GRID_HINT, btn("재료비", "addRow", ["materials"], "sm", "", "plus"))}
        ${gridTable({ list: "materials", rows: p.materials || [], newLabel: "구분 입력", cols: [
          { b: "category", t: "text", l: "구분", th: `<th class="l">구분</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.category`, r.category, { cls: "w-md", label: "구분" })}</td>` },
          { b: "content", t: "text", l: "내용", th: `<th class="l">내용</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.content`, r.content, { cls: "w-xl", label: "내용" })}</td>` },
          { b: "amount", t: "num", l: "금액", th: `<th>금액(원)</th>`, cell: (r, base) => `<td>${inp(`${base}.amount`, r.amount, { money: true, cls: "w-lg", label: "금액" })}</td>` }],
          foot: (p.materials || []).length ? `<tr class="total"><td></td><td colspan="2" class="l">재료비 합계</td><td class="num">${won(b.material)}</td></tr>` : "" })}</section>
      <section>${subHead("인비례 경비 (자동)", "계정별로 인력 MM × 기준금액 — 내부는 [인비례]에 체크한 인력, 외주는 전체 외주 인력. 계정 추가 · 삭제는 [기준정보 > 경비 계정]")}
        <div class="tbl-wrap"><table><thead>${mmHead}</thead><tbody>${mmRows || empty(months.length + 4, "인비례 계정이 없습니다. [기준정보 > 경비 계정]에서 구분을 '인비례'로 추가하세요.")}
        ${mmRows ? `<tr class="total"><td colspan="3" class="l">인비례 합계</td>${months.map(m => `<td class="num">${won(num(d.internal_overhead_monthly[m]) + num(d.external_overhead_monthly[m]))}</td>`).join("")}<td class="num">${won(b.internal_overhead + b.external_overhead)}</td></tr>` : ""}</tbody></table></div></section>
      <section>${subHead("기타경비", "", sortBtn("custom_costs") + btn("기타경비", "addRow", ["custom_costs"], "sm", "", "plus"))}${expenseGrid("custom_costs", p, c, "기타경비")}</section>
    </div>`;
  }

  // ---------------- 분할계약 ----------------
  function phase(p, c) {
    const editor = (p.phases || []).map((ph, i) => `<tr>
      <td class="c">${del("phases", i)}</td>
      <td class="l">${inp(`phases.${i}.title`, ph.title, { cls: "w-lg", label: "차수명" })}</td>
      <td>${inp(`phases.${i}.start_ym`, ph.start_ym, { type: "month", label: "시작월" })}</td>
      <td>${inp(`phases.${i}.end_ym`, ph.end_ym, { type: "month", label: "종료월" })}</td>
      <td>${inp(`phases.${i}.revenue`, ph.revenue, { money: true, cls: "w-lg", label: "회차 매출" })}</td>
      <td>${inp(`phases.${i}.override_cost`, ph.override_cost, { money: true, kind: "blank-num", cls: "w-lg", ph: "비우면 기간 원가 자동", label: "별도 원가" })}</td></tr>`).join("");
    const summary = c.phases.map(x => `<tr><td class="l"><b>${esc(x.title)}</b></td><td class="c">${esc(x.start_ym)} ~ ${esc(x.end_ym)}</td>
      <td class="num">${won(x.revenue)}</td><td class="num"><b>${won(x.cost)}</b>${x.overridden ? `<span class="cell-cost">별도 원가 · 계산 ${won(x.calc_cost)}</span>` : ""}</td>
      <td class="num">${won(x.internal_labor)}</td><td class="num">${won(x.external_labor)}</td><td class="num">${won(x.expense)}</td><td class="num"><span class="tag ${x.margin_rate >= c.summary.hurdle_rate ? "ok" : "bad"}">${x.margin_rate}%</span></td>
      <td class="num">${won(x.cum_revenue)}</td><td class="num">${won(x.cum_cost)}</td><td class="num">${won(x.rem_revenue)}</td><td class="num">${won(x.rem_cost)}</td></tr>`).join("");
    const revSum = (p.phases || []).reduce((a, x) => a + num(x.revenue), 0);
    const gap = Math.round(num(p.contract_revenue) - revSum);
    const full = (p.phases || []).length >= 5;
    return head("split", "분할계약", gap && p.phases.length ? `회차 매출 합이 계약 매출보다 ${won(Math.abs(gap))}원 ${gap > 0 ? "적습니다" : "많습니다"}` : "최대 5회차",
      full ? `<button class="btn primary sm" disabled title="최대 5회차">${icon("plus")}회차</button>` : btn("회차", "addRow", ["phases"], "primary sm", "", "plus"))
      + `<div class="stack"><div class="tbl-wrap"><table class="fill"><thead><tr><th class="w-del"></th><th class="l">차수명</th><th>시작월</th><th>종료월</th><th>회차 매출(원)</th><th>별도 원가(원)</th></tr></thead>
        <tbody>${editor || empty(6, "등록된 분할계약이 없습니다.")}</tbody></table></div>
      <div class="tbl-wrap"><table><thead><tr><th class="l" rowspan="2">회차</th><th rowspan="2">기간</th><th rowspan="2">회차 매출</th><th colspan="4" class="grp">회차 원가</th><th rowspan="2">이익률</th><th rowspan="2">누적 매출</th><th rowspan="2">누적 원가</th><th rowspan="2">잔여 매출</th><th rowspan="2">잔여 원가</th></tr>
        <tr><th>합계</th><th>내부 인건비</th><th>외주 인건비</th><th title="일반 · 인비례 · 기타경비">경비</th></tr></thead>
        <tbody>${summary || empty(12, "회차를 추가하면 회차 · 누적 · 잔여 손익이 보입니다.")}</tbody></table></div>${phaseDepts(p, c)}</div>`;
  }

  // ---------------- 견적서 ----------------
  const ROUND_UNITS = [{ v: "1", t: "절사 안 함" }, { v: "10", t: "10원 단위" }, { v: "100", t: "100원 단위" }, { v: "1000", t: "1,000원 단위" }, { v: "10000", t: "10,000원 단위" }, { v: "100000", t: "100,000원 단위" }, { v: "1000000", t: "1,000,000원 단위" }];
  function estimate(p, c) {
    const est = c.estimate, st = p.settings;
    const roles = (st.proposal_roles || []).map(x => x.role), levels = st.proposal_levels || [];
    const labor = est.labor.map((l, i) => `<tr${l.missing ? ' class="bad-row"' : ""}>
      <td class="c">${del("estimate.labor", i)}</td>
      <td>${sel(`estimate.labor.${i}.role`, l.role || "", [{ v: "", t: "고르세요" }].concat(roles), { label: "제안 직무", cls: "s-md" })}</td>
      <td>${sel(`estimate.labor.${i}.level`, l.level || "", [{ v: "", t: "고르세요" }].concat(levels), { label: "제안 등급", cls: "s-sm" })}</td>
      <td>${mmInp(`estimate.labor.${i}.mm`, l.mm || "", "견적 MM", "w-xs")}</td>
      <td>${inp(`estimate.labor.${i}.unit_price`, l.unit_price, { money: true, kind: "blank-num", cls: "w-lg", ph: won(l.master_rate), label: "제안 단가" })}</td>
      <td class="num"><b>${won(l.amount)}</b></td>
      <td class="l">${inp(`estimate.labor.${i}.memo`, l.memo, { cls: "w-lg", label: "비고" })}</td></tr>`).join("");
    const items = est.items.map((it, i) => `<tr>
      <td class="c">${del("estimate.items", i)}</td>
      <td class="l" colspan="2">${inp(`estimate.items.${i}.item_name`, it.item_name, { cls: "w-xl", label: "항목명" })}</td>
      <td>${inp(`estimate.items.${i}.qty`, it.qty, { type: "number", step: "0.01", cls: "w-xs", label: "수량" })}</td>
      <td>${inp(`estimate.items.${i}.unit_price`, it.unit_price, { money: true, cls: "w-lg", label: "단가" })}</td>
      <td class="num"><b>${won(it.amount)}</b></td>
      <td class="l">${inp(`estimate.items.${i}.memo`, it.memo, { cls: "w-lg", label: "비고" })}</td></tr>`).join("");
    const planMm = sumBy(c.details.internal_labor.concat(c.details.external_labor), r => r.total_mm), estMm = sumBy(est.labor, l => l.mm);
    const missing = est.labor.filter(l => l.missing).length;
    return head("receipt", "견적서", est.applied_to_revenue ? "공급가액이 계약 매출에 적용됨" : "인건비는 고객 제안 직무 · 등급 × 견적 MM × 제안 단가",
      btn("공급가액 → 계약 매출", "applyEstimate", [], "primary sm", "견적 공급가액을 계약 매출로 씁니다"))
      + `<div class="form-grid compact">${field("견적번호", inp("estimate.estimate_no", p.estimate.estimate_no, { ph: "EST-2026-001" }))}${field("견적일자", inp("estimate.estimate_date", p.estimate.estimate_date, { type: "date" }))}${field("비고", inp("estimate.note", p.estimate.note, { ph: "조건 · 유효기간" }))}</div>
      <section class="mt">${subHead("인건비 (고객 제안 직무 · 등급)", `견적 MM ${estMm.toFixed(2)} · 인력 계획 MM ${planMm.toFixed(2)}${missing ? ` · 직무 · 등급이 빠진 줄 ${missing}개` : ""}`,
        btn("인력 계획에서 가져오기", "estimateFromPlan", [], "sm", "인력의 제안 직무 · 등급별로 MM 을 모아 줄을 만듭니다 (견적 MM 은 고쳐도 됨)", "users") + btn("줄", "addRow", ["estimate.labor"], "sm", "", "plus"))}
      <div class="tbl-wrap"><table class="fill est"><thead><tr><th class="w-del"></th><th>제안 직무</th><th>제안 등급</th><th>견적 MM</th><th title="비우면 기준정보의 제안 단가">제안 단가(원/MM)</th><th>금액(원)</th><th class="l">비고</th></tr></thead><tbody>
        ${labor || empty(7, "[인력 계획에서 가져오기] 또는 [+ 줄]로 제안 직무 · 등급별 인건비를 넣습니다.")}
        <tr class="total"><td></td><td colspan="2" class="l">인건비 소계</td><td class="num">${estMm.toFixed(2)}</td><td></td><td class="num">${won(est.labor_total)}</td><td></td></tr></tbody></table></div></section>
      <section class="mt">${subHead("기타 항목 (SW · HW · 용역 등)", "", btn("품목", "addRow", ["estimate.items"], "sm", "", "plus"))}
      <div class="tbl-wrap"><table class="fill est"><thead><tr><th class="w-del"></th><th class="l" colspan="2">항목명</th><th>수량</th><th>단가(원)</th><th>금액(원)</th><th class="l">비고</th></tr></thead><tbody>
        ${items || empty(7, "등록된 품목이 없습니다.")}
        <tr class="total"><td></td><td colspan="4" class="l">기타 항목 소계</td><td class="num">${won(est.items_total)}</td><td></td></tr>
      </tbody></table></div></section>
      <section class="mt">${subHead("견적 금액", "합계 → 할인 → 절사 → 견적 금액")}
      <div class="tbl-wrap"><table class="est-sum"><tbody>
        <tr><td class="l">합계 (인건비 + 기타 항목)</td><td></td><td class="num">${won(est.subtotal)}</td></tr>
        <tr><td class="l">할인</td><td class="c nowrap">${inp("estimate.discount_rate", p.estimate.discount_rate || "", { type: "number", step: "0.1", kind: "num", cls: "w-xs", ph: "0", label: "할인율(%)" })} %</td><td class="num">${est.discount ? `− ${won(est.discount)}` : "-"}</td></tr>
        <tr><td class="l">절사</td><td class="c">${sel("estimate.round_unit", String(p.estimate.round_unit || 1), ROUND_UNITS, { kind: "num", label: "절사 단위", cls: "s-md" })}</td><td class="num">${est.truncation ? `− ${won(est.truncation)}` : "-"}</td></tr>
        <tr class="total"><td class="l">견적 금액</td><td></td><td class="num">${won(est.supply_total)}</td></tr>
      </tbody></table></div></section>`;
  }

  // ---------------- 기준정보: 작은 탭으로 한 번에 한 목록 (목록 길이가 달라도 화면이 어긋나지 않게) ----------------
  // 기준정보 = 프로젝트 설정(단가 · 금액 · 프로젝트마다 다른 것) + 마스터(전사 공통 목록 — 금액 없음, master_rates.xlsx 로 배포)
  const SETTING_GROUPS = [
    { id: "project", label: "프로젝트 설정", desc: "이 프로젝트에만 적용 — 단가 · 금액 · 시스템", tabs: [
      { id: "rates", label: "인건비 단가" },
      { id: "vendors", label: "외주업체 단가", count: st => (st.vendor_rates || []).length },
      { id: "proposal_rates", label: "제안 단가" },
      { id: "systems", label: "시스템", count: st => (st.systems || []).length },
      { id: "forms", label: "보고서 양식 · 샘플" }] },
    { id: "master", label: "마스터 (전사 공통)", desc: "전사가 같이 쓰는 목록 — 금액 없음. 고친 뒤 [마스터 파일 > 마스터로 저장]으로 배포", tabs: [
      { id: "grades", label: "직급 · 기술등급", count: st => (st.internal_grades || []).length + (st.external_grades || []).length },
      { id: "proposal", label: "제안 직무 · 등급", count: st => (st.proposal_roles || []).length },
      { id: "areas", label: "영역", count: st => (st.areas || []).length },
      { id: "sites", label: "근무지", count: st => (st.work_sites || []).length },
      { id: "accounts", label: "경비 계정", count: st => (st.expense_accounts || []).length },
      { id: "depts", label: "부서", count: st => (st.departments || []).length },
      { id: "approvers", label: "결재자" },
      { id: "masterfile", label: "마스터 파일" }] }
  ];
  const ALL_SETTING_TABS = SETTING_GROUPS.flatMap(g => g.tabs.map(t => Object.assign({ group: g.id }, t)));
  function settings(p, c, ctx) {
    const st = p.settings;
    const sub = ALL_SETTING_TABS.some(t => t.id === ctx.sub) ? ctx.sub : "rates";
    const group = SETTING_GROUPS.find(g => g.tabs.some(t => t.id === sub));
    const groupBar = `<div class="set-groups">${SETTING_GROUPS.map(g => `<button class="set-group${g === group ? " active" : ""}" data-act="showSettings" data-args="${args(g.tabs[0].id)}"><b>${esc(g.label)}</b><small>${esc(g.desc)}</small></button>`).join("")}</div>`;
    const subTabs = `<div class="seg" role="tablist">${group.tabs.map(t => `<button class="seg-btn${t.id === sub ? " active" : ""}" data-act="showSettings" data-args="${args(t.id)}">${esc(t.label)}${t.count ? `<span class="cnt">${t.count(st)}</span>` : ""}</button>`).join("")}</div>`;
    const narrow = html => `<div class="narrow">${html}</div>`;
    const text = (b, label, cls) => ({ b, t: "text", l: label, th: `<th class="l">${esc(label)}</th>`, cell: (r, base) => `<td class="l">${inp(b ? `${base}.${b}` : base, b ? r[b] : r, { cls: cls || "w-lg", label })}</td>` });
    const roCol = (b, label) => ({ b: "@ro", t: "skip", l: label, th: `<th class="l">${esc(label)}</th>`, cell: r => `<td class="l ro">${esc(b ? r[b] : r)}</td>` });   // 마스터에서 정한 이름 (여기서는 못 고침)
    const moneyCol = (b, label) => ({ b, t: "num", l: label, th: `<th>${esc(label)}</th>`, cell: (r, base) => `<td>${inp(`${base}.${b}`, r[b], { money: true, cls: "w-md", label })}</td>` });
    const blankMoney = (b, label, ph) => ({ b, t: "blank-num", l: label, th: `<th title="비우면 인상 없음">${esc(label)}</th>`, cell: (r, base) => `<td>${inp(`${base}.${b}`, r[b], { money: true, kind: "blank-num", cls: "w-md", ph: ph || "인상 없음", label })}</td>` });
    const fromMaster = what => `<p class="muted">${esc(what)}의 목록은 [마스터] 쪽에서 더하거나 고칩니다.</p>`;
    const raiseYm = PC.Calc.raiseMonth(st, c.months);
    let body = "";
    // ---- 프로젝트 설정 ----
    if (sub === "rates") body = narrow(subHead("내부 직급별 월 기준단가 · 연봉 인상", "직급을 고르면 인력의 기준단가가 이 값으로 들어갑니다 · " + GRID_HINT)
      + `<div class="form-grid compact raise-bar">${field("연봉 인상월 (MM)", inp("settings.raise_month", st.raise_month, { ph: "예: 03", kind: "month2b" }), { help: raiseYm ? `이 프로젝트에서는 ${raiseYm} 부터 인상단가` : "비우거나 기간 안에 그 달이 없으면 인상 없음" })}</div>`
      + gridTable({ fixed: true, list: "settings.internal_grades", rows: st.internal_grades || [], cols: [roCol("grade", "직급"), moneyCol("rate", "월 기준단가 (원)"), blankMoney("raise_rate", "인상단가 (원)")] })
      + fromMaster("직급"));
    else if (sub === "vendors") body = narrow(subHead("외주업체별 기술등급 단가", "외주 인력의 적용단가를 비워 두면 이 단가 · 협상으로 바뀌면 인상월 · 인상단가 · " + GRID_HINT)
      + gridTable({ order: true, list: "settings.vendor_rates", rows: st.vendor_rates || [], newLabel: "업체 입력", cols: [text("vendor", "업체"),
        { b: "grade", t: "text", l: "기술등급", th: `<th>기술등급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", st.external_grades || [], { label: "기술등급", cls: "s-sm" })}</td>` },
        moneyCol("rate", "월 단가 (원)"),
        { b: "raise_ym", t: "month", l: "인상월", th: `<th title="협상으로 단가가 바뀌는 달">인상월</th>`, cell: (r, base) => `<td>${inp(`${base}.raise_ym`, r.raise_ym, { type: "month", label: "인상월" })}</td>` },
        blankMoney("raise_rate", "인상단가 (원)")] }));
    else if (sub === "proposal_rates") {
      const levels = st.proposal_levels || [];
      body = subHead("고객 제안 단가 (원/MM)", "견적서의 인건비 = 견적 MM × 이 단가 · " + GRID_HINT)
        + gridTable({ fixed: true, list: "settings.proposal_roles", rows: st.proposal_roles || [], cols: [roCol("role", "제안 직무"),
          ...levels.map(l => ({ b: `rates.${l}`, t: "blank-num", l, th: `<th>${esc(l)}</th>`, cell: (r, base) => `<td>${inp(`${base}.rates.${l}`, (r.rates || {})[l], { money: true, kind: "blank-num", cls: "w-md", label: `${r.role} ${l}` })}</td>` }))] })
        + fromMaster("제안 직무 · 등급");
    }
    else if (sub === "systems") body = narrow(subHead("시스템", "인력마다 시스템 하나 — [투입 현황]에서 시스템별로 봅니다")
      + gridTable({ order: true, list: "settings.systems", rows: st.systems || [], newLabel: "시스템 입력", cols: [text("", "시스템")] }));
    else if (sub === "forms") body = narrow(`<section class="sub-card">${subHead("보고서 양식", "머리줄 [사업검토서]")}<p class="muted">자료 폴더의 templates 에 있는 엑셀 양식(xlsx)의 {{자리표시}} 칸에 값을 넣어 새 파일로 만듭니다. 양식은 Excel 에서 직접 고치고, 값이 들어갈 칸에 아래 이름을 {{이름}} 으로 적으면 됩니다. 칸에 자리표시 하나만 있으면 숫자는 숫자로 들어가 양식의 통화 · % 서식이 적용됩니다(비율은 0.4667 같은 소수).</p>
        <details class="fold"><summary>자리표시 목록 (${PC.Template.KEYS.length}개) — 눌러서 펼치기 · 접기</summary>
        <div class="tbl-wrap flat"><table class="keys fit"><thead><tr><th class="l">자리표시</th><th class="l">값</th><th class="l">이 프로젝트에서</th></tr></thead><tbody>${(() => { const v = PC.Template.values(p, c, {}); return PC.Template.KEYS.map(([k, t]) => `<tr><td class="l"><code>{{${esc(k)}}}</code></td><td class="l">${esc(t)}</td><td class="l">${esc(typeof v[k] === "number" ? (Math.abs(v[k]) < 1 && v[k] !== 0 ? `${(v[k] * 100).toFixed(2)}%` : won(v[k])) : v[k])}</td></tr>`).join(""); })()}</tbody></table></div></details></section>
      <section class="sub-card">${subHead("샘플 프로젝트")}<p class="muted">화면을 익히거나 시연할 때 — 예시 자료가 든 새 프로젝트를 엽니다 (지금 프로젝트는 그대로). 직급 · 직무 · 영역 등 이름은 이 PC 의 마스터 목록으로, 단가 · 금액은 예시 값으로 채웁니다 (결재자도 가짜 이름).</p>${btn("샘플 프로젝트 추가", "addSample", [], "", "예시 자료로 새 프로젝트", "plus")}</section>
      <section class="sub-card admin-card"><div class="sec-head"><h3>배포용 기본 양식 만들기</h3><span class="tag admin-tag">관리자용</span><span class="ph-desc">프로그램을 배포하기 전에 양식을 바꿀 때</span></div>
        <ol class="muted steps-list"><li>고친 양식(xlsx)을 연결한 자료 폴더의 templates 에 넣습니다 (같은 이름이면 기본 양식을 바꾸고, 새 이름이면 기본 양식에 더해짐).</li><li>[배포용 기본 양식 만들기] → builtin-templates.js 를 내려받습니다.</li><li>프로그램 폴더의 js/builtin-templates.js 를 이 파일로 바꾼 뒤 압축해 배포합니다.</li></ol>
        <p class="muted">폴더를 연결하기 전에는 프로그램 안의 기본 양식으로 출력하므로, 프로그램 폴더의 엑셀만 고쳐서는 바뀌지 않습니다.</p>
        ${btn("배포용 기본 양식 만들기", "buildBuiltinTemplates", [], "", "연결한 폴더의 templates 양식을 담은 js/builtin-templates.js 를 내려받음", "download")}</section>`);
    // ---- 마스터 (전사 공통 목록) ----
    else if (sub === "grades") body = `<div class="split-2 even"><div>${subHead("내부 직급", "위에서부터 높은 직급 순서로 — 단가는 [프로젝트 설정 > 인건비 단가]")}${gridTable({ order: true, list: "settings.internal_grades", rows: st.internal_grades || [], newLabel: "직급 입력", cols: [text("grade", "직급", "w-md")] })}</div>
      <div>${subHead("기술등급", "위에서부터 높은 등급 순서로 · 내부 · 외주 인력이 같이 씀 — 업체별 단가는 [프로젝트 설정 > 외주업체 단가]")}${gridTable({ order: true, list: "settings.external_grades", rows: st.external_grades || [], newLabel: "기술등급 입력", cols: [text("", "기술등급", "w-md")] })}</div></div>`;
    else if (sub === "proposal") body = `<div class="split-2 even"><div>${subHead("고객 제안 직무", "단가는 [프로젝트 설정 > 제안 단가]")}${gridTable({ order: true, list: "settings.proposal_roles", rows: st.proposal_roles || [], newLabel: "직무 입력", cols: [text("role", "제안 직무", "w-md")] })}</div>
      <div>${subHead("제안 등급", "모든 직무가 같은 등급 체계 · 위에서부터 높은 등급 순서로 (Lvl1 · 1등급 처럼 번호로만 다르면 순서 상관없음)")}${gridTable({ order: true, list: "settings.proposal_levels", rows: st.proposal_levels || [], newLabel: "등급 입력", cols: [text("", "제안 등급", "w-md")] })}</div></div>`;
    else if (sub === "areas") body = narrow(subHead("영역", "인력마다 고르는 영역 (컨설팅 · 개발 …) — [투입 현황]에서 영역별로 MM · 금액을 봅니다")
      + gridTable({ order: true, list: "settings.areas", rows: st.areas || [], newLabel: "영역 입력", cols: [text("", "영역")] }));
    else if (sub === "sites") body = narrow(subHead("근무지", "인력마다 고르는 근무지 (본사 · 고객사 · 원격 …)")
      + gridTable({ order: true, list: "settings.work_sites", rows: st.work_sites || [], newLabel: "근무지 입력", cols: [text("", "근무지")] }));
    else if (sub === "accounts") body = subHead("경비 계정", "인비례 계정은 인력 MM × 기준금액(원/MM)으로 자동 계산 — 기준금액은 프로젝트마다 [경비] 탭에서 넣습니다")
      + gridTable({ order: true, list: "settings.expense_accounts", rows: st.expense_accounts || [], newLabel: "계정 입력", cols: [text("account", "계정"),
        { b: "category", t: "cat", l: "구분", th: `<th>구분</th>`, cell: (r, base) => `<td>${sel(`${base}.category`, r.category || "general", M.CATEGORIES, { label: "구분", cls: "s-md" })}</td>` },
        text("desc", "설명", "w-xl")] });
    else if (sub === "depts") body = narrow(subHead("부서", "부서를 2개 이상 넣으면 인력 · 경비에 부서 칸과 [부서별] 탭이 나타납니다. 하나 이하면 지금 화면 그대로입니다")
      + gridTable({ order: true, list: "settings.departments", rows: st.departments || [], newLabel: "부서 입력", cols: [text("", "부서")] }));
    else if (sub === "approvers") body = narrow(`<section class="sub-card">${subHead("결재자", "보고서 양식의 {{결재1}} … {{결재5}} — 인사이동이 있을 때만 바꿈")}<div class="approvers">${APPROVER_TITLES.map((t, i) => `<label><small>${esc(t)}</small>${inp(`settings.approvers.${i}`, (st.approvers || [])[i], { ph: "이름", label: t })}</label>`).join("")}</div></section>`);
    else {
      const mv = (ctx.master && ctx.master.master_version) || {}, pv = st.master_version || {};
      const mvText = v => (v && (v.name || v.date) ? `${v.name || "이름 없음"}${v.date ? ` (${v.date})` : ""}` : "이름 없음");
      body = narrow(`<section class="sub-card">${subHead("마스터 파일 (master/master_rates.xlsx)", "직급 · 기술등급 · 제안 직무/등급 · 영역 · 근무지 · 경비 계정 · 부서 · 결재자 — 금액은 넣지 않음")}
        <p class="muted">마스터는 전사가 같이 쓰는 목록을 배포하는 기본값입니다. 새 프로젝트는 마스터의 목록으로 시작하고, 프로젝트에서 항목을 더 넣을 수 있습니다. 단가 · 금액 · 외주업체 · 시스템은 프로젝트마다 [프로젝트 설정]에서 넣습니다.</p>
        <div class="mini-kpis"><div class="mini-kpi"><span>폴더의 마스터</span><b>${esc(mvText(mv))}</b></div><div class="mini-kpi"><span>이 프로젝트가 쓴 마스터</span><b>${esc(mvText(pv))}</b></div></div>
        <div class="bar-group">${btn("마스터를 이 프로젝트에 반영", "applyMaster", [], "primary", "마스터에만 있는 항목을 이 프로젝트에 더함 (이미 있는 항목 · 금액은 그대로)", "download")}
          <label class="btn file-btn" title="본사에서 받은 master_rates.xlsx 로 폴더의 마스터를 바꿉니다">${icon("upload")}마스터 파일 올리기<input type="file" accept=".xlsx" id="masterInput"></label>
          ${btn("이 목록을 마스터로 저장", "saveDefaults", [], "", "이 프로젝트의 마스터 목록을 폴더의 마스터로 (금액 · 프로젝트 전용 제외)", "save")}${btn("마스터 내보내기", "exportMaster", [], "", "배포용 master_rates.xlsx 내려받기", "report")}</div></section>`);
    }
    return head("settings", "기준정보", group.desc, "") + groupBar + subTabs + `<div class="stack mt">${body}</div>`;
  }

  // ---------------- 투입 현황: 영역 · 등급 · 역할 · 업체별 MM · 금액 (분할계약 기간만도) ----------------
  const STAT_BY = [{ v: "area", t: "영역" }, { v: "system", t: "시스템" }, { v: "grade", t: "직급 · 기술등급" }, { v: "proposal", t: "제안 직무 · 등급" }, { v: "role", t: "역할" }, { v: "vendor", t: "업체 (내부 / 외주업체)" }, { v: "kind", t: "내부 · 외주" }];
  const STAT_SCOPE = [{ v: "all", t: "전체 인력" }, { v: "internal", t: "내부만" }, { v: "external", t: "외주만" }];
  function stats(p, c, ctx) {
    const o = ctx.stats || {}, by = o.by === "dept" && !M.multiDept(p) ? "area" : o.by || "area", scope = o.scope || "all", ph = c.phases[o.phase - 1];
    const months = ph ? PC.monthRange(ph.start_ym, ph.end_ym) : null;
    const r = PC.Calc.staffStats(c, { by, scope, months, settings: p.settings });
    const seg = (key, list, cur) => `<div class="seg">${list.map(x => `<button class="seg-btn${x.v === cur ? " active" : ""}" data-act="setStats" data-args="${args(key, x.v)}">${esc(x.t)}</button>`).join("")}</div>`;
    const phaseSel = `<select data-change-act="setStats" data-key="phase" aria-label="기간"><option value="0">전체 기간 (${c.months[0] || ""} ~ ${c.months[c.months.length - 1] || ""})</option>${c.phases.map((x, i) => `<option value="${i + 1}"${o.phase === i + 1 ? " selected" : ""}>${esc(x.title)} (${x.start_ym} ~ ${x.end_ym})</option>`).join("")}</select>`;
    const byList = M.multiDept(p) ? [{ v: "dept", t: "부서" }].concat(STAT_BY) : STAT_BY;
    const label = (byList.find(x => x.v === by) || STAT_BY[0]).t;
    const mh = r.months.map(m => `<th class="m">${esc(m.slice(2).replace("-", "."))}</th>`).join("");
    const rows = r.rows.map(g => `<tr><td class="l"><b>${esc(g.key)}</b></td><td class="c">${g.people}</td>${r.months.map(m => `<td class="num${g.mm[m] ? "" : " muted"}">${g.mm[m] ? g.mm[m] : "-"}</td>`).join("")}
      <td class="num"><b>${g.total_mm}</b></td><td class="bar-cell"><span class="bar"><i class="k-internal_labor" style="width:${Math.min(100, g.mm_share)}%"></i></span></td><td class="num">${g.mm_share.toFixed(1)}%</td><td class="num">${won(g.cost)}</td><td class="num">${g.share.toFixed(1)}%</td></tr>`).join("");
    const kpi = (lbl, v) => `<div class="mini-kpi"><span>${esc(lbl)}</span><b>${v}</b></div>`;
    return head("chart", "투입 현황", "인력 계획을 영역 · 시스템 · 직급 · 제안 직무 · 역할 · 업체별로 모아 봅니다", "")
      + `<div class="stat-bar"><label class="field"><span>모아 볼 기준</span>${seg("by", byList, by)}</label><label class="field"><span>인력</span>${seg("scope", STAT_SCOPE, scope)}</label><label class="field"><span>기간 (분할계약)</span>${phaseSel}</label></div>
      <div class="mini-kpis">${kpi("기간", `${r.months.length}개월${ph ? ` · ${esc(ph.title)}` : ""}`)}${kpi("투입 인원", `${r.total.people}명`)}${kpi("총 MM", r.total.total_mm)}${kpi("인건비", `${won(r.total.cost)}원`)}</div>
      <div class="tbl-wrap"><table><thead><tr><th class="l">${esc(label)}</th><th>인원</th>${mh}<th>총 MM</th><th class="bar-cell">MM 비중</th><th></th><th>인건비(원)</th><th>금액 비중</th></tr></thead>
      <tbody>${rows || empty(r.months.length + 7, "이 조건에 맞는 인력이 없습니다.")}
      ${rows ? `<tr class="total"><td class="l">합계</td><td class="c">${r.total.people}</td>${r.months.map(m => `<td class="num">${r.total.mm[m]}</td>`).join("")}<td class="num">${r.total.total_mm}</td><td></td><td class="num">100%</td><td class="num">${won(r.total.cost)}</td><td class="num">100%</td></tr>` : ""}</tbody></table></div>`;
  }

  // ---------------- 부서별 손익 (부서 2개 이상) ----------------
  //  r: Calc.deptSummary 결과, revPath: 부서 매출 입력 경로 앞부분 (프로젝트: dept_revenue / 회차: phases.i.dept_revenue)
  function deptTable(p, c, r, revPath, revenueOf, totalLabel) {
    const t = r.total;
    const row = (g, editable) => `<tr${editable ? "" : ' class="total"'}><td class="l"><b>${esc(g.dept)}</b></td>
      <td class="num">${editable && g.dept !== "(미지정)" ? inp(`${revPath}.${g.dept}`, revenueOf[g.dept], { money: true, cls: "w-lg", ph: "0", label: `${g.dept} 매출` }) : won(g.revenue)}</td>
      <td class="num">${won(g.internal_labor)}</td><td class="num">${won(g.external_labor)}</td><td class="num">${won(g.overhead)}</td><td class="num">${won(g.expense)}</td><td class="num">${won(g.common)}</td>
      <td class="num"><b>${won(g.cost)}</b></td><td class="num">${won(g.profit)}</td><td class="num">${g.revenue > 0 ? `<span class="tag ${g.margin_rate >= c.summary.hurdle_rate ? "ok" : "bad"}">${g.margin_rate.toFixed(1)}%</span>` : "-"}</td>
      <td class="num">${g.mm.toFixed(2)}</td><td class="c">${g.people}</td></tr>`;
    return `<div class="tbl-wrap"><table class="dept-tbl"><thead><tr><th class="l">부서</th><th>매출 (입력)</th><th>내부 인건비</th><th>외주 인건비</th><th>인비례 경비</th><th>직접 경비</th><th>공통경비 배분</th><th>원가 합</th><th>이익</th><th>이익률</th><th>MM</th><th>인원</th></tr></thead>
      <tbody>${r.rows.map(g => row(g, true)).join("")}${row(Object.assign({ dept: totalLabel || "합계 (통합)" }, t), false)}</tbody></table></div>`;
  }
  const deptWarn = (r, what, target) => {
    const w = [];
    if (r.revenue_gap) w.push(`부서 매출 합 ${won(r.total.revenue)}원이 ${what} ${won(target)}원보다 ${won(Math.abs(r.revenue_gap))}원 ${r.revenue_gap > 0 ? "적습니다" : "많습니다"}.`);
    if (r.unassigned) w.push("부서를 고르지 않은 인력이 있습니다 — (미지정) 줄을 확인하세요.");
    if (r.overridden) w.push("이 회차는 별도 원가를 넣어서, 계산한 부서 비율대로 별도 원가에 맞췄습니다.");
    return w.length ? `<p class="note-warn">${w.map(esc).join("<br>")}</p>` : "";
  };
  // [부서별] 탭: 기간에서 '전체' 또는 분할계약 회차를 고름
  function dept(p, c, ctx) {
    const pi = Number((ctx.deptPhase) || 0), ph = c.phases[pi - 1];
    const r = ph ? PC.Calc.phaseDeptSummary(p, c)[pi - 1] : PC.Calc.deptSummary(p, c), t = r.total;
    const phaseSel = `<select data-change-act="setDeptPhase" data-key="phase" aria-label="기간"><option value="0">전체 기간 (${c.months[0] || ""} ~ ${c.months[c.months.length - 1] || ""})</option>${c.phases.map((x, i) => `<option value="${i + 1}"${pi === i + 1 ? " selected" : ""}>${esc(x.title)} (${x.start_ym} ~ ${x.end_ym})</option>`).join("")}</select>`;
    const kpi = (a, b) => `<div class="mini-kpi"><span>${esc(a)}</span><b>${b}</b></div>`;
    return head("building", "부서별 손익", "인건비 · 인비례는 인력의 부서, 경비는 항목의 부서(공통은 기준에 따라 나눔), 매출은 부서마다 입력", "")
      + `<div class="stat-bar"><label class="field"><span>기간 (분할계약)</span>${phaseSel}</label><label class="field"><span>공통경비 나누는 기준</span>${sel("settings.dept_common_basis", p.settings.dept_common_basis || "cost", PC.Calc.DEPT_BASIS, { label: "공통경비 기준" })}</label></div>
      <div class="mini-kpis">${kpi(ph ? "회차 부서 매출 합" : "부서 매출 합", `${won(t.revenue)}원`)}${kpi(ph ? "회차 원가" : "통합 원가", `${won(t.cost)}원`)}${kpi("공통경비", `${won(r.common)}원`)}${kpi("이익률", `${t.margin_rate.toFixed(1)}%`)}</div>
      ${deptWarn(r, ph ? "회차 매출" : "계약 매출", ph ? ph.revenue : c.summary.revenue)}
      ${ph ? deptTable(p, c, r, `phases.${pi - 1}.dept_revenue`, p.phases[pi - 1].dept_revenue || {}, "회차 합계") : deptTable(p, c, r, "dept_revenue", p.dept_revenue || {})}`;
  }
  // 분할계약 탭 아래: 회차마다 부서별 손익 (부서 2개 이상일 때)
  function phaseDepts(p, c) {
    if (!M.multiDept(p) || !c.phases.length) return "";
    return `<section>${subHead("회차별 부서 손익", "회차 매출을 부서마다 한 번 넣으면 됩니다 · 기간은 회차 기간 그대로")}`
      + PC.Calc.phaseDeptSummary(p, c).map((r, i) => `<div class="phase-dept"><h4>${esc(r.phase.title)} <span class="muted">${esc(r.phase.start_ym)} ~ ${esc(r.phase.end_ym)} · 회차 매출 ${won(r.phase.revenue)}원</span></h4>
        ${deptWarn(r, "회차 매출", r.phase.revenue)}${deptTable(p, c, r, `phases.${i}.dept_revenue`, p.phases[i].dept_revenue || {}, "회차 합계")}</div>`).join("") + `</section>`;
  }

  // ---------------- 버전 (기준선) ----------------
  //  ctx.cmp = { a: 비교 기준(버전 번호), b: 비교 대상(0 = 지금 화면) }
  function versions(p, c, ctx) {
    const list = p.versions || [];
    const calcOf = v => PC.Calc.calculateProjectCost(M.normalize(PC.clone(v.data), p.settings), p.settings);
    const vc = list.map(calcOf);
    const rows = list.map((v, i) => { const s = vc[i].summary; return `<tr><td class="c">${v.no}</td>
      <td class="l">${inp(`versions.${i}.name`, v.name, { cls: "w-md", label: "버전 이름" })}</td><td class="nowrap">${esc(v.saved_at || "")}</td>
      <td class="l">${inp(`versions.${i}.memo`, v.memo, { cls: "w-xl", ph: "메모", label: "메모" })}</td>
      <td class="num">${won(s.revenue)}</td><td class="num">${won(s.total_cost)}</td><td class="num">${s.margin_rate.toFixed(1)}%</td>
      <td class="nowrap">${btn("덮어쓰기", "overwriteVersion", [i], "sm", "이 버전을 지금 자료로 덮어쓰고 저장")}${btn("불러오기", "restoreVersion", [i], "sm", "이 버전의 자료를 지금 화면으로 (저장해야 파일에 반영)")}${btn("", "removeVersion", [i], "ghost-danger icon-only", "이 버전 지우기", "trash")}</td></tr>`; }).join("");
    const listHtml = `<div class="tbl-wrap"><table class="ver-list"><thead><tr><th>번호</th><th class="l">이름</th><th>저장 일시</th><th class="l">메모</th><th>계약 매출</th><th>총 원가</th><th>이익률</th><th></th></tr></thead>
      <tbody>${rows || empty(8, "저장한 버전이 없습니다. 머리줄의 [버전 저장]으로 지금 자료를 이름(초기 계획 · 실계약 · 수정계약 …)과 메모를 붙여 남기세요.")}</tbody></table></div>`;
    let cmpHtml = "";
    if (list.length) {
      const cmp = ctx.cmp || {}, ai = list.some(v => v.no === cmp.a) ? cmp.a : list[list.length - 1].no;
      const bi = cmp.b === 0 || list.some(v => v.no === cmp.b) ? (cmp.b || 0) : 0;
      const pick = no => (no ? vc[list.findIndex(v => v.no === no)] : c);
      const nameOf = no => (no ? `${no}. ${(list.find(v => v.no === no) || {}).name || ""}` : "지금 화면");
      const opt = (cur, withNow) => (withNow ? [`<option value="0"${cur === 0 ? " selected" : ""}>지금 화면 (저장 안 한 변경 포함)</option>`] : []).concat(list.map(v => `<option value="${v.no}"${v.no === cur ? " selected" : ""}>${v.no}. ${esc(v.name)} (${esc(String(v.saved_at || "").slice(0, 10))})</option>`)).join("");
      const r = PC.Calc.compareVersions(pick(ai), pick(bi));
      const fmt = (v, u) => (u === "text" ? esc(v) : u === "pct" ? `${num(v).toFixed(1)}%` : u === "mm" ? num(v).toFixed(2) : won(v));
      const dcls = (d, key) => (!d ? "muted" : (key === "revenue" || key === "profit" || key === "margin_rate") === d > 0 ? "up-good" : "up-bad");
      const dtxt = (d, u) => (d === null ? "" : !d ? "-" : `${d > 0 ? "+" : ""}${fmt(d, u)}`);
      const sumRows = r.summary.map(x => `<tr${x.key === "total_cost" || x.key === "profit" ? ' class="em"' : ""}><td class="l">${esc(x.label)}</td><td class="num">${fmt(x.a, x.unit)}</td><td class="num">${fmt(x.b, x.unit)}</td><td class="num ${dcls(x.diff, x.key)}">${dtxt(x.diff, x.unit)}</td></tr>`).join("");
      const staffRows = r.staff.map(x => `<tr><td class="c">${esc(x.kind)}</td><td class="l">${esc(x.name)}</td><td>${esc(x.area)}</td><td>${esc(x.grade)}</td>
        <td class="num">${x.a_mm.toFixed(2)}</td><td class="num">${x.b_mm.toFixed(2)}</td><td class="num ${x.d_mm ? (x.d_mm > 0 ? "up-bad" : "up-good") : "muted"}">${x.d_mm ? `${x.d_mm > 0 ? "+" : ""}${x.d_mm.toFixed(2)}` : "-"}</td>
        <td class="num">${won(x.a_cost)}</td><td class="num">${won(x.b_cost)}</td><td class="num ${x.d_cost ? (x.d_cost > 0 ? "up-bad" : "up-good") : "muted"}">${x.d_cost ? `${x.d_cost > 0 ? "+" : ""}${won(x.d_cost)}` : "-"}</td></tr>`).join("");
      cmpHtml = `${subHead("버전 비교", "기준과 비교 대상을 고르면 금액 · MM 의 차이를 보여 줍니다")}
        <div class="stat-bar"><label class="field"><span>기준</span><select data-change-act="setCompare" data-key="a" aria-label="비교 기준">${opt(ai, false)}</select></label>
        <label class="field"><span>비교 대상</span><select data-change-act="setCompare" data-key="b" aria-label="비교 대상">${opt(bi, true)}</select></label></div>
        <div class="ver-cmp"><div class="tbl-wrap"><table><thead><tr><th class="l">항목</th><th>${esc(nameOf(ai))}</th><th>${esc(nameOf(bi))}</th><th>차이</th></tr></thead><tbody>${sumRows}</tbody></table></div>
        <div class="tbl-wrap"><table><thead><tr><th>구분</th><th class="l">성명</th><th>영역</th><th>등급</th><th>MM (기준)</th><th>MM (대상)</th><th>차이</th><th>인건비 (기준)</th><th>인건비 (대상)</th><th>차이</th></tr></thead>
        <tbody>${staffRows || empty(10, "인력이 없습니다.")}</tbody></table></div></div>`;
    }
    return head("flag", "버전", "초기 계획 · 실계약 · 수정계약처럼 이름을 붙여 남긴 자료 — 프로젝트 파일 안에 함께 저장됩니다",
      btn("버전 저장", "saveVersion", [], "primary sm", "지금 자료를 이름 · 메모를 붙여 남기고 저장", "flag"))
      + listHtml + cmpHtml;
  }

  const TABS = [
    { id: "info", label: "프로젝트 개요", icon: "info", view: info },
    { id: "internal", label: "내부 인력", icon: "users", view: internal, count: p => p.internal_labor.length },
    { id: "external", label: "외주 인력", icon: "partner", view: external, count: p => p.external_labor.length },
    { id: "stats", label: "투입 현황", icon: "chart", view: stats },
    { id: "dept", label: "부서별", icon: "building", view: dept, show: p => M.multiDept(p) },
    { id: "expense", label: "경비", icon: "wallet", view: expense, count: p => p.general_expenses.length + p.custom_costs.length },
    { id: "phase", label: "분할계약", icon: "split", view: phase, count: p => p.phases.length },
    { id: "estimate", label: "견적서", icon: "receipt", view: estimate, count: p => p.estimate.items.length },
    { id: "versions", label: "버전", icon: "flag", view: versions, count: p => (p.versions || []).length },
    { id: "settings", label: "기준정보", icon: "settings", view: settings }
  ];

  PC.Views = { TABS };
})(window);
