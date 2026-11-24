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
    return `<select data-bind="${esc(path)}" data-type="${o.kind || "text"}"${unknown ? ` class="invalid" title="기준정보에 없는 값입니다"` : ""}${o.label ? ` aria-label="${esc(o.label)}"` : ""}>${opts.map(x => `<option value="${esc(x.v)}"${String(x.v) === String(value) ? " selected" : ""}>${esc(x.t)}</option>`).join("")}</select>`;
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
    const body = o.rows.map((r, i) => `<tr data-row="${i}"><td class="c">${del(o.list, i)}</td>${o.cols.map(x => x.cell(r, `${o.list}.${i}`, i)).join("")}${o.tail ? o.tail(r) : ""}</tr>`).join("");
    const newRow = `<tr class="new-row"><td></td><td class="l"><input type="text" data-new="${o.list}" placeholder="＋ ${esc(o.newLabel)}" aria-label="${esc(o.newLabel)}"></td><td class="l hint-cell" colspan="${o.cols.length - 1 + (o.tail ? 1 : 0)}">이름을 넣으면 한 줄이 추가됩니다 · 엑셀에서 복사한 영역을 붙여 넣어도 됩니다</td></tr>`;
    return `<div class="tbl-wrap"><table class="grid" data-list="${o.list}" data-cols="${esc(JSON.stringify(meta))}"><thead><tr><th class="w-del"></th>${o.cols.map(x => x.th).join("")}${o.tailHead || ""}</tr></thead>
      <tbody>${body}${newRow}${o.foot || ""}</tbody></table></div>`;
  }
  const monthCols = (months, path, cellFn) => months.map(m => ({ b: `${path}.${m}`, t: "mm", l: m, fill: true, th: `<th class="m">${esc(m.slice(2).replace("-", "."))}</th>`, cell: cellFn(m) }));

  // ---------------- 프로젝트 정보 + 원가 구성 ----------------
  function info(p, c, ctx) {
    const dateType = p.date_mode === "date" ? "date" : "month";
    const toMode = v => (dateType === "date" ? (String(v || "").length === 7 ? `${v}-01` : v) : String(v || "").slice(0, 7));
    const fileText = ctx.file ? `projects/${ctx.file}` : "저장 전 — 처음 저장할 때 프로젝트 ID 로 파일 이름이 정해집니다";
    const s = c.summary;
    const comp = s.composition.map(g => `
      <tr class="cg"><td class="l"><b>${esc(g.label)}</b></td><td class="num"><b>${won(g.amount)}</b></td><td class="bar-cell"><span class="bar"><i class="k-${g.key}" style="width:${Math.min(100, g.share)}%"></i></span></td><td class="num"><b>${g.share.toFixed(1)}%</b></td></tr>
      ${g.items.map(it => `<tr class="ci"><td class="l">${esc(it.name)}</td><td class="num">${won(it.amount)}</td><td class="bar-cell"><span class="bar thin"><i class="k-${g.key}" style="width:${Math.min(100, it.share)}%"></i></span></td><td class="num">${it.share.toFixed(1)}%</td></tr>`).join("")}`).join("");
    return head("info", "프로젝트 정보", fileText,
      btn("데모 채우기", "loadDemo", [], "sm", "예시 자료로 화면을 채웁니다") + (ctx.file ? btn("프로젝트 지우기", "deleteProject", [], "sm danger", "파일은 backups/deleted 로 옮겨집니다", "trash") : ""))
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
      ${field("프로젝트 설명", `<textarea rows="4" data-bind="description" data-type="text" placeholder="배경 · 범위 · 특이사항">${esc(p.description)}</textarea>`, { full: true })}
    </div>
    <section class="sub-card comp-card">${subHead("원가 구성", `총 원가 ${won(s.total_cost)}원 대비 비중`)}
      <table class="comp"><tbody>${s.total_cost > 0 ? comp : empty(4, "인력 · 경비를 넣으면 항목별 금액과 비중이 보입니다.")}</tbody>
      <tfoot><tr><td class="l">총 원가</td><td class="num">${won(s.total_cost)}</td><td></td><td class="num">${s.total_cost > 0 ? "100%" : "-"}</td></tr>
      <tr class="${s.revenue > 0 ? (s.hurdle_met ? "ok" : "bad") : ""}"><td class="l">예상 이익 (이익률)</td><td class="num">${won(s.profit)}</td><td></td><td class="num">${s.margin_rate.toFixed(1)}%</td></tr></tfoot></table>
    </section></div>`;
  }

  // ---------------- 인력 표 (내부 · 외주 공통) ----------------
  function laborTab(list, p, c, cols, o) {
    const months = c.months, rows = c.details[list];
    const all = cols.concat(monthCols(months, "mm", m => (r, base) => `<td>${mmInp(`${base}.mm.${m}`, r.calc[m].mm || "", `${m} MM`)}${r.calc[m].cost ? `<span class="cell-cost">${man(r.calc[m].cost)}만</span>` : ""}</td>`));
    const foot = rows.length ? `<tr class="total"><td></td><td colspan="${cols.length}" class="l">합계 · ${sumBy(rows, r => r.total_mm).toFixed(2)} MM</td>${months.map(m => `<td class="num">${man(sumBy(rows, r => r.calc[m].cost))}만</td>`).join("")}<td class="num">${won(o.total)}</td></tr>` : "";
    return head(o.icon, o.title, GRID_HINT, btn(o.addLabel, "addRow", [list], "primary sm", "", "plus"))
      + gridTable({ list, rows, cols: all, newLabel: "이름 입력", foot, tailHead: `<th>합계(원)</th>`,
        tail: r => `<td class="num"><b>${won(r.total_cost)}</b><span class="cell-cost">${r.total_mm} MM</span></td>` });
  }
  const mmTarget = (list, r, base, i) => `<td class="c nowrap">${mmInp(`${base}.target_mm`, r.target_mm !== undefined ? r.target_mm : r.total_mm, "총 MM")}<button class="btn sm icon-only" data-act="distribute" data-args="${args(list, i)}" title="총 MM 을 기간에 고르게 나눔" aria-label="배분">${icon("distribute")}</button></td>`;

  function internal(p, c) {
    const grades = (p.settings.internal_grades || []).map(g => g.grade);
    const p0 = r => (r.promotions && r.promotions[0]) || {};
    return laborTab("internal_labor", p, c, [
      { b: "name", t: "text", l: "성명", th: `<th class="l">성명</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.name`, r.name, { cls: "w-sm", label: "성명" })}</td>` },
      { b: "role", t: "text", l: "역할", th: `<th class="l">역할</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.role`, r.role, { cls: "w-sm", ph: "PM/개발", label: "역할" })}</td>` },
      { b: "grade", t: "text", l: "등급", th: `<th>등급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", grades, { label: "등급" })}</td>` },
      { b: "base_rate", t: "blank-num", l: "기준단가", th: `<th>기준단가</th>`, cell: (r, base) => `<td>${inp(`${base}.base_rate`, r.base_rate, { money: true, kind: "blank-num", cls: "w-md", label: "기준단가" })}</td>` },
      { b: "overhead_eligible", t: "bool", l: "인비례", th: `<th title="내부 인비례 경비(MM × 기준금액)를 매길 인력">인비례</th>`, cell: (r, base) => `<td class="c">${chk(`${base}.overhead_eligible`, r.overhead_eligible, "인비례")}</td>` },
      { b: "promotions.0.effective_ym", t: "month", l: "승진월", th: `<th>승진월</th>`, cell: (r, base) => `<td>${inp(`${base}.promotions.0.effective_ym`, p0(r).effective_ym, { type: "month", label: "승진월" })}</td>` },
      { b: "promotions.0.rate", t: "blank-num", l: "승진단가", th: `<th>승진단가</th>`, cell: (r, base) => `<td>${inp(`${base}.promotions.0.rate`, p0(r).rate, { money: true, kind: "blank-num", cls: "w-md", ph: "승진단가", label: "승진단가" })}</td>` },
      { b: "target_mm", t: "mm", l: "총MM", th: `<th>총 MM</th>`, cell: (r, base, i) => mmTarget("internal_labor", r, base, i) }
    ], { icon: "users", title: "내부 인력 투입 계획", addLabel: "내부 인력", total: c.summary.breakdown.internal_labor });
  }

  function external(p, c) {
    const vendors = [...new Set((p.settings.vendor_rates || []).map(v => v.vendor))];
    const grades = p.settings.external_grades || [];
    return laborTab("external_labor", p, c, [
      { b: "name", t: "text", l: "성명", th: `<th class="l">성명/구분</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.name`, r.name, { cls: "w-sm", label: "성명" })}</td>` },
      { b: "role", t: "text", l: "역할", th: `<th class="l">역할</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.role`, r.role, { cls: "w-sm", ph: "개발/디자인", label: "역할" })}</td>` },
      { b: "vendor", t: "text", l: "외주업체", th: `<th>외주업체</th>`, cell: (r, base) => `<td>${sel(`${base}.vendor`, r.vendor || "", vendors, { label: "외주업체" })}</td>` },
      { b: "grade", t: "text", l: "등급", th: `<th>등급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", grades, { label: "등급" })}</td>` },
      { b: "custom_rate", t: "blank-num", l: "적용단가", th: `<th title="비워 두면 업체 · 등급의 기준단가">적용단가</th>`, cell: (r, base) => `<td>${inp(`${base}.custom_rate`, r.custom_rate, { money: true, kind: "blank-num", cls: "w-md", ph: won(r.default_rate), label: "적용단가" })}${r.is_overridden ? `<span class="cell-cost">기준 ${won(r.default_rate)}</span>` : ""}</td>` },
      { b: "target_mm", t: "mm", l: "총MM", th: `<th>총 MM</th>`, cell: (r, base, i) => mmTarget("external_labor", r, base, i) }
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
      { b: "dist_mode", t: "mode", l: "분배", th: `<th>분배 방식</th>`, cell: (r, base) => `<td>${sel(`${base}.dist_mode`, r.dist_mode || "equal_total", MODES, { label: "분배 방식" })}</td>` },
      { b: "@amount", t: "num", l: "금액", th: `<th title="총액 균등: 전체 금액 / 매달 같은 금액: 한 달 금액">입력 금액</th>`, cell: (r, base) => r.dist_mode === "manual" ? `<td class="c"><span class="tag">달마다</span></td>`
        : `<td>${inp(r.dist_mode === "fill_all" ? `${base}.unit_input` : `${base}.total_input`, r.dist_mode === "fill_all" ? r.unit_input : r.total_input, { money: true, cls: "w-md", label: "입력 금액" })}</td>` },
      ...months.map(m => ({ b: `monthly.${m}`, t: "num", l: m, fill: true, th: `<th class="m">${esc(m.slice(2).replace("-", "."))}</th>`,
        cell: (r, base) => (r.dist_mode === "manual" ? `<td>${inp(`${base}.monthly.${m}`, r.calc_monthly[m] || "", { money: true, cls: "w-sm", ph: "0", label: `${m} 금액` })}</td>` : `<td class="num muted">${won(r.calc_monthly[m])}</td>`) }))
    ];
    const foot = rows.length ? `<tr class="total"><td></td><td colspan="3" class="l">${label} 합계</td>${months.map(m => `<td class="num">${won(sumBy(rows, r => r.calc_monthly[m]))}</td>`).join("")}<td class="num">${won(sumBy(rows, r => r.total_cost))}</td></tr>` : "";
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
    return head("wallet", "경비", `일반 ${won(b.general_expense)} · 인비례 ${won(b.internal_overhead + b.external_overhead)} · 기타 ${won(b.custom_cost)} = ${won(totalAll)}원`, "")
      + `<div class="stack">
      <section>${subHead("일반경비", "금액을 넣고 달에 나누는 방식을 고릅니다 · " + GRID_HINT, btn("일반경비", "addRow", ["general_expenses"], "primary sm", "", "plus"))}${expenseGrid("general_expenses", p, c, "일반경비")}</section>
      <section>${subHead("인비례 경비 (자동)", "계정별로 인력 MM × 기준금액 — 내부는 [인비례]에 체크한 인력, 외주는 전체 외주 인력. 계정 추가 · 삭제는 [기준정보 > 경비 계정]")}
        <div class="tbl-wrap"><table><thead>${mmHead}</thead><tbody>${mmRows || empty(months.length + 4, "인비례 계정이 없습니다. [기준정보 > 경비 계정]에서 구분을 '인비례'로 추가하세요.")}
        ${mmRows ? `<tr class="total"><td colspan="3" class="l">인비례 합계</td>${months.map(m => `<td class="num">${won(num(d.internal_overhead_monthly[m]) + num(d.external_overhead_monthly[m]))}</td>`).join("")}<td class="num">${won(b.internal_overhead + b.external_overhead)}</td></tr>` : ""}</tbody></table></div></section>
      <section>${subHead("기타경비", "", btn("기타경비", "addRow", ["custom_costs"], "sm", "", "plus"))}${expenseGrid("custom_costs", p, c, "기타경비")}</section>
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
      <td class="num">${won(x.revenue)}</td><td class="num">${won(x.cost)}</td><td class="num"><span class="tag ${x.margin_rate >= c.summary.hurdle_rate ? "ok" : "bad"}">${x.margin_rate}%</span></td>
      <td class="num">${won(x.cum_revenue)}</td><td class="num">${won(x.cum_cost)}</td><td class="num">${won(x.rem_revenue)}</td><td class="num">${won(x.rem_cost)}</td></tr>`).join("");
    const revSum = (p.phases || []).reduce((a, x) => a + num(x.revenue), 0);
    const gap = Math.round(num(p.contract_revenue) - revSum);
    const full = (p.phases || []).length >= 5;
    return head("split", "분할계약", gap && p.phases.length ? `회차 매출 합이 계약 매출보다 ${won(Math.abs(gap))}원 ${gap > 0 ? "적습니다" : "많습니다"}` : "최대 5회차",
      full ? `<button class="btn primary sm" disabled title="최대 5회차">${icon("plus")}회차</button>` : btn("회차", "addRow", ["phases"], "primary sm", "", "plus"))
      + `<div class="stack"><div class="tbl-wrap"><table><thead><tr><th class="w-del"></th><th class="l">차수명</th><th>시작월</th><th>종료월</th><th>회차 매출(원)</th><th>별도 원가(원)</th></tr></thead>
        <tbody>${editor || empty(6, "등록된 분할계약이 없습니다.")}</tbody></table></div>
      <div class="tbl-wrap"><table><thead><tr><th class="l">회차</th><th>기간</th><th>회차 매출</th><th>회차 원가</th><th>이익률</th><th>누적 매출</th><th>누적 원가</th><th>잔여 매출</th><th>잔여 원가</th></tr></thead>
        <tbody>${summary || empty(9, "회차를 추가하면 회차 · 누적 · 잔여 손익이 보입니다.")}</tbody></table></div></div>`;
  }

  // ---------------- 견적서 ----------------
  function estimate(p, c) {
    const est = c.estimate;
    const body = est.items.map((it, i) => `<tr>
      <td class="c">${del("estimate.items", i)}</td>
      <td class="l">${inp(`estimate.items.${i}.item_name`, it.item_name, { cls: "w-xl", label: "항목명" })}</td>
      <td class="l">${inp(`estimate.items.${i}.spec`, it.spec, { cls: "w-lg", label: "규격/등급" })}</td>
      <td>${inp(`estimate.items.${i}.qty`, it.qty, { type: "number", step: "0.01", cls: "w-xs", label: "수량" })}</td>
      <td>${inp(`estimate.items.${i}.unit_price`, it.unit_price, { money: true, cls: "w-lg", label: "단가" })}</td>
      <td class="num"><b>${won(it.amount)}</b></td>
      <td class="l">${inp(`estimate.items.${i}.memo`, it.memo, { cls: "w-lg", label: "비고" })}</td></tr>`).join("");
    return head("receipt", "견적서", est.applied_to_revenue ? "공급가액이 계약 매출에 적용됨" : "",
      btn("품목", "addRow", ["estimate.items"], "sm", "", "plus") + btn("공급가액 → 계약 매출", "applyEstimate", [], "primary sm", "견적 공급가액을 계약 매출로 씁니다"))
      + `<div class="form-grid compact">${field("견적번호", inp("estimate.estimate_no", p.estimate.estimate_no, { ph: "EST-2026-001" }))}${field("견적일자", inp("estimate.estimate_date", p.estimate.estimate_date, { type: "date" }))}${field("비고", inp("estimate.note", p.estimate.note, { ph: "조건 · 유효기간" }))}</div>
      <div class="tbl-wrap"><table><thead><tr><th class="w-del"></th><th class="l">항목명</th><th class="l">규격/등급</th><th>수량</th><th>단가(원)</th><th>금액(원)</th><th class="l">비고</th></tr></thead><tbody>
        ${body || empty(7, "등록된 품목이 없습니다.")}
        <tr class="total"><td></td><td colspan="4" class="l">공급가액</td><td class="num">${won(est.supply_total)}</td><td></td></tr>
        <tr class="auto"><td></td><td colspan="4" class="l">부가세 (10%)</td><td class="num">${won(est.vat)}</td><td></td></tr>
        <tr class="total"><td></td><td colspan="4" class="l">견적 총액 (VAT 포함)</td><td class="num">${won(est.grand_total)}</td><td></td></tr>
      </tbody></table></div>`;
  }

  // ---------------- 기준정보 (두 줄 배치 — 카드 높이가 달라도 흐트러지지 않게) ----------------
  function settings(p) {
    const st = p.settings;
    const card = (title, desc, addAct, body) => `<section class="sub-card">${subHead(title, desc, addAct ? btn("추가", "addRow", [addAct], "sm", "", "plus") : "")}${body}</section>`;
    const tbl = (heads, rows, n) => `<div class="tbl-wrap flat"><table class="fit"><thead><tr>${heads}</tr></thead><tbody>${rows || empty(n, "없음")}</tbody></table></div>`;
    const grades = card("내부 등급 · 기준단가", "월 단가 (원/MM)", "settings.internal_grades", tbl(`<th class="l">등급</th><th>월 기준단가 (원)</th><th class="w-del"></th>`,
      (st.internal_grades || []).map((g, i) => `<tr><td class="l">${inp(`settings.internal_grades.${i}.grade`, g.grade, { cls: "full", label: "등급" })}</td><td>${inp(`settings.internal_grades.${i}.rate`, g.rate, { money: true, cls: "full", label: "월 기준단가" })}</td><td class="c">${del("settings.internal_grades", i)}</td></tr>`).join(""), 3));
    const extGrades = card("외주 등급", "", "settings.external_grades", tbl(`<th class="l">등급</th><th class="w-del"></th>`,
      (st.external_grades || []).map((g, i) => `<tr><td class="l">${inp(`settings.external_grades.${i}`, g, { cls: "full", label: "외주 등급" })}</td><td class="c">${del("settings.external_grades", i)}</td></tr>`).join(""), 2));
    const vendors = card("외주업체별 등급 단가", "월 단가 (원/MM)", "settings.vendor_rates", tbl(`<th class="l">업체</th><th>등급</th><th>월 단가 (원)</th><th class="w-del"></th>`,
      (st.vendor_rates || []).map((v, i) => `<tr><td class="l">${inp(`settings.vendor_rates.${i}.vendor`, v.vendor, { cls: "full", label: "업체" })}</td><td>${sel(`settings.vendor_rates.${i}.grade`, v.grade || "", st.external_grades || [], { label: "등급" })}</td><td>${inp(`settings.vendor_rates.${i}.rate`, v.rate, { money: true, cls: "full", label: "월 단가" })}</td><td class="c">${del("settings.vendor_rates", i)}</td></tr>`).join(""), 4));
    const basics = card("기본값", "", "", `<div class="form-grid compact">${field("내부 기본 승진월 (MM)", inp("settings.default_promotion_month", st.default_promotion_month || "03", { ph: "03", kind: "month2" }), { help: "새 내부 인력의 승진월 기본값" })}</div>`);
    const accounts = card("경비 계정", "인비례 계정은 인력 MM × 기준금액(원/MM)으로 자동 계산", "settings.expense_accounts", tbl(`<th class="l">계정</th><th>구분</th><th>기준금액 (원/MM)</th><th class="l">설명</th><th class="w-del"></th>`,
      (st.expense_accounts || []).map((a, i) => `<tr><td class="l">${inp(`settings.expense_accounts.${i}.account`, a.account, { cls: "full", label: "계정" })}</td><td>${sel(`settings.expense_accounts.${i}.category`, a.category || "general", M.CATEGORIES, { label: "구분" })}</td>
        <td>${M.isMmCategory(a.category) ? inp(`settings.expense_accounts.${i}.unit_amount`, a.unit_amount, { money: true, cls: "full", label: "기준금액" }) : `<span class="muted">-</span>`}</td>
        <td class="l">${inp(`settings.expense_accounts.${i}.desc`, a.desc, { cls: "full", label: "설명" })}</td><td class="c">${del("settings.expense_accounts", i)}</td></tr>`).join(""), 5));
    return head("settings", "기준정보", "이 프로젝트 전용 — 여기서 고친 값은 이 프로젝트에만 적용됩니다",
      btn("기본값 다시 불러오기", "resetSettings", [], "sm", "폴더의 기준정보(master_rates.xlsx)로 바꿉니다") + btn("새 프로젝트 기본값으로 저장", "saveDefaults", [], "sm", "master_rates.xlsx 에 저장 — 앞으로 만드는 프로젝트에 적용"))
      + `<div class="settings-cols"><div class="stack">${grades}${extGrades}${basics}</div><div class="stack">${vendors}</div></div><div class="stack mt">${accounts}</div>`;
  }

  const TABS = [
    { id: "info", label: "프로젝트 정보", icon: "info", view: info },
    { id: "internal", label: "내부 인력", icon: "users", view: internal, count: p => p.internal_labor.length },
    { id: "external", label: "외주 인력", icon: "partner", view: external, count: p => p.external_labor.length },
    { id: "expense", label: "경비", icon: "wallet", view: expense, count: p => p.general_expenses.length + p.custom_costs.length },
    { id: "phase", label: "분할계약", icon: "split", view: phase, count: p => p.phases.length },
    { id: "estimate", label: "견적서", icon: "receipt", view: estimate, count: p => p.estimate.items.length },
    { id: "settings", label: "기준정보", icon: "settings", view: settings }
  ];

  PC.Views = { TABS };
})(window);
