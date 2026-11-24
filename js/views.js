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
    const body = o.rows.map((r, i) => `<tr data-row="${i}"><td class="c">${del(o.list, i)}</td>${o.cols.map(x => x.cell(r, `${o.list}.${i}`, i)).join("")}${o.tail ? o.tail(r) : ""}</tr>${o.after ? o.after(r, i, o.cols.length + 1 + (o.tail ? 1 : 0)) : ""}`).join("");
    const newRow = `<tr class="new-row"><td></td><td class="l"><input type="text" data-new="${o.list}" placeholder="＋ ${esc(o.newLabel)}" aria-label="${esc(o.newLabel)}"></td><td class="l hint-cell" colspan="${o.cols.length - 1 + (o.tail ? 1 : 0)}">값을 넣으면 한 줄이 추가됩니다 · 엑셀에서 복사한 영역을 붙여 넣어도 됩니다</td></tr>`;
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
      ${field("프로젝트 설명", `<textarea rows="4" data-bind="description" data-type="text" placeholder="배경 · 범위 · 특이사항">${esc(p.description)}</textarea>`, { full: true })}
    </div>
    <section class="sub-card comp-card">${subHead("원가 구성", `총 원가 ${won(s.total_cost)}원 대비 비중`)}
      <table class="comp"><tbody>${s.total_cost > 0 ? comp : empty(4, "인력 · 경비를 넣으면 항목별 금액과 비중이 보입니다.")}</tbody>
      <tfoot><tr><td class="l">총 원가</td><td class="num">${won(s.total_cost)}</td><td></td><td class="num">${s.total_cost > 0 ? "100%" : "-"}</td></tr>
      <tr class="${s.revenue > 0 ? (s.hurdle_met ? "ok" : "bad") : ""}"><td class="l">예상 이익 (이익률)</td><td class="num">${won(s.profit)}</td><td></td><td class="num">${s.margin_rate.toFixed(1)}%</td></tr></tfoot></table>
    </section></div>`;
  }

  // ---------------- 인력 표 (내부 · 외주 공통) ----------------
  // 단가 인상(rate_changes — 내부: 연봉 인상 / 외주: 협상)과 내부 승진(promotions)은 따로 등록. 그 달부터 끝까지 새 단가로 계산
  //  표의 칸 = 목록의 첫 줄 (여러 번이면 파일의 자료로 계산되고 칸 아래에 "외 n건")
  const first = (r, f) => (r[f] || [])[0] || {};
  const more = (r, f) => ((r[f] || []).length > 1 ? `<span class="cell-cost">외 ${(r[f] || []).length - 1}건</span>` : "");
  const stepAt = (r, m) => ["rate_changes", "promotions"].some(f => (r[f] || []).some(x => String(x.effective_ym || "").slice(0, 7) === m));
  const monthCol = (f, l, why) => ({ b: `${f}.0.effective_ym`, t: "month", l, th: `<th title="${why}">${l}</th>`, cell: (r, base) => `<td>${inp(`${base}.${f}.0.effective_ym`, first(r, f).effective_ym, { type: "month", label: l })}</td>` });
  const rateCol = (f, l, why, ph) => ({ b: `${f}.0.rate`, t: "blank-num", l, th: `<th title="${why}">${l}</th>`, cell: (r, base) => `<td>${inp(`${base}.${f}.0.rate`, first(r, f).rate, { money: true, kind: "blank-num", cls: "w-md", ph, label: l })}${more(r, f)}</td>` });
  const raiseCols = why => [monthCol("rate_changes", "인상월", `${why} — 이 달부터 인상단가`), rateCol("rate_changes", "인상단가", `${why} — 비우면 인상 없음`, "인상단가")];

  function laborTab(list, p, c, cols, o) {
    const months = c.months, rows = c.details[list];
    const all = cols.concat(monthCols(months, "mm", m => (r, base) => { const st = stepAt(r, m); return `<td${st ? ` class="rate-step" title="${esc(m)}부터 단가 ${won(r.calc[m].rate)}원"` : ""}>${mmInp(`${base}.mm.${m}`, r.calc[m].mm || "", `${m} MM`)}${r.calc[m].cost ? `<span class="cell-cost">${man(r.calc[m].cost)}만</span>` : ""}</td>`; }));
    const foot = rows.length ? `<tr class="total"><td></td><td colspan="${cols.length}" class="l">합계 · ${sumBy(rows, r => r.total_mm).toFixed(2)} MM</td>${months.map(m => `<td class="num">${man(sumBy(rows, r => r.calc[m].cost))}만</td>`).join("")}<td class="num">${won(o.total)}</td></tr>` : "";
    return head(o.icon, o.title, GRID_HINT, btn(o.addLabel, "addRow", [list], "primary sm", "", "plus"))
      + gridTable({ list, rows, cols: all, newLabel: "이름 입력", foot, tailHead: `<th>합계(원)</th>`, after: o.after,
        tail: r => `<td class="num"><b>${won(r.total_cost)}</b><span class="cell-cost">${r.total_mm} MM</span></td>` });
  }
  const mmTarget = (list, r, base, i) => `<td class="c nowrap">${mmInp(`${base}.target_mm`, r.target_mm !== undefined ? r.target_mm : r.total_mm, "총 MM")}<button class="btn sm icon-only" data-act="distribute" data-args="${args(list, i)}" title="총 MM 을 기간에 고르게 나눔" aria-label="배분">${icon("distribute")}</button></td>`;

  function internal(p, c, ctx) {
    const grades = (p.settings.internal_grades || []).map(g => g.grade), areas = p.settings.areas || [];
    return laborTab("internal_labor", p, c, [
      { b: "name", t: "text", l: "성명", th: `<th class="l">성명</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.name`, r.name, { cls: "w-sm", label: "성명" })}</td>` },
      { b: "role", t: "text", l: "역할", th: `<th class="l">역할</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.role`, r.role, { cls: "w-sm", ph: "PM/개발", label: "역할" })}</td>` },
      { b: "area", t: "text", l: "담당 영역", th: `<th>담당 영역</th>`, cell: (r, base) => `<td>${sel(`${base}.area`, r.area || "", [{ v: "", t: "-" }].concat(areas), { label: "담당 영역" })}</td>` },
      { b: "grade", t: "text", l: "등급", th: `<th>등급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", grades, { label: "등급" })}</td>` },
      { b: "base_rate", t: "blank-num", l: "기준단가", th: `<th>기준단가</th>`, cell: (r, base) => `<td>${inp(`${base}.base_rate`, r.base_rate, { money: true, kind: "blank-num", cls: "w-md", label: "기준단가" })}</td>` },
      { b: "overhead_eligible", t: "bool", l: "인비례", th: `<th title="내부 인비례 경비(MM × 기준금액)를 매길 인력">인비례</th>`, cell: (r, base) => `<td class="c">${chk(`${base}.overhead_eligible`, r.overhead_eligible, "인비례")}</td>` },
      ...raiseCols("연봉 인상"),
      monthCol("promotions", "승진월", "승진하는 달 — 이 달부터 승진 등급 · 단가"),
      { b: "promotions.0.grade", t: "text", l: "승진등급", th: `<th title="승진 후 등급 — 승진단가를 비우면 이 등급의 단가">승진등급</th>`, cell: (r, base) => `<td>${sel(`${base}.promotions.0.grade`, first(r, "promotions").grade || "", [{ v: "", t: "-" }].concat(grades), { label: "승진등급" })}</td>` },
      rateCol("promotions", "승진단가", "비우면 승진등급의 단가", "등급 단가"),
      { b: "target_mm", t: "mm", l: "총MM", th: `<th>총 MM</th>`, cell: (r, base, i) => mmTarget("internal_labor", r, base, i) }
    ], { icon: "users", title: "내부 인력 투입 계획", addLabel: "내부 인력", total: c.summary.breakdown.internal_labor });
  }

  function external(p, c, ctx) {
    const vendors = [...new Set((p.settings.vendor_rates || []).map(v => v.vendor))];
    const grades = p.settings.external_grades || [], areas = p.settings.areas || [];
    return laborTab("external_labor", p, c, [
      { b: "name", t: "text", l: "성명", th: `<th class="l">성명/구분</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.name`, r.name, { cls: "w-sm", label: "성명" })}</td>` },
      { b: "role", t: "text", l: "역할", th: `<th class="l">역할</th>`, cell: (r, base) => `<td class="l">${inp(`${base}.role`, r.role, { cls: "w-sm", ph: "개발/디자인", label: "역할" })}</td>` },
      { b: "area", t: "text", l: "담당 영역", th: `<th>담당 영역</th>`, cell: (r, base) => `<td>${sel(`${base}.area`, r.area || "", [{ v: "", t: "-" }].concat(areas), { label: "담당 영역" })}</td>` },
      { b: "vendor", t: "text", l: "외주업체", th: `<th>외주업체</th>`, cell: (r, base) => `<td>${sel(`${base}.vendor`, r.vendor || "", vendors, { label: "외주업체" })}</td>` },
      { b: "grade", t: "text", l: "등급", th: `<th>등급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", grades, { label: "등급" })}</td>` },
      { b: "custom_rate", t: "blank-num", l: "기준단가", th: `<th title="비워 두면 업체 · 등급 단가표의 단가">기준단가</th>`, cell: (r, base) => `<td>${inp(`${base}.custom_rate`, r.custom_rate, { money: true, kind: "blank-num", cls: "w-md", ph: won(r.default_rate), label: "기준단가" })}${r.is_overridden ? `<span class="cell-cost">기준 ${won(r.default_rate)}</span>` : ""}</td>` },
      ...raiseCols("협상으로 단가가 바뀜"),
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

  // ---------------- 기준정보: 작은 탭으로 한 번에 한 목록 (목록 길이가 달라도 화면이 어긋나지 않게) ----------------
  const SETTING_TABS = [
    { id: "grades", label: "내부 등급 · 단가", count: st => (st.internal_grades || []).length },
    { id: "ext", label: "외주 등급", count: st => (st.external_grades || []).length },
    { id: "vendors", label: "외주업체 단가", count: st => (st.vendor_rates || []).length },
    { id: "accounts", label: "경비 계정", count: st => (st.expense_accounts || []).length },
    { id: "areas", label: "담당 영역", count: st => (st.areas || []).length },
    { id: "basics", label: "기본값 · 데모" }
  ];
  function settings(p, c, ctx) {
    const st = p.settings, sub = SETTING_TABS.some(t => t.id === ctx.sub) ? ctx.sub : "grades";
    const subTabs = `<div class="seg" role="tablist">${SETTING_TABS.map(t => `<button class="seg-btn${t.id === sub ? " active" : ""}" data-act="showSettings" data-args="${args(t.id)}">${esc(t.label)}${t.count ? `<span class="cnt">${t.count(st)}</span>` : ""}</button>`).join("")}</div>`;
    const narrow = html => `<div class="narrow">${html}</div>`;
    const text = (b, label, cls) => ({ b, t: "text", l: label, th: `<th class="l">${esc(label)}</th>`, cell: (r, base) => `<td class="l">${inp(b ? `${base}.${b}` : base, b ? r[b] : r, { cls: cls || "w-lg", label })}</td>` });
    const moneyCol = (b, label) => ({ b, t: "num", l: label, th: `<th>${esc(label)}</th>`, cell: (r, base) => `<td>${inp(`${base}.${b}`, r[b], { money: true, cls: "w-md", label })}</td>` });
    let body = "";
    if (sub === "grades") body = narrow(subHead("내부 등급 · 월 기준단가", "등급을 고르면 인력의 기준단가가 이 값으로 들어갑니다 · " + GRID_HINT)
      + gridTable({ list: "settings.internal_grades", rows: st.internal_grades || [], newLabel: "등급 입력", cols: [text("grade", "등급"), moneyCol("rate", "월 기준단가 (원)")] }));
    else if (sub === "ext") body = narrow(subHead("외주 등급", "외주 인력 · 업체 단가에서 고르는 등급 목록")
      + gridTable({ list: "settings.external_grades", rows: st.external_grades || [], newLabel: "등급 입력", cols: [text("", "등급")] }));
    else if (sub === "vendors") body = narrow(subHead("외주업체별 등급 단가", "외주 인력의 적용단가를 비워 두면 이 단가 · " + GRID_HINT)
      + gridTable({ list: "settings.vendor_rates", rows: st.vendor_rates || [], newLabel: "업체 입력", cols: [text("vendor", "업체"),
        { b: "grade", t: "text", l: "등급", th: `<th>등급</th>`, cell: (r, base) => `<td>${sel(`${base}.grade`, r.grade || "", st.external_grades || [], { label: "등급" })}</td>` },
        moneyCol("rate", "월 단가 (원)")] }));
    else if (sub === "accounts") body = subHead("경비 계정", "인비례 계정은 인력 MM × 기준금액(원/MM)으로 자동 계산 — 내부: [인비례]에 체크한 인력, 외주: 외주 인력")
      + gridTable({ list: "settings.expense_accounts", rows: st.expense_accounts || [], newLabel: "계정 입력", cols: [text("account", "계정"),
        { b: "category", t: "cat", l: "구분", th: `<th>구분</th>`, cell: (r, base) => `<td>${sel(`${base}.category`, r.category || "general", M.CATEGORIES, { label: "구분" })}</td>` },
        { b: "unit_amount", t: "num", l: "기준금액", th: `<th>기준금액 (원/MM)</th>`, cell: (r, base) => (M.isMmCategory(r.category) ? `<td>${inp(`${base}.unit_amount`, r.unit_amount, { money: true, cls: "w-md", label: "기준금액" })}</td>` : `<td class="c muted">-</td>`) },
        text("desc", "설명", "w-xl")] });
    else if (sub === "areas") body = narrow(subHead("담당 영역 (시스템)", "인력마다 고르는 담당 영역 — [투입 현황]에서 영역별로 MM · 금액을 봅니다")
      + gridTable({ list: "settings.areas", rows: st.areas || [], newLabel: "영역 입력", cols: [text("", "담당 영역")] }));
    else body = narrow(`<section class="sub-card">${subHead("기본값")}<div class="form-grid compact">${field("내부 기본 인상 · 승진월 (MM)", inp("settings.default_promotion_month", st.default_promotion_month || "03", { ph: "03", kind: "month2" }), { help: "내부 인력의 인상 · 승진 단가만 넣었을 때 쓰는 달" })}</div></section>
      <section class="sub-card">${subHead("기준정보 가져오기 · 내보내기")}<p class="muted">폴더의 기준정보(master/master_rates.xlsx)는 새 프로젝트를 만들 때의 기본값입니다.</p>
        <div class="bar-group">${btn("폴더 기본값으로 바꾸기", "resetSettings", [], "", "이 프로젝트의 기준정보를 폴더의 기본값으로")}${btn("이 기준정보를 새 프로젝트 기본값으로 저장", "saveDefaults", [], "", "master_rates.xlsx 에 저장")}</div></section>
      <section class="sub-card">${subHead("데모")}<p class="muted">화면을 익히거나 시연할 때 — 지금 프로젝트의 내용을 예시 자료로 바꿉니다.</p>${btn("데모 채우기", "loadDemo", [], "", "예시 자료로 채우기")}</section>`);
    return head("settings", "기준정보", "이 프로젝트 전용 — 여기서 고친 값은 이 프로젝트에만 적용됩니다", "") + subTabs + `<div class="stack mt">${body}</div>`;
  }

  // ---------------- 투입 현황: 영역 · 등급 · 역할 · 업체별 MM · 금액 (분할계약 기간만도) ----------------
  const STAT_BY = [{ v: "area", t: "담당 영역" }, { v: "grade", t: "등급" }, { v: "role", t: "역할" }, { v: "vendor", t: "업체 (내부 / 외주업체)" }, { v: "kind", t: "내부 · 외주" }];
  const STAT_SCOPE = [{ v: "all", t: "전체 인력" }, { v: "internal", t: "내부만" }, { v: "external", t: "외주만" }];
  function stats(p, c, ctx) {
    const o = ctx.stats || {}, by = o.by || "area", scope = o.scope || "all", ph = c.phases[o.phase - 1];
    const months = ph ? PC.monthRange(ph.start_ym, ph.end_ym) : null;
    const r = PC.Calc.staffStats(c, { by, scope, months });
    const seg = (key, list, cur) => `<div class="seg">${list.map(x => `<button class="seg-btn${x.v === cur ? " active" : ""}" data-act="setStats" data-args="${args(key, x.v)}">${esc(x.t)}</button>`).join("")}</div>`;
    const phaseSel = `<select data-change-act="setStats" data-key="phase" aria-label="기간"><option value="0">전체 기간 (${c.months[0] || ""} ~ ${c.months[c.months.length - 1] || ""})</option>${c.phases.map((x, i) => `<option value="${i + 1}"${o.phase === i + 1 ? " selected" : ""}>${esc(x.title)} (${x.start_ym} ~ ${x.end_ym})</option>`).join("")}</select>`;
    const label = (STAT_BY.find(x => x.v === by) || STAT_BY[0]).t;
    const mh = r.months.map(m => `<th class="m">${esc(m.slice(2).replace("-", "."))}</th>`).join("");
    const rows = r.rows.map(g => `<tr><td class="l"><b>${esc(g.key)}</b></td><td class="c">${g.people}</td>${r.months.map(m => `<td class="num${g.mm[m] ? "" : " muted"}">${g.mm[m] ? g.mm[m] : "-"}</td>`).join("")}
      <td class="num"><b>${g.total_mm}</b></td><td class="bar-cell"><span class="bar"><i class="k-internal_labor" style="width:${Math.min(100, g.mm_share)}%"></i></span></td><td class="num">${g.mm_share.toFixed(1)}%</td><td class="num">${won(g.cost)}</td><td class="num">${g.share.toFixed(1)}%</td></tr>`).join("");
    const kpi = (lbl, v) => `<div class="mini-kpi"><span>${esc(lbl)}</span><b>${v}</b></div>`;
    return head("chart", "투입 현황", "인력 계획을 담당 영역 · 등급 · 역할 · 업체별로 모아 봅니다", "")
      + `<div class="stat-bar"><label class="field"><span>모아 볼 기준</span>${seg("by", STAT_BY, by)}</label><label class="field"><span>인력</span>${seg("scope", STAT_SCOPE, scope)}</label><label class="field"><span>기간 (분할계약)</span>${phaseSel}</label></div>
      <div class="mini-kpis">${kpi("기간", `${r.months.length}개월${ph ? ` · ${esc(ph.title)}` : ""}`)}${kpi("투입 인원", `${r.total.people}명`)}${kpi("총 MM", r.total.total_mm)}${kpi("인건비", `${won(r.total.cost)}원`)}</div>
      <div class="tbl-wrap"><table><thead><tr><th class="l">${esc(label)}</th><th>인원</th>${mh}<th>총 MM</th><th class="bar-cell">MM 비중</th><th></th><th>인건비(원)</th><th>금액 비중</th></tr></thead>
      <tbody>${rows || empty(r.months.length + 7, "이 조건에 맞는 인력이 없습니다.")}
      ${rows ? `<tr class="total"><td class="l">합계</td><td class="c">${r.total.people}</td>${r.months.map(m => `<td class="num">${r.total.mm[m]}</td>`).join("")}<td class="num">${r.total.total_mm}</td><td></td><td class="num">100%</td><td class="num">${won(r.total.cost)}</td><td class="num">100%</td></tr>` : ""}</tbody></table></div>`;
  }

  const TABS = [
    { id: "info", label: "프로젝트 개요", icon: "info", view: info },
    { id: "internal", label: "내부 인력", icon: "users", view: internal, count: p => p.internal_labor.length },
    { id: "external", label: "외주 인력", icon: "partner", view: external, count: p => p.external_labor.length },
    { id: "stats", label: "투입 현황", icon: "chart", view: stats },
    { id: "expense", label: "경비", icon: "wallet", view: expense, count: p => p.general_expenses.length + p.custom_costs.length },
    { id: "phase", label: "분할계약", icon: "split", view: phase, count: p => p.phases.length },
    { id: "estimate", label: "견적서", icon: "receipt", view: estimate, count: p => p.estimate.items.length },
    { id: "settings", label: "기준정보", icon: "settings", view: settings }
  ];

  PC.Views = { TABS };
})(window);
