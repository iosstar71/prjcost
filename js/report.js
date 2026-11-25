// =============================================================================
// 원가 분석 보고서 (머리줄 [보고서]) — 회사 사업검토서와 같은 꼴: 맑은 고딕 · 회색 제목 칸(D9D9D9) · 가는 테두리 · 로마자 소제목
//  01_통합보고   프로젝트 개요 · 수익성 요약 · 원가 구성(세부 항목) · 분할계약 · 인력 투입 · 부서별 손익 (A4 세로 한 장 너비)
//  02_항목별상세  내부 · 외주 인력의 달별 MM · 인건비, 경비의 달별 금액, 달별 원가 합계 (A4 가로)
//  03_투입현황   부서 · 담당 영역 · 등급 · 업체별 달별 MM · 인건비
//  04_부서별손익  부서가 2개 이상일 때 — 전체 기간과 분할계약 회차별
//  값은 모두 calc.js 의 계산 결과를 그대로 씀 (여기서 새로 계산하지 않음 — 비율 표시를 위한 나눗셈만)
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const { num, pyRound } = PC;

  const ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ"];
  const FMT = { won: '#,##0;-#,##0;"-"', mm: '0.00;-0.00;"-"', pct: '0.0%;-0.0%;"-"', pct2: "0.00%", int: '0;-0;"-"' };
  const GRAY = "D9D9D9", LIGHT = "F2F2F2";
  const ext = (a, b) => Object.assign({}, a, b, a.font || (b && b.font) ? { font: Object.assign({}, a.font, b && b.font) } : {});
  const base = { border: "thin", v: "center" };
  const S = {
    title: { font: { b: 1, sz: 16 }, h: "center", v: "center" },
    subtitle: { font: { sz: 11, color: "404040" }, h: "center", v: "center" },
    meta: { font: { sz: 9, color: "7F7F7F" }, h: "right", v: "center" },
    note: { font: { sz: 9, color: "7F7F7F" }, h: "left", v: "center" },
    sec: { font: { b: 1, sz: 11 }, v: "center" },
    th: ext(base, { fill: GRAY, h: "center", wrap: 1 }),
    label: ext(base, { fill: GRAY, h: "center", wrap: 1 }),
    text: ext(base, { h: "left", indent: 1 }),
    center: ext(base, { h: "center" }),
    wrapTop: ext(base, { h: "left", v: "top", wrap: 1, indent: 1 }),
    won: ext(base, { h: "right", fmt: FMT.won }),
    mm: ext(base, { h: "right", fmt: FMT.mm }),
    pct: ext(base, { h: "right", fmt: FMT.pct }),
    pct2: ext(base, { h: "right", fmt: FMT.pct2 }),
    int: ext(base, { h: "center", fmt: FMT.int })
  };
  // 줄 꾸밈: 소계(옅은 회색 · 진하게), 합계(회색 · 진하게), 좋음 · 나쁨(글자색)
  const sub = st => ext(st, { fill: LIGHT, font: { b: 1 } });
  const tot = st => ext(st, { fill: GRAY, font: { b: 1 } });
  const good = (st, ok) => ext(st, { font: { b: 1, color: ok ? "1E7B34" : "C00000" } });
  const ratio = (a, b) => (b ? a / b : 0);
  const dot = s => String(s || "").slice(0, 10).replace(/-/g, ".");
  const lastDay = ym => { const [y, m] = ym.split("-").map(Number); return `${ym}-${String(new Date(y, m, 0).getDate()).padStart(2, "0")}`; };
  const period = (p, months) => {
    const a = String(p.start_date || ""), b = String(p.end_date || "");
    return `${dot(a.length === 7 ? `${a}-01` : a)} ~ ${dot(b.length === 7 ? lastDay(b) : b)} (${months.length}개월)`;
  };
  // 글 길이에 맞춘 줄 높이 (한글은 한 글자에 두 칸)
  const textHeight = (text, width) => {
    const lines = String(text || "").split(/\r?\n/).reduce((n, l) => n + Math.max(1, Math.ceil([...l].reduce((w, ch) => w + (ch.charCodeAt(0) > 255 ? 2 : 1), 0) / Math.max(10, width - 4))), 0);
    return Math.max(22, Math.min(300, lines * 15 + 8));
  };

  // 칸 놓기 도구 (줄 · 열 1부터, B = 2)
  function Canvas() {
    const rows = [], merges = [], heights = {};
    const col = c => PC.Xlsx.colName(c - 1);
    const put = (r, c, v, s) => { const row = rows[r - 1] || (rows[r - 1] = { cells: [] }); row.cells[c - 1] = { v: v === undefined || v === null ? "" : v, s }; };
    return {
      put,
      // 여러 칸을 합쳐 값 하나 (테두리가 끊기지 않게 모든 칸에 같은 꾸밈)
      box(r1, c1, r2, c2, v, s) {
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) put(r, c, r === r1 && c === c1 ? v : "", s);
        if (r2 > r1 || c2 > c1) merges.push(`${col(c1)}${r1}:${col(c2)}${r2}`);
      },
      ht(r, h) { heights[r] = h; },
      sheet(name, o) {
        const n = Math.max(rows.length, ...Object.keys(heights).map(Number));
        const out = [];
        for (let i = 0; i < n; i++) out.push({ cells: (rows[i] && rows[i].cells) || [], ht: heights[i + 1] });
        return Object.assign({ name, rows: out, merges, gridLines: false, rowHeight: 20 }, o);
      }
    };
  }
  // 열 묶음으로 표 한 줄: spans = [[첫 열, 끝 열] …], vals = [[값, 꾸밈] …]
  const line = (C, r, spans, vals) => spans.forEach(([a, b], i) => C.box(r, a, r, b, vals[i][0], vals[i][1]));
  const header = (C, r, spans, titles) => line(C, r, spans, titles.map(t => [t, S.th]));
  // 제목 세 줄 (제목 · 프로젝트 · 작성일)
  function heading(C, last, title, p) {
    C.box(1, 2, 1, last, title, S.title); C.ht(1, 34);
    C.box(2, 2, 2, last, `${p.project_name || ""}${p.project_id ? `  (${p.project_id})` : ""}`, S.subtitle);
    C.box(3, 2, 3, last, `작성일 ${dot(new Date().toISOString())} · 금액 단위: 원 (VAT 별도)`, S.meta);
  }
  const sec = (C, r, i, t) => { C.put(r, 2, `${ROMAN[i]}. ${t}`, S.sec); C.ht(r, 24); };

  // ---------------- 01_통합보고 ----------------
  function summarySheet(p, c) {
    const C = Canvas(), s = c.summary, b = s.breakdown, est = c.estimate, st = p.settings || {};
    const rev = s.revenue, cost = s.total_cost;
    const depts = (st.departments || []).filter(x => String(x).trim()), multi = PC.Model.multiDept(p);
    heading(C, 10, "프로젝트 원가 분석 보고서", p);
    let r = 5, k = 0;

    // Ⅰ. 프로젝트 개요: 이름 B · 값 C:E · 이름 F · 값 G:J
    sec(C, r++, k++, "프로젝트 개요");
    const pair = (a, av, bl, bv, avs, bvs) => { C.put(r, 2, a, S.label); C.box(r, 3, r, 5, av, avs || S.text); C.put(r, 6, bl, S.label); C.box(r, 7, r, 10, bv, bvs || S.text); r++; };
    const wide = (a, v, s2, h) => { C.put(r, 2, a, S.label); C.box(r, 3, r, 10, v, s2 || S.text); if (h) C.ht(r, h); r++; };
    pair("프로젝트 ID", p.project_id || "", "고객사", p.client_name || "-");
    wide("프로젝트명", p.project_name || "");
    pair("프로젝트 PM", p.pm_name || "-", "영업대표", p.sales_rep || "-");
    pair("수행 기간", period(p, c.months), "프로젝트 Type", p.project_type || "-");
    pair("계약 매출", rev, "견적 공급가", est.supply_total || "", ext(S.won, { h: "left", indent: 1 }), ext(S.won, { h: "left", indent: 1 }));
    pair("목표 허들", num(p.target_hurdle_rate) / 100, depts.length ? "부서" : "분할계약", depts.length ? depts.join(" · ") : `${c.phases.length}회차`, ext(S.pct2, { h: "left", indent: 1 }));
    wide("프로젝트 설명", p.description || "", S.wrapTop, textHeight(p.description, 8 * 13));
    r++;

    // Ⅱ. 수익성 요약: 구분 B:C · 금액 D:E · 매출 대비 F · 비고 G:J
    sec(C, r++, k++, "수익성 요약");
    const SP = [[2, 3], [4, 5], [6, 6], [7, 10]];
    header(C, r++, SP, ["구분", "금액 · 비율", "매출 대비", "비고"]);
    const gap = ratio(s.profit, rev) - num(p.target_hurdle_rate) / 100;
    [
      ["계약 매출", rev, S.won, ratio(rev, rev), "VAT 별도"],
      ["총 원가", cost, S.won, ratio(cost, rev), `인건비 ${Math.round(ratio(b.internal_labor + b.external_labor, cost) * 100)}% · 경비 ${Math.round(ratio(cost - b.internal_labor - b.external_labor, cost) * 100)}%`],
      ["예상 이익 (매출마진)", s.profit, sub(S.won), ratio(s.profit, rev), ""],
      ["이익률 (수주마진율)", ratio(s.profit, rev), sub(S.pct2), "", "(매출 − 총 원가) ÷ 매출"],
      ["목표 허들 (H/R)", num(p.target_hurdle_rate) / 100, S.pct2, "", ""],
      ["허들 대비 (Gap)", gap, good(S.pct2, gap >= 0), "", rev > 0 ? (gap >= 0 ? "목표 달성" : "목표 미달") : "매출 미입력"]
    ].forEach(([a, v, vs, rr, note]) => line(C, r++, SP, [[a, S.center], [v, vs], [rr, S.pct], [note, S.text]]));
    r++;

    // Ⅲ. 원가 구성: 구분 B:C(세로 합침) · 세부 항목 D:F · 금액 G:H · 원가 비중 I · 매출 대비 J
    sec(C, r++, k++, "원가 구성");
    header(C, r++, [[2, 3], [4, 6], [7, 8], [9, 9], [10, 10]], ["구분", "세부 항목", "금액", "원가 비중", "매출 대비"]);
    const byGroup = (scope, by) => PC.Calc.staffStats(c, { by, scope, settings: st }).rows.map(g => ({ name: `${g.key}  (${g.people}명 · ${g.total_mm} MM)`, amount: g.cost }));
    s.composition.forEach(g => {
      const items = g.key === "internal_labor" ? byGroup("internal", "grade") : g.key === "external_labor" ? byGroup("external", "vendor") : g.items;
      const r0 = r;
      line(C, r++, [[4, 6], [7, 8], [9, 9], [10, 10]], [["소계", sub(S.center)], [g.amount, sub(S.won)], [ratio(g.amount, cost), sub(S.pct)], [ratio(g.amount, rev), sub(S.pct)]]);
      items.forEach(it => line(C, r++, [[4, 6], [7, 8], [9, 9], [10, 10]], [[it.name, S.text], [it.amount, S.won], [ratio(it.amount, cost), S.pct], [ratio(it.amount, rev), S.pct]]));
      C.box(r0, 2, r - 1, 3, g.label, sub(S.center));
    });
    line(C, r++, [[2, 6], [7, 8], [9, 9], [10, 10]], [["총 원가", tot(S.center)], [cost, tot(S.won)], [cost ? 1 : 0, tot(S.pct)], [ratio(cost, rev), tot(S.pct)]]);
    r++;

    // Ⅳ. 분할계약: 회차 B:C · 기간 D · 매출 E · 원가 F · 내부 G · 외주 H · 경비 I · 이익률 J
    if (c.phases.length) {
      sec(C, r++, k++, "분할계약 현황");
      const PS = [[2, 3], [4, 4], [5, 5], [6, 6], [7, 7], [8, 8], [9, 9], [10, 10]];
      header(C, r++, PS, ["회차", "기간", "회차 매출", "회차 원가", "내부 인건비", "외주 인건비", "경비", "이익률"]);
      c.phases.forEach(ph => line(C, r++, PS, [[ph.title + (ph.overridden ? " *" : ""), S.text], [`${ph.start_ym.replace("-", ".")}~${ph.end_ym.slice(2).replace("-", ".")}`, S.center], [ph.revenue, S.won], [ph.cost, S.won], [ph.internal_labor, S.won], [ph.external_labor, S.won], [ph.expense, S.won], [ph.margin_rate / 100, good(S.pct, ph.margin_rate >= s.hurdle_rate)]]));
      const sum = f => pyRound(c.phases.reduce((a, ph) => a + num(ph[f]), 0), 2);
      line(C, r++, PS, [["합계", tot(S.center)], ["", tot(S.center)], [sum("revenue"), tot(S.won)], [sum("cost"), tot(S.won)], [sum("internal_labor"), tot(S.won)], [sum("external_labor"), tot(S.won)], [sum("expense"), tot(S.won)], [ratio(sum("revenue") - sum("cost"), sum("revenue")), tot(S.pct)]]);
      if (c.phases.some(ph => ph.overridden)) C.box(r++, 2, r - 1, 10, "* 별도 원가를 넣은 회차 — 내부 · 외주 · 경비는 계산한 금액", S.note);
      r++;
    }

    // Ⅴ. 인력 투입: 구분 B:C · 인원 D · 총 MM E · MM 비중 F · 인건비 G:H · 금액 비중 I · 평균 단가 J
    const by = multi ? "dept" : "area";
    const ss = PC.Calc.staffStats(c, { by, settings: st });
    if (ss.rows.length) {
      sec(C, r++, k++, `인력 투입 (${multi ? "부서" : "담당 영역"}별)`);
      const IS = [[2, 3], [4, 4], [5, 5], [6, 6], [7, 8], [9, 9], [10, 10]];
      header(C, r++, IS, [multi ? "부서" : "담당 영역", "인원", "총 MM", "MM 비중", "인건비", "금액 비중", "MM당 단가"]);
      ss.rows.forEach(g => line(C, r++, IS, [[g.key, S.center], [g.people, S.int], [g.total_mm, S.mm], [g.mm_share / 100, S.pct], [g.cost, S.won], [g.share / 100, S.pct], [g.total_mm ? Math.round(g.cost / g.total_mm) : "", S.won]]));
      line(C, r++, IS, [["합계", tot(S.center)], [ss.total.people, tot(S.int)], [ss.total.total_mm, tot(S.mm)], [ss.total.total_mm ? 1 : 0, tot(S.pct)], [ss.total.cost, tot(S.won)], [ss.total.cost ? 1 : 0, tot(S.pct)], [ss.total.total_mm ? Math.round(ss.total.cost / ss.total.total_mm) : "", tot(S.won)]]);
      r++;
    }

    // Ⅵ. 부서별 손익 (부서 2개 이상): 부서 B:C · 매출 D · 내부 E · 외주 F · 경비 G · 원가 H · 이익 I · 이익률 J
    if (multi) {
      const ds = PC.Calc.deptSummary(p, c);
      sec(C, r++, k++, "부서별 손익");
      const DS = [[2, 3], [4, 4], [5, 5], [6, 6], [7, 7], [8, 8], [9, 9], [10, 10]];
      header(C, r++, DS, ["부서", "매출", "내부 인건비", "외주 인건비", "경비", "원가 합", "이익", "이익률"]);
      const dl = (g, t) => line(C, r++, DS, [[g.dept, t ? tot(S.center) : S.center], [g.revenue, t ? tot(S.won) : S.won], [g.internal_labor, t ? tot(S.won) : S.won], [g.external_labor, t ? tot(S.won) : S.won],
        [pyRound(g.overhead + g.expense + g.common, 2), t ? tot(S.won) : S.won], [g.cost, t ? tot(S.won) : S.won], [g.profit, t ? tot(S.won) : S.won], [g.margin_rate / 100, t ? tot(S.pct) : good(S.pct, g.margin_rate >= s.hurdle_rate)]]);
      ds.rows.forEach(g => dl(g));
      dl(Object.assign({ dept: "합계 (통합)" }, ds.total), true);
      C.box(r, 2, r, 10, `경비 = 인비례 + 직접 경비 + 공통경비 배분(${(PC.Calc.DEPT_BASIS.find(x => x.v === ds.basis) || {}).t || ""}) · 회차별은 04_부서별손익 시트`, S.note); r++;
    }
    return C.sheet("01_통합보고", { cols: [2, 14, 13, 13, 13, 14, 13, 13, 13, 13].map((w, i) => (i === 0 ? { w } : w)), freeze: false, page: { fitWidth: true } });
  }

  // ---------------- 02_항목별상세 ----------------
  const MODE = { equal_total: "총액 균등", fill_all: "매달 같은 금액", manual: "월별 직접" };
  function detailSheet(p, c) {
    const C = Canvas(), months = c.months, d = c.details, b = c.summary.breakdown;
    const M0 = 7, cMM = M0 + months.length, cAmt = cMM + 1;   // G 부터 달, 그 뒤 총 MM · 금액
    heading(C, cAmt, "원가 항목별 상세", p);
    let r = 5, k = 0;
    const mHead = months.map(m => `${m.slice(2, 4)}.${m.slice(5)}`);
    const monthSpans = months.map((m, i) => [M0 + i, M0 + i]);

    const labor = (title, rows, kind) => {
      sec(C, r++, k++, title);
      const SP = [[2, 2], [3, 3], [4, 4], [5, 5], [6, 6], ...monthSpans, [cMM, cMM], [cAmt, cAmt]];
      header(C, r++, SP, ["성명", "역할", kind === "internal" ? "담당 영역" : "외주업체", "등급", "월 단가", ...mHead, "총 MM", "인건비"]);
      rows.forEach(x => line(C, r++, SP, [[x.name || "", S.text], [x.role || "", S.text], [(kind === "internal" ? x.area : x.vendor) || "", S.center], [x.grade || "", S.center],
        [months.length ? x.calc[months[0]].rate : "", S.won], ...months.map(m => [x.calc[m].mm, S.mm]), [x.total_mm, S.mm], [x.total_cost, S.won]]));
      const sumM = (f, m) => pyRound(rows.reduce((a, x) => a + num(x.calc[m][f]), 0), 2);
      const total = pyRound(rows.reduce((a, x) => a + num(x.total_cost), 0), 2), tmm = pyRound(rows.reduce((a, x) => a + num(x.total_mm), 0), 2);
      line(C, r++, [[2, 6], ...monthSpans, [cMM, cMM], [cAmt, cAmt]], [["MM 합계", sub(S.center)], ...months.map(m => [sumM("mm", m), sub(S.mm)]), [tmm, sub(S.mm)], ["", sub(S.won)]]);
      line(C, r++, [[2, 6], ...monthSpans, [cMM, cMM], [cAmt, cAmt]], [["월 인건비", tot(S.center)], ...months.map(m => [sumM("cost", m), tot(S.won)]), ["", tot(S.won)], [total, tot(S.won)]]);
      r++;
    };
    labor("내부 인력 투입 · 인건비", d.internal_labor, "internal");
    labor("외주 인력 투입 · 인건비", d.external_labor, "external");

    sec(C, r++, k++, "경비 (일반 · 인비례 · 기타)");
    const ES = [[2, 2], [3, 4], [5, 5], [6, 6], ...monthSpans, [cMM, cAmt]];
    header(C, r++, ES, ["구분", "계정 · 항목", "분배 방식", "기준 금액", ...mHead, "합계"]);
    const exp = [
      ...d.general_expenses.map(e => ["일반경비", e.name, MODE[e.dist_mode] || "", e.dist_mode === "fill_all" ? num(e.unit_input) : e.dist_mode === "manual" ? "" : num(e.total_input), e.calc_monthly, e.total_cost]),
      ...(d.mm_overheads || []).map(o => [o.category === "mm_internal" ? "인비례 (내부)" : "인비례 (외주)", o.account, "MM × 기준금액", o.unit_amount, o.calc_monthly, o.total_cost]),
      ...d.custom_costs.map(e => ["기타경비", e.name, MODE[e.dist_mode] || "", e.dist_mode === "fill_all" ? num(e.unit_input) : e.dist_mode === "manual" ? "" : num(e.total_input), e.calc_monthly, e.total_cost])
    ];
    exp.forEach(([g, name, mode, basev, monthly, total]) => line(C, r++, ES, [[g, S.center], [name || "", S.text], [mode, S.center], [basev, S.won], ...months.map(m => [monthly[m], S.won]), [total, S.won]]));
    const expTotal = pyRound(b.general_expense + b.internal_overhead + b.external_overhead + b.custom_cost, 2);
    const expM = m => pyRound(exp.reduce((a, x) => a + num(x[4][m]), 0), 2);
    line(C, r++, [[2, 6], ...monthSpans, [cMM, cAmt]], [["경비 합계", tot(S.center)], ...months.map(m => [expM(m), tot(S.won)]), [expTotal, tot(S.won)]]);
    r++;

    sec(C, r++, k++, "달별 원가 합계");
    header(C, r++, [[2, 6], ...monthSpans, [cMM, cAmt]], ["구분", ...mHead, "합계"]);
    const lab = (list, m) => pyRound(list.reduce((a, x) => a + num(x.calc[m].cost), 0), 2);
    const monthTotal = m => pyRound(lab(d.internal_labor, m) + lab(d.external_labor, m) + expM(m), 2);
    [["내부 인건비", m => lab(d.internal_labor, m), b.internal_labor], ["외주 인건비", m => lab(d.external_labor, m), b.external_labor], ["경비", expM, expTotal]]
      .forEach(([name, f, t]) => line(C, r++, [[2, 6], ...monthSpans, [cMM, cAmt]], [[name, S.center], ...months.map(m => [f(m), S.won]), [t, S.won]]));
    line(C, r++, [[2, 6], ...monthSpans, [cMM, cAmt]], [["월 원가 합계", tot(S.center)], ...months.map(m => [monthTotal(m), tot(S.won)]), [c.summary.total_cost, tot(S.won)]]);
    let cum = 0;
    line(C, r++, [[2, 6], ...monthSpans, [cMM, cAmt]], [["누적 원가", sub(S.center)], ...months.map(m => { cum = pyRound(cum + monthTotal(m), 2); return [cum, sub(S.won)]; }), ["", sub(S.won)]]);
    return C.sheet("02_항목별상세", { cols: [{ w: 2 }, 13, 12, 12, 12, 12, ...months.map(() => 11.5), 9, 14], freeze: { x: 6 }, page: { landscape: true, fitWidth: true } });
  }

  // ---------------- 03_투입현황 ----------------
  function staffSheet(p, c) {
    const C = Canvas(), months = c.months, st = p.settings || {}, multi = PC.Model.multiDept(p);
    const M0 = 4, cEnd = M0 + months.length;   // B 구분 · C 인원 · D~ 달 · 총 MM · MM 비중 · 인건비 · 금액 비중
    heading(C, cEnd + 3, "인력 투입 현황", p);
    let r = 5, k = 0;
    const mHead = months.map(m => `${m.slice(2, 4)}.${m.slice(5)}`);
    const SP = [[2, 2], [3, 3], ...months.map((m, i) => [M0 + i, M0 + i]), [cEnd, cEnd], [cEnd + 1, cEnd + 1], [cEnd + 2, cEnd + 2], [cEnd + 3, cEnd + 3]];
    (multi ? [["dept", "부서"]] : []).concat([["area", "담당 영역"], ["grade", "등급"], ["vendor", "업체 (내부 / 외주업체)"], ["role", "역할"]]).forEach(([by, title]) => {
      const ss = PC.Calc.staffStats(c, { by, settings: st });
      if (!ss.rows.length) return;
      sec(C, r++, k++, `${title}별`);
      header(C, r++, SP, ["구분", "인원", ...mHead, "총 MM", "MM 비중", "인건비", "금액 비중"]);
      ss.rows.forEach(g => line(C, r++, SP, [[g.key, S.center], [g.people, S.int], ...months.map(m => [g.mm[m], S.mm]), [g.total_mm, S.mm], [g.mm_share / 100, S.pct], [g.cost, S.won], [g.share / 100, S.pct]]));
      line(C, r++, SP, [["합계", tot(S.center)], [ss.total.people, tot(S.int)], ...months.map(m => [ss.total.mm[m], tot(S.mm)]), [ss.total.total_mm, tot(S.mm)], [ss.total.total_mm ? 1 : 0, tot(S.pct)], [ss.total.cost, tot(S.won)], [ss.total.cost ? 1 : 0, tot(S.pct)]]);
      r++;
    });
    return C.sheet("03_투입현황", { cols: [{ w: 2 }, 16, 7, ...months.map(() => 8.5), 9, 9, 14, 9], freeze: { x: 3 }, page: { landscape: true, fitWidth: true } });
  }

  // ---------------- 04_부서별손익 (부서 2개 이상) ----------------
  function deptSheet(p, c) {
    const C = Canvas();
    heading(C, 13, "부서별 손익", p);
    let r = 5, k = 0;
    const SP = [[2, 2], [3, 3], [4, 4], [5, 5], [6, 6], [7, 7], [8, 8], [9, 9], [10, 10], [11, 11], [12, 12], [13, 13]];
    const table = (title, ds, totalLabel, extra) => {
      sec(C, r++, k++, title);
      if (extra) { C.box(r, 2, r, 13, extra, S.note); r++; }
      header(C, r++, SP, ["부서", "매출", "내부 인건비", "외주 인건비", "인비례 경비", "직접 경비", "공통경비 배분", "원가 합", "이익", "이익률", "MM", "인원"]);
      const dl = (g, t) => { const f = x => (t ? tot(x) : x); line(C, r++, SP, [[g.dept, f(S.center)], [g.revenue, f(S.won)], [g.internal_labor, f(S.won)], [g.external_labor, f(S.won)], [g.overhead, f(S.won)], [g.expense, f(S.won)], [g.common, f(S.won)], [g.cost, f(S.won)], [g.profit, f(S.won)], [g.margin_rate / 100, t ? f(S.pct) : good(S.pct, g.margin_rate >= c.summary.hurdle_rate)], [g.mm, f(S.mm)], [g.people, f(S.int)]]); };
      ds.rows.forEach(g => dl(g));
      dl(Object.assign({ dept: totalLabel }, ds.total), true);
      if (ds.revenue_gap) { C.box(r, 2, r, 13, `부서 매출 합이 ${totalLabel === "합계 (통합)" ? "계약" : "회차"} 매출보다 ${Math.abs(ds.revenue_gap).toLocaleString("ko-KR")}원 ${ds.revenue_gap > 0 ? "적음" : "많음"}`, good(S.note, false)); r++; }
      r++;
    };
    const ds = PC.Calc.deptSummary(p, c);
    table("전체 기간", ds, "합계 (통합)", `공통경비 ${Math.round(ds.common).toLocaleString("ko-KR")}원은 ${(PC.Calc.DEPT_BASIS.find(x => x.v === ds.basis) || {}).t || ""}로 나눔`);
    PC.Calc.phaseDeptSummary(p, c).forEach(pr => table(`${pr.phase.title} (${pr.phase.start_ym} ~ ${pr.phase.end_ym})`, pr, "회차 합계", `회차 매출 ${pr.phase.revenue.toLocaleString("ko-KR")}원${pr.overridden ? " · 별도 원가를 부서 비율대로 맞춤" : ""}`));
    return C.sheet("04_부서별손익", { cols: [{ w: 2 }, 14, 14, 13, 13, 12, 12, 13, 14, 14, 9, 8, 7], freeze: false, page: { landscape: true, fitWidth: true } });
  }

  function sheets(project, calc) {
    return [summarySheet(project, calc), detailSheet(project, calc), staffSheet(project, calc), ...(PC.Model.multiDept(project) ? [deptSheet(project, calc)] : [])];
  }

  PC.Report = { sheets };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Report;
})(typeof window !== "undefined" ? window : globalThis);
