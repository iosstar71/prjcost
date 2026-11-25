// =============================================================================
// 원가 계산 — v4 calculator.py 를 옮기고, 인비례 경비를 계정별 MM × 기준금액으로 바꿈. 화면 · 저장과 무관한 순수 계산
//  회귀 시험: tests/run.js 가 tests/golden.json(저장해 둔 결과)과 비교. 규칙을 일부러 바꾸면 `node tools/make_golden.js` 로 기준을 다시 만든다
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const { num, isBlank, pyRound, normalizeYm, monthRange, nowText } = PC;

  // 연봉 인상월: 프로젝트 기간 안에서 기준정보의 인상월(MM)이 처음 오는 달 (기간이 1년 이내라 한 번). 없으면 ""
  function raiseMonth(settings, months) {
    const mm = String((settings && settings.raise_month) || "").padStart(2, "0");
    return (/^\d\d$/.test(mm) && mm !== "00" && months.find(m => m.slice(5) === mm)) || "";
  }
  // 등급의 그 달 단가: 인상월부터는 인상단가(넣은 등급만)
  function gradeRate(g, ym, raiseYm) {
    if (!g) return undefined;
    return raiseYm && ym >= raiseYm && !isBlank(g.raise_rate) ? num(g.raise_rate) : num(g.rate);
  }

  // 내부 인력의 그 달 단가: 기준단가(없으면 등급 단가) → 인상월부터 등급의 인상단가(넣은 등급만)
  function resolveInternalRate(member, ym, settings, raiseYm) {
    const own = (settings.internal_grades || []).find(g => g && String(g.grade || "").trim() === String(member.grade || "").trim());
    let rate = !isBlank(member.base_rate) ? num(member.base_rate) : num(gradeRate(own, ym, raiseYm));
    if (own && raiseYm && ym >= raiseYm && !isBlank(own.raise_rate)) rate = num(own.raise_rate);
    return pyRound(rate, 2);
  }

  // ---------- 기준정보 순서 ----------
  //  목록에 있는 값은 그 순서, 없는 값은 뒤(나온 순서대로), 빈 값 · (미지정)은 맨 뒤
  function rankOf(list) {
    const m = new Map();
    (list || []).forEach((v, i) => { const k = String(v === null || v === undefined ? "" : v).trim(); if (k && !m.has(k)) m.set(k, i); });
    return v => { const k = String(v === null || v === undefined ? "" : v).trim(); return m.has(k) ? m.get(k) : k && k !== "(미지정)" ? 1e6 : 2e6; };
  }
  // 여러 기준으로 차례로 비교 (같으면 원래 순서 — 안정 정렬)
  const sortByRanks = (rows, ranks) => rows.map((r, i) => [r, i]).sort((x, y) => {
    for (const f of ranks) { const d = f(x[0]) - f(y[0]); if (d) return d; }
    return x[1] - y[1];
  }).map(x => x[0]);
  const gradeNames = st => (st.internal_grades || []).map(g => g.grade);
  const vendorNames = st => [...new Set((st.vendor_rates || []).map(v => v.vendor))];
  // 목록(인력 · 경비)을 기준정보 순서로: 내부 = 영역 → 등급, 외주 = 영역 → 업체 → 등급, 경비 = 계정
  function orderRows(list, rows, st) {
    const area = rankOf(st.areas), acc = rankOf((st.expense_accounts || []).map(a => a.account));
    if (list === "internal_labor") { const g = rankOf(gradeNames(st)); return sortByRanks(rows, [r => area(r.area), r => g(r.grade)]); }
    if (list === "external_labor") { const v = rankOf(vendorNames(st)), g = rankOf(st.external_grades); return sortByRanks(rows, [r => area(r.area), r => v(r.vendor), r => g(r.grade)]); }
    return sortByRanks(rows, [r => acc(r.name)]);
  }

  // 경비를 달마다 나눔: 총액 균등 · 매달 같은 금액 · 직접 입력
  function distributeMonthly(item, months) {
    const out = {};
    if (!months.length) return out;
    const mode = item.dist_mode || "equal_total";
    if (mode === "equal_total") {
      const per = pyRound(num(item.total_input) / months.length, 2);
      months.forEach(m => { out[m] = per; });
    } else if (mode === "fill_all") {
      const unit = pyRound(num(item.unit_input), 2);
      months.forEach(m => { out[m] = unit; });
    } else {
      const monthly = item.monthly || {};
      months.forEach(m => { out[m] = pyRound(num(monthly[m]), 2); });
    }
    return out;
  }

  const sumValues = obj => Object.values(obj).reduce((a, b) => a + b, 0);

  function calculateEstimate(estimate) {
    let supply = 0;
    const items = (estimate.items || []).map(it => {
      const qty = num(it.qty), unit = num(it.unit_price);
      const amount = pyRound(qty * unit, 2);
      supply = pyRound(supply + amount, 2);
      return Object.assign({}, it, { qty, unit_price: unit, amount });
    });
    const vat = pyRound(supply * 0.1, 2);
    return {
      estimate_no: String(estimate.estimate_no || ""),
      estimate_date: String(estimate.estimate_date || ""),
      note: String(estimate.note || ""),
      applied_to_revenue: !!estimate.applied_to_revenue,
      items, supply_total: supply, vat, grand_total: pyRound(supply + vat, 2)
    };
  }

  // 프로젝트 전용 기준정보(project.settings)를 먼저 쓰고, 없으면 기본 기준정보로
  function calculateProjectCost(project, fallbackSettings) {
    const settings = (project.settings && Object.keys(project.settings).length ? project.settings : null) || (fallbackSettings && Object.keys(fallbackSettings).length ? fallbackSettings : null) || {};
    const months = monthRange(String(project.start_date || project.start_ym || "2026-01"), String(project.end_date || project.end_ym || "2026-06"));
    const zero = () => Object.fromEntries(months.map(m => [m, 0]));
    // 인비례 경비 계정: 내부(인비례 대상 인력 MM) · 외주(외주 인력 MM) × 계정별 기준금액(원/MM)
    const mmAccounts = (settings.expense_accounts || []).filter(a => a && (a.category === "mm_internal" || a.category === "mm_external"));

    // 외주 업체|등급 → 단가 (+ 협상으로 바뀌는 달 · 인상단가)
    const extRate = new Map(), extRaise = new Map();
    (settings.vendor_rates || []).forEach(r => {
      const v = String(r.vendor || "").trim(), g = String(r.grade || "").trim();
      if (!v || !g) return;
      extRate.set(`${v}|${g}`, num(r.rate));
      const eff = normalizeYm(r.raise_ym, "");
      if (eff && !isBlank(r.raise_rate)) extRaise.set(`${v}|${g}`, [eff, num(r.raise_rate)]);
    });
    const raiseYm = raiseMonth(settings, months);
    Object.entries(settings.external_rates || {}).forEach(([k, v]) => { if (!extRate.has(k)) extRate.set(k, num(v)); });

    // 1. 내부 인건비
    const internalMonthly = zero(), eligibleMm = zero();
    const internalDetails = (project.internal_labor || []).map(mem => {
      const mmMap = mem.mm || {}, eligible = !!mem.overhead_eligible, calc = {};
      let rowTotal = 0, totalMm = 0;
      months.forEach(m => {
        const mm = num(mmMap[m]);
        const rate = resolveInternalRate(mem, m, settings, raiseYm);
        const cost = pyRound(mm * rate, 2);
        calc[m] = { mm, rate, cost };
        internalMonthly[m] = pyRound(internalMonthly[m] + cost, 2);
        if (eligible) eligibleMm[m] = pyRound(eligibleMm[m] + mm, 4);
        rowTotal = pyRound(rowTotal + cost, 2);
        totalMm = pyRound(totalMm + mm, 4);
      });
      return Object.assign({}, mem, { overhead_eligible: eligible, calc, total_mm: pyRound(totalMm, 2), total_cost: rowTotal });
    });

    // 2. 외주 인건비 (적용단가를 직접 넣으면 그 값, 아니면 업체 · 등급 단가)
    const externalMonthly = zero(), externalMm = zero();
    const externalDetails = (project.external_labor || []).map(ext => {
      const key = `${String(ext.vendor || "").trim()}|${String(ext.grade || "").trim()}`;
      const defaultRate = extRate.has(key) ? num(extRate.get(key)) : num(ext.base_rate);
      const hasCustom = !isBlank(ext.custom_rate);
      const applied = hasCustom ? pyRound(num(ext.custom_rate), 2) : pyRound(defaultRate, 2);
      const mmMap = ext.mm || {}, calc = {};
      let rowTotal = 0, totalMm = 0;
      months.forEach(m => {
        const mm = num(mmMap[m]);
        const up = extRaise.get(key), rate = up && m >= up[0] ? up[1] : applied;   // 협상 인상월부터 인상단가
        const cost = pyRound(mm * rate, 2);
        calc[m] = { mm, rate, cost };
        externalMonthly[m] = pyRound(externalMonthly[m] + cost, 2);
        externalMm[m] = pyRound(externalMm[m] + mm, 4);
        rowTotal = pyRound(rowTotal + cost, 2);
        totalMm = pyRound(totalMm + mm, 4);
      });
      return Object.assign({}, ext, { default_rate: pyRound(defaultRate, 2), applied_rate: applied, is_overridden: hasCustom, calc, total_mm: pyRound(totalMm, 2), total_cost: rowTotal });
    });

    // 3. 일반경비 · 인비례경비(계정별 MM × 기준금액) · 기타비용
    const expenseRows = (list, monthly) => (list || []).map(exp => {
      const mMap = distributeMonthly(exp, months);
      months.forEach(m => { monthly[m] = pyRound(monthly[m] + mMap[m], 2); });
      return Object.assign({}, exp, { calc_monthly: mMap, total_cost: pyRound(sumValues(mMap), 2) });
    });
    const generalMonthly = zero();
    const generalDetails = expenseRows(project.general_expenses, generalMonthly);
    const intOvh = zero(), extOvh = zero();
    const mmOverheads = mmAccounts.map(acc => {
      const internal = acc.category === "mm_internal", base = internal ? eligibleMm : externalMm, unit = num(acc.unit_amount);
      const mMap = Object.fromEntries(months.map(m => [m, pyRound(base[m] * unit, 2)]));
      const target = internal ? intOvh : extOvh;
      months.forEach(m => { target[m] = pyRound(target[m] + mMap[m], 2); });
      return { account: String(acc.account || ""), category: acc.category, unit_amount: pyRound(unit, 2),
        mm_monthly: Object.fromEntries(months.map(m => [m, pyRound(base[m], 2)])), calc_monthly: mMap, total_cost: pyRound(sumValues(mMap), 2) };
    });
    const customMonthly = zero();
    const customDetails = expenseRows(project.custom_costs, customMonthly);

    const tInternal = pyRound(sumValues(internalMonthly), 2), tExternal = pyRound(sumValues(externalMonthly), 2);
    const tGeneral = pyRound(sumValues(generalMonthly), 2), tIntOvh = pyRound(sumValues(intOvh), 2);
    const tExtOvh = pyRound(sumValues(extOvh), 2), tCustom = pyRound(sumValues(customMonthly), 2);
    const totalCost = pyRound(tInternal + tExternal + tGeneral + tIntOvh + tExtOvh + tCustom, 2);
    const revenue = pyRound(num(project.contract_revenue), 2);
    const profit = pyRound(revenue - totalCost, 2);
    const margin = revenue > 0 ? pyRound((profit / revenue) * 100.0, 2) : 0;
    const hurdle = pyRound(num(project.target_hurdle_rate), 2);
    const monthCost = m => (internalMonthly[m] || 0) + (externalMonthly[m] || 0) + (generalMonthly[m] || 0) + (intOvh[m] || 0) + (extOvh[m] || 0) + (customMonthly[m] || 0);

    // 원가 구성: 항목별 금액 · 총 원가 대비 비중(%) — 경비는 같은 이름끼리 합침
    const share = v => (totalCost > 0 ? pyRound((v / totalCost) * 100.0, 2) : 0);
    const accRank = rankOf((settings.expense_accounts || []).map(a => a.account));
    const byName = (rows, key) => {   // 같은 이름끼리 합치고 기준정보의 계정 순서로
      const acc = new Map();
      rows.forEach(r => { const n = String(r[key] || "").trim() || "(이름 없음)"; acc.set(n, pyRound((acc.get(n) || 0) + r.total_cost, 2)); });
      return sortByRanks([...acc].map(([name, amount]) => ({ name, amount, share: share(amount) })), [x => accRank(x.name)]);
    };
    // 경비 = 일반경비 + 인비례 경비 (항목은 일반경비 이름별 · 인비례 계정별), 기타경비는 따로
    const expTotal = pyRound(tGeneral + tIntOvh + tExtOvh, 2);
    const composition = [
      { key: "internal_labor", label: "내부 인건비", amount: tInternal, share: share(tInternal), items: [] },
      { key: "external_labor", label: "외주 인건비", amount: tExternal, share: share(tExternal), items: [] },
      { key: "expense", label: "경비", amount: expTotal, share: share(expTotal), items: byName(generalDetails.concat(mmOverheads.map(o => ({ name: o.account, total_cost: o.total_cost }))), "name") },
      { key: "custom_cost", label: "기타경비", amount: tCustom, share: share(tCustom), items: byName(customDetails, "name") }
    ];

    // 4. 분할계약 (최대 5회차): 별도 원가를 넣지 않으면 그 기간 달의 원가 합 — 내부 · 외주 인건비 · 경비(일반 · 인비례 · 기타)로도 나눔
    let cumRev = 0, cumCost = 0;
    const phases = (project.phases || []).slice(0, 5).map((ph, i) => {
      const idx = i + 1;
      const ps = normalizeYm(ph.start_ym, ""), pe = normalizeYm(ph.end_ym, "");
      const phMonths = ps && pe ? monthRange(ps, pe) : [];
      const rev = pyRound(num(ph.revenue), 2);
      const sumOf = f => pyRound(phMonths.reduce((a, m) => a + f(m), 0), 2);
      const pInt = sumOf(m => internalMonthly[m] || 0), pExt = sumOf(m => externalMonthly[m] || 0);
      const pExp = sumOf(m => (generalMonthly[m] || 0) + (intOvh[m] || 0) + (extOvh[m] || 0) + (customMonthly[m] || 0));
      const calcCost = sumOf(monthCost), overridden = !isBlank(ph.override_cost);
      const cost = overridden ? pyRound(num(ph.override_cost), 2) : calcCost;
      cumRev = pyRound(cumRev + rev, 2);
      cumCost = pyRound(cumCost + cost, 2);
      return {
        phase_no: idx, title: String(ph.title || `${idx}차 계약`), start_ym: ps, end_ym: pe,
        revenue: rev, cost, calc_cost: calcCost, overridden, internal_labor: pInt, external_labor: pExt, expense: pExp, profit: pyRound(rev - cost, 2), margin_rate: rev > 0 ? pyRound(((rev - cost) / rev) * 100.0, 2) : 0,
        cum_revenue: cumRev, cum_cost: cumCost, rem_revenue: pyRound(revenue - cumRev, 2), rem_cost: pyRound(totalCost - cumCost, 2)
      };
    });

    return {
      months,
      summary: {
        revenue, total_cost: totalCost, profit, margin_rate: margin, hurdle_rate: hurdle,
        hurdle_gap: pyRound(margin - hurdle, 2), hurdle_met: margin >= hurdle, composition,
        breakdown: { internal_labor: tInternal, external_labor: tExternal, general_expense: tGeneral, internal_overhead: tIntOvh, external_overhead: tExtOvh, custom_cost: tCustom }
      },
      details: {
        internal_labor: internalDetails, external_labor: externalDetails, general_expenses: generalDetails,
        internal_overhead_monthly: intOvh, external_overhead_monthly: extOvh, mm_overheads: mmOverheads, custom_costs: customDetails
      },
      phases,
      estimate: calculateEstimate(project.estimate || {}),
      computed_at: nowText()
    };
  }

  // ---------- 인력 투입 현황 (화면 '투입 현황' · 보고서) ----------
  //  calc: calculateProjectCost 결과, o: { by: area|grade|role|vendor|kind, scope: all|internal|external, months: 집계할 달 (없으면 전체) }
  //  → { months, rows: [{ key, mm{달}, total_mm, cost, share(금액 비중 %), people }], total: {...} }
  const BY = {
    area: r => String(r.area || "").trim() || "(미지정)",
    grade: r => String(r.grade || "").trim() || "(미지정)",
    role: r => String(r.role || "").trim() || "(미지정)",
    vendor: r => (r.__kind === "internal" ? "내부" : String(r.vendor || "").trim() || "(미지정)"),
    kind: r => (r.__kind === "internal" ? "내부" : "외주")
  };
  function staffStats(calc, o = {}) {   // o.settings 가 있으면 기준정보 순서로
    const months = o.months && o.months.length ? calc.months.filter(m => o.months.includes(m)) : calc.months;
    const scope = o.scope || "all", keyOf = BY[o.by] || BY.area;
    const people = [];
    if (scope !== "external") calc.details.internal_labor.forEach(r => people.push(Object.assign({ __kind: "internal" }, r)));
    if (scope !== "internal") calc.details.external_labor.forEach(r => people.push(Object.assign({ __kind: "external" }, r)));
    const groups = new Map();
    people.forEach(r => {
      const k = keyOf(r);
      if (!groups.has(k)) groups.set(k, { key: k, mm: Object.fromEntries(months.map(m => [m, 0])), total_mm: 0, cost: 0, people: 0 });
      const g = groups.get(k);
      let any = false;
      months.forEach(m => {
        const c = r.calc[m] || { mm: 0, cost: 0 };
        g.mm[m] = pyRound(g.mm[m] + c.mm, 2); g.total_mm = pyRound(g.total_mm + c.mm, 2); g.cost = pyRound(g.cost + c.cost, 2);
        if (c.mm) any = true;
      });
      if (any) g.people += 1;
    });
    const rows = [...groups.values()];
    const totalCost = rows.reduce((a, g) => pyRound(a + g.cost, 2), 0), totalMm = rows.reduce((a, g) => pyRound(a + g.total_mm, 2), 0);
    rows.forEach(g => { g.share = totalCost > 0 ? pyRound((g.cost / totalCost) * 100, 2) : 0; g.mm_share = totalMm > 0 ? pyRound((g.total_mm / totalMm) * 100, 2) : 0; });
    const st = o.settings || {};
    const order = { area: st.areas, grade: gradeNames(st).concat(st.external_grades || []), vendor: ["내부"].concat(vendorNames(st)), kind: ["내부", "외주"] }[o.by];
    const rank = rankOf(order || []);
    const sorted = sortByRanks(rows, [g => rank(g.key)]);   // 기준정보에 없는 값(역할 등)은 나온 순서
    rows.length = 0; rows.push(...sorted);
    return { months, rows, total: { mm: Object.fromEntries(months.map(m => [m, pyRound(rows.reduce((a, g) => a + g.mm[m], 0), 2)])), total_mm: totalMm, cost: totalCost, people: rows.reduce((a, g) => a + g.people, 0) } };
  }

  // 두 계산 결과(버전 · 지금) 비교 — 요약 줄과 사람별 MM · 인건비 (이름 + 내부/외주로 맞춤)
  function compareVersions(a, b) {
    const sa = a.summary, sb = b.summary, comp = (s, k) => ((s.composition.find(g => g.key === k) || {}).amount || 0);
    const mmOf = c => pyRound(c.details.internal_labor.concat(c.details.external_labor).reduce((t, r) => t + num(r.total_mm), 0), 2);
    const row = (key, label, va, vb, unit) => ({ key, label, a: va, b: vb, diff: typeof va === "number" ? pyRound(vb - va, 2) : null, unit: unit || "won" });
    const summary = [
      row("period", "기간", `${a.months[0] || ""} ~ ${a.months[a.months.length - 1] || ""} (${a.months.length}개월)`, `${b.months[0] || ""} ~ ${b.months[b.months.length - 1] || ""} (${b.months.length}개월)`, "text"),
      row("revenue", "계약 매출", sa.revenue, sb.revenue),
      row("internal_labor", "내부 인건비", comp(sa, "internal_labor"), comp(sb, "internal_labor")),
      row("external_labor", "외주 인건비", comp(sa, "external_labor"), comp(sb, "external_labor")),
      row("expense", "경비", comp(sa, "expense"), comp(sb, "expense")),
      row("custom_cost", "기타경비", comp(sa, "custom_cost"), comp(sb, "custom_cost")),
      row("total_cost", "총 원가", sa.total_cost, sb.total_cost),
      row("profit", "예상 이익", sa.profit, sb.profit),
      row("margin_rate", "이익률(%)", sa.margin_rate, sb.margin_rate, "pct"),
      row("total_mm", "총 MM", mmOf(a), mmOf(b), "mm")
    ];
    const people = new Map();
    const add = (c, side) => [["내부", c.details.internal_labor], ["외주", c.details.external_labor]].forEach(([kind, rows]) => rows.forEach(r => {
      const k = `${kind}|${String(r.name || "").trim()}`;
      const p = people.get(k) || { kind, name: r.name || "", area: "", grade: "", a_mm: 0, b_mm: 0, a_cost: 0, b_cost: 0 };
      p[`${side}_mm`] = pyRound(p[`${side}_mm`] + num(r.total_mm), 2);
      p[`${side}_cost`] = pyRound(p[`${side}_cost`] + num(r.total_cost), 2);
      if (side === "b" || !p.grade) { p.area = r.area || p.area; p.grade = r.grade || p.grade; }
      people.set(k, p);
    }));
    add(a, "a"); add(b, "b");
    const staff = [...people.values()].map(p => Object.assign(p, { d_mm: pyRound(p.b_mm - p.a_mm, 2), d_cost: pyRound(p.b_cost - p.a_cost, 2) }));
    return { summary, staff };
  }

  PC.Calc = { orderRows, rankOf, compareVersions, calculateProjectCost, calculateEstimate, resolveInternalRate, raiseMonth, distributeMonthly, staffStats };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Calc;
})(typeof window !== "undefined" ? window : globalThis);
