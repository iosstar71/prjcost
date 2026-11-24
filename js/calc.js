// =============================================================================
// 원가 계산 — v4 calculator.py 를 옮기고, 인비례 경비를 계정별 MM × 기준금액으로 바꿈. 화면 · 저장과 무관한 순수 계산
//  회귀 시험: tests/run.js 가 tests/golden.json(저장해 둔 결과)과 비교. 규칙을 일부러 바꾸면 `node tools/make_golden.js` 로 기준을 다시 만든다
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const { num, isBlank, pyRound, normalizeYm, monthRange, nowText } = PC;

  // 내부 인력의 그 달 단가: 기준단가(없으면 등급 단가) → 적용월이 지난 연봉 인상 · 승진 중 마지막 것 (같은 달이면 승진)
  function resolveInternalRate(member, ym, settings) {
    const grades = new Map();
    (settings.internal_grades || []).forEach(g => { if (g && g.grade) grades.set(String(g.grade).trim(), num(g.rate)); });
    const gradeName = String(member.grade || "").trim();
    let rate = !isBlank(member.base_rate) ? num(member.base_rate) : num(grades.get(gradeName));
    rateSteps(member.rate_changes).concat(rateSteps(member.promotions, grades))
      .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))   // 같은 달은 순서 유지 → 승진이 뒤
      .forEach(([eff, pRate]) => { if (eff <= ym) rate = pRate; });
    return pyRound(rate, 2);
  }

  // 외주 인력의 그 달 단가: 적용단가(없으면 업체 · 등급 단가) → 적용월이 지난 단가 변동 중 마지막 것 (협상으로 바뀐 단가)
  function resolveExternalRate(ext, ym, baseRate) {
    let rate = baseRate;
    rateSteps(ext.rate_changes).forEach(([eff, r]) => { if (eff <= ym) rate = r; });
    return pyRound(rate, 2);
  }
  // 단가 변동 목록 → [[YYYY-MM, 단가]] (적용월 순). 적용월 · 단가가 없는 줄은 빼고, 등급만 있으면 등급표 단가
  function rateSteps(list, grades) {
    const out = [];
    (list || []).forEach(p => {
      const eff = normalizeYm(p && p.effective_ym, "");
      if (!eff) return;
      const g = String((p && p.grade) || "").trim();
      if (!isBlank(p.rate)) out.push([eff, num(p.rate)]);
      else if (grades && g && grades.has(g)) out.push([eff, num(grades.get(g))]);
    });
    return out.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
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

    // 외주 업체|등급 → 단가
    const extRate = new Map();
    (settings.vendor_rates || []).forEach(r => {
      const v = String(r.vendor || "").trim(), g = String(r.grade || "").trim();
      if (v && g) extRate.set(`${v}|${g}`, num(r.rate));
    });
    Object.entries(settings.external_rates || {}).forEach(([k, v]) => { if (!extRate.has(k)) extRate.set(k, num(v)); });

    // 1. 내부 인건비
    const internalMonthly = zero(), eligibleMm = zero();
    const internalDetails = (project.internal_labor || []).map(mem => {
      const mmMap = mem.mm || {}, eligible = !!mem.overhead_eligible, calc = {};
      let rowTotal = 0, totalMm = 0;
      months.forEach(m => {
        const mm = num(mmMap[m]);
        const rate = resolveInternalRate(mem, m, settings);
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
        const rate = resolveExternalRate(ext, m, applied);
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
    const byName = (rows, key) => {
      const acc = new Map();
      rows.forEach(r => { const n = String(r[key] || "").trim() || "(이름 없음)"; acc.set(n, pyRound((acc.get(n) || 0) + r.total_cost, 2)); });
      return [...acc].map(([name, amount]) => ({ name, amount, share: share(amount) }));
    };
    // 경비 = 일반경비 + 인비례 경비 (항목은 일반경비 이름별 · 인비례 계정별), 기타경비는 따로
    const expTotal = pyRound(tGeneral + tIntOvh + tExtOvh, 2);
    const composition = [
      { key: "internal_labor", label: "내부 인건비", amount: tInternal, share: share(tInternal), items: [] },
      { key: "external_labor", label: "외주 인건비", amount: tExternal, share: share(tExternal), items: [] },
      { key: "expense", label: "경비", amount: expTotal, share: share(expTotal), items: byName(generalDetails.concat(mmOverheads.map(o => ({ name: o.account, total_cost: o.total_cost }))), "name") },
      { key: "custom_cost", label: "기타경비", amount: tCustom, share: share(tCustom), items: byName(customDetails, "name") }
    ];

    // 4. 분할계약 (최대 5회차): 별도 원가를 넣지 않으면 그 기간 달의 원가 합
    let cumRev = 0, cumCost = 0;
    const phases = (project.phases || []).slice(0, 5).map((ph, i) => {
      const idx = i + 1;
      const ps = normalizeYm(ph.start_ym, ""), pe = normalizeYm(ph.end_ym, "");
      const phMonths = ps && pe ? monthRange(ps, pe) : [];
      const rev = pyRound(num(ph.revenue), 2);
      const cost = !isBlank(ph.override_cost) ? pyRound(num(ph.override_cost), 2) : pyRound(phMonths.reduce((a, m) => a + monthCost(m), 0), 2);
      cumRev = pyRound(cumRev + rev, 2);
      cumCost = pyRound(cumCost + cost, 2);
      return {
        phase_no: idx, title: String(ph.title || `${idx}차 계약`), start_ym: ps, end_ym: pe,
        revenue: rev, cost, profit: pyRound(rev - cost, 2), margin_rate: rev > 0 ? pyRound(((rev - cost) / rev) * 100.0, 2) : 0,
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
  function staffStats(calc, o = {}) {
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
    rows.sort((a, b) => b.total_mm - a.total_mm || a.key.localeCompare(b.key, "ko"));
    return { months, rows, total: { mm: Object.fromEntries(months.map(m => [m, pyRound(rows.reduce((a, g) => a + g.mm[m], 0), 2)])), total_mm: totalMm, cost: totalCost, people: rows.reduce((a, g) => a + g.people, 0) } };
  }

  PC.Calc = { calculateProjectCost, calculateEstimate, resolveInternalRate, resolveExternalRate, distributeMonthly, staffStats };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Calc;
})(typeof window !== "undefined" ? window : globalThis);
