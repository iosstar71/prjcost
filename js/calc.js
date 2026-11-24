// =============================================================================
// 원가 계산 — 원본 backend/calculator.py (v4) 를 줄 단위로 옮김. 화면 · 저장과 무관한 순수 계산
//  결과는 원본과 1원 단위까지 같아야 함: tests/run.js 가 tests/golden.json(원본 파이썬 계산 결과)과 비교
//  계산 규칙을 바꾸면 tests/reference/calculator_v4.py 도 같이 고치고 `python tools/make_golden.py` 로 기준을 다시 만든다
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC || require("./util.js");
  const { num, isBlank, pyRound, normalizeYm, monthRange, nowText } = PC;

  // 내부 인력의 그 달 단가: 기준단가(없으면 등급 단가) → 승진월이 지난 승진 중 마지막 것
  function resolveInternalRate(member, ym, settings) {
    const grades = new Map();
    (settings.internal_grades || []).forEach(g => { if (g && g.grade) grades.set(String(g.grade).trim(), num(g.rate)); });
    const gradeName = String(member.grade || "").trim();
    let rate = !isBlank(member.base_rate) ? num(member.base_rate) : num(grades.get(gradeName));
    const promos = [];
    (member.promotions || []).forEach(p => {
      const eff = normalizeYm(p && p.effective_ym, "");
      if (!eff) return;
      const pGrade = String((p && p.grade) || "").trim();
      let pRate;
      if (!isBlank(p.rate)) pRate = num(p.rate);
      else if (pGrade && grades.has(pGrade)) pRate = num(grades.get(pGrade));
      else return;
      promos.push([eff, pRate]);
    });
    promos.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
    promos.forEach(([eff, pRate]) => { if (eff <= ym) rate = pRate; });
    return pyRound(rate, 2);
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
    const intRatio = num(settings.internal_overhead_ratio), extRatio = num(settings.external_overhead_ratio);

    // 외주 업체|등급 → 단가
    const extRate = new Map();
    (settings.vendor_rates || []).forEach(r => {
      const v = String(r.vendor || "").trim(), g = String(r.grade || "").trim();
      if (v && g) extRate.set(`${v}|${g}`, num(r.rate));
    });
    Object.entries(settings.external_rates || {}).forEach(([k, v]) => { if (!extRate.has(k)) extRate.set(k, num(v)); });

    // 1. 내부 인건비
    const internalMonthly = zero(), eligibleMonthly = zero();
    const internalDetails = (project.internal_labor || []).map(mem => {
      const mmMap = mem.mm || {}, eligible = !!mem.overhead_eligible, calc = {};
      let rowTotal = 0, totalMm = 0;
      months.forEach(m => {
        const mm = num(mmMap[m]);
        const rate = resolveInternalRate(mem, m, settings);
        const cost = pyRound(mm * rate, 2);
        calc[m] = { mm, rate, cost };
        internalMonthly[m] = pyRound(internalMonthly[m] + cost, 2);
        if (eligible) eligibleMonthly[m] = pyRound(eligibleMonthly[m] + cost, 2);
        rowTotal = pyRound(rowTotal + cost, 2);
        totalMm = pyRound(totalMm + mm, 4);
      });
      return Object.assign({}, mem, { overhead_eligible: eligible, calc, total_mm: pyRound(totalMm, 2), total_cost: rowTotal });
    });

    // 2. 외주 인건비 (적용단가를 직접 넣으면 그 값, 아니면 업체 · 등급 단가)
    const externalMonthly = zero();
    const externalDetails = (project.external_labor || []).map(ext => {
      const key = `${String(ext.vendor || "").trim()}|${String(ext.grade || "").trim()}`;
      const defaultRate = extRate.has(key) ? num(extRate.get(key)) : num(ext.base_rate);
      const hasCustom = !isBlank(ext.custom_rate);
      const applied = hasCustom ? pyRound(num(ext.custom_rate), 2) : pyRound(defaultRate, 2);
      const mmMap = ext.mm || {}, calc = {};
      let rowTotal = 0, totalMm = 0;
      months.forEach(m => {
        const mm = num(mmMap[m]);
        const cost = pyRound(mm * applied, 2);
        calc[m] = { mm, rate: applied, cost };
        externalMonthly[m] = pyRound(externalMonthly[m] + cost, 2);
        rowTotal = pyRound(rowTotal + cost, 2);
        totalMm = pyRound(totalMm + mm, 4);
      });
      return Object.assign({}, ext, { default_rate: pyRound(defaultRate, 2), applied_rate: applied, is_overridden: hasCustom, calc, total_mm: pyRound(totalMm, 2), total_cost: rowTotal });
    });

    // 3. 일반경비 · 인비례경비(내부: 대상 인력만, 외주: 전체) · 기타비용
    const expenseRows = (list, monthly) => (list || []).map(exp => {
      const mMap = distributeMonthly(exp, months);
      months.forEach(m => { monthly[m] = pyRound(monthly[m] + mMap[m], 2); });
      return Object.assign({}, exp, { calc_monthly: mMap, total_cost: pyRound(sumValues(mMap), 2) });
    });
    const generalMonthly = zero();
    const generalDetails = expenseRows(project.general_expenses, generalMonthly);
    const intOvh = Object.fromEntries(months.map(m => [m, pyRound(eligibleMonthly[m] * intRatio, 2)]));
    const extOvh = Object.fromEntries(months.map(m => [m, pyRound(externalMonthly[m] * extRatio, 2)]));
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
        hurdle_gap: pyRound(margin - hurdle, 2), hurdle_met: margin >= hurdle,
        breakdown: { internal_labor: tInternal, external_labor: tExternal, general_expense: tGeneral, internal_overhead: tIntOvh, external_overhead: tExtOvh, custom_cost: tCustom }
      },
      details: {
        internal_labor: internalDetails, external_labor: externalDetails, general_expenses: generalDetails,
        internal_overhead_monthly: intOvh, external_overhead_monthly: extOvh, custom_costs: customDetails
      },
      phases,
      estimate: calculateEstimate(project.estimate || {}),
      computed_at: nowText()
    };
  }

  PC.Calc = { calculateProjectCost, calculateEstimate, resolveInternalRate, distributeMonthly };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Calc;
})(typeof window !== "undefined" ? window : globalThis);
