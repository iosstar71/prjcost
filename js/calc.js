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

  // 제안 단가: 기준정보의 제안 직무 × 등급 표
  const proposalRate = (settings, role, level) => {
    const r = ((settings || {}).proposal_roles || []).find(x => x && x.role === role);
    return r && r.rates ? num(r.rates[level]) : 0;
  };
  // 견적: 인건비(제안 직무 · 등급 · 견적 MM × 제안 단가 — 단가를 직접 넣으면 그 값) + 기타 항목(수량 × 단가)
  function calculateEstimate(estimate, settings) {
    let supply = 0;
    const labor = (estimate.labor || []).map(l => {
      const mm = num(l.mm), master = proposalRate(settings, l.role, l.level);
      const unit = isBlank(l.unit_price) ? master : num(l.unit_price);
      const amount = pyRound(mm * unit, 2);
      supply = pyRound(supply + amount, 2);
      return Object.assign({}, l, { mm, master_rate: master, applied_rate: unit, amount, missing: !l.role || !l.level });
    });
    const laborTotal = supply;
    const items = (estimate.items || []).map(it => {
      const qty = num(it.qty), unit = num(it.unit_price);
      const amount = pyRound(qty * unit, 2);
      supply = pyRound(supply + amount, 2);
      return Object.assign({}, it, { qty, unit_price: unit, amount });
    });
    // 할인(합계 × 할인율 %) → 절사(남은 금액을 절사 단위 아래로 버림, 예: 100원 단위면 9,999 → 9,900) → 공급가액
    const subtotal = supply;
    const rate = Math.min(100, Math.max(0, num(estimate.discount_rate)));
    const discount = pyRound((subtotal * rate) / 100, 2);
    const afterDiscount = pyRound(subtotal - discount, 2);
    const unit = num(estimate.round_unit) > 1 ? num(estimate.round_unit) : 0;
    supply = unit ? Math.floor(afterDiscount / unit) * unit : afterDiscount;
    const truncation = pyRound(afterDiscount - supply, 2);
    return {
      estimate_no: String(estimate.estimate_no || ""),
      estimate_date: String(estimate.estimate_date || ""),
      note: String(estimate.note || ""),
      applied_to_revenue: !!estimate.applied_to_revenue,
      labor, labor_total: laborTotal, items, items_total: pyRound(subtotal - laborTotal, 2), subtotal, discount_rate: rate, discount, round_unit: unit || 1, truncation,
      supply_total: supply   // 견적에는 부가세를 넣지 않음
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
    // OT비용: 프로젝트 전체 금액 하나 — 내부 인건비(노무비)에 넣고 달에는 고르게
    const otMonthly = distributeMonthly({ dist_mode: "equal_total", total_input: num(project.ot_cost) }, months);
    months.forEach(m => { internalMonthly[m] = pyRound(internalMonthly[m] + otMonthly[m], 2); });
    const tOt = pyRound(sumValues(otMonthly), 2);
    // 재료비: 구분 · 내용 · 금액 — 총 원가에 넣고 달에는 고르게
    const materialMonthly = zero();
    const materialDetails = (project.materials || []).map(x => {
      const mMap = distributeMonthly({ dist_mode: "equal_total", total_input: num(x.amount) }, months);
      months.forEach(m => { materialMonthly[m] = pyRound(materialMonthly[m] + mMap[m], 2); });
      return Object.assign({}, x, { calc_monthly: mMap, total_cost: pyRound(sumValues(mMap), 2) });
    });
    const tMaterial = pyRound(sumValues(materialMonthly), 2);

    const tInternal = pyRound(sumValues(internalMonthly), 2), tExternal = pyRound(sumValues(externalMonthly), 2);
    const tGeneral = pyRound(sumValues(generalMonthly), 2), tIntOvh = pyRound(sumValues(intOvh), 2);
    const tExtOvh = pyRound(sumValues(extOvh), 2), tCustom = pyRound(sumValues(customMonthly), 2);
    const totalCost = pyRound(tInternal + tExternal + tMaterial + tGeneral + tIntOvh + tExtOvh + tCustom, 2);
    const revenue = pyRound(num(project.contract_revenue), 2);
    const profit = pyRound(revenue - totalCost, 2);
    const margin = revenue > 0 ? pyRound((profit / revenue) * 100.0, 2) : 0;
    const hurdle = pyRound(num(project.target_hurdle_rate), 2);
    const monthCost = m => (internalMonthly[m] || 0) + (externalMonthly[m] || 0) + (materialMonthly[m] || 0) + (generalMonthly[m] || 0) + (intOvh[m] || 0) + (extOvh[m] || 0) + (customMonthly[m] || 0);

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
      { key: "internal_labor", label: "내부 인건비", amount: tInternal, share: share(tInternal), items: tOt ? [{ name: "OT비용", amount: tOt, share: share(tOt) }] : [] },
      { key: "external_labor", label: "외주 인건비", amount: tExternal, share: share(tExternal), items: [] },
      ...(materialDetails.length ? [{ key: "material", label: "재료비", amount: tMaterial, share: share(tMaterial), items: (() => {
        const acc = new Map();
        materialDetails.forEach(x => { const n = String(x.category || x.content || "").trim() || "(구분 없음)"; acc.set(n, pyRound((acc.get(n) || 0) + x.total_cost, 2)); });
        return [...acc].map(([name, amount]) => ({ name, amount, share: share(amount) }));
      })() }] : []),
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
      const pExp = sumOf(m => (materialMonthly[m] || 0) + (generalMonthly[m] || 0) + (intOvh[m] || 0) + (extOvh[m] || 0) + (customMonthly[m] || 0));   // 경비 열 = 재료비 + 경비
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
        breakdown: { internal_labor: tInternal, ot_cost: tOt, external_labor: tExternal, material: tMaterial, general_expense: tGeneral, internal_overhead: tIntOvh, external_overhead: tExtOvh, custom_cost: tCustom }
      },
      details: {
        internal_labor: internalDetails, external_labor: externalDetails, general_expenses: generalDetails,
        internal_overhead_monthly: intOvh, external_overhead_monthly: extOvh, mm_overheads: mmOverheads, custom_costs: customDetails,
        ot_monthly: otMonthly, materials: materialDetails, material_monthly: materialMonthly
      },
      phases,
      estimate: calculateEstimate(project.estimate || {}, settings),
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
    kind: r => (r.__kind === "internal" ? "내부" : "외주"),
    dept: r => String(r.dept || "").trim() || "(미지정)",
    system: r => String(r.system || "").trim() || "(미지정)",
    proposal: r => (r.prop_role ? `${r.prop_role} ${r.prop_level || ""}`.trim() : "(미지정)")
  };
  // 인력 계획 → 견적 인건비 줄: 제안 직무 · 등급별 MM 합 (missing: 제안 직무 · 등급이 없는 인력)
  function estimateFromPlan(calc) {
    const groups = new Map(), missing = [];
    calc.details.internal_labor.concat(calc.details.external_labor).forEach(r => {
      if (!num(r.total_mm)) return;
      if (!r.prop_role || !r.prop_level) { missing.push(r.name || "?"); return; }
      const k = `${r.prop_role}|${r.prop_level}`;
      const g = groups.get(k) || { role: r.prop_role, level: r.prop_level, mm: 0, unit_price: "", memo: "" };
      g.mm = pyRound(g.mm + num(r.total_mm), 2);
      groups.set(k, g);
    });
    return { labor: [...groups.values()], missing };
  }
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
    const order = { dept: st.departments, area: st.areas, system: st.systems, proposal: (st.proposal_roles || []).flatMap(r => (st.proposal_levels || []).map(l => `${r.role} ${l}`)), grade: gradeNames(st).concat(st.external_grades || []), vendor: ["내부"].concat(vendorNames(st)), kind: ["내부", "외주"] }[o.by];
    const rank = rankOf(order || []);
    const sorted = sortByRanks(rows, [g => rank(g.key)]);   // 기준정보에 없는 값(역할 등)은 나온 순서
    rows.length = 0; rows.push(...sorted);
    return { months, rows, total: { mm: Object.fromEntries(months.map(m => [m, pyRound(rows.reduce((a, g) => a + g.mm[m], 0), 2)])), total_mm: totalMm, cost: totalCost, people: rows.reduce((a, g) => a + g.people, 0) } };
  }

  // ---------- 부서별 손익 ----------
  //  인건비 · 인비례 경비는 사람의 부서, 일반 · 기타경비는 항목의 부서(비우면 공통 → 기준에 따라 나눔), 매출은 부서별로 직접 입력
  //  부서 합계 = 통합 원가 (나눌 때 생기는 1원 단위 끝수는 가장 큰 부서에)
  const DEPT_BASIS = [{ v: "cost", t: "직접 원가 비율" }, { v: "mm", t: "MM 비율" }, { v: "revenue", t: "매출 비율" }, { v: "equal", t: "똑같이" }];
  //  o: { months: 이 달들만 (분할계약 회차), revenue: 부서별 매출 {부서: 원} (없으면 project.dept_revenue),
  //       override: 회차에 별도 원가를 넣었으면 그 금액 — 계산한 부서 비율대로 맞춤, target_revenue: 비교할 매출 (회차 매출) }
  function deptSummary(project, calc, o = {}) {
    const st = project.settings || {}, depts = [...new Set((st.departments || []).map(d => String(d).trim()).filter(Boolean))];
    const basis = DEPT_BASIS.some(b => b.v === st.dept_common_basis) ? st.dept_common_basis : "cost";
    const UNK = "(미지정)", revIn = o.revenue || project.dept_revenue || {};
    const months = o.months ? calc.months.filter(m => o.months.includes(m)) : calc.months;
    const sumM = f => pyRound(months.reduce((a, m) => a + num(f(m)), 0), 2);
    const rows = new Map();
    const get = k => {
      if (!rows.has(k)) rows.set(k, { dept: k, revenue: k === UNK ? 0 : pyRound(num(revIn[k]), 2), internal_labor: 0, external_labor: 0, overhead: 0, expense: 0, common: 0, mm: 0, people: 0 });
      return rows.get(k);
    };
    depts.forEach(get);
    const keyOf = d => (depts.includes(String(d || "").trim()) ? String(d).trim() : UNK);
    const add = (g, f, v) => { g[f] = pyRound(g[f] + v, 2); };
    // 끝수 맞추기: 나눈 합이 전체와 다르면 가장 큰 몫에
    const settle = (parts, total, f) => {
      const sum = parts.reduce((a, [g, v]) => a + v, 0), diff = pyRound(total - sum, 2);
      parts.forEach(([g, v]) => add(g, f, v));
      if (diff && parts.length) add(parts.reduce((m, x) => (x[1] > m[1] ? x : m))[0], f, diff);
    };
    // 달마다 나눠서 더함 → 회차(기간)별 부서 금액을 더하면 전체 기간과 같음
    const dm = (g, m, v, f) => { const k = f || "d"; g[k] = g[k] || {}; g[k][m] = (g[k][m] || 0) + v; };   // 부서의 달별 직접 원가(d) · MM(mmM)
    [["internal_labor", calc.details.internal_labor], ["external_labor", calc.details.external_labor]].forEach(([f, list]) => list.forEach(r => {
      const g = get(keyOf(r.dept));
      let mm = 0;
      months.forEach(m => { const c = r.calc[m] || {}; add(g, f, num(c.cost)); dm(g, m, num(c.cost)); dm(g, m, num(c.mm), "mmM"); if (f === "internal_labor") dm(g, m, num(c.cost), "im"); mm += num(c.mm); });
      add(g, "mm", pyRound(mm, 2));
      if (mm) g.people += 1;
    }));
    (calc.details.mm_overheads || []).forEach(o => {   // 인비례: 계정 · 달마다 그 사람 MM × 기준금액
      const people = o.category === "mm_internal" ? calc.details.internal_labor.filter(r => r.overhead_eligible) : calc.details.external_labor;
      months.forEach(m => {
        const parts = people.map(r => [get(keyOf(r.dept)), pyRound(num((r.calc[m] || {}).mm) * num(o.unit_amount), 2)]).filter(x => x[1]);
        settle(parts, num(o.calc_monthly[m]), "overhead");
        parts.forEach(([g, v]) => dm(g, m, v));
      });
    });
    // OT비용(내부 인건비): 그 달 부서별 내부 인건비 비율로 (없으면 등록된 부서에 똑같이)
    months.forEach(m => {
      const ot = num((calc.details.ot_monthly || {})[m]);
      if (!ot) return;
      const base = depts.map(d => rows.get(d)), w = g => num((g.im || {})[m]);
      const ws = base.reduce((a, g) => a + w(g), 0);
      settle(base.map(g => [g, pyRound(ws > 0 ? (ot * w(g)) / ws : ot / Math.max(1, base.length), 2)]), ot, "internal_labor");
      base.forEach(g => dm(g, m, ws > 0 ? (ot * w(g)) / ws : ot / Math.max(1, base.length)));
    });
    const commonM = {};
    months.forEach(m => { const v = num((calc.details.material_monthly || {})[m]); if (v) commonM[m] = pyRound((commonM[m] || 0) + v, 2); });   // 재료비는 공통
    calc.details.general_expenses.concat(calc.details.custom_costs).forEach(e => {
      const d = String(e.dept || "").trim();
      months.forEach(m => {
        const v = num(e.calc_monthly[m]);
        if (!v) return;
        if (depts.includes(d)) { add(get(d), "expense", v); dm(get(d), m, v); } else commonM[m] = pyRound((commonM[m] || 0) + v, 2);
      });
    });
    // 공통경비 나누기 (등록된 부서에만, 달마다): 그 달의 직접 원가 · MM 비율, 매출 비율, 또는 똑같이
    const real = depts.map(d => rows.get(d));
    Object.entries(commonM).forEach(([m, cm]) => {
      if (!cm || !real.length) return;
      const weight = g => (basis === "mm" ? num((g.mmM || {})[m]) : basis === "revenue" ? g.revenue : basis === "equal" ? 1 : num((g.d || {})[m]));
      let wsum = real.reduce((a, g) => a + weight(g), 0);
      const w = wsum > 0 ? weight : () => 1;
      wsum = wsum > 0 ? wsum : real.length;
      settle(real.map(g => [g, pyRound((cm * w(g)) / wsum, 2)]), cm, "common");
    });
    rows.forEach(g => { delete g.d; delete g.mmM; delete g.im; });
    const out = [...rows.values()].filter(g => g.dept !== UNK || g.internal_labor || g.external_labor || g.overhead || g.expense);
    const PARTS = ["internal_labor", "external_labor", "overhead", "expense", "common"];
    const costOf = g => pyRound(PARTS.reduce((a, f) => a + g[f], 0), 2);
    // 별도 원가: 계산한 금액을 같은 비율로 늘이거나 줄여 별도 원가에 맞춤 (끝수는 가장 큰 부서의 공통경비 칸에)
    const calcCost = pyRound(out.reduce((a, g) => a + costOf(g), 0), 2);
    const overridden = o.override !== undefined && o.override !== null && calcCost > 0 && pyRound(o.override, 2) !== calcCost;
    if (overridden) {
      const k = o.override / calcCost;
      out.forEach(g => PARTS.forEach(f => { g[f] = pyRound(g[f] * k, 2); }));
      const diff = pyRound(o.override - out.reduce((a, g) => a + costOf(g), 0), 2);
      if (diff) { const big = out.reduce((m, g) => (costOf(g) > costOf(m) ? g : m)); add(big, "common", diff); }
    }
    out.forEach(g => {
      g.cost = costOf(g);
      g.profit = pyRound(g.revenue - g.cost, 2);
      g.margin_rate = g.revenue > 0 ? pyRound((g.profit / g.revenue) * 100, 2) : 0;
    });
    const sum = f => pyRound(out.reduce((a, g) => a + g[f], 0), 2);
    const total = { revenue: sum("revenue"), internal_labor: sum("internal_labor"), external_labor: sum("external_labor"), overhead: sum("overhead"), expense: sum("expense"), common: sum("common"), cost: sum("cost"), mm: sum("mm"), people: out.reduce((a, g) => a + g.people, 0) };
    total.profit = pyRound(total.revenue - total.cost, 2);
    total.margin_rate = total.revenue > 0 ? pyRound((total.profit / total.revenue) * 100, 2) : 0;
    const target = o.target_revenue !== undefined ? o.target_revenue : calc.summary.revenue;
    return { rows: out, total, basis, common: sum("common"), months, overridden, revenue_gap: pyRound(target - total.revenue, 2), unassigned: out.some(g => g.dept === UNK) };
  }
  // 분할계약 회차마다 부서별 손익 (기간은 회차 기간 그대로, 매출은 회차 × 부서로 입력: phases[i].dept_revenue)
  function phaseDeptSummary(project, calc) {
    return calc.phases.map((ph, i) => Object.assign({ phase: ph }, deptSummary(project, calc, {
      months: ph.start_ym && ph.end_ym ? monthRange(ph.start_ym, ph.end_ym) : [],
      revenue: ((project.phases || [])[i] || {}).dept_revenue || {},
      override: ph.overridden ? ph.cost : undefined, target_revenue: ph.revenue })));
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
      row("material", "재료비", comp(sa, "material"), comp(sb, "material")),
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

  PC.Calc = { proposalRate, estimateFromPlan, DEPT_BASIS, deptSummary, phaseDeptSummary, orderRows, rankOf, compareVersions, calculateProjectCost, calculateEstimate, resolveInternalRate, raiseMonth, distributeMonthly, staffStats };
  if (typeof module !== "undefined" && module.exports) module.exports = PC.Calc;
})(typeof window !== "undefined" ? window : globalThis);
