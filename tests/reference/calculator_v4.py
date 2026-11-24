from datetime import datetime
from typing import Dict, List, Any


def normalize_ym(val: str, fallback: str) -> str:
    """'YYYY-MM-DD' 또는 'YYYY-MM' 문자열을 'YYYY-MM'으로 정규화"""
    s = str(val or "").strip()
    if len(s) >= 7 and s[4] == "-":
        return s[:7]
    return fallback


def month_range(start_val: str, end_val: str) -> List[str]:
    start_ym = normalize_ym(start_val, "2026-01")
    end_ym = normalize_ym(end_val, "2026-06")
    try:
        s_year, s_month = map(int, start_ym.split("-"))
        e_year, e_month = map(int, end_ym.split("-"))
    except Exception:
        return ["2026-01", "2026-02", "2026-03"]

    if not (1 <= s_month <= 12 and 1 <= e_month <= 12):
        return ["2026-01", "2026-02", "2026-03"]

    if (s_year, s_month) > (e_year, e_month):
        return [start_ym]

    months: List[str] = []
    y, m = s_year, s_month
    while (y, m) <= (e_year, e_month):
        months.append(f"{y:04d}-{m:02d}")
        if m == 12:
            y += 1
            m = 1
        else:
            m += 1
        if len(months) > 120:
            break
    return months


def resolve_internal_rate(member: Dict[str, Any], ym: str, settings: Dict[str, Any]) -> float:
    """내부인력 등급/기준단가 및 승진월·승진등급/단가 반영"""
    internal_grades = {
        str(g.get("grade") or "").strip(): float(g.get("rate") or 0.0)
        for g in (settings.get("internal_grades") or [])
        if g.get("grade")
    }
    grade_name = str(member.get("grade") or "").strip()
    raw_base = member.get("base_rate")
    if raw_base is not None and str(raw_base).strip() != "":
        base_rate = float(raw_base)
    else:
        base_rate = float(internal_grades.get(grade_name, 0.0))

    rate = base_rate
    promotions = member.get("promotions") or []
    valid_promos = []
    for p in promotions:
        eff = normalize_ym(p.get("effective_ym"), "")
        if not eff:
            continue
        p_rate = p.get("rate")
        p_grade = str(p.get("grade") or "").strip()
        if p_rate is not None and str(p_rate).strip() != "":
            resolved_p_rate = float(p_rate)
        elif p_grade and p_grade in internal_grades:
            resolved_p_rate = float(internal_grades[p_grade])
        else:
            continue
        valid_promos.append((eff, resolved_p_rate))

    valid_promos.sort(key=lambda x: x[0])
    for eff, p_rate in valid_promos:
        if eff <= ym:
            rate = p_rate
    return round(rate, 2)


def distribute_expense_monthly(item: Dict[str, Any], months: List[str]) -> Dict[str, float]:
    if not months:
        return {}
    mode = item.get("dist_mode") or "equal_total"
    monthly_in = item.get("monthly") or {}

    if mode == "equal_total":
        total_val = float(item.get("total_input") or 0.0)
        per_m = round(total_val / len(months), 2)
        return {m: per_m for m in months}

    if mode == "fill_all":
        unit_val = float(item.get("unit_input") or 0.0)
        return {m: round(unit_val, 2) for m in months}

    return {m: round(float(monthly_in.get(m) or 0.0), 2) for m in months}


def calculate_estimate(estimate: Dict[str, Any]) -> Dict[str, Any]:
    items = estimate.get("items") or []
    calc_items = []
    supply_total = 0.0
    for it in items:
        qty = float(it.get("qty") or 0.0)
        unit_price = float(it.get("unit_price") or 0.0)
        amount = round(qty * unit_price, 2)
        supply_total = round(supply_total + amount, 2)
        calc_items.append({**it, "qty": qty, "unit_price": unit_price, "amount": amount})
    vat = round(supply_total * 0.1, 2)
    grand_total = round(supply_total + vat, 2)
    return {
        "estimate_no": str(estimate.get("estimate_no") or ""),
        "estimate_date": str(estimate.get("estimate_date") or ""),
        "note": str(estimate.get("note") or ""),
        "applied_to_revenue": bool(estimate.get("applied_to_revenue", False)),
        "items": calc_items,
        "supply_total": supply_total,
        "vat": vat,
        "grand_total": grand_total,
    }


def calculate_project_cost(project: Dict[str, Any], fallback_settings: Dict[str, Any]) -> Dict[str, Any]:
    """프로젝트 전용 기준정보(project['settings'])를 우선 사용하여 원가·수익성·분할계약·견적 산출"""
    settings = project.get("settings") or fallback_settings or {}
    months = month_range(
        str(project.get("start_date") or project.get("start_ym") or "2026-01"),
        str(project.get("end_date") or project.get("end_ym") or "2026-06"),
    )

    int_overhead_ratio = float(settings.get("internal_overhead_ratio") or 0.0)
    ext_overhead_ratio = float(settings.get("external_overhead_ratio") or 0.0)

    # 외주 업체|등급 -> 단가 매핑
    ext_rate_map: Dict[str, float] = {}
    for row in settings.get("vendor_rates") or []:
        v = str(row.get("vendor") or "").strip()
        g = str(row.get("grade") or "").strip()
        if v and g:
            ext_rate_map[f"{v}|{g}"] = float(row.get("rate") or 0.0)
    for k, v in (settings.get("external_rates") or {}).items():
        if k not in ext_rate_map:
            ext_rate_map[k] = float(v or 0.0)

    # 1. 내부 인건비
    internal_monthly = {m: 0.0 for m in months}
    eligible_internal_monthly = {m: 0.0 for m in months}
    internal_details = []

    for mem in project.get("internal_labor") or []:
        mm_map = mem.get("mm") or {}
        eligible = bool(mem.get("overhead_eligible", False))
        calc_map = {}
        row_total = 0.0
        total_mm = 0.0
        for m in months:
            mm = float(mm_map.get(m) or 0.0)
            rate = resolve_internal_rate(mem, m, settings)
            cost = round(mm * rate, 2)
            calc_map[m] = {"mm": mm, "rate": rate, "cost": cost}
            internal_monthly[m] = round(internal_monthly[m] + cost, 2)
            if eligible:
                eligible_internal_monthly[m] = round(eligible_internal_monthly[m] + cost, 2)
            row_total = round(row_total + cost, 2)
            total_mm = round(total_mm + mm, 4)
        internal_details.append({
            **mem,
            "overhead_eligible": eligible,
            "calc": calc_map,
            "total_mm": round(total_mm, 2),
            "total_cost": row_total,
        })

    # 2. 외주 인건비
    external_monthly = {m: 0.0 for m in months}
    external_details = []

    for ext in project.get("external_labor") or []:
        vendor = str(ext.get("vendor") or "").strip()
        grade = str(ext.get("grade") or "").strip()
        key = f"{vendor}|{grade}"
        default_rate = float(ext_rate_map.get(key, ext.get("base_rate") or 0.0))
        custom_val = ext.get("custom_rate")
        has_custom = custom_val is not None and str(custom_val).strip() != ""
        applied_rate = round(float(custom_val), 2) if has_custom else round(default_rate, 2)

        mm_map = ext.get("mm") or {}
        calc_map = {}
        row_total = 0.0
        total_mm = 0.0
        for m in months:
            mm = float(mm_map.get(m) or 0.0)
            cost = round(mm * applied_rate, 2)
            calc_map[m] = {"mm": mm, "rate": applied_rate, "cost": cost}
            external_monthly[m] = round(external_monthly[m] + cost, 2)
            row_total = round(row_total + cost, 2)
            total_mm = round(total_mm + mm, 4)
        external_details.append({
            **ext,
            "default_rate": round(default_rate, 2),
            "applied_rate": applied_rate,
            "is_overridden": has_custom,
            "calc": calc_map,
            "total_mm": round(total_mm, 2),
            "total_cost": row_total,
        })

    # 3. 일반경비, 내부/외주 인비례경비, 기타비용
    general_monthly = {m: 0.0 for m in months}
    general_details = []
    for exp in project.get("general_expenses") or []:
        m_map = distribute_expense_monthly(exp, months)
        for m in months:
            general_monthly[m] = round(general_monthly[m] + m_map[m], 2)
        general_details.append({
            **exp,
            "calc_monthly": m_map,
            "total_cost": round(sum(m_map.values()), 2),
        })

    internal_overhead_monthly = {
        m: round(eligible_internal_monthly[m] * int_overhead_ratio, 2) for m in months
    }
    external_overhead_monthly = {
        m: round(external_monthly[m] * ext_overhead_ratio, 2) for m in months
    }

    custom_monthly = {m: 0.0 for m in months}
    custom_details = []
    for item in project.get("custom_costs") or []:
        m_map = distribute_expense_monthly(item, months)
        for m in months:
            custom_monthly[m] = round(custom_monthly[m] + m_map[m], 2)
        custom_details.append({
            **item,
            "calc_monthly": m_map,
            "total_cost": round(sum(m_map.values()), 2),
        })

    total_internal = round(sum(internal_monthly.values()), 2)
    total_external = round(sum(external_monthly.values()), 2)
    total_general = round(sum(general_monthly.values()), 2)
    total_int_ovh = round(sum(internal_overhead_monthly.values()), 2)
    total_ext_ovh = round(sum(external_overhead_monthly.values()), 2)
    total_custom = round(sum(custom_monthly.values()), 2)

    total_cost = round(
        total_internal + total_external + total_general + total_int_ovh + total_ext_ovh + total_custom,
        2,
    )
    revenue = round(float(project.get("contract_revenue") or 0.0), 2)
    profit = round(revenue - total_cost, 2)
    margin_rate = round((profit / revenue) * 100.0, 2) if revenue > 0 else 0.0
    hurdle_rate = round(float(project.get("target_hurdle_rate") or 0.0), 2)
    hurdle_gap = round(margin_rate - hurdle_rate, 2)

    # 4. 분할계약 (최대 5회차)
    raw_phases = (project.get("phases") or [])[:5]
    cum_revenue = 0.0
    cum_cost = 0.0
    phase_summaries = []

    for idx, ph in enumerate(raw_phases, start=1):
        p_start = normalize_ym(ph.get("start_ym"), "")
        p_end = normalize_ym(ph.get("end_ym"), "")
        ph_months = month_range(p_start, p_end) if (p_start and p_end) else []
        ph_rev = round(float(ph.get("revenue") or 0.0), 2)

        override_val = ph.get("override_cost")
        if override_val is not None and str(override_val).strip() != "":
            ph_cost = round(float(override_val), 2)
        else:
            ph_cost = round(
                sum(
                    internal_monthly.get(m, 0.0)
                    + external_monthly.get(m, 0.0)
                    + general_monthly.get(m, 0.0)
                    + internal_overhead_monthly.get(m, 0.0)
                    + external_overhead_monthly.get(m, 0.0)
                    + custom_monthly.get(m, 0.0)
                    for m in ph_months
                ),
                2,
            )

        cum_revenue = round(cum_revenue + ph_rev, 2)
        cum_cost = round(cum_cost + ph_cost, 2)
        rem_revenue = round(revenue - cum_revenue, 2)
        rem_cost = round(total_cost - cum_cost, 2)
        ph_margin = round(((ph_rev - ph_cost) / ph_rev) * 100.0, 2) if ph_rev > 0 else 0.0

        phase_summaries.append({
            "phase_no": idx,
            "title": str(ph.get("title") or f"{idx}차 계약"),
            "start_ym": p_start,
            "end_ym": p_end,
            "revenue": ph_rev,
            "cost": ph_cost,
            "profit": round(ph_rev - ph_cost, 2),
            "margin_rate": ph_margin,
            "cum_revenue": cum_revenue,
            "cum_cost": cum_cost,
            "rem_revenue": rem_revenue,
            "rem_cost": rem_cost,
        })

    # 5. 견적서 계산
    estimate_calc = calculate_estimate(project.get("estimate") or {})

    return {
        "months": months,
        "summary": {
            "revenue": revenue,
            "total_cost": total_cost,
            "profit": profit,
            "margin_rate": margin_rate,
            "hurdle_rate": hurdle_rate,
            "hurdle_gap": hurdle_gap,
            "hurdle_met": margin_rate >= hurdle_rate,
            "breakdown": {
                "internal_labor": total_internal,
                "external_labor": total_external,
                "general_expense": total_general,
                "internal_overhead": total_int_ovh,
                "external_overhead": total_ext_ovh,
                "custom_cost": total_custom,
            },
        },
        "details": {
            "internal_labor": internal_details,
            "external_labor": external_details,
            "general_expenses": general_details,
            "internal_overhead_monthly": internal_overhead_monthly,
            "external_overhead_monthly": external_overhead_monthly,
            "custom_costs": custom_details,
        },
        "phases": phase_summaries,
        "estimate": estimate_calc,
        "computed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
