# -*- coding: utf-8 -*-
"""원본(v4) 파이썬 계산으로 시험 기준(tests/golden.json)을 만든다

  python tools/make_golden.py
  JS 계산(js/calc.js)이 원본과 1원 단위까지 같은지 tests/run.js 가 이 파일로 비교한다.
  계산 규칙을 일부러 바꿀 때만 다시 만든다 (tests/reference/calculator_v4.py 를 같이 고친 뒤).
"""
import json
import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests" / "reference"))
from calculator_v4 import calculate_project_cost  # noqa: E402

SETTINGS = {
    "internal_overhead_ratio": 0.15, "external_overhead_ratio": 0.05, "default_promotion_month": "03",
    "internal_grades": [{"grade": "수석", "rate": 8500000}, {"grade": "책임", "rate": 7200000}, {"grade": "선임", "rate": 6000000}, {"grade": "전임", "rate": 4800000}],
    "external_grades": ["특급", "고급", "중급", "초급"],
    "vendor_rates": [{"vendor": "파트너A", "grade": "특급", "rate": 9500000}, {"vendor": "파트너A", "grade": "고급", "rate": 7800000},
                     {"vendor": "파트너A", "grade": "중급", "rate": 6200000}, {"vendor": "파트너B", "grade": "특급", "rate": 9200000},
                     {"vendor": "파트너B", "grade": "고급", "rate": 7500000}, {"vendor": "파트너B", "grade": "중급", "rate": 6000000}],
    "expense_accounts": [],
}


def demo():
    return {
        "project_id": "PRJ-2026-01", "start_date": "2026-01", "end_date": "2026-06", "contract_revenue": 240000000, "target_hurdle_rate": 18.0,
        "settings": SETTINGS,
        "internal_labor": [
            {"name": "김철수", "grade": "수석", "base_rate": 8500000, "overhead_eligible": True, "promotions": [{"effective_ym": "2026-03", "rate": 9000000}],
             "mm": {"2026-01": 1, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 1}},
            {"name": "이민지", "grade": "선임", "base_rate": 6000000, "overhead_eligible": False, "promotions": [],
             "mm": {"2026-01": 0.5, "2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1, "2026-06": 0}},
        ],
        "external_labor": [{"name": "외주개발1", "vendor": "파트너A", "grade": "고급", "custom_rate": "",
                            "mm": {"2026-02": 1, "2026-03": 1, "2026-04": 1, "2026-05": 1}}],
        "general_expenses": [{"name": "출장비", "dist_mode": "equal_total", "total_input": 9000000, "unit_input": 0, "monthly": {}}],
        "custom_costs": [{"name": "사전심의·인증비", "dist_mode": "fill_all", "total_input": 0, "unit_input": 400000, "monthly": {}}],
        "phases": [{"title": "1차 선행계약", "start_ym": "2026-01", "end_ym": "2026-03", "revenue": 110000000, "override_cost": ""},
                   {"title": "2차 본계약", "start_ym": "2026-04", "end_ym": "2026-06", "revenue": 130000000, "override_cost": ""}],
        "estimate": {"estimate_no": "EST-1", "items": [{"item_name": "컨설팅", "qty": 1, "unit_price": 95000000}, {"item_name": "개발", "qty": 1.5, "unit_price": 103333333}]},
    }


def months_of(start, end):
    y, m = map(int, start.split("-"))
    ey, em = map(int, end.split("-"))
    out = []
    while (y, m) <= (ey, em):
        out.append(f"{y:04d}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def random_case(rnd, i):
    sy, sm = 2025 + rnd.randint(0, 2), rnd.randint(1, 12)
    span = rnd.randint(0, 20)
    ey, em = sy + (sm - 1 + span) // 12, (sm - 1 + span) % 12 + 1
    start = f"{sy}-{sm:02d}" + (f"-{rnd.randint(1, 28):02d}" if rnd.random() < 0.3 else "")
    end = f"{ey}-{em:02d}" + (f"-{rnd.randint(1, 28):02d}" if rnd.random() < 0.3 else "")
    months = months_of(start[:7], end[:7])
    mm_vals = [0, 0.1, 0.25, 0.3, 0.5, 0.75, 1, 1, 1, 0.333, 0.125, 0.875]
    s = json.loads(json.dumps(SETTINGS))
    s["internal_overhead_ratio"] = rnd.choice([0, 0.15, 0.125, 0.0333, 0.2])
    s["external_overhead_ratio"] = rnd.choice([0, 0.05, 0.075, 0.015])
    if rnd.random() < 0.3:
        s["external_rates"] = {"파트너C|고급": 7123456.5}
    grades = [g["grade"] for g in s["internal_grades"]]
    internal = []
    for k in range(rnd.randint(0, 6)):
        mem = {"name": f"I{k}", "grade": rnd.choice(grades + ["없는등급"]), "overhead_eligible": rnd.random() < 0.6,
               "mm": {m: rnd.choice(mm_vals) for m in months if rnd.random() < 0.85}}
        r = rnd.random()
        if r < 0.5:
            mem["base_rate"] = rnd.choice([6000000, 7250000, 5333333, 8500000.5, ""])
        promos = []
        for _ in range(rnd.randint(0, 2)):
            p = {"effective_ym": rnd.choice(months + ["", "bad"])}
            if rnd.random() < 0.5:
                p["rate"] = rnd.choice([9000000, 7777777, ""])
            if rnd.random() < 0.5:
                p["grade"] = rnd.choice(grades + ["없는등급"])
            promos.append(p)
        mem["promotions"] = promos
        internal.append(mem)
    external = []
    for k in range(rnd.randint(0, 5)):
        external.append({"name": f"E{k}", "vendor": rnd.choice(["파트너A", "파트너B", "파트너C", "없는업체"]), "grade": rnd.choice(["특급", "고급", "중급", "초급"]),
                         "custom_rate": rnd.choice(["", "", 8000000, 6543210.75]), "base_rate": rnd.choice([None, 5000000]),
                         "mm": {m: rnd.choice(mm_vals) for m in months if rnd.random() < 0.8}})

    def exps(n):
        out = []
        for k in range(n):
            mode = rnd.choice(["equal_total", "fill_all", "manual", ""])
            out.append({"name": f"X{k}", "dist_mode": mode, "total_input": rnd.choice([0, 9000000, 1000001, 333333.33]), "unit_input": rnd.choice([0, 400000, 12345.675]),
                        "monthly": {m: rnd.choice([0, 1000, 2500.5, 0.125]) for m in months if rnd.random() < 0.5}})
        return out
    phases = []
    for k in range(rnd.randint(0, 6)):
        a, b = sorted(rnd.sample(months, 2)) if len(months) >= 2 else (months[0], months[0])
        phases.append({"title": rnd.choice(["", f"{k + 1}차"]), "start_ym": rnd.choice([a, ""]) if rnd.random() < 0.1 else a, "end_ym": b,
                       "revenue": rnd.choice([0, 50000000, 123456789]), "override_cost": rnd.choice(["", "", 30000000])})
    est = {"estimate_no": f"E{i}", "applied_to_revenue": rnd.random() < 0.5,
           "items": [{"item_name": f"품목{k}", "qty": rnd.choice([1, 2, 0.5, 1.5, 0.125]), "unit_price": rnd.choice([1000000, 333333.33, 95000000])} for k in range(rnd.randint(0, 4))]}
    return {"project_id": f"R{i}", "start_date": start, "end_date": end, "contract_revenue": rnd.choice([0, 240000000, 150000000.5, 99999999]),
            "target_hurdle_rate": rnd.choice([18.0, 0, 12.5, 25]), "settings": s, "internal_labor": internal, "external_labor": external,
            "general_expenses": exps(rnd.randint(0, 3)), "custom_costs": exps(rnd.randint(0, 2)), "phases": phases, "estimate": est}


def main():
    rnd = random.Random(20261124)
    cases = [demo()] + [random_case(rnd, i) for i in range(40)]
    # 기준정보가 비어 있으면 기본 기준정보 (fallback) 를 씀
    fallback_case = demo()
    fallback_case["settings"] = {}
    cases.append(fallback_case)
    out = []
    for c in cases:
        r = calculate_project_cost(json.loads(json.dumps(c)), SETTINGS)
        r.pop("computed_at", None)
        out.append({"input": c, "fallback": SETTINGS, "expected": r})
    path = ROOT / "tests" / "golden.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"[OK] {len(out)} cases -> {path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
