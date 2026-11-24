# 자료 정의

프로젝트 파일의 숨은 시트 `State_JSON`(A열, 3만 자마다 다음 줄로 나눔)에 아래 JSON 이 들어간다. v4 는 A1 한 칸 — 읽을 때 둘 다 받음.
기준정보 파일 `master_rates.xlsx` 의 숨은 시트 `Master_JSON` 에는 `settings` 와 같은 모양이 들어간다.

## project
| 필드 | 뜻 |
|---|---|
| project_id · project_name · client_name · pm_name · sales_rep · description | 기본 정보 (ID 가 처음 저장할 때 파일 이름) |
| date_mode | `month`(YYYY-MM) / `date`(YYYY-MM-DD) — 화면 입력 단위만 바뀜, 계산은 월 단위 |
| start_date · end_date | 기간 (예전 자료의 start_ym · end_ym 도 읽음) |
| contract_revenue · target_hurdle_rate | 계약 매출(원) · 목표 이익률(%) |
| settings | 이 프로젝트 전용 기준정보 (아래) — 비어 있으면 폴더의 기준정보 |
| internal_labor[] | name · role · grade · base_rate(비우면 등급 단가) · overhead_eligible · target_mm · promotions[{effective_ym, rate 또는 grade}] · mm{YYYY-MM: MM} |
| external_labor[] | name · role · vendor · grade · custom_rate(비우면 업체 · 등급 단가) · base_rate(단가표에 없을 때) · target_mm · mm{} |
| general_expenses[] · custom_costs[] | name · dist_mode(`equal_total` 총액 균등 / `fill_all` 매달 같은 금액 / `manual` 달마다) · total_input · unit_input · monthly{} · memo |
| phases[] (최대 5) | title · start_ym · end_ym · revenue · override_cost(비우면 기간 원가 합) |
| estimate | estimate_no · estimate_date · note · applied_to_revenue · items[{item_name, spec, qty, unit_price, memo}] |

## settings
internal_overhead_ratio · external_overhead_ratio (0.15 = 15%) · default_promotion_month("03") ·
internal_grades[{grade, rate}] · external_grades["특급", …] · vendor_rates[{vendor, grade, rate}] · external_rates{"업체|등급": 단가}(예전 형식) ·
expense_accounts[{account, category: general|custom, desc}]

## 계산 (js/calc.js — v4 와 같음)
- 내부 인건비 = MM × 그 달 단가 (승진월이 지나면 승진 단가). 내부 인비례 = 인비례 대상 인력 인건비 × 내부 비율
- 외주 인건비 = MM × 적용단가. 외주 인비례 = 외주 인건비 × 외주 비율
- 이익률 = (매출 − 총 원가) ÷ 매출 × 100. 반올림은 파이썬 round(…, 2) 와 같음 (한가운데는 짝수 쪽)
- 견적: 금액 = 수량 × 단가, 부가세 10%
