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
| internal_labor[] | name · role · area(담당 영역) · dept(부서) · grade · base_rate(비우면 등급 단가) · overhead_eligible · target_mm(총 MM — 달별 MM 을 고치면 그 합) · mm{YYYY-MM: MM} |
| external_labor[] | name · role · area · dept · vendor · grade · custom_rate(적용단가, 비우면 업체 · 등급 단가) · base_rate(단가표에 없을 때) · target_mm · mm{} |
| dept_revenue | 부서별 매출 {부서 이름: 원} — 부서 2개 이상일 때 [부서별] 탭에서 입력 |
| project_type | 보고서 양식용 프로젝트 Type |
| versions[] | 버전(기준선): no · name(초기 계획 · 실계약 …) · memo · saved_at · data(그때의 프로젝트 자료 — versions 는 빼고) |
| general_expenses[] · custom_costs[] | name · dept(비우면 공통) · dist_mode(`equal_total` 총액 균등 / `fill_all` 매달 같은 금액 / `manual` 달마다) · total_input · unit_input · monthly{} · memo |
| phases[] (최대 5) | title · start_ym · end_ym · revenue · override_cost(비우면 기간 원가 합) · dept_revenue{부서: 원}(회차 × 부서 매출) |
| estimate | estimate_no · estimate_date · note · applied_to_revenue · items[{item_name, spec, qty, unit_price, memo}] |

## settings
raise_month("03" — 연봉 인상월, 비우면 인상 없음) · internal_grades[{grade, rate, raise_rate(인상단가)}] · external_grades["특급", …] · vendor_rates[{vendor, grade, rate, raise_ym(인상월 YYYY-MM), raise_rate(인상단가)}] · external_rates{"업체|등급": 단가}(예전 형식) ·
areas["PMO", …](담당 영역) · approvers[5](보고서 결재자 이름: 기안(PM) · 검토 · 사업팀장 · 영업팀장 · 본부장) · departments["SI사업부", …](부서 — 2개 이상이면 부서 화면, 이름에 '.' 대신 '·') · dept_common_basis(`cost` 직접 원가 비율 / `mm` / `revenue` / `equal` — 공통경비 나누는 기준) ·
expense_accounts[{account, category, unit_amount, desc}] — category: `general` 일반경비 · `custom` 기타경비 · `mm_internal` 인비례(내부) · `mm_external` 인비례(외주), unit_amount = 인비례 기준금액(원/MM)
- 예전 internal_overhead_ratio · external_overhead_ratio(인건비의 %)는 더 쓰지 않음 — 열 때 알리고, 저장할 때 지움

## 계산 (js/calc.js)
- 내부 인건비 = MM × 그 달 단가 (연봉 인상월이 지나면 등급의 인상단가)
- 외주 인건비 = MM × 그 달 단가 (인상월이 지나면 인상단가)
- 연봉 인상월 = 프로젝트 기간 안에서 raise_month(MM)가 처음 오는 달 (1년 이내 기준 — 한 번)
- 내부 그 달 단가 = 기준단가 → 인상월부터 그 등급의 인상단가(넣은 등급만, 개인 기준단가보다 우선)
- 분할계약(phases) = 회차 원가(cost — 별도 원가가 있으면 그 값, calc_cost 는 계산값)와 그 기간의 internal_labor · external_labor · expense(일반 + 인비례 + 기타)
- 부서별 손익(Calc.deptSummary) = 인건비 · 인비례(계정마다 그 사람 MM × 기준금액)는 인력의 dept, 일반 · 기타경비는 항목의 dept(비거나 목록에 없으면 공통 → dept_common_basis 로 나눔), 매출은 dept_revenue(회차는 phases[i].dept_revenue — Calc.phaseDeptSummary). 인비례 · 공통경비는 달마다 나눠 더함(회차 합 = 전체). 회차에 별도 원가가 있으면 부서 비율대로 맞춤. 부서를 고르지 않은 인력은 '(미지정)' 줄. 끝수는 가장 큰 부서에 → 합계 = 통합
- 기준정보 순서(Calc.rankOf · orderRows) = 목록에 있는 값은 그 순서, 없는 값은 뒤, 빈 값은 맨 뒤. 투입 현황 · 원가 구성 경비 항목 · [기준정보 순서로] 줄 세우기에 씀
- 외주 그 달 단가 = 적용단가(비우면 업체 · 등급 단가) → 업체 · 등급의 인상월부터 인상단가
- 버전 비교(Calc.compareVersions) = 두 계산 결과의 요약(매출 · 인건비 · 경비 · 기타경비 · 총 원가 · 이익 · 이익률 · MM)과 사람별(내부/외주 + 이름) MM · 인건비 차이
- 예전 자료 변환(Model.normalize): 사람별 승진(promotions, v4) · 인상(rate_changes, v0.9.1)은 지움(열 때 Model.legacyNote 로 알림), 외주 rate_changes 는 그 업체 · 등급 단가의 인상(비어 있을 때), settings.default_promotion_month 지움, target_mm = 달별 MM 합
- 인비례 경비 = 계정마다 (내부: 인비례 대상 인력 MM 합 / 외주: 외주 인력 MM 합) × 기준금액 — 달마다
- 원가 구성(summary.composition) = internal_labor · external_labor · expense(일반경비 이름별 + 인비례 계정별) · custom_cost(기타경비 이름별)의 금액 · 총 원가 대비 비중
- 투입 현황(Calc.staffStats) = 인력 계획을 area · grade · role · vendor · kind 로 묶은 인원 · 달별 MM · 인건비 (months 를 주면 그 달만 — 분할계약 기간)
- 이익률 = (매출 − 총 원가) ÷ 매출 × 100. 반올림은 파이썬 round(…, 2) 와 같음 (한가운데는 짝수 쪽)
- 견적: 금액 = 수량 × 단가, 부가세 10%
