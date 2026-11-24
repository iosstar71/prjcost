# PM Hub Cost — 작업 규칙

서버 없는 프로젝트 원가 관리 (브라우저 + 엑셀 파일 = DB). 사용자와는 **한국어**, 커밋 메시지 · 도구 출력은 영어.
PM Hub(저장소 `ioagent2`)의 자매 프로그램이라 화면 색(Slate · Obsidian)과 문구 습관을 맞춘다.

## 구조 (위치 지도)
| 무엇 | 어디 |
|---|---|
| 원가 계산 (원본 v4 파이썬과 같은 결과) | `js/calc.js` — 기준: `tests/reference/calculator_v4.py`, `tests/golden.json` |
| 반올림 · 월 계산 · 표시 도구 | `js/util.js` (`pyRound` = 파이썬 round) |
| 기본 기준정보 · 새 프로젝트 · 데모 · 예전 자료 채우기 | `js/model.js` |
| 엑셀 만들기 · 읽기 (외부 라이브러리 없음) | `js/xlsx.js` |
| 엑셀 시트 구성 (v4 와 같은 시트) | `js/workbook.js` |
| 폴더 저장 (File System Access API) · 백업 · 지우기 | `js/store.js` |
| 탭 화면 | `js/views.js` (`TABS`) |
| 엑셀처럼 쓰는 표 (키보드 이동 · 붙여넣기 · 새 줄 · 오른쪽 채우기) | `js/grid.js` — 열 정의는 `views.js` 의 `laborGrid` 열 배열 한 곳 (머리글 · 칸 · 붙여넣기 함께) |
| 흐름 (입력 반영 · 저장 · 열기 · 단축키) | `js/app.js` |
| 알림 · 확인 창 | `js/ui.js` |
| 꾸밈 (색 변수) | `assets/app.css` |

## 규칙
- **계산을 바꾸면** `tests/reference/calculator_v4.py` 를 같이 고치고 `python tools/make_golden.py` → `node tests/run.js`. 계산은 `calc.js` 에만 (화면에서 비율 · 합계를 새로 만들지 않음, 표 합계 줄 같은 표시용 합만 예외)
- **파일 형식**: 시트 이름 · 열은 v4 와 같게 유지 (예전 프로그램 · Excel 사용자와 호환). 자료를 바꾸면 `docs/DATA_MODEL.md` 와 `Model.normalize` 를 함께
- 일반 `<script>` 로 불러옴 (file:// 에서 ES 모듈이 막힘). 파일마다 `PC.xxx` 에 붙임, Node 시험용 `module.exports`
- 화면: 브라우저 `alert` · `confirm` 금지 (`PC.UI.toast` · `confirm` · `choose`). 새 코드에 인라인 `style=""` · 고정 색 금지 — 색은 `app.css` 의 변수
- 입력칸은 `data-bind="경로" data-type="text|num|blank-num|bool|pct|month2"`, 버튼은 `data-act` + `data-args`(JSON). 이벤트는 `app.js` 한 곳에서 받음
- 버전 · 이력은 코드 주석이 아니라 `RELEASE_NOTES.md` 에. 버전 번호는 `js/app.js` 의 `VERSION`
- 시험: 작은 수정 → `node --check` + `node tests/run.js`, 화면 · 저장 수정 → `node tests/e2e.js` 도
