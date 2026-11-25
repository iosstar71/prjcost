# =============================================================================
# 보고서 양식 만들기 (개발할 때만 — openpyxl 필요): python3 tools/make_templates.py
#  → data/templates/Offline사업검토서.xlsx  (값이 들어갈 칸에 {{자리표시}})
#  → js/builtin-templates.js               (같은 파일을 base64 로 — 자료 폴더에 양식이 없으면 처음 연결할 때 넣어 줌)
#  양식은 이후 Excel 에서 직접 고쳐도 됨. 자리표시 이름은 js/template.js 의 KEYS · docs/TEMPLATE_KEYS.md
# =============================================================================
import sys, os, base64, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "templates", "Offline사업검토서.xlsx")
FONT = "맑은 고딕"
GRAY = PatternFill("solid", fgColor="D9D9D9")
WHITE = PatternFill("solid", fgColor="FFFFFF")
thin, thick = Side(style="thin", color="000000"), Side(style="medium", color="000000")
f10, f10b = Font(name=FONT, size=10), Font(name=FONT, size=10, bold=True)

# 결재 칸 직함 (이름은 {{결재1}} … {{결재5}})
APPROVERS = ["기안(PM)", "검토", "사업팀장", "영업팀장", "본부장"]

wb = Workbook()
ws = wb.active
ws.title = "사업개요"
# 기본 글꼴 (통합 문서 기본 스타일) · 기본 행 높이 18 · 기본 열 너비 10
wb._fonts[0] = Font(name=FONT, size=10)
ws.sheet_format.defaultRowHeight = 18
ws.sheet_format.customHeight = True
ws.sheet_format.defaultColWidth = 10
for col in "ABCDEFGHIJ":
    ws.column_dimensions[col].width = 10
ws.column_dimensions["A"].hidden = True
ws.column_dimensions["C"].width = 2
ws.column_dimensions["D"].width = 2


def cell(ref, value=None, font=f10, fill=None, h="left", v="center", wrap=False, fmt=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font
    if fill:
        c.fill = fill
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if fmt:
        c.number_format = fmt
    return c


def box(rng, outer=thin, inner=thin):
    """rng 의 모든 칸에 테두리: 바깥 outer, 안쪽 inner"""
    rows = list(ws[rng])
    for ri, row in enumerate(rows):
        for ci, c in enumerate(row):
            c.border = Border(
                top=outer if ri == 0 else inner, bottom=outer if ri == len(rows) - 1 else inner,
                left=outer if ci == 0 else inner, right=outer if ci == len(row) - 1 else inner)


def merge(rng, *a, **k):
    ws.merge_cells(rng)
    return cell(rng.split(":")[0], *a, **k)


# 1) 제목
merge("B1:J1", "[Offline 사업검토서]", font=Font(name=FONT, size=14, bold=True), h="center")

# 3) 결재 칸 F3:J4 — 진한 바깥 테두리, 흰 바탕, 4행 높이 24
for i, title in enumerate(APPROVERS):
    col = "FGHIJ"[i]
    cell(f"{col}3", title, fill=WHITE, h="center")
    cell(f"{col}4", "{{결재%d}}" % (i + 1), fill=WHITE, h="center")
ws.row_dimensions[4].height = 24
box("F3:J4", outer=thick, inner=thin)

# 4) 소제목 (로마자, 진하게 11pt)
heads = {6: "Ⅰ. 사업개요", 11: "Ⅱ. 사업원가", 21: "Ⅲ. 수주마진율", 30: "Ⅳ. RM부서 의견", 33: "Ⅴ. 기타사항", 36: "[수주협 5차]", 42: "[의사결정 사항]"}
for r, t in heads.items():
    cell(f"B{r}", t, font=Font(name=FONT, size=11, bold=True))

# 5) 7행 사업개요 — 내용은 위 · 왼쪽 맞춤, 줄바꿈
merge("B7:C7", "사업개요", fill=GRAY, h="center")
merge("D7:J7", "1. 과제명 : {{과제명}}\n2. 과제 내용\n{{과제내용}}", h="left", v="top", wrap=True)
ws.row_dimensions[7].height = 18 * 4 + 4   # 4줄

# 6) 8행 매출액 · 프로젝트 Type
merge("B8:C8", "매출액(자사/전체)\nVAT별도", fill=GRAY, h="center", wrap=True)
merge("D8:G8", "{{매출액}}", h="left", fmt='"₩"#,##0')
cell("H8", "프로젝트 Type", fill=GRAY, h="center", wrap=True)
merge("I8:J8", "{{프로젝트Type}}", h="center")
ws.row_dimensions[8].height = 46   # 라벨이 3줄 (B+C 너비 12)

# 7) 9행 사업기간 · 매출 귀속부서
merge("B9:C9", "사업기간", fill=GRAY, h="center")
merge("D9:G9", "{{시작일}} ~ {{종료일}} ({{개월수}}개월)", h="left")
cell("H9", "매출 귀속부서", fill=GRAY, h="center", wrap=True)
merge("I9:J9", "{{귀속부서}}", h="center")
ws.row_dimensions[9].height = 30

box("B7:J9")

f11 = Font(name=FONT, size=11)
LIGHT = PatternFill("solid", fgColor="F2F2F2")
MONEY, PCT = "#,##0", "0.00%"


def header(row, spans, titles):
    """제목 행: 회색, 11pt 보통 글씨, 가운데"""
    for (a, b), t in zip(spans, titles):
        if a != b:
            ws.merge_cells(f"{a}{row}:{b}{row}")
        cell(f"{a}{row}", t, font=f11, fill=GRAY, h="center")


def spans_merge(row, spans):
    for a, b in spans:
        if a != b:
            ws.merge_cells(f"{a}{row}:{b}{row}")


# Ⅱ. 사업원가 (12~19행): 구분 B:C · 금액 D:E · 백분율 F:G · 비고 H:J
S2 = [("B", "C"), ("D", "E"), ("F", "G"), ("H", "J")]
header(12, S2, ["구분", "금액", "백분율", "비고"])
# 재료비 · AS비는 지금 관리하지 않아 빈칸
cost_rows = [("노무비", "노무비", "내부 인건비"), ("외주비", "외주비", "외주 인건비"), ("재료비", None, ""),
             ("경비", "경비", "일반 · 인비례 · 기타경비 (AI예비비 제외)"), ("AI예비비", "AI예비비", "기타경비의 AI예비비"), ("AS비", None, ""), ("합계", "총원가", "")]
for i, (name, key, note) in enumerate(cost_rows):
    r = 13 + i
    spans_merge(r, S2)
    cell(f"B{r}", name, fill=LIGHT, h="center")
    cell(f"D{r}", "{{%s}}" % key if key else None, h="right", fmt=MONEY)
    cell(f"F{r}", "{{%s율}}" % key if key else None, h="center", fmt=PCT)
    cell(f"H{r}", note or None, h="left")
box("B12:J19")

# Ⅲ. 수주마진율 (22~28행): 구분 B:D · 계 E:F · 용역 G:I · 상품 J
S3 = [("B", "D"), ("E", "F"), ("G", "I"), ("J", "J")]
header(22, S3, ["구분", "계", "용역", "상품"])
margin_rows = [("매출액", "{{매출액}}", MONEY), ("원가", "{{총원가}}", MONEY), ("매출마진", "{{매출마진}}", MONEY),
               ("수주마진율", "{{수주마진율}}", PCT), ("H/R", "{{허들}}", PCT), ("Gap", "{{Gap}}", PCT)]
for i, (name, v, fmt) in enumerate(margin_rows):
    r = 23 + i
    spans_merge(r, S3)
    cell(f"B{r}", name, h="center")
    cell(f"E{r}", v, h="right", fmt=fmt)
    cell(f"G{r}", None); cell(f"J{r}", None)
box("B22:J28")

# Ⅳ. RM부서 의견 (31행 높이 100) · Ⅴ. 기타사항 (34행 높이 40): B:J 병합, 나중에 적도록 위 · 왼쪽
for r, h in [(31, 100), (34, 40)]:
    merge(f"B{r}:J{r}", None, h="left", v="top", wrap=True)
    ws.row_dimensions[r].height = h
    box(f"B{r}:J{r}")

# [수주협 5차] (37~40행): 구분 B:D · 매출 E:F · 매출마진 G:H · Gap사유 I:J (38~40 병합)
S5 = [("B", "D"), ("E", "F"), ("G", "H"), ("I", "J")]
header(37, S5, ["구분", "매출", "매출마진", "Gap 사유"])
# 수주협 4차 = 가장 최근 저장한 버전의 매출 · 마진율 (5차 · Gap 은 비워 둠)
for i, name in enumerate(["수주협 4차", "수주협 5차", "Gap"]):
    r = 38 + i
    spans_merge(r, S5[:3])
    cell(f"B{r}", name, h="center")
    cell(f"E{r}", "{{최근버전매출}}" if i == 0 else None, h="right", fmt=MONEY)
    cell(f"G{r}", "{{최근버전마진율}}" if i == 0 else None, h="right", fmt=PCT)
merge("I38:J40", None, h="left", v="top", wrap=True)
box("B37:J40")

# [의사결정 사항] (43행 높이 50): B = 진행사항 및 검토의견 (회색, 진하게 11pt), C:J 병합 — 진한 테두리
cell("B43", "진행사항 및\n검토의견", font=Font(name=FONT, size=11, bold=True), fill=GRAY, h="center", wrap=True)
merge("C43:J43", None, h="left", v="top", wrap=True)
ws.row_dimensions[43].height = 50
box("B43:J43", outer=thick, inner=thick)

# 인쇄: A4 세로 한 장
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1   # A4 한 장
ws.print_area = "B1:J45"
wb.save(OUT)
b64 = base64.b64encode(open(OUT, "rb").read()).decode()
js = os.path.join(ROOT, "js", "builtin-templates.js")
with open(js, "w", encoding="utf-8") as f:
    f.write("// 기본 보고서 양식 (tools/make_templates.py 가 만듦 — 고치지 말 것). 자료 폴더 templates 에 없으면 처음 연결할 때 넣어 줌\n")
    f.write("(function (root) {\n  \"use strict\";\n  const PC = root.PC || (root.PC = {});\n")
    f.write("  PC.BUILTIN_TEMPLATES = { %s: \"%s\" };\n" % (json.dumps(os.path.basename(OUT), ensure_ascii=False), b64))
    f.write("})(typeof window !== \"undefined\" ? window : globalThis);\n")
print("[OK]", OUT, os.path.getsize(OUT), "bytes ->", js)
