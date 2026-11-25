# =============================================================================
# 기본 보고서 양식 넣기 (개발할 때): python3 tools/make_templates.py
#  data/templates/ 의 보고서 양식(xlsx, 업로드 양식 제외)을 js/builtin-templates.js 에 base64 로 담음
#  → 자료 폴더 templates 에 같은 이름의 양식이 없으면 처음 연결할 때 넣어 줌
#  양식은 회사에서 Excel 로 고쳐서 data/templates/ 에 넣은 뒤 이 도구를 다시 돌리면 됨. 자리표시는 docs/TEMPLATE_KEYS.md
# =============================================================================
import base64, json, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "data", "templates")
names = sorted(n for n in os.listdir(SRC) if n.lower().endswith(".xlsx") and not n.startswith("~$") and not n.lower().startswith("project_upload_template"))
items = {n: base64.b64encode(open(os.path.join(SRC, n), "rb").read()).decode() for n in names}
out = os.path.join(ROOT, "js", "builtin-templates.js")
with open(out, "w", encoding="utf-8") as f:
    f.write("// 기본 보고서 양식 (tools/make_templates.py 가 만듦 — 고치지 말 것). 자료 폴더 templates 에 없으면 처음 연결할 때 넣어 줌\n")
    f.write("(function (root) {\n  \"use strict\";\n  const PC = root.PC || (root.PC = {});\n  PC.BUILTIN_TEMPLATES = {\n")
    f.write(",\n".join("    %s: \"%s\"" % (json.dumps(n, ensure_ascii=False), b) for n, b in items.items()))
    f.write("\n  };\n})(typeof window !== \"undefined\" ? window : globalThis);\n")
print("[OK]", ", ".join(f"{n} ({len(b) * 3 // 4:,} bytes)" for n, b in items.items()), "->", out)
