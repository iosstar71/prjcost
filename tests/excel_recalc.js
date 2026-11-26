// =============================================================================
// 사업검토서 수식 시험 (개발용, 약 10초) — node tests/excel_recalc.js
//  여러 모양의 프로젝트로 사업검토서를 만든 뒤 LibreOffice 로 모든 수식을 다시 계산해
//   1) 「검증」 시트: 양식 수식 값(B) = 프로그램 값(C)
//   2) 모든 수식 칸: LibreOffice 값 = 이 프로그램 계산기(js/formula.js)가 넣은 값
//  필요: soffice(LibreOffice) · python3 + openpyxl. 없으면 건너뜀
// =============================================================================
"use strict";
const fs = require("fs"), path = require("path"), os = require("os");
const { execFileSync } = require("child_process");
const PC = require("../js/util.js");
["calc", "xlsx", "model", "formula", "template"].forEach(f => require(`../js/${f}.js`));
const M = PC.Model, T = PC.Template;

const which = c => { try { execFileSync("which", [c], { stdio: "ignore" }); return true; } catch (e) { return false; } };
if (!which("soffice")) { console.log("[SKIP] soffice (LibreOffice) not found"); process.exit(0); }

const months = (start, n) => Array.from({ length: n }, (_, i) => { const d = new Date(Number(start.slice(0, 4)), Number(start.slice(5)) - 1 + i, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; });
const spread = (ms, f) => Object.fromEntries(ms.map((m, i) => [m, f(i)]).filter(([, v]) => v));

// 시험할 프로젝트 모양
function cases() {
  const out = [];
  const base = () => M.sampleProject(M.defaultSettings());
  out.push(["sample", base()]);
  {   // 여러 명 · 연봉 인상 · OT · 재료비 · AI예비비 · 부서 2개 · 소수 MM
    const p = base(), ms = months("2026-01", 12);
    p.end_date = "2026-12"; p.phases[1].end_ym = "2026-12";
    p.settings.departments = ["SI사업부", "AI사업부"];
    p.ot_cost = 4500000;
    p.materials = [{ category: "HW", content: "서버", amount: 12000000 }, { category: "SW", content: "DB 라이선스", amount: 3333333 }];
    const g = ["수석", "책임", "선임", "선임", "책임", "수석", "선임", "책임"];
    p.internal_labor = g.map((grade, i) => ({ name: `내부${i + 1}`, role: "개발", grade, dept: i % 2 ? "AI사업부" : "SI사업부", site: "본사", base_rate: 0, overhead_eligible: i % 3 === 0, mm: spread(ms, k => ((k + i) % 4 === 0 ? 0 : [1, 0.5, 0.25, 0.75][(k + i) % 4])) }));
    p.external_labor = Array.from({ length: 5 }, (_, i) => ({ name: `외주${i + 1}`, role: "개발", vendor: "파트너A", grade: "고급", dept: "AI사업부", custom_rate: i === 4 ? 7123456 : "", mm: spread(ms, k => (k >= i ? 1 : 0)) }));
    p.custom_costs.push({ name: "AI예비비", dist_mode: "equal_total", total_input: 7777777, unit_input: 0, monthly: {}, memo: "" });
    p.general_expenses.push({ name: "회식비", dist_mode: "manual", total_input: 0, unit_input: 0, monthly: { "2026-03": 333333, "2026-09": 1000001 }, memo: "" });
    out.push(["many-people", p]);
  }
  {   // 외주 · 경비 없음
    const p = base();
    p.external_labor = []; p.general_expenses = []; p.custom_costs = [];
    out.push(["no-external-no-expense", p]);
  }
  {   // 아무도 없음
    const p = base();
    p.internal_labor = []; p.external_labor = []; p.general_expenses = []; p.custom_costs = []; p.contract_revenue = 0;
    out.push(["empty", p]);
  }
  {   // 15개월 (양식은 1년까지 — 5Tier 는 앞 12개월)
    const p = base(), ms = months("2026-01", 15);
    p.end_date = "2027-03"; p.phases[1].end_ym = "2027-03";
    p.internal_labor.forEach(r => { r.mm = spread(ms, () => 1); });
    out.push(["15-months", p]);
  }
  return out;
}

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "prjcost-recalc-"));
  let fail = 0;
  const list = [];
  for (const tpl of ["단가고정", "단가변경"]) cases().forEach(([n, p]) => list.push([`${tpl}-${n}`, p, tpl]));
  for (const [name, p, tpl] of list) {
    const src = fs.readFileSync(path.join(__dirname, "..", "data", "templates", `Offline사업검토서_${tpl}.xlsx`));
    const r = await T.fill(src, T.values(p, PC.Calc.calculateProjectCost(p, p.settings), {}));
    if (r.check.failed.length || r.check.problems.length) { fail++; console.log(`  [FAIL] ${name}: built-in check`, JSON.stringify(r.check.failed), JSON.stringify(r.check.problems)); }
    const ours = path.join(dir, `${name}.xlsx`);
    fs.writeFileSync(ours, Buffer.from(await r.blob.arrayBuffer()));
    // LibreOffice 가 계산하도록 수식 칸의 값은 지운 파일
    const files = await PC.Xlsx.unzip(fs.readFileSync(ours));
    const dec = new TextDecoder();
    const strip = Object.keys(files).map(n => ({ name: n, data: /^xl\/worksheets\//.test(n) ? new TextEncoder().encode(dec.decode(files[n]).replace(/(<c\b[^>]*?)\s+t="(?:str|e|b)"([^>]*>)(?=<f)/g, "$1$2").replace(/(<\/f>)<v>[^<]*<\/v>/g, "$1")) : files[n] }));
    fs.mkdirSync(path.join(dir, "in"), { recursive: true });
    fs.writeFileSync(path.join(dir, "in", `${name}.xlsx`), Buffer.from(await PC.Xlsx.zip(strip).arrayBuffer()));
  }
  execFileSync("soffice", ["--headless", "--calc", "--convert-to", "xlsx", "--outdir", path.join(dir, "lo"), ...list.map(([n]) => path.join(dir, "in", `${n}.xlsx`))], { stdio: "ignore", env: Object.assign({}, process.env, { HOME: dir }) });
  // 비교 (openpyxl 로 두 파일의 값 읽기)
  const py = `
import json, sys, openpyxl
ours = openpyxl.load_workbook(sys.argv[1], data_only=True); lo = openpyxl.load_workbook(sys.argv[2], data_only=True); fx = openpyxl.load_workbook(sys.argv[1])
out = {"cells": 0, "diff": [], "check": []}
def num(v): return 0 if v in (None, "") else v
for ws in fx:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                out["cells"] += 1
                a, b = num(ours[ws.title][c.coordinate].value), num(lo[ws.title][c.coordinate].value)
                if not (isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) <= max(1e-6, abs(b) * 1e-9)) and a != b:
                    out["diff"].append([ws.title, c.coordinate, c.value, a, b])
ck = lo["검증"]
for r in range(2, ck.max_row + 1):
    b, c = num(ck.cell(r, 2).value), num(ck.cell(r, 3).value)
    if not (isinstance(b, (int, float)) and abs(b - c) <= max(1e-6, abs(c) * 1e-9)): out["check"].append([ck.cell(r, 1).value, b, c])
print(json.dumps(out, ensure_ascii=False, default=str))
`;
  for (const [name] of list) {
    const res = JSON.parse(execFileSync("python3", ["-c", py, path.join(dir, `${name}.xlsx`), path.join(dir, "lo", `${name}.xlsx`)]).toString());
    const bad = res.diff.length || res.check.length;
    if (bad) fail++;
    console.log(`  [${bad ? "FAIL" : "OK"}] ${name}: ${res.cells} formula cells same as LibreOffice, check sheet ${res.check.length ? "differs" : "matches"}`);
    res.diff.slice(0, 5).forEach(d => console.log("     cell", JSON.stringify(d)));
    res.check.slice(0, 5).forEach(d => console.log("     check", JSON.stringify(d)));
  }
  fs.rmSync(dir, { recursive: true, force: true });
  console.log(`\nExcel recalc: ${list.length - fail} / ${list.length} cases passed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
