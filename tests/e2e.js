// =============================================================================
// 화면 시험 (브라우저, 약 20초) — node tests/e2e.js [--shots <폴더>]
//  실제 폴더 대신 브라우저 안의 폴더(OPFS)를 연결해 저장 · 열기 · 지우기 · 백업 · 올리기를 시험한다
//  예전 v4 프로그램으로 만든 파일(tests/fixtures/v4_project.xlsx)을 열어 같은 원가가 나오는지도 본다
// =============================================================================
"use strict";
const { chromium } = require("playwright");
const { spawn, execFileSync } = require("child_process");
const fs = require("fs"), path = require("path"), net = require("net");

const ROOT = path.resolve(__dirname, "..");
const shotDir = process.argv.includes("--shots") ? process.argv[process.argv.indexOf("--shots") + 1] : "";
const results = [];
const check = (name, ok, detail) => { results.push(ok); console.log(`  [${ok ? "PASS" : "FAIL"}] ${name}${ok || !detail ? "" : "  -- " + detail}`); };
const freePort = () => new Promise(r => { const s = net.createServer(); s.listen(0, "127.0.0.1", () => { const p = s.address().port; s.close(() => r(p)); }); });
const b64 = f => fs.readFileSync(f).toString("base64");

(async () => {
  const port = await freePort();
  const srv = spawn("python3", ["-m", "http.server", String(port), "--bind", "127.0.0.1"], { cwd: ROOT, stdio: "ignore" });
  await new Promise(r => setTimeout(r, 900));
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  const errors = [];
  page.on("pageerror", e => errors.push(e.message));
  page.on("console", m => { if (m.type() === "error") errors.push(m.text()); });
  try {
    await page.goto(`http://127.0.0.1:${port}/index.html`);
    await page.waitForFunction(() => window.PC && PC.App && PC.App.S.calc);
    check("Opens with no errors; asks to connect a folder", await page.isVisible("#noticeBox") && (await page.textContent("#noticeBox")).includes("폴더"));

    // 브라우저 안의 폴더(OPFS)에 예전 data 폴더 모양으로 파일을 넣고 연결
    await page.evaluate(async ([master, v4]) => {
      const root = await navigator.storage.getDirectory();
      for await (const [name] of root.entries()) await root.removeEntry(name, { recursive: true });
      const data = await root.getDirectoryHandle("data", { create: true });
      const put = async (dir, name, b) => { const fh = await (await data.getDirectoryHandle(dir, { create: true })).getFileHandle(name, { create: true }); const w = await fh.createWritable(); await w.write(Uint8Array.from(atob(b), c => c.charCodeAt(0))); await w.close(); };
      await put("master", "master_rates.xlsx", master);
      await put("projects", "PRJ-V4-OLD.xlsx", v4);
      await PC.Store.useHandle(root);
      await PC.App.afterConnect(true);
    }, [b64(path.join(ROOT, "data/master/master_rates.xlsx")), b64(path.join(ROOT, "tests/fixtures/v4_project.xlsx"))]);
    check("Connecting the app folder uses its data/ folder (old v4 layout) and lists the projects", (await page.textContent("#folderChip")).includes("/data") && (await page.$$eval("#projectSelect option", o => o.map(x => x.value))).includes("PRJ-V4-OLD.xlsx"));

    // v4 파일 열기 → 인건비 · 경비는 v4 와 같고, 예전 비율 방식 인비례 경비는 빠짐 (열 때 알림)
    await page.selectOption("#projectSelect", "PRJ-V4-OLD.xlsx");
    await page.waitForFunction(() => PC.App.S.file === "PRJ-V4-OLD.xlsx");
    await page.waitForSelector(".dialog");
    const note = await page.textContent(".dialog");
    await page.click('.dialog button[data-id="ok"]');
    const s1 = await page.evaluate(() => PC.App.S.calc.summary);
    check("A v4 file opens: labor and expenses as in v4 (internal 38,000,000 · external 22,500,000), old ratio overhead dropped with a notice", s1.breakdown.internal_labor === 38000000 && s1.breakdown.external_labor === 22500000 && s1.total_cost === 66500000 && note.includes("기준금액"), JSON.stringify(s1.breakdown) + " " + note.slice(0, 60));

    // 입력 → 저장 안 됨 표시 → Ctrl+S → 백업 · 다시 열면 그대로
    await page.click('[data-act="showTab"][data-args*="internal"]');
    const cell = page.locator('input[data-bind="internal_labor.0.mm.2026-05"]');
    await cell.fill("1"); await cell.press("Tab");
    await page.waitForFunction(() => PC.App.S.calc.summary.breakdown.internal_labor === 41600000, null, { timeout: 3000 }).catch(() => {});
    const ed = await page.evaluate(() => ({ state: document.getElementById("saveState").textContent, internal: PC.App.S.calc.summary.breakdown.internal_labor, focus: document.activeElement.dataset.bind || document.activeElement.tagName }));
    check("Editing a month MM recalculates at once and marks the project unsaved; Tab moves to the next month", ed.state.includes("저장 안 됨") && ed.internal === 41600000 && ed.focus === "internal_labor.0.mm.2026-06", JSON.stringify(ed));
    await page.keyboard.press("Control+s");
    await page.waitForFunction(() => !document.getElementById("saveState").classList.contains("dirty"));
    const after = await page.evaluate(async () => {
      const d = PC.Store.dir, names = async dir => { const o = []; for await (const [n] of dir.entries()) o.push(n); return o; };
      const backups = (await names(await d.getDirectoryHandle("backups"))).filter(n => n.endsWith(".xlsx"));
      await PC.App.setProject(PC.Model.newProject(PC.App.S.master), null, null);
      await PC.App.openFile("PRJ-V4-OLD.xlsx");
      const wb = await PC.Xlsx.read(await (await (await d.getDirectoryHandle("projects")).getFileHandle("PRJ-V4-OLD.xlsx")).getFile());
      return { backups, mm: PC.App.S.proj.internal_labor[0].mm["2026-05"], audit: wb.sheet("99_Audit_Log").rows.length, sheets: wb.sheets.map(s => s.name) };
    });
    check("Ctrl+S saves: previous file copied to backups, the change is there after reopening, save history grows", after.backups.length === 1 && after.mm === 1 && after.audit === 3, JSON.stringify(after));
    check("Saved file keeps the v4 sheet layout (01…07, 99_Audit_Log, State_JSON)", after.sheets.join() === "01_Project_Info,02_Project_Settings,03_Internal_Labor,04_External_Labor,05_Expenses,06_Phases,07_Estimate,99_Audit_Log,State_JSON", after.sheets.join());

    // 다른 곳(Excel)에서 먼저 바뀌면 저장 전에 알림
    await page.evaluate(async () => {
      const fh = await (await PC.Store.dir.getDirectoryHandle("projects")).getFileHandle("PRJ-V4-OLD.xlsx");
      const f = await fh.getFile(); const buf = await f.arrayBuffer();
      await new Promise(r => setTimeout(r, 1100));
      const w = await fh.createWritable(); await w.write(buf); await w.close();
      PC.App.S.proj.client_name = "변경"; PC.App.render();
    });
    await page.keyboard.press("Control+s");
    await page.waitForSelector(".dialog");
    check("A file changed elsewhere after opening is not overwritten silently (asks: reload / overwrite)", (await page.textContent(".dialog")).includes("바뀌었습니다"));
    await page.click('.dialog button[data-id="force"]');
    await page.waitForFunction(() => !document.getElementById("saveState").classList.contains("dirty"));

    // 새 프로젝트 → 데모 → 저장 → 목록 · 지우기
    await page.click('[data-act="newProject"]');
    await page.click('[data-act="showTab"][data-args*="info"]');
    await page.click('[data-act="loadDemo"]'); await page.click('.dialog button[data-id="ok"]');
    const demo = await page.evaluate(() => PC.App.S.calc.summary);
    check("Demo project: same figures as the v4 demo (revenue 240,000,000)", demo.revenue === 240000000 && demo.total_cost > 0, JSON.stringify(demo));
    // 금액 칸: 쉼표로 보이고, 쉼표를 넣어 입력해도 숫자로 읽음
    const rev = page.locator('input[data-bind="contract_revenue"]');
    const shown = await rev.inputValue();
    await rev.fill("250,000,000"); await rev.press("Tab");
    await page.waitForFunction(() => PC.App.S.proj.contract_revenue === 250000000, null, { timeout: 3000 }).catch(() => {});
    const typed = await page.evaluate(() => ({ v: PC.App.S.proj.contract_revenue, shown: document.querySelector('input[data-bind="contract_revenue"]').value }));
    check("Money fields show thousands separators and accept typed commas", shown === "240,000,000" && typed.v === 250000000 && typed.shown === "250,000,000", JSON.stringify({ shown, typed }));
    await rev.fill("240000000"); await rev.press("Tab");

    // ---- 인력 표를 엑셀처럼 ----
    await page.click('[data-act="showTab"][data-args*="internal"]');
    const B = b => `[data-bind="${b}"]`, focused = () => page.evaluate(() => document.activeElement.dataset.bind || document.activeElement.dataset.new || document.activeElement.tagName);
    await page.focus(B("internal_labor.0.mm.2026-01"));
    await page.keyboard.press("ArrowDown"); const f1 = await focused();
    await page.keyboard.press("ArrowRight"); const f2 = await focused();
    await page.keyboard.press("Enter"); const f3 = await focused();
    check("Arrow keys and Enter move between cells like a spreadsheet (Enter on the last person goes to the blank new row)", f1 === "internal_labor.1.mm.2026-01" && f2 === "internal_labor.1.mm.2026-02" && f3 === "internal_labor", JSON.stringify([f1, f2, f3]));   // 마지막 사람 줄에서 Enter → 새 줄(이름 칸)
    const paste = (sel, text) => page.evaluate(([s2, t]) => { const el = document.querySelector(s2); el.focus(); const dt = new DataTransfer(); dt.setData("text/plain", t); el.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true })); }, [sel, text]);
    // 엑셀에서 복사한 여러 줄 (성명 · 역할 · 등급) 을 새 줄에 붙여넣기 → 인력 2명 추가, 없는 등급은 빨갛게
    await paste('input[data-new="internal_labor"]', "박신규\t개발\t책임\r\n최추가\tQA\t없는등급\r\n");
    await page.waitForFunction(() => PC.App.S.proj.internal_labor.length === 4, null, { timeout: 3000 }).catch(() => {});
    const pv = await page.evaluate(() => ({ rows: PC.App.S.proj.internal_labor.map(r => [r.name, r.role, r.grade, r.base_rate]), invalid: document.querySelectorAll("select.invalid").length }));
    check("Pasting rows copied from Excel adds people (grade sets the rate; unknown grade shown in red)", pv.rows.length === 4 && pv.rows[2].join() === "박신규,개발,책임,7200000" && pv.rows[3][2] === "없는등급" && pv.invalid === 1, JSON.stringify(pv));
    // 머리글째 붙여넣기: 머리글 이름(성명 · 월)으로 열을 맞춤
    await paste(B("internal_labor.2.name"), "성명\t2026-02\t2026-03\n박신규\t0.5\t0.75");
    await page.waitForFunction(() => PC.App.S.proj.internal_labor[2].mm["2026-03"] === 0.75, null, { timeout: 3000 }).catch(() => {});
    const hv = await page.evaluate(() => PC.App.S.proj.internal_labor[2].mm);
    check("Pasting with a header row maps columns by header name (month columns)", hv["2026-02"] === 0.5 && hv["2026-03"] === 0.75, JSON.stringify(hv));
    // Ctrl+R 오른쪽 채우기 → Ctrl+Z 되돌리기 → Ctrl+Y 다시
    const before = await page.evaluate(() => JSON.stringify(PC.App.S.proj.internal_labor[3].mm));
    await page.fill(B("internal_labor.3.mm.2026-02"), "0.3");
    await page.keyboard.press("Control+r");
    await page.waitForFunction(() => PC.App.S.proj.internal_labor[3].mm["2026-06"] === 0.3, null, { timeout: 3000 }).catch(() => {});
    const filled = await page.evaluate(() => PC.App.S.proj.internal_labor[3].mm);
    await page.evaluate(() => document.activeElement.blur());
    await page.keyboard.press("Control+z");
    const undone = await page.evaluate(() => JSON.stringify(PC.App.S.proj.internal_labor[3].mm));
    await page.keyboard.press("Control+y");
    const redone = await page.evaluate(() => PC.App.S.proj.internal_labor[3].mm["2026-06"]);
    check("Ctrl+R fills the month to the right; Ctrl+Z undoes it, Ctrl+Y redoes", ["2026-02", "2026-03", "2026-04", "2026-05", "2026-06"].every(m => filled[m] === 0.3) && undone === before && redone === 0.3, JSON.stringify({ filled, undone, before, redone }));
    // 새 줄에 이름을 치고 Enter → 한 줄 추가, 그 줄의 역할 칸으로
    await page.fill('input[data-new="internal_labor"]', "정신입");
    await page.keyboard.press("Enter");
    await page.waitForFunction(() => PC.App.S.proj.internal_labor.length === 5, null, { timeout: 3000 }).catch(() => {});
    check("Typing a name in the blank last row adds a person and moves to the role cell", (await page.evaluate(() => PC.App.S.proj.internal_labor[4].name)) === "정신입" && (await focused()) === "internal_labor.4.role", await focused());
    await page.evaluate(() => { while (PC.App.S.proj.internal_labor.length > 2) PC.App.S.proj.internal_labor.pop(); PC.App.render(); });

    // MM 은 소수 둘째 자리까지
    await page.fill(B("internal_labor.1.mm.2026-06"), "0.25"); await page.keyboard.press("Tab");
    await page.waitForFunction(() => PC.App.S.proj.internal_labor[1].mm["2026-06"] === 0.25, null, { timeout: 3000 }).catch(() => {});
    check("MM accepts two decimals (0.25)", (await page.evaluate(() => PC.App.S.proj.internal_labor[1].mm["2026-06"])) === 0.25);
    await page.evaluate(() => { PC.App.S.proj.internal_labor[1].mm["2026-06"] = 0; PC.App.render(); });

    // 인비례 경비 = 계정별 MM × 기준금액 (데모: 복리후생비 300,000 × 인비례 대상 6 MM, 외주 관리비 100,000 × 외주 4 MM)
    const ov = await page.evaluate(() => PC.App.S.calc.summary.breakdown);
    check("Overhead is per account: MM x base amount (1,800,000 internal · 400,000 external in the demo)", ov.internal_overhead === 1800000 && ov.external_overhead === 400000, JSON.stringify(ov));

    // 경비 표도 엑셀처럼: 새 줄에 붙여넣기(분배 방식은 글자로) · 머리글째 달별 금액 붙여넣기 → 월별 직접 입력
    await page.click('[data-act="showTab"][data-args*="expense"]');
    await paste('input[data-new="general_expenses"]', "회의비\t매달 같은 금액\t500,000");
    await page.waitForFunction(() => PC.App.S.proj.general_expenses.length === 2, null, { timeout: 3000 }).catch(() => {});
    const ex1 = await page.evaluate(() => { const r = PC.App.S.proj.general_expenses[1]; return [r.name, r.dist_mode, r.unit_input, PC.App.S.calc.details.general_expenses[1].total_cost]; });
    await paste(B("general_expenses.0.name"), "계정\t2026-03\t2026-04\n출장비\t100\t200");
    await page.waitForFunction(() => PC.App.S.proj.general_expenses[0].dist_mode === "manual", null, { timeout: 3000 }).catch(() => {});
    const ex2 = await page.evaluate(() => ({ mode: PC.App.S.proj.general_expenses[0].dist_mode, m: PC.App.S.proj.general_expenses[0].monthly }));
    check("Expense table works like the staffing table (paste rows; month paste switches to monthly input, other months kept)", ex1.join() === "회의비,fill_all,500000,3000000" && ex2.mode === "manual" && ex2.m["2026-03"] === 100 && ex2.m["2026-04"] === 200 && ex2.m["2026-01"] === 1500000, JSON.stringify({ ex1, ex2 }));
    await page.keyboard.press("Control+z"); await page.keyboard.press("Control+z");

    // 프로젝트 정보: 원가 구성 (항목별 금액 · 비중)
    await page.click('[data-act="showTab"][data-args*="info"]');
    const comp = await page.$$eval("table.comp tr", rs => rs.map(r => r.innerText.replace(/\s+/g, " ").trim()));
    check("Project info shows the cost breakdown with shares (labor, general by item, overhead by account, other)", ["내부 인건비", "외주 인건비", "일반경비", "출장비", "인비례 경비", "복리후생비", "기타경비", "총 원가"].every(k => comp.some(t => t.startsWith(k))) && comp.some(t => /%$/.test(t)), comp.join(" | "));

    // 보고서 (2시트 엑셀) 내려받기
    // 파일 이름은 프로그램이 정한 이름으로 확인 (헤드리스 크롬은 한글 이름을 "download" 로 알려 줌 — 실제 크롬 · 엣지는 그대로 씀)
    await page.evaluate(() => { const orig = PC.UI.download; PC.UI.download = (b, n) => { window.__lastDownload = n; return orig(b, n); }; });
    const [dl] = await Promise.all([page.waitForEvent("download"), page.click('[data-act="exportReport"]')]);
    const rpath = path.join(require("os").tmpdir(), `pmhubcost_report_${Date.now()}.xlsx`);
    await dl.saveAs(rpath);
    const PCn = require("../js/util.js"); require("../js/xlsx.js");
    const rwb = await PCn.Xlsx.read(fs.readFileSync(rpath));
    check("Report button downloads a 2-sheet Excel report (통합보고 · 항목별상세)", rwb.sheets.map(x => x.name).join() === "01_통합보고,02_항목별상세" && /^PRJ-2026-01_원가보고서_\d{8}_\d{6}\.xlsx$/.test(await page.evaluate(() => window.__lastDownload)), await page.evaluate(() => window.__lastDownload));
    check("Theme button says Dark in light mode", (await page.textContent("#btnTheme")).trim() === "Dark");
    if (shotDir) {
      fs.mkdirSync(shotDir, { recursive: true });
      await page.screenshot({ path: path.join(shotDir, "light_info.png") });
      for (const tab of ["internal", "expense", "phase", "settings"]) { await page.click(`[data-act="showTab"][data-args*="${tab}"]`); await page.screenshot({ path: path.join(shotDir, `light_${tab}.png`), fullPage: true }); }
      for (let k = 0; k < 4; k++) await page.click('[data-act="addRow"][data-args*="settings.internal_grades"]');
      await page.screenshot({ path: path.join(shotDir, "light_settings_more.png"), fullPage: true });
      for (let k = 0; k < 4; k++) await page.keyboard.press("Control+z");
      await page.click('[data-act="toggleTheme"]');
      for (const tab of ["internal", "info"]) { await page.click(`[data-act="showTab"][data-args*="${tab}"]`); await page.screenshot({ path: path.join(shotDir, `dark_${tab}.png`) }); }
      await page.click('[data-act="toggleTheme"]');
    }
    await page.keyboard.press("Control+s");
    await page.waitForFunction(() => PC.App.S.file === "PRJ-2026-01.xlsx");
    check("Saving a new project creates projects/<ID>.xlsx and lists it", (await page.$$eval("#projectSelect option", o => o.map(x => x.value))).includes("PRJ-2026-01.xlsx"));
    await page.click('[data-act="showTab"][data-args*="info"]');
    await page.click('[data-act="deleteProject"]'); await page.click('.dialog button[data-id="ok"]');
    await page.waitForFunction(() => !PC.App.S.file);
    const del = await page.evaluate(async () => { const o = []; for await (const [n] of (await (await PC.Store.dir.getDirectoryHandle("backups")).getDirectoryHandle("deleted")).entries()) o.push(n); return o; });
    check("Deleting moves the file to backups/deleted", del.length === 1 && del[0].startsWith("PRJ-2026-01_deleted_"), del.join());

    // 바뀐 내용이 있을 때 다른 프로젝트로 가면 묻기
    await page.evaluate(() => { PC.App.S.proj.project_name = "고침"; PC.App.render(); });
    await page.selectOption("#projectSelect", "PRJ-V4-OLD.xlsx");
    await page.waitForSelector(".dialog");
    check("Leaving unsaved changes asks first (save / discard / cancel)", (await page.textContent(".dialog")).includes("저장하지 않은 변경"));
    await page.click('.dialog button[data-id="cancel"]');
    check("Cancel keeps the current project", await page.evaluate(() => PC.App.S.proj.project_name === "고침" && document.getElementById("projectSelect").value === ""));
    await page.evaluate(() => PC.App.setProject(PC.Model.newProject(PC.App.S.master), null, null));

    // 업로드 양식 열기
    await page.setInputFiles("#uploadInput", path.join(ROOT, "data/templates/project_upload_template.xlsx"));
    await page.waitForFunction(() => PC.App.S.proj.project_id !== PC.Model.newProject(PC.App.S.master).project_id);
    const up = await page.evaluate(() => ({ id: PC.App.S.proj.project_id, name: PC.App.S.proj.project_name, dirty: document.getElementById("saveState").textContent }));
    check("The v4 upload template opens as a new unsaved project", up.id === "PRJ-2026-02" && up.dirty.includes("저장 안 됨"), JSON.stringify(up));

    // 저장한 파일을 원본 v4(openpyxl)가 읽을 수 있는지
    const bytes = await page.evaluate(async () => { const f = await (await (await PC.Store.dir.getDirectoryHandle("projects")).getFileHandle("PRJ-V4-OLD.xlsx")).getFile(); return Array.from(new Uint8Array(await f.arrayBuffer())); });
    const tmp = path.join(require("os").tmpdir(), `pmhubcost_${Date.now()}.xlsx`);
    fs.writeFileSync(tmp, Buffer.from(bytes));
    let py = "";
    try { py = execFileSync("python3", ["-c", `import openpyxl,json,sys;wb=openpyxl.load_workbook(sys.argv[1]);print(json.loads(wb["State_JSON"]["A1"].value)["project_id"])`, tmp], { encoding: "utf8" }).trim(); } catch (e) { py = "python/openpyxl unavailable: " + e.message.split("\n")[0]; }
    check("A file saved by v5 is readable by v4 (openpyxl, State_JSON)", py === "PRJ-V4-OLD", py);
    check("No page errors", errors.length === 0, errors.join(" | "));
  } finally {
    await browser.close();
    srv.kill();
  }
  const ok = results.filter(Boolean).length;
  console.log(`\nResult: ${ok} / ${results.length} passed`);
  process.exit(ok === results.length ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
