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

    // v4 파일 열기 → 원본 파이썬과 같은 원가
    await page.selectOption("#projectSelect", "PRJ-V4-OLD.xlsx");
    await page.waitForFunction(() => PC.App.S.file === "PRJ-V4-OLD.xlsx");
    const s1 = await page.evaluate(() => PC.App.S.calc.summary);
    check("A file saved by v4 opens with the same cost as v4 (total 73,325,000 · margin 75.56%)", s1.total_cost === 73325000 && s1.margin_rate === 75.56, JSON.stringify(s1));

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
    if (shotDir) {
      fs.mkdirSync(shotDir, { recursive: true });
      await page.screenshot({ path: path.join(shotDir, "light_info.png") });
      for (const tab of ["internal", "expense", "phase", "settings"]) { await page.click(`[data-act="showTab"][data-args*="${tab}"]`); await page.screenshot({ path: path.join(shotDir, `light_${tab}.png`) }); }
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
