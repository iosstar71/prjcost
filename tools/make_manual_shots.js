// =============================================================================
// 사용자 매뉴얼 그림 만들기 — node tools/make_manual_shots.js
//  샘플 프로젝트(부서 2개)로 화면을 열어 manual/img/*.png 를 다시 찍는다 (화면이 바뀌면 다시 실행)
//  번호 표시(①②…)는 찍기 전에 화면 위에 잠깐 얹는다. 엑셀 결과 그림은 LibreOffice(soffice) · pdftoppm 이 있으면 만든다
// =============================================================================
"use strict";
const { chromium } = require("playwright");
const { spawn, execFileSync } = require("child_process");
const fs = require("fs"), path = require("path"), net = require("net"), os = require("os");

const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "manual", "img");
const W = 1280, H = 800;
const freePort = () => new Promise(r => { const s = net.createServer(); s.listen(0, "127.0.0.1", () => { const p = s.address().port; s.close(() => r(p)); }); });
const b64 = f => fs.readFileSync(f).toString("base64");
const has = c => { try { execFileSync("which", [c], { stdio: "ignore" }); return true; } catch (e) { return false; } };

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const port = await freePort();
  const srv = spawn("python3", ["-m", "http.server", String(port), "--bind", "127.0.0.1"], { cwd: ROOT, stdio: "ignore" });
  await new Promise(r => setTimeout(r, 900));
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: W, height: H }, deviceScaleFactor: 1, acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", e => errors.push(e.message));
  const shot = async (name, o = {}) => {
    const file = path.join(OUT, `${name}.png`);
    if (!o.keepToast) await page.evaluate(() => { document.getElementById("toasts").innerHTML = ""; });
    if (o.el && o.maxH) { const b = await page.locator(o.el).first().boundingBox(); await page.screenshot({ path: file, fullPage: true, clip: { x: b.x, y: b.y + (await page.evaluate(() => window.scrollY)), width: b.width, height: Math.min(b.height, o.maxH) } }); }
    else if (o.el) { await page.locator(o.el).first().screenshot({ path: file }); }
    else await page.screenshot({ path: file, fullPage: !!o.full, clip: o.clip });
    console.log("  [OK]", name);
  };
  // 번호 표시: [[선택자, 번호, 위치("tl" 왼쪽 위 · "tr" 오른쪽 위 · "l" 왼쪽 가운데)]]
  const mark = marks => page.evaluate(ms => {
    ms.forEach(([sel, n, pos]) => {
      const el = document.querySelector(sel);
      if (!el) { console.warn("mark: no", sel); return; }
      const r = el.getBoundingClientRect(), b = document.createElement("span");
      b.className = "__mark"; b.textContent = n;
      Object.assign(b.style, { position: "absolute", zIndex: 9999, width: "22px", height: "22px", borderRadius: "11px", background: "#e11d48", color: "#fff", font: "700 12px/22px sans-serif", textAlign: "center", boxShadow: "0 0 0 2px #fff", pointerEvents: "none" });
      const x = pos === "tr" ? r.right - 10 : pos === "l" ? r.left - 26 : r.left - 8, y = pos === "l" ? r.top + r.height / 2 - 11 : r.top - 10;
      b.style.left = `${x + window.scrollX}px`; b.style.top = `${Math.max(2, y) + window.scrollY}px`;
      document.body.appendChild(b);
      const o = document.createElement("span");
      o.className = "__mark";
      Object.assign(o.style, { position: "absolute", zIndex: 9998, left: `${r.left + window.scrollX - 2}px`, top: `${r.top + window.scrollY - 2}px`, width: `${r.width + 4}px`, height: `${r.height + 4}px`, border: "2px solid #e11d48", borderRadius: "6px", pointerEvents: "none" });
      document.body.appendChild(o);
    });
  }, marks);
  const unmark = () => page.evaluate(() => document.querySelectorAll(".__mark").forEach(x => x.remove()));
  const tab = id => page.click(`[data-act="showTab"][data-args*='"${id}"']`);
  const sub = id => page.evaluate(s => PC.App.A.showSettings(s), id);
  const closeDialog = () => page.evaluate(() => document.querySelectorAll(".overlay").forEach(o => o.remove()));

  try {
    await page.goto(`http://127.0.0.1:${port}/index.html`);
    await page.waitForFunction(() => window.PC && PC.App && PC.App.S.calc);
    await page.addStyleTag({ content: ".topbar { position: static !important; }" });   // 긴 화면을 찍을 때 머리줄이 내용을 덮지 않게
    // 1) 처음 화면: 폴더 연결 안내
    await mark([["#folderChip", "1", "tl"], ["#noticeBox", "2", "tl"]]);
    await shot("01_start", { clip: { x: 0, y: 0, width: W, height: 330 } });
    await unmark();

    // 브라우저 안 폴더(OPFS)를 자료 폴더로 연결하고 샘플 프로젝트(부서 2개 · 버전 1개)를 저장해 둠
    await page.evaluate(async ([master]) => {
      const root = await navigator.storage.getDirectory();
      for await (const [name] of root.entries()) await root.removeEntry(name, { recursive: true });
      const data = await root.getDirectoryHandle("data", { create: true });
      const fh = await (await data.getDirectoryHandle("master", { create: true })).getFileHandle("master_rates.xlsx", { create: true });
      const w = await fh.createWritable(); await w.write(Uint8Array.from(atob(master), c => c.charCodeAt(0))); await w.close();
      await PC.Store.useHandle(root);
      await PC.App.afterConnect(true);
      const p = PC.Model.sampleProject(PC.Model.defaultSettings());
      p.project_type = "SI-AD";
      p.settings.departments = ["SI사업부", "AI사업부"];
      p.settings.approvers = ["김철수", "이검토", "박사업", "최영업", "정본부"];
      p.internal_labor[0].dept = "SI사업부"; p.internal_labor[1].dept = "AI사업부"; p.external_labor[0].dept = "AI사업부";
      p.dept_revenue = { "SI사업부": 140000000, "AI사업부": 100000000 };
      p.phases[0].dept_revenue = { "SI사업부": 70000000, "AI사업부": 40000000 };
      p.phases[1].dept_revenue = { "SI사업부": 70000000, "AI사업부": 60000000 };
      p.ot_cost = 3000000;
      p.materials = [{ category: "HW", content: "개발 서버", amount: 5000000 }];
      p.custom_costs.push({ name: "AI예비비", dist_mode: "equal_total", total_input: 3000000, unit_input: 0, monthly: {}, memo: "" });
      p.versions = [{ no: 1, name: "초기 계획", memo: "착수 전 1차 산정", saved_at: "2026-01-05 10:00:00", data: Object.assign(PC.Model.snapshot(p), { contract_revenue: 230000000 }) }];
      PC.App.setProject(p, null, null);
      await PC.App.A.save();
    }, [b64(path.join(ROOT, "data/master/master_rates.xlsx"))]);
    await page.waitForFunction(() => PC.App.S.file);
    await tab("info");
    await page.evaluate(() => window.scrollTo(0, 0));

    // 2) 화면 구성
    await mark([["#folderChip", "1", "tl"], ["#projectSelect", "2", "tl"], ['[data-act="newProject"]', "3", "tl"], [".file-btn", "4", "tl"], ["#btnUndo", "5", "tl"],
      ['[data-act="exportReport"]', "6", "tl"], ['[data-act="exportTemplate"]', "7", "tl"], ['header [data-act="saveVersion"]', "8", "tl"], ['header [data-act="save"]', "9", "tl"], ["#btnManual", "10", "tl"],
      [".kpis", "11", "tl"], ["#tabs", "12", "tl"]]);
    await shot("02_layout", { clip: { x: 0, y: 0, width: W, height: 330 } });
    await unmark();

    // 3) 탭마다
    await shot("03_info", { el: "#panel" });
    await tab("internal");
    await mark([['[data-act="distribute"]', "1", "tl"], ['[data-act="distributeRange"]', "2", "tl"], ['[data-act="sortRows"]', "3", "tl"]]);
    await shot("04_internal", { el: "#panel" });
    await unmark();
    await page.click('[data-act="distributeRange"]');
    await page.waitForSelector(".dialog");
    await shot("05_distribute", { el: ".dialog" });
    await closeDialog();
    await tab("external"); await shot("06_external", { el: "#panel" });
    await tab("expense"); await shot("07_expense", { el: "#panel" });
    await tab("stats"); await shot("08_stats", { el: "#panel" });
    await tab("dept"); await shot("09_dept", { el: "#panel" });
    await tab("phase"); await shot("10_phase", { el: "#panel" });
    await tab("estimate"); await shot("11_estimate", { el: "#panel" });
    await tab("versions"); await shot("12_versions", { el: "#panel" });
    await page.click('header [data-act="saveVersion"]');
    await page.waitForSelector(".dialog");
    await shot("13_save_version", { el: ".dialog" });
    await closeDialog();
    await tab("settings");
    for (const s of ["rates", "vendors", "proposal_rates", "forms", "grades", "masterfile"]) { await sub(s); await shot(`14_set_${s}`, { el: "#panel" }); }

    // 4) 사업검토서 출력: 양식 고르기 → 만든 뒤 알림 (검증 결과)
    await tab("info");
    await page.click('[data-act="exportTemplate"]');
    await page.waitForSelector(".dialog-form");
    await page.selectOption(".dialog-form select", "Offline사업검토서_단가변경.xlsx");   // 샘플은 단가가 바뀌는 사람이 있어 _단가변경 양식으로
    await shot("15_template_dialog", { el: ".dialog" });
    const [dl] = await Promise.all([page.waitForEvent("download"), page.click('.dialog button[data-id="ok"]')]);
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "manual-"));
    const tfile = path.join(tmp, "review.xlsx"); await dl.saveAs(tfile);
    await page.waitForSelector(".toast");
    await shot("16_template_toast", { el: "#toasts", keepToast: true });
    const [dl2] = await Promise.all([page.waitForEvent("download"), page.click('[data-act="exportReport"]')]);
    const rfile = path.join(tmp, "report.xlsx"); await dl2.saveAs(rfile);

    // 5) 저장 상태: 고치면 "저장 안 됨"
    await tab("internal");
    const c = page.locator('input[data-bind="internal_labor.0.mm.2026-05"]');
    await c.fill("0.5"); await c.press("Tab");
    await mark([["#saveState", "1", "tl"], ['header [data-act="save"]', "2", "tl"]]);
    await shot("17_unsaved", { clip: { x: 640, y: 0, width: W - 640, height: 60 } });
    await unmark();

    // 6) 밝게 / 어둡게
    await page.click("#btnTheme");
    await tab("info");
    await shot("18_dark", { clip: { x: 0, y: 0, width: W, height: H } });
    await page.click("#btnTheme");

    // 7) 엑셀 결과 그림 (LibreOffice → PDF → PNG)
    if (has("soffice") && has("pdftoppm")) {
      execFileSync("soffice", ["--headless", "--convert-to", "pdf", "--outdir", tmp, tfile, rfile], { stdio: "ignore", env: Object.assign({}, process.env, { HOME: tmp }) });
      const png = (pdf, page1, name) => { execFileSync("pdftoppm", ["-png", "-r", "100", "-f", String(page1), "-l", String(page1), "-singlefile", path.join(tmp, pdf), path.join(OUT, name)]); console.log("  [OK]", name); };
      png("review.pdf", 1, "20_review_summary");
      png("review.pdf", 2, "21_review_detail");
      png("report.pdf", 1, "22_report");
    } else console.log("  [SKIP] Excel pictures (soffice / pdftoppm not found)");
    fs.rmSync(tmp, { recursive: true, force: true });
  } finally {
    await browser.close(); srv.kill();
  }
  if (errors.length) { console.error("page errors:", errors); process.exit(1); }
})().catch(e => { console.error(e); process.exit(1); });
