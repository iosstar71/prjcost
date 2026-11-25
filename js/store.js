// =============================================================================
// 자료 폴더 — 고른 폴더의 엑셀 파일을 DB 처럼 읽고 씀 (File System Access API: 크롬 · 엣지)
//  폴더 구조는 원본 v4 의 data 폴더와 같음 → 지금 쓰던 data 폴더를 그대로 고르면 됨
//    master/master_rates.xlsx   기준정보 (새 프로젝트의 기본값)
//    projects/<ID>.xlsx          프로젝트 하나에 파일 하나
//    backups/                    저장 직전 파일 (프로젝트마다 최근 BACKUP_KEEP 개), deleted/ 지운 파일
//  고른 폴더는 브라우저(IndexedDB)에 기억 — 다시 열면 [폴더 다시 연결] 한 번으로 이어서 씀
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC;
  const { Xlsx: X, Workbook: W, Model: M, stamp, safeId } = PC;
  const BACKUP_KEEP = 20;
  const DB = "prjcost", KV = "kv", HANDLE_KEY = "dataDir";

  // ---- 고른 폴더 기억 (IndexedDB) ----
  const idb = mode => new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(KV);
    req.onsuccess = () => resolve(req.result.transaction(KV, mode).objectStore(KV));
    req.onerror = () => reject(req.error);
  });
  const kvGet = async k => { try { const s = await idb("readonly"); return await new Promise(r => { const q = s.get(k); q.onsuccess = () => r(q.result); q.onerror = () => r(null); }); } catch (e) { console.warn("[Store] idb get", e); return null; } };
  const kvSet = async (k, v) => { try { const s = await idb("readwrite"); await new Promise(r => { const q = s.put(v, k); q.onsuccess = q.onerror = () => r(); }); } catch (e) { console.warn("[Store] idb set", e); } };

  const Store = {
    dir: null,          // 자료 폴더 (data)
    label: "",
    supported: typeof window !== "undefined" && "showDirectoryPicker" in window,

    async _use(handle) {
      // 앱 폴더를 골랐으면 그 안의 data 를, data 폴더를 골랐으면 그대로
      let dir = handle;
      try {
        const sub = await handle.getDirectoryHandle("data");
        const hasData = await sub.getDirectoryHandle("projects").then(() => true, () => false) || await sub.getDirectoryHandle("master").then(() => true, () => false);
        if (hasData) dir = sub;
      } catch (e) { /* data 하위 폴더 없음 → 고른 폴더를 그대로 */ }
      for (const name of ["master", "projects", "backups", "templates", "reports"]) await dir.getDirectoryHandle(name, { create: true });
      await (await dir.getDirectoryHandle("backups")).getDirectoryHandle("deleted", { create: true });
      this.dir = dir;
      try { await this.ensureTemplates(); } catch (e) { console.warn("[Store] builtin templates", e); }
      this.label = dir === handle ? handle.name : `${handle.name}/data`;
      return true;
    },
    // 처음 연결: 폴더 고르기
    async connect() {
      const handle = await window.showDirectoryPicker({ id: "prjcost-data", mode: "readwrite" });
      await kvSet(HANDLE_KEY, handle);
      return this._use(handle);
    },
    // 다시 열었을 때: 기억한 폴더. 권한이 '묻기' 상태면 사용자가 버튼을 눌러야 함 (needsClick)
    async restore(interactive) {
      const handle = await kvGet(HANDLE_KEY);
      if (!handle) return { ok: false };
      const opts = { mode: "readwrite" };
      let perm = await handle.queryPermission(opts);
      if (perm !== "granted" && interactive) perm = await handle.requestPermission(opts);
      if (perm !== "granted") return { ok: false, needsClick: true, name: handle.name };
      await this._use(handle);
      return { ok: true };
    },
    // 시험용: 브라우저 내부 폴더(OPFS) 등 이미 가진 폴더로
    async useHandle(handle) { return this._use(handle); },

    async _sub(name) { return this.dir.getDirectoryHandle(name, { create: true }); },
    async _write(dirHandle, name, blob) {
      const fh = await dirHandle.getFileHandle(name, { create: true });
      const w = await fh.createWritable();
      await w.write(blob);
      await w.close();
      return (await fh.getFile()).lastModified;
    },
    async _file(dirHandle, name) { try { return await (await dirHandle.getFileHandle(name)).getFile(); } catch (e) { return null; } },
    // 손상된 파일은 지우지 않고 backups 로 옮겨 둠
    async _quarantine(dirHandle, name, prefix) {
      const f = await this._file(dirHandle, name);
      if (!f) return;
      await this._write(await this._sub("backups"), `${prefix}_${stamp()}.xlsx`, f);
      await dirHandle.removeEntry(name);
    },

    // ---- 기준정보 ----
    async loadMaster() {
      const dir = await this._sub("master");
      const f = await this._file(dir, "master_rates.xlsx");
      if (f) {
        try { return W.settingsFrom(await X.read(f), M.defaultSettings()); }
        catch (e) { console.warn("[Store] master file damaged", e); await this._quarantine(dir, "master_rates.xlsx", "bad_master"); }
      }
      const defaults = M.defaultSettings();
      await this.saveMaster(defaults);
      return defaults;
    },
    // 마스터 저장: 바꾸기 전 파일은 backups/master_rates_<시각>.xlsx 로
    async saveMaster(settings) {
      const dir = await this._sub("master"), cur = await this._file(dir, "master_rates.xlsx");
      if (cur) { try { await this._write(await this._sub("backups"), `master_rates_${stamp()}.xlsx`, cur); } catch (e) { console.warn("[Store] master backup", e); } }
      await this._write(dir, "master_rates.xlsx", X.build(W.masterSheets(M.masterOf(settings))));
    },

    // ---- 프로젝트 ----
    async listProjects() {
      const dir = await this._sub("projects"), out = [];
      for await (const [name, h] of dir.entries()) {
        if (h.kind !== "file" || !/\.xlsx$/i.test(name) || name.startsWith("~$")) continue;   // ~$… = Excel 이 열어 둔 잠금 파일
        const f = await h.getFile();
        out.push({ file: name, id: name.replace(/\.xlsx$/i, ""), mtime: f.lastModified });
      }
      return out.sort((a, b) => a.id.localeCompare(b.id, "ko"));
    },
    async loadProject(file, master) {
      const dir = await this._sub("projects");
      const f = await this._file(dir, file);
      if (!f) throw new Error(`파일을 찾을 수 없습니다: ${file}`);
      let wb;
      try { wb = await X.read(f); }
      catch (e) { await this._quarantine(dir, file, "bad_project"); throw new Error(`파일이 손상되어 backups 폴더로 옮겼습니다: ${file}`); }
      const json = X.readJson(wb.sheet("State_JSON"));
      if (!json) throw new Error(`${file} 에 프로젝트 자료(State_JSON 시트)가 없습니다. 이 프로그램으로 저장한 파일인지 확인하세요.`);
      return { project: M.normalize(PC.clone(json), master), raw: json, mtime: f.lastModified };
    },
    fileFor(project) { return `${safeId(project.project_id)}.xlsx`; },
    async exists(file) { return !!(await this._file(await this._sub("projects"), file)); },
    // 저장: ① 다른 곳에서 먼저 바뀌었으면 멈춤 ② 지금 파일을 backups 로 복사 ③ 새로 씀. 이전 저장 이력은 이어 붙임
    async saveProject(project, calc, file, expectedMtime) {
      const dir = await this._sub("projects");
      const cur = await this._file(dir, file);
      let audit = [];
      if (cur) {
        if (expectedMtime && Math.abs(cur.lastModified - expectedMtime) > 1) {
          const err = new Error("저장하기 전에 파일이 다른 곳(Excel 등)에서 바뀌었습니다. 다시 불러온 뒤 고쳐 저장하세요.");
          err.code = "CHANGED";
          throw err;
        }
        try { audit = W.auditRowsOf(await X.read(cur)); } catch (e) { console.warn("[Store] previous audit rows unreadable", e); }
        await this._backup(cur, file);
      }
      return this._write(dir, file, X.build(W.projectSheets(project, calc, audit)));
    },
    async _backup(fileObj, file) {
      const dir = await this._sub("backups"), stem = file.replace(/\.xlsx$/i, "");
      await this._write(dir, `${stem}_${stamp()}.xlsx`, fileObj);
      const mine = [];
      for await (const [name, h] of dir.entries()) if (h.kind === "file" && name.startsWith(`${stem}_`) && /^\d{8}_\d{6}\.xlsx$/.test(name.slice(stem.length + 1))) mine.push(name);
      mine.sort();
      for (const old of mine.slice(0, Math.max(0, mine.length - BACKUP_KEEP))) await dir.removeEntry(old);
    },
    // 지우기 = backups/deleted 로 옮김 (되살리려면 그 파일을 projects 로 다시 옮기면 됨)
    async deleteProject(file) {
      const dir = await this._sub("projects");
      const f = await this._file(dir, file);
      if (!f) throw new Error(`지울 파일이 없습니다: ${file}`);
      const dest = `${file.replace(/\.xlsx$/i, "")}_deleted_${stamp()}.xlsx`;
      await this._write(await (await this._sub("backups")).getDirectoryHandle("deleted", { create: true }), dest, f);
      await dir.removeEntry(file);
      return `backups/deleted/${dest}`;
    },

    async saveReport(blob, name) { await this._write(await this._sub("reports"), name, blob); },

    // ---- 보고서 양식 (templates/*.xlsx, 업로드 양식 제외) ----
    async listTemplates() {
      const dir = await this._sub("templates"), out = [];
      for await (const [name, h] of dir.entries()) if (h.kind === "file" && PC.Template.isReportTemplate(name)) out.push(name);
      return out.sort((a, b) => a.localeCompare(b, "ko"));
    },
    async templateFile(name) { return this._file(await this._sub("templates"), name); },
    // 기본 양식이 폴더에 없으면 넣어 줌 (이미 있으면 사용자가 고친 것이므로 그대로)
    async ensureTemplates() {
      const dir = await this._sub("templates"), all = PC.BUILTIN_TEMPLATES || {};
      for (const [name, b64] of Object.entries(all)) {
        if (await this._file(dir, name)) continue;
        const bin = atob(b64), bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        await this._write(dir, name, new Blob([bytes]));
      }
    },

    // ---- 올리기 (폴더 밖의 엑셀): 이 프로그램이 저장한 파일 또는 업로드 양식 ----
    async readUpload(fileObj, master) {
      const wb = await X.read(fileObj);
      const json = X.readJson(wb.sheet("State_JSON"));
      if (json) return M.normalize(json, master);
      const first = wb.sheets.find(s => !s.hidden);
      if (!first) throw new Error("읽을 시트가 없습니다.");
      return M.fromTemplateRows(first.rows.slice(1), master);
    }
  };

  PC.Store = Store;
})(window);
