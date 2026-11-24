// =============================================================================
// 화면 공통 — 알림(toast) · 확인 대화상자 · 파일 내려받기
//  브라우저 기본 alert / confirm 은 쓰지 않는다 (모양이 다르고 화면을 멈춤)
// =============================================================================
(function (root) {
  "use strict";
  const PC = root.PC;
  const { esc } = PC;

  function toast(msg, kind) {
    const box = document.getElementById("toasts");
    const el = document.createElement("div");
    el.className = `toast ${kind || ""}`;
    el.textContent = msg;
    box.appendChild(el);
    while (box.children.length > 3) box.firstChild.remove();   // 최근 3개만
    setTimeout(() => el.remove(), kind === "error" ? 7000 : 3500);
  }

  // 버튼 여러 개 중 하나를 고르는 창 → 고른 버튼의 id (Esc · 바깥 = cancel)
  function choose(title, message, buttons) {
    return new Promise(resolve => {
      const ov = document.createElement("div");
      ov.className = "overlay";
      ov.innerHTML = `<div class="dialog" role="dialog" aria-modal="true" aria-labelledby="dlgTitle"><h3 id="dlgTitle">${esc(title)}</h3><p>${esc(message)}</p>
        <div class="actions">${buttons.map(b => `<button class="btn ${b.cls || ""}" data-id="${esc(b.id)}">${esc(b.label)}</button>`).join("")}</div></div>`;
      const done = id => { ov.remove(); document.removeEventListener("keydown", onKey, true); resolve(id); };
      const onKey = e => { if (e.key === "Escape") { e.stopPropagation(); done("cancel"); } };
      ov.addEventListener("click", e => { const b = e.target.closest("button[data-id]"); if (b) done(b.dataset.id); else if (e.target === ov) done("cancel"); });
      document.addEventListener("keydown", onKey, true);
      document.body.appendChild(ov);
      const primary = ov.querySelector(".btn.primary") || ov.querySelector("button");
      if (primary) primary.focus();
    });
  }
  const confirm = async (title, message, okLabel, danger) =>
    (await choose(title, message, [{ id: "cancel", label: "취소" }, { id: "ok", label: okLabel || "확인", cls: danger ? "danger" : "primary" }])) === "ok";

  // 입력 칸이 있는 창 → { key: 값 } (취소 = null). fields: [{ key, label, value, ph, options(골라 쓰는 값), multiline }]
  function form(title, message, fields, okLabel) {
    return new Promise(resolve => {
      const ov = document.createElement("div");
      ov.className = "overlay";
      const ctl = (f, i) => f.multiline
        ? `<textarea rows="3" data-key="${esc(f.key)}" placeholder="${esc(f.ph || "")}">${esc(f.value || "")}</textarea>`
        : `<input type="text" data-key="${esc(f.key)}" value="${esc(f.value || "")}" placeholder="${esc(f.ph || "")}"${f.options ? ` list="dlgList${i}"` : ""}>${f.options ? `<datalist id="dlgList${i}">${f.options.map(o => `<option value="${esc(o)}"></option>`).join("")}</datalist>` : ""}`;
      ov.innerHTML = `<div class="dialog" role="dialog" aria-modal="true" aria-labelledby="dlgTitle"><h3 id="dlgTitle">${esc(title)}</h3>${message ? `<p>${esc(message)}</p>` : ""}
        <div class="dialog-form">${fields.map((f, i) => `<label class="field"><span>${esc(f.label)}</span>${ctl(f, i)}</label>`).join("")}</div>
        <div class="actions"><button class="btn" data-id="cancel">취소</button><button class="btn primary" data-id="ok">${esc(okLabel || "확인")}</button></div></div>`;
      const done = ok => {
        const out = {};
        ov.querySelectorAll("[data-key]").forEach(el => { out[el.dataset.key] = el.value.trim(); });
        ov.remove(); document.removeEventListener("keydown", onKey, true);
        resolve(ok ? out : null);
      };
      const onKey = e => {
        if (e.key === "Escape") { e.stopPropagation(); done(false); }
        else if (e.key === "Enter" && e.target.tagName === "INPUT" && ov.contains(e.target)) { e.preventDefault(); e.stopPropagation(); done(true); }
      };
      ov.addEventListener("click", e => { const b = e.target.closest("button[data-id]"); if (b) done(b.dataset.id === "ok"); else if (e.target === ov) done(false); });
      document.addEventListener("keydown", onKey, true);
      document.body.appendChild(ov);
      const first = ov.querySelector("[data-key]");
      if (first) { first.focus(); first.select && first.select(); }
    });
  }

  function download(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  PC.UI = { toast, choose, confirm, form, download };
})(window);
