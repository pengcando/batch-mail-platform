/* ==========================================================================
   应用外壳
   - 功能注册表：window.WB.register({ id, name, desc, icon, mount })
   - 路由：URL hash（#/功能id）
   - 通用工具：api / 确认弹窗 / 提示
   新增功能时不需要修改本文件。
   ========================================================================== */
(function () {
  "use strict";

  const ICONS = {
    mail: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="2.5" y="4.5" width="15" height="11" rx="2"/><path d="m3 5.5 7 5 7-5"/></svg>',
    table: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="2.5" y="3.5" width="15" height="13" rx="2"/><path d="M2.5 8h15M8 8v8.5"/></svg>',
    doc: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M5 2.5h6l4 4v11H5z"/><path d="M11 2.5v4h4"/></svg>',
    clock: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="10" cy="10" r="7.5"/><path d="M10 6v4.3l2.8 1.7"/></svg>',
    upload: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M10 13V3.5M6.5 7 10 3.5 13.5 7"/><path d="M3.5 12.5v3a1 1 0 0 0 1 1h11a1 1 0 0 0 1-1v-3"/></svg>',
    download: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M10 3.5V13M6.5 9.5 10 13l3.5-3.5"/><path d="M3.5 15.5h13"/></svg>',
    play: '<svg viewBox="0 0 20 20" fill="currentColor"><path d="M6.5 4.2v11.6L15.5 10z"/></svg>',
    stop: '<svg viewBox="0 0 20 20" fill="currentColor"><rect x="5.5" y="5.5" width="9" height="9" rx="1.5"/></svg>',
    check: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.7"><circle cx="10" cy="10" r="7.5"/><path d="m6.8 10.2 2.2 2.2 4.2-4.6"/></svg>',
    eye: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M2.5 10S5.5 4.8 10 4.8 17.5 10 17.5 10 14.5 15.2 10 15.2 2.5 10 2.5 10z"/><circle cx="10" cy="10" r="2.4"/></svg>',
    empty: '<svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="2"><rect x="7" y="12" width="34" height="26" rx="4"/><path d="M7 26h10l2.5 4h9l2.5-4h10"/></svg>',
    info: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="10" cy="10" r="7.5"/><path d="M10 9v5M10 6.4v.2"/></svg>',
    warn: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M10 3.2 18 16.8H2z"/><path d="M10 8v3.6M10 14.2v.2"/></svg>',
  };

  const FEATURES = [];
  let currentId = null;

  /* ---------------------------------------------------------------- api */

  async function handle(res) {
    const type = res.headers.get("content-type") || "";
    const payload = type.includes("application/json") ? await res.json() : await res.text();
    if (!res.ok) {
      const msg = (payload && payload.detail) || payload || `请求失败（${res.status}）`;
      throw new Error(typeof msg === "string" ? msg : JSON.stringify(msg));
    }
    return payload;
  }

  const api = {
    get: (url) => fetch(url).then(handle),
    postForm: (url, form) => fetch(url, { method: "POST", body: form }).then(handle),
    post: (url) => fetch(url, { method: "POST" }).then(handle),
  };

  /* ---------------------------------------------------------------- ui */

  const modal = document.getElementById("modal");
  const modalTitle = document.getElementById("modal-title");
  const modalBody = document.getElementById("modal-body");
  const modalOk = document.getElementById("modal-ok");
  const modalCancel = document.getElementById("modal-cancel");

  function confirmDialog({ title, bodyHtml, okText = "确认", danger = false }) {
    return new Promise((resolve) => {
      modalTitle.textContent = title || "确认操作";
      modalBody.innerHTML = bodyHtml || "";
      modalOk.textContent = okText;
      modalOk.className = "btn " + (danger ? "btn-danger" : "btn-primary");
      modal.hidden = false;
      // 安全默认：焦点落在「取消」
      requestAnimationFrame(() => modalCancel.focus());

      const done = (value) => {
        modal.hidden = true;
        modalOk.removeEventListener("click", onOk);
        modalCancel.removeEventListener("click", onCancel);
        modal.removeEventListener("keydown", onKey);
        document.removeEventListener("keydown", onKey);
        resolve(value);
      };
      const onOk = () => done(true);
      const onCancel = () => done(false);
      const onKey = (e) => {
        if (e.key === "Escape") { e.preventDefault(); done(false); }
        if (e.key === "Enter" && document.activeElement === modalOk) done(true);
      };
      modalOk.addEventListener("click", onOk);
      modalCancel.addEventListener("click", onCancel);
      document.addEventListener("keydown", onKey);
    });
  }

  const previewModal = document.getElementById("preview-modal");
  const previewFrame = document.getElementById("preview-frame");
  document.getElementById("preview-close").addEventListener("click", () => { previewModal.hidden = true; });
  previewModal.addEventListener("click", (e) => { if (e.target === previewModal) previewModal.hidden = true; });

  function showPreview(html) {
    previewFrame.srcdoc = html || "<p style='font:14px sans-serif;padding:20px;color:#666'>暂无预览</p>";
    previewModal.hidden = false;
  }

  const ui = { ICONS, confirm: confirmDialog, showPreview };

  /* ---------------------------------------------------------------- nav */

  function register(feature) {
    FEATURES.push(feature);
  }

  function renderNav() {
    const nav = document.getElementById("nav");
    nav.innerHTML = "";
    const group = document.createElement("div");
    group.className = "nav-group";
    group.textContent = "功能";
    nav.appendChild(group);

    FEATURES.forEach((f) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "nav-item" + (f.id === currentId ? " active" : "");
      btn.innerHTML =
        (ICONS[f.icon] || ICONS.doc) +
        `<span>${f.name}</span>` +
        (f.status === "wip" ? '<span class="badge">开发中</span>' : "");
      btn.addEventListener("click", () => { location.hash = "#/" + f.id; });
      nav.appendChild(btn);
    });
  }

  function mountFeature(id) {
    const feature = FEATURES.find((f) => f.id === id) || FEATURES[0];
    if (!feature) return;
    currentId = feature.id;
    document.getElementById("page-title").textContent = feature.name;
    document.getElementById("page-desc").textContent = feature.desc || "";
    document.getElementById("topbar-actions").innerHTML = "";
    const view = document.getElementById("view");
    view.innerHTML = "";
    renderNav();
    try {
      feature.mount(view, { api, ui });
    } catch (err) {
      view.innerHTML = `<div class="card"><div class="card-body"><div class="alert alert-danger">${ICONS.warn}<div>功能加载失败：${err.message}</div></div></div></div>`;
    }
  }

  function route() {
    const id = (location.hash || "").replace(/^#\/?/, "") || (FEATURES[0] && FEATURES[0].id);
    mountFeature(id);
  }

  window.WB = { register, api, ui, ICONS, FEATURES };
  window.addEventListener("hashchange", route);

  // 等所有功能脚本注册完毕后再启动
  window.addEventListener("load", () => {
    if (!FEATURES.length) {
      document.getElementById("view").innerHTML =
        '<div class="empty">没有已注册的功能。</div>';
      return;
    }
    route();
  });
})();
