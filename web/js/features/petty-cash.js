/* ==========================================================================
   功能：对账单批量作业
   流程：城市 + 年月 → 按商户抽数生成对账单 xlsx → 渲染开票信息邮件 → 批量发送
   ========================================================================== */
(function () {
  "use strict";

  const { api, ui } = window.WB;
  const ICON = window.WB.ICONS;
  const BASE = "/api/petty-cash";

  const state = {
    mFile: null,
    tFile: null,
    cities: [],
    jobId: null,
    timer: null,
    rows: [],
    finished: false,
    meta: {},
  };

  /* ------------------------------------------------------------ templates */

  function uploadBox(id, title, subtitle, acceptLabel) {
    return `
      <label class="upload" id="${id}-box">
        <span class="upload-icon">${ICON.upload}</span>
        <span class="upload-text">
          <strong>${title}</strong>
          <span id="${id}-name">${subtitle}</span>
        </span>
        <input type="file" id="${id}" accept=".xlsx,.xlsm" />
      </label>`;
  }

  function toggleRow(id, title, subtitle, on) {
    return `
      <div class="switch">
        <div class="switch-text"><strong>${title}</strong><span>${subtitle}</span></div>
        <button type="button" class="toggle ${on ? "on" : ""}" id="${id}" aria-pressed="${on}"></button>
      </div>`;
  }

  function formHtml() {
    return `
    <div class="workspace">
      <div class="card sticky-col">
        <div class="card-head"><h2>运行参数</h2><span class="sub">批量作业</span></div>
        <div class="card-body">

          <div class="field">
            <label>源文件</label>
            ${uploadBox("m-file", "① 不开票商户明细", "不开票商户明细-{城市}.xlsx")}
            <div style="height:9px"></div>
            ${uploadBox("t-file", "② 商户对账工具", "商户对账工具-开票情况-V2.0.xlsx")}
            <div class="hint">两份文件每次运行都需重新上传，服务端不留存、不写死路径。</div>
          </div>

          <div class="field-row">
            <div class="field">
              <label for="city">城市</label>
              <select id="city" disabled><option value="">请先上传工具文件</option></select>
              <div class="hint" id="city-hint">与工具文件中的城市工作表联动</div>
            </div>
            <div class="field">
              <label for="ym">账期年月</label>
              <input type="month" id="ym" />
              <div class="hint">用于文件名与邮件主题</div>
            </div>
          </div>

          <details class="advanced">
            <summary>高级设置</summary>
            <div class="adv-body">
              ${toggleRow("skip-done", "跳过已处理行", "只处理名单中「进度」列为空的行，避免重跑时重复发信", true)}
              ${toggleRow("numeric-format", "对账单整表设为数值格式", "复刻原宏的「设为数值格式」；会覆盖模板里的百分比/百分比列格式，默认关闭", false)}
              <div class="field" style="margin-top:8px">
                <label for="month-mode">「月份」列取值</label>
                <select id="month-mode">
                  <option value="run_date">运行日期往前一个月（复刻原流程 EDATE(TODAY(),-1)）</option>
                  <option value="ym">直接取本次账期年月</option>
                </select>
              </div>
            </div>
          </details>

          <div class="field">
            <label for="smtp-pass">邮箱授权码</label>
            <div style="display:flex;gap:8px">
              <input type="password" id="smtp-pass" placeholder="发件邮箱的 SMTP 授权码" autocomplete="new-password" />
              <button type="button" class="btn btn-sm" id="verify-smtp" style="white-space:nowrap">测试连接</button>
            </div>
            <div class="hint" id="smtp-hint">发件人 <code>sender@example.com</code> · <code>smtp.example.com:465</code>（SSL）——请在 <code>app/config.py</code> 或环境变量里改成你自己的邮箱<br>授权码仅在本次请求内存中使用，不落盘、不记录。</div>
          </div>

          <div style="border-top:1px solid var(--border);margin:4px 0 14px"></div>

          ${toggleRow("dry-run", "预览模式（推荐）", "只生成文件与邮件预览，不发送任何邮件、不改动名单文件", true)}

          <div class="actions" style="margin-top:14px">
            <button class="btn btn-primary grow" id="run" type="button" disabled>${ICON.play}<span>开始处理</span></button>
            <button class="btn" id="reset" type="button" disabled>重置</button>
          </div>
          <div id="form-msg" style="margin-top:10px"></div>
        </div>
      </div>

      <div class="card">
        <div class="card-head">
          <h2>运行结果</h2>
          <div style="display:flex;gap:8px">
            <button class="btn btn-sm" id="preview-btn" type="button" disabled>${ICON.eye}邮件预览</button>
            <button class="btn btn-sm btn-primary" id="download-btn" type="button" disabled>${ICON.download}下载结果 ZIP</button>
          </div>
        </div>
        <div id="stats-slot"></div>
        <div id="table-slot"></div>
        <!-- 实时进度条：贴在结果下方，滚动看表格时始终可见 -->
        <div id="progress-slot"></div>
      </div>
    </div>`;
  }

  /* -------------------------------------------------------------- helpers */

  function el(id) { return document.getElementById(id); }

  function setToggle(button, on, danger) {
    button.classList.toggle("on", on);
    button.classList.toggle("danger", !!danger && on);
    button.setAttribute("aria-pressed", String(on));
  }

  function toggleIsOn(id) { return el(id).classList.contains("on"); }

  function prevMonth() {
    const d = new Date();
    d.setDate(1);
    d.setMonth(d.getMonth() - 1);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }

  function msg(html, kind) {
    const map = { info: "alert-info", warn: "alert-warn", danger: "alert-danger", ok: "alert-ok" };
    el("form-msg").innerHTML = html ? `<div class="alert ${map[kind] || "alert-info"}">${ICON.info}<div>${html}</div></div>` : "";
  }

  function statusTag(status) {
    if (status === "已完成") return '<span class="tag tag-ok">已完成</span>';
    if (status === "未完成") return '<span class="tag tag-fail">未完成</span>';
    return `<span class="tag tag-muted">${status || "-"}</span>`;
  }

  function esc(text) {
    return String(text == null ? "" : text)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function renderEmpty(text) {
    el("table-slot").innerHTML = `<div class="empty">${ICON.empty}<div>${text}</div></div>`;
  }

  function renderRows(rows) {
    if (!rows.length) { renderEmpty("暂无记录"); return; }
    const body = rows.map((r) => `
      <tr>
        <td class="merchant">${esc(r.merchant)}</td>
        <td class="mail">${esc(r.to) || '<span style="color:#c0392b">无有效收件人</span>'}</td>
        <td class="mail">${esc(r.cc) || "-"}</td>
        <td>${statusTag(r.status)}</td>
        <td class="num">${r.rows || 0}</td>
        <td class="num">${r.adjustRows || 0}</td>
        <td class="detail">${esc(r.detail)}</td>      </tr>`).join("");
    el("table-slot").innerHTML = `
      <div class="table-scroll">
        <table class="data">
          <thead><tr>
            <th>商户</th><th>收件人</th><th>抄送</th><th>状态</th>
            <th class="text-right">明细</th><th class="text-right">增减项</th><th>说明</th>
          </tr></thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;
  }

  function renderStats(job) {
    const s = job.summary || {};
    const has = Object.keys(s).length > 0;
    if (!has) { el("stats-slot").innerHTML = ""; return; }
    el("stats-slot").innerHTML = `
      <div class="stats">
        <div class="stat"><div class="k">商户总数</div><div class="v">${s.total ?? "-"}</div></div>
        <div class="stat"><div class="k">本月无流水</div><div class="v muted">${s.noFlow ?? 0}</div></div>
        <div class="stat"><div class="k">${s.dryRun ? "生成文件" : "发送成功"}</div><div class="v ok">${s.dryRun ? (s.files ?? 0) : (s.sent ?? 0)}</div></div>
        <div class="stat"><div class="k">未完成</div><div class="v fail">${s.failed ?? 0}</div></div>
        <div class="stat"><div class="k">耗时</div><div class="v muted">${s.elapsed ?? 0}s</div></div>
      </div>
      ${s.note ? `<div class="notice">${esc(s.note)}</div>` : ""}`;
  }

  /* ------------------------------------------------------ 实时进度条（吸底） */

  const PROGRESS_SKELETON = `
    <div class="progress-wrap" data-state="idle">
      <div class="progress-top">
        <span class="progress-stage"><i class="pdot" hidden></i><span class="p-text">等待运行</span></span>
        <span class="progress-count"><b class="progress-pct">0%</b><span class="p-num">—</span></span>
      </div>
      <div class="progress-track"><div class="progress-bar"></div></div>
      <div class="progress-meta">
        <span class="p-time"></span><span class="p-eta"></span><span class="p-speed"></span><span class="p-err"></span>
      </div>
    </div>`;

  const STATUS_LABEL = { pending: "排队中", running: "运行中", done: "已完成", failed: "失败", cancelled: "已取消" };
  const ACTIVE_STATUS = ["pending", "running"];

  // 进度条只建一次 DOM，之后原地改文本 / 宽度，CSS 过渡才走得动（不整体重绘）
  let PB = null;

  function ensureProgress() {
    if (PB && document.body.contains(PB.root)) return PB;
    const slot = el("progress-slot");
    slot.innerHTML = PROGRESS_SKELETON;
    const q = (sel) => slot.querySelector(sel);
    PB = {
      root: q(".progress-wrap"), text: q(".p-text"), dot: q(".pdot"),
      pct: q(".progress-pct"), num: q(".p-num"), bar: q(".progress-bar"),
      time: q(".p-time"), eta: q(".p-eta"), speed: q(".p-speed"), err: q(".p-err"),
    };
    return PB;
  }

  function resetProgressBar() {
    PB = null;
    el("progress-slot").innerHTML = "";
    renderProgress({ status: "idle" });
  }

  function fmtSec(sec) {
    const s = Math.max(0, Math.round(sec || 0));
    if (s < 60) return `${s}s`;
    const m = Math.floor(s / 60);
    const r = s % 60;
    return r ? `${m}分${String(r).padStart(2, "0")}秒` : `${m}分`;
  }

  function renderProgress(job) {
    const P = ensureProgress();
    job = job || {};

    const total = job.total || 0;
    const done = job.processed || 0;
    const raw = job.status || "idle";

    const state = raw === "done" ? "done"
      : (raw === "failed" || raw === "cancelled") ? "failed"
        : ACTIVE_STATUS.includes(raw) ? "running" : "idle";

    const pct = total ? Math.min(100, (done / total) * 100) : (state === "done" ? 100 : 0);
    // 还没拿到总数、或一个商户都还没跑完时，用流动条纹表示「正在忙」
    const indeterminate = state === "running" && (total === 0 || done === 0);

    // 预估剩余 / 处理速率：有样本之后才算，避免开局抖动
    const elapsed = job.elapsed || 0;
    let eta = null;
    let speed = null;
    if (state === "running" && done > 0 && total > done) {
      eta = (elapsed / done) * (total - done);
      speed = done / Math.max(elapsed, 0.1);
    }

    let headline;
    if (state === "idle") headline = "等待运行";
    else if (state === "done") headline = "已完成";
    else if (raw === "cancelled") headline = "已取消";
    else if (raw === "failed") headline = "运行失败";
    else headline = job.stage || STATUS_LABEL[raw] || "运行中";
    if (state === "running" && job.current) headline += " · " + job.current;

    P.root.dataset.state = state;
    P.text.textContent = headline;
    P.dot.hidden = state !== "running";
    P.pct.textContent = indeterminate ? "···" : `${Math.round(pct)}%`;
    P.num.textContent = total ? `${done} / ${total}` : (done ? String(done) : "—");

    P.bar.classList.toggle("indeterminate", indeterminate);
    P.bar.classList.toggle("done", state === "done");
    P.bar.classList.toggle("failed", state === "failed");
    if (!indeterminate) P.bar.style.width = `${pct}%`;

    P.time.textContent = (state === "running" || state === "failed") ? `已运行 ${fmtSec(elapsed)}` : "";
    P.eta.textContent = eta != null
      ? `预估剩余 ${fmtSec(eta)}`
      : (state === "done" ? `总耗时 ${fmtSec(elapsed)}` : "");
    P.speed.textContent = speed != null ? `${speed.toFixed(1)} 商户/秒` : "";
    P.err.textContent = job.error ? `错误：${job.error}` : "";
  }

  /* --------------------------------------------------------------- 事件 */

  async function onToolFile(file) {
    state.tFile = file;
    el("t-file-name").textContent = `${file.name} · ${(file.size / 1024).toFixed(0)} KB`;
    el("t-file-box").classList.add("filled");
    const select = el("city");
    select.disabled = true;
    select.innerHTML = '<option value="">解析中…</option>';
    try {
      const fd = new FormData();
      fd.append("toolFile", file);
      const info = await api.postForm(`${BASE}/inspect`, fd);
      state.cities = info.cities || [];
      if (!state.cities.length) {
        select.innerHTML = '<option value="">未识别到城市工作表</option>';
        msg("工具文件里没有找到以城市命名的工作表，请确认上传的是《商户对账工具-开票情况-V2.0.xlsx》。", "danger");
      } else {
        select.innerHTML = state.cities
          .map((c) => `<option value="${esc(c.name)}">${esc(c.name)}（${c.rows} 行）</option>`)
          .join("");
        select.disabled = false;
        const preferred = state.cities.find((c) => c.rows > 0) || state.cities[0];
        select.value = preferred.name;
        el("city-hint").textContent = `共 ${state.cities.length} 个城市工作表`;
      }
      if (info.missingSheets && info.missingSheets.length) {
        msg(`工具文件缺少工作表：${info.missingSheets.join("、")}，处理会失败。`, "warn");
      }
      syncRunButton();
    } catch (err) {
      select.innerHTML = '<option value="">解析失败</option>';
      msg(`工具文件解析失败：${esc(err.message)}`, "danger");
      syncRunButton();
    }
  }

  function onMerchantFile(file) {
    state.mFile = file;
    el("m-file-name").textContent = `${file.name} · ${(file.size / 1024).toFixed(0)} KB`;
    el("m-file-box").classList.add("filled");
    syncRunButton();
  }

  function syncRunButton() {
    el("run").disabled = !(state.mFile && state.tFile && state.cities.length) || (state.jobId && !state.finished);
  }

  function bindUpload(inputId, handler) {
    const input = el(inputId);
    input.addEventListener("change", () => { if (input.files[0]) handler(input.files[0]); });

    const box = el(inputId + "-box");
    ["dragenter", "dragover"].forEach((evt) =>
      box.addEventListener(evt, (e) => { e.preventDefault(); box.classList.add("dragover"); }));
    ["dragleave", "drop"].forEach((evt) =>
      box.addEventListener(evt, (e) => { e.preventDefault(); box.classList.remove("dragover"); }));
    box.addEventListener("drop", (e) => {
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) { input.files = e.dataTransfer.files; handler(file); }
    });
  }

  async function startRun() {
    const city = el("city").value;
    const ym = el("ym").value;
    if (!city) { msg("请选择城市。", "warn"); return; }
    if (!ym) { msg("请选择账期年月。", "warn"); return; }

    const dryRun = toggleIsOn("dry-run");
    const password = el("smtp-pass").value.trim();

    if (!dryRun) {
      if (!password) { msg("正式发送必须填写邮箱授权码。", "warn"); return; }
      const confirmed = await ui.confirm({
        title: "确认真实发送邮件",
        danger: true,
        okText: "确认发送",
        bodyHtml: `
          <p>即将向商户 <strong>真实发送</strong> 对账单邮件，且会回写《不开票商户明细》的「进度」列。</p>
          <ul>
            <li>城市：<strong>${esc(city)}</strong></li>
            <li>账期：<strong>${esc(ym)}</strong></li>
            <li>发送范围：<strong>${esc(state.mFile.name)}</strong>${toggleIsOn("skip-done") ? "（跳过已有进度的行）" : "（包含已有进度的行，将全量重发）"}</li>
          </ul>
          <p style="margin-top:10px;color:#d93025">邮件一经发出无法撤回，请确认参数无误。</p>`,
      });
      if (!confirmed) return;
    }

    const fd = new FormData();
    fd.append("merchantFile", state.mFile);
    fd.append("toolFile", state.tFile);
    fd.append("city", city);
    fd.append("ym", ym);
    fd.append("smtpPassword", password);
    fd.append("dryRun", String(dryRun));
    fd.append("skipDone", String(toggleIsOn("skip-done")));
    fd.append("monthMode", el("month-mode").value);
    fd.append("numericFormat", String(toggleIsOn("numeric-format")));

    msg("", null);
    state.jobId = null;
    state.finished = false;
    state.rows = [];
    el("download-btn").disabled = true;
    el("preview-btn").disabled = true;
    el("reset").disabled = false;
    el("run").disabled = true;
    el("run").innerHTML = '<span class="spinner"></span><span>处理中…</span>';
    resetProgressBar();
    renderProgress({ status: "running", stage: "排队中", elapsed: 0 });
    renderEmpty("正在处理，结果会实时出现在这里…");

    try {
      const res = await api.postForm(`${BASE}/run`, fd);
      state.jobId = res.jobId;
      poll();
    } catch (err) {
      msg(`启动失败：${esc(err.message)}`, "danger");
      el("run").disabled = false;
      el("run").innerHTML = `${ICON.play}<span>开始处理</span>`;
      renderEmpty("尚未开始");
    }
  }

  const ROWS_TAIL = 400;

  function poll() {
    clearInterval(state.timer);
    const tick = async () => {
      if (!state.jobId) return;
      let job;
      try {
        job = await api.get(`${BASE}/jobs/${state.jobId}`);
      } catch (err) {
        return;
      }
      renderProgress(job);
      if (job.rows.length !== state.rows.length || job.status !== "running") {
        state.rows = job.rows;
        renderRows(job.rows.slice(-ROWS_TAIL));
      }
      renderStats(job);

      if (job.status === "done" || job.status === "failed" || job.status === "cancelled") {
        clearInterval(state.timer);
        state.finished = true;
        el("run").disabled = false;
        el("run").innerHTML = `${ICON.play}<span>开始处理</span>`;
        el("download-btn").disabled = !job.hasResult;
        el("preview-btn").disabled = !(job.summary && job.summary.previewHtml);
        if (job.status === "done") {
          const s = job.summary || {};
          if (s.note) {
            // 空结果（如全部被「跳过已处理行」过滤）：直说原因，不假装成功
            msg(s.note, "warn");
          } else if (s.dryRun) {
            msg(
              `预览完成：共 ${s.total} 个商户，${s.noFlow} 个本月无流水，生成 ${s.files} 份对账单。未发送任何邮件。`,
              "ok"
            );
          } else {
            msg(`发送完成：成功 ${s.sent}，未完成 ${s.failed}，本月无流水 ${s.noFlow}。`, "ok");
          }
        } else if (job.status === "failed") {
          msg(`任务失败：${esc(job.error)}`, "danger");
        } else {
          msg("任务已取消。", "warn");
        }
      }
    };
    state.timer = setInterval(tick, 700);
    tick();
  }

  async function resetAll() {
    clearInterval(state.timer);
    state.mFile = state.tFile = null;
    state.jobId = null;
    state.rows = [];
    state.finished = false;
    el("run").innerHTML = `${ICON.play}<span>开始处理</span>`;

    ["m-file", "t-file"].forEach((id) => {
      el(id).value = "";
      el(id + "-box").classList.remove("filled");
      el(id + "-name").textContent = id === "m-file"
        ? "不开票商户明细-{城市}.xlsx"
        : "商户对账工具-开票情况-V2.0.xlsx";
    });
    const select = el("city");
    select.innerHTML = '<option value="">请先上传工具文件</option>';
    select.disabled = true;
    el("city-hint").textContent = "与工具文件中的城市工作表联动";
    el("smtp-pass").value = "";
    resetProgressBar();
    el("stats-slot").innerHTML = "";
    el("download-btn").disabled = true;
    el("preview-btn").disabled = true;
    el("reset").disabled = true;
    setToggle(el("skip-done"), true);
    setToggle(el("dry-run"), true);
    setToggle(el("numeric-format"), false);
    el("month-mode").value = "run_date";
    msg("", null);
    renderEmpty("还没有运行记录");
    syncRunButton();
  }

  /* ---------------------------------------------------------------- mount */

  function mount(root, ctx) {
    root.innerHTML = formHtml();
    const A = ctx.api, U = ctx.ui;
    // 用注入的实例，避免与外壳重复
    Object.assign(window.WB, { api: A, ui: U });

    el("ym").value = prevMonth();

    bindUpload("m-file", (f) => onMerchantFile(f));
    bindUpload("t-file", (f) => onToolFile(f));

    el("dry-run").addEventListener("click", () => {
      const on = !toggleIsOn("dry-run");
      setToggle(el("dry-run"), on, !on);
      el("smtp-pass").disabled = on;
      el("verify-smtp").disabled = on;
    });
    el("skip-done").addEventListener("click", () => setToggle(el("skip-done"), !toggleIsOn("skip-done")));
    el("numeric-format").addEventListener("click", () => setToggle(el("numeric-format"), !toggleIsOn("numeric-format")));

    el("verify-smtp").addEventListener("click", async () => {
      const password = el("smtp-pass").value.trim();
      if (!password) { msg("请先填写邮箱授权码。", "warn"); return; }
      const btn = el("verify-smtp");
      btn.disabled = true;
      btn.textContent = "检测中…";
      try {
        const fd = new FormData();
        fd.append("password", password);
        const res = await api.postForm(`${BASE}/verify-smtp`, fd);
        msg(res.ok ? "SMTP 登录成功，授权码可用。" : `SMTP 登录失败：${esc(res.message)}`, res.ok ? "ok" : "danger");
      } catch (err) {
        msg(`检测失败：${esc(err.message)}`, "danger");
      } finally {
        btn.disabled = false;
        btn.textContent = "测试连接";
      }
    });

    el("run").addEventListener("click", startRun);
    el("reset").addEventListener("click", resetAll);

    el("download-btn").addEventListener("click", () => {
      if (state.jobId) window.location.href = `${BASE}/jobs/${state.jobId}/download`;
    });
    el("preview-btn").addEventListener("click", async () => {
      if (!state.jobId) return;
      const job = await api.get(`${BASE}/jobs/${state.jobId}`);
      ui.showPreview(job.summary && job.summary.previewHtml);
    });

    renderEmpty("还没有运行记录");
    renderProgress({ status: "idle" });

    api.get(`${BASE}/meta`).then((meta) => {
      state.meta = meta;
      if (!meta.allowRealSend) {
        el("dry-run").disabled = true;
        msg("当前服务已锁定为预览模式，无法真实发信。", "warn");
      }
      el("smtp-hint").innerHTML =
        `发件人 ${esc(meta.sender)} · ${esc(meta.smtp.host)}:${meta.smtp.port}（SSL）<br>授权码仅在本次请求内存中使用，不落盘、不记录。`;
    }).catch(() => {});
  }

  window.WB.register({
    id: "petty-cash",
    name: "对账单批量作业",
    desc: "上传《不开票商户明细》与《商户对账工具》，选择城市与账期，自动生成每个商户的对账单文件、渲染开票信息邮件并批量发送。",
    icon: "mail",
    status: "ready",
    mount,
  });
})();
