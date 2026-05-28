// =====================================================================
// 1. マスターデータ & 定数
// =====================================================================

const HOURLY_WAGE_MAP = {
  "営業(マネージャー)":             6100,
  "営業(ミドル)":                    4800,
  "営業(ジュニア)":                  3300,
  "プランナー(マネージャー)":        4800,
  "プランナー(ミドル・ジュニア)":    3300,
  "デジマケ(マネージャー・ミドル)":  4800,
  "スタジオ(一律)":                  3100,
  "P(マネージャー)":                 4700,
  "P(ミドル・ジュニア)":             3600,
  "FROGMAN":                         8300
};

const TYPE_THRESHOLDS = {
  "制作": { gross: { good: 70, warn: 50 }, op: { good: 60, warn: 40 } },
  "広告運用": { gross: { good: 20, warn: 10 }, op: null },
  "複合": { gross: { good: 50, warn: 30 }, op: { good: 40, warn: 25 } }
};

// Google API接続用の設定
const GOOGLE_CLIENT_ID = "276815950800-td0eanufiu9m3p4pv3dckuo0je4s2bqt.apps.googleusercontent.com";
const SPREADSHEET_ID   = "1z_fYeAj4_RBRfla269yhctOsl5WZIBkKNCSe1wIi0yE";
const SCOPES           = 'https://www.googleapis.com/auth/spreadsheets';

let tokenClient = null;
let googleAccessToken = null;

let currentProject = null;
let allProjectsInMemory = []; 
let pollingIntervalId = null;

// =====================================================================
// 2. 初期化 ＆ Google API読み込み
// =====================================================================

document.addEventListener("DOMContentLoaded", () => {
  const dateEl = document.getElementById("record_date");
  if (dateEl) dateEl.value = todayStr();

  // イベント登録
  document.getElementById("btn_google_auth").addEventListener("click", handleAuthClick);
  document.getElementById("btn_new_top").addEventListener("click", () => { initNewProject(); showEditView(); });
  document.getElementById("search_input").addEventListener("input", onSearchInput);
  document.getElementById("search_clear").addEventListener("click", clearSearch);
  
  // 🌟 一覧へ戻るボタンを押した時は、URLから案件IDを消去して一覧に戻る
  document.getElementById("btn_back").addEventListener("click", () => { 
    clearUrlParam();
    showTopView(); 
    syncFromGoogleSheets(); 
  });
  
  document.getElementById("btn_new_edit").addEventListener("click", () => { if (confirm("編集中の内容を破棄して新規作成しますか？")) initNewProject(); });
  document.getElementById("add_sales_row").addEventListener("click", () => addSalesBlock());
  document.getElementById("add_internal_row").addEventListener("click", () => addInternalRow());
  document.getElementById("project_type").addEventListener("change", () => { triggerAutoSave(); });
  document.getElementById("btn_export").addEventListener("click", exportProjectJSON);
  document.getElementById("btn_import_trigger").addEventListener("click", () => document.getElementById("json_file_input").click());
  document.getElementById("json_file_input").addEventListener("change", importProjectJSON);
  document.getElementById("btn_add_plan_new").addEventListener("click", addNewPlan);
  document.getElementById("btn_add_plan_copy").addEventListener("click", addCopyPlan);

  document.getElementById("btn_save").addEventListener("click", saveProjectExplicitly);
  document.getElementById("btn_protect").addEventListener("click", toggleEditProtect);

  // 基本情報の入力変更
  ["client_name", "project_name", "manager_name", "record_date"].forEach(id => {
    document.getElementById(id).addEventListener("input", () => {
      if (currentProject) currentProject[id] = document.getElementById(id).value.trim();
      triggerAutoSave();
    });
  });

  // Google API クライアントの初期化起動
  gapi.load('client', initGapiClient);
  
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: SCOPES,
    callback: (resp) => {
      if (resp.error) return;
      googleAccessToken = resp.access_token;
      gapi.client.setToken({ access_token: googleAccessToken });
      
      localStorage.setItem("gdrive_access_token", googleAccessToken);
      updateAuthButtonStatus(true);
      syncFromGoogleSheets();
    },
  });

  showTopView();
});

async function initGapiClient() {
  await gapi.client.init({ 
    discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4"]
  });
  
  const savedToken = localStorage.getItem("gdrive_access_token");
  if (savedToken) {
    googleAccessToken = savedToken;
    gapi.client.setToken({ access_token: googleAccessToken });
    updateAuthButtonStatus(true);
    syncFromGoogleSheets();
  }
}

function handleAuthClick() {
  tokenClient.requestAccessToken({ prompt: googleAccessToken ? 'none' : 'consent' });
}

function updateAuthButtonStatus(isConnected) {
  const btn = document.getElementById("btn_google_auth");
  if (!btn) return;
  if (isConnected) {
    btn.textContent = "✅ 社内DB同期中";
    btn.style.background = "#10b981";
    btn.style.color = "#fff";
  } else {
    btn.textContent = "🌐 社内DBにログイン";
    btn.style.background = "";
    btn.style.color = "";
  }
}

// =====================================================================
// 3. Googleスプレッドシート同期システム（全員で一元管理）
// =====================================================================

async function syncFromGoogleSheets() {
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A1',
    });
    const rows = response.result.values;
    if (rows && rows.length > 0 && rows[0][0]) {
      allProjectsInMemory = JSON.parse(rows[0][0]);
    } else {
      allProjectsInMemory = [];
    }
    
    // 🌟【直リンク対応】URLに `?id=xxxx` が含まれているかチェック
    const urlParams = new URLSearchParams(window.location.search);
    const targetId = urlParams.get('id');
    
    if (targetId) {
      // 対象の案件データが存在すれば、最初から編集画面を開く
      const existProj = allProjectsInMemory.find(p => p.id === targetId);
      if (existProj) {
        loadProjectIntoForm(targetId);
        showEditView();
        return;
      }
    }

    renderProjectGrid();
    startPolling();
  } catch (err) {
    console.error("Sheets同期エラー:", err);
    if (err.status === 401) {
      localStorage.removeItem("gdrive_access_token");
      googleAccessToken = null;
      updateAuthButtonStatus(false);
      tokenClient.requestAccessToken({ prompt: 'consent' });
    } else {
      alert("社内データベースとの同期に失敗しました。対象スプレッドシートのタブ名が「Sheet1」になっているか確認してください。");
    }
  }
}

async function syncToGoogleSheets() {
  if (!googleAccessToken) return;
  
  const jsonStr = JSON.stringify(allProjectsInMemory);
  try {
    await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A1',
      valueInputOption: 'RAW',
      resource: { values: [[jsonStr]] }
    });
    console.log("スプレッドシート共有DBに自動保存完了");
  } catch (err) {
    console.error("自動保存エラー:", err);
  }
}

function triggerAutoSave() {
  if (!currentProject) return;
  
  saveCurrentPlanState();
  calculateAndDisplay();

  if (currentProject.project_name) {
    if (!currentProject.id) {
      currentProject.id = "proj_" + Date.now();
      currentProject.createdAt = new Date().toISOString();
      currentProject.updatedAt = currentProject.createdAt;
      allProjectsInMemory.push(currentProject);
      // 🌟 新規作成されてIDが確定した瞬間、URLを書き換える
      updateUrlParam(currentProject.id);
    } else {
      const idx = allProjectsInMemory.findIndex(p => p.id === currentProject.id);
      if (idx !== -1) {
        currentProject.updatedAt = new Date().toISOString();
        allProjectsInMemory[idx] = currentProject;
      }
    }
  }
}

// 🌟【URL操作】アドレスバーに案件IDを埋め込む関数（画面はリロードされません）
function updateUrlParam(id) {
  const newUrl = window.location.protocol + "//" + window.location.host + window.location.pathname + '?id=' + id;
  window.history.pushState({ path: newUrl }, '', newUrl);
}

// 🌟【URL操作】一覧に戻った時にURLのパラメータを綺麗に消去する関数
function clearUrlParam() {
  const cleanUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
  window.history.pushState({ path: cleanUrl }, '', cleanUrl);
}

// =====================================================================
// 4. 画面レンダリング ＆ 計算コア
// =====================================================================

function showTopView() {
  document.getElementById("view_top").classList.remove("d-none");
  document.getElementById("view_edit").classList.add("d-none");
  startPolling();
}

function showEditView() {
  document.getElementById("view_top").classList.add("d-none");
  document.getElementById("view_edit").classList.remove("d-none");
  stopPolling();
}

function renderProjectGrid(filterWord = "") {
  const grid = document.getElementById("project_cards_grid");
  const noMsg = document.getElementById("no_projects_msg");
  const noSearch = document.getElementById("no_search_results");

  grid.innerHTML = "";
  noMsg.classList.add("d-none");
  noSearch.classList.add("d-none");

  if (allProjectsInMemory.length === 0) {
    noMsg.classList.remove("d-none");
    updateSearchCount(0, 0);
    return;
  }

  const kw = filterWord.trim().toLowerCase();
  const filtered = kw ? allProjectsInMemory.filter(p => matchesKeyword(p, kw)) : allProjectsInMemory;

  filtered.sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt));

  if (filtered.length === 0) {
    noSearch.classList.remove("d-none");
    updateSearchCount(0, allProjectsInMemory.length);
    return;
  }

  updateSearchCount(filtered.length, allProjectsInMemory.length);

  filtered.forEach(proj => {
    grid.appendChild(buildProjectCard(proj));
  });
}

function matchesKeyword(proj, kw) {
  const planFields = (proj.plans || []).flatMap(plan => [
    plan.planName,
    ...(plan.salesRows || []).map(r => r.name),
    ...(plan.salesRows || []).flatMap(r => (r.outsourcingRows || []).map(e => e.name)),
    ...(plan.internalRows || []).map(r => r.rank)
  ]);
  const fields = [proj.project_name, proj.client_name, proj.manager_name, proj.record_date, ...planFields];
  return fields.some(f => f && f.toLowerCase().includes(kw));
}

function updateSearchCount(filtered, total) {
  const el = document.getElementById("search_count");
  if (!el) return;
  const kw = document.getElementById("search_input").value.trim();
  el.textContent = kw ? `${total}件中 ${filtered}件 該当` : `${total}件の案件`;
}

function onSearchInput(e) {
  const kw = e.target.value;
  document.getElementById("search_clear").classList.toggle("d-none", !kw);
  renderProjectGrid(kw);
}

function clearSearch() {
  document.getElementById("search_input").value = "";
  document.getElementById("search_clear").classList.add("d-none");
  renderProjectGrid("");
}

function buildProjectCard(proj) {
  const activeIdx = proj.activePlanIndex || 0;
  const plan = proj.plans?.[activeIdx] || proj.plans?.[0] || {};
  const s  = plan.summary || {};
  const gm = s.grossMargin  ?? 0;
  const om = s.opMargin     ?? 0;
  const op = s.opProfit     ?? 0;
  const pt = plan.projectType || "制作";
  const th = TYPE_THRESHOLDS[pt] || TYPE_THRESHOLDS["制作"];

  const grossCls = getRateClass(gm, th.gross, s.grossProfit ?? 0);
  const opCls    = th.op ? getRateClass(om, th.op, op) : "val-neutral";
  const tagClass = { "制作": "tag-seisaku", "広告運用": "tag-koukoku", "複合": "tag-fukugo" }[pt] || "tag-seisaku";

  const card = document.createElement("div");
  card.className = `project-card ${proj.isProtected ? "protected" : ""}`;

  card.innerHTML = `
    <div class="project-card-header">
      <span class="project-card-name">${escapeHtml(proj.project_name || "（案件名未入力）")}</span>
      <span class="project-type-tag ${tagClass}">${escapeHtml(pt)}</span>
    </div>
    <div class="project-card-meta">
      <span class="meta-item">🏢 ${escapeHtml(proj.client_name || "-")}</span>
      <span class="meta-item">👤 ${escapeHtml(proj.manager_name || "-")}</span>
      <span class="meta-item">📅 ${escapeHtml(proj.record_date || "-")}</span>
    </div>
    <div class="project-card-metrics">
      <div class="metric-box">
        <span class="metric-label">売上総額</span>
        <span class="metric-value">${formatCurrency(s.totalSales ?? 0)}</span>
      </div>
      <div class="metric-box">
        <span class="metric-label">粗利率</span>
        <span class="metric-value ${grossCls}">${formatPercent(gm)}</span>
      </div>
      <div class="metric-box">
        <span class="metric-label">営利率</span>
        <span class="metric-value ${opCls}">${th.op ? formatPercent(om) : "-"}</span>
      </div>
    </div>
    <div class="project-card-footer">
      <button class="btn ${proj.isProtected ? 'btn-warning' : 'btn-secondary'} btn-xs protect-btn" data-id="${proj.id}">
        ${proj.isProtected ? '🔒 保護中' : '🔓 保護'}
      </button>
      <button class="btn btn-secondary btn-xs edit-btn" data-id="${proj.id}">✏️ 編集</button>
      <button class="btn btn-danger btn-xs delete-btn" data-id="${proj.id}" ${proj.isProtected ? 'disabled' : ''}>🗑️ 削除</button>
    </div>
  `;

  card.addEventListener("click", (e) => {
    if (e.target.closest(".protect-btn") || e.target.closest(".edit-btn") || e.target.closest(".delete-btn")) return;
    loadProjectIntoForm(proj.id);
    showEditView();
  });
  card.querySelector(".protect-btn").addEventListener("click", (e) => { e.stopPropagation(); toggleProjectProtectInline(proj.id); });
  card.querySelector(".edit-btn").addEventListener("click", (e) => { e.stopPropagation(); loadProjectIntoForm(proj.id); showEditView(); });
  card.querySelector(".delete-btn").addEventListener("click", (e) => { e.stopPropagation(); deleteProject(proj.id, proj.project_name); });
  return card;
}

function addSalesBlock(name = "", amount = "", extRows = []) {
  const container = document.getElementById("sales_rows_container");
  const block = document.createElement("div");
  block.className = "sales-block";

  block.innerHTML = `
    <div class="sales-row">
      <div class="sales-cell-left">
        <input type="text" class="table-input input-name sales-name" placeholder="売上項目名" value="${escapeAttr(name)}">
        <input type="text" class="table-input input-amount sales-amount" placeholder="¥0" value="${formatInputCurrency(amount)}">
        <button class="btn btn-danger btn-icon-only sales-del-btn" title="売上行を削除"><svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
      </div>
      <div class="sales-cell-right">
        <div class="ext-rows-container"></div>
        <div class="ext-add-row">
          <button class="btn btn-secondary btn-xs ext-add-btn"><svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>外注を追加</button>
          <span class="ext-subtotal">外注小計: <strong class="ext-subtotal-val">¥0</strong></span>
        </div>
      </div>
    </div>
  `;

  const salesNameInp   = block.querySelector(".sales-name");
  const salesAmountInp = block.querySelector(".sales-amount");
  const extContainer   = block.querySelector(".ext-rows-container");
  const extAddBtn      = block.querySelector(".ext-add-btn");
  
  block.querySelector(".sales-del-btn").addEventListener("click", () => {
    if (extContainer.children.length > 0 && !confirm("紐づく外注費もすべて削除されます。よろしいですか？")) return;
    block.remove(); triggerAutoSave();
  });

  salesNameInp.addEventListener("input", () => { triggerAutoSave(); });
  salesAmountInp.addEventListener("input", (e) => { applyLiveCurrencyFormat(e.target); triggerAutoSave(); });
  extAddBtn.addEventListener("click", () => addExtRow(block, extContainer));

  container.appendChild(block);
  extRows.forEach(er => addExtRow(block, extContainer, er.name, er.amount));
  
  if (!name && !amount) saveCurrentPlanState();
  calculateAndDisplay();
}

function addExtRow(salesBlock, extContainer, name = "", amount = "") {
  const row = document.createElement("div");
  row.className = "ext-row";
  row.innerHTML = `
    <span class="ext-indent">└</span>
    <input type="text" class="table-input input-name ext-name" placeholder="外注項目名" value="${escapeAttr(name)}">
    <input type="text" class="table-input input-amount ext-amount" placeholder="¥0" value="${formatInputCurrency(amount)}">
    <button class="btn btn-danger btn-icon-only ext-del-btn" title="外注行を削除"><svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
  `;

  row.querySelector(".ext-amount").addEventListener("input", (e) => { applyLiveCurrencyFormat(e.target); updateExtSubtotal(salesBlock); triggerAutoSave(); });
  row.querySelector(".ext-name").addEventListener("input", () => { triggerAutoSave(); });
  row.querySelector(".ext-del-btn").addEventListener("click", () => { row.remove(); updateExtSubtotal(salesBlock); triggerAutoSave(); });

  extContainer.appendChild(row);
  updateExtSubtotal(salesBlock);
  if (!name && !amount) saveCurrentPlanState();
  calculateAndDisplay();
}

function updateExtSubtotal(salesBlock) {
  let sub = 0;
  salesBlock.querySelectorAll(".ext-amount").forEach(inp => { sub += parseNumber(inp.value); });
  if (salesBlock.querySelector(".ext-subtotal-val")) salesBlock.querySelector(".ext-subtotal-val").textContent = formatCurrency(sub);
}

function addInternalRow(rank = "", wage = "", hours = "") {
  const tbody = document.getElementById("internal_table_body");
  const tr = document.createElement("tr");

  let opts = '<option value="">選択してください</option>';
  for (const key in HOURLY_WAGE_MAP) { opts += `<option value="${key}" ${key === rank ? "selected" : ""}>${key}</option>`; }

  tr.innerHTML = `
    <td><select class="table-select internal-rank">${opts}</select></td>
    <td><input type="text" class="table-input text-right internal-wage" value="${formatInputCurrency(wage)}" readonly placeholder="¥0"></td>
    <td><input type="number" class="table-input text-right internal-hours" value="${hours}" placeholder="0.0" step="0.1" min="0"></td>
    <td><input type="text" class="table-input text-right internal-cost" readonly placeholder="¥0"></td>
    <td class="col-actions"><button class="btn btn-danger btn-icon-only" title="行を削除"><svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button></td>
  `;

  const rankSel  = tr.querySelector(".internal-rank");
  const wageInp  = tr.querySelector(".internal-wage");
  const hoursInp = tr.querySelector(".internal-hours");
  const costInp  = tr.querySelector(".internal-cost");

  function calcRowCost() {
    const w = parseNumber(wageInp.value);
    const h = parseNumber(hoursInp.value);
    const c = Math.round(w * h);
    costInp.value = c > 0 ? formatInputCurrency(c) : "";
    triggerAutoSave();
  }

  rankSel.addEventListener("change", () => { wageInp.value = formatInputCurrency(rankSel.value ? (HOURLY_WAGE_MAP[rankSel.value] ?? "") : ""); calcRowCost(); });
  hoursInp.addEventListener("input", calcRowCost);
  tr.querySelector("button").addEventListener("click", () => { tr.remove(); triggerAutoSave(); });

  tbody.appendChild(tr);
  if (rank && wage) calcRowCost();
  else { calculateAndDisplay(); }
}

function parseNumber(val) {
  if (!val) return 0;
  const n = parseFloat(String(val).replace(/[¥,]/g, ""));
  return isNaN(n) ? 0 : n;
}

function formatCurrency(amount) {
  return "¥" + Math.round(parseNumber(amount)).toLocaleString("ja-JP");
}

function formatInputCurrency(amount) {
  if (!amount) return "";
  const n = Math.round(parseNumber(amount));
  return n === 0 ? "" : "¥" + n.toLocaleString("ja-JP");
}

function applyLiveCurrencyFormat(inputEl) {
  let nums = inputEl.value.replace(/[^\d]/g, "");
  inputEl.value = nums === "" ? "" : "¥" + parseInt(nums, 10).toLocaleString("ja-JP");
}

function formatPercent(rate) {
  return (rate === null || isNaN(rate)) ? "-" : rate.toFixed(1) + "%";
}

function calculateAndDisplay() {
  let tSales = 0, tOut = 0, tInt = 0;
  document.querySelectorAll(".sales-amount").forEach(inp => { tSales += parseNumber(inp.value); });
  document.querySelectorAll(".ext-amount").forEach(inp => { tOut += parseNumber(inp.value); });
  document.querySelectorAll(".internal-cost").forEach(inp => { tInt += parseNumber(inp.value); });

  document.getElementById("sales_total_display").textContent = formatCurrency(tSales);
  document.getElementById("outsourcing_total_display").textContent = formatCurrency(tOut);
  document.getElementById("internal_total_display").textContent = formatCurrency(tInt);

  const grossProfit = tSales - tOut;
  const opProfit    = tSales - tOut - tInt;
  const grossMargin = tSales > 0 ? (grossProfit / tSales) * 100 : null;
  const opMargin    = tSales > 0 ? (opProfit    / tSales) * 100 : null;

  document.getElementById("summary_total_sales").textContent = formatCurrency(tSales);
  document.getElementById("summary_total_outsourcing").textContent = formatCurrency(tOut);
  document.getElementById("summary_gross_profit").textContent = formatCurrency(grossProfit);
  document.getElementById("summary_gross_profit_rate").textContent = formatPercent(grossMargin);
  document.getElementById("summary_total_internal").textContent = formatCurrency(tInt);
  document.getElementById("summary_operating_profit").textContent = formatCurrency(opProfit);
  document.getElementById("summary_operating_profit_rate").textContent = formatPercent(opMargin);

  updateStatusBadges(grossProfit, grossMargin, opProfit, opMargin);
}

function updateStatusBadges(grossProfit, grossMargin, opProfit, opMargin) {
  const pt = document.getElementById("project_type")?.value || "制作";
  const th = TYPE_THRESHOLDS[pt] || TYPE_THRESHOLDS["制作"];
  setStatusBadge(document.getElementById("gross_status_badge"), document.getElementById("gross_status_text"), grossProfit, grossMargin, th.gross, "粗利");
  
  const opBadge = document.getElementById("op_status_badge");
  if (th.op === null) {
    opBadge.className = "status-badge-sm status-neutral d-none";
  } else {
    opBadge.classList.remove("d-none");
    setStatusBadge(opBadge, document.getElementById("op_status_text"), opProfit, opMargin, th.op, "営利");
  }
}

function setStatusBadge(badge, textEl, profit, margin, threshold, label) {
  badge.className = "status-badge-sm";
  if (margin === null) { badge.classList.add("status-neutral"); textEl.textContent = "データ未入力"; return; }
  if (profit < 0) { badge.className = "status-badge-sm status-critical"; textEl.textContent = `赤字 (${label}マイナス)`; }
  else if (margin >= threshold.good) { badge.className = "status-badge-sm status-good"; textEl.textContent = `良好（${formatPercent(margin)}）`; }
  else if (margin >= threshold.warn) { badge.className = "status-badge-sm status-warning"; textEl.textContent = `注意（${formatPercent(margin)}）`; }
  else { badge.className = "status-badge-sm status-danger"; textEl.textContent = `警告（${formatPercent(margin)}）`; }
}

function getRateClass(margin, th, profit) {
  if (!th) return "val-neutral";
  if (profit < 0) return "val-critical";
  if (margin >= th.good) return "val-good";
  if (margin >= th.warn) return "val-warning";
  return "val-danger";
}

// =====================================================================
// 5. タブ・複数プラン切り替え ＆ メモリ保存
// =====================================================================

function saveCurrentPlanState() {
  if (!currentProject) return;
  const activeIdx = currentProject.activePlanIndex;
  const salesRows = [];
  
  document.querySelectorAll(".sales-block").forEach(block => {
    const name = block.querySelector(".sales-name")?.value.trim() || "";
    const amount = parseNumber(block.querySelector(".sales-amount")?.value);
    const outsourcingRows = [];
    block.querySelectorAll(".ext-row").forEach(extRow => {
      const eName = extRow.querySelector(".ext-name")?.value.trim() || "";
      const eAmount = parseNumber(extRow.querySelector(".ext-amount")?.value);
      if (eName || eAmount > 0) outsourcingRows.push({ name: eName, amount: eAmount });
    });
    if (name || amount > 0 || outsourcingRows.length > 0) salesRows.push({ name, amount, outsourcingRows });
  });

  const internalRows = [];
  document.querySelectorAll("#internal_table_body tr").forEach(tr => {
    const rank = tr.querySelector(".internal-rank")?.value || "";
    const wage = parseNumber(tr.querySelector(".internal-wage")?.value);
    const hours = parseNumber(tr.querySelector(".internal-hours")?.value);
    if (rank || hours > 0) internalRows.push({ rank, wage, hours });
  });

  let tSales = 0, tOut = 0, tInt = 0;
  salesRows.forEach(r => { tSales += r.amount; (r.outsourcingRows || []).forEach(er => { tOut += er.amount; }); });
  internalRows.forEach(r => { tInt += Math.round(r.wage * r.hours); });

  currentProject.plans[activeIdx] = {
    planName: currentProject.plans[activeIdx]?.planName || `パターン ${activeIdx + 1}`,
    projectType: document.getElementById("project_type").value,
    salesRows, internalRows,
    summary: { totalSales: tSales, totalOutsourcing: tOut, totalInternal: tInt, grossProfit: tSales - tOut, grossMargin: tSales > 0 ? ((tSales - tOut)/tSales)*100 : 0, opProfit: tSales - tOut - tInt, opMargin: tSales > 0 ? ((tSales - tOut - tInt)/tSales)*100 : 0 }
  };
}

function renderTabs() {
  const tabsBar = document.getElementById("tabs_bar");
  if (!tabsBar || !currentProject) return;
  tabsBar.innerHTML = "";

  currentProject.plans.forEach((plan, idx) => {
    const tab = document.createElement("button");
    tab.className = `tab-item ${idx === currentProject.activePlanIndex ? "active" : ""}`;
    tab.type = "button";
    
    tab.innerHTML = `
      <span class="tab-text-view">${escapeHtml(plan.planName)}</span>
      <input type="text" class="tab-text-input d-none" value="${escapeAttr(plan.planName)}">
      ${currentProject.plans.length > 1 ? '<span class="tab-close">✕</span>' : ''}
    `;

    const textView  = tab.querySelector(".tab-text-view");
    const textInput = tab.querySelector(".tab-text-input");

    function startRename() {
      if (idx !== currentProject.activePlanIndex) return;
      textView.classList.add("d-none");
      textInput.classList.remove("d-none");
      textInput.focus();
      textInput.select();
    }

    tab.addEventListener("dblclick", (e) => {
      if (e.target.classList.contains("tab-close")) return;
      startRename();
    });

    tab.addEventListener("click", (e) => {
      if (e.target.classList.contains("tab-close")) return;
      if (idx === currentProject.activePlanIndex && !textInput.classList.contains("d-none")) return;
      
      if (idx === currentProject.activePlanIndex) {
        startRename();
      } else {
        switchPlan(idx);
      }
    });

    function finishRename() {
      const newName = textInput.value.trim();
      if (newName) {
        plan.planName = newName;
      }
      triggerAutoSave();
    }

    textInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); finishRename(); }
      if (e.key === "Escape") { textInput.value = plan.planName; textView.classList.remove("d-none"); textInput.classList.add("d-none"); }
    });

    textInput.addEventListener("blur", finishRename);

    if (currentProject.plans.length > 1) {
      tab.querySelector(".tab-close").addEventListener("click", (e) => {
        e.stopPropagation();
        removePlan(idx);
      });
    }

    tabsBar.appendChild(tab);
  });
}

function renderComparisonSummary() {
  const container = document.getElementById("comparison_summary");
  if (!container || !currentProject) return;
  container.innerHTML = "";
  currentProject.plans.forEach((plan, idx) => {
    const s = plan.summary || {};
    const item = document.createElement("div");
    item.className = `comparison-item ${idx === currentProject.activePlanIndex ? "active" : ""}`;
    item.innerHTML = `
      <div class="comp-title">${escapeHtml(plan.planName)}</div>
      <div class="comp-metrics">
        <div class="comp-metric"><span class="comp-lbl">売上:</span> <strong>${formatCurrency(s.totalSales ?? 0)}</strong></div>
        <div class="comp-metric"><span class="comp-lbl">粗利:</span> <strong class="val-good">${formatPercent(s.grossMargin ?? 0)}</strong></div>
        <div class="comp-metric"><span class="comp-lbl">営利:</span> <strong class="val-warning">${formatPercent(s.opMargin ?? 0)}</strong></div>
      </div>`;
    item.addEventListener("click", () => switchPlan(idx));
    container.appendChild(item);
  });
}

function switchPlan(index) { saveCurrentPlanState(); currentProject.activePlanIndex = index; loadPlanIntoForm(index); }
function addNewPlan() { saveCurrentPlanState(); currentProject.plans.push({ planName: `パターン ${currentProject.plans.length + 1}`, projectType: "制作", salesRows: [], internalRows: [], summary: {} }); currentProject.activePlanIndex = currentProject.plans.length - 1; loadPlanIntoForm(currentProject.activePlanIndex); }
function addCopyPlan() { saveCurrentPlanState(); const cp = JSON.parse(JSON.stringify(currentProject.plans[currentProject.activePlanIndex])); cp.planName = `${cp.planName} (コピー)`; currentProject.plans.push(cp); currentProject.activePlanIndex = currentProject.plans.length - 1; loadPlanIntoForm(currentProject.activePlanIndex); }
function removePlan(index) { if (currentProject.plans.length <= 1 || !confirm("削除しますか？")) return; currentProject.plans.splice(index, 1); if (currentProject.activePlanIndex >= currentProject.plans.length) currentProject.activePlanIndex = currentProject.plans.length - 1; loadPlanIntoForm(currentProject.activePlanIndex); triggerAutoSave(); }

function loadPlanIntoForm(index) {
  const plan = currentProject.plans[index];
  if (!plan) return;
  document.getElementById("project_type").value = plan.projectType || "制作";
  document.getElementById("sales_rows_container").innerHTML = "";
  document.getElementById("internal_table_body").innerHTML  = "";
  if (plan.salesRows?.length) plan.salesRows.forEach(r => addSalesBlock(r.name, r.amount, r.outsourcingRows || [])); else addSalesBlock();
  if (plan.internalRows?.length) plan.internalRows.forEach(r => addInternalRow(r.rank, r.wage, r.hours)); else addInternalRow();
  
  calculateAndDisplay();
  renderTabs();
  renderComparisonSummary();
}

// =====================================================================
// 6. データ読み込み ＆ 削除 コア
// =====================================================================

function initNewProject() {
  currentProject = { id: null, client_name: "", project_name: "", manager_name: "", record_date: todayStr(), activePlanIndex: 0, plans: [{ planName: "パターン 1", projectType: "制作", salesRows: [], internalRows: [], summary: {} }], isProtected: false };
  document.getElementById("client_name").value  = "";
  document.getElementById("project_name").value = "";
  document.getElementById("manager_name").value = "";
  document.getElementById("record_date").value  = todayStr();
  clearUrlParam(); // 新規作成時はURLパラメータをクリア
  loadPlanIntoForm(0);
  updateEditProtectButtonDisplay();
}

function loadProjectIntoForm(id) {
  const proj = allProjectsInMemory.find(p => p.id === id);
  if (!proj) return;
  currentProject = JSON.parse(JSON.stringify(proj));
  if (!currentProject.activePlanIndex) currentProject.activePlanIndex = 0;

  document.getElementById("client_name").value  = currentProject.client_name  || "";
  document.getElementById("project_name").value = currentProject.project_name || "";
  document.getElementById("manager_name").value = currentProject.manager_name || "";
  document.getElementById("record_date").value  = currentProject.record_date  || todayStr();

  // 🌟 編集画面を読み込んだ瞬間にアドレスバーのURLを固有のものに上書き
  updateUrlParam(id);

  loadPlanIntoForm(currentProject.activePlanIndex);
  updateEditProtectButtonDisplay();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function deleteProject(id, name) {
  if (!googleAccessToken) {
    alert("Google連携がされていません。");
    return;
  }
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A1',
    });
    const rows = response.result.values;
    let latestProjects = [];
    if (rows && rows.length > 0 && rows[0][0]) {
      latestProjects = JSON.parse(rows[0][0]);
    }

    const targetProj = latestProjects.find(p => p.id === id);
    if (targetProj && targetProj.isProtected) {
      alert(`案件「${name || "（無題）"}」は保護されているため削除できません。保護を解除してから削除してください。`);
      renderProjectGrid(document.getElementById("search_input").value);
      return;
    }

    if (!confirm(`案件「${name || "（無題）"}」を削除してもよろしいですか？`)) return;

    allProjectsInMemory = latestProjects.filter(p => p.id !== id);
    
    const jsonStr = JSON.stringify(allProjectsInMemory);
    await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A1',
      valueInputOption: 'RAW',
      resource: { values: [[jsonStr]] }
    });

    if (currentProject && currentProject.id === id) initNewProject();
    renderProjectGrid(document.getElementById("search_input").value);
  } catch (err) {
    console.error("削除エラー:", err);
    alert("削除の実行に失敗しました。最新のデータを取得できませんでした。");
  }
}

// =====================================================================
// 7. ローカルIO·ユーティリティ
// =====================================================================

function exportProjectJSON() {
  if (!currentProject) return; saveCurrentPlanState();
  if (!currentProject.project_name) { alert("案件名を入力してください。"); return; }
  const blob = new Blob([JSON.stringify(currentProject, null, 2)], { type: "application/json" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `${currentProject.project_name}_原価シート.json`; a.click();
}

function importProjectJSON(event) {
  const file = event.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      currentProject = data;
      if (!currentProject.plans) { currentProject.plans = [{ planName: "パターン 1", projectType: data.projectType||"制作", salesRows: data.salesRows||[], internalRows: data.internalRows||[], summary: data.summary||{} }]; currentProject.activePlanIndex = 0; }
      document.getElementById("client_name").value = currentProject.client_name || "";
      document.getElementById("project_name").value = currentProject.project_name || "";
      document.getElementById("manager_name").value = currentProject.manager_name || "";
      document.getElementById("record_date").value = currentProject.record_date || todayStr();
      loadPlanIntoForm(currentProject.activePlanIndex || 0);
      triggerAutoSave(); 
    } catch (err) { alert("JSON読み込み失敗"); }
  };
  reader.readAsText(file);
}

function todayStr() { return new Date().toISOString().split("T")[0]; }
function escapeHtml(str) { return str ? str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;") : ""; }
function escapeAttr(str) { return str ? str.replace(/"/g, "&quot;") : ""; }

// =====================================================================
// 8. 複数人上書き競合防止・自動ポーリング・保護機能用 新規関数
// =====================================================================

function startPolling() {
  if (pollingIntervalId) clearInterval(pollingIntervalId);
  pollingIntervalId = setInterval(async () => {
    console.log("定期自動同期（ポーリング）実行中...");
    await syncFromGoogleSheetsSilent();
  }, 30000); // 30秒間隔
}

function stopPolling() {
  if (pollingIntervalId) {
    console.log("定期自動同期（ポーリング）を停止しました。");
    clearInterval(pollingIntervalId);
    pollingIntervalId = null;
  }
}

async function syncFromGoogleSheetsSilent() {
  if (!googleAccessToken) return;
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A1',
    });
    const rows = response.result.values;
    if (rows && rows.length > 0 && rows[0][0]) {
      allProjectsInMemory = JSON.parse(rows[0][0]);
    } else {
      allProjectsInMemory = [];
    }
    
    // TOP画面が表示されているときのみグリッドを再描画
    const viewTop = document.getElementById("view_top");
    if (viewTop && !viewTop.classList.contains("d-none")) {
      renderProjectGrid(document.getElementById("search_input").value);
    }
  } catch (err) {
    console.warn("自動同期（サイレント）エラー:", err);
  }
}

async function saveProjectExplicitly() {
  if (!currentProject) return;

  saveCurrentPlanState();

  const projName = document.getElementById("project_name").value.trim();
  if (!projName) {
    alert("案件名を入力してください。");
    return;
  }

  currentProject.project_name = projName;
  currentProject.client_name = document.getElementById("client_name").value.trim();
  currentProject.manager_name = document.getElementById("manager_name").value.trim();
  currentProject.record_date = document.getElementById("record_date").value.trim();

  const saveBtn = document.getElementById("btn_save");
  const origText = saveBtn.textContent;
  saveBtn.disabled = true;
  saveBtn.textContent = "⏳ 保存中...";

  try {
    if (googleAccessToken) {
      // 最新のデータを再取得して上書き競合を防ぐ
      const response = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Sheet1!A1',
      });
      const rows = response.result.values;
      let latestProjects = [];
      if (rows && rows.length > 0 && rows[0][0]) {
        latestProjects = JSON.parse(rows[0][0]);
      }

      // 保護チェック：サーバー上ですでに保護されており、かつローカルで保護が解除されている（かつ以前とステータスが異なる）場合、
      // 意図せぬ競合上書きを防ぐためブロックする
      if (currentProject.id) {
        const serverProj = latestProjects.find(p => p.id === currentProject.id);
        if (serverProj && serverProj.isProtected && !currentProject.isProtected) {
          alert("この案件は他のメンバーによって保護されています。保存するには先に保護を解除するか、一旦一覧に戻って状態を確認してください。");
          saveBtn.disabled = false;
          saveBtn.textContent = origText;
          return;
        }
      }

      // ID確定とタイムスタンプの更新
      if (!currentProject.id) {
        currentProject.id = "proj_" + Date.now();
        currentProject.createdAt = new Date().toISOString();
        currentProject.updatedAt = currentProject.createdAt;
      } else {
        currentProject.updatedAt = new Date().toISOString();
      }

      // 最新プロジェクトリストに現在のデータのみをマージ
      const idx = latestProjects.findIndex(p => p.id === currentProject.id);
      if (idx !== -1) {
        latestProjects[idx] = currentProject;
      } else {
        latestProjects.push(currentProject);
      }

      allProjectsInMemory = latestProjects;

      // 保存処理
      const jsonStr = JSON.stringify(allProjectsInMemory);
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Sheet1!A1',
        valueInputOption: 'RAW',
        resource: { values: [[jsonStr]] }
      });

      updateUrlParam(currentProject.id);
      updateEditProtectButtonDisplay();

      saveBtn.textContent = "✅ 保存完了";
      setTimeout(() => {
        saveBtn.textContent = origText;
        saveBtn.disabled = false;
      }, 1500);
    } else {
      alert("Google連携がされていません。スプレッドシートへの保存はログインが必要です。");
      saveBtn.disabled = false;
      saveBtn.textContent = origText;
    }
  } catch (err) {
    console.error("保存エラー:", err);
    alert("スプレッドシートへの保存に失敗しました。通信状況を確認してください。");
    saveBtn.disabled = false;
    saveBtn.textContent = origText;
  }
}

function toggleEditProtect() {
  if (!currentProject) return;
  currentProject.isProtected = !currentProject.isProtected;
  updateEditProtectButtonDisplay();
}

function updateEditProtectButtonDisplay() {
  const btn = document.getElementById("btn_protect");
  if (!btn || !currentProject) return;
  if (currentProject.isProtected) {
    btn.textContent = "🔒 保護中";
    btn.className = "btn btn-warning btn-sm";
  } else {
    btn.textContent = "🔓 保護";
    btn.className = "btn btn-secondary btn-sm";
  }
}

async function toggleProjectProtectInline(id) {
  if (!googleAccessToken) {
    alert("Google連携がされていません。");
    return;
  }
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A1',
    });
    const rows = response.result.values;
    let latestProjects = [];
    if (rows && rows.length > 0 && rows[0][0]) {
      latestProjects = JSON.parse(rows[0][0]);
    }

    const targetProj = latestProjects.find(p => p.id === id);
    if (targetProj) {
      targetProj.isProtected = !targetProj.isProtected;
      allProjectsInMemory = latestProjects;

      const jsonStr = JSON.stringify(allProjectsInMemory);
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Sheet1!A1',
        valueInputOption: 'RAW',
        resource: { values: [[jsonStr]] }
      });

      renderProjectGrid(document.getElementById("search_input").value);
    }
  } catch (err) {
    console.error("保護ステータス変更エラー:", err);
    alert("保護状態の切り替えに失敗しました。通信状況を確認してください。");
  }
}