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
  "制作": { gross: { good: 70, warn: 50 }, direct: { good: 60, warn: 40 }, business: { good: 45, warn: 25 }, op: { good: 60, warn: 40 } },
  "広告運用": { gross: { good: 20, warn: 10 }, direct: null, business: null, op: null },
  "複合": { gross: { good: 50, warn: 30 }, direct: { good: 40, warn: 25 }, business: { good: 25, warn: 10 }, op: { good: 40, warn: 25 } }
};

// Google API接続用の設定
const GOOGLE_CLIENT_ID = "276815950800-td0eanufiu9m3p4pv3dckuo0je4s2bqt.apps.googleusercontent.com";
const SPREADSHEET_ID   = "1z_fYeAj4_RBRfla269yhctOsl5WZIBkKNCSe1wIi0yE";
const SCOPES           = 'https://www.googleapis.com/auth/spreadsheets';

// 認証＆24時間セッション管理定数
const STORAGE_KEY_TOKEN         = "gdrive_access_token";
const STORAGE_KEY_EXPIRES_AT    = "gdrive_token_expires_at";
const STORAGE_KEY_SESSION_UNTIL = "gdrive_session_until";
const SESSION_DURATION_MS       = 24 * 60 * 60 * 1000; // 24時間

let tokenClient = null;
let googleAccessToken = null;
let isRefreshingToken = false;
let refreshResolvers = [];
let tokenMaintenanceIntervalId = null;

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
    callback: handleTokenResponse,
  });

  showTopView();
});

// セッション・トークン管理ヘルパー
function isSessionActive() {
  const sessionUntil = parseInt(localStorage.getItem(STORAGE_KEY_SESSION_UNTIL) || "0", 10);
  return !!sessionUntil && Date.now() < sessionUntil;
}

function isTokenNearExpiry() {
  if (!googleAccessToken) return true;
  const expiresAt = parseInt(localStorage.getItem(STORAGE_KEY_EXPIRES_AT) || "0", 10);
  // 有効期限の5分前（300,000ms）を過ぎていたら再取得対象
  return Date.now() > (expiresAt - 300000);
}

function saveAuthSession(accessToken, expiresInSec) {
  googleAccessToken = accessToken;
  gapi.client.setToken({ access_token: googleAccessToken });
  
  const expiresIn = (expiresInSec && !isNaN(expiresInSec)) ? parseInt(expiresInSec, 10) : 3599;
  const expiresAt = Date.now() + (expiresIn * 1000);
  
  // 既存の有効な24時間セッションがあれば維持、なければ新規24時間後を設定
  let sessionUntil = parseInt(localStorage.getItem(STORAGE_KEY_SESSION_UNTIL) || "0", 10);
  if (!sessionUntil || sessionUntil < Date.now()) {
    sessionUntil = Date.now() + SESSION_DURATION_MS;
  }
  
  localStorage.setItem(STORAGE_KEY_TOKEN, googleAccessToken);
  localStorage.setItem(STORAGE_KEY_EXPIRES_AT, String(expiresAt));
  localStorage.setItem(STORAGE_KEY_SESSION_UNTIL, String(sessionUntil));
  
  updateAuthButtonStatus(true);
}

function clearAuthSession() {
  googleAccessToken = null;
  localStorage.removeItem(STORAGE_KEY_TOKEN);
  localStorage.removeItem(STORAGE_KEY_EXPIRES_AT);
  localStorage.removeItem(STORAGE_KEY_SESSION_UNTIL);
  updateAuthButtonStatus(false);
}

function handleTokenResponse(resp) {
  const resolvers = refreshResolvers.slice();
  refreshResolvers = [];
  isRefreshingToken = false;

  if (resp.error) {
    console.warn("OAuthトークン取得エラー:", resp.error);
    resolvers.forEach(r => r.reject(new Error(resp.error)));
    if (!isSessionActive() || isTokenNearExpiry()) {
      clearAuthSession();
    }
    return;
  }

  saveAuthSession(resp.access_token, resp.expires_in);
  resolvers.forEach(r => r.resolve(resp.access_token));
  
  // 初回データ未読み込みであれば同期
  if (allProjectsInMemory.length === 0) {
    syncFromGoogleSheets();
  }
}

// バックグラウンドでサイレントにトークンを自動更新する関数
function refreshAccessTokenSilently() {
  return new Promise((resolve, reject) => {
    if (!tokenClient) {
      return reject(new Error("tokenClient が初期化されていません"));
    }

    refreshResolvers.push({ resolve, reject });
    if (isRefreshingToken) return;
    isRefreshingToken = true;

    try {
      tokenClient.requestAccessToken({ prompt: '' });
    } catch (err) {
      isRefreshingToken = false;
      const resolvers = refreshResolvers.slice();
      refreshResolvers = [];
      resolvers.forEach(r => r.reject(err));
    }
  });
}

// API呼び出し前にトークンが新鮮であることを保証する
async function ensureFreshToken() {
  if (!isSessionActive()) {
    clearAuthSession();
    return false;
  }

  if (isTokenNearExpiry()) {
    try {
      await refreshAccessTokenSilently();
      return true;
    } catch (err) {
      console.warn("トークンのサイレント更新に失敗しました。既存トークンで試行します:", err);
      return !!googleAccessToken;
    }
  }
  return true;
}

// 定期的にトークンを監視し、切れる前にサイレント更新するタイマー
function startTokenMaintenance() {
  if (tokenMaintenanceIntervalId) clearInterval(tokenMaintenanceIntervalId);
  tokenMaintenanceIntervalId = setInterval(async () => {
    if (isSessionActive() && isTokenNearExpiry()) {
      console.log("トークン有効期限接近中：バックグラウンド自動サイレント更新を実行...");
      try {
        await refreshAccessTokenSilently();
      } catch (e) {
        console.warn("バックグラウンドトークン更新失敗:", e);
      }
    }
  }, 60000); // 1分ごとにチェック
}

async function initGapiClient() {
  await gapi.client.init({ 
    discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4"]
  });
  
  if (isSessionActive()) {
    const savedToken = localStorage.getItem(STORAGE_KEY_TOKEN);
    if (savedToken && !isTokenNearExpiry()) {
      googleAccessToken = savedToken;
      gapi.client.setToken({ access_token: googleAccessToken });
      updateAuthButtonStatus(true);
      syncFromGoogleSheets();
    } else {
      // 24時間以内であれば初回ロード時にバックグラウンドでサイレント更新して接続
      updateAuthButtonStatus(true);
      try {
        await refreshAccessTokenSilently();
        syncFromGoogleSheets();
      } catch (e) {
        console.warn("初回サイレント更新失敗:", e);
        if (savedToken) {
          googleAccessToken = savedToken;
          gapi.client.setToken({ access_token: googleAccessToken });
          syncFromGoogleSheets();
        } else {
          updateAuthButtonStatus(false);
        }
      }
    }
  } else {
    clearAuthSession();
  }

  startTokenMaintenance();
}

function handleAuthClick() {
  // ユーザーが手動でログインボタンを押した時は新規24時間セッションを開始
  localStorage.setItem(STORAGE_KEY_SESSION_UNTIL, String(Date.now() + SESSION_DURATION_MS));
  tokenClient.requestAccessToken({ prompt: 'select_account' });
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

function parseSheetsResponse(response) {
  const rows = response.result.values;
  if (rows && rows.length > 0) {
    let fullJsonStr = "";
    for (const row of rows) {
      if (row[0]) fullJsonStr += row[0];
    }
    if (fullJsonStr) {
      try {
        return JSON.parse(fullJsonStr);
      } catch (e) {
        console.error("JSON parse error:", e);
        return [];
      }
    }
  }
  return [];
}

async function fetchProjectsFromGoogleSheets() {
  await ensureFreshToken();
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A:A',
    });
    return parseSheetsResponse(response);
  } catch (err) {
    if (err.status === 401 && isSessionActive()) {
      console.log("fetchProjects: 401を検知、サイレント更新してリトライします...");
      try {
        await refreshAccessTokenSilently();
        const retryResp = await gapi.client.sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: 'Sheet1!A:A',
        });
        return parseSheetsResponse(retryResp);
      } catch (retryErr) {
        console.error("リトライ取得失敗:", retryErr);
        throw retryErr;
      }
    }
    throw err;
  }
}

async function saveProjectsToGoogleSheets(projects) {
  await ensureFreshToken();
  const executeSave = async () => {
    const jsonStr = JSON.stringify(projects);
    const chunkSize = 40000;
    const chunks = [];
    for (let i = 0; i < jsonStr.length; i += chunkSize) {
      chunks.push([jsonStr.substring(i, i + chunkSize)]);
    }

    await gapi.client.sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A:A',
    });

    await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sheet1!A1',
      valueInputOption: 'RAW',
      resource: { values: chunks }
    });
  };

  try {
    await executeSave();
  } catch (err) {
    if (err.status === 401 && isSessionActive()) {
      console.log("saveProjects: 401を検知、サイレント更新してリトライします...");
      await refreshAccessTokenSilently();
      await executeSave();
      return;
    }
    throw err;
  }
}

async function syncFromGoogleSheets() {
  try {
    allProjectsInMemory = await fetchProjectsFromGoogleSheets();
    
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
      clearAuthSession();
      alert("ログインの有効期限が切れました。再度「社内DBにログイン」を押してください。");
    } else {
      alert("社内データベースとの同期に失敗しました。対象スプレッドシートのタブ名が「Sheet1」になっているか確認してください。");
    }
  }
}

async function syncToGoogleSheets() {
  if (!googleAccessToken) return;
  
  try {
    await saveProjectsToGoogleSheets(allProjectsInMemory);
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
  const tSales = s.totalSales ?? 0;
  const gm = s.grossMargin  ?? 0;
  const dm = s.directMargin ?? s.opMargin ?? 0;
  const dp = s.directProfit ?? s.opProfit ?? 0;
  const alloc = s.allocation ?? Math.round(tSales * 0.15);
  const bp = s.businessProfit ?? (dp - alloc);
  const bm = s.businessMargin ?? (tSales > 0 ? (bp / tSales) * 100 : 0);
  const pt = plan.projectType || "制作";
  const th = TYPE_THRESHOLDS[pt] || TYPE_THRESHOLDS["制作"];

  const grossCls  = getRateClass(gm, th.gross, s.grossProfit ?? 0);
  const directCls = (th.direct || th.op) ? getRateClass(dm, th.direct || th.op, dp) : "val-neutral";
  const bizCls    = th.business ? getRateClass(bm, th.business, bp) : (tSales > 0 ? (bp < 0 ? "val-critical" : "val-good") : "val-neutral");
  const tagClass  = { "制作": "tag-seisaku", "広告運用": "tag-koukoku", "複合": "tag-fukugo" }[pt] || "tag-seisaku";

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
        <span class="metric-value">${formatCurrency(tSales)}</span>
      </div>
      <div class="metric-box">
        <span class="metric-label">粗利率</span>
        <span class="metric-value ${grossCls}">${formatPercent(gm)}</span>
      </div>
      <div class="metric-box">
        <span class="metric-label">直利率</span>
        <span class="metric-value ${directCls}">${(th.direct || th.op) ? formatPercent(dm) : "-"}</span>
      </div>
      <div class="metric-box">
        <span class="metric-label">事業利益率</span>
        <span class="metric-value ${bizCls}">${th.business ? formatPercent(bm) : (tSales > 0 ? formatPercent(bm) : "-")}</span>
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

let draggedSalesBlock = null;

function setupSalesBlockDragAndDrop(block) {
  block.setAttribute("draggable", "true");

  block.addEventListener("dragstart", (e) => {
    draggedSalesBlock = block;
    block.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", "");
  });

  block.addEventListener("dragend", () => {
    block.classList.remove("dragging");
    document.querySelectorAll(".sales-block").forEach(b => b.classList.remove("drag-over"));
    draggedSalesBlock = null;
    triggerAutoSave();
  });

  block.addEventListener("dragover", (e) => {
    e.preventDefault();
    if (!draggedSalesBlock || draggedSalesBlock === block) return;
    e.dataTransfer.dropEffect = "move";

    const container = document.getElementById("sales_rows_container");
    const siblings = [...container.querySelectorAll(".sales-block:not(.dragging)")];
    const nextSibling = siblings.find(sibling => {
      const box = sibling.getBoundingClientRect();
      return e.clientY < box.top + box.height / 2;
    });

    if (nextSibling) {
      container.insertBefore(draggedSalesBlock, nextSibling);
    } else {
      container.appendChild(draggedSalesBlock);
    }
  });

  block.addEventListener("dragenter", (e) => {
    if (draggedSalesBlock && draggedSalesBlock !== block) {
      block.classList.add("drag-over");
    }
  });

  block.addEventListener("dragleave", () => {
    block.classList.remove("drag-over");
  });

  block.addEventListener("drop", (e) => {
    e.preventDefault();
    block.classList.remove("drag-over");
  });
}

function autoResizeTextarea(el) {
  if (!el) return;
  el.style.height = "auto";
  const borderOffset = (el.offsetHeight - el.clientHeight) || 2;
  el.style.height = Math.max(33, el.scrollHeight + borderOffset) + "px";
}

window.addEventListener("resize", () => {
  document.querySelectorAll(".sales-name, .ext-name").forEach(autoResizeTextarea);
});

function addSalesBlock(name = "", amount = "", extRows = []) {
  const container = document.getElementById("sales_rows_container");
  const block = document.createElement("div");
  block.className = "sales-block";

  block.innerHTML = `
    <div class="sales-row">
      <div class="sales-cell-left">
        <div class="drag-control">
          <span class="drag-handle" title="ドラッグして並び替え">⠿</span>
          <div class="move-btns">
            <button type="button" class="btn-move btn-move-up" title="上に移動">▲</button>
            <button type="button" class="btn-move btn-move-down" title="下に移動">▼</button>
          </div>
        </div>
        <textarea class="table-input input-name sales-name" placeholder="売上項目名" rows="1">${escapeHtml(name)}</textarea>
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
      <div class="sales-cell-profit">
        <span class="cell-profit-title">項目粗利</span>
        <span class="cell-profit-amount sales-profit-val">¥0</span>
        <span class="cell-profit-rate">粗利率: <strong class="sales-profit-rate">0.0%</strong></span>
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

  salesNameInp.addEventListener("input", (e) => { 
    autoResizeTextarea(e.target);
    triggerAutoSave(); 
  });
  salesNameInp.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      salesAmountInp.focus();
    }
  });

  salesAmountInp.addEventListener("input", (e) => { 
    applyLiveCurrencyFormat(e.target); 
    updateExtSubtotal(block);
    triggerAutoSave(); 
  });
  salesAmountInp.addEventListener("blur", (e) => {
    if (e.target.value === "-" || e.target.value === "-¥" || e.target.value === "¥") {
      e.target.value = "";
      updateExtSubtotal(block);
      triggerAutoSave();
    }
  });
  extAddBtn.addEventListener("click", () => addExtRow(block, extContainer));

  // 上下移動ボタンのハンドラ
  block.querySelector(".btn-move-up").addEventListener("click", (e) => {
    e.stopPropagation();
    const prev = block.previousElementSibling;
    if (prev && prev.classList.contains("sales-block")) {
      container.insertBefore(block, prev);
      triggerAutoSave();
    }
  });

  block.querySelector(".btn-move-down").addEventListener("click", (e) => {
    e.stopPropagation();
    const next = block.nextElementSibling;
    if (next && next.classList.contains("sales-block")) {
      container.insertBefore(next, block);
      triggerAutoSave();
    }
  });

  // ドラッグ＆ドロップ機能のセットアップ
  setupSalesBlockDragAndDrop(block);

  container.appendChild(block);
  autoResizeTextarea(salesNameInp);
  extRows.forEach(er => addExtRow(block, extContainer, er.name, er.amount));
  
  updateExtSubtotal(block);

  if (!name && !amount) saveCurrentPlanState();
  calculateAndDisplay();
}

function addExtRow(salesBlock, extContainer, name = "", amount = "") {
  const row = document.createElement("div");
  row.className = "ext-row";
  row.innerHTML = `
    <span class="ext-indent">└</span>
    <textarea class="table-input input-name ext-name" placeholder="外注項目名" rows="1">${escapeHtml(name)}</textarea>
    <input type="text" class="table-input input-amount ext-amount" placeholder="¥0" value="${formatInputCurrency(amount)}">
    <button class="btn btn-danger btn-icon-only ext-del-btn" title="外注行を削除"><svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
  `;

  const extNameInp = row.querySelector(".ext-name");
  const extAmountInp = row.querySelector(".ext-amount");

  extAmountInp.addEventListener("input", (e) => { applyLiveCurrencyFormat(e.target); updateExtSubtotal(salesBlock); triggerAutoSave(); });
  extAmountInp.addEventListener("blur", (e) => {
    if (e.target.value === "-" || e.target.value === "-¥" || e.target.value === "¥") {
      e.target.value = "";
      updateExtSubtotal(salesBlock);
      triggerAutoSave();
    }
  });
  extNameInp.addEventListener("input", (e) => { 
    autoResizeTextarea(e.target);
    triggerAutoSave(); 
  });
  extNameInp.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      extAmountInp.focus();
    }
  });
  row.querySelector(".ext-del-btn").addEventListener("click", () => { row.remove(); updateExtSubtotal(salesBlock); triggerAutoSave(); });

  extContainer.appendChild(row);
  autoResizeTextarea(extNameInp);
  updateExtSubtotal(salesBlock);
  if (!name && !amount) saveCurrentPlanState();
  calculateAndDisplay();
}

function updateExtSubtotal(salesBlock) {
  let extTotal = 0;
  salesBlock.querySelectorAll(".ext-amount").forEach(inp => { extTotal += parseNumber(inp.value); });
  
  if (salesBlock.querySelector(".ext-subtotal-val")) {
    salesBlock.querySelector(".ext-subtotal-val").textContent = formatCurrency(extTotal);
  }

  // 項目ごとの粗利・粗利率計算
  const salesAmount = parseNumber(salesBlock.querySelector(".sales-amount")?.value);
  const profit = salesAmount - extTotal;
  // 売上がプラスの場合のみ粗利率を計算（値引き・マイナス行や0の場合は「-」表示）
  const rate = salesAmount > 0 ? (profit / salesAmount) * 100 : null;

  const profitValEl = salesBlock.querySelector(".sales-profit-val");
  const profitRateEl = salesBlock.querySelector(".sales-profit-rate");
  const profitCell = salesBlock.querySelector(".sales-cell-profit");

  if (profitValEl) profitValEl.textContent = formatCurrency(profit);
  if (profitRateEl) profitRateEl.textContent = formatPercent(rate);

  if (profitCell) {
    profitCell.classList.remove("profit-plus", "profit-minus", "profit-zero");
    if (salesAmount === 0 && extTotal === 0) {
      profitCell.classList.add("profit-zero");
    } else if (profit < 0) {
      profitCell.classList.add("profit-minus");
    } else {
      profitCell.classList.add("profit-plus");
    }
  }
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
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  let str = String(val)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[¥,￥\s]/g, "")
    .replace(/[ー−―－▲△]/g, "-");
  
  const isNegative = str.includes("-");
  const cleaned = str.replace(/[^\d.]/g, "");
  if (!cleaned) return 0;
  const n = parseFloat(cleaned);
  if (isNaN(n)) return 0;
  return isNegative ? -n : n;
}

function formatCurrency(amount) {
  const n = Math.round(parseNumber(amount));
  if (n < 0) {
    return "-¥" + Math.abs(n).toLocaleString("ja-JP");
  }
  return "¥" + n.toLocaleString("ja-JP");
}

function formatInputCurrency(amount) {
  if (amount === "" || amount === null || amount === undefined) return "";
  const n = Math.round(parseNumber(amount));
  if (n === 0 && amount !== 0 && amount !== "0") return "";
  if (n < 0) {
    return "-¥" + Math.abs(n).toLocaleString("ja-JP");
  }
  return "¥" + n.toLocaleString("ja-JP");
}

function applyLiveCurrencyFormat(inputEl) {
  const raw = inputEl.value;
  const selStart = inputEl.selectionStart ?? raw.length;

  const textBeforeCursor = raw.slice(0, selStart);
  const normBefore = textBeforeCursor
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[ー−―－▲△]/g, "-");
  const digitsBefore = (normBefore.match(/\d/g) || []).length;

  const normAll = raw
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[ー−―－▲△]/g, "-");

  const isNegative = normAll.includes("-");
  const digitsOnly = normAll.replace(/[^\d]/g, "");

  if (!digitsOnly) {
    if (isNegative) {
      inputEl.value = "-";
      inputEl.setSelectionRange(1, 1);
    } else {
      inputEl.value = "";
    }
    return;
  }

  const numVal = parseInt(digitsOnly, 10);
  const formattedDigits = numVal.toLocaleString("ja-JP");
  const prefix = isNegative ? "-¥" : "¥";
  const formattedVal = prefix + formattedDigits;

  inputEl.value = formattedVal;

  let targetPos = prefix.length;
  let countedDigits = 0;
  for (let i = prefix.length; i < formattedVal.length; i++) {
    if (/\d/.test(formattedVal[i])) {
      countedDigits++;
      if (countedDigits === digitsBefore) {
        targetPos = i + 1;
        break;
      }
    }
  }

  if (digitsBefore === 0) {
    targetPos = prefix.length;
  }

  inputEl.setSelectionRange(targetPos, targetPos);
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

  // 1. 売上: tSales
  // 2. 粗利 (売上 - 外注費)
  const grossProfit = tSales - tOut;
  const grossMargin = tSales > 0 ? (grossProfit / tSales) * 100 : null;

  // 3. 直接利益 (粗利 - 社内費)
  const directProfit = grossProfit - tInt;
  const directMargin = tSales > 0 ? (directProfit / tSales) * 100 : null;

  // 4. 事業利益 (直利 - 配賦) ※配賦は売上の15%自動計算
  const allocation = Math.round(tSales * 0.15);
  const businessProfit = directProfit - allocation;
  const businessMargin = tSales > 0 ? (businessProfit / tSales) * 100 : null;

  // ① 売上
  const elTotalSales = document.getElementById("summary_total_sales");
  if (elTotalSales) elTotalSales.textContent = formatCurrency(tSales);

  // ② 粗利
  const elTotalOut = document.getElementById("summary_total_outsourcing");
  if (elTotalOut) elTotalOut.textContent = formatCurrency(tOut);
  const elGrossProfit = document.getElementById("summary_gross_profit");
  if (elGrossProfit) elGrossProfit.textContent = formatCurrency(grossProfit);
  const elGrossMargin = document.getElementById("summary_gross_profit_rate");
  if (elGrossMargin) elGrossMargin.textContent = formatPercent(grossMargin);

  // ③ 直接利益
  const elTotalInt = document.getElementById("summary_total_internal");
  if (elTotalInt) elTotalInt.textContent = formatCurrency(tInt);
  const elDirectProfit = document.getElementById("summary_direct_profit");
  if (elDirectProfit) elDirectProfit.textContent = formatCurrency(directProfit);
  const elDirectMargin = document.getElementById("summary_direct_profit_rate");
  if (elDirectMargin) elDirectMargin.textContent = formatPercent(directMargin);

  // 旧ID互換（もし要素が存在していれば更新）
  const elOpProfit = document.getElementById("summary_operating_profit");
  if (elOpProfit) elOpProfit.textContent = formatCurrency(directProfit);
  const elOpMargin = document.getElementById("summary_operating_profit_rate");
  if (elOpMargin) elOpMargin.textContent = formatPercent(directMargin);

  // ④ 事業利益
  const elAlloc = document.getElementById("summary_allocation");
  if (elAlloc) elAlloc.textContent = formatCurrency(allocation);
  const elBusinessProfit = document.getElementById("summary_business_profit");
  if (elBusinessProfit) elBusinessProfit.textContent = formatCurrency(businessProfit);
  const elBusinessMargin = document.getElementById("summary_business_profit_rate");
  if (elBusinessMargin) elBusinessMargin.textContent = formatPercent(businessMargin);

  updateStatusBadges(grossProfit, grossMargin, directProfit, directMargin, businessProfit, businessMargin);
}

function updateStatusBadges(grossProfit, grossMargin, directProfit, directMargin, businessProfit, businessMargin) {
  const pt = document.getElementById("project_type")?.value || "制作";
  const th = TYPE_THRESHOLDS[pt] || TYPE_THRESHOLDS["制作"];
  
  // 粗利判定
  setStatusBadge(document.getElementById("gross_status_badge"), document.getElementById("gross_status_text"), grossProfit, grossMargin, th.gross, "粗利");
  
  // 直利判定（旧営利判定）
  const directBadge = document.getElementById("direct_status_badge") || document.getElementById("op_status_badge");
  const directText = document.getElementById("direct_status_text") || document.getElementById("op_status_text");
  const directTh = th.direct || th.op;
  if (!directTh) {
    if (directBadge) directBadge.className = "status-badge-sm status-neutral d-none";
  } else {
    if (directBadge) {
      directBadge.classList.remove("d-none");
      setStatusBadge(directBadge, directText, directProfit, directMargin, directTh, "直利");
    }
  }

  // 事業利益判定
  const bizBadge = document.getElementById("business_status_badge");
  const bizText = document.getElementById("business_status_text");
  if (bizBadge && bizText) {
    if (th.business === null) {
      bizBadge.className = "status-badge-sm status-neutral d-none";
    } else {
      bizBadge.classList.remove("d-none");
      setStatusBadge(bizBadge, bizText, businessProfit, businessMargin, th.business, "事利");
    }
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
      if (eName || eAmount !== 0) outsourcingRows.push({ name: eName, amount: eAmount });
    });
    if (name || amount !== 0 || outsourcingRows.length > 0) salesRows.push({ name, amount, outsourcingRows });
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

  const grossProfit = tSales - tOut;
  const directProfit = grossProfit - tInt;
  const allocation = Math.round(tSales * 0.15);
  const businessProfit = directProfit - allocation;

  const grossMargin = tSales > 0 ? ((grossProfit) / tSales) * 100 : 0;
  const directMargin = tSales > 0 ? ((directProfit) / tSales) * 100 : 0;
  const businessMargin = tSales > 0 ? ((businessProfit) / tSales) * 100 : 0;

  currentProject.plans[activeIdx] = {
    planName: currentProject.plans[activeIdx]?.planName || `パターン ${activeIdx + 1}`,
    projectType: document.getElementById("project_type").value,
    salesRows, internalRows,
    summary: { 
      totalSales: tSales, 
      totalOutsourcing: tOut, 
      totalInternal: tInt, 
      grossProfit: grossProfit, 
      grossMargin: grossMargin, 
      directProfit: directProfit,
      directMargin: directMargin,
      allocation: allocation,
      businessProfit: businessProfit,
      businessMargin: businessMargin,
      // 過去データ互換用プロパティ
      opProfit: directProfit, 
      opMargin: directMargin 
    }
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
    const tSales = s.totalSales ?? 0;
    const gm = s.grossMargin ?? 0;
    const dm = s.directMargin ?? s.opMargin ?? 0;
    const dp = s.directProfit ?? s.opProfit ?? 0;
    const alloc = s.allocation ?? Math.round(tSales * 0.15);
    const bp = s.businessProfit ?? (dp - alloc);
    const bm = s.businessMargin ?? (tSales > 0 ? (bp / tSales) * 100 : 0);

    const item = document.createElement("div");
    item.className = `comparison-item ${idx === currentProject.activePlanIndex ? "active" : ""}`;
    item.innerHTML = `
      <div class="comp-title">${escapeHtml(plan.planName)}</div>
      <div class="comp-metrics">
        <div class="comp-metric"><span class="comp-lbl">売上:</span> <strong>${formatCurrency(tSales)}</strong></div>
        <div class="comp-metric"><span class="comp-lbl">粗利:</span> <strong class="val-good">${formatPercent(gm)}</strong></div>
        <div class="comp-metric"><span class="comp-lbl">直利:</span> <strong class="val-warning">${formatPercent(dm)}</strong></div>
        <div class="comp-metric"><span class="comp-lbl">事利:</span> <strong class="${bp < 0 ? 'val-critical' : 'val-good'}">${formatPercent(bm)}</strong></div>
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
  requestAnimationFrame(() => {
    document.querySelectorAll(".sales-name, .ext-name").forEach(autoResizeTextarea);
  });
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
  if (!googleAccessToken && isSessionActive()) {
    await ensureFreshToken();
  }
  if (!googleAccessToken) {
    alert("Google連携がされていません。");
    return;
  }
  try {
    let latestProjects = await fetchProjectsFromGoogleSheets();

    const targetProj = latestProjects.find(p => p.id === id);
    if (targetProj && targetProj.isProtected) {
      alert(`案件「${name || "（無題）"}」は保護されているため削除できません。保護を解除してから削除してください。`);
      renderProjectGrid(document.getElementById("search_input").value);
      return;
    }

    if (!confirm(`案件「${name || "（無題）"}」を削除してもよろしいですか？`)) return;

    allProjectsInMemory = latestProjects.filter(p => p.id !== id);
    await saveProjectsToGoogleSheets(allProjectsInMemory);

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
    allProjectsInMemory = await fetchProjectsFromGoogleSheets();
    
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
    if (!googleAccessToken && isSessionActive()) {
      await ensureFreshToken();
    }

    if (googleAccessToken) {
      // 最新のデータを再取得して上書き競合を防ぐ
      let latestProjects = await fetchProjectsFromGoogleSheets();

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
      await saveProjectsToGoogleSheets(allProjectsInMemory);

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
  if (!googleAccessToken && isSessionActive()) {
    await ensureFreshToken();
  }
  if (!googleAccessToken) {
    alert("Google連携がされていません。");
    return;
  }
  try {
    let latestProjects = await fetchProjectsFromGoogleSheets();

    const targetProj = latestProjects.find(p => p.id === id);
    if (targetProj) {
      targetProj.isProtected = !targetProj.isProtected;
      allProjectsInMemory = latestProjects;
      await saveProjectsToGoogleSheets(allProjectsInMemory);

      renderProjectGrid(document.getElementById("search_input").value);
    }
  } catch (err) {
    console.error("保護ステータス変更エラー:", err);
    alert("保護状態の切り替えに失敗しました。通信状況を確認してください。");
  }
}