/**
 * 品牌成交系統提案 — 統一後端
 * 一份 GAS 專案，同時處理：
 *   ① 客戶在提案網頁「產生摘要」時回傳的選擇資料 → submissions 分頁
 *   ② 提案產生器（後臺）的案件草稿雲端儲存/讀取 → cases 分頁
 *   ③ 一鍵發布到 GitHub Pages → publishToGithub action
 *
 * 設定步驟：
 * 1. 開一份新的 Google Sheets
 * 2. 建立兩個分頁，分別命名為 submissions 與 cases（大小寫需一致）
 * 3. 該 Sheets 點「擴充功能 > Apps Script」，貼上這份程式碼，取代原本內容
 * 4. 上方執行一次 setupHeaders() 建立兩個分頁的表頭
 * 5. 部署 > 新增部署作業 > 類型選「網頁應用程式」
 *    - 執行身份：我
 *    - 誰可以存取：任何人
 * 6. 複製部署後的網址：
 *    - 貼到「提案產生器」頁面最上方的「後台管理 GAS 網址」欄位
 *    - 貼到每個客戶提案的「GAS 回傳網址」欄位（產生器會自動寫入該客戶提案檔的 CONFIG.gasEndpoint）
 *
 * GitHub Pages 發布設定（publishToGithub 功能）：
 *    在「專案設定 > 指令碼屬性」手動新增以下三個屬性：
 *    - GITHUB_TOKEN  → 您的 GitHub Personal Access Token（需有 repo 或 contents 寫入權限）
 *    - GITHUB_OWNER  → 您的 GitHub 帳號名稱（例如 angel0973180707）
 *    - GITHUB_REPO   → 目標 repo 名稱（例如 proposal-system）
 */

const SUBMISSIONS_SHEET = "submissions";
const CASES_SHEET = "cases";
const CARD_CASES_SHEET = "card_cases";
const CARD_SELECTIONS_SHEET = "card_selections";

function setupHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let subSheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!subSheet) subSheet = ss.insertSheet(SUBMISSIONS_SHEET);
  subSheet.getRange(1, 1, 1, 10).setValues([[
    "送出時間", "案件代號", "客戶名稱", "填寫人", "主標語",
    "IP頻率設定", "智慧名片保留區塊", "經營前三優先",
    "內容矩陣配置", "補充說明"
  ]]);
  subSheet.getRange(1, 1, 1, 10).setFontWeight("bold");

  let caseSheet = ss.getSheetByName(CASES_SHEET);
  if (!caseSheet) caseSheet = ss.insertSheet(CASES_SHEET);
  caseSheet.getRange(1, 1, 1, 5).setValues([[
    "案件代號", "客戶名稱", "設定內容(JSON)", "更新時間", "交付網址"
  ]]);
  caseSheet.getRange(1, 1, 1, 5).setFontWeight("bold");

  // 智慧名片內容協作系統
  let cardCasesSheet = ss.getSheetByName(CARD_CASES_SHEET);
  if (!cardCasesSheet) cardCasesSheet = ss.insertSheet(CARD_CASES_SHEET);
  cardCasesSheet.getRange(1, 1, 1, 5).setValues([["caseCode", "clientName", "contentJSON", "createdDate", "status"]]);
  cardCasesSheet.getRange(1, 1, 1, 5).setFontWeight("bold");

  let cardSelSheet = ss.getSheetByName(CARD_SELECTIONS_SHEET);
  if (!cardSelSheet) cardSelSheet = ss.insertSheet(CARD_SELECTIONS_SHEET);
  cardSelSheet.getRange(1, 1, 1, 18).setValues([[
    "caseCode", "clientName", "submittedAt",
    "field1_role", "field1_slogan", "field2_services", "field3_story",
    "field4_line", "field4_wechat", "field4_phone", "field4_email", "field4_address",
    "field5_marquee", "field6_youtube", "field6_fb", "field7_cta", "field8_photos",
    "otherNotes"
  ]]);
  cardSelSheet.getRange(1, 1, 1, 18).setFontWeight("bold");
}

/* ---------------- doGet：讀取類請求（列表、載入單一案件），支援 JSONP 避開瀏覽器 CORS 限制 ---------------- */
function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let result;

  try {
    if (action === "listCases") {
      result = { status: "ok", cases: listCases() };
    } else if (action === "loadCase") {
      result = { status: "ok", case: loadCase(e.parameter.caseCode) };
    } else if (action === "listCardCases") {
      result = { status: "ok", cases: listCardCases_() };
    } else if (action === "getCase") {
      result = { status: "ok", data: getCardCase_(e.parameter.caseCode) };
    } else if (action === "getSelection") {
      result = { status: "ok", data: getCardSelection_(e.parameter.caseCode) };
    } else {
      result = { status: "error", message: "未知的 action" };
    }
  } catch (err) {
    result = { status: "error", message: err.message };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function listCases() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASES_SHEET);
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      caseCode: data[i][0],
      clientName: data[i][1],
      updatedAt: data[i][3],
      deliveryUrl: data[i][4] || ""
    });
  }
  return out;
}

function loadCase(caseCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASES_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === caseCode) {
      return JSON.parse(data[i][2]);
    }
  }
  return null;
}

/* ---------------- doPost：寫入類請求（提交摘要、儲存/刪除案件、發布到 GitHub） ---------------- */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || "submitProposal"; // 未帶 action 視為舊版客戶提交摘要
    let result = { status: "ok" };

    if (action === "submitProposal") {
      submitProposal(data);
    } else if (action === "saveCase") {
      saveCase(data);
    } else if (action === "deleteCase") {
      deleteCase(data.caseCode);
    } else if (action === "publishToGithub") {
      result = publishToGithub(data);
    } else if (action === "createCase") {
      createCardCase_(data);
    } else if (action === "submitSelection") {
      result = submitCardSelection_(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function submitProposal(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUBMISSIONS_SHEET);

  const ipFreqText = Object.entries(data.ipFreq || {}).map(([k, v]) => k + ":" + v).join("、");
  const blocksText = (data.blocks || []).join("、");
  const priorityText = (data.priorities || []).join("、");
  const matrixText = (data.matrix || []).map(m => m.label + " " + m.val + "%").join("、");

  sheet.appendRow([
    new Date(),
    data.caseCode || "",
    data.clientName || "",
    data.contact || "",
    data.slogan || "",
    ipFreqText,
    blocksText,
    priorityText,
    matrixText,
    data.note || ""
  ]);
}

function saveCase(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASES_SHEET);
  const values = sheet.getDataRange().getValues();
  const configJson = JSON.stringify(data.config || {});
  const deliveryUrl = data.deliveryUrl || "";
  const now = new Date();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.caseCode) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[data.caseCode, data.clientName || "", configJson, now, deliveryUrl]]);
      return;
    }
  }
  sheet.appendRow([data.caseCode, data.clientName || "", configJson, now, deliveryUrl]);
}

function deleteCase(caseCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CASES_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === caseCode) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

/* ================== GitHub Pages 發布功能 ================== */

/**
 * 把產生的提案 HTML 推送到 GitHub repo（proposals/{caseCode}.html）
 * Token 等機密從 Script Properties 讀取，不寫死在程式碼裡
 *
 * 需在「專案設定 > 指令碼屬性」設定：
 *   GITHUB_TOKEN  → Personal Access Token（scope: repo 或 contents:write）
 *   GITHUB_OWNER  → GitHub 帳號名稱
 *   GITHUB_REPO   → repo 名稱
 */
function publishToGithub(data) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER');
  const repo  = props.getProperty('GITHUB_REPO');

  if (!token || !owner || !repo) {
    return {
      status: 'error',
      message: '請先在 Apps Script > 專案設定 > 指令碼屬性設定 GITHUB_TOKEN、GITHUB_OWNER、GITHUB_REPO'
    };
  }

  const caseCode    = (data.caseCode || '').trim();
  const htmlContent = data.htmlContent || '';
  if (!caseCode || !htmlContent) {
    return { status: 'error', message: '缺少 caseCode 或 htmlContent' };
  }

  const filePath = 'proposals/' + caseCode + '.html';
  const apiUrl   = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + filePath;
  const headers  = {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent': 'ProposalGenerator-GAS'
  };

  // GET：先確認檔案是否已存在，若存在需取得 sha 才能更新
  let sha = null;
  try {
    const getResp = UrlFetchApp.fetch(apiUrl, {
      method: 'GET',
      headers: headers,
      muteHttpExceptions: true
    });
    if (getResp.getResponseCode() === 200) {
      sha = JSON.parse(getResp.getContentText()).sha;
    }
  } catch (e) {
    // 檔案不存在是正常情況，sha 維持 null
  }

  // PUT：上傳（新增或更新）
  const putBody = {
    message: '提案發布：' + caseCode,
    content: Utilities.base64Encode(htmlContent, Utilities.Charset.UTF_8)
  };
  if (sha) putBody.sha = sha; // 更新現有檔案時必須帶 sha

  const putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'PUT',
    headers: headers,
    payload: JSON.stringify(putBody),
    muteHttpExceptions: true
  });

  const code = putResp.getResponseCode();
  if (code === 200 || code === 201) {
    return {
      status: 'ok',
      url: 'https://' + owner + '.github.io/' + repo + '/proposals/' + caseCode + '.html'
    };
  } else {
    const errText = putResp.getContentText();
    return {
      status: 'error',
      message: 'GitHub API 回應 ' + code + '：' + errText.slice(0, 300)
    };
  }
}

/* ================== 智慧名片內容協作系統 ================== */

function listCardCases_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CARD_CASES_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      caseCode: String(data[i][0]),
      clientName: String(data[i][1]),
      createdDate: data[i][3] ? String(data[i][3]).slice(0, 10) : '',
      status: data[i][4] || '待填寫'
    });
  }
  return out;
}

function getCardCase_(caseCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CARD_CASES_SHEET);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(caseCode)) {
      return {
        caseCode: String(data[i][0]),
        clientName: String(data[i][1]),
        contentJSON: String(data[i][2]),
        createdDate: data[i][3] ? String(data[i][3]).slice(0, 10) : '',
        status: data[i][4] || '待填寫'
      };
    }
  }
  return null;
}

function getCardSelection_(caseCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CARD_SELECTIONS_SHEET);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  let lastRow = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(caseCode)) lastRow = data[i];
  }
  if (!lastRow) return null;
  const result = {};
  headers.forEach((h, i) => { result[h] = lastRow[i] !== undefined ? String(lastRow[i]) : ''; });
  return result;
}

function createCardCase_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CARD_CASES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CARD_CASES_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([["caseCode", "clientName", "contentJSON", "createdDate", "status"]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][0]) === String(data.caseCode)) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[
        data.caseCode, data.clientName, data.contentJSON, existing[i][3], existing[i][4] || '待填寫'
      ]]);
      return;
    }
  }
  sheet.appendRow([data.caseCode, data.clientName, data.contentJSON, new Date(), '待填寫']);
}

function submitCardSelection_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let selSheet = ss.getSheetByName(CARD_SELECTIONS_SHEET);
  if (!selSheet) {
    selSheet = ss.insertSheet(CARD_SELECTIONS_SHEET);
    selSheet.getRange(1, 1, 1, 18).setValues([[
      "caseCode", "clientName", "submittedAt",
      "field1_role", "field1_slogan", "field2_services", "field3_story",
      "field4_line", "field4_wechat", "field4_phone", "field4_email", "field4_address",
      "field5_marquee", "field6_youtube", "field6_fb", "field7_cta", "field8_photos", "otherNotes"
    ]]);
    selSheet.getRange(1, 1, 1, 18).setFontWeight("bold");
  }
  selSheet.appendRow([
    data.caseCode || '', data.clientName || '', new Date(),
    data.field1_role || '', data.field1_slogan || '',
    data.field2_services || '', data.field3_story || '',
    data.field4_line || '', data.field4_wechat || '',
    data.field4_phone || '', data.field4_email || '',
    data.field4_address || '', data.field5_marquee || '',
    data.field6_youtube || '', data.field6_fb || '',
    data.field7_cta || '', data.field8_photos || '',
    data.otherNotes || ''
  ]);
  const casesSheet = ss.getSheetByName(CARD_CASES_SHEET);
  if (casesSheet) {
    const casesData = casesSheet.getDataRange().getValues();
    for (let i = 1; i < casesData.length; i++) {
      if (String(casesData[i][0]) === String(data.caseCode)) {
        casesSheet.getRange(i + 1, 5).setValue('已送出');
        break;
      }
    }
  }
  return { status: 'ok', message: '選擇已儲存' };
}
