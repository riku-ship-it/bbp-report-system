/**
 * BBP 回報系統後端代碼 - 旗艦穩定版
 * 包含：狀態(進行中)、優先序(1-嚴重)
 * 更新：週期截止時間改為每週三 19:00 台灣時間
 * 更新：新案件 / 使用者補充 自動寄信通知管理者
 * 更新：新增知識產出（BBP知識技能樹）寫入 / 讀取功能
 */
const SHEET_NAME           = "工作表1";         // 回報中心工作表名稱
const KNOWLEDGE_SHEET_NAME = "BBP知識技能樹";   // 知識產出工作表名稱（請依實際分頁名稱修改）
const ADMIN_EMAIL          = "riku@hcatwn.com"; // 管理者 Email（通知收件人，同時也排除自己的回報）
const NOTIFY_LABEL         = "#回報中心";        // 信件主旨標籤

/**
 * 測試用：手動執行這個函式確認寄信功能正常
 * 步驟：上方下拉選單切換到 testEmail → 點執行
 */
function testEmail() {
  MailApp.sendEmail({
    to:      ADMIN_EMAIL,
    subject: NOTIFY_LABEL + " 【測試】通知功能測試",
    body:    "這是測試信，如果收到代表寄信功能正常。"
  });
}

/**
 * doGet: 處理讀取請求（回報中心主表）
 */
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * doPost: 處理寫入請求 (使用者回報 / 管理者回覆 / 知識產出)
 */
function doPost(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const now  = new Date();

    // ── 知識產出分流 ──────────────────────────────────────────────
    // [修正] 原本只判斷 data.type === 'knowledge'
    //        前端實際送出的是 data.action === 'knowledge'，加入 || 相容兩種寫法
    if (data.action === 'knowledge' || data.type === 'knowledge') {
      return handleKnowledge(ss, data, now);
    }

    // ── 原有回報中心邏輯 ──────────────────────────────────────────
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // 1. 核心參數處理
    const ticketId = data.id || Utilities.getUuid();
    const cycleId  = getCycleId(now);
    const sender   = data.sender || "使用者";

    // 2. 狀態與優先序
    const status   = data.status   || "待處理";
    const priority = data.priority || "2-優先";

    const email       = data.email       || "";
    const type        = data.type        || "";
    const title       = data.title       || "";
    const description = data.description || "";

    // 3. 判斷是否為新案件
    const isNewTicket = !data.id;

    // 4. 執行寫入 (appendRow)
    // 欄位順序：ID, 時間, Email, 類別, 標題, 描述, 發言人, 狀態, 更新時間, 優先序, 週期
    sheet.appendRow([
      ticketId,     // A
      now,          // B
      email,        // C
      type,         // D
      title,        // E
      description,  // F
      sender,       // G
      status,       // H
      now,          // I
      priority,     // J
      cycleId       // K
    ]);

    // 5. 同步舊紀錄
    syncOldStatus(sheet, ticketId, status, priority);

    // 6. 通知邏輯
    //    條件A：新案件，且回報者不是管理者自己
    //    條件B：使用者補充（sender = 使用者，且是既有案件）
    const senderIsAdmin  = email.trim().toLowerCase() === ADMIN_EMAIL.toLowerCase();
    const isUserFollowUp = !isNewTicket && sender === "使用者";

    if ((isNewTicket && !senderIsAdmin) || isUserFollowUp) {
      sendNotification(isNewTicket, title, email, type, priority, ticketId);
    }

    return ContentService.createTextOutput(JSON.stringify({
      "result":   "success",
      "id":       ticketId,
      "cycle":    cycleId,
      "status":   status,
      "priority": priority
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result":  "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ════════════════════════════════════════════════════════════════════
//  知識產出功能
// ════════════════════════════════════════════════════════════════════

/**
 * [修正] 狀態值正規化：前端送英文代碼，GAS 統一存中文標籤
 *
 *   前端值   →  試算表顯示
 *   'idea'   →  '想法'
 *   'draft'  →  '草稿'
 *   'wip'    →  '進行中'
 *   'pub'    →  '已發布'
 *   其他     →  原值（已是中文則直接存入，向下相容）
 */
function normalizeKnowledgeStatus(status) {
  const map = { idea: '想法', draft: '草稿', wip: '進行中', pub: '已發布' };
  return map[status] || status;
}

/**
 * handleKnowledge: 知識產出的讀取 / 新增 / 更新
 *
 * 工作表欄位（共 8 欄）：
 *   A: ID          B: 建立時間  C: 發布時間
 *   D: 標題        E: 解決的問題  F: 備注
 *   G: 附件網址    H: 狀態
 *
 * 前端傳入 JSON 欄位說明（已修正欄位名稱對應）：
 *   data.action      → 'knowledge'（路由用）
 *   data.id          → 前端節點 ID（如 'n1'），作為 A 欄主鍵；有值=更新，查無此 ID=新增
 *   data.createdAt   → 節點建立時間（ISO 字串），新增時寫入 B 欄
 *   data.title       → 標題（D 欄）
 *   data.prob        → 解決的問題（E 欄）← [修正] 原本期待 data.problem
 *   data.note        → 備注（F 欄）
 *   data.url         → 附件網址（G 欄）← [修正] 原本期待 data.attachUrl
 *   data.status      → 英文代碼，正規化後存入 H 欄；'pub' 時自動補寫發布時間（C 欄）
 *   data.publishedAt → 前端記錄的發布時間（ISO 字串），優先使用；沒有時以 now 補上
 */
function handleKnowledge(ss, data, now) {
  const sheet = ss.getSheetByName(KNOWLEDGE_SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      result:  'error',
      message: '找不到知識產出工作表：' + KNOWLEDGE_SHEET_NAME
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ── GET：回傳整張工作表資料 ──────────────────────────────────────
  if (data.action === 'get') {
    const allData = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(allData))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── WRITE：欄位讀取（含向下相容舊欄位名稱）────────────────────────
  const title     = data.title                   || '';
  const problem   = data.prob    || data.problem  || '';   // [修正] 前端送 prob
  const note      = data.note                    || '';
  const attachUrl = data.url     || data.attachUrl || '';  // [修正] 前端送 url
  const statusRaw = data.status                  || '';
  const status    = normalizeKnowledgeStatus(statusRaw);   // [修正] 英文→中文

  // [修正] 前端永遠帶 data.id（節點 ID），直接用作主鍵搜尋
  //        原本 getNextKnowledgeId 仍保留供未來其他呼叫方使用
  const existingRow = data.id ? findKnowledgeRow(sheet, data.id) : -1;

  if (existingRow === -1) {
    // ── 新增 ────────────────────────────────────────────────────────
    // [修正] 直接使用前端的 data.id 作為 A 欄值，確保下次能以相同 id 找到並更新
    //        若前端未帶 id（理論上不會），才退而求其次用自動編號
    const rowId = data.id || getNextKnowledgeId(sheet);

    // 建立時間：優先採用前端記錄的 createdAt，否則用 now
    const createdTime = data.createdAt ? new Date(data.createdAt) : now;

    // 發布時間：狀態為「已發布」時才寫入
    const publishTime = (status === '已發布') ? (data.publishedAt ? new Date(data.publishedAt) : now) : '';

    sheet.appendRow([
      rowId,         // A: ID
      createdTime,   // B: 建立時間
      publishTime,   // C: 發布時間（非已發布時留空）
      title,         // D: 標題
      problem,       // E: 解決的問題
      note,          // F: 備注
      attachUrl,     // G: 附件網址
      status         // H: 狀態（中文）
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      id:     rowId,
      action: 'created'
    })).setMimeType(ContentService.MimeType.JSON);

  } else {
    // ── 更新 ────────────────────────────────────────────────────────
    const rowData             = sheet.getRange(existingRow, 1, 1, 8).getValues()[0];
    const existingPublishTime = rowData[2]; // C 欄：發布時間

    // 發布時間：僅在狀態第一次變成「已發布」且尚未記錄過時才寫入
    let newPublishTime = existingPublishTime;
    if (status === '已發布' && !existingPublishTime) {
      newPublishTime = data.publishedAt ? new Date(data.publishedAt) : now;
    }

    // B 欄（建立時間）保持不動；只更新 C~H
    sheet.getRange(existingRow, 3).setValue(newPublishTime); // C: 發布時間
    sheet.getRange(existingRow, 4).setValue(title);          // D: 標題
    sheet.getRange(existingRow, 5).setValue(problem);        // E: 解決的問題
    sheet.getRange(existingRow, 6).setValue(note);           // F: 備注
    sheet.getRange(existingRow, 7).setValue(attachUrl);      // G: 附件網址
    sheet.getRange(existingRow, 8).setValue(status);         // H: 狀態

    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      id:     data.id,
      action: 'updated'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * getNextKnowledgeId: 計算下一個 K-XXX 流水號
 * 掃描整欄 ID，取最大號碼 +1；工作表為空時從 K-001 開始
 * （保留供日後其他呼叫情境使用）
 */
function getNextKnowledgeId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'K-001';
  const ids    = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let   maxNum = 0;
  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0]);
    if (/^K-\d+$/.test(id)) {
      const num = parseInt(id.substring(2), 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return 'K-' + String(maxNum + 1).padStart(3, '0');
}

/**
 * findKnowledgeRow: 依 ID 找到知識產出工作表的列號（1-indexed）
 * 找不到回傳 -1
 */
function findKnowledgeRow(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // +2 = 跳過標頭列，且 0-indexed → 1-indexed
  }
  return -1;
}

// ════════════════════════════════════════════════════════════════════
//  原有輔助函式（保持不動）
// ════════════════════════════════════════════════════════════════════

/**
 * 寄送通知信給管理者
 */
function sendNotification(isNew, title, email, type, priority, ticketId) {
  const eventLabel = isNew ? "新案件回報" : "使用者補充說明";
  const subject    = NOTIFY_LABEL + " 【" + eventLabel + "】" + title;
  let body = "";
  if (isNew) {
    body = [
      "有新的案件回報，請至後台查看。",
      "",
      "─────────────────────",
      "標題：" + title,
      "分類：" + type,
      "優先序：" + priority,
      "回報者：" + email,
      "案件 ID：" + ticketId,
      "─────────────────────",
      "",
      "請登入 BBP 回報中心後台進行處理。"
    ].join("\n");
  } else {
    body = [
      "已退回的案件收到使用者補充說明，請至後台查看。",
      "",
      "─────────────────────",
      "標題：" + title,
      "回報者：" + email,
      "案件 ID：" + ticketId,
      "─────────────────────",
      "",
      "請登入 BBP 回報中心後台進行處理。"
    ].join("\n");
  }
  MailApp.sendEmail({
    to:      ADMIN_EMAIL,
    subject: subject,
    body:    body
  });
}

/**
 * 輔助：計算週期編號
 * 規則：每週三 19:00 台灣時間（UTC+8）為截止點
 */
function getCycleId(date) {
  const twOffset = 8 * 60 * 60 * 1000;
  const tw       = new Date(date.getTime() + twOffset);
  const day    = tw.getUTCDay();
  const hour   = tw.getUTCHours();
  const minute = tw.getUTCMinutes();
  const pastCutoff = (day === 3 && (hour > 19 || (hour === 19 && minute > 0)))
                     || day === 4 || day === 5 || day === 6 || day === 0;
  const base = new Date(tw);
  let diffToWed;
  if (!pastCutoff) {
    diffToWed = 3 - day;
  } else {
    diffToWed = (3 - day + 7) % 7 || 7;
  }
  base.setUTCDate(tw.getUTCDate() + diffToWed);
  const year        = base.getUTCFullYear();
  const jan4        = new Date(Date.UTC(year, 0, 4));
  const jan4Day     = jan4.getUTCDay() || 7;
  const firstMonday = new Date(jan4);
  firstMonday.setUTCDate(jan4.getUTCDate() - (jan4Day - 1));
  const weekNum = Math.floor((base.getTime() - firstMonday.getTime()) / (7 * 24 * 60 * 60 * 1000)) + 1;
  return `${year}-W${String(weekNum).padStart(2, '0')}`;
}

/**
 * 輔助：狀態同步（回報中心主表用）
 */
function syncOldStatus(sheet, ticketId, newStatus, newPriority) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range  = sheet.getRange(2, 1, lastRow - 1, 10);
  const values = range.getValues();
  const now    = new Date();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === ticketId) {
      sheet.getRange(i + 2, 8).setValue(newStatus);
      sheet.getRange(i + 2, 10).setValue(newPriority);
      sheet.getRange(i + 2, 9).setValue(now);
    }
  }
}
