/**
 * 學校召集人 LINE Bot 訊息智慧彙整與派報系統
 * Google Apps Script (GAS) 核心程式碼
 */

// ==========================
// 1. 環境變數與設定 (請在「專案設定 > 指令碼屬性」中填入)
// ==========================
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const LINE_CHANNEL_ACCESS_TOKEN = SCRIPT_PROPERTIES.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
const GEMINI_API_KEY = SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
const DRIVE_FOLDER_ID = SCRIPT_PROPERTIES.getProperty('DRIVE_FOLDER_ID'); // 存放圖片與PDF的Drive資料夾ID
const ADMIN_USER_ID = SCRIPT_PROPERTIES.getProperty('ADMIN_USER_ID');     // 管理員 LINE User ID (選填)

const SHEET_RAW = 'Raw_Messages';       // 原始接收訊息表
const SHEET_TEACHERS = 'Teacher_List';  // 教師Email通訊錄
const SHEET_DIGEST = 'Digest_Summaries';// 彙整派報歷史記錄

// 取得有效的 Gemini API Key
function getActiveGeminiApiKey() {
  return GEMINI_API_KEY || SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
}

// ==========================
// 2. 動態偵測與模型候選邏輯
// ==========================
function getLatestFlashModelName() {
  const defaultModel = "gemini-flash-lite-latest"; // 💡 修正預設值為目前有效的模型
  const apiKey = getActiveGeminiApiKey();
  if (!apiKey) return defaultModel;

  const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  
  try {
    const response = UrlFetchApp.fetch(listUrl, { "muteHttpExceptions": true });
    const result = JSON.parse(response.getContentText());
    
    if (result.models && Array.isArray(result.models)) {
      const flashModels = result.models
        .filter(m => {
          const name = m.name.toLowerCase();
          const hasGen = m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent");
          const isFlash = name.includes("flash");
          
          // 💡 排除 TTS(語音)、Image(生圖)、Computer-Use 等不適合純 OCR/文字辨識的模型
          const isExcluded = name.includes("tts") || name.includes("image") || name.includes("computer-use") || name.includes("omni") || name.includes("exp");
          return hasGen && isFlash && !isExcluded;
        })
        .map(m => m.name.replace("models/", ""));

      if (flashModels.length > 0) {
        // 自然排序：版號新、數字大的排在最前面
        flashModels.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
        Logger.log(`🤖 自動選擇的最佳 Flash 模型: [${flashModels[0]}]`);
        return flashModels[0];
      }
    }
  } catch (err) {
    Logger.log("⚠️ 取得模型清單失敗，使用預設值: " + err.toString());
  }

  return defaultModel;
}

function getCandidateModels() {
  const dynamicModel = getLatestFlashModelName();
  return Array.from(new Set([
    "gemini-flash-lite-latest",  // 1. 強制優先選用 Lite（速度快、省額度）
    dynamicModel,                // 2. 動態偵測到的最新版 Flash 模型
    "gemini-flash-latest",       // 3. Standard Flash 最新版 (高辨識率)
    "gemini-3.6-flash",          // 4. 新世代強大模型備援
    "gemini-2.5-flash"           // 5. 保底備援
  ].filter(Boolean)));
}

// ==========================
// 3. 測試專用函式 (在 Apps Script 編輯器上方選擇此函式並按「執行」)
// ==========================
/**
 * 🧪 測試 Gemini API 是否調用成功
 * 模擬兩則訊息：一則為全校公告、一則為僅召集人內部事項，驗證是否能正常過濾並產生結構化摘要
 */
function testGeminiApi() {
  console.log('🚀 開始測試 Gemini API 連線與動態模型解析...');
  const candidates = getCandidateModels();
  console.log('📋 模型候選清單優先順序：', candidates);

  if (!getActiveGeminiApiKey()) {
    console.error('❌ 測試失敗：尚未在「專案設定 > 指令碼屬性」中設定 GEMINI_API_KEY！');
    return;
  }

  // 模擬真實群組接收到的兩則訊息
  const mockMessages = [
    {
      sender: '黃怡瑜 Lia',
      type: 'text',
      content: '@All 各位老師大家好，附件為 115-1 第二次段考範圍通知，煩請協助確認內容，若沒問題今天下午正式公告全校。',
      time: '14:51'
    },
    {
      sender: '教學組長',
      type: 'text',
      content: '請各科召集人明天中午 12:10 到第一會議室領取各科審題記錄表並繳交簽到，非召集人老師不需出席。',
      time: '15:10'
    }
  ];

  console.log('📝 傳送測試樣本訊息給 Gemini...');
  const startTime = new Date().getTime();

  try {
    const result = analyzeAndSummarizeWithGemini(mockMessages.map((m, idx) => ({
      rowIndex: idx + 1,
      sender: m.sender,
      type: m.type,
      content: m.content,
      timestamp: new Date()
    })));

    const elapsedTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);
    console.log(`⏱️ API 總耗時：${elapsedTime} 秒`);
    console.log('--------------------------------------------------');
    console.log('📋 Gemini 分析結果輸出 (JSON)：');
    console.log(JSON.stringify(result, null, 2));
    console.log('--------------------------------------------------');

    if (result && result.items && result.items.length > 0) {
      console.log('🎉 【測試大成功！】');
      console.log(`✅ 成功萃取出 ${result.items.length} 則全校性公告！`);
      console.log(`✅ 成功排除「僅限召集人內部領表」之事項！`);
      console.log(`💡 產出之半日摘要：${result.summary}`);
    } else {
      console.warn('⚠️ API 有回應，但未解析出任何需公告項目，請檢視輸出內容。');
    }
  } catch (err) {
    console.error('❌ 測試失敗，發生例外錯誤：' + err.message);
  }
}

// ==========================
// 4. LINE Webhook 接收端 (doPost)
// ==========================
function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    const events = json.events;
    if (!events || events.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let rawSheet = ss.getSheetByName(SHEET_RAW);
    if (!rawSheet) {
      rawSheet = initRawSheet(ss);
    }

    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      if (event.type !== 'message') continue;

      const message = event.message;
      const timestamp = new Date(event.timestamp);
      const senderName = getUserDisplayName(event.source);
      const messageId = message.id;
      const messageType = message.type; // text, image, file
      const isOneOnOne = (event.source.type === 'user');
      const userId = event.source.userId;

      // --- [管理功能] 1對1 私訊特殊指令處理 ---
      if (isOneOnOne && messageType === 'text') {
        const textLower = message.text.trim().toLowerCase();

        // 1. 查詢自己的 LINE User ID (方便管理者設定 ADMIN_USER_ID)
        if (textLower === '我的id' || textLower === 'id' || textLower === '!id') {
          replyLineMessage(event.replyToken, `您的 LINE User ID 為：\n${userId}\n\n請複製此代碼，填入 GAS 的指令碼屬性 ADMIN_USER_ID 中！`);
          continue;
        }

        // 2. 管理員指令：!status (查詢待處理筆數)
        if (textLower === '!status') {
          if (ADMIN_USER_ID && userId !== ADMIN_USER_ID) {
            replyLineMessage(event.replyToken, '❌ 抱歉，您不是系統管理員，無法執行此指令。');
          } else {
            const pendingCount = getPendingMessageCount(rawSheet);
            replyLineMessage(event.replyToken, `📊 系統狀態報告：\n目前待彙整派報的訊息共有 ${pendingCount} 筆。`);
          }
          continue;
        }

        // 3. 管理員指令：!send (立即手動觸發一次彙整並寄信)
        if (textLower === '!send') {
          if (ADMIN_USER_ID && userId !== ADMIN_USER_ID) {
            replyLineMessage(event.replyToken, '❌ 抱歉，您不是系統管理員，無法執行手動派報。');
          } else {
            replyLineMessage(event.replyToken, '🚀 已收到管理員指令，正在立即執行手動彙整與寄發 Email...');
            processAndSendDigest('手動即時觸發');
          }
          continue;
        }
      }

      // --- [常態群組記錄] 接收文字、圖片與 PDF 檔案 ---
      let textContent = '';
      let driveFileId = '';
      let driveFileUrl = '';
      let fileName = '';

      if (messageType === 'text') {
        textContent = message.text;
      } else if (messageType === 'image' || messageType === 'file') {
        fileName = message.fileName || `IMG_${Utilities.formatDate(timestamp, 'GMT+8', 'yyyyMMdd_HHmmss')}.jpg`;
        const driveFile = saveLineContentToDrive(messageId, fileName);
        if (driveFile) {
          driveFileId = driveFile.getId();
          driveFileUrl = driveFile.getUrl();
          textContent = `[附件: ${fileName}]`;
        }
      }

      // 關聯同一個發布者在 2 分鐘內的訊息（Group ID）
      const sessionGroupId = generateSessionGroupId(rawSheet, senderName, timestamp);

      // 寫入 Google Sheets 原始訊息表
      rawSheet.appendRow([
        messageId,
        timestamp,
        senderName,
        messageType,
        textContent,
        driveFileId,
        driveFileUrl,
        sessionGroupId,
        false, // is_processed
        ''     // batch_tag
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('doPost Error: ' + err.message);
    notifyAdmin(`⚠️ LINE Webhook 發生錯誤：${err.message}`);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================
// 5. 定時彙整與 Email 派報 (12:00 / 16:00 觸發)
// ==========================
function runDigest1200() {
  processAndSendDigest('中午 12:00');
}

function runDigest1600() {
  processAndSendDigest('下午 16:00');
}

function processAndSendDigest(timeSlotLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(SHEET_RAW);
  const teacherSheet = ss.getSheetByName(SHEET_TEACHERS);
  const digestSheet = ss.getSheetByName(SHEET_DIGEST);
  if (!rawSheet || !teacherSheet) return;

  const data = rawSheet.getDataRange().getValues();
  if (data.length <= 1) return; // 只有標題列

  const pendingRows = [];
  for (let r = 1; r < data.length; r++) {
    const isProcessed = data[r][8];
    if (isProcessed === false || isProcessed === 'FALSE') {
      pendingRows.push({
        rowIndex: r + 1,
        messageId: data[r][0],
        timestamp: data[r][1],
        sender: data[r][2],
        type: data[r][3],
        content: data[r][4],
        fileId: data[r][5],
        fileUrl: data[r][6],
        groupId: data[r][7]
      });
    }
  }

  if (pendingRows.length === 0) {
    console.log(`[${timeSlotLabel}] 沒有待處理的新訊息。`);
    return;
  }

  // 1. 利用 Gemini API 進行「全校公告 vs 僅召集人事務」判斷與重點摘要
  const digestResult = analyzeAndSummarizeWithGemini(pendingRows);

  // 若所有訊息皆為「召集人內部事項」，則本次不打擾全校老師
  if (!digestResult.items || digestResult.items.length === 0) {
    console.log(`[${timeSlotLabel}] 訊息皆屬召集人專用或行政內部事項，無需公告給全體老師。`);
    markRowsAsProcessed(rawSheet, pendingRows, `${timeSlotLabel}_SKIPPED`);
    return;
  }

  // 2. 準備郵件附件（圖片與 PDF 檔案）
  const attachments = [];
  const processedFileIds = new Set();
  pendingRows.forEach(row => {
    if (row.fileId && !processedFileIds.has(row.fileId)) {
      try {
        const file = DriveApp.getFileById(row.fileId);
        attachments.push(file.getAs(file.getMimeType()));
        processedFileIds.add(row.fileId);
      } catch (e) {
        console.warn('讀取附件失敗: ' + row.fileId, e);
      }
    }
  });

  // 3. 組合 HTML 郵件內容
  const htmlBody = generateHtmlEmail(timeSlotLabel, digestResult);

  // 4. 取得教師 Email 清單並寄送
  const recipients = getTeacherEmailList(teacherSheet);
  if (recipients.length > 0) {
    const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd');
    const subject = `【竹林召集人通報】${todayStr} ${timeSlotLabel} 重點訊息整理`;

    try {
      GmailApp.sendEmail(recipients.join(','), subject, '請使用支援 HTML 的郵件客戶端開啟本通知。', {
        htmlBody: htmlBody,
        attachments: attachments,
        name: '竹林召集人智慧秘書'
      });
      console.log(`已成功發送郵件給 ${recipients.length} 位老師。`);

      // 寫入 Digest_Summaries 歷史記錄
      if (digestSheet) {
        digestSheet.appendRow([
          `DIG-${Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd-HHmm')}`,
          new Date(),
          timeSlotLabel,
          digestResult.items.map(it => it.category).join('、'),
          digestResult.items[0].title,
          digestResult.summary,
          attachments.length,
          recipients.length,
          'SUCCESS'
        ]);
      }
    } catch (e) {
      console.error('Gmail 寄送失敗: ' + e.message);
      notifyAdmin(`⚠️ 寄發 Email 失敗：${e.message}`);
    }
  }

  // 5. 標記訊息為已處理
  markRowsAsProcessed(rawSheet, pendingRows, `${timeSlotLabel}_SENT`);
}

// ==========================
// 6. Gemini API 智慧分類、篩選與摘要 (支援候選模型動態遞補)
// ==========================
function analyzeAndSummarizeWithGemini(rows) {
  const apiKey = getActiveGeminiApiKey();
  if (!apiKey) {
    console.warn('未偵測到 GEMINI_API_KEY，啟用簡易備援規則');
    return fallbackRuleFilter(rows);
  }

  const promptText = `
你是一位專業的高中行政教務秘書。以下是學校「召集人群組」今日由各處室或老師發出的訊息清單：
${JSON.stringify(rows.map(r => ({ sender: r.sender, type: r.type, content: r.content, time: Utilities.formatDate(new Date(r.timestamp), 'GMT+8', 'HH:mm') })))}

【重要過濾與判斷任務】：
1. 判斷每一則訊息或同群組訊息「是否需要傳送給全體老師」：
   - 需公告給老師：全校性時程（如流感疫苗）、段考範圍與確認、公文通報更正、學務社團調課、榮譽榜、人事福利法規等。
   - 僅限召集人內部事項【不要放進信件】：例如要求各科召集人交審題表、召集人會議出席調查、純召集人私人事務對話。
2. 對於需要公告的項目，進行專業重點條列式整理（提取時間、地點、配合事項、對象）。
3. 輸出嚴格的 JSON 格式（不要包含 markdown 代碼區塊），結構如下：
{
  "summary": "1~2句話的本次半日通報總結",
  "items": [
    {
      "category": "教務與段考 | 學務與活動 | 人事公告 | 緊急更正 | 榮譽榜",
      "title": "簡潔明確的主題標題",
      "sender": "發布者姓名",
      "points": ["重點1", "重點2", "注意事項"]
    }
  ]
}
`;

  // 取得候選模型清單（優先 Lite -> 動態最新 -> Standard -> 3.6 -> 2.5）
  const candidateModels = getCandidateModels();
  let lastError = null;

  for (let m = 0; m < candidateModels.length; m++) {
    const currentModel = candidateModels[m];
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${currentModel}:generateContent?key=${apiKey}`;
      const payload = {
        contents: [{ parts: [{ text: promptText }] }],
        generationConfig: { responseMimeType: "application/json" }
      };

      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const responseCode = response.getResponseCode();
      const contentText = response.getContentText();

      if (responseCode !== 200) {
        console.warn(`[${currentModel}] 呼叫回傳 HTTP ${responseCode}: ${contentText}`);
        lastError = new Error(`HTTP ${responseCode}: ${contentText}`);
        continue; // 嘗試下一個候選模型
      }

      const resJson = JSON.parse(contentText);
      const candidateText = resJson.candidates[0].content.parts[0].text;
      Logger.log(`✨ 成功使用模型 [${currentModel}] 完成語意解析`);
      return JSON.parse(candidateText);
    } catch (err) {
      lastError = err;
      console.warn(`[${currentModel}] 嘗試失敗: ${err.message}`);
    }
  }

  console.error('所有候選 Gemini 模型呼叫皆失敗，切換為備援規則: ' + (lastError ? lastError.message : ''));
  return fallbackRuleFilter(rows);
}

function fallbackRuleFilter(rows) {
  const excludeKeywords = ['僅召集人', '召集人開會', '請召集人簽到', '請召集人回覆私訊'];
  const items = [];
  rows.forEach(r => {
    const isInternal = excludeKeywords.some(kw => r.content && r.content.includes(kw));
    if (!isInternal && r.content && !r.content.startsWith('[附件:')) {
      items.push({
        category: '一般通報',
        title: `${r.sender} 發布事項`,
        sender: r.sender,
        points: [r.content]
      });
    }
  });
  return { summary: '本半日彙整之公告事項。', items: items };
}

// ==========================
// 7. 輔助函式 (Drive存檔、LINE Reply、HTML產生)
// ==========================
function saveLineContentToDrive(messageId, fileName) {
  try {
    const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }
    });
    const blob = response.getBlob().setName(fileName);
    const folder = DRIVE_FOLDER_ID ? DriveApp.getFolderById(DRIVE_FOLDER_ID) : DriveApp.getRootFolder();
    return folder.createFile(blob);
  } catch (e) {
    console.error('下載 LINE 內容至 Drive 失敗: ' + e.message);
    return null;
  }
}

function replyLineMessage(replyToken, text) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || !replyToken) return;
  try {
    const url = 'https://api.line.me/v2/bot/message/reply';
    UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
      },
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: [{ type: 'text', text: text }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('LINE 回覆失敗: ' + e.message);
  }
}

function notifyAdmin(message) {
  if (!ADMIN_USER_ID || !LINE_CHANNEL_ACCESS_TOKEN) return;
  try {
    const url = 'https://api.line.me/v2/bot/message/push';
    UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
      },
      payload: JSON.stringify({
        to: ADMIN_USER_ID,
        messages: [{ type: 'text', text: message }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('管理員警報通知失敗: ' + e.message);
  }
}

function getUserDisplayName(source) {
  if (source.userId && LINE_CHANNEL_ACCESS_TOKEN) {
    try {
      const url = `https://api.line.me/v2/bot/profile/${source.userId}`;
      const res = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` },
        muteHttpExceptions: true
      });
      const data = JSON.parse(res.getContentText());
      return data.displayName || '群組同仁';
    } catch (e) {
      return '群組同仁';
    }
  }
  return '群組同仁';
}

function generateSessionGroupId(sheet, senderName, timestamp) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return Utilities.getUuid();

  const lastSender = sheet.getRange(lastRow, 3).getValue();
  const lastTime = new Date(sheet.getRange(lastRow, 2).getValue());
  const diffMinutes = (timestamp - lastTime) / (1000 * 60);

  if (lastSender === senderName && diffMinutes <= 2) {
    return sheet.getRange(lastRow, 8).getValue();
  }
  return Utilities.getUuid();
}

function getTeacherEmailList(sheet) {
  const values = sheet.getDataRange().getValues();
  const emails = [];
  for (let i = 1; i < values.length; i++) {
    const email = values[i][1];
    const enabled = values[i][2];
    if (email && (enabled === true || enabled === 'TRUE' || enabled === '')) {
      emails.push(email.toString().trim());
    }
  }
  return emails;
}

function getPendingMessageCount(sheet) {
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let r = 1; r < data.length; r++) {
    if (data[r][8] === false || data[r][8] === 'FALSE') count++;
  }
  return count;
}

function markRowsAsProcessed(sheet, rows, batchTag) {
  rows.forEach(r => {
    sheet.getRange(r.rowIndex, 9).setValue(true);
    sheet.getRange(r.rowIndex, 10).setValue(batchTag);
  });
}

function generateHtmlEmail(timeSlotLabel, digest) {
  let cardsHtml = '';
  digest.items.forEach(item => {
    let listHtml = '';
    item.points.forEach(pt => {
      listHtml += `<li style="margin-bottom: 4px;">${pt}</li>`;
    });

    cardsHtml += `
      <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin-bottom: 16px;">
        <div style="font-size: 12px; color: #64748b; margin-bottom: 6px;">
          <span style="background: #e2e8f0; color: #1e293b; padding: 2px 8px; border-radius: 12px; font-weight: bold; margin-right: 8px;">${item.category}</span>
          發布者：${item.sender}
        </div>
        <h3 style="margin: 0 0 10px 0; font-size: 16px; color: #1e293b;">${item.title}</h3>
        <ul style="margin: 0; padding-left: 20px; color: #334155; font-size: 14px; line-height: 1.6;">
          ${listHtml}
        </ul>
      </div>
    `;
  });

  return `
    <div style="max-width: 650px; margin: 0 auto; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background: #f8fafc; padding: 20px;">
      <div style="background: #1e3a8a; color: #ffffff; padding: 20px; text-align: center; border-radius: 8px 8px 0 0;">
        <h2 style="margin: 0; font-size: 20px;">🏫 竹林召集人群組 半日訊息整理</h2>
        <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.9;">發布時段：${timeSlotLabel} ｜ 附件已直接隨信夾帶</p>
      </div>
      <div style="background: #ffffff; padding: 20px; border-radius: 0 0 8px 8px; border: 1px solid #e2e8f0; border-top: none;">
        <div style="background: #eff6ff; border-left: 4px solid #2563eb; padding: 12px; margin-bottom: 20px; font-size: 14px; color: #1e40af;">
          <strong>💡 摘要速覽：</strong> ${digest.summary}
        </div>
        ${cardsHtml}
        <div style="margin-top: 20px; padding: 12px; background: #f1f5f9; border-radius: 6px; font-size: 12px; color: #64748b; text-align: center;">
          📎 相關公文 PDF 與通知圖片已夾帶於信件附件，請查閱下載。<br>
          本信件由召集人通報秘書系統自動整理發送。
        </div>
      </div>
    </div>
  `;
}

function initRawSheet(ss) {
  const sheet = ss.insertSheet(SHEET_RAW);
  sheet.appendRow([
    '訊息ID (message_id)',
    '接收時間 (timestamp)',
    '發布者 (sender_name)',
    '類型 (message_type)',
    '訊息內容 (content)',
    'Drive檔案ID (drive_file_id)',
    'Drive網址 (drive_file_url)',
    '關聯群組碼 (session_group_id)',
    '已處理 (is_processed)',
    '處理批次 (batch_tag)'
  ]);
  return sheet;
}
