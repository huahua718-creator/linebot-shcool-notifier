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

const SHEET_RAW = 'Raw_Messages';       // 原始接收訊息表
const SHEET_TEACHERS = 'Teacher_List';  // 教師Email通訊錄
const SHEET_DIGEST = 'Digest_Summaries';// 彙整派報歷史記錄

// ==========================
// 2. LINE Webhook 接收端 (doPost)
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

      // 寫入 Google Sheets
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
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================
// 3. 定時彙整與 Email 派報 (12:00 / 16:00 觸發)
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

    // 使用 Bcc 密件副本保障隱私，或批次寄送
    GmailApp.sendEmail(recipients.join(','), subject, '請使用支援 HTML 的郵件客戶端開啟本通知。', {
      htmlBody: htmlBody,
      attachments: attachments,
      name: '竹林召集人智慧秘書'
    });
    console.log(`已成功發送郵件給 ${recipients.length} 位老師。`);
  }

  // 5. 標記訊息為已處理
  markRowsAsProcessed(rawSheet, pendingRows, `${timeSlotLabel}_SENT`);
}

// ==========================
// 4. Gemini API 智慧分類、篩選與摘要
// ==========================
function analyzeAndSummarizeWithGemini(rows) {
  if (!GEMINI_API_KEY) {
    // 若尚未填寫 Gemini API Key，走簡易本地關鍵字規則作為備援
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

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
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

    const resJson = JSON.parse(response.getContentText());
    const candidateText = resJson.candidates[0].content.parts[0].text;
    return JSON.parse(candidateText);
  } catch (err) {
    console.error('Gemini API 處理失敗，改用規則模式: ' + err.message);
    return fallbackRuleFilter(rows);
  }
}

// 規則式備援機制（當無 API Key 時運作）
function fallbackRuleFilter(rows) {
  const excludeKeywords = ['僅召集人', '召集人開會', '請召集人簽到', '請召集人回覆私訊'];
  const items = [];
  rows.forEach(r => {
    const isInternal = excludeKeywords.some(kw => r.content.includes(kw));
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
// 5. 輔助函式 (Drive存檔、LINE Token、HTML產生)
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
    return sheet.getRange(lastRow, 8).getValue(); // 繼承上一筆的 Group ID
  }
  return Utilities.getUuid();
}

function getTeacherEmailList(sheet) {
  const values = sheet.getDataRange().getValues();
  const emails = [];
  for (let i = 1; i < values.length; i++) {
    const email = values[i][1]; // 假設 B 欄是 Email
    const enabled = values[i][2]; // 假設 C 欄是 是否接收 (TRUE/FALSE)
    if (email && (enabled === true || enabled === 'TRUE' || enabled === '')) {
      emails.push(email.toString().trim());
    }
  }
  return emails;
}

function markRowsAsProcessed(sheet, rows, batchTag) {
  rows.forEach(r => {
    sheet.getRange(r.rowIndex, 9).setValue(true);     // is_processed = true
    sheet.getRange(r.rowIndex, 10).setValue(batchTag); // batch_tag
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
