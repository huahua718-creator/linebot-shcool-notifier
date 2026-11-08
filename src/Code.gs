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
const TARGET_GROUP_ID = SCRIPT_PROPERTIES.getProperty('TARGET_GROUP_ID'); // 指定監聽的群組ID (選填)

const SHEET_RAW = 'Raw_Messages';           // 原始接收訊息表
const SHEET_TEACHERS = 'Teacher_List';      // 教師Email通訊錄
const SHEET_DIGEST = 'Digest_Summaries';    // 彙整派報歷史記錄
const SHEET_MEMBERS = 'Member_Directory';   // 群組發言人與處室單位對照表

// 取得有效的 Gemini API Key
function getActiveGeminiApiKey() {
  return GEMINI_API_KEY || SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
}

// ==========================
// 2. 測試專用函式專區
// ==========================

/**
 * 📧 【測試寄信功能】：只寄給管理員測試信件
 */
function testSendEmailToAdmin() {
  console.log('🚀 開始執行管理員寄信測試 (testSendEmailToAdmin)...');
  const adminEmail = Session.getActiveUser().getEmail();
  console.log(`📬 測試信收件者 Email：${adminEmail}`);
  
  if (!adminEmail) {
    console.error('❌ 無法自動偵測管理員 Email，請手動指定。');
    return;
  }

  const remainingQuota = MailApp.getRemainingDailyQuota();
  console.log(`📊 今日剩餘寄信額度：${remainingQuota} 封`);

  const mockDigest = {
    summary: '此為系統寄件測試信。教務處已發布段考核對提醒，學務處公告下週疫苗施打。',
    items: [
      {
        department: '教務處',
        category: '教務與段考',
        title: '【教務處公告】115-1 第二次段考範圍通知確認',
        official_sender: '教務處教學組 (黃怡瑜老師)',
        points: [
          '第二次段考各科範圍表已由教學組擬定完成',
          '煩請各科召集人與任課老師確實檢查內容',
          '若無疑義，今日下午將正式公告全校'
        ]
      },
      {
        department: '學務處',
        category: '學務與活動',
        title: '【學務處通知】下週一 (11/09) 全校流感疫苗施打時程',
        official_sender: '學務處體衛組 (曉光老師)',
        points: [
          '施打當日請全校師生務必攜帶「健保卡」',
          '各班請按表定時程於活動中心舞台前集合',
          '遇跨兩節課之任課教師，請利用下課時間提前5分鐘交接'
        ]
      }
    ]
  };

  const htmlBody = generateHtmlEmail('管理員寄件測試時段', mockDigest);
  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm');
  const subject = `【測試信件】竹林召集人群組 半日訊息整理 (${todayStr})`;

  try {
    GmailApp.sendEmail(adminEmail, subject, '這是一封測試郵件，請使用支援 HTML 的郵件閱讀器檢視。', {
      htmlBody: htmlBody,
      name: '竹林召集人智慧秘書 (測試)'
    });
    console.log(`🎉 【測試信發送成功！】已寄至：${adminEmail}`);
  } catch (err) {
    console.error('❌ 寄件失敗：' + err.message);
  }
}

/**
 * 🧪 測試 Gemini API 是否調用成功
 */
function testGeminiApi() {
  console.log('🚀 開始測試 Gemini API 連線與動態模型解析...');
  const candidates = getCandidateModels();
  if (!getActiveGeminiApiKey()) return '❌ 尚未設定 GEMINI_API_KEY！';

  const mockMessages = [
    {
      sender: '教務處教學組 (黃怡瑜 Lia 老師)',
      department: '教務處',
      type: 'text',
      content: '@All 各位老師大家好，附件為 115-1 第二次段考範圍通知，煩請協助確認內容，若沒問題今天下午正式公告全校。',
      time: '14:51'
    }
  ];

  try {
    const result = analyzeAndSummarizeWithGemini(mockMessages.map((m, idx) => ({
      rowIndex: idx + 1,
      sender: m.sender,
      department: m.department,
      type: m.type,
      content: m.content,
      timestamp: new Date()
    })));

    if (result && result.items && result.items.length > 0) {
      return `🎉 【Gemini 測試成功！】\n✅ 成功辨識公告：${result.items[0].title}\n💡 摘要：${result.summary}`;
    }
    return '⚠️ API 有回應但未成功解析出公告。';
  } catch (err) {
    return '❌ 測試失敗：' + err.message;
  }
}

// ==========================
// 3. 動態偵測與模型候選邏輯
// ==========================
function getLatestFlashModelName() {
  const defaultModel = "gemini-flash-lite-latest";
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
          const isExcluded = name.includes("tts") || name.includes("image") || name.includes("computer-use") || name.includes("omni") || name.includes("exp");
          return hasGen && isFlash && !isExcluded;
        })
        .map(m => m.name.replace("models/", ""));

      if (flashModels.length > 0) {
        flashModels.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
        return flashModels[0];
      }
    }
  } catch (err) {
    Logger.log("⚠️ 取得模型清單失敗: " + err.toString());
  }
  return defaultModel;
}

function getCandidateModels() {
  const dynamicModel = getLatestFlashModelName();
  return Array.from(new Set([
    "gemini-flash-lite-latest",
    dynamicModel,
    "gemini-flash-latest",
    "gemini-3.6-flash",
    "gemini-2.5-flash"
  ].filter(Boolean)));
}

// ==========================
// 4. LINE Webhook 接收端
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
    if (!rawSheet) rawSheet = initRawSheet(ss);

    let memberSheet = ss.getSheetByName(SHEET_MEMBERS);
    if (!memberSheet) memberSheet = initMemberSheet(ss);

    let teacherSheet = ss.getSheetByName(SHEET_TEACHERS);
    if (!teacherSheet) teacherSheet = initTeacherSheet(ss);

    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const sourceType = event.source.type;
      const userId = event.source.userId || '';
      const rawSenderName = getUserDisplayName(event.source);

      // --- [新好友加入或解除封鎖事件：引導輸入姓名與 Email 綁定] ---
      if (event.type === 'follow') {
        sendWelcomeAndBindingPrompt(event.replyToken, rawSenderName);
        continue;
      }

      if (event.type !== 'message') continue;

      const message = event.message;
      const timestamp = new Date(event.timestamp);
      const groupId = event.source.groupId || event.source.roomId || '';
      const messageId = message.id;
      const messageType = message.type;

      // 🛡️【安全邊界】：1 對 1 私人對話處理 (處理身分綁定或指令，絕不記錄發送)
      if (sourceType === 'user') {
        handleOneOnOneChat(event, userId, rawSenderName, message, rawSheet, memberSheet, teacherSheet);
        continue;
      }

      // 🛡️【群組過濾】：若設定 TARGET_GROUP_ID，則鎖定指定召集人群組
      if (TARGET_GROUP_ID && groupId !== TARGET_GROUP_ID) {
        continue;
      }

      // --- [群組公務收集流程] ---
      const affiliation = registerOrUpdateMember(memberSheet, userId, rawSenderName, timestamp);

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

      const sessionGroupId = generateSessionGroupId(rawSheet, affiliation.officialSender, timestamp);

      rawSheet.appendRow([
        messageId,
        timestamp,
        userId,
        rawSenderName,
        affiliation.department,
        affiliation.officialSender,
        messageType,
        textContent,
        driveFileId,
        driveFileUrl,
        sessionGroupId,
        false,
        ''
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
// 5. 1 對 1 私聊處理：引導同仁綁定身分與管理指令
// ==========================
function handleOneOnOneChat(event, userId, rawSenderName, message, rawSheet, memberSheet, teacherSheet) {
  if (message.type !== 'text') {
    replyWithMenu(event.replyToken, 'ℹ️ 您好！本機器人僅在「召集人群組」中收集公務通知。\n若要設定接收每日郵件通知，請直接回覆您的姓名與 Email 進行綁定！');
    return;
  }

  const text = message.text.trim();
  const textLower = text.toLowerCase();

  // 1. 偵測是否為身分綁定輸入 (包含 Email 格式或含有「綁定」關鍵字)
  const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/;
  const emailMatch = text.match(emailRegex);

  if (emailMatch || textLower === '身分綁定' || textLower === '綁定' || textLower === '!bind') {
    if (!emailMatch) {
      // 點擊按鈕但未帶 Email，提示範例格式
      const promptMsg = `📝【教師身分綁定引導】\n您好 ${rawSenderName} 老師！請直接回覆您的姓名與公務 Email：\n\n👉 範例格式：\n綁定 王大明 daming@school.tp.edu.tw\n\n（系統將自動在 Teacher_List 名冊中連結您的身分與 LINE 帳號！）`;
      replyWithMenu(event.replyToken, promptMsg);
      return;
    }

    // 含有 Email，執行綁定邏輯
    const extractedEmail = emailMatch[0].trim();
    // 移除「綁定」字樣與 Email，剩下的文字作為姓名
    let extractedName = text.replace(extractedEmail, '').replace(/綁定|老師|主任|組長|：|:/g, '').trim();
    if (!extractedName) extractedName = rawSenderName; // 若未填姓名則暫用 LINE 暱稱

    const bindResult = bindTeacherProfile(teacherSheet, userId, rawSenderName, extractedName, extractedEmail);
    replyWithMenu(event.replyToken, bindResult);
    return;
  }

  // 2. 查詢自己的 LINE User ID
  if (textLower === '我的id' || textLower === 'id' || textLower === '!id') {
    replyWithMenu(event.replyToken, `🆔 您的 LINE User ID 為：\n${userId}\n\nLINE 暱稱：${rawSenderName}`);
    return;
  }

  // 3. 指令選單 / 說明
  if (textLower === '選單' || textLower === 'menu' || textLower === '!menu' || textLower === '說明' || textLower === '!help') {
    sendAdminMenuCard(event.replyToken, rawSenderName);
    return;
  }

  // 4. 管理員指令：!status
  if (textLower === '!status' || textLower === '查詢狀態' || textLower === '狀態') {
    if (ADMIN_USER_ID && userId !== ADMIN_USER_ID) {
      replyWithMenu(event.replyToken, '❌ 抱歉，此功能為系統管理員專用指令。');
    } else {
      const pendingCount = getPendingMessageCount(rawSheet);
      replyWithMenu(event.replyToken, `📊 系統狀態報告：\n目前試算表中待彙整派報的群組公務訊息共有 【${pendingCount}】 筆。\n\n💡 提示：若該時段待處理為 0 筆，系統將自動跳過不寄信。`);
    }
    return;
  }

  // 5. 管理員指令：!members
  if (textLower === '!members' || textLower === '人員名冊' || textLower === '名冊') {
    if (ADMIN_USER_ID && userId !== ADMIN_USER_ID) {
      replyWithMenu(event.replyToken, '❌ 抱歉，此功能為系統管理員專用指令。');
    } else {
      const memberSummary = getMemberDirectorySummary(memberSheet);
      replyWithMenu(event.replyToken, memberSummary);
    }
    return;
  }

  // 6. 管理員指令：!send
  if (textLower === '!send' || textLower === '立即派報' || textLower === '手動派報') {
    if (ADMIN_USER_ID && userId !== ADMIN_USER_ID) {
      replyWithMenu(event.replyToken, '❌ 抱歉，此功能為系統管理員專用指令。');
    } else {
      const pendingCount = getPendingMessageCount(rawSheet);
      if (pendingCount === 0) {
        replyWithMenu(event.replyToken, 'ℹ️ 目前群組【無任何待處理新訊息】，依系統規則本次不寄發信件。');
      } else {
        replyWithMenu(event.replyToken, `🚀 已收到管理員指令！偵測到有 ${pendingCount} 筆待處理群組訊息，系統正在彙整派發中...`);
        processAndSendDigest('管理員手動即時觸發');
      }
    }
    return;
  }

  // 7. 管理員指令：!test
  if (textLower === '!test' || textLower === '測試ai' || textLower === '測試') {
    if (ADMIN_USER_ID && userId !== ADMIN_USER_ID) {
      replyWithMenu(event.replyToken, '❌ 抱歉，此功能為系統管理員專用指令。');
    } else {
      replyWithMenu(event.replyToken, '🧪 正在連線測試 Gemini AI 模型中，請稍候約 1~3 秒...');
      const testResult = testGeminiApi();
      notifyAdmin(testResult);
    }
    return;
  }

  // 預設回傳引導選單
  sendAdminMenuCard(event.replyToken, rawSenderName);
}

/**
 * 🔗 核心功能：在 Teacher_List 比對同仁姓名與 Email，連帶記錄 UID、暱稱與綁定時間
 */
function bindTeacherProfile(sheet, userId, displayName, nameInput, emailInput) {
  if (!sheet) return '⚠️ 系統尚未建立 Teacher_List 工作表，請聯繫管理員！';

  const data = sheet.getDataRange().getValues();
  let matchedRow = -1;
  let department = '';
  let officialName = nameInput;

  // 清理比對用字串 (去除空格與「老師」後綴)
  const cleanInputName = nameInput.replace(/\s+|老師/g, '').toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const sheetName = String(data[i][0] || '').replace(/\s+|老師/g, '').toLowerCase();
    const sheetEmail = String(data[i][1] || '').trim().toLowerCase();

    // 優先比對姓名或 Email
    if ((cleanInputName && sheetName && (sheetName.includes(cleanInputName) || cleanInputName.includes(sheetName))) ||
        (sheetEmail && sheetEmail === emailInput.toLowerCase())) {
      matchedRow = i + 1;
      officialName = data[i][0] || nameInput;
      department = data[i][3] || '一般科別';
      break;
    }
  }

  const timeStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss');

  if (matchedRow > 0) {
    // 找到該同仁：更新 Email、啟用通知、回填 UID、暱稱與綁定時間
    sheet.getRange(matchedRow, 2).setValue(emailInput); // B欄: Email
    sheet.getRange(matchedRow, 3).setValue(true);       // C欄: 接收通知 = TRUE
    sheet.getRange(matchedRow, 6).setValue(userId);     // F欄: LINE UID
    sheet.getRange(matchedRow, 7).setValue(displayName);// G欄: LINE 暱稱
    sheet.getRange(matchedRow, 8).setValue(timeStr);    // H欄: 綁定時間

    notifyAdmin(`📢 同仁身分綁定通知：【${officialName}】已完成 LINE 綁定！\nEmail：${emailInput}\nLINE 暱稱：${displayName}`);

    return `🎉【身分綁定成功！】\n\n` +
           `👤 教師姓名：${officialName}\n` +
           `📧 通知信箱：${emailInput}\n` +
           `🏢 所屬處室/科別：${department}\n` +
           `📱 綁定 LINE 暱稱：${displayName}\n\n` +
           `✅ 系統已將您的 LINE 與教師通訊錄成功串接！\n` +
           `每日 12:00 與 16:00 的公務半日整理將準時寄送至您的 Email。`;
  } else {
    // 名冊未建檔：自動追加一列新同仁資料
    sheet.appendRow([
      nameInput,
      emailInput,
      true,
      '待確認處室',
      '由LINE同仁自助綁定新增',
      userId,
      displayName,
      timeStr
    ]);

    notifyAdmin(`📢 新同仁自助註冊：【${nameInput}】已登記至名冊！\nEmail：${emailInput}\nLINE 暱稱：${displayName}`);

    return `🎉【註冊與綁定完成！】\n\n` +
           `👤 教師姓名：${nameInput}\n` +
           `📧 通知信箱：${emailInput}\n` +
           `📱 綁定 LINE 暱稱：${displayName}\n\n` +
           `✅ 系統已為您在教師名冊中建立檔案並開啟通知！每日 12:00 與 16:00 派報將寄送至該信箱。`;
  }
}

// 引導好友加入之歡迎詞
function sendWelcomeAndBindingPrompt(replyToken, displayName) {
  const welcomeText = `👋 您好 ${displayName} 老師！歡迎加入「竹林召集人通報秘書」！

本系統每日於 12:00 與 16:00 自動彙整召集人群組之重要行政公文、段考範圍與活動調移，派送至全體同仁信箱。

📝【請進行身分綁定】：
請直接依照下方範例格式回覆您的「姓名」與「公務Email」：
👉 範例：綁定 王大明 daming@school.tp.edu.tw

點擊下方「📝 身分綁定」即可開始！`;

  replyWithMenu(replyToken, welcomeText);
}

// ==========================
// 6. 定時彙整與 Email 派報 (12:00 / 16:00 觸發)
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
  if (data.length <= 1) return;

  const pendingRows = [];
  for (let r = 1; r < data.length; r++) {
    const isProcessed = data[r][11];
    if (isProcessed === false || isProcessed === 'FALSE') {
      pendingRows.push({
        rowIndex: r + 1,
        messageId: data[r][0],
        timestamp: data[r][1],
        userId: data[r][2],
        rawSender: data[r][3],
        department: data[r][4],
        sender: data[r][5] || data[r][3],
        type: data[r][6],
        content: data[r][7],
        fileId: data[r][8],
        fileUrl: data[r][9],
        groupId: data[r][10]
      });
    }
  }

  if (pendingRows.length === 0) {
    console.log(`[${timeSlotLabel}] 該時段完全無待處理新訊息，系統略過不寄信。`);
    return;
  }

  const digestResult = analyzeAndSummarizeWithGemini(pendingRows);

  if (!digestResult.items || digestResult.items.length === 0) {
    console.log(`[${timeSlotLabel}] 訊息皆屬召集人內部工作，不打擾全校老師，系統不寄信。`);
    markRowsAsProcessed(rawSheet, pendingRows, `${timeSlotLabel}_SKIPPED`);
    return;
  }

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

  const htmlBody = generateHtmlEmail(timeSlotLabel, digestResult);
  const recipients = getTeacherEmailList(teacherSheet);

  if (recipients.length === 0) {
    notifyAdmin('⚠️ 派報中斷：Teacher_List 工作表中沒有啟用 (TRUE) 的教師 Email！');
    return;
  }

  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd');
  const subject = `【竹林召集人通報】${todayStr} ${timeSlotLabel} 重點訊息整理`;

  try {
    GmailApp.sendEmail(recipients.join(','), subject, '請使用支援 HTML 的郵件客戶端開啟本通知。', {
      htmlBody: htmlBody,
      attachments: attachments,
      name: '竹林召集人智慧秘書'
    });

    if (digestSheet) {
      digestSheet.appendRow([
        `DIG-${Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd-HHmm')}`,
        new Date(),
        timeSlotLabel,
        digestResult.items.map(it => it.department || it.category).join('、'),
        digestResult.items[0].title,
        digestResult.summary,
        attachments.length,
        recipients.length,
        'SUCCESS'
      ]);
    }

    notifyAdmin(`✅ 已完成 [${timeSlotLabel}] 訊息派報，共整理 ${digestResult.items.length} 則公告，發送給 ${recipients.length} 位老師！`);
  } catch (e) {
    console.error('Gmail 寄送失敗: ' + e.message);
    notifyAdmin(`⚠️ 寄發 Email 失敗：${e.message}`);
  }

  markRowsAsProcessed(rawSheet, pendingRows, `${timeSlotLabel}_SENT`);
}

// ==========================
// 7. Gemini API 智慧分類與單位身分發文
// ==========================
function analyzeAndSummarizeWithGemini(rows) {
  const apiKey = getActiveGeminiApiKey();
  if (!apiKey) return fallbackRuleFilter(rows);

  const promptText = `
你是一位專業的高中行政教務秘書。以下是學校「召集人群組」今日由各處室幹事、主任或召集人發出的訊息清單：
${JSON.stringify(rows.map(r => ({
  sender: r.sender,
  department: r.department || "未指定處室",
  type: r.type,
  content: r.content,
  time: Utilities.formatDate(new Date(r.timestamp), 'GMT+8', 'HH:mm')
})))}

【重要任務與發文身分規範】：
1. 判斷每一則訊息「是否需要公告給全體老師」：
   - 需公告：全校時程、段考範圍與核對、公文更正、學務活動調課、榮譽榜、人事退休福利等。
   - 僅限召集人內部事項【剔除不收錄】：例如催繳審題表、召集人會議領取試卷、私訊答覆等。
2. 【處室單位名義連結】：
   - 若訊息發布者的文字未自帶單位名稱，請務必直接連結其後端對應的處室單位（如「教務處」、「學務處」、「人事室」），以「單位名義」擬定正式標題（例如：【教務處公告】第二次段考範圍核對）。
3. 輸出嚴格的 JSON 格式（不要包含 markdown 代碼區塊），結構如下：
{
  "summary": "1~2句話的本次半日通報總結",
  "items": [
    {
      "department": "發布處室單位",
      "category": "教務與段考 | 學務與活動 | 人事公告 | 緊急更正 | 榮譽榜",
      "title": "以單位名義起頭之簡潔主旨標題",
      "official_sender": "官方顯示代表",
      "points": ["重點1", "重點2", "注意事項與時限"]
    }
  ]
}
`;

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
        lastError = new Error(`HTTP ${responseCode}: ${contentText}`);
        continue;
      }

      const resJson = JSON.parse(contentText);
      const candidateText = resJson.candidates[0].content.parts[0].text;
      return JSON.parse(candidateText);
    } catch (err) {
      lastError = err;
    }
  }

  console.error('Gemini 呼叫失敗，改用備援規則: ' + (lastError ? lastError.message : ''));
  return fallbackRuleFilter(rows);
}

function fallbackRuleFilter(rows) {
  const excludeKeywords = ['僅召集人', '召集人開會', '請召集人簽到', '請召集人回覆私訊'];
  const items = [];
  rows.forEach(r => {
    const isInternal = excludeKeywords.some(kw => r.content && r.content.includes(kw));
    if (!isInternal && r.content && !r.content.startsWith('[附件:')) {
      const dept = r.department || '學校行政處室';
      items.push({
        department: dept,
        category: '一般通報',
        title: `【${dept}】${r.sender} 發布事項`,
        official_sender: r.sender,
        points: [r.content]
      });
    }
  });
  return { summary: '本半日彙整之公告事項。', items: items };
}

// ==========================
// 8. 發言人名冊管理 (Member_Directory)
// ==========================
function registerOrUpdateMember(sheet, userId, displayName, timestamp) {
  if (!sheet) return { department: '', title: '', officialSender: displayName };

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  let department = '';
  let title = '';
  let officialSender = '';

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      foundRow = i + 1;
      department = data[i][2];
      title = data[i][3];
      officialSender = data[i][4];
      break;
    }
  }

  const timeStr = Utilities.formatDate(timestamp, 'GMT+8', 'yyyy-MM-dd HH:mm:ss');

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2).setValue(displayName);
    sheet.getRange(foundRow, 6).setValue(timeStr);
  } else {
    sheet.appendRow([userId, displayName, '', '', '', timeStr]);
  }

  let finalSender = officialSender;
  if (!finalSender) {
    if (department) {
      finalSender = `${department} (${displayName})`;
    } else {
      finalSender = displayName;
    }
  }

  return {
    department: department,
    title: title,
    officialSender: finalSender
  };
}

function getMemberDirectorySummary(sheet) {
  if (!sheet) return '⚠️ 尚未建立 Member_Directory 工作表。';
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return '👥 目前群組名冊中尚未有任何人員發言紀錄。';

  let total = data.length - 1;
  let unsetCount = 0;
  let listStr = '';

  for (let i = 1; i < data.length; i++) {
    const name = data[i][1] || '未知名稱';
    const dept = data[i][2] || '❌ 未設定處室';
    if (!data[i][2]) unsetCount++;
    listStr += `• ${name} ➡️ ${dept}\n`;
  }

  return `👥【群組發言人與單位名冊】\n總計已記錄 ${total} 位發言人（其中 ${unsetCount} 位待設定處室）：\n\n${listStr}\n💡 提示：請至 Google 試算表「Member_Directory」分頁填寫對應處室，系統將自動以該處室名義發文！`;
}

// ==========================
// 9. LINE 快捷按鈕與面板
// ==========================
function getQuickReplyItems() {
  return {
    items: [
      { type: "action", action: { type: "message", label: "📝 身分綁定", text: "身分綁定" } },
      { type: "action", action: { type: "message", label: "📊 查詢狀態", text: "!status" } },
      { type: "action", action: { type: "message", label: "👥 人員名冊", text: "!members" } },
      { type: "action", action: { type: "message", label: "🚀 立即派報", text: "!send" } },
      { type: "action", action: { type: "message", label: "🧪 測試AI", text: "!test" } },
      { type: "action", action: { type: "message", label: "🆔 我的ID", text: "!id" } }
    ]
  };
}

function replyWithMenu(replyToken, text) {
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
        messages: [{
          type: 'text',
          text: text,
          quickReply: getQuickReplyItems()
        }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('replyWithMenu 失敗: ' + e.message);
  }
}

function sendAdminMenuCard(replyToken, userName) {
  const menuText = `🛠️【召集人通報秘書・同仁服務面板】
您好 ${userName} 老師！請直接點擊下方快捷按鈕操作：

📝 身分綁定 (輸入：綁定 姓名 信箱)
  將您的 LINE 帳號與學校教師通訊錄連結，準時接收半日公務通報

📊 狀態查詢 (管理員指令 !status)
  查看目前累積待派報之公務訊息筆數

👥 人員名冊 (管理員指令 !members)
  查看群組發言人與單位綁定狀況

🚀 立即派報 (管理員指令 !send)
  立即手動執行半日彙整並以處室名義發送 Email

🔒 安全提示：個人私訊傳送的內容絕不會被公開發送，請安心使用！`;

  replyWithMenu(replyToken, menuText);
}

// ==========================
// 10. 輔助函式 (Drive存檔、LINE Reply、HTML產生)
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
        messages: [{
          type: 'text',
          text: message,
          quickReply: getQuickReplyItems()
        }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('管理員通知失敗: ' + e.message);
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

  const lastSender = sheet.getRange(lastRow, 6).getValue();
  const lastTime = new Date(sheet.getRange(lastRow, 2).getValue());
  const diffMinutes = (timestamp - lastTime) / (1000 * 60);

  if (lastSender === senderName && diffMinutes <= 2) {
    return sheet.getRange(lastRow, 11).getValue();
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
    if (data[r][11] === false || data[r][11] === 'FALSE') count++;
  }
  return count;
}

function markRowsAsProcessed(sheet, rows, batchTag) {
  rows.forEach(r => {
    sheet.getRange(r.rowIndex, 12).setValue(true);
    sheet.getRange(r.rowIndex, 13).setValue(batchTag);
  });
}

function generateHtmlEmail(timeSlotLabel, digest) {
  let cardsHtml = '';
  digest.items.forEach(item => {
    let listHtml = '';
    item.points.forEach(pt => {
      listHtml += `<li style="margin-bottom: 4px;">${pt}</li>`;
    });

    const deptTag = item.department ? `<span style="background: #1e40af; color: #ffffff; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 6px;">${item.department}</span>` : '';
    const senderDisplay = item.official_sender || item.sender || '學校行政處室';

    cardsHtml += `
      <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin-bottom: 16px;">
        <div style="font-size: 12px; color: #64748b; margin-bottom: 8px;">
          ${deptTag}
          <span style="background: #f1f5f9; color: #334155; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 8px;">${item.category}</span>
          發布單位/人員：<strong>${senderDisplay}</strong>
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
        <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.9;">發布時段：${timeSlotLabel} ｜ 公務通知已依處室整合派發</p>
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
    '發言人LINE ID (user_id)',
    '發言人暱稱 (display_name)',
    '所屬處室 (department)',
    '官方代表身分 (official_sender)',
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

function initMemberSheet(ss) {
  const sheet = ss.insertSheet(SHEET_MEMBERS);
  sheet.appendRow([
    '發言人 LINE ID (user_id)',
    'LINE 顯示名稱 (display_name)',
    '所屬單位/處室 (department)',
    '職稱/身分 (title)',
    '官方發文代表名稱 (official_sender)',
    '最後發言時間 (last_seen)'
  ]);
  return sheet;
}

function initTeacherSheet(ss) {
  const sheet = ss.insertSheet(SHEET_TEACHERS);
  sheet.appendRow([
    '教師姓名 (Name)',
    '電子信箱 (Email)',
    '接收通知 (IsActive)',
    '任教科別/處室 (Department)',
    '備註說明 (Notes)',
    'LINE UID (user_id)',
    'LINE 暱稱 (display_name)',
    '綁定時間 (bound_at)'
  ]);
  return sheet;
}
