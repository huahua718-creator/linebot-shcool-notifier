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

function getActiveGeminiApiKey() {
  return GEMINI_API_KEY || SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
}

function isUserAdmin(userId) {
  if (!ADMIN_USER_ID) return false;
  return userId === ADMIN_USER_ID;
}

// ==========================
// 2. 測試專用函式專區
// ==========================
function testSendEmailToAdmin() {
  console.log('開始執行管理員寄信測試 (testSendEmailToAdmin)...');
  const adminEmail = Session.getActiveUser().getEmail();
  console.log(`測試信收件者 Email：${adminEmail}`);
  if (!adminEmail) {
    console.error('無法自動偵測管理員 Email，請手動指定。');
    return;
  }

  const remainingQuota = MailApp.getRemainingDailyQuota();
  console.log(`今日剩餘寄信額度：${remainingQuota} 封`);

  const mockDigest = {
    summary: '此為系統寄件編碼防亂碼測試信。教務處發布段考核對提醒，學務處公告健檢時程。',
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
        department: '人事室',
        category: '人事公告',
        title: '【人事室公告】115年度全體教職員工健康檢查注意事項與時程',
        official_sender: '人事室',
        points: [
          '檢查日期：115年11月17日(星期二)上午07:10~11:30 (11:15截止報到)，地點在活動中心。',
          '報到請帶身分證；健檢前需空腹8小時。',
          '請於11月13日(星期五)下午4點前至人事室領取健檢注意事項及採集管。'
        ]
      }
    ]
  };

  const htmlBody = generateHtmlEmail('管理員寄件測試時段', mockDigest);
  const plainText = generatePlainTextEmail('管理員寄件測試時段', mockDigest);
  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm');
  const subject = `【測試信件】竹林召集人群組 半日訊息整理 (${todayStr})`;

  try {
    GmailApp.sendEmail(adminEmail, subject, plainText, {
      htmlBody: htmlBody,
      name: '竹林召集人通報秘書'
    });
    console.log(`【測試信發送成功！】已寄至：${adminEmail}，無任何亂碼問號！`);
  } catch (err) {
    console.error('寄件失敗：' + err.message);
  }
}

function testGeminiApi() {
  console.log('開始測試 Gemini API 連線與動態模型解析...');
  if (!getActiveGeminiApiKey()) return '尚未設定 GEMINI_API_KEY！';

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
      return `【Gemini 測試成功！】\n成功辨識公告：${result.items[0].title}\n摘要：${result.summary}`;
    }
    return 'API 有回應但未成功解析出公告。';
  } catch (err) {
    return '測試失敗：' + err.message;
  }
}

// ==========================
// 3. 動態偵測與模型候選邏輯
// ==========================
function getLatestFlashModelName() {
  const defaultModel = 'gemini-flash-lite-latest';
  const apiKey = getActiveGeminiApiKey();
  if (!apiKey) return defaultModel;

  const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  try {
    const response = UrlFetchApp.fetch(listUrl, { 'muteHttpExceptions': true });
    const result = JSON.parse(response.getContentText());
    if (result.models && Array.isArray(result.models)) {
      const flashModels = result.models
        .filter(m => {
          const name = m.name.toLowerCase();
          const hasGen = m.supportedGenerationMethods && m.supportedGenerationMethods.includes('generateContent');
          const isFlash = name.includes('flash');
          const isExcluded = name.includes('tts') || name.includes('image') || name.includes('computer-use') || name.includes('omni') || name.includes('exp');
          return hasGen && isFlash && !isExcluded;
        })
        .map(m => m.name.replace('models/', ''));

      if (flashModels.length > 0) {
        flashModels.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
        return flashModels[0];
      }
    }
  } catch (err) {
    Logger.log('取得模型清單失敗: ' + err.toString());
  }
  return defaultModel;
}

function getCandidateModels() {
  const dynamicModel = getLatestFlashModelName();
  return Array.from(new Set([
    'gemini-flash-lite-latest',
    dynamicModel,
    'gemini-flash-latest',
    'gemini-3.6-flash',
    'gemini-2.5-flash'
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

      // --- [新好友加入或解除封鎖事件] ---
      if (event.type === 'follow') {
        const boundProfile = findTeacherProfileByUid(teacherSheet, userId);
        sendWelcomeAndBindingPrompt(event.replyToken, rawSenderName, boundProfile, userId);
        continue;
      }

      if (event.type !== 'message') continue;

      const message = event.message;
      const timestamp = new Date(event.timestamp);
      const groupId = event.source.groupId || event.source.roomId || '';
      const messageId = message.id;
      const messageType = message.type;

      // 1 對 1 私人對話處理
      if (sourceType === 'user') {
        handleOneOnOneChat(event, userId, rawSenderName, message, rawSheet, memberSheet, teacherSheet);
        continue;
      }

      // 群組過濾
      if (TARGET_GROUP_ID && groupId !== TARGET_GROUP_ID) {
        continue;
      }

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
    notifyAdmin(`[警告] LINE Webhook 發生錯誤：${err.message}`);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================
// 5. 1 對 1 私聊處理：身分綁定、信箱修改、管理員審核預覽
// ==========================
function handleOneOnOneChat(event, userId, rawSenderName, message, rawSheet, memberSheet, teacherSheet) {
  const isAdmin = isUserAdmin(userId);
  const boundProfile = findTeacherProfileByUid(teacherSheet, userId);

  if (message.type !== 'text') {
    replyForUser(event.replyToken, userId, boundProfile, '您好！本對話僅提供身分綁定與信箱設定，群組公務請直接在召集人群組中發布！');
    return;
  }

  const text = message.text.trim();
  const textLower = text.toLowerCase();
  const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/;
  const emailMatch = text.match(emailRegex);

  // 1. 身分綁定 / 信箱查詢與修改流程
  if (textLower === '身分綁定' || textLower === '綁定' || textLower === '修改email' || textLower === '修改信箱' || textLower === '!bind') {
    if (boundProfile) {
      const msg = `【教師身分綁定狀態】\n\n` +
                  `• 姓名：${boundProfile.name} 老師\n` +
                  `• 處室/科別：${boundProfile.department}\n` +
                  `• 目前通知信箱 (舊Mail)：\n  ${boundProfile.email}\n` +
                  `• 接收狀態：${boundProfile.isActive ? '啟用中' : '暫停接收'}\n\n` +
                  `[如需修改接收信箱]：\n` +
                  `請直接在此回覆您的「新Email」即可！\n` +
                  `範例：修改信箱 new_email@school.tp.edu.tw\n` +
                  `（或直接輸入新的信箱地址）`;
      replyForUser(event.replyToken, userId, boundProfile, msg);
      return;
    } else {
      const promptMsg = `【教師身分綁定】\n` +
                        `您好 ${rawSenderName} 老師！請依照以下格式回覆您的姓名與公務 Email：\n\n` +
                        `範例：\n` +
                        `綁定 王大明 daming@school.tp.edu.tw\n\n` +
                        `（系統將自動在 Teacher_List 名冊中連結您的身分與 LINE 帳號！）`;
      replyForUser(event.replyToken, userId, boundProfile, promptMsg);
      return;
    }
  }

  // 2. 收到 Email 格式輸入 (處理初次綁定 或 修改舊信箱)
  if (emailMatch) {
    const newEmail = emailMatch[0].trim();

    if (boundProfile) {
      const oldEmail = boundProfile.email;
      teacherSheet.getRange(boundProfile.rowIndex, 2).setValue(newEmail);
      teacherSheet.getRange(boundProfile.rowIndex, 3).setValue(true);
      teacherSheet.getRange(boundProfile.rowIndex, 8).setValue(Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd HH:mm:ss'));

      notifyAdmin(`[同仁變更信箱] 【${boundProfile.name}】已將通知信箱更新！\n舊信箱：${oldEmail}\n新信箱：${newEmail}`);

      const successMsg = `【通知信箱已成功修改！】\n\n` +
                         `• 教師姓名：${boundProfile.name} 老師\n` +
                         `• 舊通知信箱：${oldEmail}\n` +
                         `• 新通知信箱 (已更新)：\n  ${newEmail}\n\n` +
                         `未來每日 12:00 與 16:00 的公務半日整理將準時寄送至您的「新信箱」！`;
      
      const updatedProfile = findTeacherProfileByUid(teacherSheet, userId);
      replyForUser(event.replyToken, userId, updatedProfile, successMsg);
      return;
    } else {
      let extractedName = text.replace(newEmail, '').replace(/綁定|修改信箱|修改email|老師|主任|組長|：|:/g, '').trim();
      if (!extractedName) extractedName = rawSenderName;

      const bindResult = bindTeacherProfile(teacherSheet, userId, rawSenderName, extractedName, newEmail);
      const updatedProfile = findTeacherProfileByUid(teacherSheet, userId);
      replyForUser(event.replyToken, userId, updatedProfile, bindResult);
      return;
    }
  }

  // 3. 查詢個人 LINE UID
  if (textLower === '我的id' || textLower === 'id' || textLower === '!id') {
    replyForUser(event.replyToken, userId, boundProfile, `您的 LINE User ID 為：\n${userId}\n\nLINE 暱稱：${rawSenderName}`);
    return;
  }

  // 4. 管理員專屬指令區
  if (isAdmin) {
    if (textLower === '!preview' || textLower === '預覽信件' || textLower === '預覽') {
      const previewMsg = buildAdminPreviewMessage(rawSheet);
      replyForUser(event.replyToken, userId, boundProfile, previewMsg);
      return;
    }

    const toggleMatch = text.match(/^(剔除|保留|切換|!toggle|!exclude|!include)\s*(\d+)$/i);
    if (toggleMatch) {
      const action = toggleMatch[1];
      const itemIndex = parseInt(toggleMatch[2], 10);
      const toggleMsg = toggleReviewItem(itemIndex, action);
      replyForUser(event.replyToken, userId, boundProfile, toggleMsg);
      return;
    }

    if (textLower === '!testmail' || textLower === '試寄給我' || textLower === '寄給我預覽') {
      const adminEmail = boundProfile ? boundProfile.email : Session.getActiveUser().getEmail();
      const sendPreviewMsg = sendPreviewEmailToAdmin(rawSheet, adminEmail);
      replyForUser(event.replyToken, userId, boundProfile, sendPreviewMsg);
      return;
    }

    if (textLower === '!confirm' || textLower === '確認寄出' || textLower === '!send' || textLower === '立即派報') {
      const resultMsg = executeApprovedDigest(rawSheet, teacherSheet);
      replyForUser(event.replyToken, userId, boundProfile, resultMsg);
      return;
    }

    if (textLower === '!status' || textLower === '查詢狀態' || textLower === '狀態') {
      const pendingCount = getPendingMessageCount(rawSheet);
      replyForUser(event.replyToken, userId, boundProfile, `系統狀態報告：\n目前試算表中待派報的群組公務訊息共有 【${pendingCount}】 筆。\n\n您可點選下方【預覽信件】先審核並挑選發布內容！`);
      return;
    }

    if (textLower === '!members' || textLower === '人員名冊' || textLower === '名冊') {
      const memberSummary = getMemberDirectorySummary(memberSheet);
      replyForUser(event.replyToken, userId, boundProfile, memberSummary);
      return;
    }

    if (textLower === '!test' || textLower === '測試ai' || textLower === '測試') {
      replyForUser(event.replyToken, userId, boundProfile, '正在連線測試 Gemini AI 模型中，請稍候約 1~3 秒...');
      const testResult = testGeminiApi();
      notifyAdmin(testResult);
      return;
    }
  } else {
    if (text.startsWith('!') || text.includes('狀態') || text.includes('派報') || text.includes('預覽') || text.includes('名冊')) {
      replyForUser(event.replyToken, userId, boundProfile, '抱歉，此功能為系統管理員專屬權限。一般同仁僅開放「身分綁定 / 修改信箱」功能。');
      return;
    }
  }

  sendAppropriateMenu(event.replyToken, userId, rawSenderName, boundProfile);
}

// ==========================
// 6. 管理員審核與預覽系統核心
// ==========================
function buildAdminPreviewMessage(rawSheet) {
  const pendingRows = getPendingRows(rawSheet);
  if (pendingRows.length === 0) {
    return '目前試算表中【無任何待處理新訊息】，暫無可預覽之信件內容。';
  }

  const digestResult = analyzeAndSummarizeWithGemini(pendingRows);
  if (!digestResult.items || digestResult.items.length === 0) {
    return '目前累積的訊息經 AI 判讀皆屬「召集人內部事項」，預計略過不寄信。';
  }

  const cacheKey = 'ACTIVE_REVIEW_DIGEST';
  const cachedJson = SCRIPT_PROPERTIES.getProperty(cacheKey);
  let reviewState = [];

  if (cachedJson) {
    try {
      reviewState = JSON.parse(cachedJson);
    } catch (e) { reviewState = []; }
  }

  const items = digestResult.items.map((item, idx) => {
    const existing = reviewState.find(r => r.title === item.title);
    return {
      index: idx + 1,
      department: item.department || '學校行政處室',
      category: item.category || '公務通知',
      title: item.title,
      official_sender: item.official_sender || item.sender || '行政同仁',
      points: item.points || [],
      included: existing ? existing.included : true
    };
  });

  SCRIPT_PROPERTIES.setProperty(cacheKey, JSON.stringify(items));
  SCRIPT_PROPERTIES.setProperty('ACTIVE_REVIEW_SUMMARY', digestResult.summary || '半日公務通報');

  return formatPreviewText(items, digestResult.summary);
}

function formatPreviewText(items, summary) {
  let listStr = '';
  let activeCount = 0;

  items.forEach(it => {
    const statusIcon = it.included ? '[發送]' : '[剔除不發]';
    if (it.included) activeCount++;
    listStr += `\n[${it.index}] ${statusIcon} ${it.title}\n` +
               `    單位：${it.department} (${it.official_sender})\n` +
               `    摘要：${(it.points[0] || '').substring(0, 35)}...\n`;
  });

  return `【即將發出之半日通報預覽】\n` +
         `本期總結：${summary}\n` +
         `預計發送：${activeCount} / ${items.length} 則項目\n` +
         `------------------------------------\n` +
         `${listStr}\n` +
         `------------------------------------\n` +
         `[如何選擇發送項目？]\n` +
         `• 輸入「剔除 1」可將第 1 項取消發送\n` +
         `• 輸入「保留 1」可重新納入發送\n\n` +
         `審核完成後，請點擊下方：\n` +
         `• 【試寄給我】：先寄一封到您信箱看版面\n` +
         `• 【確認寄出】：正式發信給全體老師！`;
}

function toggleReviewItem(index, action) {
  const cacheKey = 'ACTIVE_REVIEW_DIGEST';
  const cachedJson = SCRIPT_PROPERTIES.getProperty(cacheKey);
  if (!cachedJson) return '尚未建立預覽清單，請先點選【預覽信件】！';

  let items = JSON.parse(cachedJson);
  const target = items.find(it => it.index === index);
  if (!target) return `找不到編號為 [${index}] 的項目，目前共有 1 ~ ${items.length} 項。`;

  if (action.includes('剔除') || action.includes('exclude')) {
    target.included = false;
  } else if (action.includes('保留') || action.includes('include')) {
    target.included = true;
  } else {
    target.included = !target.included;
  }

  SCRIPT_PROPERTIES.setProperty(cacheKey, JSON.stringify(items));
  const summary = SCRIPT_PROPERTIES.getProperty('ACTIVE_REVIEW_SUMMARY') || '';

  return `已更新第 [${index}] 項狀態為：${target.included ? '[發送]' : '[剔除不發]'}\n\n` +
         formatPreviewText(items, summary);
}

function sendPreviewEmailToAdmin(rawSheet, adminEmail) {
  const cacheKey = 'ACTIVE_REVIEW_DIGEST';
  const cachedJson = SCRIPT_PROPERTIES.getProperty(cacheKey);
  if (!cachedJson) return '尚未建立預覽清單，請先點選【預覽信件】！';

  const items = JSON.parse(cachedJson);
  const approvedItems = items.filter(it => it.included);

  if (approvedItems.length === 0) {
    return '目前所有項目皆被設為「剔除」，無任何內容可寄送！請至少保留一項。';
  }

  const summary = SCRIPT_PROPERTIES.getProperty('ACTIVE_REVIEW_SUMMARY') || '管理員預覽信件';
  const digestData = { summary: summary, items: approvedItems };
  const htmlBody = generateHtmlEmail('管理員審核預覽', digestData);
  const plainText = generatePlainTextEmail('管理員審核預覽', digestData);
  const pendingRows = getPendingRows(rawSheet);
  const attachments = collectAttachments(pendingRows);

  try {
    GmailApp.sendEmail(adminEmail, `【管理員審核預覽】竹林召集人半日通報 (${approvedItems.length}則)`, plainText, {
      htmlBody: htmlBody,
      attachments: attachments,
      name: '竹林召集人通報秘書 (預覽)'
    });
    return `預覽信已成功發送至您的信箱：\n${adminEmail}\n\n請開啟手機 Gmail 查看排版，確認無誤後點擊【確認寄出】！`;
  } catch (err) {
    return `試寄失敗：${err.message}`;
  }
}

function executeApprovedDigest(rawSheet, teacherSheet) {
  const pendingRows = getPendingRows(rawSheet);
  if (pendingRows.length === 0) {
    return '目前無待發送的新訊息。';
  }

  const cacheKey = 'ACTIVE_REVIEW_DIGEST';
  const cachedJson = SCRIPT_PROPERTIES.getProperty(cacheKey);
  let approvedItems = [];
  let summary = '本半日公務通報';

  if (cachedJson) {
    const items = JSON.parse(cachedJson);
    approvedItems = items.filter(it => it.included);
    summary = SCRIPT_PROPERTIES.getProperty('ACTIVE_REVIEW_SUMMARY') || summary;
  } else {
    const digestResult = analyzeAndSummarizeWithGemini(pendingRows);
    approvedItems = digestResult.items || [];
    summary = digestResult.summary || summary;
  }

  if (approvedItems.length === 0) {
    return '所有項目皆已被管理員剔除或為召集人內部事項，本次不寄發信件。';
  }

  const digestData = { summary: summary, items: approvedItems };
  const htmlBody = generateHtmlEmail('半日公務彙整', digestData);
  const plainText = generatePlainTextEmail('半日公務彙整', digestData);
  const attachments = collectAttachments(pendingRows);
  const recipients = getTeacherEmailList(teacherSheet);

  if (recipients.length === 0) {
    return '派報中斷：Teacher_List 工作表中沒有啟用 (TRUE) 的教師 Email！';
  }

  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd');
  const subject = `【竹林召集人通報】${todayStr} 重點訊息整理`;

  try {
    GmailApp.sendEmail(recipients.join(','), subject, plainText, {
      htmlBody: htmlBody,
      attachments: attachments,
      name: '竹林召集人通報秘書'
    });

    SCRIPT_PROPERTIES.deleteProperty(cacheKey);
    SCRIPT_PROPERTIES.deleteProperty('ACTIVE_REVIEW_SUMMARY');
    markRowsAsProcessed(rawSheet, pendingRows, 'ADMIN_APPROVED_SENT');

    return `【派報完成！】已成功將 ${approvedItems.length} 則經您審核確認之公務通報發送給全體 ${recipients.length} 位老師！`;
  } catch (e) {
    return `寄件失敗：${e.message}`;
  }
}

function getPendingRows(rawSheet) {
  const data = rawSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

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
  return pendingRows;
}

function collectAttachments(rows) {
  const attachments = [];
  const processedFileIds = new Set();
  rows.forEach(row => {
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
  return attachments;
}

// ==========================
// 7. 身分綁定核心資料庫處理函式
// ==========================
function findTeacherProfileByUid(sheet, userId) {
  if (!sheet || !userId) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === userId) {
      return {
        rowIndex: i + 1,
        name: data[i][0],
        email: data[i][1],
        isActive: (data[i][2] === true || data[i][2] === 'TRUE'),
        department: data[i][3] || '一般科別',
        notes: data[i][4],
        userId: data[i][5],
        displayName: data[i][6],
        boundAt: data[i][7]
      };
    }
  }
  return null;
}

function bindTeacherProfile(sheet, userId, displayName, nameInput, emailInput) {
  if (!sheet) return '系統尚未建立 Teacher_List 工作表，請聯繫管理員！';

  const data = sheet.getDataRange().getValues();
  let matchedRow = -1;
  let department = '';
  let officialName = nameInput;

  const cleanInputName = nameInput.replace(/\s+|老師/g, '').toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const sheetName = String(data[i][0] || '').replace(/\s+|老師/g, '').toLowerCase();
    const sheetEmail = String(data[i][1] || '').trim().toLowerCase();

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
    sheet.getRange(matchedRow, 2).setValue(emailInput);
    sheet.getRange(matchedRow, 3).setValue(true);
    sheet.getRange(matchedRow, 6).setValue(userId);
    sheet.getRange(matchedRow, 7).setValue(displayName);
    sheet.getRange(matchedRow, 8).setValue(timeStr);

    notifyAdmin(`[同仁身分綁定成功] 【${officialName}】已完成 LINE 綁定！\nEmail：${emailInput}\nLINE 暱稱：${displayName}`);

    return `【身分綁定成功！】\n\n` +
           `• 教師姓名：${officialName}\n` +
           `• 通知信箱：${emailInput}\n` +
           `• 所屬處室/科別：${department}\n` +
           `• 綁定 LINE 暱稱：${displayName}\n\n` +
           `系統已將您的 LINE 與教師通訊錄成功串接！\n` +
           `每日 12:00 與 16:00 的公務半日整理將準時寄送至您的 Email。\n\n` +
           `若未來欲更換信箱，隨時在此點選「修改Email」即可！`;
  } else {
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

    notifyAdmin(`[新同仁自助註冊] 【${nameInput}】已登記至名冊！\nEmail：${emailInput}\nLINE 暱稱：${displayName}`);

    return `【註冊與綁定完成！】\n\n` +
           `• 教師姓名：${nameInput}\n` +
           `• 通知信箱：${emailInput}\n` +
           `• 綁定 LINE 暱稱：${displayName}\n\n` +
           `系統已為您在教師名冊中建立檔案並開啟通知！每日 12:00 與 16:00 派報將寄送至該信箱。`;
  }
}

function sendWelcomeAndBindingPrompt(replyToken, displayName, boundProfile, userId) {
  if (boundProfile) {
    const welcomeBackText = `您好 ${boundProfile.name} 老師！歡迎回到「竹林召集人通報秘書」！\n\n` +
                            `您目前已綁定通知信箱：\n${boundProfile.email}\n\n` +
                            `若需修改接收信箱，請直接點擊下方【修改Email】！`;
    replyForUser(replyToken, userId, boundProfile, welcomeBackText);
  } else {
    const welcomeText = `您好 ${displayName} 老師！歡迎加入「竹林召集人通報秘書」！\n\n` +
                        `本系統每日於 12:00 與 16:00 自動彙整召集人群組之重要行政公文、段考範圍與活動調移，派送至全體同仁信箱。\n\n` +
                        `[請先進行身分綁定]：\n` +
                        `請直接依照下方範例格式回覆您的「姓名」與「公務Email」：\n` +
                        `範例：綁定 王大明 daming@school.tp.edu.tw\n\n` +
                        `點擊下方「身分綁定」即可開始！`;
    replyForUser(replyToken, userId, boundProfile, welcomeText);
  }
}

// ==========================
// 8. 定時自動派報 (12:00 / 16:00 觸發)
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

  const pendingRows = getPendingRows(rawSheet);
  if (pendingRows.length === 0) {
    console.log(`[${timeSlotLabel}] 該時段完全無待處理新訊息，系統略過不寄信。`);
    return;
  }

  const cacheKey = 'ACTIVE_REVIEW_DIGEST';
  const cachedJson = SCRIPT_PROPERTIES.getProperty(cacheKey);
  let approvedItems = [];
  let summary = '本半日公務通報';

  if (cachedJson) {
    const items = JSON.parse(cachedJson);
    approvedItems = items.filter(it => it.included);
    summary = SCRIPT_PROPERTIES.getProperty('ACTIVE_REVIEW_SUMMARY') || summary;
  } else {
    const digestResult = analyzeAndSummarizeWithGemini(pendingRows);
    approvedItems = digestResult.items || [];
    summary = digestResult.summary || summary;
  }

  if (approvedItems.length === 0) {
    console.log(`[${timeSlotLabel}] 訊息皆屬召集人內部工作或已被管理員剔除，系統不寄信。`);
    markRowsAsProcessed(rawSheet, pendingRows, `${timeSlotLabel}_SKIPPED`);
    return;
  }

  const digestData = { summary: summary, items: approvedItems };
  const htmlBody = generateHtmlEmail(timeSlotLabel, digestData);
  const plainText = generatePlainTextEmail(timeSlotLabel, digestData);
  const attachments = collectAttachments(pendingRows);
  const recipients = getTeacherEmailList(teacherSheet);

  if (recipients.length === 0) {
    notifyAdmin('[警告] 派報中斷：Teacher_List 工作表中沒有啟用 (TRUE) 的教師 Email！');
    return;
  }

  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd');
  const subject = `【竹林召集人通報】${todayStr} ${timeSlotLabel} 重點訊息整理`;

  try {
    GmailApp.sendEmail(recipients.join(','), subject, plainText, {
      htmlBody: htmlBody,
      attachments: attachments,
      name: '竹林召集人通報秘書'
    });

    if (digestSheet) {
      digestSheet.appendRow([
        `DIG-${Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd-HHmm')}`,
        new Date(),
        timeSlotLabel,
        approvedItems.map(it => it.department || it.category).join('、'),
        approvedItems[0].title,
        summary,
        attachments.length,
        recipients.length,
        'SUCCESS'
      ]);
    }

    SCRIPT_PROPERTIES.deleteProperty(cacheKey);
    SCRIPT_PROPERTIES.deleteProperty('ACTIVE_REVIEW_SUMMARY');
    notifyAdmin(`[派報完成] 已完成 [${timeSlotLabel}] 訊息派報，共發布 ${approvedItems.length} 則公務公告，發送給 ${recipients.length} 位老師！`);
  } catch (e) {
    console.error('Gmail 寄送失敗: ' + e.message);
    notifyAdmin(`[錯誤] 寄發 Email 失敗：${e.message}`);
  }

  markRowsAsProcessed(rawSheet, pendingRows, `${timeSlotLabel}_SENT`);
}

// ==========================
// 9. Gemini API 智慧分類與單位身分發文
// ==========================
function analyzeAndSummarizeWithGemini(rows) {
  const apiKey = getActiveGeminiApiKey();
  if (!apiKey) return fallbackRuleFilter(rows);

  const promptText = `
你是一位專業的高中行政教務秘書。以下是學校「召集人群組」今日由各處室幹事、主任或召集人發出的訊息清單：
${JSON.stringify(rows.map(r => ({
  sender: r.sender,
  department: r.department || '未指定處室',
  type: r.type,
  content: r.content,
  time: Utilities.formatDate(new Date(r.timestamp), 'GMT+8', 'HH:mm')
})))}

【重要任務與發文身分規範】：
1. 判斷每一則訊息「是否需要公告給全體老師」：
   - 需公告：全校時程、段考範圍與核對、公文更正、學務活動調課、榮譽榜、人事退休福利、健檢等。
   - 僅限召集人內部事項【剔除不收錄】：例如催繳審題表、召集人會議領取試卷、私訊答覆等。
2. 【處室單位名義連結】：
   - 若訊息發布者的文字未自帶單位名稱，請務必直接連結其後端對應的處室單位（如「教務處」、「學務處」、「人事室」），以「單位名義」擬定正式標題（例如：【教務處公告】第二次段考範圍核對、或【人事室公告】教職員健檢注意事項）。
3. 請勿在輸出文字中使用會產生編碼衝突的特殊 Emoji 符號，保持純文字乾淨專業。
4. 輸出嚴格的 JSON 格式（不要包含 markdown 代碼區塊），結構如下：
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
        generationConfig: { responseMimeType: 'application/json' }
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
// 10. 發言人名冊管理 (Member_Directory)
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
  if (!sheet) return '尚未建立 Member_Directory 工作表。';
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return '目前群組名冊中尚未有任何人員發言紀錄。';

  let total = data.length - 1;
  let unsetCount = 0;
  let listStr = '';

  for (let i = 1; i < data.length; i++) {
    const name = data[i][1] || '未知名稱';
    const dept = data[i][2] || '[未設定處室]';
    if (!data[i][2]) unsetCount++;
    listStr += `• ${name} 處室：${dept}\n`;
  }

  return `【群組發言人與單位名冊】\n總計已記錄 ${total} 位發言人（其中 ${unsetCount} 位待設定處室）：\n\n${listStr}\n提示：請至 Google 試算表「Member_Directory」分頁填寫對應處室，系統將自動以該處室名義發文！`;
}

// ==========================
// 11. 權限隔離選單與 Quick Reply
// ==========================
function getQuickReplyForUser(userId, boundProfile) {
  const isAdmin = isUserAdmin(userId);

  if (isAdmin) {
    return {
      items: [
        { type: 'action', action: { type: 'message', label: '預覽信件', text: '!preview' } },
        { type: 'action', action: { type: 'message', label: '試寄給我', text: '!testmail' } },
        { type: 'action', action: { type: 'message', label: '確認寄出', text: '!confirm' } },
        { type: 'action', action: { type: 'message', label: '查詢狀態', text: '!status' } },
        { type: 'action', action: { type: 'message', label: '人員名冊', text: '!members' } },
        { type: 'action', action: { type: 'message', label: boundProfile ? '修改Email' : '身分綁定', text: boundProfile ? '修改Email' : '身分綁定' } },
        { type: 'action', action: { type: 'message', label: '我的ID', text: '!id' } }
      ]
    };
  } else {
    const bindLabel = boundProfile ? '修改Email' : '身分綁定';
    const bindText = boundProfile ? '修改Email' : '身分綁定';
    return {
      items: [
        { type: 'action', action: { type: 'message', label: bindLabel, text: bindText } },
        { type: 'action', action: { type: 'message', label: '我的ID', text: '!id' } }
      ]
    };
  }
}

function replyForUser(replyToken, userId, boundProfile, text) {
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
          quickReply: getQuickReplyForUser(userId, boundProfile)
        }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('replyForUser 失敗: ' + e.message);
  }
}

function sendAppropriateMenu(replyToken, userId, rawSenderName, boundProfile) {
  const isAdmin = isUserAdmin(userId);

  if (isAdmin) {
    const adminMenuText = `【召集人通報秘書・管理員審核控制台】\n` +
                          `您好 管理員！您可以先預覽、挑選並決定要發出的公告：\n\n` +
                          `• 預覽信件 (!preview)：查看即將發出之公告清單並選擇項目\n` +
                          `• 試寄給我 (!testmail)：將挑選好的內容先寄給您看版面\n` +
                          `• 確認寄出 (!confirm)：確認無誤後，正式群發給全體老師\n` +
                          `• 狀態查詢 (!status)：查看目前待派報訊息筆數\n` +
                          `• 人員名冊 (!members)：查看群組發言人處室設定\n` +
                          `${boundProfile ? '• 修改Email：變更您的個人接收信箱' : '• 身分綁定：綁定您的個人教師資料'}\n` +
                          `• 我的ID (!id)：查詢您的 33 碼 LINE User ID`;

    replyForUser(replyToken, userId, boundProfile, adminMenuText);
  } else {
    let teacherMenuText = '';
    if (boundProfile) {
      teacherMenuText = `【竹林召集人通報秘書・同仁專區】\n` +
                        `您好 ${boundProfile.name} 老師！\n\n` +
                        `您目前已完成身分綁定：\n` +
                        `• 通知信箱：${boundProfile.email}\n` +
                        `• 所屬處室：${boundProfile.department}\n\n` +
                        `若需更換公務通知信箱，請點擊下方【修改Email】！\n` +
                        `（每日 12:00 與 16:00 之半日公務通報將準時寄至您的信箱）`;
    } else {
      teacherMenuText = `【竹林召集人通報秘書・同仁專區】\n` +
                        `您好 ${rawSenderName} 老師！\n\n` +
                        `本機器人專門整理「召集人群組」之公務通知，並定時發送至同仁信箱。\n` +
                        `您目前尚未綁定教師信箱，請點擊下方【身分綁定】依引導設定！`;
    }

    replyForUser(replyToken, userId, boundProfile, teacherMenuText);
  }
}

// ==========================
// 12. 輔助函式 (Drive存檔、LINE Reply、純文字與 HTML 電子郵件產生)
// ==========================

/**
 * 徹底解決「問號亂碼」的核心：完整標準 HTML5 + UTF-8 標頭 + 移除任何易亂碼的 4-byte Emoji
 */
function generateHtmlEmail(timeSlotLabel, digest) {
  let cardsHtml = '';
  digest.items.forEach(item => {
    let listHtml = '';
    item.points.forEach(pt => {
      listHtml += `<li style="margin-bottom: 6px; color: #334155; line-height: 1.6;">${escapeHtml(pt)}</li>`;
    });

    const deptTag = item.department ? `<span style="display: inline-block; background-color: #1e40af; color: #ffffff; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; margin-right: 6px;">${escapeHtml(item.department)}</span>` : '';
    const catTag = item.category ? `<span style="display: inline-block; background-color: #e2e8f0; color: #334155; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; margin-right: 8px;">${escapeHtml(item.category)}</span>` : '';
    const senderDisplay = escapeHtml(item.official_sender || item.sender || '學校行政處室');

    cardsHtml += `
      <div style="background-color: #ffffff; border: 1px solid #cbd5e1; border-left: 4px solid #1e40af; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
        <div style="margin-bottom: 10px; font-size: 12px; color: #64748b;">
          ${deptTag}
          ${catTag}
          <span style="color: #475569;">發布單位/人員：<strong>${senderDisplay}</strong></span>
        </div>
        <h3 style="margin: 0 0 12px 0; font-size: 17px; color: #0f172a; font-weight: bold; line-height: 1.4;">${escapeHtml(item.title)}</h3>
        <ul style="margin: 0; padding-left: 20px; font-size: 14px;">
          ${listHtml}
        </ul>
      </div>
    `;
  });

  const escapedSummary = escapeHtml(digest.summary || '今日半日重要公務通報');

  return `<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>竹林召集人群組 半日訊息整理</title>
</head>
<body style="margin: 0; padding: 20px; background-color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Microsoft JhengHei', sans-serif; color: #334155; -webkit-font-smoothing: antialiased;">
  <!-- 防預覽亂碼隱藏前綴 (Preheader) -->
  <div style="display: none; max-height: 0px; overflow: hidden; font-size: 1px; line-height: 1px; color: #ffffff;">
    ${escapedSummary}
  </div>

  <div style="max-width: 680px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">
    <!-- 頂端校務標題列 -->
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #0369a1 100%); color: #ffffff; padding: 24px; text-align: center;">
      <h1 style="margin: 0; font-size: 21px; font-weight: bold; letter-spacing: 0.5px;">竹林召集人群組 半日訊息整理</h1>
      <p style="margin: 8px 0 0 0; font-size: 13px; opacity: 0.9;">發布時段：${escapeHtml(timeSlotLabel)} ｜ 公務通知已依處室整合派發</p>
    </div>

    <div style="padding: 24px;">
      <!-- 摘要速覽區塊 -->
      <div style="background-color: #eff6ff; border-left: 4px solid #2563eb; border-radius: 4px; padding: 14px 16px; margin-bottom: 24px;">
        <div style="font-size: 14px; font-weight: bold; color: #1e40af; margin-bottom: 4px;">【重點速覽】</div>
        <div style="font-size: 14px; color: #1e3a8a; line-height: 1.6;">${escapedSummary}</div>
      </div>

      <!-- 公告卡片列表 -->
      ${cardsHtml}

      <!-- 底部附件提示與版權 -->
      <div style="margin-top: 24px; padding: 16px; background-color: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; text-align: center;">
        <div style="font-size: 13px; font-weight: bold; color: #0369a1; margin-bottom: 4px;">[附件下載說明]</div>
        <div style="font-size: 12px; color: #64748b; line-height: 1.5;">
          相關公文 PDF 與通知圖片已隨信夾帶於下方附件，請同仁查閱下載。<br>
          本信件由「竹林召集人通報秘書」系統自動整理發送。
        </div>
      </div>
    </div>
  </div>
</body>
</html>`;
}

/**
 * 乾淨純文字備援字串 (防止舊版客戶端或預覽 Snippet 抓出 Emoji 亂碼)
 */
function generatePlainTextEmail(timeSlotLabel, digest) {
  let plain = `【竹林召集人群組 半日訊息整理】\n` +
              `發布時段：${timeSlotLabel}\n\n` +
              `【重點速覽】\n${digest.summary}\n\n` +
              `----------------------------------------\n`;

  digest.items.forEach((item, idx) => {
    plain += `\n[${idx + 1}] ${item.title}\n` +
             `發布單位：${item.department} (${item.official_sender || item.sender})\n`;
    item.points.forEach(pt => {
      plain += `• ${pt}\n`;
    });
  });

  plain += `\n----------------------------------------\n` +
           `[附件說明] 相關公文 PDF 與通知圖片已夾帶於信件附件，請查閱下載。\n` +
           `本信件由召集人通報秘書系統自動整理發送。`;

  return plain;
}

function escapeHtml(text) {
  if (!text) return '';
  return text.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

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
          text: message
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
