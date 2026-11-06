# 🏫 學校召集人群組 LINE Bot 智慧訊息彙整與定時郵件派報系統

本專案透過 LINE Messaging API 監聽學校「召集人群組」，將日常發布之重要行政通知、公文附件（PDF）、活動調課圖卡與榮譽榜即時轉存至 **Google 試算表（Google Sheets）** 與 **Google 雲端硬碟（Google Drive）**。每天 **12:00** 與 **16:00** 自動透過 **Gemini AI** 語意過濾「僅限召集人內部事項」，將需公告之重點條列摘要，並直接夾帶附件發送至全體教師 Email 信箱。

---

## 📂 專案目錄結構

```text
linebot-school-notifier/
├── README.md                      # 本說明文件（含完整部署指南）
├── linebot_school_template.xlsx   # Google 試算表範本（含教師通訊錄、原始日誌、派報紀錄）
├── src/
│   └── Code.gs                    # Google Apps Script (GAS) 核心程式碼
├── docs/
│   ├── message_schema.md          # 訊息欄位規格、連續圖文關聯邏輯與篩選準則
│   └── system_architecture.md     # 系統架構評估與序列圖
└── templates/
    └── email_digest_preview.html  # 派發給老師的 HTML 郵件版型預覽
```

---

## 🚀 完整部署操作步驟指南 (Step-by-Step)

本系統採用 **Google Apps Script (GAS) 純雲端架構**，無須租用主機，完全免費且與 Google 服務原生整合。請依照下列步驟進行設定：

### 步驟 1：建立 Google 試算表 (Google Sheets)
1. 前往 [Google 雲端硬碟](https://drive.google.com/)。
2. 將本專案內的 `linebot_school_template.xlsx` 上傳至雲端硬碟。
3. 點開後選擇「以 Google 試算表開啟」，或將其另存為 Google 試算表。
4. 在 **`Teacher_List`** 分頁中，填入學校老師的姓名、Email 信箱，並確保「接收通知」欄位設定為 `TRUE`。

---

### 步驟 2：建立 Google Drive 附件資料夾
1. 在 Google 雲端硬碟建立一個新資料夾（例如命名為：`學校召集人通知附件庫`）。
2. 點入該資料夾，複製網址列中 `folders/` 後方的英數代碼，這就是您的 **`DRIVE_FOLDER_ID`**。
   > 例如網址為 `https://drive.google.com/drive/folders/1a2b3c4d5e6f7...`，代碼即為 `1a2b3c4d5e6f7...`。

---

### 步驟 3：在 Google 試算表建立 Apps Script 專案
1. 在剛才建立的 Google 試算表上方選單，點選 **「擴充功能」 ➡️ 「Apps Script」**。
2. 將專案名稱重新命名為：`召集人LINE智慧通報秘書`。
3. 清空左側 `程式碼.gs` 的內容，將本專案 `src/Code.gs` 的完整程式碼複製並貼上，按上方「儲存 (Ctrl+S)」。

---

### 步驟 4：設定「指令碼屬性 (Script Properties)」
為了資安與維護便利，金鑰與 ID 均統一放在屬性中：
1. 在 Apps Script 左側側邊欄點選 **⚙️「專案設定」**。
2. 往下滾動至 **「指令碼屬性」** 區塊，點擊 **「新增指令碼屬性」**，依序新增以下項目：

| 屬性名稱 (Key) | 說明與範例值 | 是否必填 |
| :--- | :--- | :--- |
| `LINE_CHANNEL_ACCESS_TOKEN` | LINE Developers 後台核發的 Channel Access Token (長權杖) | **必填** |
| `GEMINI_API_KEY` | [Google AI Studio](https://aistudio.google.com/) 免費取得的 API Key | **必填** |
| `DRIVE_FOLDER_ID` | 步驟 2 取得的 Google Drive 資料夾代碼 | **必填** |
| `ADMIN_USER_ID` | 管理員個人的 LINE User ID（以 `U` 開頭之 33 碼） | 選填（建議設定） |

---

### 步驟 5：將 GAS 專案部署為 Web 應用程式 (Web App)
1. 點擊右上角藍色 **「部署」 ➡️ 「新增部署」**。
2. 點選左上角齒輪 ⚙️ 圖示，選擇 **「網頁應用程式 (Web app)」**。
3. 填寫部署資訊：
   - **說明**：`LINE Bot Webhook v1`
   - **執行身分**：選擇 **「我 (您的 Google 帳號)」**
   - **誰可以存取**：務必選擇 **「所有人 (Anyone)」**（這樣 LINE 伺服器才能把訊息送進來）
4. 點擊 **「部署」**，首次執行會要求授權存取 Google 試算表、雲端硬碟與 Gmail，依指示點擊允許。
5. 部署完成後，複製畫面上顯示的 **「網頁應用程式網址 (Web app URL)」**。

---

### 步驟 6：至 LINE Developers 設定 Webhook
1. 前往 [LINE Developers Console](https://developers.line.biz/)，點入您的 Messaging API Channel。
2. 切換至 **「Messaging API」** 標籤頁：
   - 找到 **Webhook URL**，將步驟 5 複製的網頁應用程式網址貼上，點擊 **Update**。
   - 開啟 **Use webhook** 開關（轉為綠色開啟狀態）。
   - 點擊 **Verify** 按鈕，若顯示 `Success` 表示連線正常。
3. 往下拉找到 **LINE Official Account features**：
   - 點選進入 LINE Official Account Manager（LINE 官方帳號管理後台）。
   - 在「回應設定」中：
     - **聊天回應模式**：選擇「Bot (聊天機器人)」。
     - **自動回應訊息**：關閉（避免機器人在群組任意亂回覆）。
     - **加入群組或多人聊天室**：開啟「允許加入群組」。

---

### 步驟 7：設定每日 12:00 與 16:00 定時觸發條件 (Triggers)
讓系統在每天固定時間自動整理並發送 Email：
1. 回到 Google Apps Script 畫面，點選左側側邊欄的 **⏰「觸發條件 (Triggers)」** 圖示。
2. 點選右下角 **「+ 新增觸發條件」**：
   - **第一組（中午派報）**：
     - 選擇要執行的功能：`runDigest1200`
     - 選取活動來源：**「時間導向」**
     - 選取類型：**「計時器 (以日為單位)」**
     - 選取時段：**「上午 11:00 至中午 12:00」** 或 **「中午 12:00 至下午 1:00」**
   - **第二組（下午派報）**：
     - 選擇要執行的功能：`runDigest1600`
     - 選取活動來源：**「時間導向」**
     - 選取時段：**「計時器 (以日為單位)」**
     - 選取時段：**「下午 4:00 至下午 5:00」**

---

## 🛠️ 管理員專用私訊功能與指令

只要加該 LINE Bot 官方帳號為好友，管理員便能享受以下專屬功能：

1. **查詢自己的 LINE User ID**：
   - 私訊傳送：`我的ID` 或 `id`
   - 機器人會立即回覆您的 33 碼 User ID（如 `U1a2b3c4...`），可直接填回步驟 4 的 `ADMIN_USER_ID`。
2. **查詢目前待處理筆數**：
   - 管理員私訊傳送：`!status`
   - 機器人回覆目前試算表中累積待發送的訊息數量。
3. **即時手動強制發送測試**：
   - 管理員私訊傳送：`!send`
   - 機器人會立即執行一次整理並寄送 Email，不必等待定時排程，隨時可用於測試驗證。
