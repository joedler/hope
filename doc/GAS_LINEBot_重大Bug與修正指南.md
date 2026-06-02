# GAS + LINE Bot 三大必修與三大進階 Bug 修正指南

> 每次新建 Google Apps Script + LINE Messaging API Webhook 專案，**必須套用這幾個修正**，否則 BOT 必定無法正常運作。

---

## Bug 1：302 Found — ContentService 造成重新導向

### 症狀
LINE Developers Console 驗證 Webhook 時永遠顯示：
```
The webhook returned an HTTP status code other than 200.(302 Found)
```

### 原因
`return ContentService.createTextOutput("OK")` 在新版 GAS 機制中會觸發跨網域 302 重新導向。瀏覽器會自動跟跳，但 LINE 的 Webhook 驗證伺服器拒絕跳轉，直接判定失敗。

### 修正

```javascript
// ❌ 錯誤寫法
function doPost(e) {
  // ... 處理邏輯 ...
  return ContentService.createTextOutput('OK');
}

// ✅ 正確寫法
function doPost(e) {
  // ... 處理邏輯 ...
  return; // 隱式回傳 200 OK 空白內容
}

// ✅ 同樣加上 doGet（某些驗證會打 GET）
function doGet(e) {
  return;
}
```

---

## Bug 2：已讀不回 — const/let 造成跨檔案全域變數失效

### 症狀
Webhook 驗證成功（過了 302），但在 LINE 群組或私訊輸入任何指令都沒有回應（已讀不回）。

### 原因
GAS V8 引擎中，用 `const` 或 `let` 宣告的**頂層變數**只在該檔案內有效，其他 `.gs` / `.js` 檔案無法存取。導致 `Config.js` 裡的 `CHANNEL_TOKEN`、`SHEETS` 等設定在其他檔案中是 `undefined`，呼叫 LINE API 時直接當機。

### 修正

```javascript
// ❌ 錯誤寫法（Config.js）
const SHEETS = { STUDENTS: '學員名單' };
const LINE_TOKEN = '...';

// ✅ 正確寫法（Config.js）
// 跨檔案全域變數必須用 var
var SHEETS = { STUDENTS: '學員名單' };
var LINE_TOKEN = '...';
```

**變數作用域規則**：
* 函式內的局部變數：可以使用 `const` / `let`（保持 JS 良好習慣）
* 檔案頂層、需要跨檔案存取的常數或設定值：**必須使用 `var`**

---

## Bug 3：UrlFetchApp 無外部連線權限

### 症狀
變數正確了，但 BOT 依然沉默。透過 Spreadsheet 寫 Log 才發現：
```
UrlFetchApp.fetch を呼び出す権限がありません
```
（沒有呼叫外部服務的權限）

### 原因
初次授權時程式碼還沒有明確使用外部連線，Google 只核發了 Spreadsheet 權限。之後加了程式碼也不會自動補授權。

### 修正

#### Step 1：在 `appsscript.json` 強制宣告 oauthScopes

```json
{
  "timeZone": "Asia/Taipei",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/drive"
  ],
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  }
}
```

#### Step 2：建立 forceAuth 函式並手動執行一次

```javascript
function forceAuth() {
  // 強制觸發外部連線與雲端硬碟授權視窗
  UrlFetchApp.fetch('https://www.google.com');
  DriveApp.getRootFolder(); 
  Logger.log('所有外部資源授權完成');
}
```
在 GAS 編輯器選 `forceAuth` → 執行 → 在彈出的授權視窗點「允許」。

---

## Bug 4：Google Drive 授權快取拒絕漏洞

### 症狀
在程式中調用 `DriveApp.getFileById()` 抓取選單圖片時，程式報出 `You do not have permission to call DriveApp...`，但 GAS 編輯器卻不彈出授權同意視窗。

### 原因
GAS 的 Scopes 快取機制偶爾會失效，若未在程式碼頂層顯式調用 DriveApp 的根節點方法（例如 `getRootFolder()`），GAS 會認為不需重新取得使用者同意。

### 修正
在強制授權函式 `forceAuth()` 內，一定要加入對 `DriveApp` 的根節點呼叫：
```javascript
function forceAuth() {
  UrlFetchApp.fetch('https://www.google.com'); // 觸發 external_request Scopes
  DriveApp.getRootFolder();                    // 觸發 drive.readonly / drive Scopes
}
```

---

## Bug 5：試算表隱形空白與大小寫匹配失效 (ORM/CRUD 致命缺陷)

### 症狀
用戶的 LINE UID 確實在教職員表 (`Staff`) 中，但 Webhook 比對卻永遠回傳 `guest` 訪客，或誤判為 `member` 學員。

### 原因
1. 人類在試算表中貼入 UID 時，極易夾帶**看不見的前後空白字元、換行符或 Tab 鍵**（例如 `" U123... "`），導致 JS 的 `===` 嚴格比對失敗。
2. 狀態欄位輸入了 `Active`、`Admin` 首字母大寫，而後端程式碼寫死小寫比對 `status === 'active'` 導致無法解析。

### 修正
在所有 Sheets ORM 比對邏輯（例如 `getRow`、`resolveRole`）與屬性提取時，**一律強制對 UID 做 `.trim()`，對 status/role 做 `.trim().toLowerCase()`**：
```typescript
// ✅ 頂級防呆對齊
const cleanLineUid = String(row.line_uid).trim();
const cleanStatus = String(row.status).trim().toLowerCase();
```

---

## Bug 6：LINE 實體圖文選單狀態非即時同步

### 症狀
在試算表中手動修改用戶的身分為管理員後，用戶在 LINE 對話框發送訊息，底部的 Rich Menu 鍵盤依然停留在學員或訪客狀態，沒有即時切換。

### 原因
1. LINE Webhook 的預設流程中，**接收到文字訊息時，LINE 伺服器並不會主動為該用戶重新刷新或關聯 Rich Menu**。
2. 用戶輸入非標準指令時，若無攔截機制，會走入預設的學員服務防呆引導，導致用戶誤以為身分仍然是學員。

### 修正
1. 在 Webhook (`LineHandler.ts`) 頂層實作強制的 **`更新`** 與 **`同步選單`** 命令攔截器。
2. 收到「更新」時，主動重新讀取資料庫，並發送 API 給 LINE 強制為其手機更新圖文選單關聯：
```typescript
if (cleanText === '更新' || cleanText === '同步選單') {
  // 1. 重新解析身分角色
  const role = resolveRole(userId); 
  // 2. 主動關聯 Rich Menu
  LineRichMenu.link(userId, role);
}
```

---

## 完整的新專案 Checklist

每次建新 GAS + LINE Bot 專案，照順序執行：

- [ ] `appsscript.json` 加入 `oauthScopes`，確保包含 `drive.readonly` 與 `script.external_request`
- [ ] 所有跨檔案全域變數改用 `var`
- [ ] 所有與試算表資料的 `line_uid` 比對處強制加上 `.trim()`，`status/role` 比對處加上 `.trim().toLowerCase()` 防呆
- [ ] `doPost` 和 `doGet` 結尾改為 `return;`
- [ ] Webhook 頂層預留 `診斷` 與 `更新` 兩個高維度調錯語意指令
- [ ] `clasp push --force`
- [ ] GAS 編輯器執行 `forceAuth()` 取得外部連線與雲端硬碟授權
- [ ] **建立全新部署**（不要更新舊部署）→ 類型：Web App → 存取：所有人（匿名）
- [ ] 確認 Webhook URL 是 `/exec` 結尾（不是 `/dev`）
- [ ] 填入 LINE Developers Console → 驗證與啟用 Webhook
