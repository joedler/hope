# 空空BOT 功能整理與全般優化建議報告

本專案是一個基於 **Google Apps Script (GAS) + LINE Messaging API** 的管理系統，主要提供講師課程登錄、請假、行政代操作、財務與免稅額試算、PDF 單據自動生成（收據/領據）與 Email 寄送、以及稅務/會計日記帳管理等功能。

以下為針對本專案的完整功能整理、環境優化、程式碼優化建議與提問清單。

---

## 1. 系統功能全貌整理

本系統的功能可歸納為以下 6 大模組：

| 模組 | 指令 / 觸發詞 | 流程與行為簡述 | 相關服務模組 |
| :--- | :--- | :--- | :--- |
| **講師功能** | `登錄` / `登記` | 彈出 Flex 學生與課程選單 → 輸入時間格式 (`MM/DD HHMM HHMM`) → 檢查時段重疊 → 寫入「授課紀錄」。 | `Core_Service.js` <br> `UI_Utils.js` |
| | `預排` | 同登錄流程 → 寫入「預排紀錄」，狀態預設為「未核銷」。 | `Core_Service.js` |
| | `核銷` / `結算` | 顯示該講師「未核銷預排」的 Carousel 卡片 → 點擊確認有上課或未上課，更新狀態。 | `Core_Service.js` |
| | `刪除` | 刪除該講師最後一筆授課或預排紀錄（自動比對最新的一筆進行安全刪除）。 | `Core_Service.js` |
| **行政代操作** <br>*(限 Admin)* | `代登 [講師名]` <br> `代排 [講師名]` <br> `代核 [講師名]` | 允許行政管理員代替指定講師進行登錄、預排、核銷操作。若無輸入講師名，會彈出講師選擇選單。 | `Core_Service.js` |
| **財務試算** <br>*(限 Admin)* | `學費試算` | 導引輸入月份 → 計算學費 → 顯示試算 Flex 卡片 → 確認後寫入「學費結算表」。 | `Finance_Service.js` |
| | `鐘點試算` | 導引輸入月份 → 計算鐘點費（自動計算扣繳稅額 10% 與補充保費 2.11%） → 寫入「鐘點結算表」。 | `Finance_Service.js` |
| | `產生繳費單` | 導引輸入月份 → 批次為該月有上課的學生產生 PDF 繳費單。 | `Finance_Service.js` <br> `Doc_Service.js` |
| | `查詢繳費單 YYYY/MM 姓名` | 直接查詢並回傳特定學生在特定月份的繳費單 PDF 連結。 | `Finance_Service.js` |
| **單據系統** | `開收據 YYYY/MM 學生` | 選擇繳費方式 → 選擇類別（學費/其他）→ 輸入日期 → 產生 PDF 預覽 → 確認存檔後自動寫入「會計日記帳」。 | `Doc_Service.js` |
| | `寄收據 YYYY/MM 學生` | 將已存檔的 PDF 收據以 Gmail 附件形式寄送給學生。 | `Doc_Service.js` |
| | `開領據 YYYY/MM 講師` | 批次或單獨產生講師的鐘點費領據 PDF 預覽 → 確認存檔。 | `Doc_Service.js` |
| | `寄領據 YYYY/MM` | 批次將已存檔的 PDF 領據以 Gmail 附件寄送給各位講師。 | `Doc_Service.js` |
| | `開一般收據` | 用於入會費、常年會費、捐款收據生成 → 產生 PDF → 自動寫入會計日記帳。 | `Doc_Service.js` |
| | `寄一般收據` | 將一般收據 PDF 以 Gmail 寄送。 | `Doc_Service.js` |
| **稅務專區** | `記帳` | 導引輸入會計分錄（借貸） → 自動寫入「會計日記帳」。 | `Finance_Service.js` |
| | `損益表 YYYY` <br> `資產負債表 YYYY` <br> `現金流量表 YYYY` | 根據「會計日記帳」的資料，即時運算並渲染該年度的財務三表 Flex 訊息。 | `Finance_Service.js` |
| | `免稅試算 YYYY` | 運算教育文化公益慈善機關免稅標準（80% 法定支出門檻）並回傳預警及圖表。 | `Finance_Service.js` |
| | `年度扣繳 YYYY` | 彙整該年度所有講師的申報資料，並寫入扣繳申報工作表。 | `Finance_Service.js` |
| | `捐款申報 YYYY` | 彙整年度捐款紀錄，比對身分證字號，產出符合國稅局格式的 CSV 申報檔。 | `Finance_Service.js` |
| **系統工具** | `選單` / `menu` / `主選單` | 呼叫主選單 Flex 訊息。 | `UI_Utils.js` |
| | `稅務` | 呼叫稅務專區 Flex 訊息。 | `UI_Utils.js` |
| | `請假` | 請假系統（事假/病假、餘額查詢、重複申請檢查）。 | `Leave_Service.js` |
| | `查ID` / `查群ID` | 查詢目前對話的 User ID 或 Group ID。 | `Main_Controller.js` |
| | `關鍵字` | 讀取試算表「功能說明」A1 儲存格內容（供快速指令參考）。 | `Main_Controller.js` |
| | `發公告` / `群組公告` | 群發訊息給所有講師或指定群組。 | `Main_Controller.js` |

---

## 2. 環境與架構優化建議 (Environment & Architecture)

由於講師與助理每月僅執行一次，極易遺忘指令，且 GAS 專案在安全性與模組化上有其侷限性。以下是我們為您規劃的環境優化方案：

### A. 解決「常忘記指令」的防呆方案 (Anticipate Needs)
1. **建立 LINE Rich Menu (圖文選單)**：
   - **痛點**：講師與秘書每月才用一次，文字指令容易打錯或忘記。
   - **方案**：利用 LINE Official Account Manager 設定常駐的 Rich Menu。
     - **普通講師版**：包含「課程登記（登錄）」、「未核銷預排（核銷）」、「我要請假」、「主選單」。
     - **行政/秘書版**（可透過 LINE API 動態切換 Link Rich Menu）：額外包含「財務試算」、「稅務專區」、「單據系統」。
   - **效果**：完全免除打字記憶負擔，一鍵觸發。
2. **優化 `關鍵字`（功能說明）動態提示卡**：
   - 目前 `關鍵字` 指令只讀取試算表 `功能說明` 的 A1。建議改為讀取整張表的指令清單，並用 Flex Message 渲染成精美的「快捷指令圖卡」，點擊卡片按鈕直接送出指令。

### B. 安全性與配置優化
3. **敏感資訊移出程式碼 (Script Properties)**：
   - **痛點**：`Config.js` 中的 `CHANNEL_ACCESS_TOKEN` 是明碼寫在程式碼裡，若代碼備份到 GitHub 或分享給他人會有安全外洩風險。
   - **方案**：將 Token 放入 Google Apps Script 的 `Script Properties`（專案設定 -> 指令碼屬性），程式碼改用以下方式動態讀取：
     ```javascript
     const CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN');
     ```
4. **試算表權限防護 (Sheet Protection)**：
   - 核心試算表（如「會計日記帳」、「授課紀錄」）中含有大量公式與關聯。建議在 Google Sheets 上設定「保護工作表與範圍」，僅限行政帳號或執行 GAS 的帳號（機器人自己）有寫入權限，防止秘書或講師不小心手動刪改關鍵公式。

### C. 開發與部署流程優化 (CI/CD)
5. **導入 TypeScript + Clasp 結構化管理**：
   - 本專案規模已相當龐大（單是 `Finance_Service.js` 與 `UI_Utils.js` 就各有千行以上）。
   - 建議使用 `clasp` (Chrome V8 Apps Script CLI) 將專案本地化，使用 TypeScript 進行靜態型別檢查，避免 runtime 發生物件屬性拼錯的低級錯誤，並透過 Git 進行完善的版本控制。

---

## 3. 程式碼優化建議 (Code Refactoring)

針對現有的 7 個 `.gs` 檔案，我們發現了以下可以大幅提升維護性、執行效能與穩定性的程式碼優化點：

### 1. 消除大量的重複 HTTP 請求代碼 (DRY Principle)
- **現況**：程式碼中充斥著 `UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", ...)` 與 `UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", ...)` 的重複呼叫（幾乎各個 Service 都有自己封裝的 fetch）。
- **優化**：在 `UI_Utils.js` 中提煉一個高内聚的 `LineClient` 類別或公用函式：
  ```javascript
  const LineClient = {
    post: function(endpoint, payload) {
      const url = "https://api.line.me/v2/bot/" + endpoint;
      const options = {
        "method": "post",
        "headers": {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN
        },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      };
      const response = UrlFetchApp.fetch(url, options);
      // 統一做 Error Handling / Log 記錄
      return response;
    },
    reply: function(replyToken, messages) {
      return this.post("message/reply", { replyToken, messages });
    },
    push: function(to, messages) {
      return this.post("message/push", { to, messages });
    }
  };
  ```

### 2. 修正 L610 `handleAdminPickStudent` 的寫死硬字串 Bug
- **現況**：在 `Main_Controller.js` 的 `handleAdminPickStudent` 函式中：
  ```javascript
  // L610 附近
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("講師名單"); 
  ```
  這裡直接寫死了 `"講師名單"` 字串。
- **優化**：應統一使用 `Config.js` 中定義的 `SHEET_NAME_TEACHER` 常數，避免日後更名時遺漏。

### 3. 清理已註解的死代碼 (Dead Code Cleanup)
- **現況**：
  - `Finance_Service.js` 中有近 130 行（L503-L639）被註解掉的舊版 `handleTaxExemptionDashboardCommand`。
  - `Doc_Service.js` 中有被註解掉的 `executeReceiptSaveOnly` (L404-L432)。
- **優化**：這些死代碼會干擾閱讀，應利用 Git 記錄歷史，並將現有代碼中的大段註解徹底刪除。

### 4. 統一 Cache Service 的生存時間 (TTL) 常數
- **現況**：專案中多處使用 `CacheService.getScriptCache()` 來維護狀態機，但過期時間混雜了 `300` 秒與 `600` 秒。
- **優化**：在 `Config.js` 定義統一的 `CACHE_TTL = 600;`，並封裝 `StateManager` 簡化狀態讀寫，避免各個 Service 直接操作原生 Cache。

### 5. 提煉 `isAdmin` 權限檢查機制
- **現況**：多個需要管理員權限的指令（如學費試算、免稅額試算等）各自在函式內讀取 `ADMIN_LIST` 並判斷是否包含 `userId`。
- **優化**：封裝公用函數 `Auth.isAdmin(userId)`，程式碼更具可讀性：
  ```javascript
  if (!Auth.isAdmin(userId)) {
    return replyTextMessage(replyToken, "⚠️ 您沒有此功能的執行權限。");
  }
  ```

### 6. 優化 Cache Lock 機制，防止併發寫入衝突
- **現況**：稅務申報、捐款申報等高耗能的批次操作雖然有使用 `LOCK_ANNUAL_` 等 Cache 鎖，但 GAS 官方推薦使用 `LockService.getScriptLock()` 來處理執行緒等級的寫入鎖（防止多用戶同時操作導致 Sheet 寫入覆蓋）。
- **優化**：對於會修改 Sheets 的試算、儲存與單據寫入操作，引入 `LockService`：
  ```javascript
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 等待最多10秒
    // 執行試算與寫入 Sheet
  } catch (e) {
    // 回傳系統繁忙提示
  } finally {
    lock.releaseLock();
  }
  ```

---

## 4. 待確認的業務邏輯提問清單 (Questions for User)

為了能夠給出最精確的優化方案，有幾個涉及**業務邏輯**與**歷史遗留代碼**的部分需要與您核對，請您協助解答：

> [!IMPORTANT]
> **請協助確認以下問題，這將直接決定我們後續的程式碼結構與體驗設計：**
> 1. **「功能說明」分頁的維護方式**：
>    目前系統中有 `關鍵字` 指令會讀取 `功能說明` 試算表 A1。該分頁的內容目前是否由秘書定期手動更新？我們是否能將其改造成一個自動生成的 Flex 格式「快速功能導覽卡」，以降低手動維護成本？
> 2. **`MAGIC_KEYWORDS` (允許同時段重疊課程) 的業務邏輯**：
>    程式碼中若課程名稱包含 `MAGIC_KEYWORDS` 就不會觸發「時段重疊檢查」。請問這個機制的實際業務場景是什麼？（例如：線上課、講座或是特定的團體課？）是否需要將這個關鍵字清單改為從試算表設定頁面動態讀取，而不是寫死在 `Config.js` 中？
> 3. **請假系統的權限對象**：
>    在 `Leave_Service.js` 的註解中提到「請假系統限秘書操作」，但程式碼中並沒有嚴格限制 `ADMIN_LIST` 才能使用 `請假` 相關指令。目前是一般講師也可以在 LINE Bot 上自己請假，還是實際上只有秘書在使用？
> 4. **已註解的「批次功能」是否需要恢復**：
>    在 `Doc_Service.js` 中，有些關於批次收據生成或自動寄送的邏輯被註解掉了。請問目前的作業流程是「一筆一筆手動開立/寄送」，還是希望未來能實現「一鍵批次生成並寄送」？
> 5. **`REPORT_SHEET_ID` 的機密性**：
>    這個 Sheet 包含免稅試算、年度扣繳等敏感財務數據，目前除了 GAS 讀取外，該 Sheet 是否有分享給其他行政人員協作？是否需要為其設計獨立的資料加密或更嚴格的存取權限審查？
> 6. **講師資料表中的 formatCode、nationality、nhiExempt 欄位**：
>    這些是計算二代健保（補充保費）與所得稅扣繳的重要欄位。請問這些欄位目前是由行政人員在 Google Sheets 後台手動登錄，還是有預留透過 LINE Bot 讓講師自己綁定/申報的規劃？
