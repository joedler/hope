# 重要 ID 與設定外移清單

最後更新：2026-06-15

本文件用於正式版切換前檢查：哪些設定應放在 GAS Script Properties、哪些應放在 Google Sheets 資料表、哪些可以保留在程式碼中。

## 一、必須放在 GAS Script Properties

這些值屬於環境設定或敏感設定，正式版不可依賴程式碼 fallback。

| Key | 用途 | 備註 |
| --- | --- | --- |
| `LINE_CHANNEL_TOKEN` | LINE Bot Push/Reply API Token | 必填，敏感資料 |
| `SPREADSHEET_ID` | 空空BOT 主資料庫 | 必填 |
| `LEAVE_SHEET_ID` | 請假系統試算表 | 必填 |
| `REPORT_SHEET_ID` | 報表/稅務輸出試算表 | 必填 |
| `SHEET_ID_MEMBER` | 會員/捐款者資料表 | 必填 |
| `TEMPLATE_ID_PAYMENT` | 繳費單 Google Docs 範本 | 必填 |
| `TEMPLATE_ID_RECEIPT` | 收據 Google Docs 範本 | 必填 |
| `TEMPLATE_ID_ALLOWANCE` | 領據 Google Docs 範本 | 必填 |
| `TEMPLATE_ID_GENERAL_RECEIPT` | 一般收據 Google Docs 範本 | 必填 |
| `PDF_FOLDER_PAYMENT_NOTICE` | 繳費單 PDF 存放資料夾 | 必填 |
| `PDF_FOLDER_RECEIPT` | 收據 PDF 存放資料夾 | 必填 |
| `PDF_FOLDER_ALLOWANCE` | 領據 PDF 存放資料夾 | 必填 |
| `PDF_FOLDER_GENERAL_RECEIPT` | 一般收據 PDF 存放資料夾 | 必填 |
| `ORG_TAX_ID` | 協會統一編號 | 稅務與收據使用 |
| `LINE_GROUP_ID` | 行政通知群組 ID | 若正式版不用群組推播，可保留但不使用 |
| `ADMIN_LINE_USER_IDS` | 行政人員 LINE User ID 清單 | 逗號、分號或換行分隔 |

## 二、前端公開設定

這些值會出現在 `docs/config.js`，屬於公開入口設定，不是 token 類機密；但正式切換時仍需核對是否指向正確環境。

| Key | 用途 | 備註 |
| --- | --- | --- |
| `activityName` | 前端系統名稱 | 應為 `空空 工作台` |
| `gasWebAppUrl` | GAS Web App URL | 前端呼叫 LIFF API 使用 |
| `liffId` | LINE LIFF ID | 與 LINE Developers LIFF 設定一致 |
| `liffUrl` | LINE LIFF URL | 圖文選單或入口使用 |
| `pagesUrl` | GitHub Pages URL | LIFF 前端網址 |

## 三、目前程式掃描結果

2026-06-15 盤點結果：

- `LINE_CHANNEL_TOKEN`、`SPREADSHEET_ID`、`LEAVE_SHEET_ID`、`REPORT_SHEET_ID`、`SHEET_ID_MEMBER` 已採必要 Script Property，缺值會報錯。
- `TEMPLATE_ID_PAYMENT`、`TEMPLATE_ID_RECEIPT`、`TEMPLATE_ID_ALLOWANCE`、`TEMPLATE_ID_GENERAL_RECEIPT` 仍有程式 fallback。
- `PDF_FOLDER_PAYMENT_NOTICE`、`PDF_FOLDER_RECEIPT`、`PDF_FOLDER_ALLOWANCE`、`PDF_FOLDER_GENERAL_RECEIPT` 仍有程式 fallback。
- `ADMIN_LINE_USER_IDS` 仍有程式 fallback 行政清單。
- `LINE_GROUP_ID` 仍有程式 fallback 群組 ID。
- `ORG_TAX_ID` 仍有程式 fallback 協會統一編號。
- `docs/config.js` 僅放公開入口設定；不可放 `LINE_CHANNEL_TOKEN`、Channel secret 或個人 LINE User ID。
- 已新增 `auditProjectProperties()`，可在 GAS 編輯器執行，僅列出缺少的 key 名稱，不輸出 token、ID 或個資值。

目前保留 fallback 的理由是避免測試環境屬性缺漏時立即中斷。正式切換前應逐項確認 Script Properties 已建立，並在下一階段將正式版重要 ID fallback 移除或改為明確錯誤。

## 四、應放在 Google Sheets 的資料

這些資料屬於營運資料，需由行政維護，不應寫死在程式碼中。

| 資料 | 建議位置 | 說明 |
| --- | --- | --- |
| 講師姓名、講師 LINE User ID、Email、鐘點費率 | `講師名單` | 講師登入、鐘點、領據、LINE push 使用 |
| 學生姓名、家長 LINE User ID、Email | `學生基本資料表` | 家長繳費單、收據、LINE push 使用；家長 LINE ID 目前放 D 欄 |
| 學生可上課程、收費模式、單價 | `課程設定表` | LIFF 下拉選單、學費試算、鐘點試算使用 |
| 會員/捐款者 Email、身分證號/統一編號 | 會員/捐款者資料表，或 `一般收據紀錄` K/L 欄 | 新增一般收據時以 K/L 欄為優先來源 |
| 行政帳務補救紀錄 | `學費調整紀錄表` | 補收/退費/補發鐘點的來源資料 |
| 單據狀態 | `單據紀錄表` | PDF、Email、LINE、作廢、補發、重產狀態的主要來源 |

## 五、可以保留在程式碼中的設定

以下屬於程式規格或固定命名，可保留在程式碼中，但若正式試算表改名，需同步修改。

- Google Sheets 分頁名稱，例如 `授課紀錄`、`預排紀錄`、`學費結算表`、`鐘點結算表`。
- 單據種類文字，例如 `繳費單`、`收據`、`領據`、`一般收據`。
- 狀態文字，例如 `待寄送`、`已寄送`、`未推播`、`已推播`、`作廢`。
- UI 顯示文字、流程名稱、按鈕名稱。
- 時間欄位顯示規格：使用者可見欄位一律 `HH:MM` 24 小時制，分鐘 15 分鐘間隔。

## 六、正式切換前需要人工確認

1. GAS Script Properties 已包含第一節所有必填 key。
2. `docs/config.js` 的 `gasWebAppUrl`、`liffId`、`liffUrl`、`pagesUrl` 指向正式入口。
3. 在 GAS 編輯器執行 `auditProjectProperties()`，確認必要屬性缺少 0 項。
4. Google Docs 範本權限允許 GAS 專案使用，`forceAuth()` 可完成四份範本授權。
5. PDF 資料夾權限允許 GAS 建立檔案，`forceAuth()` 可讀取四個資料夾。
6. `學生基本資料表` 已填入需要 LINE push 的家長 LINE User ID 與 Email。
7. `講師名單` 已填入講師 LINE User ID 與 Email。
8. 行政人員 LINE User ID 已放入 `ADMIN_LINE_USER_IDS` 或等效正式權限來源。
9. 正式版若移除程式 fallback，需先在測試 GAS 執行一次完整流程驗收。
