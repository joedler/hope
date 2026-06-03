// ==========================================
// ⚙️ Config.ts : 全域設定 (TypeScript + 安全重構版)
// ==========================================

// 輔助函式：安全獲取 Script 屬性；敏感設定不得寫死在程式碼中
function getRequiredProperty(key: string): string {
  try {
    const prop = PropertiesService.getScriptProperties().getProperty(key);
    if (prop) return prop;
  } catch (e) {
    // Apps Script 尚未初始化 PropertiesService 時，讓下方錯誤訊息接手。
  }
  throw new Error("缺少必要 Script Property: " + key);
}

// LINE Bot 設定
var CHANNEL_TOKEN: string = getRequiredProperty("LINE_CHANNEL_TOKEN");

// Google 試算表 ID (主資料庫、請假、報表、會員名單)
var SHEET_ID: string = getRequiredProperty("SPREADSHEET_ID");
var LEAVE_SHEET_ID: string = getRequiredProperty("LEAVE_SHEET_ID");
var REPORT_SHEET_ID: string = getRequiredProperty("REPORT_SHEET_ID");
var SHEET_ID_MEMBER: string = getRequiredProperty("SHEET_ID_MEMBER");

// ★ 協會專屬稅務設定 ★
var ORG_TAX_ID: string = "91622132"; // 協會統一編號

// ★ 分頁名稱設定 (請確保試算表分頁名稱一致) ★
var SHEET_NAME_COURSE: string = "課程設定表";
var SHEET_NAME_STUDENT: string = "學生基本資料表";  
var SHEET_NAME_TEACHER: string = "講師名單";
var SHEET_NAME_RECORD: string = "授課紀錄";
var SHEET_NAME_PLAN: string = "預排紀錄";
var SHEET_NAME_FIN_FEE: string = "學費結算表";
var SHEET_NAME_FIN_PAY: string = "鐘點結算表";
var SHEET_NAME_TUITION_ADJUSTMENT: string = "學費調整紀錄表";

// 群組 ID (推播用)
var GROUP_ID: string = "C0227dd553381f2503d344481ae1b4453";

// 通關關鍵字
var MAGIC_KEYWORDS: string[] = ["諮詢師", "家長", "同時段", "人際情緒支持兒童團體"];

// 管理員與秘書清單
var ADMIN_LIST: string[] = [
  "U65c06840e57dd0fa7dee49fbcc9ca5c6", // 管理員
  "U0ba285786d5ab40ddcb30e7c394ca384",  // 秘書處
  "U8d812c4a0b5e44b34fcc2c1d86b08e87"   // 小白
];

// 一般收據專用試算表與範本 ID
var TEMPLATE_ID_GEN_RECEIPT: string = "1cIMvNBr_j8que87efpJwNF9wqK03uiz8pXVAG7sQXIY"; // 一般收據範本
var FOLDER_ID_GEN_RECEIPT: string = "17jkNW3fslGa_4nc4MSghFkGsrh3Iwd_E";
var SHEET_NAME_GEN_RECORD: string = "一般收據紀錄"; // 核心試算表中的紀錄分頁

// 一次性設定 Script 屬性的初始化工具，供首次部署新客戶時執行
function setupProjectProperties() {
  const properties = {
    LINE_CHANNEL_TOKEN: CHANNEL_TOKEN,
    SPREADSHEET_ID: SHEET_ID,
    LEAVE_SHEET_ID: LEAVE_SHEET_ID,
    REPORT_SHEET_ID: REPORT_SHEET_ID,
    SHEET_ID_MEMBER: SHEET_ID_MEMBER
  };
  PropertiesService.getScriptProperties().setProperties(properties);
}

function forceAuth() {
  UrlFetchApp.fetch("https://www.google.com");
  DriveApp.getRootFolder();
  SpreadsheetApp.openById(SHEET_ID).getSheets();
  GmailApp.getAliases();
  Logger.log("所有外部資源授權完成");
}

