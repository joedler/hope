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

function getOptionalProperty(key: string, fallback: string): string {
  try {
    const prop = PropertiesService.getScriptProperties().getProperty(key);
    if (prop) return prop;
  } catch (e) {
    // 保留 fallback，避免屬性尚未建立時影響既有正式功能。
  }
  return fallback;
}

function getOptionalListProperty(key: string, fallback: string[]): string[] {
  try {
    const prop = PropertiesService.getScriptProperties().getProperty(key);
    if (prop) {
      return prop
        .split(/[\n,;]+/)
        .map((item) => item.replace(/^\s+|\s+$/g, ""))
        .filter((item) => item !== "");
    }
  } catch (e) {
    // 保留 fallback，避免屬性尚未建立時影響既有正式功能。
  }
  return fallback;
}

// LINE Bot 設定
var CHANNEL_TOKEN: string = getRequiredProperty("LINE_CHANNEL_TOKEN");

// Google 試算表 ID (主資料庫、請假、報表、會員名單)
var SHEET_ID: string = getRequiredProperty("SPREADSHEET_ID");
var LEAVE_SHEET_ID: string = getRequiredProperty("LEAVE_SHEET_ID");
var REPORT_SHEET_ID: string = getRequiredProperty("REPORT_SHEET_ID");
var SHEET_ID_MEMBER: string = getRequiredProperty("SHEET_ID_MEMBER");

// ★ 協會專屬稅務設定 ★
var ORG_TAX_ID: string = getOptionalProperty("ORG_TAX_ID", "91622132"); // 協會統一編號

// ★ 分頁名稱設定 (請確保試算表分頁名稱一致) ★
var SHEET_NAME_COURSE: string = "課程設定表";
var SHEET_NAME_STUDENT: string = "學生基本資料表";  
var SHEET_NAME_TEACHER: string = "講師名單";
var SHEET_NAME_RECORD: string = "授課紀錄";
var SHEET_NAME_PLAN: string = "預排紀錄";
var SHEET_NAME_FIN_FEE: string = "學費結算表";
var SHEET_NAME_FIN_PAY: string = "鐘點結算表";
var SHEET_NAME_TUITION_ADJUSTMENT: string = "學費調整紀錄表";
var SHEET_NAME_DOCUMENT_RECORD: string = "單據紀錄表";

// 群組 ID (推播用)
var GROUP_ID: string = getOptionalProperty("LINE_GROUP_ID", "C0227dd553381f2503d344481ae1b4453");

// 通關關鍵字
var MAGIC_KEYWORDS: string[] = ["諮詢師", "家長", "同時段", "人際情緒支持兒童團體"];

// 管理員與秘書清單
var ADMIN_LIST: string[] = getOptionalListProperty("ADMIN_LINE_USER_IDS", [
  "U65c06840e57dd0fa7dee49fbcc9ca5c6", // 管理員
  "U0ba285786d5ab40ddcb30e7c394ca384",  // 秘書處
  "U8d812c4a0b5e44b34fcc2c1d86b08e87"   // 小白
]);

// 一般收據專用試算表與範本 ID
var TEMPLATE_ID_GEN_RECEIPT: string = getOptionalProperty("TEMPLATE_ID_GENERAL_RECEIPT", "1cIMvNBr_j8que87efpJwNF9wqK03uiz8pXVAG7sQXIY"); // 一般收據範本
var FOLDER_ID_GEN_RECEIPT: string = getOptionalProperty("PDF_FOLDER_GENERAL_RECEIPT", "17jkNW3fslGa_4nc4MSghFkGsrh3Iwd_E");
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
  const templateIds = [
    TEMPLATE_ID_PAYMENT,
    TEMPLATE_ID_RECEIPT,
    TEMPLATE_ID_ALLOWANCE,
    TEMPLATE_ID_GEN_RECEIPT
  ];
  for (let i = 0; i < templateIds.length; i++) {
    const authDocFile = DriveApp.getFileById(templateIds[i]).makeCopy("forceAuth_document_scope_check");
    DocumentApp.openById(authDocFile.getId()).getBody().getText();
    authDocFile.setTrashed(true);
  }
  DriveApp.getFolderById(PDF_FOLDER_CONFIG.PAYMENT_NOTICE).getName();
  DriveApp.getFolderById(PDF_FOLDER_CONFIG.RECEIPT).getName();
  DriveApp.getFolderById(PDF_FOLDER_CONFIG.ALLOWANCE).getName();
  DriveApp.getFolderById(FOLDER_ID_GEN_RECEIPT).getName();
  GmailApp.getAliases();
  Logger.log("所有外部資源授權完成");
}

function auditProjectProperties() {
  const requiredKeys = [
    "LINE_CHANNEL_TOKEN",
    "SPREADSHEET_ID",
    "LEAVE_SHEET_ID",
    "REPORT_SHEET_ID",
    "SHEET_ID_MEMBER"
  ];
  const recommendedKeys = [
    "TEMPLATE_ID_PAYMENT",
    "TEMPLATE_ID_RECEIPT",
    "TEMPLATE_ID_ALLOWANCE",
    "TEMPLATE_ID_GENERAL_RECEIPT",
    "PDF_FOLDER_PAYMENT_NOTICE",
    "PDF_FOLDER_RECEIPT",
    "PDF_FOLDER_ALLOWANCE",
    "PDF_FOLDER_GENERAL_RECEIPT",
    "ORG_TAX_ID",
    "LINE_GROUP_ID",
    "ADMIN_LINE_USER_IDS"
  ];
  const props = PropertiesService.getScriptProperties();
  const missingRequired: string[] = [];
  const missingRecommended: string[] = [];

  requiredKeys.forEach(function(key) {
    if (!props.getProperty(key)) missingRequired.push(key);
  });
  recommendedKeys.forEach(function(key) {
    if (!props.getProperty(key)) missingRecommended.push(key);
  });

  Logger.log("正式版設定檢查：必要屬性缺少 " + missingRequired.length + " 項。");
  if (missingRequired.length > 0) Logger.log("缺少必要屬性：" + missingRequired.join(", "));
  Logger.log("正式版設定檢查：建議屬性缺少 " + missingRecommended.length + " 項。");
  if (missingRecommended.length > 0) Logger.log("缺少建議屬性：" + missingRecommended.join(", "));
  Logger.log("注意：本函式只輸出 key 名稱，不輸出 token、ID 或個資值。");

  return {
    ok: missingRequired.length === 0,
    missingRequired,
    missingRecommended
  };
}

function setupRecommendedProjectProperties() {
  const props = PropertiesService.getScriptProperties();
  const recommendedDefaults: any = {
    TEMPLATE_ID_PAYMENT: TEMPLATE_ID_PAYMENT,
    TEMPLATE_ID_RECEIPT: TEMPLATE_ID_RECEIPT,
    TEMPLATE_ID_ALLOWANCE: TEMPLATE_ID_ALLOWANCE,
    TEMPLATE_ID_GENERAL_RECEIPT: TEMPLATE_ID_GEN_RECEIPT,
    PDF_FOLDER_PAYMENT_NOTICE: PDF_FOLDER_CONFIG.PAYMENT_NOTICE,
    PDF_FOLDER_RECEIPT: PDF_FOLDER_CONFIG.RECEIPT,
    PDF_FOLDER_ALLOWANCE: PDF_FOLDER_CONFIG.ALLOWANCE,
    PDF_FOLDER_GENERAL_RECEIPT: FOLDER_ID_GEN_RECEIPT,
    LINE_GROUP_ID: GROUP_ID,
    ADMIN_LINE_USER_IDS: ADMIN_LIST.join("\n")
  };
  const addedKeys: string[] = [];
  const keptKeys: string[] = [];

  Object.keys(recommendedDefaults).forEach(function(key) {
    if (props.getProperty(key)) {
      keptKeys.push(key);
      return;
    }
    props.setProperty(key, recommendedDefaults[key]);
    addedKeys.push(key);
  });

  Logger.log("已補入建議屬性 " + addedKeys.length + " 項。");
  if (addedKeys.length > 0) Logger.log("補入 key：" + addedKeys.join(", "));
  Logger.log("已存在未覆蓋 " + keptKeys.length + " 項。");
  Logger.log("注意：本函式只輸出 key 名稱，不輸出 token、ID 或個資值。");

  return {
    addedKeys,
    keptKeys
  };
}

function auditFormalSpreadsheetStructure() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const requiredSheets = [
    SHEET_NAME_COURSE,
    SHEET_NAME_STUDENT,
    SHEET_NAME_TEACHER,
    SHEET_NAME_RECORD,
    SHEET_NAME_PLAN,
    SHEET_NAME_FIN_FEE,
    SHEET_NAME_FIN_PAY,
    SHEET_NAME_TUITION_ADJUSTMENT,
    SHEET_NAME_DOCUMENT_RECORD,
    SHEET_NAME_GEN_RECORD,
    "會計日記帳"
  ];
  const minimumColumns: any = {};
  minimumColumns[SHEET_NAME_COURSE] = 7;
  minimumColumns[SHEET_NAME_STUDENT] = 4;
  minimumColumns[SHEET_NAME_TEACHER] = 14;
  minimumColumns[SHEET_NAME_RECORD] = 11;
  minimumColumns[SHEET_NAME_PLAN] = 12;
  minimumColumns[SHEET_NAME_FIN_FEE] = 17;
  minimumColumns[SHEET_NAME_FIN_PAY] = 15;

  const exactHeaders: any = {};
  exactHeaders[SHEET_NAME_FIN_PAY] = [
    "結算月份", "講師姓名", "課程/學生", "課程日期/時間(金額)", "時數/費率",
    "單科試算", "應付小計", "單據編號", "存檔時間", "備註",
    "檔案連結", "Email狀態", "扣繳稅額", "補充保費", "實發金額"
  ];
  exactHeaders[SHEET_NAME_TUITION_ADJUSTMENT] = [
    "建立時間", "調整月份", "學生姓名", "課程名稱", "原上課日期",
    "開始時間", "結束時間", "時數", "單價", "調整金額",
    "調整類型", "關聯原單號", "原因", "狀態", "操作人",
    "備註", "原錯誤月份", "補收單號", "補收PDF", "補收單狀態"
  ];
  exactHeaders[SHEET_NAME_DOCUMENT_RECORD] = [
    "建立時間", "處理月份", "單據類型", "對象類型", "對象姓名",
    "單據編號", "來源表", "來源鍵值", "金額", "PDF連結",
    "產生狀態", "Email狀態", "LINE狀態", "寄送時間", "作廢狀態",
    "備註", "操作人"
  ];
  exactHeaders[SHEET_NAME_GEN_RECORD] = [
    "建立時間", "單據編號", "收據日期", "姓名/單位", "金額", "類別",
    "收款方式", "PDF連結", "Email狀態", "操作人", "身分證號/統一編號", "Email收件人"
  ];

  const missingSheets: string[] = [];
  const errors: string[] = [];
  const warnings: string[] = [];

  requiredSheets.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      missingSheets.push(sheetName);
      return;
    }

    const minCols = minimumColumns[sheetName] || 0;
    if (minCols > 0 && sheet.getMaxColumns() < minCols) {
      errors.push(sheetName + " 欄位數不足：目前 " + sheet.getMaxColumns() + " 欄，至少需要 " + minCols + " 欄。");
    }

    const headers = exactHeaders[sheetName];
    if (headers) {
      const actual = sheet.getRange(1, 1, 1, headers.length).getValues()[0].map(function(value) {
        return String(value || "").trim();
      });
      for (let i = 0; i < headers.length; i++) {
        if (actual[i] !== headers[i]) {
          errors.push(sheetName + " 第 " + (i + 1) + " 欄表頭不一致：應為「" + headers[i] + "」，目前為「" + (actual[i] || "空白") + "」。");
        }
      }
    }
  });

  if (missingSheets.length > 0) {
    missingSheets.forEach(function(sheetName) {
      errors.push("缺少必要分頁：" + sheetName);
    });
  }

  const studentSheet = ss.getSheetByName(SHEET_NAME_STUDENT);
  if (studentSheet && studentSheet.getMaxColumns() >= 4) {
    warnings.push("已確認學生基本資料表至少有 D 欄；家長/學生 LINE User ID 應放 D 欄。");
  }
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (teacherSheet && teacherSheet.getMaxColumns() >= 2) {
    warnings.push("已確認講師名單至少有 B 欄；講師 LINE User ID 應放 B 欄。");
  }

  Logger.log("正式試算表結構檢查：必要錯誤 " + errors.length + " 項。");
  errors.forEach(function(item) { Logger.log("錯誤：" + item); });
  Logger.log("正式試算表結構檢查：提醒 " + warnings.length + " 項。");
  warnings.forEach(function(item) { Logger.log("提醒：" + item); });
  Logger.log("注意：本函式只檢查分頁、欄位數與表頭，不輸出試算表資料內容。");

  return {
    ok: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings
  };
}

