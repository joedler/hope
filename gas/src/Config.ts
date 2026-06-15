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
  const headerAliases: any = {};
  headerAliases[SHEET_NAME_GEN_RECORD] = {
    "建立時間": ["開立時間"],
    "單據編號": ["收據編號"],
    "收據日期": ["收款年月"],
    "姓名/單位": ["姓名"],
    "類別": ["收入類別"],
    "收款方式": ["繳費方式"],
    "PDF連結": ["PDF檔案連結"],
    "Email狀態": ["Email寄送狀態"],
    "身分證號/統一編號": ["身分證字號/統一編號"]
  };

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
        const expected = headers[i];
        const aliases = (headerAliases[sheetName] && headerAliases[sheetName][expected]) || [];
        if (actual[i] === expected) {
          continue;
        }
        if (aliases.indexOf(actual[i]) > -1) {
          warnings.push(sheetName + " 第 " + (i + 1) + " 欄使用舊表頭「" + actual[i] + "」，位置與語意可相容；建議日後改名為「" + expected + "」。");
        } else {
          errors.push(sheetName + " 第 " + (i + 1) + " 欄表頭不一致：應為「" + expected + "」，目前為「" + (actual[i] || "空白") + "」。");
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

function auditFormalDataSources() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const errors: string[] = [];
  const warnings: string[] = [];

  const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  const studentSheet = ss.getSheetByName(SHEET_NAME_STUDENT);
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);

  if (!courseSheet) errors.push("缺少課程設定表，無法檢查課程資料來源。");
  if (!studentSheet) errors.push("缺少學生基本資料表，無法檢查學生 Email 與 LINE ID。");
  if (!teacherSheet) errors.push("缺少講師名單，無法檢查講師 Email、LINE ID 與行政權限。");

  const validEmailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const knownStudents: any = {};
  const knownTeachers: any = {};

  if (studentSheet) {
    const sourceStudentData = studentSheet.getDataRange().getValues();
    for (let i = 1; i < sourceStudentData.length; i++) {
      const studentName = String(sourceStudentData[i][0] || "").trim();
      if (studentName) knownStudents[studentName] = true;
    }
  }
  if (teacherSheet) {
    const sourceTeacherData = teacherSheet.getDataRange().getValues();
    for (let i = 1; i < sourceTeacherData.length; i++) {
      const teacherName = String(sourceTeacherData[i][0] || "").trim();
      if (teacherName) knownTeachers[teacherName] = true;
    }
  }

  let validCourseCount = 0;
  let duplicateCourseCount = 0;
  if (courseSheet) {
    const courseData = courseSheet.getDataRange().getValues();
    const seenCourseKeys: any = {};
    for (let i = 1; i < courseData.length; i++) {
      const rowNo = i + 1;
      const teacherName = String(courseData[i][1] || "").trim();
      const studentName = String(courseData[i][2] || "").trim();
      const courseName = String(courseData[i][3] || "").trim();
      const fee = parseFloat(courseData[i][4]);
      const ratio = parseFloat(courseData[i][5]);
      const mode = String(courseData[i][6] || "").trim();
      const hasAnyValue = teacherName || studentName || courseName || String(courseData[i][4] || "").trim() || mode;
      if (!hasAnyValue) continue;

      const rowErrors: string[] = [];
      if (!teacherName) rowErrors.push("講師空白");
      if (!studentName) rowErrors.push("學生空白");
      if (!courseName) rowErrors.push("課程空白");
      if (!fee || fee <= 0) rowErrors.push("單價需大於 0");
      if (String(courseData[i][5] || "").trim() !== "" && (isNaN(ratio) || ratio < 0 || ratio > 1)) rowErrors.push("鐘點比例需介於 0 到 1");
      if (mode !== "預收" && mode !== "後收") rowErrors.push("收費模式需為預收或後收");

      if (rowErrors.length > 0) {
        errors.push("課程設定表第 " + rowNo + " 列：" + rowErrors.join("、") + "。");
      } else {
        validCourseCount++;
        if (teacherName && !knownTeachers[teacherName]) warnings.push("課程設定表第 " + rowNo + " 列講師未出現在講師名單，鐘點、領據或綁定可能受影響。");
        if (studentName && !knownStudents[studentName]) warnings.push("課程設定表第 " + rowNo + " 列學生未出現在學生基本資料表，Email 或 LINE push 可能受影響。");
        if (String(courseData[i][5] || "").trim() === "") warnings.push("課程設定表第 " + rowNo + " 列鐘點比例空白，鐘點試算可能為 0。");
        const key = teacherName + "|" + studentName + "|" + courseName;
        if (seenCourseKeys[key]) {
          duplicateCourseCount++;
          warnings.push("課程設定表第 " + rowNo + " 列與第 " + seenCourseKeys[key] + " 列講師/學生/課程相同，請確認是否為刻意重複。");
        } else {
          seenCourseKeys[key] = rowNo;
        }
      }
    }
    if (validCourseCount === 0) errors.push("課程設定表沒有可用課程列。");
  }

  let studentCount = 0;
  let studentEmailCount = 0;
  let studentLineCount = 0;
  let invalidStudentEmailCount = 0;
  if (studentSheet) {
    const studentData = studentSheet.getDataRange().getValues();
    for (let i = 1; i < studentData.length; i++) {
      const rowNo = i + 1;
      const studentName = String(studentData[i][0] || "").trim();
      if (!studentName) continue;
      studentCount++;
      const email = String(studentData[i][1] || "").trim();
      const lineId = String(studentData[i][3] || "").trim();
      if (email && validEmailPattern.test(email)) studentEmailCount++;
      if (email && !validEmailPattern.test(email)) {
        invalidStudentEmailCount++;
        warnings.push("學生基本資料表第 " + rowNo + " 列 Email 格式可能錯誤。");
      }
      if (lineId) studentLineCount++;
    }
    if (studentCount === 0) warnings.push("學生基本資料表目前沒有學生姓名資料。");
    if (studentCount > 0 && studentEmailCount === 0) warnings.push("學生基本資料表沒有有效 Email；繳費單與收據 Email 將無法寄送。");
    if (studentCount > 0 && studentLineCount === 0) warnings.push("學生基本資料表 D 欄沒有 LINE User ID；繳費單與收據 LINE push 將無法使用。");
  }

  let teacherCount = 0;
  let teacherEmailCount = 0;
  let teacherLineCount = 0;
  let adminLineMatchedCount = 0;
  let invalidTeacherEmailCount = 0;
  if (teacherSheet) {
    const teacherData = teacherSheet.getDataRange().getValues();
    for (let i = 1; i < teacherData.length; i++) {
      const rowNo = i + 1;
      const teacherName = String(teacherData[i][0] || "").trim();
      if (!teacherName) continue;
      teacherCount++;
      const lineId = String(teacherData[i][1] || "").trim();
      const email = String(teacherData[i][2] || "").trim();
      if (lineId) teacherLineCount++;
      if (lineId && ADMIN_LIST.indexOf(lineId) > -1) adminLineMatchedCount++;
      if (email && validEmailPattern.test(email)) teacherEmailCount++;
      if (email && !validEmailPattern.test(email)) {
        invalidTeacherEmailCount++;
        warnings.push("講師名單第 " + rowNo + " 列 Email 格式可能錯誤。");
      }
    }
    if (teacherCount === 0) errors.push("講師名單沒有講師姓名資料。");
    if (teacherCount > 0 && teacherLineCount === 0) warnings.push("講師名單 B 欄沒有 LINE User ID；講師綁定、領據 LINE push 與行政權限會受影響。");
    if (teacherCount > 0 && teacherEmailCount === 0) warnings.push("講師名單沒有有效 Email；領據 Email 將無法寄送。");
    if (adminLineMatchedCount === 0) warnings.push("講師名單沒有任何 B 欄 LINE User ID 命中 ADMIN_LINE_USER_IDS；請確認行政人員是否已綁定。");
  }

  Logger.log("正式資料來源檢查：必要錯誤 " + errors.length + " 項。");
  errors.forEach(function(item) { Logger.log("錯誤：" + item); });
  Logger.log("正式資料來源檢查：提醒 " + warnings.length + " 項。");
  warnings.forEach(function(item) { Logger.log("提醒：" + item); });
  Logger.log("統計：可用課程 " + validCourseCount + " 筆，重複課程提醒 " + duplicateCourseCount + " 筆。");
  Logger.log("統計：學生 " + studentCount + " 位，有效 Email " + studentEmailCount + " 位，LINE ID " + studentLineCount + " 位，Email 格式提醒 " + invalidStudentEmailCount + " 筆。");
  Logger.log("統計：講師 " + teacherCount + " 位，有效 Email " + teacherEmailCount + " 位，LINE ID " + teacherLineCount + " 位，行政 LINE 命中 " + adminLineMatchedCount + " 位，Email 格式提醒 " + invalidTeacherEmailCount + " 筆。");
  Logger.log("注意：本函式只輸出列號與統計，不輸出姓名、Email、LINE ID 或其他個資值。");

  return {
    ok: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    stats: {
      validCourseCount,
      studentCount,
      studentEmailCount,
      studentLineCount,
      teacherCount,
      teacherEmailCount,
      teacherLineCount,
      adminLineMatchedCount
    },
    errors,
    warnings
  };
}

