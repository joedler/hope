// ==========================================
// 📄 Doc_Service.ts : 單據生成與收據系統 (TypeScript + DRY 重構版)
// ==========================================

const TEMPLATE_ID_PAYMENT: string = "1oOSkm4aJ980AVqz7sV7zWf6-9YJMRYzApJmNrd5_d3A"; // 繳費單範本
const TEMPLATE_ID_RECEIPT: string = "1DXIWvcl3NvXicSE2-aXXj2dooNbwYlazLuVT4cBcjw8"; // 收據範本
const TEMPLATE_ID_ALLOWANCE: string = "156NO2In8nKKp6Xj52rYBkAw7ZlNOdqYqojc-zVfEP_w"; // 領據範本

const PDF_FOLDER_CONFIG: any = {
  "PAYMENT_NOTICE": "1N_8vwCGkCcCYbxqldQM6FPjw-PFntCpJ", 
  "RECEIPT":        "1TZCwAGoLE0umWMg7TU5qlKta86iFBiz_", 
  "ALLOWANCE":      "16uLXMoEExNvtGPLLxFKMV9Ln0-AUn4UB"  
};

const EMAIL_CONFIG: any = {
  "STUDENT": {
    "SUBJECT": "【收據】{{Month}} 費用收據 - {{Name}}",
    "BODY": "親愛的 {{Name}} 家長您好：\n\n感謝您的繳費，附件為本期收據，請查收。\n\n臺灣撐出空間教育關懷協會 敬上"
  },
  "PAYMENT_NOTICE": {
    "SUBJECT": "【繳費通知】{{Month}} 費用繳費單 - {{Name}}",
    "BODY": "親愛的 {{Name}} 家長您好：\n\n附件為本期繳費通知單，請查收。完成繳費後請回覆告知，以利匯款確認。\n\n臺灣撐出空間教育關懷協會 敬上"
  },
  "TEACHER": {
    "SUBJECT": "【領據】{{Month}} 講師鐘點費 - {{Name}}",
    "BODY": "親愛的 {{Name}} 講師您好：\n\n附件為本月份鐘點費領據，請查收確認。簽名後回傳。\n\n臺灣撐出空間教育關懷協會 敬上"
  }
};

const PDF_MIME_TYPE = "application/pdf";

function sanitizeDocFilePart(value: any): string {
  return String(value || "")
    .replace(/[\\\/:*?"<>|#%\{\}\[\]~&]/g, "_")
    .replace(/\s+/g, "")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function buildDocPdfFileName(docId: string, docType: string, name: string, suffix?: string): string {
  const parts = [docId, docType, name, suffix]
    .filter((part) => part !== undefined && part !== null && String(part) !== "")
    .map((part) => sanitizeDocFilePart(part));
  return parts.join("_") + ".pdf";
}

function findExistingFileByNames(folder: any, fileNames: string[]) {
  for (let i = 0; i < fileNames.length; i++) {
    const files = folder.getFilesByName(fileNames[i]);
    while (files.hasNext()) {
      const file = files.next();
      if (file.isTrashed && file.isTrashed()) continue;
      return file;
    }
  }
  return null;
}

function debugTestSendEmail() {
  const email = Session.getActiveUser().getEmail();
  try {
    GmailApp.sendEmail(email, "機器人權限測試", "您的機器人 Email 寄送功能正常！");
    Logger.log("✅ 成功！郵件已寄送給：" + email);
  } catch(e) { Logger.log("❌ 失敗：" + e.toString()); }
}

// ==========================================
// 🗑️ 共用模組：取消即刪檔 (垃圾桶機制)
// ==========================================
function executeDocCancel(event: any, cachePrefix: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = cachePrefix + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);

  if (cached) {
    const state = JSON.parse(cached);
    if (state.pdfId) {
      try {
        const fileToTrash = DriveApp.getFileById(state.pdfId);
        fileToTrash.setTrashed(true);
      } catch (e) {
        Logger.log("刪除暫存檔失敗: " + e.toString());
      }
    }
    CacheService.getScriptCache().remove(cacheKey);
  }
  
  replyLineMessage(replyToken, "❌ 已取消作業，並自動銷毀雲端的預覽檔案。");
}

// ==========================================
// 🧾 功能模組 A：收據系統 (學生)
// ==========================================

function handleBatchReceiptCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const targetId = event.source.groupId || userId;
  const parts = userMsg.split(" ");
  if (parts.length < 2) { replyLineMessage(replyToken, "❌ 格式：批次收據 YYYY/MM"); return; }
  const targetMonth = parts[1];
  replyLineMessage(replyToken, "⏳ 正在為 " + targetMonth + " 產生收據 PDF...\n(完成後會推播通知)");
  try {
    const result = processBatchReceiptGeneration(targetMonth);
    pushLineMessage(targetId, result);
  } catch (e) { pushLineMessage(targetId, "❌ 產生失敗：" + e.toString()); }
}

function processBatchReceiptGeneration(targetMonth: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("學費結算表");
  const baseSheet = ss.getSheetByName("學生基本資料表") || ss.getSheetByName("課程設定表");
  if (!sheet || !baseSheet) return "❌ 找不到必要分頁";
  const data = sheet.getDataRange().getValues();
  const baseData = baseSheet.getDataRange().getValues();
  
  const emailMap: any = {};
  for (let k = 1; k < baseData.length; k++) emailMap[baseData[k][0]] = { email: baseData[k][1], pid: baseData[k][2] };

  let count = 0;
  const successList: string[] = [];
  const folder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.RECEIPT);

  for (let i = 1; i < data.length; i++) {
    let rowMonth = data[i][0];
    if (rowMonth instanceof Date) rowMonth = Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM");
    
    const method = data[i][12]; const category = data[i][13]; const dateRaw = data[i][14]; const link = data[i][15];
    if (rowMonth == targetMonth && method && category && dateRaw && (!link || link === "")) {
       const sName = data[i][1]; const amount = data[i][8]; const docId = data[i][9]; const pid = data[i][11];
       const dateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), "yyyy/MM/dd") : String(dateRaw);
       if (amount && docId) {
           const info = emailMap[sName] || {};
           const state = {
             name: sName, amount: amount, docId: docId, month: targetMonth, method: method, 
             category: category, date: dateStr, email: info.email, pid: pid || info.pid, detail: data[i][4]
           };
           try {
             const result = generateReceiptPDF(state, folder);
             sheet.getRange(i+1, 16).setValue(result.url);
             sheet.getRange(i+1, 17).setValue("待寄送");   
             successList.push(sName); count++;
           } catch (err) { Logger.log("Gen Skip: " + sName + " Error: " + err); }
       }
    }
  }
  if (count === 0) return "⚠️ " + targetMonth + " 無需產生的資料。";
  return "✅ 已產生 " + count + " 份收據 PDF！狀態更新為「待寄送」。";
}

function handleBatchSendEmailCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const targetId = event.source.groupId || userId;
  const parts = userMsg.split(" ");
  if (parts.length < 2) { replyLineMessage(replyToken, "❌ 格式：批次寄送 YYYY/MM"); return; }
  const targetMonth = parts[1];
  replyLineMessage(replyToken, "⏳ 正在寄送 " + targetMonth + " 的收據 Email...");
  try {
    const result = processBatchEmailSend(targetMonth);
    pushLineMessage(targetId, result);
  } catch (e) { pushLineMessage(targetId, "❌ 寄送執行錯誤：" + e.toString()); }
}

function processBatchEmailSend(targetMonth: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("學費結算表");
  const baseSheet = ss.getSheetByName("學生基本資料表") || ss.getSheetByName("課程設定表");
  if (!sheet || !baseSheet) return "❌ 資料表缺失";
  const data = sheet.getDataRange().getValues();
  const baseData = baseSheet.getDataRange().getValues();
  const emailMap: any = {};
  for (let k = 1; k < baseData.length; k++) emailMap[baseData[k][0]] = baseData[k][1];

  let count = 0; const failList: string[] = []; const successList: string[] = [];
  for (let i = 1; i < data.length; i++) {
    let rowMonth = data[i][0];
    if (rowMonth instanceof Date) rowMonth = Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM");
    
    const status = data[i][16]; const pdfUrl = data[i][15]; const sName = data[i][1];
    if (rowMonth == targetMonth && status === "待寄送" && pdfUrl && pdfUrl !== "") {
       const email = emailMap[sName]; let newStatus = "寄送失敗";
       if (email && email.indexOf("@") > -1) {
         try {
           const fileIdMatch = pdfUrl.match(/[-\w]{25,}/);
           if (fileIdMatch) {
             const file = DriveApp.getFileById(fileIdMatch[0]);
             const subject = EMAIL_CONFIG.STUDENT.SUBJECT.replace("{{Month}}", targetMonth).replace("{{Name}}", sName);
             const body = EMAIL_CONFIG.STUDENT.BODY.replace("{{Name}}", sName);
             GmailApp.sendEmail(email, subject, body, { attachments: [file.getAs(PDF_MIME_TYPE)] });
             newStatus = "已寄送"; successList.push(sName); count++;
           } else { newStatus = "失敗(連結錯)"; failList.push(sName + "(連結錯)"); }
         } catch (err) { newStatus = "失敗(" + err.message + ")"; failList.push(sName + "(權限/配額)"); }
       } else { newStatus = "失敗(無Email)"; failList.push(sName + "(無Email)"); }
       sheet.getRange(i+1, 17).setValue(newStatus);
    }
  }
  let msg = "✅ 寄送作業結束。\n成功：" + count + " 封\n";
  if (failList.length > 0) msg += "❌ 失敗/跳過：" + failList.length + " 封\n(" + failList.join(", ") + ")";
  if (count === 0 && failList.length === 0) msg += "⚠️ 沒有發現狀態為「待寄送」的項目。";
  return msg;
}

function handleReceiptCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const parts = userMsg.split(" ");
  if (parts.length < 3) { replyLineMessage(replyToken, "❌ 格式：開收據 YYYY/MM 學生姓名"); return; }
  const targetMonth = parts[1]; const targetStudent = parts[2];
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("學費結算表");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  let rowData = null; let rowIndex = -1; const detailsArr: string[] = [];
  
  for (let i = 1; i < data.length; i++) {
    let rowMonth = data[i][0];
    if (rowMonth instanceof Date) rowMonth = Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM");
    if (rowMonth == targetMonth && data[i][1] == targetStudent) {
        if (data[i][4]) detailsArr.push(data[i][4]);
        if (data[i][8] && data[i][8] !== "" && data[i][9] && data[i][9] !== "") { 
            rowData = data[i]; rowIndex = i + 1; 
        }
    }
  }

  if (!rowData) { replyLineMessage(replyToken, "❌ 找不到 " + targetStudent + " 的主資料。"); return; }
  const currentStatus = rowData[16];
  if (currentStatus && (currentStatus.indexOf("已寄送") > -1)) {
     replyLineMessage(replyToken, "⚠️ 該筆收據已寄出過！若需重開，請先手動清空 Q 欄。"); return;
  }
  
  const state = { step: "WAIT_METHOD", rowIndex: rowIndex, month: targetMonth, name: targetStudent, amount: rowData[8], docId: rowData[9], detail: detailsArr.join("\n\n") };
  CacheService.getScriptCache().put("RCPT_" + userId, JSON.stringify(state), 600);

  const buttons = [
    { "type": "postback", "label": "現金", "data": "action=rcpt_method&val=現金" },
    { "type": "postback", "label": "匯款", "data": "action=rcpt_method&val=匯款" }
  ];
  replyButtons(replyToken, "💰 金額：$" + rowData[8] + "\n請選擇繳費方式：", buttons);
}

function handleReceiptMethod(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = "RCPT_" + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (!cached) { replyLineMessage(replyToken, "⚠️ 介面逾時。"); return; }
  const state = JSON.parse(cached);
  state.method = postbackData.split("val=")[1]; state.step = "WAIT_CATEGORY";
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(state), 600);

  const buttons = [
    { "type": "postback", "label": "課程", "data": "action=rcpt_cat&val=課程" },
    { "type": "postback", "label": "講座", "data": "action=rcpt_cat&val=講座" }
  ];
  replyButtons(replyToken, "已選擇：" + state.method + "\n請選擇類別：", buttons);
}

function handleReceiptCategory(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = "RCPT_" + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (!cached) { replyLineMessage(replyToken, "⚠️ 介面逾時。"); return; }
  
  const state = JSON.parse(cached);
  state.category = postbackData.split("val=")[1]; state.step = "WAIT_RECEIPT_DATE";
  CacheService.getScriptCache().put(userId, JSON.stringify({ status: "WAITING_RECEIPT_DATE", rcptState: state }), 600);
  CacheService.getScriptCache().remove(cacheKey); 

  replyLineMessage(replyToken, "📅 請輸入「實際收款日期」\n格式：YYYY/MM/DD (或輸入「今天」)");
}

function handleReceiptDateInput(event: any, userMsg: string, globalState: any) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const state = globalState.rcptState;
  const timeZone = Session.getScriptTimeZone();
  const dateStr = userMsg.trim(); let finalDate = "";
  
  try {
    if (dateStr === "今天") finalDate = Utilities.formatDate(new Date(), timeZone, "yyyy/MM/dd");
    else {
      const parts = dateStr.split(/[\/\-\.]/); const now = new Date(); const currentYear = now.getFullYear(); let inputDate;
      if (parts.length >= 3) { let y = parseInt(parts[0]); if (y < 1911) y += 1911; inputDate = new Date(y, parseInt(parts[1])-1, parseInt(parts[2])); } 
      else if (parts.length === 2) {
        const m = parseInt(parts[0]); const d = parseInt(parts[1]); const tempDate = new Date(currentYear, m - 1, d);
        const diffTime = tempDate.getTime() - now.getTime(); const diffDays = diffTime / (1000 * 3600 * 24);
        if (diffDays > 180) inputDate = new Date(currentYear - 1, m - 1, d);
        else if (diffDays < -180) inputDate = new Date(currentYear + 1, m - 1, d); else inputDate = tempDate;
      } else { throw "格式不符"; }
      if (isNaN(inputDate.getTime())) throw "日期無效";
      finalDate = Utilities.formatDate(inputDate, timeZone, "yyyy/MM/dd");
    }
  } catch (e) { replyLineMessage(replyToken, "❌ 日期辨識失敗。"); return; }
  
  state.date = finalDate; state.step = "FINAL_CONFIRM";
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const baseSheet = ss.getSheetByName("學生基本資料表") || ss.getSheetByName("課程設定表");
  if (!baseSheet) return;
  const baseData = baseSheet.getDataRange().getValues();
  for (let i = 1; i < baseData.length; i++) {
    if (baseData[i][0] === state.name) { state.email = baseData[i][1]; state.pid = baseData[i][2]; break; }
  }

  try {
    const folder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.RECEIPT);
    const pdfResult = generateReceiptPDF(state, folder);
    state.pdfUrl = pdfResult.url; 
    state.pdfId = pdfResult.id; 
    
    CacheService.getScriptCache().put("RCPT_" + userId, JSON.stringify(state), 600);
    CacheService.getScriptCache().remove(userId);
    replyReceiptPreview(replyToken, state);
  } catch(e) { replyLineMessage(replyToken, "❌ 預覽產生失敗：" + e.toString()); }
}

// ==========================================
// 🧾 收入端：確認寫入並自動切傳票 (學費收據)
// ==========================================
function executeReceiptSaveOnly(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = "RCPT_" + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);
  
  if (!cached) { replyLineMessage(replyToken, "⚠️ 逾時，請重新操作。"); return; }
  const state = JSON.parse(cached);
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("學費結算表");
    const journalSheet = ss.getSheetByName("會計日記帳");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let confirmRow = -1;

    for(let i=1; i<data.length; i++) {
        if(data[i][9] == state.docId) { confirmRow = i+1; break; }
    }
    
    if (confirmRow > 0) {
        sheet.getRange(confirmRow, 12).setValue(state.pid || "");
        sheet.getRange(confirmRow, 13).setValue(state.method);
        sheet.getRange(confirmRow, 14).setValue(state.category);
        sheet.getRange(confirmRow, 15).setValue(state.date);
        sheet.getRange(confirmRow, 16).setValue(state.pdfUrl);
        sheet.getRange(confirmRow, 17).setValue("待寄送");

        if (journalSheet) {
          const now = new Date();
          const summary = "收到學費 - " + state.name + " (" + state.month + ")";
          const cashCode = (state.method === "現金") ? "1101" : "1102";
          const cashName = (state.method === "現金") ? "現金" : "銀行存款";
          
          const journalRows = [
            [now, state.docId, cashCode, cashName, summary, state.amount, "", state.docId, "系統自動"],
            [now, state.docId, "4101", "學費收入", summary, "", state.amount, state.docId, "系統自動"]
          ];
          journalSheet.getRange(journalSheet.getLastRow() + 1, 1, 2, 9).setValues(journalRows);
        }
    } else {
        replyLineMessage(replyToken, "⚠️ 寫入失敗：找不到原始單據。"); return;
    }
  } catch (e) {
    replyLineMessage(replyToken, "⚠️ 寫入失敗：" + e.toString());
    return;
  }

  CacheService.getScriptCache().remove(cacheKey);
  replyLineMessage(replyToken, "✅ 學費收據已寫入，並自動完成會計分錄！\n📂 檔案：" + state.pdfUrl);
}

// ==========================================
// 🧾 收入端：確認寫入並自動切傳票 (一般收據)
// ==========================================
function executeGenReceiptSave(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = "GEN_RCPT_" + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);
  
  if (!cached) return; 
  const state = JSON.parse(cached);
  CacheService.getScriptCache().remove(cacheKey); 

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME_GEN_RECORD);
    const journalSheet = ss.getSheetByName("會計日記帳");
    if (!sheet) return;
    
    let operatorName = "ID:" + userId;
    const tSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
    if(tSheet) {
      const tData = tSheet.getDataRange().getValues();
      for(let i=1; i<tData.length; i++) {
        if(tData[i][1] === userId) { operatorName = tData[i][0]; break; }
      }
    }

    sheet.appendRow([
      new Date(), state.docId, state.date, state.name, state.amount, state.category, state.method, state.pdfUrl, "待寄送", operatorName
    ]);

    if (journalSheet) {
      const now = new Date();
      const summary = state.category + " - " + state.name;
      
      const drCode = (state.method === "現金") ? "1101" : "1102";
      const drName = (state.method === "現金") ? "現金" : "銀行存款";
      
      const crCode = (state.category === "捐款") ? "4102" : "4103";
      const crName = (state.category === "捐款") ? "捐款收入" : "入會費及常年會費";
      
      const journalRows = [
        [now, state.docId, drCode, drName, summary, state.amount, "", state.docId, operatorName],
        [now, state.docId, crCode, crName, summary, "", state.amount, state.docId, operatorName]
      ];
      journalSheet.getRange(journalSheet.getLastRow() + 1, 1, 2, 9).setValues(journalRows);
    }

    replyLineMessage(replyToken, "✅ 一般收據存檔成功，會計分錄已自動拋轉！\n編號：" + state.docId);
  } catch (e) {
    replyLineMessage(replyToken, "⚠️ 存檔失敗：" + e.toString());
  }
}

function handleSingleSendReceiptCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const parts = userMsg.split(" ");
  if (parts.length < 3) { replyLineMessage(replyToken, "❌ 格式：寄收據 YYYY/MM 姓名"); return; }
  const targetMonth = parts[1]; const targetStudent = parts[2];
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("學費結算表");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  let targetRow = -1; let pdfUrl = ""; 
  for (let i = 1; i < data.length; i++) {
     let rowMonth = data[i][0];
     if (rowMonth instanceof Date) rowMonth = Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM");
     if (rowMonth == targetMonth && data[i][1] == targetStudent && data[i][9]) {
         targetRow = i + 1; pdfUrl = data[i][15]; break;
     }
  }
  
  if (targetRow == -1) { replyLineMessage(replyToken, "❌ 找不到紀錄。"); return; }
  if (!pdfUrl || pdfUrl === "") { replyLineMessage(replyToken, "⚠️ 該筆資料尚未產生 PDF (P欄為空)。"); return; }
  
  replyLineMessage(replyToken, "⏳ 處理中...");
  const baseSheet = ss.getSheetByName("學生基本資料表") || ss.getSheetByName("課程設定表");
  if (!baseSheet) return;
  const baseData = baseSheet.getDataRange().getValues();
  let email = "";
  for (let k = 1; k < baseData.length; k++) { if (baseData[k][0] === targetStudent) { email = baseData[k][1]; break; } }
  
  let newStatus = "寄送失敗";
  if (email && email.indexOf("@") > -1) {
     try {
       const fileIdMatch = pdfUrl.match(/[-\w]{25,}/);
       if (fileIdMatch) {
          const file = DriveApp.getFileById(fileIdMatch[0]);
          const subject = EMAIL_CONFIG.STUDENT.SUBJECT.replace("{{Month}}", targetMonth).replace("{{Name}}", targetStudent);
          const body = EMAIL_CONFIG.STUDENT.BODY.replace("{{Name}}", targetStudent);
          GmailApp.sendEmail(email, subject, body, { attachments: [file.getAs(PDF_MIME_TYPE)] });
          newStatus = "已寄送";
       }
     } catch(e: any) { newStatus = "失敗:" + e.message; }
  } else { newStatus = "無Email"; }
  
  sheet.getRange(targetRow, 17).setValue(newStatus);
  const targetId = event.source.groupId || userId;
  pushLineMessage(targetId, "📧 寄送結果：" + newStatus);
}

function handlePaymentQueryCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const parts = userMsg.split(" ");
  if (parts.length < 3) { replyLineMessage(replyToken, "❌ 格式：查詢繳費單 YYYY/MM 姓名"); return; }
  const targetMonth = parts[1]; const targetStudent = parts[2];
  const ss = SpreadsheetApp.openById(SHEET_ID); const sheet = ss.getSheetByName("學費結算表");
  if (!sheet) { replyLineMessage(replyToken, "❌ 找不到結算表"); return; }
  const data = sheet.getDataRange().getValues(); let docId = "";
  for (let i = data.length - 1; i >= 1; i--) {
    let rowMonth = data[i][0];
    if (rowMonth instanceof Date) rowMonth = Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM");
    if (rowMonth == targetMonth && data[i][1] == targetStudent) {
      docId = data[i][9]; if (docId && docId !== "") break;
    }
  }
  if (docId === "") { replyLineMessage(replyToken, "❌ 找不到紀錄。"); return; }
  const fileName = buildDocPdfFileName(docId, "繳費單", targetStudent);
  const legacyFileName = "繳費單_" + targetStudent + "_" + docId + ".pdf";
  try {
    const folder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.PAYMENT_NOTICE);
    const file = findExistingFileByNames(folder, [fileName, legacyFileName]);
    if (file) { replyLineMessage(replyToken, "📂 繳費單:\n" + file.getUrl()); } 
    else { replyLineMessage(replyToken, "⚠️ 找不到 PDF。"); }
  } catch (e) { replyLineMessage(replyToken, "❌ 查詢錯誤：" + e.toString()); }
}

function handleTuitionAdjustmentCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const isAdmin = ADMIN_LIST.indexOf(userId) > -1;
  if (!isAdmin) { replyLineMessage(replyToken, "❌ 權限不足：限行政人員使用。"); return; }

  const parts = userMsg.trim().split(/\s+/);
  if (parts.length === 1 || parts[1] === "說明") {
    replyLineMessage(replyToken, getTuitionAdjustmentHelpText());
    ensureTuitionAdjustmentSheet();
    return;
  }

  if (parts[1] === "查詢") {
    if (parts.length < 4) {
      replyLineMessage(replyToken, "❌ 格式：帳務補救 查詢 YYYY/MM 學生姓名");
      return;
    }
    replyLineMessage(replyToken, queryTuitionAdjustments(parts[2], parts[3]));
    return;
  }

  const type = parts[1];
  if (type !== "補收" && type !== "退費") {
    replyLineMessage(replyToken, "❌ 調整類型只支援：補收、退費。\n\n" + getTuitionAdjustmentHelpText());
    return;
  }
  if (parts.length < 11) {
    replyLineMessage(replyToken, "❌ 格式不足。\n\n" + getTuitionAdjustmentHelpText());
    return;
  }

  const targetMonth = parts[2];
  const studentName = parts[3];
  const courseName = parts[4];
  const lessonDate = parts[5].replace(/-/g, "/");
  const startTime = parts[6].replace(/:/g, "");
  const endTime = parts[7].replace(/:/g, "");
  const hours = parseFloat(parts[8]);
  const unitFee = parseFloat(parts[9]);
  const relatedDocId = parts[10];
  const reason = parts.slice(11).join(" ") || "未填寫";

  if (!targetMonth.match(/^\d{4}\/\d{2}$/)) { replyLineMessage(replyToken, "❌ 調整月份格式需為 YYYY/MM。"); return; }
  if (!lessonDate.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) { replyLineMessage(replyToken, "❌ 原上課日期格式需為 YYYY/MM/DD。"); return; }
  if (!startTime.match(/^\d{3,4}$/) || !endTime.match(/^\d{3,4}$/)) { replyLineMessage(replyToken, "❌ 時間格式需為 HHMM，例如 1400 1530。"); return; }
  if (!hours || hours <= 0 || !unitFee || unitFee <= 0) { replyLineMessage(replyToken, "❌ 時數與單價必須大於 0。"); return; }

  const operatorName = getOperatorNameByUserId(userId);
  const amount = Math.round(hours * unitFee) * (type === "退費" ? -1 : 1);
  const sheet = ensureTuitionAdjustmentSheet();
  sheet.appendRow([
    new Date(),
    targetMonth,
    studentName,
    courseName,
    lessonDate,
    startTime,
    endTime,
    hours,
    unitFee,
    amount,
    type,
    relatedDocId,
    reason,
    "待處理",
    operatorName,
    "",
    targetMonth
  ]);

  replyLineMessage(replyToken,
    "✅ 已建立學費調整紀錄\n" +
    "類型：" + type + "\n" +
    "月份：" + targetMonth + "\n" +
    "學生：" + studentName + "\n" +
    "課程：" + courseName + "\n" +
    "金額：" + formatMoney(amount) + "\n" +
    "狀態：待處理"
  );
}

function handleLiffTuitionAdjustment(params: any) {
  const lineUserId = String(params.lineUserId || "").trim();
  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) return { ok: false, message: "權限不足：限行政人員使用。" };

  const type = String(params.adjustmentType || "").trim();
  const targetMonth = String(params.processingMonth || params.month || "").replace("-", "/").trim();
  const originalMonth = String(params.originalMonth || params.month || "").replace("-", "/").trim();
  const studentName = String(params.student || "").trim();
  const courseName = String(params.course || "").trim();
  const lessonDate = String(params.lessonDate || "").replace(/-/g, "/").trim();
  const startTime = String(params.startTime || "").replace(":", "").trim();
  const endTime = String(params.endTime || "").replace(":", "").trim();
  const hours = parseFloat(params.hours);
  const unitFee = parseFloat(params.unitFee);
  const relatedDocId = String(params.relatedDocId || "").trim();
  const reason = String(params.reason || "").trim() || "未填寫";

  if (type !== "補收" && type !== "退費") return { ok: false, message: "調整類型只支援：補收、退費。" };
  if (!targetMonth.match(/^\d{4}\/\d{2}$/)) return { ok: false, message: "調整月份格式需為 YYYY/MM。" };
  if (!originalMonth.match(/^\d{4}\/\d{2}$/)) return { ok: false, message: "原錯誤月份格式需為 YYYY/MM。" };
  if (!studentName || !courseName) return { ok: false, message: "請填寫學生與課程。" };
  if (!lessonDate.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) return { ok: false, message: "原上課日期格式需為 YYYY/MM/DD。" };
  if (!startTime.match(/^\d{3,4}$/) || !endTime.match(/^\d{3,4}$/)) return { ok: false, message: "時間格式需為 HHMM，例如 1400 1530。" };
  if (!hours || hours <= 0 || !unitFee || unitFee <= 0) return { ok: false, message: "時數與單價必須大於 0。" };
  if (!relatedDocId) return { ok: false, message: "請填寫關聯原單號。" };

  const operatorName = getOperatorNameByUserId(lineUserId);
  const amount = Math.round(hours * unitFee) * (type === "退費" ? -1 : 1);
  const sheet = ensureTuitionAdjustmentSheet();
  sheet.appendRow([
    new Date(),
    targetMonth,
    studentName,
    courseName,
    lessonDate,
    startTime,
    endTime,
    hours,
    unitFee,
    amount,
    type,
    relatedDocId,
    reason,
    "待處理",
    operatorName,
    "",
    originalMonth
  ]);

  return {
    ok: true,
    message: `已建立${type}紀錄：${studentName} / ${courseName} / ${formatMoney(amount)}，狀態：待處理。`
  };
}

function handleLiffTuitionAdjustmentOptions(params: any) {
  const lineUserId = String(params.lineUserId || "").trim();
  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) return { ok: false, message: "權限不足：限行政人員使用。" };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  if (!courseSheet) return { ok: false, message: "找不到課程設定表。" };

  const data = courseSheet.getDataRange().getValues();
  const studentSet = new Set<string>();
  const coursesByStudent: any = {};

  for (let i = 1; i < data.length; i++) {
    const studentName = String(data[i][2] || "").trim();
    const courseName = String(data[i][3] || "").trim();
    const unitFee = parseFloat(data[i][4]) || 0;
    const mode = String(data[i][6] || "").trim() === "預收" ? "預收" : "後收";
    if (!studentName || !courseName) continue;

    studentSet.add(studentName);
    if (!coursesByStudent[studentName]) coursesByStudent[studentName] = [];
    let exists = false;
    for (let c = 0; c < coursesByStudent[studentName].length; c++) {
      if (coursesByStudent[studentName][c].name === courseName && coursesByStudent[studentName][c].unitFee === unitFee) {
        exists = true;
        break;
      }
    }
    if (!exists) {
      coursesByStudent[studentName].push({
        name: courseName,
        unitFee,
        mode
      });
    }
  }

  const originalMonth = String(params.originalMonth || params.month || "").replace("-", "/").trim();
  const studentName = String(params.student || "").trim();
  const docIds = originalMonth && studentName ? findTuitionDocIds(originalMonth, studentName) : [];

  return {
    ok: true,
    options: {
      students: Array.from(studentSet).sort(),
      coursesByStudent,
      docIds
    }
  };
}

function getTuitionAdjustmentHelpText() {
  return "🧾 帳務補救\n\n" +
    "建立補收/退費紀錄，不會直接更改舊繳費單。\n\n" +
    "格式：\n" +
    "帳務補救 補收 YYYY/MM 學生 課程 YYYY/MM/DD 開始 結束 時數 單價 原單號 原因\n\n" +
    "範例：\n" +
    "帳務補救 補收 2026/05 均逸 陪伴對話/自學紀錄同行 2026/04/18 1400 1530 1.5 1000 R_2026_04_019 講師漏登\n\n" +
    "查詢：\n" +
    "帳務補救 查詢 2026/05 均逸";
}

function ensureTuitionAdjustmentSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME_TUITION_ADJUSTMENT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_TUITION_ADJUSTMENT);
    sheet.appendRow([
      "建立時間",
      "調整月份",
      "學生姓名",
      "課程名稱",
      "原上課日期",
      "開始時間",
      "結束時間",
      "時數",
      "單價",
      "調整金額",
      "調整類型",
      "關聯原單號",
      "原因",
      "狀態",
      "操作人",
      "備註",
      "原錯誤月份",
      "補收單號",
      "補收PDF",
      "補收單狀態"
    ]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < 20) {
    const headers = ["原錯誤月份", "補收單號", "補收PDF", "補收單狀態"];
    for (let i = 0; i < headers.length; i++) {
      sheet.getRange(1, 17 + i).setValue(headers[i]);
    }
  }
  return sheet;
}

function handleLiffAdjustmentPaymentPreview(params: any) {
  const lineUserId = String(params.lineUserId || "").trim();
  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) return { ok: false, message: "權限不足：限行政人員使用。" };

  const month = String(params.month || params.processingMonth || "").replace("-", "/").trim();
  if (!month.match(/^\d{4}\/\d{2}$/)) return { ok: false, message: "處理月份格式需為 YYYY/MM。" };

  const preview = buildAdjustmentPaymentPreview(month);
  return { ok: true, preview };
}

function handleLiffGenerateAdjustmentPayment(params: any) {
  const lineUserId = String(params.lineUserId || "").trim();
  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) return { ok: false, message: "權限不足：限行政人員使用。" };

  const month = String(params.month || params.processingMonth || "").replace("-", "/").trim();
  if (!month.match(/^\d{4}\/\d{2}$/)) return { ok: false, message: "處理月份格式需為 YYYY/MM。" };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = ensureTuitionAdjustmentSheet();
    const preview = buildAdjustmentPaymentPreview(month);
    if (!preview.hasPending) return { ok: false, message: "沒有可產生補救明細/補收通知的待處理補收資料。" };

    const folder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.PAYMENT_NOTICE);
    const data = sheet.getDataRange().getValues();
    const nextSerial = getNextAdjustmentPaymentSerial(data, month);
    let serial = nextSerial;
    const results: string[] = [];

    for (let i = 0; i < preview.students.length; i++) {
      const student = preview.students[i];
      const docId = "ADJ_" + month.replace("/", "_") + "_" + padSerial(serial);
      serial++;
      const stuData = {
        name: student.name,
        docId,
        saveTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
        total: student.total,
        courses: student.rows.map(function(row: any) {
          return {
            title: "補收：" + row.course,
            detail:
              "原單號：" + row.relatedDocId + "\n" +
              "原上課：" + row.lessonDate + " " + row.startTime + "-" + row.endTime + "\n" +
              "時數/單價：" + row.hours + "hr x $" + row.unitFee + "\n" +
              "原因：" + row.reason + "\n" +
              "補收金額：NT$ " + formatMoney(row.amount)
          };
        })
      };
      const fileUrl = generateSinglePaymentPDF(stuData, folder, month);
      for (let r = 0; r < student.rows.length; r++) {
        const rowNumber = student.rows[r].rowNumber;
        sheet.getRange(rowNumber, 14).setValue("已產生補收單");
        sheet.getRange(rowNumber, 18).setValue(docId);
        sheet.getRange(rowNumber, 19).setValue(fileUrl);
        sheet.getRange(rowNumber, 20).setValue("補收單已產");
      }
      results.push(student.name + "：" + docId + " / NT$ " + formatMoney(student.total));
    }

    return {
      ok: true,
      message: "已產生補救明細/補收通知：\n" + results.join("\n"),
      preview: buildAdjustmentPaymentPreview(month)
    };
  } catch (e) {
    return { ok: false, message: "產生補救明細/補收通知失敗：" + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function buildAdjustmentPaymentPreview(month: string) {
  const sheet = ensureTuitionAdjustmentSheet();
  const data = sheet.getDataRange().getValues();
  const studentsMap: any = {};
  const pendingStudentsMap: any = {};
  let totalAmount = 0;
  let rowCount = 0;
  let pendingAmount = 0;
  let pendingCount = 0;

  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeSheetMonth(data[i][1]);
    const type = String(data[i][10] || "").trim();
    const status = String(data[i][13] || "").trim();
    const paymentDocId = String(data[i][17] || "").trim();
    const paymentPdf = String(data[i][18] || "").trim();
    const amount = parseFloat(data[i][9]) || 0;
    if (rowMonth !== month || type !== "補收" || amount <= 0) continue;

    const studentName = String(data[i][2] || "").trim();
    if (!studentsMap[studentName]) studentsMap[studentName] = { name: studentName, total: 0, rows: [] };
    const row = {
      rowNumber: i + 1,
      course: String(data[i][3] || "").trim(),
      lessonDate: formatSheetDate(data[i][4]),
      startTime: formatSheetTime(data[i][5]),
      endTime: formatSheetTime(data[i][6]),
      hours: parseFloat(data[i][7]) || 0,
      unitFee: parseFloat(data[i][8]) || 0,
      amount,
      relatedDocId: String(data[i][11] || "").trim(),
      reason: String(data[i][12] || "").trim(),
      originalMonth: normalizeSheetMonth(data[i][16]),
      status,
      paymentDocId,
      paymentPdf
    };
    studentsMap[studentName].rows.push(row);
    studentsMap[studentName].total += amount;
    totalAmount += amount;
    rowCount++;

    if (status === "待處理" && !paymentDocId) {
      if (!pendingStudentsMap[studentName]) pendingStudentsMap[studentName] = { name: studentName, total: 0, rows: [] };
      pendingStudentsMap[studentName].rows.push(row);
      pendingStudentsMap[studentName].total += amount;
      pendingAmount += amount;
      pendingCount++;
    }
  }

  const students = Object.keys(studentsMap).sort().map(function(name) { return studentsMap[name]; });
  const pendingStudents = Object.keys(pendingStudentsMap).sort().map(function(name) { return pendingStudentsMap[name]; });
  const items: string[] = [];
  students.forEach(function(student: any) {
    student.rows.forEach(function(row: any) {
      items.push(
        student.name +
        "\n補收金額：NT$ " + formatMoney(row.amount) +
        "\n原月份：" + (row.originalMonth || "未填") +
        "\n上課時間：" + row.lessonDate + " " + row.startTime + "-" + row.endTime +
        "\n課程：" + row.course +
        "\n時數/單價：" + row.hours + "hr x NT$ " + formatMoney(row.unitFee) +
        "\n原單：" + (row.relatedDocId || "未填") +
        "\n狀態：" + (row.status || "未填") +
        (row.paymentDocId ? "\n通知單：" + row.paymentDocId : "") +
        "\n原因：" + (row.reason || "未填")
      );
    });
  });
  if (items.length === 0) items.push(month + " 目前沒有補救明細/補收通知相關補收資料。");

  return {
    title: "補救明細/補收通知預覽",
    month,
    summary: month + " 補救明細/補收通知只讀預覽：" + students.length + " 位學生，" + rowCount + " 筆補收，總金額 NT$ " + formatMoney(totalAmount) + "；待產生 " + pendingCount + " 筆，待產生金額 NT$ " + formatMoney(pendingAmount) + "。",
    items,
    students: pendingStudents,
    totalAmount,
    rowCount,
    pendingAmount,
    pendingCount,
    hasPending: pendingCount > 0
  };
}

function getNextAdjustmentPaymentSerial(data: any[][], month: string) {
  const prefix = "ADJ_" + month.replace("/", "_") + "_";
  let maxSerial = 0;
  for (let i = 1; i < data.length; i++) {
    const docId = String(data[i][17] || "").trim();
    if (docId.indexOf(prefix) !== 0) continue;
    const serial = parseInt(docId.substring(prefix.length), 10);
    if (serial > maxSerial) maxSerial = serial;
  }
  return maxSerial + 1;
}

function padSerial(value: number) {
  if (value < 10) return "00" + value;
  if (value < 100) return "0" + value;
  return String(value);
}

function uniqueValues(values: string[]) {
  const seen: any = {};
  const result: string[] = [];
  for (let i = 0; i < values.length; i++) {
    const value = String(values[i] || "").trim();
    if (!value || seen[value]) continue;
    seen[value] = true;
    result.push(value);
  }
  return result;
}

function findTuitionDocIds(targetMonth: string, studentName: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_FEE);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const docIds: string[] = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const rowMonth = normalizeSheetMonth(data[i][0]);
    const rowStudent = String(data[i][1] || "").trim();
    const docId = String(data[i][9] || "").trim();
    if (rowMonth === targetMonth && rowStudent === studentName && docId && docIds.indexOf(docId) === -1) {
      docIds.push(docId);
    }
  }
  return docIds;
}

function queryTuitionAdjustments(targetMonth: string, studentName: string) {
  const sheet = ensureTuitionAdjustmentSheet();
  const data = sheet.getDataRange().getValues();
  const lines: string[] = [];
  let total = 0;

  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeSheetMonth(data[i][1]);
    const rowStudent = String(data[i][2] || "").trim();
    if (rowMonth !== targetMonth || rowStudent !== studentName) continue;

    const amount = parseFloat(data[i][9]) || 0;
    total += amount;
    lines.push(
      data[i][10] + " " +
      data[i][4] + " " +
      data[i][5] + "-" + data[i][6] + " " +
      data[i][3] + " " +
      formatMoney(amount) + "\n" +
      "原單：" + data[i][11] + "｜狀態：" + data[i][13] + "\n" +
      "原因：" + data[i][12]
    );
  }

  if (lines.length === 0) return "ℹ️ 找不到 " + targetMonth + "「" + studentName + "」的帳務補救紀錄。";
  return "🧾 帳務補救查詢\n" +
    "月份：" + targetMonth + "\n" +
    "學生：" + studentName + "\n\n" +
    lines.join("\n\n") + "\n\n" +
    "調整合計：" + formatMoney(total);
}

function normalizeSheetMonth(value: any) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy/MM");
  const text = String(value || "").trim();
  if (text.match(/^\d{4}[\/-]\d{2}/)) return text.substring(0, 7).replace("-", "/");
  return text;
}

function formatSheetDate(value: any) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy/MM/dd");
  const text = String(value || "").trim().replace(/-/g, "/");
  if (text.match(/^\d{4}\/\d{1,2}\/\d{1,2}/)) {
    const parts = text.split("/");
    return parts[0] + "/" + parts[1].padStart(2, "0") + "/" + parts[2].padStart(2, "0");
  }
  return text;
}

function formatSheetTime(value: any) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  const digits = String(value || "").trim().replace(":", "").padStart(4, "0");
  if (!digits.match(/^\d{4}$/)) return String(value || "").trim();
  return digits.substring(0, 2) + ":" + digits.substring(2, 4);
}

function getOperatorNameByUserId(userId: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return userId;
  const data = teacherSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === userId) return String(data[i][0] || userId);
  }
  return userId;
}

function handlePaymentDocCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const targetId = event.source.groupId || userId;
  const isAdmin = ADMIN_LIST.indexOf(userId) > -1;
  if (!isAdmin) { replyLineMessage(replyToken, "❌ 權限不足"); return; }

  const parts = userMsg.split(" ");
  let targetMonth = ""; let targetStudent = "";
  if (parts.length >= 2) targetMonth = parts[1];
  else {
    const nextM = new Date(); nextM.setMonth(nextM.getMonth() + 1);
    targetMonth = Utilities.formatDate(nextM, Session.getScriptTimeZone(), "yyyy/MM");
  }
  if (parts.length >= 3) targetStudent = parts[2];
  if (!targetMonth.match(/^\d{4}\/\d{2}$/)) { replyLineMessage(replyToken, "❌ 月份格式錯誤"); return; }
  
  try {
    const resultMsg = createPaymentNoticesBatch(targetMonth, targetStudent);
    replyLineMessage(replyToken, resultMsg);
  } catch (e) {
    replyLineMessage(replyToken, "❌ 失敗：" + e.toString());
  }
}

function createPaymentNoticesBatch(targetMonth: string, targetName: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("學費結算表");
  if (!sheet) return "❌ 找不到學費結算表";
  const targetFolder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.PAYMENT_NOTICE); 
  const data = sheet.getDataRange().getValues();
  const studentsMap: any = {};
  for (let i = 1; i < data.length; i++) {
    const rowMonthStr = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "yyyy/MM") : data[i][0];
    if (rowMonthStr !== targetMonth) continue;
    const sName = data[i][1];
    if (targetName && targetName !== "" && sName !== targetName) continue;
    if (!studentsMap[sName]) { studentsMap[sName] = { name: sName, docId: "", saveTime: "", total: 0, courses: [], updateRow: -1 }; }
    studentsMap[sName].courses.push({ title: data[i][2], detail: data[i][4].replace(/^\s+|\s+$/g, '').replace(/\n+/g, "\n") });
    if (data[i][8] !== "" && data[i][9] !== "") {
      studentsMap[sName].total = data[i][8]; studentsMap[sName].docId = data[i][9];
      studentsMap[sName].saveTime = (data[i][10] instanceof Date) ? Utilities.formatDate(data[i][10], Session.getScriptTimeZone(), "yyyy/MM/dd") : data[i][10];
      studentsMap[sName].updateRow = i + 1;
    }
  }
  const results: string[] = []; let count = 0;
  for (const name in studentsMap) {
    const stuData = studentsMap[name];
    if (!stuData.docId || stuData.total === 0) continue;
    const fileName = buildDocPdfFileName(stuData.docId, "繳費單", stuData.name);
    const legacyFileName = "繳費單_" + stuData.name + "_" + stuData.docId + ".pdf";
    const existingFile = findExistingFileByNames(targetFolder, [fileName, legacyFileName]);
    let fileUrl = existingFile ? existingFile.getUrl() : "";
    if (fileUrl) {
      results.push("✅ (既有) " + name + " - " + fileUrl); 
    } else { 
      fileUrl = generateSinglePaymentPDF(stuData, targetFolder, targetMonth);
      results.push("🆕 (新增) " + name + " - " + fileUrl); 
    }
    
    if (stuData.updateRow > 0) {
      sheet.getRange(stuData.updateRow, 16).setValue(fileUrl);
      sheet.getRange(stuData.updateRow, 17).setValue("繳費單已產");
    }
    recordDocumentEntry({
      month: targetMonth,
      docType: "繳費單",
      targetType: "學生",
      targetName: name,
      docId: stuData.docId,
      sourceSheet: SHEET_NAME_FIN_FEE,
      sourceKey: targetMonth + "|" + name,
      amount: stuData.total,
      pdfUrl: fileUrl,
      generateStatus: "已產生",
      emailStatus: "待寄送",
      lineStatus: "未推播",
      note: "繳費單"
    });
    count++;
  }
  if (count === 0) return "⚠️ " + targetMonth + " 無資料。";
  
  const msg = "📁 [學費單_" + targetMonth.replace("/", "-") + "]\n\n" + results.join("\n\n");
  return msg;
}

function generateSinglePaymentPDF(stuData: any, folder: any, targetMonth: string) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID_PAYMENT);
  const fileName = buildDocPdfFileName(stuData.docId, "繳費單", stuData.name);
  const copyName = fileName.replace(/\.pdf$/i, "");
  const newFile = templateFile.makeCopy(copyName, folder);
  const newDoc = DocumentApp.openById(newFile.getId());
  const body = newDoc.getBody();
  const parts = targetMonth.split("/"); const deadlineDate = new Date(parseInt(parts[0]), parseInt(parts[1]), 10);
  const deadlineStr = Utilities.formatDate(deadlineDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  const rowsText = [];
  for (let i = 0; i < stuData.courses.length; i++) rowsText.push(stuData.courses[i].title + "\n" + stuData.courses[i].detail);
  body.replaceText("{{存檔時間}}", stuData.saveTime || ""); body.replaceText("{{單據編號}}", stuData.docId || ""); 
  body.replaceText("{{學生姓名}}", stuData.name || ""); body.replaceText("{{課程名稱}}", rowsText.join("\n\n"));
  body.replaceText("{{日期/時間\\(金額\\)}}", ""); 
  body.replaceText("{{個人小計}}", formatMoney(stuData.total)); 
  body.replaceText("{{個人小計大寫}}", digitToChinese(stuData.total)); 
  body.replaceText("{{繳費期限}}", deadlineStr);
  newDoc.saveAndClose();
  const pdfBlob = newFile.getAs(PDF_MIME_TYPE); pdfBlob.setName(fileName); 
  const pdfFile = folder.createFile(pdfBlob); newFile.setTrashed(true);
  return pdfFile.getUrl();
}

function handleBatchAllowanceCommand(event: any, userMsg: string) {
  replyLineMessage(event.replyToken, "⚠️ 批次開領據已停用。請使用：開領據 YYYY/MM 講師名");
}

function handleBatchSendAllowanceEmailCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const targetId = event.source.groupId || userId;
  const parts = userMsg.split(" ");
  if (parts.length < 2) { replyLineMessage(replyToken, "❌ 格式：寄送領據 YYYY/MM"); return; }
  const targetMonth = parts[1];
  replyLineMessage(replyToken, "⏳ 正在寄送 " + targetMonth + " 的領據 Email...");
  try {
    pushLineMessage(targetId, processBatchAllowanceEmail(targetMonth));
  } catch (e) { pushLineMessage(targetId, "❌ 寄送執行錯誤：" + e.toString()); }
}

function processBatchAllowanceEmail(targetMonth: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("鐘點結算表");
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!sheet || !teacherSheet) return "❌ 資料表缺失";
  const data = sheet.getDataRange().getValues();
  const teacherData = teacherSheet.getDataRange().getValues();
  const emailMap: any = {};
  for (let k = 1; k < teacherData.length; k++) emailMap[teacherData[k][0]] = teacherData[k][2];

  let count = 0; const failList: string[] = [];
  for (let i = 1; i < data.length; i++) {
    let rowMonth = data[i][0];
    if (rowMonth instanceof Date) rowMonth = Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM");
    const status = data[i][11]; const pdfUrl = data[i][10]; const tName = data[i][1];
    
    if (rowMonth == targetMonth && status === "待寄送" && pdfUrl && pdfUrl !== "") {
       const email = emailMap[tName]; let newStatus = "寄送失敗";
       if (email && email.indexOf("@") > -1) {
         try {
           const fileIdMatch = pdfUrl.match(/[-\w]{25,}/);
           if (fileIdMatch) {
             const file = DriveApp.getFileById(fileIdMatch[0]);
             const subject = EMAIL_CONFIG.TEACHER.SUBJECT.replace("{{Month}}", targetMonth).replace("{{Name}}", tName);
             const body = EMAIL_CONFIG.TEACHER.BODY.replace("{{Name}}", tName);
             GmailApp.sendEmail(email, subject, body, { attachments: [file.getAs(PDF_MIME_TYPE)] });
             newStatus = "已寄送"; count++;
           } else { newStatus = "失敗(連結錯)"; failList.push(tName); }
         } catch (err: any) { newStatus = "失敗(" + err.message + ")"; failList.push(tName); }
       } else { newStatus = "失敗(無Email)"; failList.push(tName); }
       sheet.getRange(i+1, 12).setValue(newStatus);
    }
  }
  return "✅ 領據寄送結束。\n成功：" + count + " 封\n失敗：" + failList.length;
}

function handleSingleAllowanceCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const parts = userMsg.split(" ");
  if (parts.length < 3) { replyLineMessage(replyToken, "❌ 格式：開領據 YYYY/MM 講師名"); return; }
  
  const targetMonth = parts[1]; const targetName = parts[2];
  const dateStr = (parts.length >= 4) ? parts[3] : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("鐘點結算表");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  let rowData = null; let rowIndex = -1;

  for (let i = data.length - 1; i >= 1; i--) {
    let rowMonth = data[i][0];
    if (rowMonth instanceof Date) rowMonth = Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM");
    if (rowMonth == targetMonth && data[i][1] == targetName) {
      if (data[i][6] && data[i][6] !== "" && data[i][7]) { rowData = data[i]; rowIndex = i + 1; break; }
    }
  }

  if (!rowData) { replyLineMessage(replyToken, "❌ 找不到結算資料。"); return; }
  
  if (rowData[10] && rowData[10] !== "") {
    replyLineMessage(replyToken, "⚠️ 此筆領據已產製過！\n📂 連結：" + rowData[10] + "\n若需重新產製，請先手動清空試算表 K 欄。");
    return;
  }
  
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return;
  const teacherData = teacherSheet.getDataRange().getValues();
  let tInfo: any = {};
  for (let k=1; k<teacherData.length; k++) {
     if(teacherData[k][0] === targetName) {
        tInfo = { email: teacherData[k][2], pid: teacherData[k][3], addr: teacherData[k][4], phone: teacherData[k][5], method: teacherData[k][6], bank: teacherData[k][7], account: teacherData[k][8] }; break;
     }
  }

  const fullContent = [rowData[2], rowData[3], rowData[4], rowData[5]].filter(function(part: any) { return String(part || "").trim() !== ""; }).join("\n");
  const state: any = {
     name: targetName, amount: rowData[6], taxAmount: rowData[12] || 0, nhiAmount: rowData[13] || 0, netAmount: rowData[14] || rowData[6],
     docId: rowData[7], date: dateStr, detail: fullContent, totalHours: rowData[5] || rowData[2], rowIndex: rowIndex,
     email: tInfo.email, pid: tInfo.pid, addr: tInfo.addr, phone: tInfo.phone, method: tInfo.method, bank: tInfo.bank, account: tInfo.account
  };

  try {
     const folder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.ALLOWANCE);
     const pdfResult = generateAllowancePDF(state, folder);
     state.pdfUrl = pdfResult.url;
     state.pdfId = pdfResult.id; 
     
     CacheService.getScriptCache().put("ALLOW_" + userId, JSON.stringify(state), 600);
     replyAllowancePreview(replyToken, state); 
  } catch(e) { replyLineMessage(replyToken, "❌ 預覽產製失敗：" + e.toString()); }
}

function executeAllowanceSaveOnly(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = "ALLOW_" + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (!cached) { replyLineMessage(replyToken, "⚠️ 逾時。"); return; }
  const state = JSON.parse(cached);
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("鐘點結算表");
    if (!sheet) return;
    sheet.getRange(state.rowIndex, 11).setValue(state.pdfUrl);
    sheet.getRange(state.rowIndex, 12).setValue("待寄送");
  } catch(e) { replyLineMessage(replyToken, "⚠️ 寫入失敗：" + e.toString()); return; }
  
  CacheService.getScriptCache().remove(cacheKey);
  replyLineMessage(replyToken, "✅ 領據已寫入紀錄表！\n狀態：待寄送\n📂 檔案：" + state.pdfUrl);
}

function generateAllowancePDF(state: any, folder: any) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID_ALLOWANCE);
  const fileName = buildDocPdfFileName(state.docId, "領據", state.name); 
  const legacyFileName = "領據_" + state.name + "_" + state.date.replace(/\//g,"") + "_" + state.docId + ".pdf"; 
  
  const existing = findExistingFileByNames(folder, [fileName, legacyFileName]);
  if (existing) {
     return { url: existing.getUrl(), id: existing.getId() };
  }

  const newFile = templateFile.makeCopy(fileName.replace(/\.pdf$/i, ""), folder);
  const newDoc = DocumentApp.openById(newFile.getId());
  const body = newDoc.getBody();
  const replacements: any = {
    "{{日期}}": state.date, "{{單據編號}}": state.docId, "{{講師姓名}}": state.name, "{{身分證字號}}": state.pid || "", 
    "{{戶籍地址}}": state.addr || "", "{{電話}}": state.phone || "", "{{支領方式}}": state.method || "", 
    "{{銀行}}": state.bank || "", "{{帳號}}": state.account || "", "{{類別}}": "鐘點費", "{{課程名稱}}": state.detail, 
    "{{課程/學生}}": state.detail, "{{課程日期/時間(金額)}}": state.detail, "{{授課內容}}": state.detail || "", 
    "{{授課總時數}}": state.totalHours || "", "{{應付總額}}": formatMoney(state.amount), "{{應付小計}}": formatMoney(state.amount), 
    "{{應付小計大寫}}": digitToChinese(state.amount), "{{扣繳稅額}}": formatMoney(state.taxAmount), 
    "{{補充保費}}": formatMoney(state.nhiAmount), "{{實發金額}}": formatMoney(state.netAmount), 
    "{{實發金額大寫}}": digitToChinese(state.netAmount), "{{地址}}": state.addr || ""
  };
  for (const key in replacements) body.replaceText(key.replace(/([()[\]{}*+?^$|#\s])/g, "\\$1"), replacements[key]);
  body.replaceText("\\{\\{.*?\\}\\}", ""); 
  newDoc.saveAndClose();
  
  const pdfBlob = newFile.getAs(PDF_MIME_TYPE); pdfBlob.setName(fileName); const pdfFile = folder.createFile(pdfBlob); newFile.setTrashed(true);
  return { url: pdfFile.getUrl(), id: pdfFile.getId() };
}

function generateReceiptPDF(state: any, folder: any) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID_RECEIPT);
  const fileName = buildDocPdfFileName(state.docId, "收據", state.name); 
  const legacyFileName = "收據_" + state.name + "_" + state.date.replace(/\//g,"") + "_" + state.docId + ".pdf"; 
  
  const existing = findExistingFileByNames(folder, [fileName, legacyFileName]);
  if (existing) {
     return { url: existing.getUrl(), id: existing.getId() };
  }

  const newFile = templateFile.makeCopy(fileName.replace(/\.pdf$/i, ""), folder);
  const newDoc = DocumentApp.openById(newFile.getId());
  const body = newDoc.getBody();
  const replacements: any = {
    "{{日期}}": state.date, "{{收款日期}}": state.date, "{{單據編號}}": state.docId, "{{編號}}": state.docId, 
    "{{繳款人}}": state.name, "{{學生姓名}}": state.name, "{{姓名}}": state.name, "{{身分證字號/統一編號}}": state.pid || "", 
    "{{身分證字號}}": state.pid || "", "{{金額}}": formatMoney(state.amount), "{{個人小計}}": formatMoney(state.amount), 
    "{{金額大寫}}": digitToChinese(state.amount), "{{個人小計(大寫)}}": digitToChinese(state.amount), 
    "{{繳費方式}}": state.method, "{{收入用途/類別}}": state.category, "{{課程名稱}}": state.detail || "" 
  };
  for (const key in replacements) body.replaceText(key.replace(/([()[\]{}*+?^$|#\s])/g, "\\$1"), replacements[key]);
  body.replaceText("\\{\\{.*?\\}\\}", ""); 
  newDoc.saveAndClose();
  
  const pdfBlob = newFile.getAs(PDF_MIME_TYPE); pdfBlob.setName(fileName); const pdfFile = folder.createFile(pdfBlob); newFile.setTrashed(true);
  return { url: pdfFile.getUrl(), id: pdfFile.getId() };
}

function replyReceiptPreview(token: string, state: any) {
  const card = {
    "type": "bubble",
    "body": {
      "type": "box", "layout": "vertical",
      "contents": [
        { "type": "text", "text": "收據預覽", "weight": "bold", "size": "xl", "color": "#1DB446" },
        { "type": "separator", "margin": "md" },
        { "type": "text", "text": "👤 " + state.name + "\n💰 " + state.method + " $" + state.amount + "\n📅 " + state.date, "wrap": true, "margin": "md" },
        { "type": "button", "style": "link", "action": { "type": "uri", "label": "📄 查看真實 PDF 檔案", "uri": state.pdfUrl }, "margin": "md" }
      ]
    },
    "footer": {
      "type": "box", "layout": "horizontal", "spacing": "sm",
      "contents": [
        { "type": "button", "style": "primary", "action": { "type": "postback", "label": "✅ 確認並寫入", "data": "action=rcpt_save_only" } },
        { "type": "button", "style": "secondary", "action": { "type": "postback", "label": "❌ 取消並銷毀", "data": "action=doc_cancel_rcpt" } }
      ]
    }
  };
  replyFlexMessage(token, "收據預覽", card);
}

function replyAllowancePreview(token: string, state: any) {
  let taxNote = "";
  if (state.taxAmount > 0) taxNote += "\n➖ 扣繳稅額：$" + formatMoney(state.taxAmount);
  if (state.nhiAmount > 0) taxNote += "\n➖ 補充保費：$" + formatMoney(state.nhiAmount);
  const card = {
    "type": "bubble",
    "body": {
      "type": "box", "layout": "vertical",
      "contents": [
        { "type": "text", "text": "領據預覽", "weight": "bold", "size": "xl", "color": "#1DB446" },
        { "type": "separator", "margin": "md" },
        { "type": "text", "text": "👨‍🏫 講師：" + state.name + "\n💰 應付總額：$" + formatMoney(state.amount) + taxNote + "\n✨ 實發總額：$" + formatMoney(state.netAmount) + "\n📅 領取日期：" + state.date, "wrap": true, "margin": "md" },
        { "type": "button", "style": "link", "action": { "type": "uri", "label": "📄 查看真實 PDF 檔案", "uri": state.pdfUrl }, "margin": "md" }
      ]
    },
    "footer": {
      "type": "box", "layout": "horizontal", "spacing": "sm",
      "contents": [
        { "type": "button", "style": "primary", "action": { "type": "postback", "label": "✅ 確認並寫入", "data": "action=allowance_save_only" } },
        { "type": "button", "style": "secondary", "action": { "type": "postback", "label": "❌ 取消並銷毀", "data": "action=doc_cancel_allowance" } }
      ]
    }
  };
  replyFlexMessage(token, "領據預覽", card);
}

function replyButtons(token: string, text: string, actions: any[]) {
  const btnTemplate = {
    "type": "template",
    "altText": "請選擇",
    "template": {
      "type": "buttons",
      "text": text.substring(0, 160),
      "actions": actions
    }
  };
  LineClient.reply(token, [btnTemplate]);
}

function getOrCreateFolder(folderName: string, rootId: string) {
  const parentFolder = (rootId && rootId !== "") ? DriveApp.getFolderById(rootId) : DriveApp.getRootFolder();
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function digitToChinese(n: number) {
  if (!/^(0|[1-9]\d*)(\.\d+)?$/.test(String(n))) return "數據非法";
  let unit = "仟佰拾億仟佰拾萬仟佰拾元角分";
  let valStr = String(n) + "00";
  const p = valStr.indexOf('.'); 
  if (p >= 0) valStr = valStr.substring(0, p) + valStr.substr(p + 1, 2);
  unit = unit.substr(unit.length - valStr.length);
  let str = "";
  for (let i = 0; i < valStr.length; i++) str += '零壹貳參肆伍陸柒捌玖'.charAt(parseInt(valStr.charAt(i))) + unit.charAt(i);
  return str.replace(/零(仟|佰|拾|角)/g, "零").replace(/(零)+/g, "零").replace(/零(萬|億|元)/g, "$1").replace(/(億)萬|壹(拾)/g, "$1$2").replace(/^元零?|零分/g, "").replace(/元$/g, "元整");
}

function formatMoney(n: number): string { return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","); }

function handleGenReceiptCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const parts = userMsg.split(" ");
  if (parts.length < 4) { replyLineMessage(replyToken, "請輸入：開一般收據 YYYY/MM/DD 姓名 金額"); return; }

  const dateStr = parts[1];
  const state = { date: dateStr, month: dateStr.substring(0, 7), name: parts[2], amount: parseInt(parts[3]), step: "WAIT_CAT" };
  if (isNaN(state.amount)) { replyLineMessage(replyToken, "金額格式錯誤。"); return; }

  CacheService.getScriptCache().put("GEN_RCPT_" + userId, JSON.stringify(state), 600);
  const buttons = [
    { "type": "postback", "label": "入會費", "data": "action=gen_rcpt_cat&val=入會費" },
    { "type": "postback", "label": "常年會費", "data": "action=gen_rcpt_cat&val=常年會費" },
    { "type": "postback", "label": "捐款", "data": "action=gen_rcpt_cat&val=捐款" },
    { "type": "postback", "label": "入會+常年會費", "data": "action=gen_rcpt_cat&val=入會費+常年會費" }
  ];
  replyButtons(replyToken, "💳 類別：" + state.name + " (" + state.date + ")\n請選擇用途：", buttons);
}

function handleGenReceiptCategory(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = "GEN_RCPT_" + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (!cached) return;
  const state = JSON.parse(cached); state.category = postbackData.split("val=")[1];
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(state), 600);
  const buttons = [ { "type": "postback", "label": "現金", "data": "action=gen_rcpt_method&val=現金" }, { "type": "postback", "label": "匯款", "data": "action=gen_rcpt_method&val=匯款" } ];
  replyButtons(replyToken, "💰 類別：" + state.category + "\n請選擇繳費方式：", buttons);
}

function handleGenReceiptMethod(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cacheKey = "GEN_RCPT_" + userId;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (!cached) return;
  
  const state = JSON.parse(cached);
  if (state.step === "PREVIEW") return;
  state.method = postbackData.split("val=")[1]; state.step = "PREVIEW";
  
  const memberInfo = lookupGeneralMemberData(state.name, state.category);
  state.pid = memberInfo.pid; state.email = memberInfo.email; state.docId = getNextGenReceiptNumber();

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID_GEN_RECEIPT);
    const pdfResult = generateGeneralReceiptPDF(state, folder);
    state.pdfUrl = pdfResult.url;
    state.pdfId = pdfResult.id; 
    
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(state), 600);
    replyGenReceiptPreview(replyToken, state); 
  } catch(e) { replyLineMessage(replyToken, "❌ 預覽產生失敗：" + e.toString()); }
}

function handleSendGenReceiptCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const parts = userMsg.trim().split(/\s+/);
  if (parts.length < 3) { replyLineMessage(replyToken, "格式：寄一般收據 YYYY/MM 姓名"); return; }
  const targetMonth = parts[1]; const targetName = parts[2];
  if (!targetMonth.match(/^\d{4}\/\d{2}$/)) { replyLineMessage(replyToken, "月份格式錯誤，請使用 YYYY/MM"); return; }
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_GEN_RECORD);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  let targetRow = -1; let pdfUrl = ""; let category = ""; let currentEmailStatus = "";
  
  for (let i = data.length - 1; i >= 1; i--) {
     const rowDateRaw = data[i][2];
     const rowMonth = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, Session.getScriptTimeZone(), "yyyy/MM") : String(rowDateRaw).substring(0, 7);
     if (rowMonth === targetMonth && String(data[i][3]).trim() === targetName) {
         targetRow = i + 1; category = data[i][5]; pdfUrl = data[i][7]; currentEmailStatus = data[i][8]; break;
     }
  }
  
  if (targetRow == -1) { replyLineMessage(replyToken, "⚠️ 找不到紀錄。"); return; }
  if (currentEmailStatus === "已寄送") { replyLineMessage(replyToken, "⚠️ 已寄送完成，系統阻擋重複寄送。"); return; }
  
  const memberInfo = lookupGeneralMemberData(targetName, category);
  const email = memberInfo.email; let newStatus = "寄送失敗"; let reportMsg = "";
  
  if (email && email.indexOf("@") > -1) {
     try {
       const fileIdMatch = pdfUrl.match(/[-\w]{25,}/);
       if (fileIdMatch) {
          const file = DriveApp.getFileById(fileIdMatch[0]);
          let subject = ""; let body = ""; const yearStr = targetMonth.split("/")[0];
          if (category.indexOf("入會費") > -1) { subject = "【入會證明】歡迎加入本會 - " + targetName; body = "附件為入會費收據。"; } 
          else if (category === "常年會費") { subject = "【收據】" + yearStr + " 年度會費收據 - " + targetName; body = "附件為常年會費收據。"; } 
          else if (category === "捐款") { subject = "【感謝狀】感謝愛心捐款 - " + targetName; body = "附件為捐款電子收據。"; } 
          else { subject = "【收據】" + category + "收據 (" + targetName + ")"; body = "附件為本期收據。"; }
          GmailApp.sendEmail(email, subject, body, { attachments: [file.getAs(PDF_MIME_TYPE)] });
          newStatus = "已寄送"; reportMsg = "📧 發送成功！";
       } else { newStatus = "失敗:無效連結"; reportMsg = "❌ 找不到 PDF 連結。"; }
     } catch(e: any) { newStatus = "失敗:" + e.message; reportMsg = "❌ 系統報錯：" + e.message; }
  } else { newStatus = "無Email"; reportMsg = "❌ 找不到 Email。"; }
  
  sheet.getRange(targetRow, 9).setValue(newStatus);
  replyLineMessage(replyToken, reportMsg);
}

function replyGenReceiptPreview(replyToken: string, state: any) {
  const card = {
    "type": "bubble",
    "body": {
      "type": "box", "layout": "vertical",
      "contents": [
        { "type": "text", "text": "🗂️ 一般收據預覽", "weight": "bold", "size": "xl", "color": "#2980b9" },
        { "type": "separator", "margin": "md" },
        { "type": "text", "text": "單號：" + state.docId + "\n姓名：" + state.name + "\n日期：" + state.date + "\n金額：$" + state.amount, "wrap": true, "margin": "md" },
        { "type": "button", "style": "link", "action": { "type": "uri", "label": "📄 查看真實 PDF 檔案", "uri": state.pdfUrl }, "margin": "md" }
      ]
    },
    "footer": {
      "type": "box", "layout": "horizontal", "spacing": "sm",
      "contents": [
        { "type": "button", "style": "primary", "action": { "type": "postback", "label": "✅ 確認並寫入", "data": "action=gen_rcpt_save" } },
        { "type": "button", "style": "secondary", "action": { "type": "postback", "label": "❌ 取消並銷毀", "data": "action=doc_cancel_gen_rcpt" } }
      ]
    }
  };
  replyFlexMessage(replyToken, "一般收據預覽", card);
}

function lookupGeneralMemberData(name: string, category: string) {
  const memberSs = SpreadsheetApp.openById(SHEET_ID_MEMBER);
  let pid = ""; let email = "";
  try {
    if (category === "捐款") {
      const sheet = memberSs.getSheetByName("表單回覆 2"); 
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        for (let i = data.length - 1; i >= 1; i--) { if (data[i][1] === name) { pid = data[i][2]; email = data[i][4]; break; } }
      }
    } else {
      const sheet = memberSs.getSheetByName("表單回應 1");
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        for (let i = data.length - 1; i >= 1; i--) { if (data[i][2] === name) { pid = data[i][5]; email = data[i][11]; break; } }
      }
    }
  } catch(e) { Logger.log("Lookup Error: " + e); }
  return { pid: pid || "", email: email || "" };
}

function getNextGenReceiptNumber() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_GEN_RECORD);
  const prefix = "M_" + new Date().getFullYear() + "_"; let maxNum = 0;
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const docId = data[i][1];
      if (docId && docId.indexOf(prefix) === 0) {
        const num = parseInt(docId.replace(prefix, ""), 10); if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
  }
  return prefix + ("000" + (maxNum + 1)).slice(-3);
}

function generateGeneralReceiptPDF(state: any, folder: any) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID_GEN_RECEIPT);
  const fileName = buildDocPdfFileName(state.docId, "一般收據", state.name, state.category);
  const legacyFileName = "一般收據_" + state.name + "_" + state.category + "_" + state.docId;
  
  const existing = findExistingFileByNames(folder, [fileName, legacyFileName, legacyFileName + ".pdf"]);
  if (existing) {
     return { url: existing.getUrl(), id: existing.getId() };
  }

  const newFile = templateFile.makeCopy(fileName.replace(/\.pdf$/i, ""), folder);
  const newDoc = DocumentApp.openById(newFile.getId());
  const body = newDoc.getBody();
  const replacements: any = { "{{收款日期}}": state.date, "{{單據編號}}": state.docId, "{{姓名}}": state.name, "{{身分證字號/統一編號}}": state.pid, "{{收入用途/類別}}": state.category, "{{繳費方式}}": state.method, "{{個人小計}}": formatMoney(state.amount), "{{個人小計(大寫)}}": digitToChinese(state.amount) };
  for (const key in replacements) body.replaceText(key.replace(/([()[\]{}*+?^$|#\s])/g, "\\$1"), replacements[key]);
  newDoc.saveAndClose();
  
  const pdfBlob = newFile.getAs(PDF_MIME_TYPE); pdfBlob.setName(fileName); const pdfFile = folder.createFile(pdfBlob); newFile.setTrashed(true); 
  return { url: pdfFile.getUrl(), id: pdfFile.getId() };
}
