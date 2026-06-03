// ==========================================
// 🎮 Main_Controller.ts : 程式入口 (LINE Bot doPost & LIFF doGet 控制器)
// ==========================================

// LINE Bot 訊息入口
function doPost(e: GoogleAppsScript.Events.DoPost) {
  try {
    const msgObj = JSON.parse(e.postData.contents);
    const event = msgObj.events[0];
    if (!event) return;

    if (event.type === 'postback') {
      handlePostback(event);
      return;
    }
    if (event.type === 'message' && event.message.type === 'text') {
      handleMessage(event);
      return;
    }
  } catch (err) {
    Logger.log("Main Error: " + err.toString());
  }
  return;
}

// LIFF API 後端入口
function doGet(e: any) {
  try {
    const action = e.parameter.action;
    const lineUserId = e.parameter.lineUserId;

    let result: any = { ok: false, message: "Unknown action" };

    switch (action) {
      case "me":
        // 讀取講師個人狀態與權限
        result = handleLiffMe(lineUserId);
        break;
      case "getFormOptions":
        // 讀取登記表單所需的下拉選項
        result = handleLiffFormOptions(lineUserId);
        break;
      case "register":
        // 登記課程或預排
        result = handleLiffRegister(e.parameter);
        break;
      case "getUnverified":
        // 取得待核銷課程
        result = handleLiffGetUnverified(lineUserId);
        break;
      case "verifySchedule":
        // 執行核銷
        result = handleLiffVerifySchedule(e.parameter);
        break;
      case "leave":
        // 提交請假申請
        result = handleLiffLeave(e.parameter);
        break;
      case "adminTask":
        // 執行行政/財務月結任務
        result = handleLiffAdminTask(e.parameter);
        break;
      case "unbind":
        // 解除 LINE 綁定
        result = handleLiffUnbind(lineUserId);
        break;
      case "verifyAndBind":
        // 講師身分驗證與綁定
        result = handleLiffVerifyAndBind(e.parameter.name, lineUserId);
        break;
      case "health":
        result = { ok: true, message: "Dev GAS API OK", time: new Date().toISOString() };
        break;
      default:
        result = { ok: false, message: `不支援的 API 指令: ${action}` };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: "API 系統錯誤: " + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleMessage(event: any) {
  const userMsg = event.message.text.trim();
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const cache = CacheService.getScriptCache();

  // 1. 呼叫主選單
  if (userMsg === "選單" || userMsg === "功能" || userMsg === "menu" || userMsg === "主選單") {
      replyMainMenu(replyToken);
      return; 
  }

  // ★ 新增：呼叫稅務專區選單 ★
  if (userMsg === "稅務") {
      replyTaxMenu(replyToken);
      return;
  }

  // 2. ★★★ 對話狀態機 (Conversational Flow) ★★★
  // 檢查是否正在回答機器人的問題
  const cachedState = cache.get(userId);
  if (cachedState) {
      const state = JSON.parse(cachedState);
      
      // (A) 正在等待輸入月份 (學費/鐘點)
      if (state.status === "WAITING_FIN_MONTH") {
          const monthInput = normalizeMonthInput(userMsg); 
          if (monthInput) {
             cache.remove(userId); // 清除狀態
             if (state.type === "FEE") {
                 handleTuitionCalculation(event, "學費試算 " + monthInput); 
             } else if (state.type === "SALARY") {
                 handleSalaryCalculation(event, "鐘點試算 " + monthInput);
             } else if (state.type === "PAYMENT_NOTICE") {
                 handlePaymentDocCommand(event, "產生繳費單 " + monthInput);
             }
          } else {
             if (userMsg === "取消") {
                 cache.remove(userId);
                 replyLineMessage(replyToken, "已取消試算。");
             } else {
                 replyLineMessage(replyToken, "❌ 日期格式不清楚，請輸入如「2026/02」或「02」(本月)。\n(若要取消請輸入「取消」)");
             }
          }
          return;
      }

      // (B) 正在等待收據日期
      if (state.status === "WAITING_RECEIPT_DATE") {
          handleReceiptDateInput(event, userMsg, state);
          return;
      }

      // (C) 正在等待請假時間
      if (state.status === "WAITING_LEAVE_TIME") {
          processLeaveInput(event, userMsg, state); 
          return;
      }

      // (D) 正在等待講師登錄時間 (簡單判斷數字)
      if (userMsg.match(/\d/)) {
          processFlowReport(event, userMsg, cachedState);
          return; 
      } else {
          cache.remove(userId);
      }
  }

  // 3. [啟動對話] 當收到關鍵字時，進入「等待輸入」狀態
  
  // ★ 學費試算 -> 進入問答模式
  if (userMsg === "學費試算") {
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM");
      CacheService.getScriptCache().put(userId, JSON.stringify({ status: "WAITING_FIN_MONTH", type: "FEE" }), 600);
      replyLineMessage(replyToken, "📊 [學費試算]\n請輸入試算月份？\n\n(例如輸入「" + nowStr + "」，或直接輸入「02」)");
      return;
  }

  // ★ 鐘點試算 -> 進入問答模式
  if (userMsg === "鐘點試算") {
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM");
      CacheService.getScriptCache().put(userId, JSON.stringify({ status: "WAITING_FIN_MONTH", type: "SALARY" }), 600);
      replyLineMessage(replyToken, "💸 [鐘點試算]\n請輸入試算月份？\n\n(例如輸入「" + nowStr + "」，或直接輸入「02」)");
      return;
  }
  
  // ★ 產生繳費單 -> 進入問答模式
  if (userMsg === "產生繳費單") {
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM");
      CacheService.getScriptCache().put(userId, JSON.stringify({ status: "WAITING_FIN_MONTH", type: "PAYMENT_NOTICE" }), 600);
      replyLineMessage(replyToken, "📄 [產生繳費單]\n請輸入產製月份？\n\n(例如輸入「" + nowStr + "」，或直接輸入「02」)");
      return; 
  }

  // 4. [其他指令]
  
  if (userMsg === "查ID") {
      replyLineMessage(replyToken, "👤 您的 User ID：\n" + userId + "\n\n(若要查別人，請輸入「查ID @對方」)");
      return;
  }
  
  if (userMsg === "開領據") { replyLineMessage(replyToken, "🧾 格式：開領據 YYYY/MM 講師名\n範例：開領據 2026/02 陳怡均"); return; }
  if (userMsg === "寄領據") { replyLineMessage(replyToken, "📧 格式：寄領據 YYYY/MM\n範例：寄領據 2026/02"); return; }
  
  if (userMsg === "開收據") { replyLineMessage(replyToken, "🧾 格式：開收據 YYYY/MM 學生名\n範例：開收據 2026/02 均逸"); return; }
  if (userMsg === "寄收據") { replyLineMessage(replyToken, "📨 格式：寄收據 YYYY/MM 學生名\n範例：寄收據 2026/02 均逸"); return; }

  if (userMsg.indexOf("開一般收據") === 0) { handleGenReceiptCommand(event, userMsg); return; }
  if (userMsg.indexOf("寄一般收據") === 0) { handleSendGenReceiptCommand(event, userMsg); return; }

  if (userMsg.indexOf("捐款申報") === 0) { handleDonationTaxExportCommand(event, userMsg); return; }

  // [講師功能]
  if (userMsg.indexOf("登錄") === 0 || userMsg.indexOf("登記") === 0) { handleStudentMenu(event, userId, false, 'record'); return; }
  if (userMsg.indexOf("預排") === 0) { handleStudentMenu(event, userId, false, 'plan'); return; }
  if (userMsg.indexOf("核銷") === 0 || userMsg.indexOf("結算") === 0) { handleVerifyMenu(event, userId); return; }
  if (userMsg.indexOf("刪除") === 0) { processDeleteRequest(event); return; } 

  // [行政代操作]
  if (userMsg.indexOf("代登") === 0) {
    const targetName = userMsg.replace("代登選單", "").replace("代登", "").trim();
    if (targetName === "") { handleAdminMenu(event, "record"); } else { handleAdminDirectSearch(event, targetName, "record"); } return;
  }
  if (userMsg.indexOf("代排") === 0) {
    const targetName = userMsg.replace("代排選單", "").replace("代排", "").trim();
    if (targetName === "") { handleAdminMenu(event, "plan"); } else { handleAdminDirectSearch(event, targetName, "plan"); } return;
  }
  if (userMsg.indexOf("代核") === 0) {
    const targetName = userMsg.replace("代核選單", "").replace("代核", "").trim();
    if (targetName === "") { handleAdminMenu(event, "verify"); } else { handleAdminDirectSearch(event, targetName, "verify"); } return;
  }

  // [帶參數的財務指令]
  if (userMsg.indexOf("學費試算") === 0) { handleTuitionCalculation(event, userMsg); return; }
  if (userMsg.indexOf("鐘點試算") === 0) { handleSalaryCalculation(event, userMsg); return; }

  if (userMsg.indexOf("年度扣繳") === 0) { handleAnnualTaxSummaryCommand(event, userMsg); return; }
  if (userMsg.indexOf("免稅試算") === 0) { handleTaxExemptionDashboardCommand(event, userMsg); return; }

  if (userMsg.indexOf("記帳") === 0) { handleManualJournalEntryCommand(event, userMsg); return; }

  // ─── 財務三表報表引擎 ───
  if (userMsg.indexOf("損益表") === 0) { handleFinancialReportCommand(event, userMsg, "IS"); return; }
  if (userMsg.indexOf("資產負債表") === 0) { handleFinancialReportCommand(event, userMsg, "BS"); return; }
  if (userMsg.indexOf("現金流量表") === 0) { handleFinancialReportCommand(event, userMsg, "CF"); return; }
  
  if (userMsg.indexOf("產生繳費單") === 0) { handlePaymentDocCommand(event, userMsg); return; }
  if (userMsg.indexOf("查詢繳費單") === 0) { handlePaymentQueryCommand(event, userMsg); return; }
  if (userMsg.indexOf("帳務補救") === 0) { handleTuitionAdjustmentCommand(event, userMsg); return; }
  
  if (userMsg.indexOf("開收據") === 0) { handleReceiptCommand(event, userMsg); return; } 
  if (userMsg.indexOf("寄收據") === 0) { handleSingleSendReceiptCommand(event, userMsg); return; } 

  if (userMsg.indexOf("開領據") === 0) { handleSingleAllowanceCommand(event, userMsg); return; }
  if (userMsg.indexOf("寄領據") === 0) { handleBatchSendAllowanceEmailCommand(event, userMsg); return; }

  // [其他]
  if (userMsg === "關鍵字") { handleKeywordQuery(event); return; }
  if (userMsg === "請假") { handleLeaveMenu(event); return; }
  if (userMsg === "查群ID") { const gid = event.source.groupId; replyLineMessage(replyToken, gid ? "群ID: " + gid : "非群組"); return; }
  if (userMsg.indexOf("查ID") > -1 && event.message.mention) { handleMentionQuery(event); return; }
  if (userMsg.indexOf("發公告") === 0) { processBroadcast(event, userMsg, false); return; }
  if (userMsg.indexOf("群組公告") === 0) { processBroadcast(event, userMsg, true); return; }
}

function handlePostback(event: any) {
  const data = event.postback.data;
  const replyToken = event.replyToken;
  const userId = event.source.userId;

  if (data === "action=cancel") { replyLineMessage(replyToken, "已取消動作。"); } 
  else if (data.indexOf("action=real_delete") === 0) { executeRealDelete(userId, replyToken, data); }
  else if (data.indexOf("action=pick_stu") === 0) { handleStudentPick(event, data, userId); }
  else if (data.indexOf("action=plan_stu") === 0) { handleStudentPick(event, data, userId); }
  else if (data.indexOf("action=admin_pick_teacher") === 0) { handleAdminPickTeacher(event, data); }
  else if (data.indexOf("action=admin_pick_stu") === 0) { handleAdminPickStudent(event, data); }
  else if (data.indexOf("action=verify_yes") === 0) { executeVerify(event, data, true); } 
  else if (data.indexOf("action=verify_no") === 0) { executeVerify(event, data, false); }
  
  else if (data.indexOf("action=fin_confirm") === 0) { executeFinancialSave(event, data); }
  else if (data.indexOf("action=fin_cancel") === 0) { CacheService.getScriptCache().remove("FIN_" + userId); replyLineMessage(replyToken, "👌 已取消，未寫入任何資料。"); }
  else if (data.indexOf("action=fin_report") === 0) { CacheService.getScriptCache().remove("FIN_" + userId); replyLineMessage(replyToken, "⚠️ 請至紀錄表修正資料後，再重新試算。"); }
  else if (data.indexOf("action=leave_pick") === 0) { handleLeavePick(event, data); }
  
  else if (data.indexOf("action=rcpt_method") === 0) { handleReceiptMethod(event, data); }
  else if (data.indexOf("action=rcpt_cat") === 0) { handleReceiptCategory(event, data); }
  else if (data.indexOf("action=rcpt_final_confirm") === 0) { executeReceiptGeneration(event, data); }
  else if (data.indexOf("action=rcpt_send_email") === 0) { executeReceiptSend(event, data); }
  else if (data.indexOf("action=rcpt_save_only") === 0) { executeReceiptSaveOnly(event, data); }
  else if (data.indexOf("action=allowance_save_only") === 0) { executeAllowanceSaveOnly(event, data); }

  // ★ 新增：單據專用的「取消並銷毀檔案」路由 ★
  else if (data.indexOf("action=doc_cancel_allowance") === 0) { executeDocCancel(event, "ALLOW_"); }
  else if (data.indexOf("action=doc_cancel_rcpt") === 0) { executeDocCancel(event, "RCPT_"); }
  else if (data.indexOf("action=doc_cancel_gen_rcpt") === 0) { executeDocCancel(event, "GEN_RCPT_"); }

  // 記帳 postback
  else if (data.indexOf("action=journal_save_only") === 0) { executeManualJournalSave(event, data); }
  else if (data.indexOf("action=doc_cancel_journal") === 0) { executeDocCancel(event, "JOURNAL_"); }

  // 一般收據 postback
  else if (data.indexOf("action=gen_rcpt_cat") === 0) { handleGenReceiptCategory(event, data); }
  else if (data.indexOf("action=gen_rcpt_method") === 0) { handleGenReceiptMethod(event, data); }
  else if (data.indexOf("action=gen_rcpt_save") === 0) { executeGenReceiptSave(event, data); }
}

// ★ 輔助函式：讓使用者輸入 "02" 也能自動轉成 "2026/02"
function normalizeMonthInput(input: string) {
  const d = new Date();
  const currentYear = d.getFullYear();
  const str = input.trim().replace(/[\/\-\.]/g, ""); // 去除符號

  // 1. 輸入 202602 -> 2026/02
  if (str.length === 6) {
     return str.substring(0, 4) + "/" + str.substring(4, 6);
  }
  // 2. 輸入 02 或 2 -> 2026/02
  if (str.length <= 2) {
     const m = parseInt(str);
     if (m >= 1 && m <= 12) {
       const mm = (m < 10 ? "0" + m : m);
       return currentYear + "/" + mm;
     }
  }
  // 3. 輸入 2026/02 格式 (原樣返回)
  if (input.match(/^\d{4}\/\d{2}$/)) {
     return input;
  }
  return null;
}

// 宣告在其他檔案中定義的外部全域變數/函式，避免 TS 強型別報錯
declare function replyMainMenu(token: string): void;
declare function replyTaxMenu(token: string): void;
declare function replyLineMessage(token: string, msg: string): void;
declare function handleTuitionCalculation(event: any, command: string): void;
declare function handleSalaryCalculation(event: any, command: string): void;
declare function handlePaymentDocCommand(event: any, command: string): void;
declare function handleReceiptDateInput(event: any, msg: string, state: any): void;
declare function processLeaveInput(event: any, msg: string, state: any): void;
declare function processFlowReport(event: any, msg: string, cachedState: string): void;
declare function handleGenReceiptCommand(event: any, cmd: string): void;
declare function handleSendGenReceiptCommand(event: any, cmd: string): void;
declare function handleDonationTaxExportCommand(event: any, cmd: string): void;
declare function handleStudentMenu(event: any, uid: string, isAdm: boolean, mode: string): void;
declare function handleVerifyMenu(event: any, uid: string): void;
declare function processDeleteRequest(event: any): void;
declare function handleAdminMenu(event: any, mode: string): void;
declare function handleAdminDirectSearch(event: any, name: string, mode: string): void;
declare function handleAnnualTaxSummaryCommand(event: any, cmd: string): void;
declare function handleTaxExemptionDashboardCommand(event: any, cmd: string): void;
declare function handleManualJournalEntryCommand(event: any, cmd: string): void;
declare function handleFinancialReportCommand(event: any, cmd: string, type: string): void;
declare function handlePaymentQueryCommand(event: any, cmd: string): void;
declare function handleTuitionAdjustmentCommand(event: any, cmd: string): void;
declare function handleReceiptCommand(event: any, cmd: string): void;
declare function handleSingleSendReceiptCommand(event: any, cmd: string): void;
declare function handleSingleAllowanceCommand(event: any, cmd: string): void;
declare function handleBatchSendAllowanceEmailCommand(event: any, cmd: string): void;
declare function handleKeywordQuery(event: any): void;
declare function handleLeaveMenu(event: any): void;
declare function handleMentionQuery(event: any): void;
declare function processBroadcast(event: any, cmd: string, isGroup: boolean): void;
declare function executeRealDelete(uid: string, token: string, data: string): void;
declare function handleStudentPick(event: any, data: string, uid: string): void;
declare function handleAdminPickTeacher(event: any, data: string): void;
declare function handleAdminPickStudent(event: any, data: string): void;
declare function executeVerify(event: any, data: string, isOk: boolean): void;
declare function executeFinancialSave(event: any, data: string): void;
declare function handleLeavePick(event: any, data: string): void;
declare function handleReceiptMethod(event: any, data: string): void;
declare function handleReceiptCategory(event: any, data: string): void;
declare function executeReceiptGeneration(event: any, data: string): void;
declare function executeReceiptSend(event: any, data: string): void;
declare function executeReceiptSaveOnly(event: any, data: string): void;
declare function executeAllowanceSaveOnly(event: any, data: string): void;
declare function executeDocCancel(event: any, prefix: string): void;
declare function executeManualJournalSave(event: any, data: string): void;
declare function handleGenReceiptCategory(event: any, data: string): void;
declare function handleGenReceiptMethod(event: any, data: string): void;
declare function executeGenReceiptSave(event: any, data: string): void;

// LIFF 專屬 API 控制器介面宣告
declare function handleLiffMe(userId: string): any;
declare function handleLiffFormOptions(lineUserId?: string): any;
declare function handleLiffRegister(params: any): any;
declare function handleLiffGetUnverified(userId: string): any;
declare function handleLiffVerifySchedule(params: any): any;
declare function handleLiffLeave(params: any): any;
declare function handleLiffAdminTask(params: any): any;
declare function handleLiffUnbind(userId: string): any;
declare function handleLiffVerifyAndBind(name: string, userId: string): any;

