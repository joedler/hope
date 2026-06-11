// ==========================================
// 💰 Finance_Service.ts : 財務計算、存檔與稅務引擎 (TypeScript + LIFF 整合重構版)
// ==========================================

// ==========================================
// 1. LIFF 專屬行政管理 API 後端實作
// ==========================================

function handleLiffAdminTask(params: any) {
  const lineUserId = params.lineUserId;
  const task = params.task; // 行政作業名稱，由 LIFF 行政端傳入
  const month = params.month; // YYYY/MM (例如 "2026/05")

  // 1. 驗證管理員權限
  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) {
    return { ok: false, message: "❌ 權限不足：限行政人員使用。" };
  }

  // 2. 模擬 event 對象
  const mockEvent = {
    replyToken: "LIFF_API_CALL", // 攔截 reply，不向 LINE 伺服器發送
    source: {
      userId: lineUserId
    }
  };

  try {
    if (task === "學費試算" || task === "学费试算") {
      return { ok: false, message: "學費試算的 LIFF 預覽/確認流程尚未開放；為避免直接寫入，目前已暫停此按鈕。" };
    } 
    
    if (task === "鐘點試算" || task === "钟点试算") {
      return { ok: false, message: "鐘點試算的 LIFF 預覽/確認流程尚未開放；為避免直接寫入，目前已暫停此按鈕。" };
    } 
    
    if (task === "產生繳費單" || task === "产生缴费单") {
      handlePaymentDocCommand(mockEvent, "產生繳費單 " + month);
      return { ok: true, message: `✅ ${month} 學生繳費單已批次生成至 Google 雲端硬碟！` };
    } 
    
    if (task === "寄送領據" || task === "寄領據" || task === "寄领据") {
      handleBatchSendAllowanceEmailCommand(mockEvent, "寄領據 " + month);
      return { ok: true, message: `✅ ${month} 講師領據 Email 已發送完畢！` };
    }

    if (task === "產生收據" || task === "寄送收據") {
      return { ok: false, message: `${task} 的 LIFF 流程尚未開放；請先維持原 LINE 對話流程，避免誤寄或重複開立。` };
    }

    if (task === "產生領據") {
      return { ok: false, message: "產生領據目前採單一講師對話流程；請在 LINE 輸入：開領據 YYYY/MM 講師名。" };
    }

    if (task === "一般收據") {
      return { ok: false, message: "一般收據目前採 LINE 對話流程；請輸入：開一般收據 YYYY/MM/DD 姓名 金額，或：寄一般收據 YYYY/MM 姓名。" };
    }

    if (task === "稅務專區") {
      return { ok: false, message: "稅務專區尚在規劃中，後續會整合年度收據、捐款/會費分類與稅務檢核。" };
    }

    return { ok: false, message: `不支援的行政任務: ${task}` };
  } catch (e) {
    return { ok: false, message: "執行失敗：" + e.toString() };
  }
}

function handleLiffAdminPreview(params: any) {
  const lineUserId = params.lineUserId;
  const feature = String(params.feature || "").trim();
  const month = normalizeAdminPreviewMonth(params.month);

  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) {
    return { ok: false, message: "❌ 權限不足：限行政人員使用。" };
  }

  const previewMap: any = {
    "學費試算": {
      summary: `預覽 ${month} 學費試算。`,
      items: ["檢查學生姓名與課程", "檢查預收/後收", "檢查補收/退費", "確認後才可寫入試算結果"],
      nextAction: "確認寫入試算結果（尚未開放）"
    },
    "鐘點試算": {
      summary: `預覽 ${month} 講師鐘點試算。`,
      items: ["檢查授課紀錄", "檢查預排核銷", "檢查補充鐘點與保費", "確認後才可寫入試算結果"],
      nextAction: "確認寫入鐘點試算（尚未開放）"
    },
    "繳費單": {
      summary: `預覽 ${month} 繳費單。`,
      items: ["檢查學生、課程與金額", "檢查補收/退費是否併入", "確認後才產生單據", "寄送 Email 或 LINE push 前需再次確認收件人"],
      nextAction: "確認產生繳費單（尚未開放）"
    },
    "收據": {
      summary: `預覽 ${month} 學費收據。`,
      items: ["檢查繳費狀態", "檢查收據金額與抬頭", "確認後才產生收據", "寄送前需確認家長收件資訊"],
      nextAction: "確認產生收據（尚未開放）"
    },
    "一般收據": {
      summary: "預覽入會費、常年會費或捐款一般收據。",
      items: ["選擇收據類型", "檢查收款人與金額", "確認後才產生或寄送", "後續需接正式收件人資料"],
      nextAction: "確認產生一般收據（尚未開放）"
    },
    "領據": {
      summary: `預覽 ${month} 講師領據。`,
      items: ["檢查講師與鐘點費", "檢查授課紀錄與補充項目", "確認後才產生領據", "LINE push 對象為講師本人"],
      nextAction: "確認產生領據（尚未開放）"
    },
    "稅務專區": {
      summary: `預覽 ${month} 稅務資料整理。`,
      items: ["年度資料彙整", "收據與捐款資料檢核", "後續需定義稅務輸出格式"],
      nextAction: "確認產生稅務資料（尚未開放）"
    }
  };

  const preview = feature === "學費試算" ? buildTuitionAdminPreview(month) :
    feature === "鐘點試算" ? buildSalaryAdminPreview(month) :
    feature === "繳費單" ? buildPaymentNoticeAdminPreview(month) :
    feature === "收據" ? buildReceiptAdminPreview(month) :
    feature === "領據" ? buildAllowanceAdminPreview(month) :
    feature === "一般收據" ? buildGeneralReceiptAdminPreview(month) :
    previewMap[feature];
  if (!preview) {
    return { ok: false, message: `不支援的行政預覽功能: ${feature}` };
  }

  const metrics = buildAdminPreviewMetrics(month, feature);

  return {
    ok: true,
    preview: {
      title: feature,
      month,
      summary: preview.summary,
      items: preview.items,
      rows: preview.rows || [],
      metrics,
      status: "後端只讀預覽，尚未寫入或寄送",
      nextAction: preview.nextAction,
      canConfirm: preview.canConfirm === true,
      confirmAction: preview.confirmAction || ""
    }
  };
}

function handleLiffAdminConfirmSettlement(params: any) {
  const lineUserId = String(params.lineUserId || "").trim();
  const feature = String(params.feature || "").trim();
  const month = normalizeAdminPreviewMonth(params.month);

  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) {
    return { ok: false, message: "❌ 權限不足：限行政人員使用。" };
  }
  if (feature !== "學費試算" && feature !== "鐘點試算") {
    return { ok: false, message: "目前只開放學費試算與鐘點試算確認寫入。" };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const targetSheetName = feature === "學費試算" ? SHEET_NAME_FIN_FEE : SHEET_NAME_FIN_PAY;
  if (feature === "學費試算") {
    const tuitionPreview: any = buildTuitionReadOnlyPreview(month);
    const selectedIds = parseAdminSelectedIds(params.selectedIds);
    const selectableRows = (tuitionPreview.rows || []).filter(function(row: any) { return row.selectable; });
    const selectedRows = selectedIds.length > 0
      ? selectableRows.filter(function(row: any) { return selectedIds.indexOf(row.id) > -1; })
      : selectableRows;
    if (selectedRows.length === 0) {
      return { ok: false, message: `${month} 沒有選取可寫入的學費試算項目；待核銷預排不可寫入。` };
    }
  } else {
    const existingSalaryCount = countSheetRowsByMonthOnly(ss, SHEET_NAME_FIN_PAY, 0, month);
    if (existingSalaryCount > 0) {
      return { ok: false, message: `${month} 已有 ${existingSalaryCount} 筆鐘點結算資料，系統已阻擋重複寫入；若需更正請走作廢、重新產生或補發流程。` };
    }
  }
  const beforeCount = countSheetRowsByMonthOnly(ss, targetSheetName, 0, month);
  const mockEvent = {
    replyToken: "LIFF_API_CALL",
    source: { userId: lineUserId }
  };
  const cacheKey = "FIN_" + lineUserId;
  CacheService.getScriptCache().remove(cacheKey);

  if (feature === "學費試算") {
    handleTuitionCalculation(mockEvent, "學費試算 " + month);
  } else {
    handleSalaryCalculation(mockEvent, "鐘點試算 " + month);
  }

  const cacheDataStr = CacheService.getScriptCache().get(cacheKey);
  if (!cacheDataStr) {
    return { ok: false, message: `${month} 目前沒有可寫入的${feature}資料，或試算資料未能建立。` };
  }
  const cacheObj = JSON.parse(cacheDataStr);
  if (feature === "學費試算") {
    const selectedIds = parseAdminSelectedIds(params.selectedIds);
    if (selectedIds.length > 0) {
      cacheObj.selectedTuitionKeys = selectedIds.map(function(id: string) { return id.replace(/^tuition:/, ""); });
      cacheObj.save = (cacheObj.save || []).filter(function(row: any[]) {
        return cacheObj.selectedTuitionKeys.indexOf(buildTuitionSelectionKey(row[1], row[2])) > -1;
      });
      if (!cacheObj.save || cacheObj.save.length === 0) {
        CacheService.getScriptCache().remove(cacheKey);
        return { ok: false, message: `${month} 勾選項目沒有可寫入的學費結算資料；待核銷預排不可寫入。` };
      }
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(cacheObj), 600);
    }
  }
  const saveCount = cacheObj.save ? cacheObj.save.length : 0;
  executeFinancialSave(mockEvent, "");
  const afterCount = countSheetRowsByMonthOnly(ss, targetSheetName, 0, month);
  const writtenCount = Math.max(0, afterCount - beforeCount);
  if (writtenCount <= 0) {
    CacheService.getScriptCache().remove(cacheKey);
    return {
      ok: false,
      message: `${month} ${feature}未寫入新資料。可能已有結算紀錄，或被重複寫入防呆擋下；請檢查結算表。`
    };
  }

  const preview = feature === "學費試算" ? buildPaymentNoticeAdminPreview(month) : buildAllowanceAdminPreview(month);
  return {
    ok: true,
    message: `${month} ${feature}已確認寫入，共新增 ${writtenCount} 筆結算資料。`,
    writtenCount,
    plannedCount: saveCount,
    preview
  };
}

function handleLiffAdminConfirmDocument(params: any) {
  const lineUserId = String(params.lineUserId || "").trim();
  const feature = String(params.feature || "").trim();
  const month = normalizeAdminPreviewMonth(params.month);

  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;
  if (!isAdmin) {
    return { ok: false, message: "❌ 權限不足：限行政人員使用。" };
  }

  if (feature !== "繳費單" && feature !== "收據" && feature !== "領據") {
    return { ok: false, message: "目前只開放繳費單、收據與領據確認產生；寄送流程仍需各自預覽確認後再開放。" };
  }

  if (feature === "收據") {
    const beforePreview = buildReceiptReadOnlyPreview(month);
    if (beforePreview.studentCount <= 0) {
      return { ok: false, message: `${month} 沒有可產生收據的學費結算資料，請先完成學費試算與繳費單流程。` };
    }
    if (beforePreview.generatedCount >= beforePreview.studentCount) {
      return { ok: false, message: `${month} 收據已全部產生，未重複產生 PDF。` };
    }

    const selectedIds = parseAdminSelectedIds(params.selectedIds);
    const selectableRows = (beforePreview.rows || []).filter(function(row: any) { return row.selectable; });
    const selectedRows = selectedIds.length > 0
      ? selectableRows.filter(function(row: any) { return selectedIds.indexOf(row.id) > -1; })
      : selectableRows;
    if (selectedRows.length === 0) {
      return { ok: false, message: `${month} 沒有選取可產生收據的學生，請先確認收款方式、收據類別與收款日期。` };
    }

    const resultMessages: string[] = [];
    for (let i = 0; i < selectedRows.length; i++) {
      resultMessages.push(createReceiptDocumentsBatch(month, selectedRows[i].name));
    }
    const resultMessage = resultMessages.join("\n\n");
    const afterPreview = buildReceiptAdminPreview(month);
    if (resultMessage.indexOf("❌") === 0 || resultMessage.indexOf("⚠️") === 0) {
      return { ok: false, message: resultMessage, preview: afterPreview };
    }

    return {
      ok: true,
      message: `${month} 收據已確認產生，共 ${selectedRows.length} 位學生。\n${resultMessage}`,
      preview: afterPreview
    };
  }

  if (feature === "領據") {
    const beforePreview = buildAllowanceReadOnlyPreview(month);
    if (beforePreview.teacherCount <= 0) {
      return { ok: false, message: `${month} 沒有可產生領據的鐘點結算資料，請先完成鐘點試算確認寫入。` };
    }
    if (beforePreview.generatedCount >= beforePreview.teacherCount) {
      return { ok: false, message: `${month} 領據已全部產生，未重複產生 PDF。` };
    }

    const selectedIds = parseAdminSelectedIds(params.selectedIds);
    const selectableRows = (beforePreview.rows || []).filter(function(row: any) { return row.selectable; });
    const selectedRows = selectedIds.length > 0
      ? selectableRows.filter(function(row: any) { return selectedIds.indexOf(row.id) > -1; })
      : selectableRows;
    if (selectedRows.length === 0) {
      return { ok: false, message: `${month} 沒有選取可產生領據的講師。` };
    }

    const resultMessages: string[] = [];
    for (let i = 0; i < selectedRows.length; i++) {
      resultMessages.push(createAllowanceDocumentsBatch(month, selectedRows[i].name));
    }
    const resultMessage = resultMessages.join("\n\n");
    const afterPreview = buildAllowanceAdminPreview(month);
    if (resultMessage.indexOf("❌") === 0 || resultMessage.indexOf("⚠️") === 0) {
      return { ok: false, message: resultMessage, preview: afterPreview };
    }

    return {
      ok: true,
      message: `${month} 領據已確認產生，共 ${selectedRows.length} 位講師。\n${resultMessage}`,
      preview: afterPreview
    };
  }

  const beforePreview = buildPaymentNoticeReadOnlyPreview(month);
  if (beforePreview.studentCount <= 0) {
    return { ok: false, message: `${month} 沒有可產生繳費單的學費結算資料，請先完成學費試算確認寫入。` };
  }
  if (beforePreview.generatedCount >= beforePreview.studentCount) {
    return { ok: false, message: `${month} 繳費單已全部產生，未重複產生 PDF。` };
  }

  const selectedIds = parseAdminSelectedIds(params.selectedIds);
  const selectableRows = (beforePreview.rows || []).filter(function(row: any) { return row.selectable; });
  const selectedRows = selectedIds.length > 0
    ? selectableRows.filter(function(row: any) { return selectedIds.indexOf(row.id) > -1; })
    : selectableRows;
  if (selectedRows.length === 0) {
    return { ok: false, message: `${month} 沒有選取可產生繳費單的學生。` };
  }

  const resultMessages: string[] = [];
  for (let i = 0; i < selectedRows.length; i++) {
    resultMessages.push(createPaymentNoticesBatch(month, selectedRows[i].name));
  }
  const resultMessage = resultMessages.join("\n\n");
  const afterPreview = buildPaymentNoticeAdminPreview(month);
  if (resultMessage.indexOf("❌") === 0 || resultMessage.indexOf("⚠️") === 0) {
    return { ok: false, message: resultMessage, preview: afterPreview };
  }

  return {
    ok: true,
    message: `${month} 繳費單已確認產生，共 ${selectedRows.length} 位學生。\n${resultMessage}`,
    preview: afterPreview
  };
}

function parseAdminSelectedIds(value: any): string[] {
  const raw = String(value || "").trim();
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (Object.prototype.toString.call(parsed) === "[object Array]") {
      return parsed.map(function(item: any) { return String(item || "").trim(); }).filter(function(item: string) { return item !== ""; });
    }
  } catch (e) {
    // Fall back to comma separated values.
  }
  return raw.split(",").map(function(item) { return String(item || "").trim(); }).filter(function(item) { return item !== ""; });
}

function normalizeAdminPreviewMonth(value: any) {
  const raw = String(value || "").trim();
  if (/^\d{4}\/\d{2}$/.test(raw)) return raw;
  if (/^\d{4}-\d{2}$/.test(raw)) return raw.replace("-", "/");
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM");
}

function buildTuitionAdminPreview(month: string) {
  try {
    const result = buildTuitionReadOnlyPreview(month);
    const items = result.items.slice(0, 12);
    if (result.items.length > items.length) {
      items.push(`另有 ${result.items.length - items.length} 筆課程彙總未列出，正式預覽頁後續再提供完整清單。`);
    }
    if ((result as any).isExistingSettlement) {
      return {
        summary: `${month} 學費既有結算摘要：${result.studentCount} 位學生，既有結算總額 ${formatCurrency(result.grandTotal)}。`,
        items,
        nextAction: "查看繳費單 / 已產生單據",
        canConfirm: false
      };
    }
    const pendingCount = (result as any).pendingCount || 0;
    const rows = (result as any).rows || [];
    return {
      summary: `${month} 學費試算只讀預覽：${result.studentCount} 位學生，預估總額 ${formatCurrency(result.grandTotal)}。${pendingCount > 0 ? " 尚有待核銷預排，需先核銷後才能正式寫入。" : ""}`,
      items,
      rows,
      nextAction: rows.some(function(row: any) { return row.selectable; }) ? "確認寫入已勾選項目" : "目前沒有可寫入項目",
      canConfirm: rows.some(function(row: any) { return row.selectable; }),
      confirmAction: "adminConfirmSettlement"
    };
  } catch (e) {
    return {
      summary: `${month} 學費試算只讀預覽讀取失敗。`,
      items: ["請先檢查 Google Sheets 分頁、欄位與 GAS 權限。", "錯誤：" + e.toString()],
      nextAction: "確認寫入試算結果（尚未開放）",
      canConfirm: false
    };
  }
}

function buildTuitionReadOnlyPreview(month: string) {
  const timeZone = Session.getScriptTimeZone();
  const splitDt = month.split("/");
  const nextDate = new Date(parseInt(splitDt[0], 10), parseInt(splitDt[1], 10), 1);
  const nextMonthStr = Utilities.formatDate(nextDate, timeZone, "yyyy/MM");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  if (!recordSheet || !planSheet || !courseSheet) throw new Error("找不到授課紀錄、預排紀錄或課程設定表。");

  const existingSettlementCount = countSheetRowsByMonthOnly(ss, SHEET_NAME_FIN_FEE, 0, month);
  if (existingSettlementCount > 0) {
    return buildExistingTuitionSettlementPreview(ss, month, existingSettlementCount);
  }

  const configMap: any = {};
  const courseData = courseSheet.getDataRange().getValues();
  for (let i = 1; i < courseData.length; i++) {
    const studentName = String(courseData[i][2] || "").trim();
    const courseName = String(courseData[i][3] || "").trim();
    if (!studentName || !courseName) continue;
    configMap[studentName + "_" + courseName] = {
      fee: parseFloat(courseData[i][4]) || 0,
      mode: String(courseData[i][6] || "").trim() === "預收" ? "預收" : "後收",
      teacher: String(courseData[i][1] || "").trim()
    };
  }

  const stats: any = {};
  function initStats(studentName: string, courseName: string, teacherName: string, fee: number, mode: string) {
    if (!stats[studentName]) stats[studentName] = {};
    if (!stats[studentName][courseName]) {
      stats[studentName][courseName] = {
        teacher: teacherName,
        fee,
        mode,
        recordBase: 0,
        planBase: 0,
        pendingPlanBase: 0,
        planNext: 0,
        detailsRec: [],
        detailsPending: [],
        detailsPlan: [],
        adjustments: []
      };
    }
  }

  const recordData = recordSheet.getDataRange().getValues();
  for (let i = 1; i < recordData.length; i++) {
    const rowMonth = normalizeFinancialMonth(recordData[i][2], timeZone);
    const settled = recordData[i][9];
    if (rowMonth !== month || settled) continue;
    const studentName = String(recordData[i][7] || "").trim();
    const courseName = String(recordData[i][8] || "").trim();
    const hours = parseFloat(recordData[i][5]) || 0;
    const conf = configMap[studentName + "_" + courseName] || { fee: 0, mode: "後收", teacher: recordData[i][1] };
    initStats(studentName, courseName, conf.teacher, conf.fee, conf.mode);
    stats[studentName][courseName].recordBase += hours;
    const perLessonAmt = Math.round(hours * conf.fee);
    stats[studentName][courseName].detailsRec.push("[實上] " + formatSheetMonthDay(recordData[i][2], timeZone) + " " + recordData[i][3] + "-" + recordData[i][4] + " (" + formatCurrency(perLessonAmt) + ")");
  }

  const planData = planSheet.getDataRange().getValues();
  for (let i = 1; i < planData.length; i++) {
    const lessonDateMonth = normalizeFinancialMonth(planData[i][2], timeZone);
    const refundSettledMonth = normalizeFinancialMonth(planData[i][11], timeZone);
    const status = String(planData[i][9] || "").trim();
    const targetMonthForPlanBase = refundSettledMonth || lessonDateMonth;
    if (status === "取消" && !refundSettledMonth) continue;

    const studentName = String(planData[i][7] || "").trim();
    const courseName = String(planData[i][8] || "").trim();
    const hours = parseFloat(planData[i][5]) || 0;
    const conf = configMap[studentName + "_" + courseName];
    if (!conf || conf.mode !== "預收") continue;
    initStats(studentName, courseName, conf.teacher, conf.fee, conf.mode);

    if (targetMonthForPlanBase === month) {
      if (status === "未核銷") {
        stats[studentName][courseName].pendingPlanBase += hours;
        const perLessonAmt = Math.round(hours * conf.fee);
        stats[studentName][courseName].detailsPending.push("[待核銷預排] " + formatSheetMonthDay(planData[i][2], timeZone) + " " + planData[i][3] + "-" + planData[i][4] + " (" + formatCurrency(perLessonAmt) + ")");
      } else {
        stats[studentName][courseName].planBase += hours;
      }
      if (status === "取消") {
        stats[studentName][courseName].detailsRec.push("[歷史取消退費] " + formatSheetMonthDay(planData[i][2], timeZone) + " " + planData[i][3] + "-" + planData[i][4]);
      }
    } else if (lessonDateMonth === nextMonthStr && status !== "取消") {
      stats[studentName][courseName].planNext += hours;
      const perLessonAmt = Math.round(hours * conf.fee);
      stats[studentName][courseName].detailsPlan.push("[下月預收] " + formatSheetMonthDay(planData[i][2], timeZone) + " " + planData[i][3] + "-" + planData[i][4] + " (" + formatCurrency(perLessonAmt) + ")");
    }
  }

  appendTuitionAdjustmentsToStats(ss, stats, configMap, month);

  const items: string[] = [];
  const rows: any[] = [];

  let grandTotal = 0;
  let studentCount = 0;
  let pendingCount = 0;
  for (const studentName in stats) {
    let studentTotal = 0;
    const courseSummaries: string[] = [];
    for (const courseName in stats[studentName]) {
      const item = stats[studentName][courseName];
      const adjustmentTotal = item.adjustments.reduce(function(sum: number, adj: any) { return sum + adj.amount; }, 0);
      let courseTotal = 0;
      let formula = "";
      if (item.mode === "預收") {
        const diff = item.pendingPlanBase > 0 ? 0 : Math.round((item.recordBase - item.planBase) * 10) / 10;
        const totalHours = Math.round((item.planNext + diff) * 10) / 10;
        courseTotal = Math.round(totalHours * item.fee) + adjustmentTotal;
        formula = item.pendingPlanBase > 0
          ? `預收 ${item.planNext}hr，尚有待核銷 ${item.pendingPlanBase}hr，暫不計入退費，調整 ${formatCurrency(adjustmentTotal)}`
          : `預收 ${item.planNext}hr，核對差異 ${diff}hr，調整 ${formatCurrency(adjustmentTotal)}`;
      } else {
        courseTotal = Math.round(item.recordBase * item.fee) + adjustmentTotal;
        formula = `後收實上 ${item.recordBase}hr，調整 ${formatCurrency(adjustmentTotal)}`;
      }
      if (courseTotal !== 0 || item.recordBase > 0 || item.planNext > 0 || item.pendingPlanBase > 0 || adjustmentTotal !== 0) {
        studentTotal += courseTotal;
        const detailParts: string[] = [];
        for (let d = 0; d < item.detailsRec.length; d++) detailParts.push(item.detailsRec[d]);
        for (let d = 0; d < item.detailsPending.length; d++) detailParts.push(item.detailsPending[d]);
        for (let d = 0; d < item.detailsPlan.length; d++) detailParts.push(item.detailsPlan[d]);
        for (let a = 0; a < item.adjustments.length; a++) {
          const adj = item.adjustments[a];
          detailParts.push(
            `[帳務補救-${adj.type}] 原月份 ${adj.originalMonth || "未填"}，` +
            `${adj.lessonDate} ${adj.startTime}-${adj.endTime}，` +
            `${adj.hours}hr x ${formatCurrency(adj.unitFee)}，金額 ${formatCurrency(adj.amount)}，` +
            `原單 ${adj.relatedDocId || "未填"}，狀態 ${adj.status || "未填"}，原因：${adj.reason || "未填"}`
          );
        }
        const detailText = detailParts.length > 0 ? `\n  明細：\n  - ${detailParts.join("\n  - ")}` : "";
        courseSummaries.push(`${courseName}（${item.mode}）\n  ${formula}\n  小計 ${formatCurrency(courseTotal)}${detailText}`);
        const selectable = item.pendingPlanBase <= 0 && courseTotal !== 0;
        rows.push({
          id: "tuition:" + buildTuitionSelectionKey(studentName, courseName),
          type: "tuition",
          name: studentName + " / " + courseName,
          amount: courseTotal,
          amountText: formatCurrency(courseTotal),
          docId: "",
          status: selectable ? "可寫入" : "待核銷預排不可寫入",
          selectable,
          selectedDefault: selectable,
          warnings: item.pendingPlanBase > 0 ? ["尚有待核銷預排，暫不可寫入"] : [],
          details: [item.mode, formula]
        });
        if (item.pendingPlanBase > 0) pendingCount++;
      }
    }
    if (courseSummaries.length > 0) {
      studentCount++;
      grandTotal += studentTotal;
      items.push(`${studentName}\n本期預估：${formatCurrency(studentTotal)}\n` + courseSummaries.join("\n"));
    }
  }

  if (items.length === 0) items.push(`${month} 目前沒有待試算學費資料。`);
  return { items, rows, grandTotal, studentCount, pendingCount };
}

function buildExistingTuitionSettlementPreview(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, month: string, existingSettlementCount: number) {
  const timeZone = Session.getScriptTimeZone();
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_FEE);
  if (!sheet) {
    return {
      items: [`${month} 已有 ${existingSettlementCount} 筆學費結算資料，但找不到學費結算表。`],
      grandTotal: 0,
      studentCount: 0
    };
  }

  const data = sheet.getDataRange().getValues();
  const studentsMap: any = {};
  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][0], timeZone);
    if (rowMonth !== month) continue;
    const studentName = String(data[i][1] || "").trim();
    if (!studentName) continue;
    if (!studentsMap[studentName]) {
      studentsMap[studentName] = {
        name: studentName,
        total: 0,
        docId: "",
        courseAmountSum: 0,
        courses: []
      };
    }
    const item = studentsMap[studentName];
    const courseName = String(data[i][2] || "").trim();
    const mode = String(data[i][3] || "").trim();
    const courseAmount = parseFloat(data[i][7]) || 0;
    if (courseName) {
      item.courses.push(`${courseName}${mode ? "（" + mode + "）" : ""}：${formatCurrency(courseAmount)}`);
    }
    item.courseAmountSum += courseAmount;
    if (data[i][8] !== "" && data[i][8] != null) item.total = parseFloat(data[i][8]) || 0;
    if (data[i][9]) item.docId = String(data[i][9]).trim();
  }

  const items: string[] = [
    `注意：${month} 已有 ${existingSettlementCount} 筆學費結算資料；此處顯示既有結算摘要，不重新試算授課、預排或補救差異。`
  ];
  let grandTotal = 0;
  let studentCount = 0;
  for (const studentName in studentsMap) {
    const item = studentsMap[studentName];
    const total = item.total || item.courseAmountSum;
    grandTotal += total;
    studentCount++;
    const courseText = item.courses.length > 0 ? "\n課程：\n- " + item.courses.join("\n- ") : "";
    items.push(`${studentName}\n金額：${formatCurrency(total)}\n課程數：${item.courses.length}\n單號：${item.docId || "未填"}${courseText}`);
  }

  if (studentCount === 0) {
    items.push(`${month} 找不到可顯示的既有學費結算摘要。`);
  }
  return { items, grandTotal, studentCount, isExistingSettlement: true };
}

function appendTuitionAdjustmentsToStats(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, stats: any, configMap: any, month: string) {
  const sheet = ss.getSheetByName(SHEET_NAME_TUITION_ADJUSTMENT);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const targetMonth = normalizeFinancialMonth(data[i][1], Session.getScriptTimeZone());
    if (targetMonth !== month) continue;
    const status = String(data[i][13] || "").trim();
    if (status === "作廢" || status === "已取消") continue;
    const studentName = String(data[i][2] || "").trim();
    const courseName = String(data[i][3] || "").trim();
    const amount = parseFloat(data[i][9]) || 0;
    const type = String(data[i][10] || "").trim();
    const conf = configMap[studentName + "_" + courseName] || { fee: parseFloat(data[i][8]) || 0, mode: "後收", teacher: "" };
    if (!stats[studentName]) stats[studentName] = {};
    if (!stats[studentName][courseName]) {
      stats[studentName][courseName] = { teacher: conf.teacher, fee: conf.fee, mode: conf.mode, recordBase: 0, planBase: 0, pendingPlanBase: 0, planNext: 0, detailsRec: [], detailsPending: [], detailsPlan: [], adjustments: [] };
    }
    stats[studentName][courseName].adjustments.push({
      amount,
      type,
      status,
      originalMonth: normalizeFinancialMonth(data[i][16], Session.getScriptTimeZone()),
      lessonDate: formatSheetMonthDay(data[i][4], Session.getScriptTimeZone()),
      startTime: formatPreviewTime(data[i][5]),
      endTime: formatPreviewTime(data[i][6]),
      hours: parseFloat(data[i][7]) || 0,
      unitFee: parseFloat(data[i][8]) || 0,
      relatedDocId: String(data[i][11] || "").trim(),
      reason: String(data[i][12] || "").trim()
    });
  }
}

function buildSalaryAdminPreview(month: string) {
  try {
    const result = buildSalaryReadOnlyPreview(month);
    const items = result.items.slice(0, 12);
    if (result.items.length > items.length) {
      items.push(`另有 ${result.items.length - items.length} 位講師彙總未列出，正式預覽頁後續再提供完整清單。`);
    }
    if ((result as any).isExistingSettlement) {
      return {
        summary: `${month} 鐘點既有結算摘要：${result.teacherCount} 位講師，應付總額 ${formatCurrency(result.grossTotal)}，實發總額 ${formatCurrency(result.netTotal)}。`,
        items,
        nextAction: "已寫入鐘點結算，請勿重複試算",
        canConfirm: false
      };
    }
    return {
      summary: `${month} 鐘點試算只讀預覽：${result.teacherCount} 位講師，應付總額 ${formatCurrency(result.grossTotal)}，實發總額 ${formatCurrency(result.netTotal)}。`,
      items,
      nextAction: "確認寫入鐘點試算",
      canConfirm: result.teacherCount > 0,
      confirmAction: "adminConfirmSettlement"
    };
  } catch (e) {
    return {
      summary: `${month} 鐘點試算只讀預覽讀取失敗。`,
      items: ["請先檢查授課紀錄、課程設定表、講師名單與 GAS 權限。", "錯誤：" + e.toString()],
      nextAction: "確認寫入鐘點試算（尚未開放）",
      canConfirm: false
    };
  }
}

function buildSalaryReadOnlyPreview(month: string) {
  const timeZone = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!recordSheet || !courseSheet || !teacherSheet) throw new Error("找不到授課紀錄、課程設定表或講師名單。");

  const courseData = courseSheet.getDataRange().getValues();
  const profitMap: any = {};
  for (let i = 1; i < courseData.length; i++) {
    const studentName = String(courseData[i][2] || "").trim();
    const courseName = String(courseData[i][3] || "").trim();
    const fee = parseFloat(courseData[i][4]) || 0;
    const ratio = parseFloat(courseData[i][5]) || 0;
    const teacher = String(courseData[i][1] || "").trim();
    if (studentName && courseName && fee && ratio) {
      profitMap[studentName + "_" + courseName] = { fee, ratio, teacher };
    }
  }

  const teacherData = teacherSheet.getDataRange().getValues();
  const taxConfigMap: any = {};
  for (let i = 1; i < teacherData.length; i++) {
    const teacherName = String(teacherData[i][0] || "").trim();
    if (!teacherName) continue;
    taxConfigMap[teacherName] = {
      formatCode: teacherData[i][11] || "9B",
      nationality: teacherData[i][12] || "本國人",
      nhiExempt: teacherData[i][13] || "否"
    };
  }

  const salaryStats: any = {};
  const recordData = recordSheet.getDataRange().getValues();
  for (let i = 1; i < recordData.length; i++) {
    const rowMonth = normalizeFinancialMonth(recordData[i][2], timeZone);
    const settled = recordData[i][10];
    const courseName = String(recordData[i][8] || "").trim();
    if (courseName.indexOf("取消") > -1 || rowMonth !== month || settled) continue;

    const teacherName = String(recordData[i][1] || "").trim();
    const studentName = String(recordData[i][7] || "").trim();
    const hours = parseFloat(recordData[i][5]) || 0;
    const conf = profitMap[studentName + "_" + courseName] || { fee: 0, ratio: 0 };
    const payRate = conf.fee * conf.ratio;
    const payAmount = Math.round(hours * payRate);
    appendSalaryPreviewItem(
      salaryStats,
      teacherName,
      studentName,
      courseName,
      formatSheetMonthDay(recordData[i][2], timeZone),
      recordData[i][3],
      recordData[i][4],
      hours,
      payAmount,
      payRate,
      "授課"
    );
  }

  appendSalaryAdjustmentsToStats(ss, salaryStats, profitMap, month, timeZone);

  const items: string[] = [];
  const existingSettlementCount = countSheetRowsByMonthOnly(ss, SHEET_NAME_FIN_PAY, 0, month);
  if (existingSettlementCount > 0) {
    return buildExistingSalarySettlementPreview(ss, month, existingSettlementCount);
  }

  let grossTotal = 0;
  let netTotal = 0;
  let teacherCount = 0;
  for (const teacherName in salaryStats) {
    const item = salaryStats[teacherName];
    const taxConfig = taxConfigMap[teacherName] || { formatCode: "9B", nationality: "本國人", nhiExempt: "否" };
    const taxAndNhi = calculateSalaryDeductions(item.total, taxConfig);
    grossTotal += item.total;
    netTotal += taxAndNhi.netAmount;
    teacherCount++;
    const lines: string[] = [
      `${teacherName}（${taxConfig.formatCode}）`,
      `總時數：${item.hours}hr`,
      `應付：${formatCurrency(item.total)}`,
      `扣繳：${formatCurrency(taxAndNhi.taxAmount)}`,
      `補充保費：${formatCurrency(taxAndNhi.nhiAmount)}`,
      `實發：${formatCurrency(taxAndNhi.netAmount)}`
    ];
    if (item.adjustmentCount > 0) lines.push(`帳務補救補發：${item.adjustmentCount} 筆`);
    if (item.missingRateCount > 0) lines.push(`提醒：${item.missingRateCount} 筆缺少鐘點比例/單價`);
    if (item.details.length > 0) {
      lines.push("明細：");
      for (let d = 0; d < item.details.length; d++) {
        lines.push("- " + item.details[d]);
      }
    }
    items.push(lines.join("\n"));
  }

  if (items.length === 0) items.push(`${month} 目前沒有待試算鐘點資料。`);
  return { items, grossTotal, netTotal, teacherCount };
}

function buildExistingSalarySettlementPreview(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, month: string, existingSettlementCount: number) {
  const timeZone = Session.getScriptTimeZone();
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_PAY);
  if (!sheet) {
    return {
      items: [`${month} 已有 ${existingSettlementCount} 筆鐘點結算資料，但找不到鐘點結算表。`],
      grossTotal: 0,
      netTotal: 0,
      teacherCount: 0,
      existingSettlementCount,
      isExistingSettlement: true
    };
  }

  const data = sheet.getDataRange().getValues();
  const items: string[] = [
    `注意：${month} 已有 ${existingSettlementCount} 筆鐘點結算資料；此處顯示既有結算摘要，不重新試算授課或帳務補救補發。`
  ];
  let grossTotal = 0;
  let netTotal = 0;
  let teacherCount = 0;

  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][0], timeZone);
    if (rowMonth !== month) continue;
    const teacherName = String(data[i][1] || "").trim();
    if (!teacherName) continue;
    const gross = parseFloat(data[i][6]) || 0;
    const docId = String(data[i][7] || "").trim();
    const note = String(data[i][9] || "").trim();
    const pdfUrl = String(data[i][10] || "").trim();
    const emailStatus = String(data[i][11] || "").trim();
    const taxAmount = parseFloat(data[i][12]) || 0;
    const nhiAmount = parseFloat(data[i][13]) || 0;
    const netAmount = parseFloat(data[i][14]) || gross;
    const detail = [
      String(data[i][2] || "").trim(),
      String(data[i][3] || "").trim(),
      String(data[i][4] || "").trim(),
      String(data[i][5] || "").trim()
    ].filter(function(part: string) { return part !== ""; }).join("\n");

    teacherCount++;
    grossTotal += gross;
    netTotal += netAmount;
    items.push(
      teacherName +
      "\n狀態：已寫入鐘點結算，不可重複試算寫入" +
      "\n應付：" + formatCurrency(gross) +
      "\n扣繳：" + formatCurrency(taxAmount) +
      "\n補充保費：" + formatCurrency(nhiAmount) +
      "\n實發：" + formatCurrency(netAmount) +
      "\n單據：" + (docId || "未填") +
      "\nPDF：" + (pdfUrl ? "已產生" : "尚未產生") +
      "\nEmail：" + (emailStatus || "未寄送") +
      (note ? "\n備註：" + note : "") +
      (detail ? "\n明細：\n- " + detail.split("\n").join("\n- ") : "")
    );
  }

  return { items, grossTotal, netTotal, teacherCount, existingSettlementCount, isExistingSettlement: true };
}

function appendSalaryPreviewItem(
  salaryStats: any,
  teacherName: string,
  studentName: string,
  courseName: string,
  dateText: string,
  startTime: any,
  endTime: any,
  hours: number,
  payAmount: number,
  payRate: number,
  sourceLabel: string
) {
  if (!teacherName) teacherName = "未設定講師";
  if (!salaryStats[teacherName]) salaryStats[teacherName] = { total: 0, hours: 0, details: [], missingRateCount: 0, adjustmentCount: 0 };
  salaryStats[teacherName].total += payAmount;
  salaryStats[teacherName].hours += hours;
  if (!payRate) salaryStats[teacherName].missingRateCount++;
  if (sourceLabel !== "授課") salaryStats[teacherName].adjustmentCount++;
  salaryStats[teacherName].details.push(`${sourceLabel}\n  學生：${studentName}\n  課程：${courseName}\n  時間：${dateText} ${startTime}-${endTime}\n  時數：${hours}hr\n  金額：${formatCurrency(payAmount)}`);
}

function appendSalaryAdjustmentsToStats(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  salaryStats: any,
  profitMap: any,
  month: string,
  timeZone: string
) {
  const sheet = ss.getSheetByName(SHEET_NAME_TUITION_ADJUSTMENT);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const targetMonth = normalizeFinancialMonth(data[i][1], timeZone);
    if (targetMonth !== month) continue;
    const type = String(data[i][10] || "").trim();
    if (type !== "補收") continue;
    const status = String(data[i][13] || "").trim();
    if (status === "作廢" || status === "已取消") continue;
    const hours = parseFloat(data[i][7]) || 0;
    if (hours <= 0) continue;

    const studentName = String(data[i][2] || "").trim();
    const courseName = String(data[i][3] || "").trim();
    const conf = profitMap[studentName + "_" + courseName];
    if (!conf || !conf.teacher) continue;

    const payRate = (parseFloat(conf.fee) || 0) * (parseFloat(conf.ratio) || 0);
    const payAmount = Math.round(hours * payRate);
    appendSalaryPreviewItem(
      salaryStats,
      conf.teacher,
      studentName,
      courseName,
      "原月份 " + (normalizeFinancialMonth(data[i][16], timeZone) || "未填") + "，原上課 " + formatSheetMonthDay(data[i][4], timeZone),
      formatPreviewTime(data[i][5]),
      formatPreviewTime(data[i][6]),
      hours,
      payAmount,
      payRate,
      "帳務補救補發"
    );
  }
}

function calculateSalaryDeductions(total: number, taxConfig: any) {
  let taxAmount = 0;
  let nhiAmount = 0;
  if (String(taxConfig.nationality || "").indexOf("外籍") > -1) {
    taxAmount = Math.round(total * 0.18);
  } else if (taxConfig.formatCode === "9B" && total > 20000) {
    taxAmount = Math.round(total * 0.10);
  } else if (taxConfig.formatCode === "50" && total >= 86001) {
    taxAmount = Math.round(total * 0.05);
  }
  if (total >= 20000 && taxConfig.nhiExempt !== "是") {
    nhiAmount = Math.round(total * 0.0211);
  }
  return { taxAmount, nhiAmount, netAmount: total - taxAmount - nhiAmount };
}

function buildPaymentNoticeAdminPreview(month: string) {
  try {
    const result = buildPaymentNoticeReadOnlyPreview(month);
    const items = result.items.slice(0, 12);
    if (result.items.length > items.length) {
      items.push(`另有 ${result.items.length - items.length} 位學生單據未列出，正式預覽頁後續再提供完整清單。`);
    }
    return {
      summary: `${month} 繳費單只讀預覽：${result.studentCount} 位學生，總金額 ${formatCurrency(result.grandTotal)}，已產生 PDF ${result.generatedCount} 份。`,
      items,
      rows: result.rows,
      nextAction: result.generatedCount >= result.studentCount && result.studentCount > 0 ? "繳費單已全部產生" : "確認產生繳費單",
      canConfirm: result.studentCount > 0 && result.generatedCount < result.studentCount,
      confirmAction: "adminConfirmDocument"
    };
  } catch (e) {
    return {
      summary: `${month} 繳費單只讀預覽讀取失敗。`,
      items: ["請先檢查學費結算表欄位、單據編號與 GAS 權限。", "錯誤：" + e.toString()],
      nextAction: "確認產生繳費單（尚未開放）",
      canConfirm: false
    };
  }
}

function buildPaymentNoticeReadOnlyPreview(month: string) {
  const timeZone = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_FEE);
  if (!sheet) throw new Error("找不到學費結算表。");
  const data = sheet.getDataRange().getValues();
  const studentsMap: any = {};

  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][0], timeZone);
    if (rowMonth !== month) continue;
    const studentName = String(data[i][1] || "").trim();
    if (!studentName) continue;
    if (!studentsMap[studentName]) {
      studentsMap[studentName] = {
        name: studentName,
        total: 0,
        docId: "",
        saveTime: "",
        pdfUrl: "",
        status: "",
        courses: [],
        missingDocId: false,
        missingTotal: false
      };
    }
    const item = studentsMap[studentName];
    item.courses.push({
      title: String(data[i][2] || "").trim(),
      mode: String(data[i][3] || "").trim(),
      detail: String(data[i][4] || "").trim()
    });
    if (data[i][8] !== "" && data[i][8] != null) item.total = parseFloat(data[i][8]) || 0;
    if (data[i][9]) item.docId = String(data[i][9]).trim();
    if (data[i][10]) item.saveTime = data[i][10] instanceof Date ? Utilities.formatDate(data[i][10], timeZone, "yyyy/MM/dd") : String(data[i][10]).trim();
    if (data[i][15]) item.pdfUrl = String(data[i][15]).trim();
    if (data[i][16]) item.status = String(data[i][16]).trim();
  }

  const items: string[] = [];
  const rows: any[] = [];
  let grandTotal = 0;
  let studentCount = 0;
  let generatedCount = 0;
  for (const studentName in studentsMap) {
    const item = studentsMap[studentName];
    studentCount++;
    grandTotal += item.total;
    if (item.pdfUrl || item.status.indexOf("已產") > -1) generatedCount++;
    const statusText = item.pdfUrl ? "已有 PDF" : item.status ? item.status : "尚未產生 PDF";
    const warnings: string[] = [];
    if (!item.docId) warnings.push("缺單據編號");
    if (!item.total) warnings.push("缺總金額");
    if (item.pdfUrl || item.status.indexOf("已產") > -1) warnings.push("已產生");
    const warningText = warnings.length > 0 ? "；提醒：" + warnings.join("、") : "";
    items.push(`${studentName}\n金額：${formatCurrency(item.total)}\n課程：${item.courses.length} 項\n單號：${item.docId || "未填"}\n狀態：${statusText}${warnings.length ? "\n提醒：" + warnings.join("、") : ""}`);
    const selectable = !!item.docId && !!item.total && !(item.pdfUrl || item.status.indexOf("已產") > -1);
    rows.push({
      id: "student:" + studentName,
      type: "student",
      name: studentName,
      amount: item.total,
      amountText: formatCurrency(item.total),
      docId: item.docId,
      status: statusText,
      selectable,
      selectedDefault: selectable,
      warnings,
      details: item.courses.map(function(course: any) {
        return `${course.title || "未填課程"}${course.mode ? "（" + course.mode + "）" : ""}`;
      })
    });
  }

  if (items.length === 0) {
    items.push(`${month} 學費結算表沒有可產生繳費單的資料；請先完成學費試算確認寫入。`);
  }
  return { items, rows, grandTotal, studentCount, generatedCount };
}

function buildReceiptAdminPreview(month: string) {
  try {
    const result = buildReceiptReadOnlyPreview(month);
    const items = result.items.slice(0, 12);
    if (result.items.length > items.length) {
      items.push(`另有 ${result.items.length - items.length} 位學生收據資料未列出，正式預覽頁後續再提供完整清單。`);
    }
    return {
      summary: `${month} 收據只讀預覽：${result.studentCount} 位學生，應開收據總額 ${formatCurrency(result.grandTotal)}，已有收據 PDF ${result.generatedCount} 份，待寄送 ${result.pendingSendCount} 份。`,
      items,
      rows: result.rows,
      nextAction: "確認產生收據",
      canConfirm: result.rows.some(function(row: any) { return row.selectable; }),
      confirmAction: "adminConfirmDocument"
    };
  } catch (e) {
    return {
      summary: `${month} 收據只讀預覽讀取失敗。`,
      items: ["請先檢查學費結算表、收款方式、收據類別、收款日期與 GAS 權限。", "錯誤：" + e.toString()],
      nextAction: "確認產生收據（尚未開放）",
      canConfirm: false
    };
  }
}

function buildReceiptReadOnlyPreview(month: string) {
  const timeZone = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_FEE);
  if (!sheet) throw new Error("找不到學費結算表。");
  const data = sheet.getDataRange().getValues();
  const studentsMap: any = {};

  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][0], timeZone);
    if (rowMonth !== month) continue;
    const studentName = String(data[i][1] || "").trim();
    if (!studentName) continue;
    if (!studentsMap[studentName]) {
      studentsMap[studentName] = {
        name: studentName,
        total: 0,
        docId: "",
        method: "",
        category: "",
        date: "",
        receiptUrl: "",
        status: "",
        courseCount: 0
      };
    }
    const item = studentsMap[studentName];
    if (data[i][2]) item.courseCount++;
    if (data[i][8] !== "" && data[i][8] != null) item.total = parseFloat(data[i][8]) || 0;
    if (data[i][9]) item.docId = String(data[i][9]).trim();
    if (data[i][12]) item.method = String(data[i][12]).trim();
    if (data[i][13]) item.category = String(data[i][13]).trim();
    if (data[i][14]) item.date = data[i][14] instanceof Date ? Utilities.formatDate(data[i][14], timeZone, "yyyy/MM/dd") : String(data[i][14]).trim();
    if (data[i][15]) item.receiptUrl = String(data[i][15]).trim();
    if (data[i][16]) item.status = String(data[i][16]).trim();
  }

  const items: string[] = [];
  const rows: any[] = [];
  let grandTotal = 0;
  let studentCount = 0;
  let generatedCount = 0;
  let pendingSendCount = 0;
  for (const studentName in studentsMap) {
    const item = studentsMap[studentName];
    studentCount++;
    grandTotal += item.total;
    if (item.receiptUrl) generatedCount++;
    if (item.status === "待寄送") pendingSendCount++;
    const warnings: string[] = [];
    if (!item.docId) warnings.push("缺單據編號");
    if (!item.method) warnings.push("缺收款方式");
    if (!item.category) warnings.push("缺收據類別");
    if (!item.date) warnings.push("缺收款日期");
    if (!item.total) warnings.push("缺收據金額");
    const receiptState = item.receiptUrl ? (item.status || "已有收據 PDF") : "尚未產生收據 PDF";
    items.push(`${studentName}\n金額：${formatCurrency(item.total)}\n課程：${item.courseCount} 項\n單號：${item.docId || "未填"}\n收款：${item.method || "方式未填"} / ${item.category || "類別未填"} / ${item.date || "日期未填"}\n狀態：${receiptState}${warnings.length ? "\n提醒：" + warnings.join("、") : ""}`);
    const selectable = !item.receiptUrl && !!item.docId && !!item.method && !!item.category && !!item.date && item.total > 0;
    rows.push({
      id: "student:" + studentName,
      type: "student",
      name: studentName,
      amount: item.total,
      amountText: formatCurrency(item.total),
      docId: item.docId,
      status: receiptState,
      selectable,
      selectedDefault: selectable,
      warnings,
      details: [
        "課程 " + item.courseCount + " 項",
        "收款：" + (item.method || "未填") + " / " + (item.category || "未填") + " / " + (item.date || "未填")
      ]
    });
  }

  if (items.length === 0) items.push(`${month} 學費結算表沒有可開收據的資料；請先完成學費試算與繳費單流程。`);
  return { items, rows, grandTotal, studentCount, generatedCount, pendingSendCount };
}

function createReceiptDocumentsBatch(targetMonth: string, targetName: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_FEE);
  const baseSheet = ss.getSheetByName("學生基本資料表") || ss.getSheetByName(SHEET_NAME_COURSE);
  if (!sheet || !baseSheet) return "❌ 找不到學費結算表或學生資料表";

  const timeZone = Session.getScriptTimeZone();
  const data = sheet.getDataRange().getValues();
  const baseData = baseSheet.getDataRange().getValues();
  const infoMap: any = {};
  for (let k = 1; k < baseData.length; k++) {
    const name = String(baseData[k][0] || "").trim();
    if (!name) continue;
    infoMap[name] = { email: baseData[k][1], pid: baseData[k][2] };
  }

  const folder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.RECEIPT);
  const studentsMap: any = {};
  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][0], timeZone);
    const studentName = String(data[i][1] || "").trim();
    if (rowMonth !== targetMonth || !studentName) continue;
    if (targetName && studentName !== targetName) continue;
    if (!studentsMap[studentName]) {
      studentsMap[studentName] = {
        name: studentName,
        total: 0,
        docId: "",
        method: "",
        category: "",
        date: "",
        receiptUrl: "",
        status: "",
        pid: "",
        detailParts: [],
        updateRow: -1
      };
    }
    const item = studentsMap[studentName];
    if (data[i][4]) item.detailParts.push(String(data[i][4]).trim());
    if (data[i][8] !== "" && data[i][8] != null) item.total = parseFloat(data[i][8]) || 0;
    if (data[i][9]) item.docId = String(data[i][9]).trim();
    if (data[i][11]) item.pid = String(data[i][11]).trim();
    if (data[i][12]) item.method = String(data[i][12]).trim();
    if (data[i][13]) item.category = String(data[i][13]).trim();
    if (data[i][14]) item.date = data[i][14] instanceof Date ? Utilities.formatDate(data[i][14], timeZone, "yyyy/MM/dd") : String(data[i][14]).trim();
    if (data[i][15]) item.receiptUrl = String(data[i][15]).trim();
    if (data[i][16]) item.status = String(data[i][16]).trim();
    if (data[i][8] !== "" && data[i][8] != null) item.updateRow = i + 1;
  }

  const results: string[] = [];
  let count = 0;
  for (const studentName in studentsMap) {
    const item = studentsMap[studentName];
    if (item.receiptUrl || !item.total || !item.docId || !item.method || !item.category || !item.date || item.updateRow < 0) continue;
    const info = infoMap[studentName] || {};
    const state = {
      name: studentName,
      amount: item.total,
      docId: item.docId,
      month: targetMonth,
      method: item.method,
      category: item.category,
      date: item.date,
      email: info.email,
      pid: item.pid || info.pid,
      detail: item.detailParts.join("\n")
    };
    const result = generateReceiptPDF(state, folder);
    sheet.getRange(item.updateRow, 16).setValue(result.url);
    sheet.getRange(item.updateRow, 17).setValue("待寄送");
    results.push(studentName + "：" + item.docId + " / " + formatCurrency(item.total));
    count++;
  }

  if (count === 0) return "⚠️ " + targetMonth + (targetName ? " " + targetName : "") + " 沒有可產生的收據，請確認收款方式、收據類別與收款日期。";
  return "✅ 已產生 " + count + " 份收據 PDF，狀態更新為待寄送。\n" + results.join("\n");
}

function buildAllowanceAdminPreview(month: string) {
  try {
    const result = buildAllowanceReadOnlyPreview(month);
    const items = result.items.slice(0, 12);
    if (result.items.length > items.length) {
      items.push(`另有 ${result.items.length - items.length} 位講師領據資料未列出，正式預覽頁後續再提供完整清單。`);
    }
    return {
      summary: `${month} 領據只讀預覽：${result.teacherCount} 位講師，應付總額 ${formatCurrency(result.grossTotal)}，實發總額 ${formatCurrency(result.netTotal)}，已有領據 PDF ${result.generatedCount} 份，待寄送 ${result.pendingSendCount} 份。`,
      items,
      rows: result.rows,
      nextAction: "確認產生領據",
      canConfirm: result.rows.some(function(row: any) { return row.selectable; }),
      confirmAction: "adminConfirmDocument"
    };
  } catch (e) {
    return {
      summary: `${month} 領據只讀預覽讀取失敗。`,
      items: ["請先檢查鐘點結算表、領據編號、PDF 與寄送狀態。", "錯誤：" + e.toString()],
      nextAction: "確認產生領據（尚未開放）",
      canConfirm: false
    };
  }
}

function buildAllowanceReadOnlyPreview(month: string) {
  const timeZone = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_PAY);
  if (!sheet) throw new Error("找不到鐘點結算表。");
  const data = sheet.getDataRange().getValues();
  const items: string[] = [];
  const rows: any[] = [];
  let grossTotal = 0;
  let netTotal = 0;
  let teacherCount = 0;
  let generatedCount = 0;
  let pendingSendCount = 0;

  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][0], timeZone);
    if (rowMonth !== month) continue;
    const teacherName = String(data[i][1] || "").trim();
    if (!teacherName) continue;
    const gross = parseFloat(data[i][6]) || parseFloat(data[i][5]) || 0;
    const net = parseFloat(data[i][14]) || gross;
    const docId = String(data[i][7] || "").trim();
    const pdfUrl = String(data[i][10] || "").trim();
    const status = String(data[i][11] || "").trim();
    const taxAmount = parseFloat(data[i][12]) || 0;
    const nhiAmount = parseFloat(data[i][13]) || 0;
    const detail = [data[i][2], data[i][3], data[i][4], data[i][5]].filter(function(part: any) { return String(part || "").trim() !== ""; }).join("；");
    teacherCount++;
    grossTotal += gross;
    netTotal += net;
    if (pdfUrl) generatedCount++;
    if (status === "待寄送") pendingSendCount++;
    const warnings: string[] = [];
    if (!docId) warnings.push("缺領據編號");
    if (!gross) warnings.push("缺應付金額");
    if (pdfUrl) warnings.push("已產生");
    const stateText = pdfUrl ? (status || "已有領據 PDF") : "尚未產生領據 PDF";
    items.push(`${teacherName}\n應付：${formatCurrency(gross)}\n扣繳：${formatCurrency(taxAmount)}\n補充保費：${formatCurrency(nhiAmount)}\n實發：${formatCurrency(net)}\n領據：${docId || "未填"}\n狀態：${stateText}${detail ? "\n明細：\n- " + detail.split("；").join("\n- ") : ""}${warnings.length ? "\n提醒：" + warnings.join("、") : ""}`);
    rows.push({
      id: "teacher:" + teacherName,
      type: "teacher",
      name: teacherName,
      amount: net,
      amountText: formatCurrency(net),
      docId,
      status: stateText,
      selectable: !pdfUrl && !!docId && gross > 0,
      selectedDefault: !pdfUrl && !!docId && gross > 0,
      warnings,
      details: detail ? detail.split("；") : []
    });
  }

  if (items.length === 0) items.push(`${month} 鐘點結算表沒有可開領據的資料；請先完成鐘點試算確認寫入。`);
  return { items, rows, grossTotal, netTotal, teacherCount, generatedCount, pendingSendCount };
}

function createAllowanceDocumentsBatch(targetMonth: string, targetName: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_FIN_PAY);
  if (!sheet) return "❌ 找不到鐘點結算表";
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return "❌ 找不到講師名單";

  const timeZone = Session.getScriptTimeZone();
  const data = sheet.getDataRange().getValues();
  const teacherData = teacherSheet.getDataRange().getValues();
  const teacherMap: any = {};
  for (let i = 1; i < teacherData.length; i++) {
    const name = String(teacherData[i][0] || "").trim();
    if (!name) continue;
    teacherMap[name] = {
      email: teacherData[i][2],
      pid: teacherData[i][3],
      addr: teacherData[i][4],
      phone: teacherData[i][5],
      method: teacherData[i][6],
      bank: teacherData[i][7],
      account: teacherData[i][8]
    };
  }

  const folder = DriveApp.getFolderById(PDF_FOLDER_CONFIG.ALLOWANCE);
  const today = Utilities.formatDate(new Date(), timeZone, "yyyy/MM/dd");
  const results: string[] = [];
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][0], timeZone);
    const teacherName = String(data[i][1] || "").trim();
    if (rowMonth !== targetMonth || !teacherName) continue;
    if (targetName && teacherName !== targetName) continue;

    const gross = parseFloat(data[i][6]) || parseFloat(data[i][5]) || 0;
    const docId = String(data[i][7] || "").trim();
    const pdfUrl = String(data[i][10] || "").trim();
    if (!gross || !docId || pdfUrl) continue;

    const detail = [data[i][2], data[i][3], data[i][4], data[i][5]]
      .filter(function(part: any) { return String(part || "").trim() !== ""; })
      .join("\n");
    const tInfo = teacherMap[teacherName] || {};
    const state: any = {
      name: teacherName,
      amount: gross,
      taxAmount: data[i][12] || 0,
      nhiAmount: data[i][13] || 0,
      netAmount: data[i][14] || gross,
      docId,
      date: today,
      detail,
      totalHours: data[i][5] || data[i][2],
      rowIndex: i + 1,
      email: tInfo.email,
      pid: tInfo.pid,
      addr: tInfo.addr,
      phone: tInfo.phone,
      method: tInfo.method,
      bank: tInfo.bank,
      account: tInfo.account
    };

    const pdfResult = generateAllowancePDF(state, folder);
    sheet.getRange(i + 1, 11).setValue(pdfResult.url);
    sheet.getRange(i + 1, 12).setValue("待寄送");
    results.push(teacherName + "：" + docId + " / " + formatCurrency(state.netAmount));
    count++;
  }

  if (count === 0) return "⚠️ " + targetMonth + (targetName ? " " + targetName : "") + " 沒有可產生的領據。";
  return "✅ 已產生 " + count + " 份領據 PDF，狀態更新為待寄送。\n" + results.join("\n");
}

function buildGeneralReceiptAdminPreview(month: string) {
  try {
    const result = buildGeneralReceiptReadOnlyPreview(month);
    return {
      summary: `${month} 一般收據只讀預覽：${result.recordCount} 筆，合計 ${formatCurrency(result.totalAmount)}，已有 PDF ${result.generatedCount} 份。`,
      items: result.items.slice(0, 12),
      nextAction: "確認產生一般收據（尚未開放）"
    };
  } catch (e) {
    return {
      summary: `${month} 一般收據只讀預覽讀取失敗。`,
      items: ["請先檢查一般收據紀錄分頁。", "錯誤：" + e.toString()],
      nextAction: "確認產生一般收據（尚未開放）"
    };
  }
}

function buildGeneralReceiptReadOnlyPreview(month: string) {
  const timeZone = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME_GEN_RECORD);
  if (!sheet) throw new Error("找不到一般收據紀錄。");
  const data = sheet.getDataRange().getValues();
  const items: string[] = [];
  let totalAmount = 0;
  let recordCount = 0;
  let generatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    const rowMonth = normalizeFinancialMonth(data[i][2] || data[i][3], timeZone);
    if (rowMonth !== month) continue;
    const name = String(data[i][1] || data[i][2] || "").trim();
    const category = String(data[i][4] || data[i][3] || "").trim();
    const amount = parseFloat(data[i][5]) || parseFloat(data[i][4]) || 0;
    const docId = String(data[i][6] || data[i][5] || "").trim();
    const pdfUrl = String(data[i][9] || data[i][8] || "").trim();
    recordCount++;
    totalAmount += amount;
    if (pdfUrl) generatedCount++;
    items.push(`${name || "未填姓名"}\n類別：${category || "類別未填"}\n金額：${formatCurrency(amount)}\n編號：${docId || "未填"}\n狀態：${pdfUrl ? "已有 PDF" : "尚未產生 PDF"}`);
  }
  if (items.length === 0) items.push(`${month} 目前沒有一般收據紀錄。`);
  return { items, totalAmount, recordCount, generatedCount };
}

function buildAdminPreviewMetrics(month: string, feature: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const metrics: any[] = [];

  if (feature === "學費試算" || feature === "繳費單" || feature === "收據") {
    metrics.push(buildSheetMonthMetric(ss, "授課紀錄", SHEET_NAME_RECORD, 2, month, "指定月份已登記授課筆數"));
    metrics.push(buildSheetMonthMetric(ss, "預排紀錄", SHEET_NAME_PLAN, 2, month, "指定月份預排資料筆數"));
    metrics.push(buildSheetMonthMetric(ss, "學費結算", SHEET_NAME_FIN_FEE, 0, month, "既有學費結算表筆數"));
    metrics.push(buildSheetMonthMetric(ss, "帳務補救", SHEET_NAME_TUITION_ADJUSTMENT, 1, month, "處理月份符合的補收/退費筆數"));
    return metrics;
  }

  if (feature === "鐘點試算" || feature === "領據") {
    metrics.push(buildSheetMonthMetric(ss, "授課紀錄", SHEET_NAME_RECORD, 2, month, "指定月份已登記授課筆數"));
    metrics.push(buildSheetMonthMetric(ss, "預排紀錄", SHEET_NAME_PLAN, 2, month, "指定月份預排資料筆數"));
    metrics.push(buildSheetMonthMetric(ss, "鐘點結算", SHEET_NAME_FIN_PAY, 0, month, "既有鐘點結算表筆數"));
    metrics.push(buildSheetMonthMetric(ss, "帳務補救", SHEET_NAME_TUITION_ADJUSTMENT, 1, month, "處理月份符合的補收/退費筆數"));
    return metrics;
  }

  if (feature === "一般收據") {
    metrics.push(buildSheetMonthMetric(ss, "一般收據紀錄", SHEET_NAME_GEN_RECORD, 2, month, "指定月份一般收據紀錄筆數"));
    return metrics;
  }

  if (feature === "稅務專區") {
    metrics.push(buildSheetMonthMetric(ss, "學費結算", SHEET_NAME_FIN_FEE, 0, month, "指定月份學費結算筆數"));
    metrics.push(buildSheetMonthMetric(ss, "鐘點結算", SHEET_NAME_FIN_PAY, 0, month, "指定月份鐘點結算筆數"));
    metrics.push(buildSheetMonthMetric(ss, "一般收據", SHEET_NAME_GEN_RECORD, 2, month, "指定月份一般收據紀錄筆數"));
  }

  return metrics;
}

function buildSheetMonthMetric(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, label: string, sheetName: string, dateColumnIndex: number, month: string, note: string) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return adminMetric(label, 0, note, "missing", `找不到分頁：${sheetName}`);
    const values = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < values.length; i++) {
      if (isValueInMonth(values[i][dateColumnIndex], month)) count++;
    }
    const state = count > 0 ? "ok" : "empty";
    const stateText = count > 0 ? "有資料" : "本月無資料";
    return adminMetric(label, count, note, state, stateText);
  } catch (e) {
    return adminMetric(label, 0, note, "error", "讀取失敗：" + e.toString());
  }
}

function adminMetric(label: string, value: number, note: string, state: string, stateText: string) {
  return { label, value, note, state, stateText };
}

function countSheetRowsByMonthOnly(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: string, dateColumnIndex: number, month: string) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  const values = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < values.length; i++) {
    if (isValueInMonth(values[i][dateColumnIndex], month)) count++;
  }
  return count;
}

function isValueInMonth(value: any, month: string) {
  if (!value) return false;
  const timeZone = Session.getScriptTimeZone();
  if (value instanceof Date) {
    return Utilities.formatDate(value, timeZone, "yyyy/MM") === month;
  }
  const text = String(value).trim().replace(/-/g, "/");
  if (/^\d{4}\/\d{2}/.test(text)) return text.substring(0, 7) === month;
  if (/^\d{4}\/\d{1}\//.test(text)) {
    const parts = text.split("/");
    return `${parts[0]}/${String(Number(parts[1])).padStart(2, "0")}` === month;
  }
  return false;
}

function formatSheetMonthDay(value: any, timeZone: string) {
  if (value instanceof Date) return Utilities.formatDate(value, timeZone, "MM/dd");
  const text = String(value || "").trim().replace(/-/g, "/");
  if (/^\d{4}\/\d{1,2}\/\d{1,2}/.test(text)) {
    const parts = text.split("/");
    return parts[1].padStart(2, "0") + "/" + parts[2].padStart(2, "0");
  }
  return text;
}

function formatPreviewTime(value: any) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  const text = String(value || "").trim();
  if (/^\d{1,2}:\d{2}$/.test(text)) {
    const parts = text.split(":");
    return parts[0].padStart(2, "0") + ":" + parts[1];
  }
  const digits = text.replace(/\D/g, "").padStart(4, "0");
  if (/^\d{4}$/.test(digits)) return digits.substring(0, 2) + ":" + digits.substring(2, 4);
  return text;
}

function formatCurrency(value: number) {
  const amount = Math.round(value || 0);
  return "NT$ " + amount.toLocaleString();
}

// ==========================================
// 2. 原 LINE 對話控制與財務引擎重構模組
// ==========================================

function sendReplyWithPushFallback(replyToken: string, userId: string, messagesArray: any[]) {
  if (replyToken === "LIFF_API_CALL") return; // LIFF 呼叫不發送 LINE
  
  try {
    const res = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
      "method": "post",
      "headers": { "Authorization": "Bearer " + CHANNEL_TOKEN, "Content-Type": "application/json" },
      "payload": JSON.stringify({ "replyToken": replyToken, "messages": messagesArray }),
      "muteHttpExceptions": true
    });
    if (res.getResponseCode() !== 200) {
      throw new Error("Reply Failed: " + res.getContentText());
    }
  } catch (e) {
    try {
      UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
        "method": "post",
        "headers": { "Authorization": "Bearer " + CHANNEL_TOKEN, "Content-Type": "application/json" },
        "payload": JSON.stringify({ "to": userId, "messages": messagesArray }),
        "muteHttpExceptions": true
      });
    } catch (e2) {
      Logger.log("Push Fallback also failed: " + e2.message);
    }
  }
}

function handleAnnualTaxSummaryCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;

  const isAdmin = ADMIN_LIST.indexOf(userId) > -1;
  if (!isAdmin) { replyLineMessage(replyToken, "❌ 權限不足：限行政人員使用。"); return; }

  const parts = userMsg.trim().split(/\s+/);
  if (parts.length < 2) {
    const warmMsg = "💡 **【年度扣繳憑單彙整】**\n" +
                  "請輸入欲彙整的年份。\n" +
                  "格式範例：年度扣繳 2026\n\n" +
                  "*( ⏳ 溫馨提醒：系統需掃描全年度的給付紀錄，資料量較大，送出後請稍候幾秒鐘讓機器人跑一下，期間請勿重複點擊哦！)*";
    replyLineMessage(replyToken, warmMsg);
    return;
  }

  const targetYear = parts[1];
  if (!targetYear.match(/^\d{4}$/)) { replyLineMessage(replyToken, "❌ 年份格式錯誤，請輸入四碼數字 (如 2026)。"); return; }

  const cacheKey = "LOCK_ANNUAL_" + userId;
  if (CacheService.getScriptCache().get(cacheKey)) {
    replyLineMessage(replyToken, "⏳ 系統已經在幫您結算中囉，請稍等片刻...");
    return;
  }
  CacheService.getScriptCache().put(cacheKey, "true", 60);

  try {
    const resultMsg = executeAnnualTaxSummary(targetYear);
    const messages = [
      { type: "text", text: "✅ 彙整作業完畢！" },
      { type: "text", text: resultMsg }
    ];
    sendReplyWithPushFallback(replyToken, userId, messages);
  } catch (e) {
    sendReplyWithPushFallback(replyToken, userId, [{ type: "text", text: "❌ 彙整失敗：" + e.toString() }]);
  } finally {
    CacheService.getScriptCache().remove(cacheKey);
  }
}

function executeAnnualTaxSummary(targetYear: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const reportSs = SpreadsheetApp.openById(REPORT_SHEET_ID);

  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) throw new Error("找不到「講師名單」分頁");
  const tData = teacherSheet.getDataRange().getValues();

  const taxConfigMap: any = {};
  for (let i = 1; i < tData.length; i++) {
    const tName = tData[i][0];
    if (tName) {
      taxConfigMap[tName] = {
        pid: tData[i][3] || "",
        addr: tData[i][4] || "",
        formatCode: tData[i][11] || "9B",
        nationality: tData[i][12] || "本國人",
        nhiExempt: tData[i][13] || "否"
      };
    }
  }

  const paySheet = ss.getSheetByName(SHEET_NAME_FIN_PAY);
  if (!paySheet) throw new Error("找不到「鐘點結算表」分頁");
  const pData = paySheet.getDataRange().getValues();

  const summaryMap: any = {}; 
  let totalRecords = 0;

  for (let i = 1; i < pData.length; i++) {
    const rowMonth = pData[i][0];
    const rowMonthStr = (rowMonth instanceof Date) ? Utilities.formatDate(rowMonth, Session.getScriptTimeZone(), "yyyy/MM") : String(rowMonth);

    if (rowMonthStr.indexOf(targetYear + "/") === 0) {
      const tName = pData[i][1];
      const totalAmount = parseFloat(pData[i][6]) || 0; 
      const taxAmount = parseFloat(pData[i][12]) || 0; 
      const nhiAmount = parseFloat(pData[i][13]) || 0; 
      const netAmount = parseFloat(pData[i][14]) || totalAmount; 

      if (totalAmount > 0) {
        const config = taxConfigMap[tName] || { pid: "未知", addr: "未知", formatCode: "9B", nationality: "本國籍", nhiExempt: "否" };
        const key = config.pid + "_" + config.formatCode;

        if (!summaryMap[key]) {
          summaryMap[key] = {
            name: tName, pid: config.pid, addr: config.addr, formatCode: config.formatCode, nationality: config.nationality,
            totalAmount: 0, taxAmount: 0, nhiAmount: 0, netAmount: 0
          };
        }

        summaryMap[key].totalAmount += totalAmount;
        summaryMap[key].taxAmount += taxAmount;
        summaryMap[key].nhiAmount += nhiAmount;
        summaryMap[key].netAmount += netAmount;
        totalRecords++;
      }
    }
  }

  if (totalRecords === 0) return "⚠️ 找不到 " + targetYear + " 年度的鐘點費給付紀錄。";

  const summarySheetName = "📅 年度扣繳彙整_" + targetYear;
  let summarySheet = reportSs.getSheetByName(summarySheetName);

  if (summarySheet) {
    summarySheet.clear(); 
  } else {
    summarySheet = reportSs.insertSheet(summarySheetName);
  }

  const headers = ["申報年度", "所得人姓名", "身分證/居留證", "所得格式", "身分別", "戶籍地址", "給付總額", "扣繳稅額", "補充保費", "實發總額"];
  summarySheet.appendRow(headers);

  const outputData: any[] = [];

  for (const key in summaryMap) {
    const item = summaryMap[key];
    outputData.push([
      targetYear, item.name, item.pid, item.formatCode, item.nationality, item.addr,
      item.totalAmount, item.taxAmount, item.nhiAmount, item.netAmount
    ]);
  }

  if (outputData.length > 0) summarySheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);

  summarySheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#D9EAD3");
  summarySheet.setFrozenRows(1);
  summarySheet.autoResizeColumns(1, headers.length);

  return "✅ " + targetYear + " 年度扣繳資料彙整完成！\n" +
         "共處理 " + outputData.length + " 筆歸戶資料。\n" +
         "📂 檔案已產出至「協會財務報表彙整」檔案中。";
}

function handleDonationTaxExportCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;

  const isAdmin = ADMIN_LIST.indexOf(userId) > -1;
  if (!isAdmin) { replyLineMessage(replyToken, "❌ 權限不足：限行政人員使用。"); return; }

  const parts = userMsg.trim().split(/\s+/);
  if (parts.length < 2) {
    const warmMsg = "💡 **【捐款扣除額上傳】**\n" +
                  "請輸入欲申報的年份。\n" +
                  "格式範例：捐款申報 2026\n\n" +
                  "*( ⏳ 溫馨提醒：系統會自動去會員表單比對捐款人的身分證字號，跨表單作業需要一點時間，送出後請耐心等待 3~5 秒！)*";
    replyLineMessage(replyToken, warmMsg);
    return;
  }

  const targetYear = parts[1];
  if (!targetYear.match(/^\d{4}$/)) { replyLineMessage(replyToken, "❌ 年份格式錯誤，請輸入四碼數字 (如 2026)。"); return; }

  const cacheKey = "LOCK_DONATION_" + userId;
  if (CacheService.getScriptCache().get(cacheKey)) { replyLineMessage(replyToken, "⏳ 系統正在為您跨表比對身分證中，請稍候..."); return; }
  CacheService.getScriptCache().put(cacheKey, "true", 60);

  try {
    const resultMsg = processDonationTaxExport(targetYear);
    const messages = [
      { type: "text", text: "✅ " + targetYear + " 捐款申報檔產生完成！" },
      { type: "text", text: resultMsg }
    ];
    sendReplyWithPushFallback(replyToken, userId, messages);
  } catch (e) {
    sendReplyWithPushFallback(replyToken, userId, [{ type: "text", text: "❌ 申報檔產生失敗：" + e.toString() }]);
  } finally {
    CacheService.getScriptCache().remove(cacheKey);
  }
}

function processDonationTaxExport(targetYear: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recordSheet = ss.getSheetByName(SHEET_NAME_GEN_RECORD);
  if (!recordSheet) throw new Error("找不到「" + SHEET_NAME_GEN_RECORD + "」分頁");

  const data = recordSheet.getDataRange().getValues();
  const csvData: any[] = [];
  const headers = ["受贈單位統編", "捐贈人身分證/統編", "捐贈人姓名", "捐贈日期", "捐贈金額", "收據字軌號碼", "捐贈用途"];
  csvData.push(headers.join(","));

  let successCount = 0; let skipCount = 0; let totalAmount = 0; const missingIdNames: string[] = [];
  const orgTaxId = ORG_TAX_ID || "91622132"; 

  for (let i = 1; i < data.length; i++) {
    const docId = data[i][1]; const dateRaw = data[i][2]; const name = String(data[i][3]).trim(); 
    const amount = parseInt(data[i][4], 10); const category = data[i][5];    

    if (category === "捐款") {
      const dateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), "yyyy/MM/dd") : String(dateRaw);
      if (dateStr.indexOf(targetYear + "/") === 0) {
        const memberInfo = lookupGeneralMemberData(name, category);
        const pid = memberInfo.pid;

        if (!pid || pid === "") {
          skipCount++; if (missingIdNames.indexOf(name) === -1) missingIdNames.push(name);
          continue;
        }

        const dateParts = dateStr.split("/");
        if (dateParts.length === 3) {
          const rocYear = parseInt(dateParts[0], 10) - 1911;
          const rocDateStr = rocYear.toString() + dateParts[1] + dateParts[2];
          const rowObj = [ orgTaxId, pid, name, rocDateStr, amount, docId, "捐款" ];
          csvData.push(rowObj.join(","));
          successCount++; totalAmount += amount;
        }
      }
    }
  }

  if (successCount === 0) return "⚠️ 找不到 " + targetYear + " 年度符合申報條件的捐款紀錄。";

  const csvString = csvData.join("\n");
  const blob = Utilities.newBlob("\uFEFF" + csvString, "text/csv", "捐款申報匯出_" + targetYear + ".csv");
  const file = DriveApp.createFile(blob);
  
  let report = "--------------------\n🔹 成功彙整：" + successCount + " 筆\n💰 申報總額：$" + totalAmount + "\n📂 檔案連結：\n" + file.getUrl();
  if (skipCount > 0) report += "\n\n⚠️ 有 " + skipCount + " 筆捐款因「缺少身分證字號」被排除。\n名單：" + missingIdNames.join(", ");
  return report;
}

function handleTaxExemptionDashboardCommand(event: any, userMsg: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;

  const isAdmin = ADMIN_LIST.indexOf(userId) > -1;
  if (!isAdmin) { replyLineMessage(replyToken, "❌ 權限不足：限行政人員使用。"); return; }

  const parts = userMsg.trim().split(/\s+/);
  if (parts.length < 2) {
    const warmMsg = "💡 **【80% 免稅預警儀表板】**\n" +
                  "請輸入您想檢視的年份。\n" +
                  "格式範例：免稅試算 2026\n\n" +
                  "*( ⏳ 溫馨提醒：系統將根據會計日記帳進行全年結算，並為您自動繪製雲端圖表，送出後請稍候片刻哦！)*";
    replyLineMessage(replyToken, warmMsg);
    return;
  }

  const targetYear = parts[1];
  if (!targetYear.match(/^\d{4}$/)) { replyLineMessage(replyToken, "❌ 年份格式錯誤，請輸入四碼數字 (如 2026)。"); return; }

  const cacheKey = "LOCK_TAX_DASH_" + userId;
  if (CacheService.getScriptCache().get(cacheKey)) { replyLineMessage(replyToken, "⏳ 系統正在為您繪製戰情儀表板，請稍候..."); return; }
  CacheService.getScriptCache().put(cacheKey, "true", 60);

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const journalSheet = ss.getSheetByName("會計日記帳");
    if (!journalSheet) throw new Error("找不到會計日記帳");

    const jData = journalSheet.getDataRange().getValues();
    const timeZone = Session.getScriptTimeZone();
    
    let totalRev = 0; let totalExp = 0;
    const expBreakdown: any = {};

    for (let i = 1; i < jData.length; i++) {
      const dateRaw = jData[i][0];
      if (!dateRaw) continue;
      const dateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, timeZone, "yyyy/MM/dd") : String(dateRaw);
      
      if (dateStr.indexOf(targetYear + "/") === 0) {
        const code = String(jData[i][2]);
        const name = String(jData[i][3]);
        const drAmt = parseFloat(jData[i][5]) || 0;
        const crAmt = parseFloat(jData[i][6]) || 0;

        if (code.charAt(0) === '4') {
          totalRev += (crAmt - drAmt);
        } else if (code.charAt(0) === '5') {
          const netExp = (drAmt - crAmt);
          totalExp += netExp;
          if (!expBreakdown[name]) expBreakdown[name] = 0;
          expBreakdown[name] += netExp;
        }
      }
    }

    const requiredExpense = totalRev * 0.8; 
    const currentRatio = (totalRev > 0) ? parseFloat(((totalExp / totalRev) * 100).toFixed(1)) : 0;
    const gapAmount = requiredExpense - totalExp;
    const isSafe = gapAmount <= 0;

    const dashboardUrl = buildCloudDashboard(targetYear, totalRev, totalExp, expBreakdown, requiredExpense, currentRatio);

    const dashData = {
      year: targetYear, totalInc: totalRev, totalExp: totalExp, 
      reqExp: requiredExpense, ratio: currentRatio, gap: isSafe ? 0 : gapAmount, isSafe: isSafe, url: dashboardUrl
    };

    const flexCard = createTaxDashboardFlexCard(dashData);
    const messages = [
      { type: "text", text: "✅ " + targetYear + " 年度免稅門檻試算與圖表繪製完成！" },
      { type: "flex", altText: "免稅預警儀表板", contents: flexCard }
    ];

    sendReplyWithPushFallback(replyToken, userId, messages);

  } catch (e) {
    sendReplyWithPushFallback(replyToken, userId, [{ type: "text", text: "❌ 試算失敗：" + e.toString() }]);
  } finally {
    CacheService.getScriptCache().remove(cacheKey);
  }
}

function buildCloudDashboard(targetYear: string, totalRev: number, totalExp: number, expBreakdown: any, requiredExpense: number, currentRatio: number) {
  const reportSs = SpreadsheetApp.openById(REPORT_SHEET_ID);
  const dashSheetName = "📈 " + targetYear + "年_戰情儀表板";
  let dashSheet = reportSs.getSheetByName(dashSheetName);
  
  if (!dashSheet) {
    dashSheet = reportSs.insertSheet(dashSheetName, 0); 
  } else {
    dashSheet.clear();
    const charts = dashSheet.getCharts();
    for (let i = 0; i < charts.length; i++) dashSheet.removeChart(charts[i]);
    reportSs.setActiveSheet(dashSheet);
    reportSs.moveActiveSheet(1);
  }

  dashSheet.getRange("A1").setValue("🏛️ 協會財務戰情室 (" + targetYear + "年度)");
  dashSheet.getRange("A1:E1").merge().setFontSize(16).setFontWeight("bold").setBackground("#2C3E50").setFontColor("white").setHorizontalAlignment("center");

  dashSheet.getRange("A3:B7").setValues([
    ["💰 總收入總計 (A)", totalRev],
    ["💸 總支出總計 (B)", totalExp],
    ["⚖️ 免稅法定 80% 門檻 (A * 80%)", requiredExpense],
    ["⚠️ 目前支出缺口 (尚需消化)", (requiredExpense - totalExp) > 0 ? (requiredExpense - totalExp) : 0],
    ["🎯 目前支出達成率", currentRatio + "%"]
  ]);
  dashSheet.getRange("A3:A7").setFontWeight("bold").setBackground("#ECF0F1");
  dashSheet.getRange("B3:B6").setNumberFormat('"$"#,##0');

  let r = 3;
  dashSheet.getRange("D2:E2").setValues([["支出項目", "金額"]]).setFontWeight("bold");
  for (const k in expBreakdown) {
    if (expBreakdown[k] > 0) {
      dashSheet.getRange(r, 4).setValue(k);
      dashSheet.getRange(r, 5).setValue(expBreakdown[k]);
      r++;
    }
  }

  dashSheet.getRange("G2:H2").setValues([["指標", "金額"]]).setFontWeight("bold");
  dashSheet.getRange("G3:H4").setValues([
    ["目前總支出", totalExp],
    ["法定 80% 門檻", requiredExpense]
  ]);

  if (r > 3) {
    const pieChart = dashSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dashSheet.getRange("D2:E" + (r - 1)))
      .setPosition(9, 1, 0, 0)
      .setOption('title', '各項支出結構佔比分析')
      .setOption('pieHole', 0.4)
      .setOption('width', 400).setOption('height', 300)
      .build();
    dashSheet.insertChart(pieChart);
  }

  const barColor = (totalExp >= requiredExpense) ? '#1DB446' : '#E74C3C';
  const barChart = dashSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dashSheet.getRange("G2:H4"))
    .setPosition(9, 5, 0, 0)
    .setOption('title', '免稅門檻 80% 達成進度對比')
    .setOption('vAxis.minValue', 0)
    .setOption('colors', [barColor])
    .setOption('legend', {position: 'none'})
    .setOption('width', 400).setOption('height', 300)
    .build();
  dashSheet.insertChart(barChart);

  dashSheet.autoResizeColumns(1, 2);
  dashSheet.setColumnWidth(3, 20);
  dashSheet.hideColumns(4, 5);

  return reportSs.getUrl() + "#gid=" + dashSheet.getSheetId();
}

function createTaxDashboardFlexCard(data: any) {
  const formatMoney = function(n: number) { return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","); };
  const statusColor = data.isSafe ? "#1DB446" : "#E74C3C"; 
  const statusText = data.isSafe ? "✅ 支出已達標 (免稅安全)" : "⚠️ 未達 80% 法定門檻！";
  const gapBlock = data.isSafe ? 
      { "type": "text", "text": "恭喜！年度支出已達法定要求", "size": "sm", "color": "#1DB446", "weight": "bold", "wrap": true } :
      { "type": "text", "text": "距免稅門檻還差 $" + formatMoney(data.gap), "size": "md", "color": "#E74C3C", "weight": "bold", "wrap": true };

  return {
    "type": "bubble", "size": "giga",
    "body": {
      "type": "box", "layout": "vertical",
      "contents": [
        { "type": "text", "text": "🏛️ " + data.year + " 免稅預警儀表板", "weight": "bold", "size": "xl", "color": "#2C3E50" },
        { "type": "text", "text": statusText, "size": "md", "color": statusColor, "weight": "bold", "margin": "md" },
        { "type": "separator", "margin": "lg" },
        { "type": "box", "layout": "horizontal", "margin": "md", "contents": [ { "type": "text", "text": "合計總收入", "size": "sm", "weight": "bold", "color": "#2980B9", "flex": 2 }, { "type": "text", "text": "$" + formatMoney(data.totalInc), "size": "sm", "weight": "bold", "color": "#2980B9", "align": "end", "flex": 3 } ] },
        { "type": "box", "layout": "horizontal", "margin": "md", "contents": [ { "type": "text", "text": "合計總支出", "size": "sm", "weight": "bold", "color": "#D35400", "flex": 2 }, { "type": "text", "text": "$" + formatMoney(data.totalExp), "size": "sm", "weight": "bold", "color": "#D35400", "align": "end", "flex": 3 } ] },
        { "type": "separator", "margin": "lg" },
        {
          "type": "box", "layout": "vertical", "margin": "lg", "backgroundColor": "#F4F6F6", "paddingAll": "md", "cornerRadius": "md",
          "contents": [
            { "type": "text", "text": "法定 80% 低標：$" + formatMoney(data.reqExp), "size": "sm", "color": "#34495E", "weight": "bold" },
            { "type": "text", "text": "目前支出比例：" + data.ratio + "%", "size": "md", "color": statusColor, "weight": "bold", "margin": "sm" },
            { "type": "separator", "margin": "md" },
            gapBlock
          ]
        },
        { "type": "button", "style": "link", "height": "sm", "margin": "md", "action": { "type": "uri", "label": "📈 查看雲端戰情室圖表", "uri": data.url } }
      ]
    }
  };
}

function handleTuitionCalculation(event: any, userMsg: string) {
  const replyToken = event.replyToken; const userId = event.source.userId;
  const isAdmin = ADMIN_LIST.indexOf(userId) > -1;
  if (!isAdmin) { replyLineMessage(replyToken, "❌ 權限不足：限行政人員使用。"); return; }

  const parts = userMsg.split(" "); let baseMonthStr = ""; const timeZone = Session.getScriptTimeZone();
  if (parts.length > 1 && parts[1].match(/^\d{4}\/\d{2}$/)) { baseMonthStr = parts[1]; } 
  else { const d = new Date(); d.setMonth(d.getMonth() - 1); baseMonthStr = Utilities.formatDate(d, timeZone, "yyyy/MM"); }

  const splitDt = baseMonthStr.split("/"); const nextDate = new Date(parseInt(splitDt[0]), parseInt(splitDt[1]) - 1 + 1, 1); const nextMonthStr = Utilities.formatDate(nextDate, timeZone, "yyyy/MM");
  const ss = SpreadsheetApp.openById(SHEET_ID); const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD); const planSheet = ss.getSheetByName(SHEET_NAME_PLAN); const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  if (!recordSheet || !planSheet || !courseSheet) { replyLineMessage(replyToken, "❌ 資料表缺失。"); return; }

  const courseData = courseSheet.getDataRange().getValues(); const configMap: any = {};
  for (let i = 1; i < courseData.length; i++) {
    const sName = courseData[i][2]; const cName = courseData[i][3]; const fee = courseData[i][4]; const mode = courseData[i][6];
    if (sName && cName) { configMap[sName + "_" + cName] = { fee: fee, mode: (mode === "預收" ? "預收" : "後收"), teacher: courseData[i][1] }; }
  }

  const stats: any = {};
  function initStats(s: string, c: string, t: string, fee: number, mode: string) {
    if (!stats[s]) stats[s] = {};
    if (!stats[s][c]) { stats[s][c] = { teacher: t, fee: fee, mode: mode, recordBase: 0, planBase: 0, pendingPlanBase: 0, planNext: 0, detailsRec: [], detailsPending: [], detailsPlan: [], adjustments: [] }; }
  }

  const rData = recordSheet.getDataRange().getValues();
  for (let i = 1; i < rData.length; i++) {
    const rowMonth = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "yyyy/MM") : String(rData[i][2]).substring(0, 7);
    const settled = rData[i][9];
    if (rowMonth == baseMonthStr && (!settled || settled === "")) {
      const sName = rData[i][7]; const cName = rData[i][8]; const hr = parseFloat(rData[i][5]); const key = sName + "_" + cName;
      const conf = configMap[key] || { fee: 0, mode: "後收", teacher: rData[i][1] };
      initStats(sName, cName, conf.teacher, conf.fee, conf.mode); stats[sName][cName].recordBase += hr;
      const perLessonAmt = Math.round(hr * conf.fee); const dText = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "MM/dd") : rData[i][2];
      stats[sName][cName].detailsRec.push("[實上] " + dText + " " + rData[i][3] + "-" + rData[i][4] + " ($" + perLessonAmt + ")");
    }
  }

  const pData = planSheet.getDataRange().getValues();
  for (let i = 1; i < pData.length; i++) {
    const lessonDateMonth = (pData[i][2] instanceof Date) ? Utilities.formatDate(pData[i][2], timeZone, "yyyy/MM") : String(pData[i][2]).substring(0, 7);
    
    const getFmtMonth = function(val: any) {
      if (!val) return "";
      if (val instanceof Date) return Utilities.formatDate(val, timeZone, "yyyy/MM");
      const s = String(val).trim();
      if (s.match(/^\d{4}[\/-]\d{2}/)) return s.substring(0, 7).replace("-", "/");
      return s;
    };

    const refundSettledMonth = getFmtMonth(pData[i][11]);
    const status = pData[i][9];
    const targetMonthForPlanBase = refundSettledMonth || lessonDateMonth;

    if (status !== "取消" || (status === "取消" && refundSettledMonth)) {
      const sName = pData[i][7]; const cName = pData[i][8]; const hr = parseFloat(pData[i][5]); const key = sName + "_" + cName; const conf = configMap[key];
      if (conf && conf.mode === "預收") {
        initStats(sName, cName, conf.teacher, conf.fee, conf.mode);
        
        if (targetMonthForPlanBase == baseMonthStr) { 
          const dText = (pData[i][2] instanceof Date) ? Utilities.formatDate(pData[i][2], timeZone, "MM/dd") : pData[i][2];
          if (String(status || "").trim() === "未核銷") {
            stats[sName][cName].pendingPlanBase += hr;
            const perLessonAmt = Math.round(hr * conf.fee);
            stats[sName][cName].detailsPending.push("[待核銷預排] " + dText + " " + pData[i][3] + "-" + pData[i][4] + " ($" + perLessonAmt + ")");
          } else {
            stats[sName][cName].planBase += hr;
          }
          if (status === "取消") {
            stats[sName][cName].detailsRec.push("[歷史補救] " + dText + " " + pData[i][3] + "-" + pData[i][4] + " (取消退費)");
          }
        } else if (lessonDateMonth == nextMonthStr && status !== "取消") {
          stats[sName][cName].planNext += hr; const perLessonAmt = Math.round(hr * conf.fee); const dText = (pData[i][2] instanceof Date) ? Utilities.formatDate(pData[i][2], timeZone, "MM/dd") : pData[i][2];
          stats[sName][cName].detailsPlan.push("[預收] " + dText + " " + pData[i][3] + "-" + pData[i][4] + " ($" + perLessonAmt + ")");
        }
      }
    }
  }

  appendTuitionAdjustmentsToStats(ss, stats, configMap, baseMonthStr);

  let report = "💰 學費試算單 (" + baseMonthStr + ")\n\n"; let saveData: any[] = []; let grandTotal = 0; let hasData = false; let hasPendingPlan = false;
  for (const sName in stats) {
    let sTotal = 0; let sDetailText = ""; const courseRows = [];
    for (const cName in stats[sName]) {
      const item = stats[sName][cName]; let finalAmount = 0; let formulaStr = ""; let fullDetails: string[] = [];
      const adjustmentTotal = item.adjustments.reduce(function(sum: number, adj: any) { return sum + adj.amount; }, 0);
      if (item.mode === "預收") {
        if (item.pendingPlanBase > 0) hasPendingPlan = true;
        const diff = item.pendingPlanBase > 0 ? 0 : Math.round((item.recordBase - item.planBase) * 10) / 10;
        const totalHr = item.planNext + diff; finalAmount = Math.round(totalHr * item.fee) + adjustmentTotal;
        const diffStr = (diff > 0) ? (" +補" + diff + "hr") : (diff < 0 ? (" -退" + Math.abs(diff) + "hr") : "");
        const adjustmentStr = adjustmentTotal !== 0 ? "，帳務補救 " + formatCurrency(adjustmentTotal) : "";
        formulaStr = item.pendingPlanBase > 0 ? ("預收" + item.planNext + "hr，待核銷" + item.pendingPlanBase + "hr 暫不退費" + adjustmentStr) : ("預收" + item.planNext + "hr" + diffStr + " = " + totalHr + "hr" + adjustmentStr);
        if (item.detailsRec.length > 0 || item.planBase > 0) {
            fullDetails.push("📋 [上月核對] 實上" + item.recordBase + " / 預繳" + item.planBase); fullDetails = fullDetails.concat(item.detailsRec); 
            if (diff !== 0) fullDetails.push("⚠️ 差異金額: " + (diff > 0 ? "補收 $":"退費 $") + Math.abs(Math.round(diff * item.fee))); else fullDetails.push("✅ 差異: 無 (已結清)"); fullDetails.push(""); 
        }
        if (item.detailsPending.length > 0) { fullDetails.push("⏳ [待核銷預排] 尚未列入退費"); fullDetails = fullDetails.concat(item.detailsPending); fullDetails.push(""); }
        if (item.detailsPlan.length > 0) { fullDetails.push("📅 [下月預收] " + nextMonthStr); fullDetails = fullDetails.concat(item.detailsPlan); }
      } else {
        finalAmount = Math.round(item.recordBase * item.fee) + adjustmentTotal; formulaStr = "實上 " + item.recordBase + " hr × $" + item.fee + (adjustmentTotal !== 0 ? "，帳務補救 " + formatCurrency(adjustmentTotal) : ""); fullDetails = item.detailsRec;
      }
      if (adjustmentTotal !== 0) {
        const adjustmentLines = item.adjustments.map(function(adj: any) {
          return "[帳務補救] " + adj.type + " " + formatCurrency(adj.amount) + "（" + adj.status + "）";
        });
        if (fullDetails.length > 0) fullDetails.push("");
        fullDetails.push("🧾 [本月帳務補救]");
        fullDetails = fullDetails.concat(adjustmentLines);
      }
      if (finalAmount !== 0 || item.recordBase > 0 || item.planNext > 0 || item.pendingPlanBase > 0 || adjustmentTotal !== 0) {
        const canWriteCourse = item.pendingPlanBase <= 0;
        sTotal += canWriteCourse ? finalAmount : 0; const detailBlock = fullDetails.join("\n"); if (sDetailText !== "") sDetailText += "\n--------------------\n";
        sDetailText += "   📘 " + cName + " (" + item.mode + ")\n" + detailBlock.replace(/^/gm, "      ") + "\n      ➤ 本科總計：$" + finalAmount;
        if (canWriteCourse) courseRows.push([ baseMonthStr, sName, cName, item.mode, detailBlock, formulaStr, item.fee, finalAmount, "", "", "" ]);
      }
    }
    if (courseRows.length > 0) {
      courseRows[courseRows.length - 1][8] = sTotal; saveData = saveData.concat(courseRows); report += "🎓 " + sName + "\n" + sDetailText + "\n   💰 本期應繳：$" + sTotal + "\n\n";
      grandTotal += sTotal; hasData = true;
    }
  }

  if (!hasData) { replyLineMessage(replyToken, "💰 學費試算 (" + baseMonthStr + ")\n無可寫入的結算資料；待核銷預排暫不可寫入。"); } else {
    report += "════════════════\n總金額： $" + grandTotal + "\n\n請確認是否寫入？";
    const cacheKey = "FIN_" + userId; const cacheData = { targetSheet: SHEET_NAME_FIN_FEE, save: saveData, updateTargetMonth: baseMonthStr, category: "學費", prefix: "R" };
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(cacheData), 600); replyConfirmationCard(replyToken, "學費試算確認", report, cacheKey);
  }
}

function handleSalaryCalculation(event: any, userMsg: string) {
  const replyToken = event.replyToken; const userId = event.source.userId;
  const isAdmin = ADMIN_LIST.indexOf(userId) > -1;
  if (!isAdmin) { replyLineMessage(replyToken, "❌ 權限不足：限行政人員使用。"); return; }

  const parts = userMsg.split(" "); let queryMonth = ""; const timeZone = Session.getScriptTimeZone();
  if (parts.length > 1 && parts[1].match(/^\d{4}\/\d{2}$/)) { queryMonth = parts[1]; } else { const d = new Date(); d.setMonth(d.getMonth() - 1); queryMonth = Utilities.formatDate(d, timeZone, "yyyy/MM"); }

  const ss = SpreadsheetApp.openById(SHEET_ID); const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD); const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE); const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER); 
  if (!recordSheet || !courseSheet || !teacherSheet) { replyLineMessage(replyToken, "❌ 找不到必要的資料表。"); return; }

  const courseData = courseSheet.getDataRange().getValues(); const profitMap: any = {};
  for (let i = 1; i < courseData.length; i++) { const key = courseData[i][2] + "_" + courseData[i][3]; const fee = courseData[i][4]; const ratio = courseData[i][5]; const teacher = String(courseData[i][1] || "").trim(); if (fee && ratio) { profitMap[key] = { fee: fee, ratio: ratio, teacher: teacher }; } }

  const tData = teacherSheet.getDataRange().getValues(); const taxConfigMap: any = {};
  for (let i = 1; i < tData.length; i++) { const tName = tData[i][0]; if (tName) { taxConfigMap[tName] = { formatCode: tData[i][11] || "9B", nationality: tData[i][12] || "本國人", nhiExempt: tData[i][13] || "否" }; } }

  const salaryStats: any = {}; const rData = recordSheet.getDataRange().getValues(); const updateRows: number[] = [];
  for (let i = 1; i < rData.length; i++) {
    const rowMonth = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "yyyy/MM") : String(rData[i][2]).substring(0, 7);
    const settled = rData[i][10]; const courseName = rData[i][8];
    if (courseName.indexOf("取消") === -1 && rowMonth == queryMonth && (!settled || settled === "")) {
      const tName = rData[i][1]; const sName = rData[i][7]; const hr = parseFloat(rData[i][5]); const key = sName + "_" + courseName; const conf = profitMap[key] || { fee: 0, ratio: 0 };
      const payRate = conf.fee * conf.ratio; const payAmount = Math.round(hr * payRate); 
      const dText = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "MM/dd") : rData[i][2];
      appendSalarySaveItem(salaryStats, tName, sName, courseName, dText, rData[i][3], rData[i][4], hr, payRate, payAmount, "");
      updateRows.push(i + 1);
    }
  }
  appendSalaryAdjustmentsToSaveStats(ss, salaryStats, profitMap, queryMonth, timeZone);

  let report = "💰 鐘點費試算 (" + queryMonth + ")\n\n"; const saveData: any[] = []; let hasData = false;
  for (const tName in salaryStats) {
    const item = salaryStats[tName]; const total = item.total; const taxConfig = taxConfigMap[tName] || { formatCode: "9B", nationality: "本國人", nhiExempt: "否" };
    let taxAmount = 0; let nhiAmount = 0;
    if (taxConfig.nationality.indexOf("外籍") > -1) { taxAmount = Math.round(total * 0.18); } else { if (taxConfig.formatCode === "9B" && total > 20000) { taxAmount = Math.round(total * 0.10); } else if (taxConfig.formatCode === "50" && total >= 86001) { taxAmount = Math.round(total * 0.05); } }
    if (total >= 20000 && taxConfig.nhiExempt !== "是") { nhiAmount = Math.round(total * 0.0211); }
    const netAmount = total - taxAmount - nhiAmount; 

    report += "👨‍🏫 " + tName + " (" + taxConfig.formatCode + ")\n"; const detailStr = item.details.join("\n"); report += detailStr.replace(/^/gm, "   ") + "\n";
    report += "   💵 應付總額：$" + total + "\n"; if (taxAmount > 0) report += "   ➖ 扣繳稅額：$" + taxAmount + "\n"; if (nhiAmount > 0) report += "   ➖ 補充保費：$" + nhiAmount + "\n"; report += "   💰 實發金額：$" + netAmount + "\n--------------------\n";
    
    const lessonCount = item.entries ? item.entries.length : item.details.length;
    const hasAdjustment = detailStr.indexOf("帳務補救補發") > -1;
    const courseStudent = (item.entries || []).map(function(entry: any) { return entry.courseStudent; }).join("\n") || detailStr;
    const dateTimeAmount = (item.entries || []).map(function(entry: any) { return entry.dateTimeAmount; }).join("\n") || detailStr;
    const hoursRate = (item.entries || []).map(function(entry: any) { return entry.hoursRate; }).join("\n");
    const singleCalc = (item.entries || []).map(function(entry: any) { return entry.singleCalc; }).join("\n") || ("共" + lessonCount + "筆");
    const note = hasAdjustment ? "帳務補救補發" : (lessonCount > 1 ? "多筆彙整" : "");
    saveData.push([ queryMonth, tName, courseStudent, dateTimeAmount, hoursRate, singleCalc, total, "", "", note, "", "", taxAmount, nhiAmount, netAmount ]);
    hasData = true;
  }

  if (!hasData) { replyLineMessage(replyToken, "💰 鐘點費試算 (" + queryMonth + ")\n無須結算資料。"); } else {
    report += "請確認是否寫入結算工作表？";
    const cacheKey = "FIN_" + userId; const cacheData = { targetSheet: SHEET_NAME_FIN_PAY, save: saveData, updateTargetMonth: queryMonth, updateRows: updateRows, category: "鐘點費", prefix: "A" };
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(cacheData), 600); replyConfirmationCard(replyToken, "鐘點費試算確認", report, cacheKey);
  }
}

function appendSalaryAdjustmentsToSaveStats(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, salaryStats: any, profitMap: any, month: string, timeZone: string) {
  const sheet = ss.getSheetByName(SHEET_NAME_TUITION_ADJUSTMENT);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const targetMonth = normalizeFinancialMonth(data[i][1], timeZone);
    if (targetMonth !== month) continue;
    const type = String(data[i][10] || "").trim();
    if (type !== "補收") continue;
    const status = String(data[i][13] || "").trim();
    if (status === "作廢" || status === "已取消") continue;
    const hr = parseFloat(data[i][7]) || 0;
    if (hr <= 0) continue;
    const sName = String(data[i][2] || "").trim();
    const courseName = String(data[i][3] || "").trim();
    const conf = profitMap[sName + "_" + courseName];
    if (!conf || !conf.teacher) continue;
    const tName = conf.teacher;
    const payRate = (parseFloat(conf.fee) || 0) * (parseFloat(conf.ratio) || 0);
    const payAmount = Math.round(hr * payRate);
    const dText = formatSheetMonthDay(data[i][4], timeZone);
    appendSalarySaveItem(
      salaryStats,
      tName,
      sName,
      courseName,
      dText,
      formatPreviewTime(data[i][5]),
      formatPreviewTime(data[i][6]),
      hr,
      payRate,
      payAmount,
      "帳務補救補發"
    );
  }
}

function appendSalarySaveItem(
  salaryStats: any,
  teacherName: string,
  studentName: string,
  courseName: string,
  dateText: string,
  startTime: any,
  endTime: any,
  hours: number,
  payRate: number,
  payAmount: number,
  note: string
) {
  if (!salaryStats[teacherName]) salaryStats[teacherName] = { total: 0, details: [], entries: [] };
  const normalizedStart = formatPreviewTime(startTime);
  const normalizedEnd = formatPreviewTime(endTime);
  const courseStudent = (note ? note + "：" : "") + studentName + "(" + courseName + ")";
  const dateTimeAmount = dateText + " " + normalizedStart + "-" + normalizedEnd + " ($" + payAmount + ")";
  const hoursRate = hours + "hr x $" + Math.round(payRate);
  const singleCalc = "$" + payAmount;
  salaryStats[teacherName].details.push(courseStudent + " " + dateTimeAmount);
  salaryStats[teacherName].entries.push({ courseStudent, dateTimeAmount, hoursRate, singleCalc, note });
  salaryStats[teacherName].total += payAmount;
}

function executeFinancialSave(event: any, postbackData: string) {
  const replyToken = event.replyToken; const userId = event.source.userId;
  const cacheKey = "FIN_" + userId; const cacheDataStr = CacheService.getScriptCache().get(cacheKey);
  if (!cacheDataStr) { replyLineMessage(replyToken, "⚠️ 快取已過期或無此結算資料。"); return; }
  const cacheObj = JSON.parse(cacheDataStr); const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 使用 LockService 以防寫入衝突 (優化建議 #6)
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 鎖定最多10秒
    
    const targetSheet = ss.getSheetByName(cacheObj.targetSheet);
    if (!targetSheet) { replyLineMessage(replyToken, "❌ 寫入失敗，找不到工作表。"); return; }

    const timeZone = Session.getScriptTimeZone();
    const now = new Date();
    
    // (A) 產生主檔的單據編號字軌，例如 R_2026_06_001 / A_2026_06_001
    const invoicePrefix = buildFinancialInvoicePrefix(cacheObj.category, cacheObj.updateTargetMonth, cacheObj.prefix);
    
    const sheetData = targetSheet.getDataRange().getValues();
    const invoiceColumnIndex = getFinancialInvoiceColumnIndex(cacheObj.category);
    const duplicateKeys = getFinancialDuplicateKeys(cacheObj.save, cacheObj.category);
    for (let j = 1; j < sheetData.length; j++) {
      const rowMonth = normalizeFinancialMonth(sheetData[j][0], timeZone);
      const rowKey = getFinancialRowKey(sheetData[j], cacheObj.category);
      if (rowMonth === cacheObj.updateTargetMonth && rowKey && duplicateKeys.indexOf(rowKey) > -1) {
        replyLineMessage(replyToken, "⚠️ " + cacheObj.updateTargetMonth + " 的「" + rowKey + "」已有結算紀錄，系統已阻擋重複寫入。若需重算，請先由行政確認並清理舊結算資料。");
        return;
      }
    }

    let maxSerial = 0;
    for (let j = 1; j < sheetData.length; j++) {
      const rowInvoice = String(sheetData[j][invoiceColumnIndex] || "");
      if (rowInvoice.indexOf(invoicePrefix) === 0) {
        const serial = parseFinancialInvoiceSerial(rowInvoice, invoicePrefix);
        if (serial > maxSerial) maxSerial = serial;
      }
    }

    const outputRows = cacheObj.save;
    for (let k = 0; k < outputRows.length; k++) {
      const shouldIssueInvoice = cacheObj.category !== "學費" || outputRows[k][8];
      if (shouldIssueInvoice) {
        maxSerial++;
        const paddedSerial = (maxSerial < 10 ? "00" : (maxSerial < 100 ? "0" : "")) + maxSerial;
        const finalInvoiceId = invoicePrefix + paddedSerial;
        outputRows[k][invoiceColumnIndex] = finalInvoiceId;
      }
      if (cacheObj.category === "學費" && shouldIssueInvoice) {
        outputRows[k][10] = now;
      }
      targetSheet.appendRow(outputRows[k]);
    }

    // (B) 回溯標記原始「授課紀錄」為已結算
    if (cacheObj.updateRows && cacheObj.updateRows.length > 0) {
      const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
      if (recordSheet) {
        for (let idx = 0; idx < cacheObj.updateRows.length; idx++) {
          const rowNum = cacheObj.updateRows[idx];
          recordSheet.getRange(rowNum, 11).setValue(cacheObj.updateTargetMonth); // K欄：鐘點費結算日期
        }
      }
    } else {
      // 學費回溯標記：只標記本次實際寫入的學生/課程，避免待核銷或未勾選項目被誤標記。
      const tuitionKeys = getTuitionKeysFromSavedRows(cacheObj.save || []);
      const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
      if (recordSheet) {
        const rData = recordSheet.getDataRange().getValues();
        for (let i = 1; i < rData.length; i++) {
          const rowMonth = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "yyyy/MM") : String(rData[i][2]).substring(0, 7);
          const settled = rData[i][9];
          const rowKey = buildTuitionSelectionKey(rData[i][7], rData[i][8]);
          if (rowMonth == cacheObj.updateTargetMonth && tuitionKeys.indexOf(rowKey) > -1 && (!settled || settled === "")) {
            recordSheet.getRange(i + 1, 10).setValue(cacheObj.updateTargetMonth); // J欄：學費結算日期
          }
        }
      }
      
      // 預排紀錄學費回溯標記
      const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
      if (planSheet) {
        const pData = planSheet.getDataRange().getValues();
        for (let i = 1; i < pData.length; i++) {
          const lessonDateMonth = (pData[i][2] instanceof Date) ? Utilities.formatDate(pData[i][2], timeZone, "yyyy/MM") : String(pData[i][2]).substring(0, 7);
          const status = pData[i][9];
          const feeSettled = pData[i][10];
          const rowKey = buildTuitionSelectionKey(pData[i][7], pData[i][8]);
          
          if (status !== "取消" && lessonDateMonth == cacheObj.updateTargetMonth && tuitionKeys.indexOf(rowKey) > -1 && (!feeSettled || feeSettled === "")) {
            planSheet.getRange(i + 1, 11).setValue(cacheObj.updateTargetMonth); // K欄：預排學費結算
          }
        }
      }
    }

    CacheService.getScriptCache().remove(cacheKey);
    replyLineMessage(replyToken, "🎉 結算資料已成功寫入，並自動回溯標記完成！");
  } catch (e) {
    replyLineMessage(replyToken, "❌ 寫入衝突或失敗，請稍候重試：" + e.toString());
  } finally {
    lock.releaseLock();
  }
}

function getFinancialInvoiceColumnIndex(category: string): number {
  if (category === "學費") return 9; // J欄：收據/單據編號
  if (category === "鐘點費") return 7; // H欄：領據編號
  return 1;
}

function buildFinancialInvoicePrefix(category: string, month: string, fallbackPrefix: string): string {
  const normalizedMonth = normalizeFinancialMonth(month, Session.getScriptTimeZone());
  const parts = normalizedMonth.split("/");
  const year = parts[0] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  const monthPart = (parts[1] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM")).padStart(2, "0");
  const prefix = category === "學費" ? "R" : (category === "鐘點費" ? "A" : (fallbackPrefix || ""));
  return prefix + "_" + year + "_" + monthPart + "_";
}

function parseFinancialInvoiceSerial(invoiceId: string, invoicePrefix: string): number {
  const serialText = invoiceId.substring(invoicePrefix.length);
  const serial = parseInt(serialText, 10);
  return isNaN(serial) ? 0 : serial;
}

function getFinancialDuplicateKeys(rows: any[], category: string): string[] {
  const keys: string[] = [];
  for (let i = 0; i < rows.length; i++) {
    const key = getFinancialRowKey(rows[i], category);
    if (key && keys.indexOf(key) === -1) keys.push(key);
  }
  return keys;
}

function getFinancialRowKey(row: any[], category: string): string {
  if (category === "學費") return buildTuitionSelectionKey(row[1], row[2]); // B/C欄：學生姓名 + 課程名稱
  if (category === "鐘點費") return String(row[1] || "").trim(); // B欄：講師姓名
  return "";
}

function buildTuitionSelectionKey(studentName: any, courseName: any): string {
  return String(studentName || "").trim() + "::" + String(courseName || "").trim();
}

function getTuitionKeysFromSavedRows(rows: any[][]): string[] {
  const keys: string[] = [];
  for (let i = 0; i < rows.length; i++) {
    const key = buildTuitionSelectionKey(rows[i][1], rows[i][2]);
    if (key !== "::" && keys.indexOf(key) === -1) keys.push(key);
  }
  return keys;
}

function normalizeFinancialMonth(value: any, timeZone: string): string {
  if (value instanceof Date) return Utilities.formatDate(value, timeZone, "yyyy/MM");
  const text = String(value || "").trim();
  if (text.match(/^\d{4}[\/-]\d{2}/)) return text.substring(0, 7).replace("-", "/");
  return text;
}

// 記帳、財務三表與會員查詢模組宣告與引進
declare function lookupGeneralMemberData(name: string, category: string): any;
declare function handleFinancialReportCommand(event: any, userMsg: string, type: string): void;
declare function handleManualJournalEntryCommand(event: any, userMsg: string): void;
declare function executeManualJournalSave(event: any, data: string): void;
declare function handlePaymentDocCommand(event: any, userMsg: string): void;
declare function createPaymentNoticesBatch(targetMonth: string, targetName: string): string;
declare function replyLineMessage(token: string, msg: string): void;
declare function replyConfirmationCard(token: string, title: string, report: string, key: string): void;

