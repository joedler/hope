// ==========================================
// 🧱 Core_Service.ts : 營運核心 (授課登記、預排核銷與 LIFF API 模組)
// ==========================================

// ==========================================
// 1. LIFF 專屬 API 後端實作模組
// ==========================================

// (1) handleLiffMe : 讀取講師個人狀態與權限
function handleLiffMe(lineUserId: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return { ok: false, error: "SYSTEM_ERROR", message: "找不到講師名單分頁。" };

  const teacherData = teacherSheet.getDataRange().getValues();
  let teacherRowIndex = -1;
  let teacherName = "";

  for (let i = 1; i < teacherData.length; i++) {
    if (String(teacherData[i][1]).trim() === lineUserId) {
      teacherRowIndex = i;
      teacherName = teacherData[i][0];
      break;
    }
  }

  if (teacherRowIndex === -1) {
    return { ok: false, error: "NOT_BOUND", message: "此 LINE 帳號尚未綁定講師身分。" };
  }

  // 統計本月已授課時數
  const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  let thisMonthHours = 0;
  if (recordSheet) {
    const recordData = recordSheet.getDataRange().getValues();
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth(); // 0-11

    for (let j = 1; j < recordData.length; j++) {
      const rowTeacher = recordData[j][1];
      const rowDate = recordData[j][2];
      const rowHours = parseFloat(recordData[j][5]) || 0;

      if (rowTeacher === teacherName && rowDate instanceof Date) {
        if (rowDate.getFullYear() === currentYear && rowDate.getMonth() === currentMonth) {
          thisMonthHours += rowHours;
        }
      }
    }
  }

  // 讀取請假額度 (自 Leave_Service 讀取，這裡提供安全預設或呼叫)
  let casualLeaveRemaining = 8;
  let sickLeaveRemaining = 24;
  try {
    const leaveQuota = getTeacherLeaveQuota(teacherName);
    casualLeaveRemaining = leaveQuota.casual;
    sickLeaveRemaining = leaveQuota.sick;
  } catch (e) {}

  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;

  return {
    ok: true,
    profile: {
      name: teacherName,
      thisMonthHours: Math.round(thisMonthHours * 100) / 100,
      casualLeaveRemaining: casualLeaveRemaining,
      sickLeaveRemaining: sickLeaveRemaining,
      isAdmin: isAdmin
    }
  };
}

// (2) handleLiffFormOptions : 讀取登記表單所需的下拉選項
function handleLiffFormOptions() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
    if (!courseSheet) return { ok: false, message: "找不到課程設定分頁。" };

    const data = courseSheet.getDataRange().getValues();
    const studentsSet = new Set<string>();
    const subjectsSet = new Set<string>();

    for (let i = 1; i < data.length; i++) {
      const studentName = String(data[i][2]).trim();
      const subjectName = String(data[i][3]).trim();
      if (studentName) studentsSet.add(studentName);
      if (subjectName) subjectsSet.add(subjectName);
    }

    return {
      ok: true,
      options: {
        students: Array.from(studentsSet),
        subjects: Array.from(subjectsSet)
      }
    };
  } catch (e) {
    return { ok: false, message: "載入下拉選項出錯：" + e.toString() };
  }
}

// (3) handleLiffRegister : LIFF 提交登記課程或預排
function handleLiffRegister(params: any) {
  const lineUserId = params.lineUserId;
  const mode = params.mode; // 'normal' (正式授課) | 'pre' (預排)
  const studentName = params.student;
  const subjectName = params.subject;
  const dateStr = params.date; // YYYY-MM-DD
  const startTimeStr = params.startTime; // HH:MM
  const endTimeStr = params.endTime; // HH:MM

  const isPlan = (mode === 'pre');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const targetSheetName = isPlan ? SHEET_NAME_PLAN : SHEET_NAME_RECORD;
  const recordSheet = ss.getSheetByName(targetSheetName);
  if (!recordSheet) return { ok: false, message: `找不到寫入分頁: ${targetSheetName}` };

  // 1. 取得講師姓名
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  const teacherData = teacherSheet.getDataRange().getValues();
  let userName = "";
  for (let i = 1; i < teacherData.length; i++) {
    if (String(teacherData[i][1]).trim() === lineUserId) {
      userName = teacherData[i][0];
      break;
    }
  }
  if (!userName) return { ok: false, message: "身分驗證失敗，找不到講師資料。" };

  // 2. 獲取單價
  const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  const courseData = courseSheet.getDataRange().getValues();
  let unitFee = 0;
  for (let i = 1; i < courseData.length; i++) {
    if (courseData[i][0] === userName && courseData[i][2] === studentName && courseData[i][3] === subjectName) {
      unitFee = parseFloat(courseData[i][4]) || 0;
      break;
    }
  }

  // 3. 計算時數與金額
  const formattedDateStr = dateStr.replace(/-/g, "/"); // YYYY/MM/DD
  const cleanStart = startTimeStr.replace(/:/g, ""); // HHMM
  const cleanEnd = endTimeStr.replace(/:/g, ""); // HHMM

  const startNum = parseTime(cleanStart);
  const endNum = parseTime(cleanEnd);
  const duration = Math.round((endNum - startNum) * 100) / 100;
  const totalPay = Math.round(duration * unitFee);

  if (duration <= 0) return { ok: false, message: "上課結束時間必須晚於開始時間。" };

  // 4. 重疊防護檢查
  const timeZone = Session.getScriptTimeZone();
  const inputDate = new Date(formattedDateStr);
  const inputFingerprint = Utilities.formatDate(inputDate, timeZone, "yyyyMMdd");

  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  const recordHistorySheet = ss.getSheetByName(SHEET_NAME_RECORD);

  const planHistory = planSheet ? planSheet.getDataRange().getValues() : [];
  const recordHistory = recordHistorySheet ? recordHistorySheet.getDataRange().getValues() : [];

  let isOverlap = checkGlobalOverlap(planHistory, userName, inputFingerprint, startNum, endNum, true, timeZone, subjectName);
  if (!isOverlap) {
    isOverlap = checkGlobalOverlap(recordHistory, userName, inputFingerprint, startNum, endNum, false, timeZone, subjectName);
  }

  if (isOverlap) {
    return { ok: false, message: `🚫 登記失敗：該時段 (${startTimeStr} - ${endTimeStr}) 已有其他授課安排！` };
  }

  // 5. 寫入資料
  const now = new Date();
  const writeDateStr = Utilities.formatDate(inputDate, timeZone, "yyyy/MM/dd");
  const rowData: any[] = [now, userName, writeDateStr, cleanStart, cleanEnd, duration, totalPay, studentName, subjectName];

  if (isPlan) {
    rowData.push("未核銷");
    rowData.push(""); // 學費結算欄留空
    rowData.push(""); // 退費結算欄留空
  } else {
    rowData.push(""); // 學費結算欄留空
    rowData.push(""); // 鐘點結算欄留空
  }
  rowData.push(userName); // 操作人為講師自己

  recordSheet.appendRow(rowData);

  return { ok: true, message: "登記成功！" };
}

// (4) handleLiffGetUnverified : 取得該講師待核銷的預排紀錄
function handleLiffGetUnverified(lineUserId: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 1. 取得講師姓名
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  const teacherData = teacherSheet.getDataRange().getValues();
  let userName = "";
  for (let i = 1; i < teacherData.length; i++) {
    if (String(teacherData[i][1]).trim() === lineUserId) {
      userName = teacherData[i][0];
      break;
    }
  }
  if (!userName) return { ok: false, message: "身分驗證失敗。" };

  // 2. 搜尋待核銷紀錄
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  if (!planSheet) return { ok: true, schedules: [] };

  const data = planSheet.getDataRange().getValues();
  const schedules: any[] = [];
  const timeZone = Session.getScriptTimeZone();

  for (let i = 1; i < data.length; i++) {
    const rowTeacher = data[i][1];
    const rowDate = data[i][2];
    const rowStart = data[i][3];
    const rowEnd = data[i][4];
    const rowHours = data[i][5];
    const rowStudent = data[i][7];
    const rowSubject = data[i][8];
    const rowStatus = String(data[i][9]).trim();

    if (rowTeacher === userName && rowStatus === "未核銷") {
      const dateFormatted = rowDate instanceof Date ? Utilities.formatDate(rowDate, timeZone, "yyyy/MM/dd") : rowDate;
      schedules.push({
        rowId: i + 1, // 記住試算表中的實際 Row 行號，以供核銷定位
        student: rowStudent,
        subject: rowSubject,
        date: dateFormatted,
        startTime: formatTimeStr(rowStart),
        endTime: formatTimeStr(rowEnd),
        hours: rowHours
      });
    }
  }

  return { ok: true, schedules: schedules };
}

// (5) handleLiffVerifySchedule : 執行預排核銷
function handleLiffVerifySchedule(params: any) {
  const lineUserId = params.lineUserId;
  const rowId = parseInt(params.rowId);
  const isVerified = params.isVerified; // 'yes' (有上課) | 'no' (未上課)

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  if (!planSheet) return { ok: false, message: "找不到預排工作表。" };

  // 1. 驗證權限：確認操作者是否擁有這筆資料或是 Admin
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  const teacherData = teacherSheet.getDataRange().getValues();
  let userName = "";
  for (let i = 1; i < teacherData.length; i++) {
    if (String(teacherData[i][1]).trim() === lineUserId) {
      userName = teacherData[i][0];
      break;
    }
  }
  if (!userName) return { ok: false, message: "身分驗證失敗。" };

  const recordTeacher = planSheet.getRange(rowId, 2).getValue();
  const isAdmin = ADMIN_LIST.indexOf(lineUserId) > -1;

  if (recordTeacher !== userName && !isAdmin) {
    return { ok: false, message: "無權限操作此預排紀錄。" };
  }

  // 2. 進行核銷
  const timeZone = Session.getScriptTimeZone();
  if (isVerified === "yes") {
    // 標記為「已實際上課」
    planSheet.getRange(rowId, 10).setValue("已實際上課");

    // 將資料複製並寫入「授課紀錄」工作表
    const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
    if (recordSheet) {
      const rowValues = planSheet.getRange(rowId, 1, 1, 13).getValues()[0];
      
      const now = new Date();
      const dateFormatted = rowValues[2] instanceof Date ? Utilities.formatDate(rowValues[2], timeZone, "yyyy/MM/dd") : rowValues[2];

      const newRow = [
        now,
        rowValues[1], // 講師
        dateFormatted, // 日期
        rowValues[3], // 開始
        rowValues[4], // 結束
        rowValues[5], // 時數
        rowValues[6], // 金額
        rowValues[7], // 學生
        rowValues[8], // 課程
        "",           // 學費結算
        "",           // 鐘點結算
        `核銷人:${userName}` // 操作人
      ];
      recordSheet.appendRow(newRow);
    }
  } else {
    // 標記為「取消」
    planSheet.getRange(rowId, 10).setValue("取消");
  }

  return { ok: true };
}

// (6) handleLiffUnbind : 解除 LINE 綁定
function handleLiffUnbind(lineUserId: string) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
    if (!teacherSheet) return { ok: false, message: "找不到講師名單工作表。" };

    const data = teacherSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === lineUserId) {
        // 清空 LINE 使用者 ID 欄位 (第二欄，即 B 欄)
        teacherSheet.getRange(i + 1, 2).setValue("");
        return { ok: true, message: "已解除綁定！" };
      }
    }
    return { ok: false, message: "找不到對應的綁定資料。" };
  } catch (e) {
    return { ok: false, message: "解綁失敗：" + e.toString() };
  }
}

// (7) handleLiffVerifyAndBind : 講師驗證姓名並綁定 LINE
function handleLiffVerifyAndBind(name: string, lineUserId: string) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
    if (!teacherSheet) return { ok: false, message: "找不到講師名單工作表。" };

    const data = teacherSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowName = String(data[i][0]).trim();
      const rowLineId = String(data[i][1]).trim();

      if (rowName === name) {
        if (rowLineId && rowLineId !== lineUserId) {
          return { ok: false, message: `此姓名已被其他 LINE 帳號綁定。若有疑問請聯絡管理員。` };
        }
        // 寫入綁定 ID
        teacherSheet.getRange(i + 1, 2).setValue(lineUserId);
        
        // 自動載入綁定成功後的 profile
        return handleLiffMe(lineUserId);
      }
    }
    return { ok: false, message: `驗證失敗：找不到姓名為「${name}」的講師，請確認輸入是否正確。` };
  } catch (e) {
    return { ok: false, message: "綁定失敗：" + e.toString() };
  }
}

// ==========================================
// 2. 原 LINE 訊息對話控制重構模組
// ==========================================

function processFlowReport(event: any, userMsg: string, cachedJson: string) {
  const replyToken = event.replyToken;
  const operatorId = event.source.userId;
  
  const state = JSON.parse(cachedJson);
  const targetUserId = state.targetId;
  const studentName = state.sName;
  const courseName = state.sCourse;
  const isProxy = state.isProxy;
  const mode = state.mode;
  const isPlan = (mode === 'plan');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const targetSheetName = isPlan ? SHEET_NAME_PLAN : SHEET_NAME_RECORD;
  let recordSheet = ss.getSheetByName(targetSheetName);
  const mainRecordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);

  if (!recordSheet && isPlan) {
    recordSheet = ss.insertSheet(SHEET_NAME_PLAN);
    recordSheet.appendRow(["時間", "講師", "日期", "開始", "結束", "時數", "金額", "學生", "課程", "狀態", "學費結算", "退費結算", "操作人"]);
  }

  const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  const userData = userSheet.getDataRange().getValues();
  
  let userName = "";
  let operatorName = "未知操作者";
  
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][1] === targetUserId) {
      userName = userData[i][0];
    }
    if (userData[i][1] === operatorId) {
      operatorName = userData[i][0];
    }
  }
  
  if (operatorName === "未知操作者") {
      operatorName = "ID:" + operatorId;
  }
  
  if (userName === "") {
    replyLineMessage(replyToken, "❌ 找不到目標講師資料。");
    CacheService.getScriptCache().remove(operatorId);
    return;
  }

  const studentSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  const stuData = studentSheet.getDataRange().getValues();
  let unitFee = 0;
  for (let i = 1; i < stuData.length; i++) {
    if (stuData[i][0] === targetUserId && stuData[i][2] === studentName && stuData[i][3] === courseName) {
      unitFee = stuData[i][4];
      break;
    }
  }
  if (!unitFee || isNaN(unitFee)) unitFee = 0;

  const lines = userMsg.split("\n");
  const planHistory = planSheet ? planSheet.getDataRange().getValues() : [];
  const recordHistory = mainRecordSheet ? mainRecordSheet.getDataRange().getValues() : [];
  const now = new Date();
  const currentYear = now.getFullYear();
  const todayDate = new Date();
  todayDate.setHours(0, 0, 0, 0);
  const timeZone = Session.getScriptTimeZone();

  let successCount = 0;
  let failCount = 0;
  let overlapCount = 0;
  let reportDetail = "";
  
  for (let k = 0; k < lines.length; k++) {
    const line = lines[k].trim();
    if (line === "") continue;

    const cleanMsg = line.replace(/:/g, "").trim();
    const parts = cleanMsg.split(/[ ]+/);
    if (parts.length < 3) {
      failCount++;
      reportDetail += "❌ 格式短：" + line + "\n";
      continue;
    }

    const dateStr = parts[0];
    const startStr = parts[1];
    const endStr = parts[2];
    if (!dateStr.match(/^(\d{4}\/)?\d{1,2}\/\d{1,2}$/)) {
      failCount++;
      reportDetail += "❌ 日期錯 (" + dateStr + ")\n";
      continue;
    }
    if (!startStr.match(/^\d{4}$/) || !endStr.match(/^\d{4}$/)) {
      failCount++;
      reportDetail += "❌ 時間錯\n";
      continue;
    }

    try {
      const monthPart = parseInt(dateStr.split("/")[0]);
      const dayPart = parseInt(dateStr.split("/")[1]);
      let inputDate = new Date(currentYear, monthPart - 1, dayPart);
      const diffTime = inputDate.getTime() - todayDate.getTime();
      const diffDays = diffTime / (1000 * 3600 * 24);
      if (diffDays > 180) {
        inputDate = new Date(currentYear - 1, monthPart - 1, dayPart);
      } else if (diffDays < -180) {
        inputDate = new Date(currentYear + 1, monthPart - 1, dayPart);
      }
      inputDate.setHours(0, 0, 0, 0);
      if (dateStr.match(/^\d{4}\//)) {
        inputDate = new Date(dateStr);
        inputDate.setHours(0, 0, 0, 0);
      }

      if (!isPlan && inputDate > todayDate) {
        failCount++;
        const dLog = Utilities.formatDate(inputDate, timeZone, "yyyy/MM/dd");
        reportDetail += "❌ 未來日期禁止登記：" + dLog + "\n";
        continue;
      }

      const startNum = parseTime(startStr);
      const endNum = parseTime(endStr);
      const duration = Math.round((endNum - startNum) * 100) / 100;
      const totalPay = Math.round(duration * unitFee);
      if (duration <= 0) {
        failCount++;
        reportDetail += "❌ 時間錯 (負數)\n";
        continue;
      }

      const inputFingerprint = Utilities.formatDate(inputDate, timeZone, "yyyyMMdd");
      
      let isOverlap = false;
      if (checkGlobalOverlap(planHistory, userName, inputFingerprint, startNum, endNum, true, timeZone, courseName)) {
        isOverlap = true;
      }
      if (!isOverlap && checkGlobalOverlap(recordHistory, userName, inputFingerprint, startNum, endNum, false, timeZone, courseName)) {
        isOverlap = true;
      }

      if (isOverlap) {
        overlapCount++;
        reportDetail += "🚫 重疊擋下：" + dateStr + " " + startStr + "-" + endStr + "\n";
        continue;
      }

      const writeDateStr = Utilities.formatDate(inputDate, timeZone, "yyyy/MM/dd");
      
      const rowData = [now, userName, writeDateStr, startStr, endStr, duration, totalPay, studentName, courseName];
      if (isPlan) {
        rowData.push("未核銷");
        rowData.push(""); // 學費結算欄留空
        rowData.push(""); // 退費結算欄留空
      } else {
        rowData.push(""); // 學費結算欄留空
        rowData.push(""); // 鐘點結算欄留空
      }
      rowData.push(operatorName);

      recordSheet.appendRow(rowData);
      
      if (isPlan) {
        planHistory.push(rowData);
      } else {
        recordHistory.push(rowData);
      }

      successCount++;
      reportDetail += "✅ " + writeDateStr + " | " + startStr + "-" + endStr + " (" + duration + "hr) $" + totalPay + "\n";
    } catch (e) {
      failCount++;
      reportDetail += "❌ 失敗\n";
    }
  }

  if (successCount > 0) {
    CacheService.getScriptCache().remove(operatorId);
    const modeTitle = isPlan ? (isProxy ? "📅 [代排成功]" : "📅 [預排成功]") : (isProxy ? "📝 [代登成功]" : "✅ [登記成功]");
    const header = "👤 " + userName + " 講師\n🎓 " + studentName + " (" + courseName + ")\n👨‍💻 操作者：" + operatorName + "\n------------------\n";
    replyLineMessage(replyToken, modeTitle + "\n" + header + reportDetail + "\n(成功:" + successCount + " / 重疊:" + overlapCount + " / 失敗:" + failCount + ")");
  } else {
    let warning = "❌ 處理失敗，請修正後重新輸入：\n\n" + reportDetail;
    if (failCount > 0) warning += "\n💡 請用格式：日期 開始 結束";
    replyLineMessage(replyToken, warning);
  }
}

function checkGlobalOverlap(
  data: any[],
  userName: string,
  inputFingerprint: string,
  startNum: number,
  endNum: number,
  isPlanSheet: boolean,
  timeZone: string,
  inputCourse?: string
): boolean {
  if (inputCourse) {
    for (let m = 0; m < MAGIC_KEYWORDS.length; m++) {
      if (inputCourse.indexOf(MAGIC_KEYWORDS[m]) > -1) return false;
    }
  }

  const cleanUser = String(userName).trim();

  for (let h = 1; h < data.length; h++) {
    if (String(data[h][1]).trim() !== cleanUser) continue;
    const rowDateRaw = data[h][2];
    let rowFingerprint = "";
    if (rowDateRaw instanceof Date) {
      rowFingerprint = Utilities.formatDate(rowDateRaw, timeZone, "yyyyMMdd");
    } else {
      try {
        const parsedDate = new Date(rowDateRaw);
        if (!isNaN(parsedDate.getTime())) {
          rowFingerprint = Utilities.formatDate(parsedDate, timeZone, "yyyyMMdd");
        } else {
          continue;
        }
      } catch (e) {
        continue;
      }
    }

    if (rowFingerprint !== inputFingerprint) continue;
    if (isPlanSheet && data[h].length > 9) {
      if (String(data[h][9]).trim() === "取消") continue;
    }

    const rowCourse = String(data[h][8] || "");
    let hasMagic = false;
    for (let m = 0; m < MAGIC_KEYWORDS.length; m++) {
      if (rowCourse.indexOf(MAGIC_KEYWORDS[m]) > -1) {
        hasMagic = true;
        break;
      }
    }
    if (hasMagic) continue;

    const rowStart = parseTime(data[h][3].toString());
    const rowEnd = parseTime(data[h][4].toString());
    if (Math.max(startNum, rowStart) < Math.min(endNum, rowEnd)) {
      return true;
    }
  }
  return false;
}

// 輔助函式：時間解析與格式化
function parseTime(timeStr: string): number {
  let val = timeStr.toString().trim();
  if (val.length === 3) {
    val = "0" + val;
  }
  if (val.length !== 4) return 0;
  const hr = parseInt(val.substring(0, 2));
  const min = parseInt(val.substring(2, 4));
  return hr + min / 60;
}

function formatTimeStr(timeVal: any): string {
  const val = String(timeVal).trim();
  if (val.length === 3) {
    return "0" + val.substring(0, 1) + ":" + val.substring(1, 3);
  }
  if (val.length === 4) {
    return val.substring(0, 2) + ":" + val.substring(2, 4);
  }
  return val;
}

// 宣告外部參考
declare function replyLineMessage(token: string, msg: string): void;
declare function getTeacherLeaveQuota(teacherName: string): { casual: number; sick: number };

