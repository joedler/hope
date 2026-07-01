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

  const isAdmin = isAdminLineUser(lineUserId);

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
function handleLiffFormOptions(lineUserId?: string) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
    if (!courseSheet) return { ok: false, message: "找不到課程設定分頁。" };

    let teacherName = "";
    if (lineUserId) {
      const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
      const teacherData = teacherSheet ? teacherSheet.getDataRange().getValues() : [];
      for (let i = 1; i < teacherData.length; i++) {
        if (String(teacherData[i][1]).trim() === String(lineUserId).trim()) {
          teacherName = String(teacherData[i][0]).trim();
          break;
        }
      }
    }

    const data = courseSheet.getDataRange().getValues();
    const studentsSet = new Set<string>();
    const subjectsSet = new Set<string>();
    const coursesByStudent: any = {};

    for (let i = 1; i < data.length; i++) {
      const rowTeacher = String(data[i][0]).trim();
      if (lineUserId && !isCourseOwner(rowTeacher, lineUserId, teacherName)) continue;
      const studentName = String(data[i][2]).trim();
      const subjectName = String(data[i][3]).trim();
      if (studentName) studentsSet.add(studentName);
      if (subjectName) subjectsSet.add(subjectName);
      if (studentName && subjectName) {
        if (!coursesByStudent[studentName]) coursesByStudent[studentName] = [];
        if (coursesByStudent[studentName].indexOf(subjectName) === -1) {
          coursesByStudent[studentName].push(subjectName);
        }
      }
    }

    const students = Array.from(studentsSet);
    const subjects = Array.from(subjectsSet);
    const isAdmin = isAdminLineUser(lineUserId || "");
    return {
      ok: true,
      message: !isAdmin && (students.length === 0 || subjects.length === 0) ? "目前沒有可登記的學生或課程，請聯絡行政確認課程設定表。" : "",
      options: {
        students,
        subjects,
        coursesByStudent
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
    const rowOwner = String(courseData[i][0]).trim();
    const rowStudent = String(courseData[i][2]).trim();
    const rowSubject = String(courseData[i][3]).trim();
    if (isCourseOwner(rowOwner, lineUserId, userName) && rowStudent === String(studentName).trim() && rowSubject === String(subjectName).trim()) {
      unitFee = parseFloat(courseData[i][4]) || 0;
      break;
    }
  }
  if (unitFee <= 0) {
    return { ok: false, message: `找不到「${userName} / ${studentName} / ${subjectName}」的課程設定或鐘點單價，請聯絡行政確認課程設定表。` };
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
  if (duration > 4) return { ok: false, message: "課程時數超過 4 小時，請確認是否填錯。" };

  // 4. 重疊防護檢查
  const timeZone = Session.getScriptTimeZone();
  const inputDate = new Date(formattedDateStr);
  const inputFingerprint = Utilities.formatDate(inputDate, timeZone, "yyyyMMdd");
  const lessonStartDate = buildLessonDateTime(formattedDateStr, cleanStart, timeZone);
  const lessonEndDate = buildLessonDateTime(formattedDateStr, cleanEnd, timeZone);
  const currentTime = new Date();
  if (!lessonStartDate || !lessonEndDate) return { ok: false, message: "日期或時間格式錯誤。" };
  if (!isPlan && lessonEndDate.getTime() > currentTime.getTime()) {
    return { ok: false, message: "授課登記限已完成課程，不能登錄未來時間。" };
  }
  if (isPlan && lessonStartDate.getTime() < currentTime.getTime()) {
    return { ok: false, message: "預排課程不能選擇已過去的時間。" };
  }

  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  const recordHistorySheet = ss.getSheetByName(SHEET_NAME_RECORD);

  const planHistory = planSheet ? planSheet.getDataRange().getValues() : [];
  const recordHistory = recordHistorySheet ? recordHistorySheet.getDataRange().getValues() : [];

  const conflicts = findScheduleConflicts(
    planHistory.concat(recordHistory),
    {
      teacher: userName,
      student: studentName,
      subject: subjectName,
      dateFingerprint: inputFingerprint,
      start: startNum,
      end: endNum
    },
    timeZone
  );

  if (conflicts.length > 0) {
    return { ok: false, message: "🚫 登記失敗：\n" + conflicts.join("\n") };
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
    const now = new Date();

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
      const lessonEnd = buildLessonEndDate(rowDate, rowEnd, timeZone);
      const canVerify = !lessonEnd || lessonEnd.getTime() <= now.getTime();
      const dateFormatted = rowDate instanceof Date ? Utilities.formatDate(rowDate, timeZone, "yyyy/MM/dd") : rowDate;
      schedules.push({
        rowId: i + 1, // 記住試算表中的實際 Row 行號，以供核銷定位
        student: rowStudent,
        subject: rowSubject,
        date: dateFormatted,
        startTime: formatTimeStr(rowStart),
        endTime: formatTimeStr(rowEnd),
        hours: rowHours,
        canVerify,
        verifyMessage: canVerify ? "" : "課程尚未結束，暫不可核銷。"
      });
    }
  }

  return { ok: true, schedules: schedules };
}

// (4b) handleLiffGetRecentRegistered : 取得該講師本月最近已登記授課
function handleLiffGetRecentRegistered(lineUserId: string, limit?: any) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const teacherName = getTeacherNameByLineUserId(lineUserId);
  if (!teacherName) return { ok: false, message: "身分驗證失敗。" };

  const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  if (!recordSheet) return { ok: true, lessons: [] };

  const maxRows = Math.max(1, Math.min(parseInt(limit || "5", 10) || 5, 20));
  const data = recordSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const now = new Date();
  const currentMonth = Utilities.formatDate(now, timeZone, "yyyyMM");
  const lessons: any[] = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const rowTeacher = String(data[i][1] || "").trim();
    if (rowTeacher !== teacherName) continue;

    const rowDate = data[i][2];
    const rowMonth = getDateFingerprint(rowDate, timeZone).substring(0, 6);
    if (rowMonth !== currentMonth) continue;

    lessons.push({
      rowId: i + 1,
      date: rowDate instanceof Date ? Utilities.formatDate(rowDate, timeZone, "yyyy/MM/dd") : String(rowDate || ""),
      startTime: formatTimeStr(data[i][3]),
      endTime: formatTimeStr(data[i][4]),
      hours: parseFloat(data[i][5]) || 0,
      student: String(data[i][7] || "").trim(),
      subject: String(data[i][8] || "").trim()
    });

    if (lessons.length >= maxRows) break;
  }

  return { ok: true, lessons };
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
  const recordStatus = String(planSheet.getRange(rowId, 10).getValue()).trim();
  const isAdmin = isAdminLineUser(lineUserId);

  if (recordTeacher !== userName && !isAdmin) {
    return { ok: false, message: "無權限操作此預排紀錄。" };
  }
  if (recordStatus !== "未核銷") {
    return { ok: false, message: `此預排紀錄目前狀態為「${recordStatus || "空白"}」，不可重複核銷。` };
  }
  const recordDate = planSheet.getRange(rowId, 3).getValue();
  const recordEnd = planSheet.getRange(rowId, 5).getValue();
  const lessonEnd = buildLessonEndDate(recordDate, recordEnd, Session.getScriptTimeZone());
  if (lessonEnd && lessonEnd.getTime() > new Date().getTime()) {
    return { ok: false, message: "此預排課程尚未到下課時間，不可提前核銷。" };
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
  planSheet.getRange(rowId, 13).setValue(recordTeacher === userName ? userName : `代核銷:${userName}`);

  return { ok: true };
}

// (5b) 行政代操作：讀取講師清單、指定講師課程選項與待核銷清單
function handleLiffAdminCourseProxyOptions(params: any) {
  const operatorLineUserId = String(params.lineUserId || "").trim();
  if (!isAdminLineUser(operatorLineUserId)) {
    return { ok: false, message: "權限不足：限行政人員使用課程代操作。" };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return { ok: false, message: "找不到講師名單。" };

  const targetKey = String(params.targetTeacher || "").trim();
  const target = targetKey ? findTeacherForCourseProxy_(teacherSheet.getDataRange().getValues(), targetKey) : null;

  return {
    ok: true,
    teachers: listCourseProxyTeachers_(teacherSheet.getDataRange().getValues()),
    selectedTeacher: target,
    options: target ? getCourseOptionsForTeacher_(ss, target) : { students: [], subjects: [], coursesByStudent: {} },
    schedules: target ? getUnverifiedSchedulesForTeacher_(ss, target.name) : []
  };
}

// (5c) 行政代操作：預覽代登、代預排或代核銷
function handleLiffAdminCourseProxyPreview(params: any) {
  const operatorLineUserId = String(params.lineUserId || "").trim();
  if (!isAdminLineUser(operatorLineUserId)) {
    return { ok: false, message: "權限不足：限行政人員使用課程代操作。" };
  }
  const actionType = String(params.proxyAction || "").trim();
  if (actionType === "verify") return previewAdminProxyVerify_(params, operatorLineUserId);
  if (actionType === "normal" || actionType === "pre") return previewAdminProxyRegister_(params, operatorLineUserId);
  return { ok: false, message: "不支援的代操作類型。" };
}

// (5d) 行政代操作：確認執行
function handleLiffAdminCourseProxyConfirm(params: any) {
  const operatorLineUserId = String(params.lineUserId || "").trim();
  if (!isAdminLineUser(operatorLineUserId)) {
    return { ok: false, message: "權限不足：限行政人員使用課程代操作。" };
  }
  const actionType = String(params.proxyAction || "").trim();
  if (actionType === "verify") return confirmAdminProxyVerify_(params, operatorLineUserId);
  if (actionType === "normal" || actionType === "pre") return confirmAdminProxyRegister_(params, operatorLineUserId);
  return { ok: false, message: "不支援的代操作類型。" };
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
    const cleanName = String(name || "").trim();
    const cleanLineUserId = String(lineUserId || "").trim();
    if (!cleanName) return { ok: false, message: "請輸入姓名。" };
    if (!cleanLineUserId) return { ok: false, message: "缺少 LINE User ID，請從 LINE LIFF 入口重新開啟。" };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
    if (!teacherSheet) return { ok: false, message: "找不到講師名單工作表。" };

    const data = teacherSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowName = String(data[i][0]).trim();
      const rowLineId = String(data[i][1]).trim();

      if (rowName === cleanName) {
        if (rowLineId && rowLineId !== cleanLineUserId) {
          return { ok: false, message: `此姓名已被其他 LINE 帳號綁定。若有疑問請聯絡管理員。` };
        }
        // 寫入綁定 ID
        teacherSheet.getRange(i + 1, 2).setValue(cleanLineUserId);
        
        // 自動載入綁定成功後的 profile
        return handleLiffMe(cleanLineUserId);
      }
    }
    return { ok: false, message: `驗證失敗：找不到姓名為「${cleanName}」的講師，請確認輸入是否正確。` };
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

function findScheduleConflicts(
  data: any[],
  input: { teacher: string; student: string; subject: string; dateFingerprint: string; start: number; end: number },
  timeZone: string
): string[] {
  const conflicts = new Set<string>();
  if (hasMagicCourse(input.subject)) return [];

  const cleanTeacher = String(input.teacher).trim();
  const cleanStudent = String(input.student).trim();
  const cleanSubject = String(input.subject).trim();

  for (let h = 1; h < data.length; h++) {
    const rowStatus = String(data[h][9] || "").trim();
    if (rowStatus === "取消") continue;

    const rowSubject = String(data[h][8] || "").trim();
    if (hasMagicCourse(rowSubject)) continue;

    const rowFingerprint = getDateFingerprint(data[h][2], timeZone);
    if (rowFingerprint !== input.dateFingerprint) continue;

    const rowStart = parseTime(String(data[h][3] || ""));
    const rowEnd = parseTime(String(data[h][4] || ""));
    const isTimeOverlap = Math.max(input.start, rowStart) < Math.min(input.end, rowEnd);
    if (!isTimeOverlap) continue;

    const rowTeacher = String(data[h][1] || "").trim();
    const rowStudent = String(data[h][7] || "").trim();
    const rowCourse = String(data[h][8] || "").trim();

    if (rowTeacher === cleanTeacher) {
      conflicts.add(`同一講師在該時段已有課程：${rowStudent} / ${rowCourse}`);
    }
    if (rowStudent === cleanStudent) {
      conflicts.add(`同一學生在該時段已有課程：${rowTeacher} / ${rowCourse}`);
    }
    if (rowTeacher === cleanTeacher && rowStudent === cleanStudent && rowCourse === cleanSubject && rowStart === input.start && rowEnd === input.end) {
      conflicts.add("偵測到同講師、同學生、同課程、同時段的重複輸入。");
    }
  }

  return Array.from(conflicts);
}

function hasMagicCourse(courseName: string): boolean {
  const course = String(courseName || "");
  for (let m = 0; m < MAGIC_KEYWORDS.length; m++) {
    if (course.indexOf(MAGIC_KEYWORDS[m]) > -1) return true;
  }
  return false;
}

function isCourseOwner(rowOwner: string, lineUserId?: string, teacherName?: string): boolean {
  const owner = String(rowOwner || "").trim();
  const cleanLineUserId = String(lineUserId || "").trim();
  const cleanTeacherName = String(teacherName || "").trim();
  return owner === cleanLineUserId || owner === cleanTeacherName;
}

function getTeacherNameByLineUserId(lineUserId: string): string {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return "";
  const teacherData = teacherSheet.getDataRange().getValues();
  const cleanLineUserId = String(lineUserId || "").trim();
  for (let i = 1; i < teacherData.length; i++) {
    if (String(teacherData[i][1] || "").trim() === cleanLineUserId) {
      return String(teacherData[i][0] || "").trim();
    }
  }
  return "";
}

function buildLessonEndDate(dateValue: any, endTimeValue: any, timeZone: string): Date | null {
  return buildLessonDateTime(dateValue, endTimeValue, timeZone);
}

function buildLessonDateTime(dateValue: any, timeValue: any, timeZone: string): Date | null {
  const dateText = dateValue instanceof Date
    ? Utilities.formatDate(dateValue, timeZone, "yyyy/MM/dd")
    : String(dateValue || "").replace(/-/g, "/").substring(0, 10);
  const cleanTime = String(timeValue || "").replace(":", "").trim();
  const normalizedTime = cleanTime.length === 3 ? "0" + cleanTime : cleanTime;
  if (!dateText || normalizedTime.length !== 4) return null;

  const year = parseInt(dateText.substring(0, 4), 10);
  const month = parseInt(dateText.substring(5, 7), 10);
  const day = parseInt(dateText.substring(8, 10), 10);
  const hour = parseInt(normalizedTime.substring(0, 2), 10);
  const minute = parseInt(normalizedTime.substring(2, 4), 10);
  if ([year, month, day, hour, minute].some(function (value) { return isNaN(value); })) return null;
  return new Date(year, month - 1, day, hour, minute, 0, 0);
}

function getDateFingerprint(value: any, timeZone: string): string {
  if (value instanceof Date) {
    return Utilities.formatDate(value, timeZone, "yyyyMMdd");
  }
  try {
    const parsedDate = new Date(value);
    if (!isNaN(parsedDate.getTime())) {
      return Utilities.formatDate(parsedDate, timeZone, "yyyyMMdd");
    }
  } catch (e) {}
  return "";
}

function listCourseProxyTeachers_(teacherData: any[][]): any[] {
  const teachers: any[] = [];
  for (let i = 1; i < teacherData.length; i++) {
    const name = String(teacherData[i][0] || "").trim();
    const lineUserId = String(teacherData[i][1] || "").trim();
    const role = String(teacherData[i][9] || "").trim();
    const status = teacherData[i][10];
    if (!name || isInactiveStaffStatus(status)) continue;
    if (isAdminRoleValue(role)) continue;
    teachers.push({
      id: lineUserId || name,
      lineUserId,
      name
    });
  }
  return teachers;
}

function findTeacherForCourseProxy_(teacherData: any[][], targetKey: string): any {
  const cleanTarget = String(targetKey || "").trim();
  if (!cleanTarget) return null;
  for (let i = 1; i < teacherData.length; i++) {
    const name = String(teacherData[i][0] || "").trim();
    const lineUserId = String(teacherData[i][1] || "").trim();
    const status = teacherData[i][10];
    if (isInactiveStaffStatus(status)) continue;
    if (cleanTarget === lineUserId || cleanTarget === name) {
      return { id: lineUserId || name, lineUserId, name };
    }
  }
  return null;
}

function getCourseOptionsForTeacher_(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, teacher: any): any {
  const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  if (!courseSheet) return { students: [], subjects: [], coursesByStudent: {} };
  const data = courseSheet.getDataRange().getValues();
  const studentsSet = new Set<string>();
  const subjectsSet = new Set<string>();
  const coursesByStudent: any = {};

  for (let i = 1; i < data.length; i++) {
    const rowTeacher = String(data[i][0] || "").trim();
    if (!isCourseOwner(rowTeacher, teacher.lineUserId, teacher.name)) continue;
    const studentName = String(data[i][2] || "").trim();
    const subjectName = String(data[i][3] || "").trim();
    if (studentName) studentsSet.add(studentName);
    if (subjectName) subjectsSet.add(subjectName);
    if (studentName && subjectName) {
      if (!coursesByStudent[studentName]) coursesByStudent[studentName] = [];
      if (coursesByStudent[studentName].indexOf(subjectName) === -1) coursesByStudent[studentName].push(subjectName);
    }
  }

  return {
    students: Array.from(studentsSet),
    subjects: Array.from(subjectsSet),
    coursesByStudent
  };
}

function getUnverifiedSchedulesForTeacher_(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, teacherName: string): any[] {
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  if (!planSheet) return [];
  const data = planSheet.getDataRange().getValues();
  const schedules: any[] = [];
  const timeZone = Session.getScriptTimeZone();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const rowTeacher = String(data[i][1] || "").trim();
    const rowStatus = String(data[i][9] || "").trim();
    if (rowTeacher !== teacherName || rowStatus !== "未核銷") continue;
    const lessonEnd = buildLessonEndDate(data[i][2], data[i][4], timeZone);
    const canVerify = !lessonEnd || lessonEnd.getTime() <= now.getTime();
    schedules.push({
      rowId: i + 1,
      student: String(data[i][7] || "").trim(),
      subject: String(data[i][8] || "").trim(),
      date: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], timeZone, "yyyy/MM/dd") : String(data[i][2] || ""),
      startTime: formatTimeStr(data[i][3]),
      endTime: formatTimeStr(data[i][4]),
      hours: data[i][5],
      canVerify,
      verifyMessage: canVerify ? "" : "課程尚未結束，暫不可核銷。"
    });
  }
  return schedules;
}

function getCourseProxyTarget_(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, targetKey: string): any {
  const teacherSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return null;
  return findTeacherForCourseProxy_(teacherSheet.getDataRange().getValues(), targetKey);
}

function previewAdminProxyRegister_(params: any, operatorLineUserId: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const target = getCourseProxyTarget_(ss, String(params.targetTeacher || ""));
  if (!target) return { ok: false, message: "請先選擇要代操作的講師。" };
  const check = validateCourseProxyRegistration_(ss, params, target);
  if (!check.ok) return check;
  const operatorName = getTeacherNameByLineUserId(operatorLineUserId) || "行政人員";
  const modeLabel = String(params.proxyAction) === "pre" ? "代預排課程" : "代新增授課";
  return {
    ok: true,
    preview: {
      title: modeLabel,
      summary: `請確認是否由 ${operatorName} 幫 ${target.name} 講師執行${modeLabel}。`,
      items: [
        `講師：${target.name}`,
        `學生：${params.student}`,
        `課程：${params.subject}`,
        `時間：${String(params.date || "").replace(/-/g, "/")} ${params.startTime}-${params.endTime}`,
        `時數：${check.hours} 小時`,
        `金額：${check.totalPay}`
      ],
      nextAction: `確認${modeLabel}`,
      canConfirm: true,
      confirmAction: "adminCourseProxyConfirm"
    }
  };
}

function confirmAdminProxyRegister_(params: any, operatorLineUserId: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const target = getCourseProxyTarget_(ss, String(params.targetTeacher || ""));
  if (!target) return { ok: false, message: "請先選擇要代操作的講師。" };
  const check = validateCourseProxyRegistration_(ss, params, target);
  if (!check.ok) return check;

  const isPlan = String(params.proxyAction) === "pre";
  const sheet = ss.getSheetByName(isPlan ? SHEET_NAME_PLAN : SHEET_NAME_RECORD);
  if (!sheet) return { ok: false, message: "找不到寫入分頁。" };
  const timeZone = Session.getScriptTimeZone();
  const now = new Date();
  const operatorName = getTeacherNameByLineUserId(operatorLineUserId) || "行政人員";
  const rowData: any[] = [
    now,
    target.name,
    check.writeDate,
    check.cleanStart,
    check.cleanEnd,
    check.hours,
    check.totalPay,
    params.student,
    params.subject
  ];
  if (isPlan) {
    rowData.push("未核銷");
    rowData.push("");
    rowData.push("");
  } else {
    rowData.push("");
    rowData.push("");
  }
  rowData.push(`代操作:${operatorName}`);
  sheet.appendRow(rowData);
  return {
    ok: true,
    message: `已完成${isPlan ? "代預排課程" : "代新增授課"}：${target.name} / ${params.student} / ${params.subject} / ${Utilities.formatDate(new Date(check.writeDate), timeZone, "yyyy/MM/dd")} ${params.startTime}-${params.endTime}`
  };
}

function validateCourseProxyRegistration_(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, params: any, target: any) {
  const isPlan = String(params.proxyAction) === "pre";
  const studentName = String(params.student || "").trim();
  const subjectName = String(params.subject || "").trim();
  const dateStr = String(params.date || "").trim();
  const startTimeStr = String(params.startTime || "").trim();
  const endTimeStr = String(params.endTime || "").trim();
  if (!studentName || !subjectName || !dateStr || !startTimeStr || !endTimeStr) {
    return { ok: false, message: "請完整填寫講師、學生、課程、日期與時間。" };
  }

  const courseSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  if (!courseSheet) return { ok: false, message: "找不到課程設定分頁。" };
  const courseData = courseSheet.getDataRange().getValues();
  let unitFee = 0;
  for (let i = 1; i < courseData.length; i++) {
    const rowOwner = String(courseData[i][0] || "").trim();
    const rowStudent = String(courseData[i][2] || "").trim();
    const rowSubject = String(courseData[i][3] || "").trim();
    if (isCourseOwner(rowOwner, target.lineUserId, target.name) && rowStudent === studentName && rowSubject === subjectName) {
      unitFee = parseFloat(courseData[i][4]) || 0;
      break;
    }
  }
  if (unitFee <= 0) {
    return { ok: false, message: `找不到「${target.name} / ${studentName} / ${subjectName}」的課程設定或鐘點單價。` };
  }

  const timeZone = Session.getScriptTimeZone();
  const formattedDateStr = dateStr.replace(/-/g, "/");
  const cleanStart = startTimeStr.replace(/:/g, "");
  const cleanEnd = endTimeStr.replace(/:/g, "");
  const startNum = parseTime(cleanStart);
  const endNum = parseTime(cleanEnd);
  const duration = Math.round((endNum - startNum) * 100) / 100;
  const totalPay = Math.round(duration * unitFee);
  if (duration <= 0) return { ok: false, message: "上課結束時間必須晚於開始時間。" };
  if (duration > 4) return { ok: false, message: "課程時數超過 4 小時，請確認是否填錯。" };

  const inputDate = new Date(formattedDateStr);
  const lessonStartDate = buildLessonDateTime(formattedDateStr, cleanStart, timeZone);
  const lessonEndDate = buildLessonDateTime(formattedDateStr, cleanEnd, timeZone);
  if (!lessonStartDate || !lessonEndDate) return { ok: false, message: "日期或時間格式錯誤。" };
  const currentTime = new Date();
  if (!isPlan && lessonEndDate.getTime() > currentTime.getTime()) {
    return { ok: false, message: "授課登記限已完成課程，不能登錄未來時間。" };
  }
  if (isPlan && lessonStartDate.getTime() < currentTime.getTime()) {
    return { ok: false, message: "預排課程不能選擇已過去的時間。" };
  }

  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  const planHistory = planSheet ? planSheet.getDataRange().getValues() : [];
  const recordHistory = recordSheet ? recordSheet.getDataRange().getValues() : [];
  const conflicts = findScheduleConflicts(
    planHistory.concat(recordHistory),
    {
      teacher: target.name,
      student: studentName,
      subject: subjectName,
      dateFingerprint: Utilities.formatDate(inputDate, timeZone, "yyyyMMdd"),
      start: startNum,
      end: endNum
    },
    timeZone
  );
  if (conflicts.length > 0) return { ok: false, message: "🚫 登記失敗：\n" + conflicts.join("\n") };

  return {
    ok: true,
    hours: duration,
    totalPay,
    cleanStart,
    cleanEnd,
    writeDate: Utilities.formatDate(inputDate, timeZone, "yyyy/MM/dd")
  };
}

function previewAdminProxyVerify_(params: any, operatorLineUserId: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const target = getCourseProxyTarget_(ss, String(params.targetTeacher || ""));
  if (!target) return { ok: false, message: "請先選擇要代操作的講師。" };
  const item = getProxyVerifySchedule_(ss, parseInt(params.rowId, 10), target.name);
  if (!item.ok) return item;
  const operatorName = getTeacherNameByLineUserId(operatorLineUserId) || "行政人員";
  const resultLabel = String(params.isVerified) === "yes" ? "有上課，轉入授課紀錄" : "未上課，取消預排";
  return {
    ok: true,
    preview: {
      title: "代核銷預排",
      summary: `請確認是否由 ${operatorName} 幫 ${target.name} 講師核銷預排。`,
      items: [
        `講師：${target.name}`,
        `學生：${item.student}`,
        `課程：${item.subject}`,
        `時間：${item.date} ${item.startTime}-${item.endTime}`,
        `核銷結果：${resultLabel}`
      ],
      nextAction: "確認代核銷",
      canConfirm: true,
      confirmAction: "adminCourseProxyConfirm"
    }
  };
}

function confirmAdminProxyVerify_(params: any, operatorLineUserId: string) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const target = getCourseProxyTarget_(ss, String(params.targetTeacher || ""));
  if (!target) return { ok: false, message: "請先選擇要代操作的講師。" };
  const item = getProxyVerifySchedule_(ss, parseInt(params.rowId, 10), target.name);
  if (!item.ok) return item;
  const result = handleLiffVerifySchedule({
    lineUserId: operatorLineUserId,
    rowId: params.rowId,
    isVerified: params.isVerified
  });
  if (!result.ok) return result;
  return { ok: true, message: `已完成代核銷：${target.name} / ${item.student} / ${item.subject}` };
}

function getProxyVerifySchedule_(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, rowId: number, targetTeacherName: string) {
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  if (!planSheet) return { ok: false, message: "找不到預排工作表。" };
  if (!rowId || rowId < 2 || rowId > planSheet.getLastRow()) return { ok: false, message: "預排資料列不存在。" };
  const rowValues = planSheet.getRange(rowId, 1, 1, 13).getValues()[0];
  const recordTeacher = String(rowValues[1] || "").trim();
  const recordStatus = String(rowValues[9] || "").trim();
  if (recordTeacher !== targetTeacherName) return { ok: false, message: "此預排紀錄不屬於指定講師。" };
  if (recordStatus !== "未核銷") return { ok: false, message: `此預排紀錄目前狀態為「${recordStatus || "空白"}」，不可重複核銷。` };
  const timeZone = Session.getScriptTimeZone();
  const lessonEnd = buildLessonEndDate(rowValues[2], rowValues[4], timeZone);
  if (lessonEnd && lessonEnd.getTime() > new Date().getTime()) {
    return { ok: false, message: "此預排課程尚未到下課時間，不可提前核銷。" };
  }
  return {
    ok: true,
    student: String(rowValues[7] || "").trim(),
    subject: String(rowValues[8] || "").trim(),
    date: rowValues[2] instanceof Date ? Utilities.formatDate(rowValues[2], timeZone, "yyyy/MM/dd") : String(rowValues[2] || ""),
    startTime: formatTimeStr(rowValues[3]),
    endTime: formatTimeStr(rowValues[4])
  };
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


