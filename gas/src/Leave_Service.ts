// ==========================================
// 🏖️ Leave_Service.ts : 請假管理系統 (TypeScript + LIFF 整合重構版)
// ==========================================

// ==========================================
// 1. LIFF 專屬請假 API 後端實作
// ==========================================

// handleLiffLeave : 處理來自 LIFF 的請假申請
function handleLiffLeave(params: any) {
  const lineUserId = params.lineUserId;
  const leaveType = params.type; // '事假' | '病假'
  const dateStr = params.date; // YYYY-MM-DD
  const startTimeStr = params.startTime; // HH:MM
  const endTimeStr = params.endTime; // HH:MM

  const ss = SpreadsheetApp.openById(LEAVE_SHEET_ID);
  const recordSheet = ss.getSheetByName("請假紀錄");
  const configSheet = ss.getSheetByName("假別設定");

  if (!recordSheet || !configSheet) {
    return { ok: false, message: "請假系統工作表尚未初始化。" };
  }

  // 1. 取得講師姓名
  const coreSs = SpreadsheetApp.openById(SHEET_ID);
  const teacherSheet = coreSs.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return { ok: false, message: "找不到講師名單。" };
  
  const teacherData = teacherSheet.getDataRange().getValues();
  let userName = "";
  for (let i = 1; i < teacherData.length; i++) {
    if (String(teacherData[i][1]).trim() === lineUserId) {
      userName = teacherData[i][0];
      break;
    }
  }
  if (!userName) return { ok: false, message: "身分驗證失敗，找不到講師資料。" };

  // 2. 計算學年度與日期
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const academicYear = (currentMonth >= 9) ? (currentYear - 1911) : (currentYear - 1912);

  const formattedDateStr = dateStr.replace(/-/g, "/"); // YYYY/MM/DD
  const inputDate = new Date(formattedDateStr);
  const finalDateStr = (inputDate.getMonth() + 1) + "/" + inputDate.getDate(); // e.g. "5/28"

  const cleanStart = startTimeStr.replace(/:/g, ""); // e.g. "0900"
  const cleanEnd = endTimeStr.replace(/:/g, ""); // e.g. "1200"

  const s = parseTimeMinutes(cleanStart);
  const e = parseTimeMinutes(cleanEnd);

  if (s >= e) {
    return { ok: false, message: "請假結束時間必須晚於開始時間。" };
  }

  // 3. 準備限額與重複檢查
  const configs = configSheet.getDataRange().getValues();
  const limits: any = {};
  for (let i = 1; i < configs.length; i++) {
    limits[configs[i][0]] = configs[i][3];
  }

  const used = { "事假": 0, "病假": 0 };
  const records = recordSheet.getDataRange().getValues();
  const existingRanges: any[] = [];

  for (let i = 1; i < records.length; i++) {
    // 比對講師姓名 (因為是用 LINE ID/姓名作為關聯，原本請假表存的是申請人ID，重構中我們支援比對姓名或 LINE ID)
    const recordOperator = records[i][2]; // 申請人ID或姓名
    const recordYear = records[i][1];
    
    // 如果請假記錄中的 operator 吻合該講師的 lineUserId 或是其姓名
    if ((recordOperator === lineUserId || recordOperator === userName) && recordYear == academicYear) {
      const type = records[i][3];
      const hr = parseFloat(records[i][5]) || 0;
      if (used[type] !== undefined) used[type] += hr;

      const dateTimeStr = records[i][4];
      const dtParts = dateTimeStr.split(" ");
      if (dtParts.length >= 2) {
        const dStr = dtParts[0];
        const tRange = dtParts[1].split("-");
        if (tRange.length === 2) {
          existingRanges.push({
            dateStr: dStr,
            start: parseTimeMinutes(tRange[0]),
            end: parseTimeMinutes(tRange[1])
          });
        }
      }
    }
  }

  // 重複申請檢核
  const isOverlap = existingRanges.some((range) => {
    return range.dateStr === finalDateStr && s < range.end && e > range.start;
  });

  if (isOverlap) {
    return { ok: false, message: `🚫 重複請假：您在 ${finalDateStr} 已有登記相同時段的請假！` };
  }

  // 4. 計算請假時數 (扣除 12:00 - 13:30 休息時間)
  const totalDuration = e - s;
  const breakStart = 12 * 60;      
  const breakEnd = 13 * 60 + 30;   
  const overlapStart = Math.max(s, breakStart);
  const overlapEnd = Math.min(e, breakEnd);
  let breakDedution = 0;
  
  if (overlapStart < overlapEnd) {
    breakDedution = overlapEnd - overlapStart;
  }

  const netMinutes = totalDuration - breakDedution;
  const netHours = netMinutes / 60;

  // 更新使用量與剩餘額度
  if (used[leaveType] !== undefined) {
    used[leaveType] += netHours;
  }
  const currentBalancePersonal = (limits["事假"] || 0) - used["事假"];
  const currentBalanceSick = (limits["病假"] || 0) - used["病假"];

  if (currentBalancePersonal < 0 && leaveType === "事假") {
    return { ok: false, message: "⚠️ 登記失敗：事假額度不足。" };
  }
  if (currentBalanceSick < 0 && leaveType === "病假") {
    return { ok: false, message: "⚠️ 登記失敗：病假額度不足。" };
  }

  // 寫入請假紀錄
  recordSheet.appendRow([
    new Date(),
    academicYear,
    lineUserId, // 原本是以 LINE ID 寫入
    leaveType,
    finalDateStr + " " + cleanStart + "-" + cleanEnd,
    netHours,
    currentBalancePersonal,
    currentBalanceSick,
    `講師姓名:${userName}` // 備註備份姓名
  ]);

  return {
    ok: true,
    message: "請假申請登記完成！",
    profile: {
      casualLeaveRemaining: currentBalancePersonal,
      sickLeaveRemaining: currentBalanceSick
    }
  };
}

// 供 Core_Service 調用計算講師剩餘請假額度
function getTeacherLeaveQuota(teacherName: string) {
  const ss = SpreadsheetApp.openById(LEAVE_SHEET_ID);
  const recordSheet = ss.getSheetByName("請假紀錄");
  const configSheet = ss.getSheetByName("假別設定");

  const fallback = { casual: 8, sick: 24 };
  if (!recordSheet || !configSheet) return fallback;

  // 取得 LINE ID
  const coreSs = SpreadsheetApp.openById(SHEET_ID);
  const teacherSheet = coreSs.getSheetByName(SHEET_NAME_TEACHER);
  if (!teacherSheet) return fallback;
  
  const teacherData = teacherSheet.getDataRange().getValues();
  let lineUserId = "";
  for (let i = 1; i < teacherData.length; i++) {
    if (teacherData[i][0] === teacherName) {
      lineUserId = String(teacherData[i][1]).trim();
      break;
    }
  }

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const academicYear = (currentMonth >= 9) ? (currentYear - 1911) : (currentYear - 1912);

  const configs = configSheet.getDataRange().getValues();
  const limits: any = {};
  for (let i = 1; i < configs.length; i++) {
    limits[configs[i][0]] = configs[i][3];
  }

  const used = { "事假": 0, "病假": 0 };
  const records = recordSheet.getDataRange().getValues();

  for (let i = 1; i < records.length; i++) {
    const recordOperator = records[i][2];
    const recordYear = records[i][1];

    if ((recordOperator === lineUserId || recordOperator === teacherName) && recordYear == academicYear) {
      const type = records[i][3];
      const hr = parseFloat(records[i][5]) || 0;
      if (used[type] !== undefined) used[type] += hr;
    }
  }

  return {
    casual: Math.max(0, (limits["事假"] || 49) - used["事假"]),
    sick: Math.max(0, (limits["病假"] || 196) - used["病假"])
  };
}

// ==========================================
// 2. 原 LINE 對話控制請假重構模組
// ==========================================

function handleLeaveMenu(event: any) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;

  const ss = SpreadsheetApp.openById(LEAVE_SHEET_ID);
  let recordSheet = ss.getSheetByName("請假紀錄");
  let configSheet = ss.getSheetByName("假別設定");

  if (!recordSheet || !configSheet) {
    if (!configSheet) {
        configSheet = ss.insertSheet("假別設定");
        configSheet.appendRow(["假別名稱", "每年天數", "每天時數", "總時數"]);
        configSheet.appendRow(["事假", 7, 7, 49]);
        configSheet.appendRow(["病假", 28, 7, 196]);
    }
    if (!recordSheet) {
        recordSheet = ss.insertSheet("請假紀錄");
        recordSheet.appendRow(["申請時間", "學年度", "申請人ID", "假別", "日期/時間", "請假時數", "事假餘額", "病假餘額", "備註"]);
    }
  }

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1; 
  const academicYear = (currentMonth >= 9) ? (currentYear - 1911) : (currentYear - 1912);

  const configs = configSheet.getDataRange().getValues();
  const limits: any = {}; 
  for (let i = 1; i < configs.length; i++) {
    limits[configs[i][0]] = configs[i][3];
  }

  const used = { "事假": 0, "病假": 0 };
  const records = recordSheet.getDataRange().getValues();
  for (let i = 1; i < records.length; i++) {
    if (records[i][2] === userId && records[i][1] == academicYear) {
      const type = records[i][3];
      const hr = parseFloat(records[i][5]);
      if (used[type] !== undefined) used[type] += hr;
    }
  }

  const personalLeaveLeft = (limits["事假"] || 0) - used["事假"];
  const sickLeaveLeft = (limits["病假"] || 0) - used["病假"];
  
  const text = "👤 秘書 (" + academicYear + "學年度)\n\n" +
             "📝 事假(" + (limits["事假"] || 0) + ")：已用 " + used["事假"] + "hr / 剩 " + personalLeaveLeft + "hr\n" +
             "😷 病假(" + (limits["病假"] || 0) + ")：已用 " + used["病假"] + "hr / 剩 " + sickLeaveLeft + "hr";

  const bubble = {
    "type": "bubble",
    "body": {
      "type": "box", "layout": "vertical",
      "contents": [
        { "type": "text", "text": "請假系統", "weight": "bold", "size": "xl", "color": "#1DB446" },
        { "type": "separator", "margin": "md" },
        { "type": "text", "text": text, "wrap": true, "margin": "md", "size": "sm", "color": "#555555" }
      ]
    },
    "footer": {
      "type": "box", "layout": "horizontal", "spacing": "sm",
      "contents": [
        { "type": "button", "style": "secondary", "action": { "type": "postback", "label": "📝 請事假", "data": "action=leave_pick&t=事假" } },
        { "type": "button", "style": "secondary", "action": { "type": "postback", "label": "😷 請病假", "data": "action=leave_pick&t=病假" } }
      ]
    }
  };

  // 統一從外部 UI_Utils 呼叫發送，免去 UrlFetchApp 重複代碼
  replyFlexMessage(replyToken, "請假餘額", bubble);
}

function handleLeavePick(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const leaveType = postbackData.split("t=")[1];

  const cache = CacheService.getScriptCache();
  const stateObj = { status: "WAITING_LEAVE_TIME", type: leaveType };
  cache.put(userId, JSON.stringify(stateObj), 300);

  replyLineMessage(replyToken, 
    "已選擇【" + leaveType + "】\n" +
    "請輸入：日期 起 迄 (支援多行)\n" +
    "範例：\n" +
    "10/05 0800 1630\n" +
    "10/06 0900 1200"
  );
}

function processLeaveInput(event: any, userMsg: string, state: any) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const leaveType = state.type;
  
  const lines = userMsg.split("\n");
  let successCount = 0;
  let failCount = 0;
  let reportDetail = "";
  
  const ss = SpreadsheetApp.openById(LEAVE_SHEET_ID);
  const recordSheet = ss.getSheetByName("請假紀錄");
  const configSheet = ss.getSheetByName("假別設定"); 
  
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const academicYear = (currentMonth >= 9) ? (currentYear - 1911) : (currentYear - 1912);
  const todayDate = new Date(); 
  todayDate.setHours(0,0,0,0);

  const configs = configSheet.getDataRange().getValues();
  const limits: any = {};
  for (let i = 1; i < configs.length; i++) limits[configs[i][0]] = configs[i][3];

  const used = { "事假": 0, "病假": 0 };
  const records = recordSheet.getDataRange().getValues();
  
  const existingRanges: any[] = [];
  for (let i = 1; i < records.length; i++) {
    if (records[i][2] === userId && records[i][1] == academicYear) {
      const t = records[i][3];
      const h = parseFloat(records[i][5]);
      if (used[t] !== undefined) used[t] += h;
      
      const dateTimeStr = records[i][4]; 
      const dtParts = dateTimeStr.split(" ");
      if (dtParts.length >= 2) {
        const dStr = dtParts[0];
        const tRange = dtParts[1].split("-");
        if (tRange.length === 2) {
          existingRanges.push({
            dateStr: dStr,
            start: parseTimeMinutes(tRange[0]),
            end: parseTimeMinutes(tRange[1])
          });
        }
      }
    }
  }
  
  for (let k = 0; k < lines.length; k++) {
    const line = lines[k].trim();
    if (line === "") continue;
    
    const cleanMsg = line.replace(/:/g, "").trim();
    const parts = cleanMsg.split(/[ ]+/);
    
    if (parts.length < 3) {
      failCount++;
      reportDetail += "❌ 格式錯：" + line + "\n";
      continue;
    }

    const dateStr = parts[0];
    const startStr = parts[1];
    const endStr = parts[2];
    
    try {
      let inputDate;
      if (dateStr.match(/^\d{4}\//)) {
         inputDate = new Date(dateStr);
      } else {
         const monthPart = parseInt(dateStr.split("/")[0]); 
         const dayPart = parseInt(dateStr.split("/")[1]);
         const tempDate = new Date(currentYear, monthPart - 1, dayPart);
         const diffTime = tempDate.getTime() - todayDate.getTime();
         const diffDays = diffTime / (1000 * 3600 * 24);
         
         if (diffDays > 180) inputDate = new Date(currentYear - 1, monthPart - 1, dayPart);
         else if (diffDays < -180) inputDate = new Date(currentYear + 1, monthPart - 1, dayPart);
         else inputDate = tempDate;
      }
      const finalDateStr = (inputDate.getMonth()+1) + "/" + inputDate.getDate();

      const s = parseTimeMinutes(startStr); 
      const e = parseTimeMinutes(endStr);
      
      if (s >= e) {
        failCount++;
        reportDetail += "❌ 時間錯 (迄<起)：" + line + "\n";
        continue;
      }

      const isOverlap = existingRanges.some((range) => {
        return range.dateStr === finalDateStr && s < range.end && e > range.start;
      });

      if (isOverlap) {
        failCount++;
        reportDetail += "❌ 重複申請：" + line + "\n";
        continue;
      }

      const totalDuration = e - s;
      const breakStart = 12 * 60;      
      const breakEnd = 13 * 60 + 30;   
      const overlapStart = Math.max(s, breakStart);
      const overlapEnd = Math.min(e, breakEnd);
      let breakDedution = 0;
      
      if (overlapStart < overlapEnd) {
        breakDedution = overlapEnd - overlapStart;
      }

      const netMinutes = totalDuration - breakDedution;
      const netHours = netMinutes / 60;

      if (used[leaveType] !== undefined) {
          used[leaveType] += netHours;
      }
      const currentBalancePersonal = (limits["事假"] || 0) - used["事假"];
      const currentBalanceSick = (limits["病假"] || 0) - used["病假"];

      recordSheet.appendRow([
        new Date(),        
        academicYear,      
        userId,            
        leaveType,         
        finalDateStr + " " + startStr + "-" + endStr, 
        netHours,          
        currentBalancePersonal, 
        currentBalanceSick,     
        ""                 
      ]);
      
      existingRanges.push({ dateStr: finalDateStr, start: s, end: e });
      successCount++;
      reportDetail += "✅ " + finalDateStr + " " + startStr + "-" + endStr + " (" + netHours + "hr)\n";
      
    } catch (e) {
      failCount++;
      reportDetail += "❌ 錯誤：" + line + "\n";
    }
  }

  CacheService.getScriptCache().remove(userId);

  if (successCount > 0) {
      const finalRemaining = (limits[leaveType] || 0) - (used[leaveType] || 0);
      replyLineMessage(replyToken, 
        "📝 請假登錄完成\n" +
        "假別：" + leaveType + "(餘" + finalRemaining + ")\n" +
        "--------------------\n" +
        reportDetail +
        "(成功 " + successCount + " / 失敗 " + failCount + ")"
      );
  } else {
      replyLineMessage(replyToken, "❌ 登錄失敗，請檢查內容：\n" + reportDetail);
  }
}

function parseTimeMinutes(tStr: string): number {
  let val = tStr;
  while(val.length < 4) val = "0" + val;
  const h = parseInt(val.substring(0, 2));
  const m = parseInt(val.substring(2, 4));
  return h * 60 + m;
}

// 宣告外部參考
declare function replyLineMessage(token: string, msg: string): void;
declare function replyFlexMessage(token: string, altText: string, bubble: any): void;

