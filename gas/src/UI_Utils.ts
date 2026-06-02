// ==========================================
// 🛠️ UI_Utils.ts : 介面與 LINE API 工具函式 (TypeScript DRY 重構版)
// ==========================================

// ==========================================
// 1. 統一的 LINE API 請求客戶端 (DRY Principle)
// ==========================================
var LineClient = {
  post: function (endpoint: string, payload: any) {
    const url = "https://api.line.me/v2/bot/" + endpoint;
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + CHANNEL_TOKEN
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) {
        Logger.log(`❌ LINE API HTTP Error (${endpoint}): ${response.getContentText()}`);
      }
      return response;
    } catch (e) {
      Logger.log(`❌ UrlFetchApp failed (${endpoint}): ${e.toString()}`);
      return null;
    }
  },

  reply: function (replyToken: string, messages: any[]) {
    return this.post("message/reply", { replyToken, messages });
  },

  push: function (to: string, messages: any[]) {
    return this.post("message/push", { to, messages });
  }
};

// --- 回覆文字訊息 ---
function replyLineMessage(token: string, text: string) {
  LineClient.reply(token, [{
    "type": "text",
    "text": text
  }]);
}

// --- 推播文字訊息 ---
function pushLineMessage(userId: string, text: string) {
  LineClient.push(userId, [{
    "type": "text",
    "text": text
  }]);
}

// --- 回覆 Flex 訊息 ---
function replyFlexMessage(token: string, altText: string, flexBubble: any) {
  LineClient.reply(token, [{
    "type": "flex",
    "altText": altText,
    "contents": flexBubble
  }]);
}

// --- [API] 回覆確認按鈕 ---
function replyConfirmButton(token: string, text: string, actionData: string) {
  const confirmTemplate = {
    "type": "template",
    "altText": "確認",
    "template": {
      "type": "confirm",
      "text": text.substring(0, 240),
      "actions": [{
        "type": "postback",
        "label": "是",
        "data": actionData
      }, {
        "type": "postback",
        "label": "否",
        "data": "action=cancel"
      }]
    }
  };
  LineClient.reply(token, [confirmTemplate]);
}

// --- [API] 財務確認卡片 ---
function replyConfirmationCard(token: string, title: string, text: string, cacheKey: string) {
  const card = {
    "type": "bubble",
    "body": {
      "type": "box",
      "layout": "vertical",
      "contents": [{
        "type": "text",
        "text": title,
        "weight": "bold",
        "size": "xl",
        "color": "#1DB446"
      }, {
        "type": "separator",
        "margin": "md"
      }, {
        "type": "text",
        "text": text,
        "wrap": true,
        "margin": "md",
        "size": "sm",
        "color": "#555555"
      }]
    },
    "footer": {
      "type": "box",
      "layout": "vertical",
      "spacing": "sm",
      "contents": [{
        "type": "button",
        "style": "primary",
        "color": "#1DB446",
        "action": {
          "type": "postback",
          "label": "✅ 確認寫入",
          "data": "action=fin_confirm"
        }
      }, {
        "type": "button",
        "style": "secondary",
        "color": "#aaaaaa",
        "action": {
          "type": "postback",
          "label": "🆗 取消",
          "data": "action=fin_cancel"
        }
      }, {
        "type": "button",
        "style": "secondary",
        "color": "#ff5555",
        "action": {
          "type": "postback",
          "label": "⚠️ 資料有誤",
          "data": "action=fin_report"
        }
      }]
    }
  };

  replyFlexMessage(token, "財務試算確認", card);
}

// ==========================================
// 2. 原 LINE 訊息對話控制介面重構模組
// ==========================================

// --- 核銷功能：處理核銷選單 ---
function handleVerifyMenu(event: any, targetUserId: string) {
  const replyToken = event.replyToken;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!userSheet) return;
  const uData = userSheet.getDataRange().getValues();
  let myName = "";

  for (let i = 1; i < uData.length; i++) {
    if (uData[i][1] === targetUserId) {
      myName = uData[i][0];
      break;
    }
  }

  if (myName === "") {
    replyLineMessage(replyToken, "❌ 找不到講師資料。");
    return;
  }

  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  if (!planSheet) {
    replyLineMessage(replyToken, "⚠️ 無預排紀錄。");
    return;
  }

  const data = planSheet.getDataRange().getValues();
  const pendingList: any[] = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][9] ? data[i][9].toString().trim() : "";
    if (data[i][1] === myName && status === "未核銷") {
      pendingList.push({
        row: i + 1,
        date: data[i][2],
        start: data[i][3],
        end: data[i][4],
        stu: data[i][7],
        course: data[i][8]
      });
    }
  }

  if (pendingList.length === 0) {
    replyLineMessage(replyToken, "🎉 " + myName + " 講師目前沒有未核銷課程。");
    return;
  }

  const bubbles: any[] = [];
  const limit = Math.min(pendingList.length, 5);
  const timeZone = Session.getScriptTimeZone();

  for (let k = 0; k < limit; k++) {
    const item = pendingList[k];
    let dateStr = item.date;
    if (item.date instanceof Date) dateStr = Utilities.formatDate(item.date, timeZone, "yyyy/MM/dd");

    const bubble = {
      "type": "bubble",
      "size": "kilo",
      "header": {
        "type": "box",
        "layout": "vertical",
        "backgroundColor": "#fff8dc",
        "contents": [{
          "type": "text",
          "text": "預排核對 (" + myName + ")",
          "weight": "bold",
          "color": "#aaaaaa",
          "size": "xxs"
        }]
      },
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [{
          "type": "text",
          "text": item.stu,
          "weight": "bold",
          "size": "xl"
        }, {
          "type": "text",
          "text": item.course,
          "size": "sm",
          "color": "#666666"
        }, {
          "type": "text",
          "text": dateStr + " " + item.start + "-" + item.end,
          "size": "md",
          "margin": "md",
          "color": "#1DB446"
        }]
      },
      "footer": {
        "type": "box",
        "layout": "horizontal",
        "spacing": "sm",
        "contents": [{
          "type": "button",
          "style": "primary",
          "color": "#1DB446",
          "action": {
            "type": "postback",
            "label": "🙆‍♂️ 有上課",
            "data": "action=verify_yes&r=" + item.row
          }
        }, {
          "type": "button",
          "style": "secondary",
          "color": "#aaaaaa",
          "action": {
            "type": "postback",
            "label": "🙅‍♀️ 沒上課",
            "data": "action=verify_no&r=" + item.row
          }
        }]
      }
    };
    bubbles.push(bubble);
  }

  replyFlexMessage(replyToken, "預排核對", {
    "type": "carousel",
    "contents": bubbles
  });
}

// --- 核銷功能：執行核銷動作 ---
function executeVerify(event: any, postbackData: string, isAttended: boolean) {
  const replyToken = event.replyToken;
  const parts = postbackData.split("&");
  let row = -1;
  for (let i = 0; i < parts.length; i++) {
    if (parts[i].indexOf("r=") === 0) row = parseInt(parts[i].substring(2));
  }
  if (row === -1) return;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
  const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  if (!planSheet || !recordSheet) return;
  
  const range = planSheet.getRange(row, 1, 1, 10);
  const rowValues = range.getValues()[0];

  if (rowValues[9].toString().trim() !== "未核銷") {
    replyLineMessage(replyToken, "⚠️ 此筆已處理過。");
    return;
  }

  const dateVal = rowValues[2];
  let dateStr = dateVal;
  const tz = Session.getScriptTimeZone();
  if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, tz, "yyyy/MM/dd");
  const timeInfo = dateStr + " " + rowValues[3] + "-" + rowValues[4];
  const settlementMonth = Utilities.formatDate(new Date(), tz, "yyyy/MM");

  if (isAttended) {
    planSheet.getRange(row, 10).setValue("已核銷");
    planSheet.getRange(row, 11).setValue(settlementMonth); 
    const recordData = rowValues.slice(0, 9);
    recordData[0] = new Date();
    recordSheet.appendRow(recordData);
    replyLineMessage(replyToken, "✅ 核銷完成：\n" + timeInfo);
  } else {
    planSheet.getRange(row, 10).setValue("取消");
    planSheet.getRange(row, 12).setValue(settlementMonth); 
    replyLineMessage(replyToken, "🗑️ 已取消預排：\n" + timeInfo);
  }
}

// --- 選單：顯示學生清單 ---
function handleStudentMenu(event: any, targetUserId: string, isProxy: boolean, mode: string, filter?: string) {
  const replyToken = event.replyToken;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!userSheet) return;
  const uData = userSheet.getDataRange().getValues();
  let teacherName = "該位";
  for (let i = 1; i < uData.length; i++) { if (uData[i][1] === targetUserId) { teacherName = uData[i][0]; break; } }

  const stuSheet = ss.getSheetByName(SHEET_NAME_COURSE);
  if (!stuSheet) return;
  const data = stuSheet.getDataRange().getValues();
  const studentMap: any = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetUserId) {
      const sName = data[i][2]; const cName = data[i][3];
      if (filter && sName.indexOf(filter) === -1) continue;
      if (!studentMap[sName]) studentMap[sName] = [];
      studentMap[sName].push(cName);
    }
  }

  const studentNames = Object.keys(studentMap);
  if (studentNames.length === 0) { replyLineMessage(replyToken, "⚠️ " + teacherName + " 講師目前沒有學生資料。"); return; }

  const bubbles: any[] = [];
  const limit = Math.min(studentNames.length, 12);
  for (let k = 0; k < limit; k++) {
    const sName = studentNames[k]; const courses = studentMap[sName];
    const isPlan = (mode === 'plan'); const labelText = isPlan ? "預排" : "登記";
    const headerText = isPlan ? (isProxy ? "[代預排] " + teacherName + " 講師" : "[預排] " + teacherName + " 講師") : (isProxy ? "[代登] " + teacherName + " 講師" : "[已上課] " + teacherName + " 講師");
    
    const courseRows: any[] = [];
    for (let c = 0; c < Math.min(courses.length, 8); c++) {
      const cName = courses[c];
      const action = isProxy ? "action=admin_pick_stu" : "action=pick_stu";
      const postData = action + "&n=" + encodeURIComponent(sName) + "&c=" + encodeURIComponent(cName) + "&m=" + mode + (isProxy ? ("&tid=" + targetUserId) : "");
      courseRows.push({
        "type": "box", "layout": "horizontal", "margin": "md",
        "contents": [
          { "type": "box", "layout": "vertical", "contents": [ { "type": "text", "text": cName, "size": "sm", "color": "#555555", "wrap": true, "weight": "bold" }, { "type": "text", "text": "點右側登記", "size": "xxs", "color": "#aaaaaa" } ], "flex": 4 },
          { "type": "button", "action": { "type": "postback", "label": labelText, "data": postData }, "style": "primary", "color": "#00aa00", "height": "sm", "flex": 2 }
        ]
      });
      if (c < courses.length - 1) courseRows.push({ "type": "separator", "margin": "md" });
    }
    bubbles.push({
      "type": "bubble", "size": "kilo",
      "header": { "type": "box", "layout": "vertical", "backgroundColor": isPlan ? "#E6E6FA" : "#F8F9FA", "contents": [ { "type": "text", "text": headerText, "size": "xxs", "color": "#aaaaaa" }, { "type": "text", "text": sName, "weight": "bold", "size": "xl", "color": "#2C3E50" } ] },
      "body": { "type": "box", "layout": "vertical", "contents": courseRows }
    });
  }

  replyFlexMessage(replyToken, "學生課程選單", {
    "type": "carousel",
    "contents": bubbles
  });
}

// --- 學生選擇處理 ---
function handleStudentPick(event: any, postbackData: string, userId: string) {
  const replyToken = event.replyToken;
  const parts = postbackData.split("&");
  let stuName = "", stuCourse = "", mode = "record";

  for (let i = 0; i < parts.length; i++) {
    if (parts[i].indexOf("n=") === 0) stuName = decodeURIComponent(parts[i].substring(2));
    if (parts[i].indexOf("c=") === 0) stuCourse = decodeURIComponent(parts[i].substring(2));
    if (parts[i].indexOf("m=") === 0) mode = parts[i].substring(2);
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!userSheet) return;
  const uData = userSheet.getDataRange().getValues();
  let myName = "您";

  for (let i = 1; i < uData.length; i++) {
    if (uData[i][1] === userId) {
      myName = uData[i][0];
      break;
    }
  }

  const cache = CacheService.getScriptCache();
  const stateObj = {
    status: "WAITING_TIME",
    targetId: userId,
    sName: stuName,
    sCourse: stuCourse,
    isProxy: false,
    isPlan: (mode === 'plan'),
    mode: mode
  };
  cache.put(userId, JSON.stringify(stateObj), 300);

  const modeTitle = (mode === 'plan') ? "📅 [預排模式]" : "✅ [已上課登記]";
  replyLineMessage(replyToken, modeTitle + " 已選擇：\n" + myName + " 講師：" + stuName + " (" + stuCourse + ")\n請輸入時間 (範例 12/28 1000 1200)");
}

// --- 行政搜尋：直接輸入姓名 ---
function handleAdminDirectSearch(event: any, targetName: string, mode: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  let isAdmin = false;

  for (let a = 0; a < ADMIN_LIST.length; a++) {
    if (userId === ADMIN_LIST[a]) {
      isAdmin = true;
      break;
    }
  }

  if (!isAdmin) {
    replyLineMessage(replyToken, "❌ 權限不足");
    return;
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!userSheet) return;
  const data = userSheet.getDataRange().getValues();
  let targetId = "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetName) {
      targetId = data[i][1];
      break;
    }
  }

  if (targetId === "") {
    replyLineMessage(replyToken, "❌ 找不到講師：" + targetName);
  } else {
    if (mode === "verify") handleVerifyMenu(event, targetId);
    else handleStudentMenu(event, targetId, true, mode);
  }
}

// --- 行政選單：顯示所有講師 ---
function handleAdminMenu(event: any, mode: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  let isAdmin = false;

  for (let a = 0; a < ADMIN_LIST.length; a++) {
    if (userId === ADMIN_LIST[a]) {
      isAdmin = true;
      break;
    }
  }

  if (!isAdmin) {
    replyLineMessage(replyToken, "❌ 權限不足：限行政人員使用。");
    return;
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!userSheet) return;
  const data = userSheet.getDataRange().getValues();
  const bubbles: any[] = [];
  const limit = Math.min(data.length, 12);
  const actionPrefix = "action=admin_pick_teacher";
  let btnLabel = "幫他填";
  if (mode === "plan") btnLabel = "幫預排";
  if (mode === "verify") btnLabel = "幫核銷";

  const pendingTeachers: string[] = [];
  if (mode === "verify") {
    const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);
    if (planSheet) {
      const pData = planSheet.getDataRange().getValues();
      for (let p = 1; p < pData.length; p++) {
        if (pData[p][9] == "未核銷") {
          pendingTeachers.push(pData[p][1]);
        }
      }
    }
  }

  for (let i = 1; i < limit; i++) {
    const tName = data[i][0];
    const tId = data[i][1];
    if (tId && tId.indexOf("U") === 0 && ADMIN_LIST.indexOf(tId) === -1) {
      if (mode === "verify" && pendingTeachers.indexOf(tName) === -1) continue;
      bubbles.push({
        "type": "bubble",
        "size": "micro",
        "body": {
          "type": "box",
          "layout": "vertical",
          "contents": [{
            "type": "text",
            "text": tName,
            "weight": "bold",
            "align": "center"
          }]
        },
        "footer": {
          "type": "box",
          "layout": "vertical",
          "contents": [{
            "type": "button",
            "style": "primary",
            "height": "sm",
            "action": {
              "type": "postback",
              "label": btnLabel,
              "data": actionPrefix + "&m=" + mode + "&tid=" + tId
            }
          }]
        }
      });
    }
  }

  if (bubbles.length === 0) {
    const emptyMsg = (mode === "verify") ? "🎉 目前沒有任何講師需要核銷！" : "⚠️ 講師名單為空。";
    replyLineMessage(replyToken, emptyMsg);
    return;
  }

  replyFlexMessage(replyToken, "講師選單", {
    "type": "carousel",
    "contents": bubbles
  });
}

// --- 行政選擇講師回傳處理 ---
function handleAdminPickTeacher(event: any, postbackData: string) {
  const parts = postbackData.split("&");
  let targetId = "", mode = "record";
  for (let i = 0; i < parts.length; i++) {
    if (parts[i].indexOf("tid=") === 0) targetId = parts[i].substring(4);
    if (parts[i].indexOf("m=") === 0) mode = parts[i].substring(2);
  }
  if (mode === "verify") handleVerifyMenu(event, targetId);
  else handleStudentMenu(event, targetId, true, mode);
}

// --- 行政選擇學生回傳處理 ---
function handleAdminPickStudent(event: any, postbackData: string) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const parts = postbackData.split("&");
  let stuName = "", stuCourse = "", targetId = "", mode = "record";

  for (let i = 0; i < parts.length; i++) {
    if (parts[i].indexOf("n=") === 0) stuName = decodeURIComponent(parts[i].substring(2));
    if (parts[i].indexOf("c=") === 0) stuCourse = decodeURIComponent(parts[i].substring(2));
    if (parts[i].indexOf("tid=") === 0) targetId = parts[i].substring(4);
    if (parts[i].indexOf("m=") === 0) mode = parts[i].substring(2);
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
  if (!userSheet) return;
  const uData = userSheet.getDataRange().getValues();
  let tName = "該位";
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][1] === targetId) {
      tName = uData[i][0];
      break;
    }
  }

  const cache = CacheService.getScriptCache();
  const isPlan = (mode === 'plan');
  const stateObj = {
    status: "WAITING_TIME",
    targetId: targetId,
    sName: stuName,
    sCourse: stuCourse,
    isProxy: true,
    isPlan: isPlan,
    mode: mode
  };
  cache.put(userId, JSON.stringify(stateObj), 300);

  const modeTitle = isPlan ? "📅 [代預排模式]" : "📝 [代填模式]";
  replyLineMessage(replyToken, modeTitle + "\n" + tName + " 講師：" + stuName + " (" + stuCourse + ")\n請輸入時間 (範例 12/28 1000 1200)");
}

// --- 刪除功能：請求刪除 (雙表比對) ---
function processDeleteRequest(event: any) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
  const planSheet = ss.getSheetByName(SHEET_NAME_PLAN);

  const getLastEntry = function(sheet: any) {
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
    if (!userSheet) return null;
    const uData = userSheet.getDataRange().getValues();
    let myName = "";
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][1] === userId) {
        myName = uData[i][0];
        break;
      }
    }
    if (myName === "") return null;
    for (let r = data.length - 1; r >= 1; r--) {
      if (data[r][1] === myName) {
        const d = data[r][2];
        const dStr = (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd") : d;
        return {
          rowIndex: r + 1,
          time: new Date(data[r][0]),
          teacher: data[r][1],
          dateStr: dStr,
          start: data[r][3],
          end: data[r][4],
          duration: data[r][5],
          student: data[r][7],
          course: data[r][8]
        };
      }
    }
    return null;
  };

  const recordLast = getLastEntry(recordSheet);
  const planLast = getLastEntry(planSheet);
  let target: any = null;
  let targetType = "";

  if (recordLast && planLast) {
    if (recordLast.time > planLast.time) {
      target = recordLast;
      targetType = "record";
    } else {
      target = planLast;
      targetType = "plan";
    }
  } else if (recordLast) {
    target = recordLast;
    targetType = "record";
  } else if (planLast) {
    target = planLast;
    targetType = "plan";
  }

  if (target) {
    const typeText = (targetType === "plan") ? "[預排]" : "[已上課]";
    const info = "👤 " + target.teacher + " 講師\n" +
      "📅 " + target.dateStr + " (" + target.duration + "hr)\n" +
      typeText + " " + target.start + " - " + target.end + "\n" +
      "🎓 " + target.student + " | " + target.course;

    replyConfirmButton(replyToken, "⚠️ 確定要刪除這筆 " + typeText + " 紀錄嗎？\n\n" + info, "action=real_delete&sht=" + targetType + "&r=" + target.rowIndex);
  } else {
    replyLineMessage(replyToken, "⚠️ 找不到您的任何紀錄。");
  }
}

// --- 刪除功能：執行刪除 ---
function executeRealDelete(userId: string, replyToken: string, postbackData: string) {
  let targetSheetType = "";
  let rowIndex = -1;
  const parts = postbackData.split("&");

  for (let i = 0; i < parts.length; i++) {
    if (parts[i].indexOf("sht=") === 0) targetSheetType = parts[i].substring(4);
    if (parts[i].indexOf("r=") === 0) rowIndex = parseInt(parts[i].substring(2));
  }

  if (targetSheetType === "" || rowIndex === -1) {
    replyLineMessage(replyToken, "⚠️ 刪除參數錯誤。");
    return;
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetName = (targetSheetType === "plan") ? SHEET_NAME_PLAN : SHEET_NAME_RECORD;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  try {
    sheet.deleteRow(rowIndex);
    replyLineMessage(replyToken, "🗑️ 已刪除該筆 " + (targetSheetType === "plan" ? "預排" : "已上課") + " 紀錄。");
  } catch (e) {
    replyLineMessage(replyToken, "⚠️ 刪除失敗，可能該筆資料已被刪除。");
  }
}

// --- 公告廣播功能 ---
function processBroadcast(event: any, userMsg: string, isGroup: boolean) {
  const replyToken = event.replyToken;
  const userId = event.source.userId;
  let isAdmin = false;

  for (let a = 0; a < ADMIN_LIST.length; a++) {
    if (userId === ADMIN_LIST[a]) {
      isAdmin = true;
      break;
    }
  }

  if (!isAdmin) {
    replyLineMessage(replyToken, "❌ 權限不足！");
    return;
  }

  const keyword = isGroup ? "群組公告" : "發公告";
  const content = userMsg.replace(keyword, "").trim();

  if (content === "") {
    replyLineMessage(replyToken, "❌ 內容是空的！");
    return;
  }

  if (isGroup) {
    try {
      pushLineMessage(GROUP_ID, "📢 【空空_公告】\n" + content);
      replyLineMessage(replyToken, "✅ 已發送到群組！");
    } catch (e) {
      replyLineMessage(replyToken, "❌ 發送失敗，請確認機器人是否在群組內？");
    }
  } else {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAME_TEACHER);
    if (!userSheet) return;
    const data = userSheet.getDataRange().getValues();
    let count = 0;

    for (let i = 1; i < data.length; i++) {
      const teacherId = data[i][1];
      if (teacherId && teacherId.indexOf("U") === 0) {
        try {
          pushLineMessage(teacherId, "📢 【空空_公告】\n" + content);
          count++;
        } catch (e) {}
      }
    }
    replyLineMessage(replyToken, "✅ 公告已發送給 " + count + " 位講師。");
  }
}

// --- ID 查詢工具 ---
function handleMentionQuery(event: any) {
  const mentionees = event.message.mention.mentionees;
  const replyToken = event.replyToken;
  
  if (mentionees && mentionees.length > 0) {
    const targetUserId = mentionees[0].userId;
    if (targetUserId) {
      replyLineMessage(replyToken, "🔍 User ID: " + targetUserId);
    }
  }
}

// --- 關鍵字查詢工具 ---
function handleKeywordQuery(event: any) {
  const replyToken = event.replyToken;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("功能說明");
  
  if (!sheet) {
    replyLineMessage(replyToken, "⚠️ 找不到「功能說明」分頁。");
    return;
  }
  
  const text = sheet.getRange("A1").getValue();
  if (!text || text === "") {
    replyLineMessage(replyToken, "⚠️ 「功能說明」A1 儲存格是空的。");
  } else {
    replyLineMessage(replyToken, text);
  }
}

// --- 渲染主選單 ---
function replyMainMenu(replyToken: string) {
  const flexContent = {
    "type": "bubble",
    "body": {
      "type": "box",
      "layout": "vertical",
      "spacing": "sm",
      "contents": [
        { "type": "text", "text": "🤖 空空BOT 選單", "weight": "bold", "size": "xl", "color": "#1DB446", "align": "center" },
        { "type": "separator", "margin": "md" },
        { "type": "text", "text": "👨‍🏫 講師功能", "weight": "bold", "size": "sm", "margin": "lg", "color": "#555555" },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📝 登錄", "text": "登錄" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📅 預排", "text": "預排" } }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "✅ 核銷", "text": "核銷" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "🗑️ 刪除", "text": "刪除最後一筆" } }
          ]
        },
        { "type": "text", "text": "💰 行政財務與單據", "weight": "bold", "size": "sm", "margin": "lg", "color": "#555555" },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📊 學費試算", "text": "學費試算" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "💸 鐘點試算", "text": "鐘點試算" } }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📄 產生繳費單", "text": "產生繳費單" } }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "🧾 開收據", "text": "開收據" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📨 寄收據", "text": "寄收據" } }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📝 開領據", "text": "開領據" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📧 寄領據", "text": "寄領據" } }
          ]
        },
        { "type": "text", "text": "💰 入會費/常年會費/捐款", "weight": "bold", "size": "sm", "margin": "lg", "color": "#555555" },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "🗂️ 開一般收據", "text": "開一般收據" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📧 寄一般收據", "text": "寄一般收據" } }
          ]
        },
        { "type": "separator", "margin": "lg" },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm", "margin": "md",
          "contents": [
            { "type": "button", "style": "primary", "color": "#2980b9", "height": "sm", "action": { "type": "message", "label": "🏛️ 稅務專區", "text": "稅務" } },
            { "type": "button", "style": "link", "height": "sm", "action": { "type": "message", "label": "🔍 查我的ID", "text": "查ID" } }
          ]
        }
      ]
    }
  };

  replyFlexMessage(replyToken, "完整功能選單", flexContent);
}

// --- 渲染稅務專區選單 ---
function replyTaxMenu(replyToken: string) {
  const flexContent = {
    "type": "bubble",
    "size": "giga",
    "body": {
      "type": "box", "layout": "vertical", "spacing": "md",
      "contents": [
        { "type": "text", "text": "🏛️ 稅務與申報專區", "weight": "bold", "size": "xl", "color": "#2980b9", "align": "center" },
        { "type": "separator" },
        { "type": "text", "text": "✍️ 日常記帳", "weight": "bold", "size": "sm", "color": "#8E44AD" },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm",
          "contents": [
            { "type": "button", "style": "primary", "color": "#8E44AD", "height": "sm", "action": { "type": "message", "label": "📝 記帳指令教學", "text": "記帳" } }
          ]
        },
        { "type": "text", "text": "📊 自動財務三表 (複式簿記結算)", "weight": "bold", "size": "sm", "color": "#27ae60" },
        {
          "type": "box", "layout": "vertical", "spacing": "sm",
          "contents": [
            {
              "type": "box", "layout": "horizontal", "spacing": "sm",
              "contents": [
                { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "📈 損益表", "text": "損益表" } },
                { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "⚖️ 資產負債表", "text": "資產負債表" } }
              ]
            },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "💸 現金流量表", "text": "現金流量表" } }
          ]
        },
        { "type": "text", "text": "🚨 稅務安全監控", "weight": "bold", "size": "sm", "color": "#555555" },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm",
          "contents": [
            { "type": "button", "style": "primary", "color": "#E74C3C", "height": "sm", "action": { "type": "message", "label": "📊 80% 免稅預警", "text": "免稅試算" } }
          ]
        },
        { "type": "text", "text": "📅 年度申報工具", "weight": "bold", "size": "sm", "color": "#555555" },
        {
          "type": "box", "layout": "horizontal", "spacing": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "👨‍🏫 講師扣繳", "text": "年度扣繳" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "💖 捐款上傳", "text": "捐款申報" } }
          ]
        }
      ]
    }
  };

  replyFlexMessage(replyToken, "稅務專區選單", flexContent);
}

// --- 渲染免稅預警儀表板 ---
function replyTaxDashboardCard(token: string, data: any) {
  const formatMoney = function(n: number) { return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","); };
  const statusColor = data.isSafe ? "#1DB446" : "#E74C3C"; 
  const statusText = data.isSafe ? "✅ 支出門檻已達標 (免稅安全)" : "⚠️ 未達 80% 法定門檻！";
  const gapBlock = data.isSafe ? 
      { "type": "text", "text": "恭喜！年度支出已達法定要求", "size": "sm", "color": "#1DB446", "weight": "bold", "wrap": true } :
      { "type": "text", "text": "距免稅門檻還差 $" + formatMoney(data.gap), "size": "md", "color": "#E74C3C", "weight": "bold", "wrap": true };

  const card = {
    "type": "bubble",
    "size": "giga",
    "body": {
      "type": "box",
      "layout": "vertical",
      "contents": [
        { "type": "text", "text": "🏛️ " + data.year + " 年度免稅預警儀表板", "weight": "bold", "size": "xl", "color": "#2C3E50" },
        { "type": "text", "text": statusText, "size": "md", "color": statusColor, "weight": "bold", "margin": "md" },
        { "type": "separator", "margin": "lg" },
        { "type": "text", "text": "💰 總收入分析", "weight": "bold", "size": "md", "color": "#2980B9", "margin": "lg" },
        {
          "type": "box", "layout": "horizontal", "margin": "md",
          "contents": [
            { "type": "text", "text": "學費收入", "size": "sm", "color": "#7F8C8D", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(data.tuition), "size": "sm", "color": "#2C3E50", "align": "end", "flex": 3 }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "margin": "sm",
          "contents": [
            { "type": "text", "text": "捐款與會費", "size": "sm", "color": "#7F8C8D", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(data.donation), "size": "sm", "color": "#2C3E50", "align": "end", "flex": 3 }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "margin": "md",
          "contents": [
            { "type": "text", "text": "合計總收入", "size": "sm", "weight": "bold", "color": "#2980B9", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(data.totalInc), "size": "sm", "weight": "bold", "color": "#2980B9", "align": "end", "flex": 3 }
          ]
        },
        { "type": "separator", "margin": "lg" },
        { "type": "text", "text": "💸 支出與達成率", "weight": "bold", "size": "md", "color": "#D35400", "margin": "lg" },
        {
          "type": "box", "layout": "horizontal", "margin": "md",
          "contents": [
            { "type": "text", "text": "系統講師費", "size": "sm", "color": "#7F8C8D", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(data.teacherExp), "size": "sm", "color": "#2C3E50", "align": "end", "flex": 3 }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "margin": "sm",
          "contents": [
            { "type": "text", "text": "系統外支出", "size": "sm", "color": "#7F8C8D", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(data.externalExp), "size": "sm", "color": "#2C3E50", "align": "end", "flex": 3 }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "margin": "md",
          "contents": [
            { "type": "text", "text": "合計總支出", "size": "sm", "weight": "bold", "color": "#D35400", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(data.totalExp), "size": "sm", "weight": "bold", "color": "#D35400", "align": "end", "flex": 3 }
          ]
        },
        { "type": "separator", "margin": "lg" },
        {
          "type": "box", "layout": "vertical", "margin": "lg", "backgroundColor": "#F4F6F6", "paddingAll": "md", "cornerRadius": "md",
          "contents": [
            { "type": "text", "text": "法定 80% 低標：$" + formatMoney(data.reqExp), "size": "sm", "color": "#34495E", "weight": "bold" },
            { "type": "text", "text": "目前支出比例：" + data.ratio + "%", "size": "md", "color": statusColor, "weight": "bold", "margin": "sm" },
            { "type": "separator", "margin": "md" },
            gapBlock
          ]
        }
      ]
    }
  };

  replyFlexMessage(token, "免稅預警儀表板", card);
}

// --- 📓 記帳傳票預覽卡片 ---
function replyJournalPreviewCard(replyToken: string, state: any) {
  const formatMoney = function(n: number) { return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","); };

  const card = {
    "type": "bubble",
    "body": {
      "type": "box",
      "layout": "vertical",
      "contents": [
        { "type": "text", "text": "📓 記帳傳票預覽", "weight": "bold", "size": "xl", "color": "#8E44AD" },
        { "type": "text", "text": "請核對會計科目與金額是否正確", "size": "xs", "color": "#aaaaaa" },
        { "type": "separator", "margin": "md" },
        { "type": "text", "text": "📅 發生日期：" + state.date, "size": "sm", "margin": "md", "weight": "bold" },
        {
          "type": "box", "layout": "horizontal", "margin": "md",
          "contents": [
            { "type": "text", "text": "(借) " + state.drName, "size": "sm", "color": "#D35400", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(state.amount), "size": "sm", "color": "#2C3E50", "align": "end", "flex": 2 }
          ]
        },
        {
          "type": "box", "layout": "horizontal", "margin": "sm",
          "contents": [
            { "type": "text", "text": "(貸) " + state.crName, "size": "sm", "color": "#2980B9", "flex": 2 },
            { "type": "text", "text": "$" + formatMoney(state.amount), "size": "sm", "color": "#2C3E50", "align": "end", "flex": 2 }
          ]
        },
        { "type": "separator", "margin": "md" },
        { "type": "text", "text": "💡 摘要：" + state.summary, "size": "xs", "color": "#7F8C8D", "margin": "md", "wrap": true }
      ]
    },
    "footer": {
      "type": "box", "layout": "horizontal", "spacing": "sm",
      "contents": [
        { "type": "button", "style": "primary", "color": "#8E44AD", "action": { "type": "postback", "label": "✅ 確認入帳", "data": "action=journal_save_only" } },
        { "type": "button", "style": "secondary", "action": { "type": "postback", "label": "❌ 取消", "data": "action=doc_cancel_journal" } }
      ]
    }
  };

  replyFlexMessage(replyToken, "記帳預覽", card);
}

// --- 📈 創建財務三表卡片結構 ---
function createReportFlexCard(data: any, type: string) {
  const formatMoney = function(n: number) { return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","); };
  let title = "", color = "", blocks: any[] = [];

  if (type === "IS") {
    title = "📊 收支決算表 (損益表)"; color = "#27AE60";
    const netColor = data.net >= 0 ? "#1DB446" : "#E74C3C";
    const netText = data.net >= 0 ? "本期結餘 (淨利)" : "本期短絀 (淨損)";
    blocks = [
      { "type": "box", "layout": "horizontal", "contents": [ { "type": "text", "text": "總收入", "color": "#7F8C8D" }, { "type": "text", "text": "$" + formatMoney(data.revenue), "align": "end", "weight": "bold", "color": "#2980B9" } ] },
      { "type": "box", "layout": "horizontal", "margin": "sm", "contents": [ { "type": "text", "text": "總支出", "color": "#7F8C8D" }, { "type": "text", "text": "$" + formatMoney(data.expense), "align": "end", "weight": "bold", "color": "#D35400" } ] },
      { "type": "separator", "margin": "md" },
      { "type": "box", "layout": "horizontal", "margin": "md", "contents": [ { "type": "text", "text": netText, "weight": "bold" }, { "type": "text", "text": "$" + formatMoney(data.net), "align": "end", "weight": "bold", "color": netColor } ] }
    ];
  } else if (type === "BS") {
    title = "⚖️ 資產負債表"; color = "#8E44AD";
    const balanceCheck = (data.assets === data.equity) ? "✅ 借貸平衡" : "❌ 借貸不平";
    blocks = [
      { "type": "box", "layout": "horizontal", "contents": [ { "type": "text", "text": "資產總計", "color": "#7F8C8D" }, { "type": "text", "text": "$" + formatMoney(data.assets), "align": "end", "weight": "bold", "color": "#2C3E50" } ] },
      { "type": "box", "layout": "horizontal", "margin": "sm", "contents": [ { "type": "text", "text": "負債總計", "color": "#7F8C8D" }, { "type": "text", "text": "$" + formatMoney(data.liab), "align": "end", "weight": "bold", "color": "#E74C3C" } ] },
      { "type": "separator", "margin": "md" },
      { "type": "box", "layout": "horizontal", "margin": "md", "contents": [ { "type": "text", "text": "負債及淨值總計", "weight": "bold" }, { "type": "text", "text": "$" + formatMoney(data.equity), "align": "end", "weight": "bold", "color": "#2C3E50" } ] },
      { "type": "text", "text": balanceCheck, "size": "xs", "align": "end", "color": "#BDC3C7", "margin": "sm" }
    ];
  } else if (type === "CF") {
    title = "💸 現金流量表"; color = "#E67E22";
    const netColor = data.netCash >= 0 ? "#1DB446" : "#E74C3C";
    blocks = [
      { "type": "box", "layout": "horizontal", "contents": [ { "type": "text", "text": "現金流入", "color": "#7F8C8D" }, { "type": "text", "text": "$" + formatMoney(data.inflow), "align": "end", "weight": "bold", "color": "#27AE60" } ] },
      { "type": "box", "layout": "horizontal", "margin": "sm", "contents": [ { "type": "text", "text": "現金流出", "color": "#7F8C8D" }, { "type": "text", "text": "$" + formatMoney(data.outflow), "align": "end", "weight": "bold", "color": "#E74C3C" } ] },
      { "type": "separator", "margin": "md" },
      { "type": "box", "layout": "horizontal", "margin": "md", "contents": [ { "type": "text", "text": "淨現金變動", "weight": "bold" }, { "type": "text", "text": "$" + formatMoney(data.netCash), "align": "end", "weight": "bold", "color": netColor } ] }
    ];
  }

  return {
    "type": "bubble",
    "body": {
      "type": "box", "layout": "vertical",
      "contents": [
        { "type": "text", "text": title, "weight": "bold", "size": "xl", "color": color },
        { "type": "text", "text": "期間：" + data.period, "size": "sm", "color": "#95A5A6", "margin": "sm" },
        { "type": "separator", "margin": "md" },
        { "type": "box", "layout": "vertical", "margin": "md", "contents": blocks },
        { "type": "separator", "margin": "lg" },
        { "type": "button", "style": "link", "height": "sm", "action": { "type": "uri", "label": "📄 查看雲端明細報表", "uri": data.url } }
      ]
    }
  };
}

