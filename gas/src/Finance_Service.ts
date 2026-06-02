// ==========================================
// 💰 Finance_Service.ts : 財務計算、存檔與稅務引擎 (TypeScript + LIFF 整合重構版)
// ==========================================

// ==========================================
// 1. LIFF 專屬行政管理 API 後端實作
// ==========================================

function handleLiffAdminTask(params: any) {
  const lineUserId = params.lineUserId;
  const task = params.task; // '学费试算' | '钟点试算' | '产生缴费单' | '寄领据'
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
    
    if (task === "寄領據" || task === "寄领据") {
      handleBatchSendAllowanceEmailCommand(mockEvent, "寄領據 " + month);
      return { ok: true, message: `✅ ${month} 講師領據 Email 已發送完畢！` };
    }

    return { ok: false, message: `不支援的行政任務: ${task}` };
  } catch (e) {
    return { ok: false, message: "執行失敗：" + e.toString() };
  }
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
    if (!stats[s][c]) { stats[s][c] = { teacher: t, fee: fee, mode: mode, recordBase: 0, planBase: 0, planNext: 0, detailsRec: [], detailsPlan: [] }; }
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
          stats[sName][cName].planBase += hr; 
          if (status === "取消") {
            const dText = (pData[i][2] instanceof Date) ? Utilities.formatDate(pData[i][2], timeZone, "MM/dd") : pData[i][2];
            stats[sName][cName].detailsRec.push("[歷史補救] " + dText + " " + pData[i][3] + "-" + pData[i][4] + " (取消退費)");
          }
        } else if (lessonDateMonth == nextMonthStr && status !== "取消") {
          stats[sName][cName].planNext += hr; const perLessonAmt = Math.round(hr * conf.fee); const dText = (pData[i][2] instanceof Date) ? Utilities.formatDate(pData[i][2], timeZone, "MM/dd") : pData[i][2];
          stats[sName][cName].detailsPlan.push("[預收] " + dText + " " + pData[i][3] + "-" + pData[i][4] + " ($" + perLessonAmt + ")");
        }
      }
    }
  }

  let report = "💰 學費試算單 (" + baseMonthStr + ")\n\n"; let saveData: any[] = []; let grandTotal = 0; let hasData = false;
  for (const sName in stats) {
    let sTotal = 0; let sDetailText = ""; const courseRows = [];
    for (const cName in stats[sName]) {
      const item = stats[sName][cName]; let finalAmount = 0; let formulaStr = ""; let fullDetails: string[] = [];
      if (item.mode === "預收") {
        const diff = Math.round((item.recordBase - item.planBase) * 10) / 10;
        const totalHr = item.planNext + diff; finalAmount = Math.round(totalHr * item.fee);
        const diffStr = (diff > 0) ? (" +補" + diff + "hr") : (diff < 0 ? (" -退" + Math.abs(diff) + "hr") : "");
        formulaStr = "預收" + item.planNext + "hr" + diffStr + " = " + totalHr + "hr";
        if (item.detailsRec.length > 0 || item.planBase > 0) {
            fullDetails.push("📋 [上月核對] 實上" + item.recordBase + " / 預繳" + item.planBase); fullDetails = fullDetails.concat(item.detailsRec); 
            if (diff !== 0) fullDetails.push("⚠️ 差異金額: " + (diff > 0 ? "補收 $":"退費 $") + Math.abs(Math.round(diff * item.fee))); else fullDetails.push("✅ 差異: 無 (已結清)"); fullDetails.push(""); 
        }
        if (item.detailsPlan.length > 0) { fullDetails.push("📅 [下月預收] " + nextMonthStr); fullDetails = fullDetails.concat(item.detailsPlan); }
      } else {
        finalAmount = Math.round(item.recordBase * item.fee); formulaStr = "實上 " + item.recordBase + " hr × $" + item.fee; fullDetails = item.detailsRec;
      }
      if (finalAmount !== 0 || item.recordBase > 0 || item.planNext > 0) {
        sTotal += finalAmount; const detailBlock = fullDetails.join("\n"); if (sDetailText !== "") sDetailText += "\n--------------------\n";
        sDetailText += "   📘 " + cName + " (" + item.mode + ")\n" + detailBlock.replace(/^/gm, "      ") + "\n      ➤ 本科總計：$" + finalAmount;
        courseRows.push([ baseMonthStr, sName, cName, item.mode, detailBlock, formulaStr, item.fee, finalAmount, "", "", "" ]);
      }
    }
    if (courseRows.length > 0) {
      courseRows[courseRows.length - 1][8] = sTotal; saveData = saveData.concat(courseRows); report += "🎓 " + sName + "\n" + sDetailText + "\n   💰 本期應繳：$" + sTotal + "\n\n";
      grandTotal += sTotal; hasData = true;
    }
  }

  if (!hasData) { replyLineMessage(replyToken, "💰 學費試算 (" + baseMonthStr + ")\n無須結算資料。"); } else {
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
  for (let i = 1; i < courseData.length; i++) { const key = courseData[i][2] + "_" + courseData[i][3]; const fee = courseData[i][4]; const ratio = courseData[i][5]; if (fee && ratio) { profitMap[key] = { fee: fee, ratio: ratio }; } }

  const tData = teacherSheet.getDataRange().getValues(); const taxConfigMap: any = {};
  for (let i = 1; i < tData.length; i++) { const tName = tData[i][0]; if (tName) { taxConfigMap[tName] = { formatCode: tData[i][11] || "9B", nationality: tData[i][12] || "本國人", nhiExempt: tData[i][13] || "否" }; } }

  const salaryStats: any = {}; const rData = recordSheet.getDataRange().getValues(); const updateRows: number[] = [];
  for (let i = 1; i < rData.length; i++) {
    const rowMonth = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "yyyy/MM") : String(rData[i][2]).substring(0, 7);
    const settled = rData[i][10]; const courseName = rData[i][8];
    if (courseName.indexOf("取消") === -1 && rowMonth == queryMonth && (!settled || settled === "")) {
      const tName = rData[i][1]; const sName = rData[i][7]; const hr = parseFloat(rData[i][5]); const key = sName + "_" + courseName; const conf = profitMap[key] || { fee: 0, ratio: 0 };
      const payRate = conf.fee * conf.ratio; const payAmount = Math.round(hr * payRate); 
      if (!salaryStats[tName]) salaryStats[tName] = { total: 0, details: [] };
      const dText = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "MM/dd") : rData[i][2];
      salaryStats[tName].details.push(sName + "(" + courseName + ") " + dText + " " + rData[i][3] + "-" + rData[i][4] + " ($" + payAmount + ")");
      salaryStats[tName].total += payAmount; updateRows.push(i + 1);
    }
  }

  let report = "💰 鐘點費試算 (" + queryMonth + ")\n\n"; const saveData: any[] = []; let hasData = false;
  for (const tName in salaryStats) {
    const item = salaryStats[tName]; const total = item.total; const taxConfig = taxConfigMap[tName] || { formatCode: "9B", nationality: "本國人", nhiExempt: "否" };
    let taxAmount = 0; let nhiAmount = 0;
    if (taxConfig.nationality.indexOf("外籍") > -1) { taxAmount = Math.round(total * 0.18); } else { if (taxConfig.formatCode === "9B" && total > 20000) { taxAmount = Math.round(total * 0.10); } else if (taxConfig.formatCode === "50" && total >= 86001) { taxAmount = Math.round(total * 0.05); } }
    if (total >= 20000 && taxConfig.nhiExempt !== "是") { nhiAmount = Math.round(total * 0.0211); }
    const netAmount = total - taxAmount - nhiAmount; 

    report += "👨‍🏫 " + tName + " (" + taxConfig.formatCode + ")\n"; const detailStr = item.details.join("\n"); report += detailStr.replace(/^/gm, "   ") + "\n";
    report += "   💵 應付總額：$" + total + "\n"; if (taxAmount > 0) report += "   ➖ 扣繳稅額：$" + taxAmount + "\n"; if (nhiAmount > 0) report += "   ➖ 補充保費：$" + nhiAmount + "\n"; report += "   💰 實發金額：$" + netAmount + "\n--------------------\n";
    
    saveData.push([ queryMonth, tName, taxConfig.formatCode, taxConfig.nationality, detailStr, total, netAmount, "", "", "", "", "", taxAmount, nhiAmount, netAmount ]);
    hasData = true;
  }

  if (!hasData) { replyLineMessage(replyToken, "💰 鐘點費試算 (" + queryMonth + ")\n無須結算資料。"); } else {
    report += "請確認是否寫入結算工作表？";
    const cacheKey = "FIN_" + userId; const cacheData = { targetSheet: SHEET_NAME_FIN_PAY, save: saveData, updateTargetMonth: queryMonth, updateRows: updateRows, category: "鐘點費", prefix: "P" };
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(cacheData), 600); replyConfirmationCard(replyToken, "鐘點費試算確認", report, cacheKey);
  }
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
    
    // (A) 產生主檔的給付編號首碼字軌，例如 "R2605_"
    const rocYearShort = (now.getFullYear() - 1911).toString().substring(1, 3);
    const mm = (now.getMonth() + 1 < 10 ? "0" : "") + (now.getMonth() + 1);
    const invoicePrefix = cacheObj.prefix + rocYearShort + mm + "_";
    
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
        const serial = parseInt(rowInvoice.split("_")[1]);
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
      // 學費回溯標記：掃描並標記已結算的學費
      const recordSheet = ss.getSheetByName(SHEET_NAME_RECORD);
      if (recordSheet) {
        const rData = recordSheet.getDataRange().getValues();
        for (let i = 1; i < rData.length; i++) {
          const rowMonth = (rData[i][2] instanceof Date) ? Utilities.formatDate(rData[i][2], timeZone, "yyyy/MM") : String(rData[i][2]).substring(0, 7);
          const settled = rData[i][9];
          if (rowMonth == cacheObj.updateTargetMonth && (!settled || settled === "")) {
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
          
          if (status !== "取消" && lessonDateMonth == cacheObj.updateTargetMonth && (!feeSettled || feeSettled === "")) {
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

function getFinancialDuplicateKeys(rows: any[], category: string): string[] {
  const keys: string[] = [];
  for (let i = 0; i < rows.length; i++) {
    const key = getFinancialRowKey(rows[i], category);
    if (key && keys.indexOf(key) === -1) keys.push(key);
  }
  return keys;
}

function getFinancialRowKey(row: any[], category: string): string {
  if (category === "學費") return String(row[1] || "").trim(); // B欄：學生姓名
  if (category === "鐘點費") return String(row[1] || "").trim(); // B欄：講師姓名
  return "";
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
declare function replyLineMessage(token: string, msg: string): void;
declare function replyConfirmationCard(token: string, title: string, report: string, key: string): void;

