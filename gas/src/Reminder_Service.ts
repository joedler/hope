// ==========================================
// ⏰ Reminder_Service.ts : 講師每月授課登記提醒
// ==========================================

var COURSE_REMINDER_TRIGGER_FUNCTION = "runMonthlyCourseRegistrationReminder";
var COURSE_REMINDER_PENDING_INSTALL_KEY = "COURSE_REMINDER_PENDING_INSTALL";
var COURSE_REMINDER_PENDING_TEST_KEY = "COURSE_REMINDER_PENDING_TEST";
var COURSE_REMINDER_LAST_SENT_KEY = "COURSE_REMINDER_LAST_SENT_MONTH";
var COURSE_REMINDER_TIME_ZONE = "Asia/Taipei";

function getMonthlyCourseReminderConfig() {
  const props = PropertiesService.getScriptProperties();
  const configuredDay = parseInt(props.getProperty("COURSE_REMINDER_DAY") || "30", 10);
  const configuredHour = parseInt(props.getProperty("COURSE_REMINDER_HOUR") || "20", 10);
  return {
    groupId: props.getProperty("COURSE_REMINDER_GROUP_ID") || GROUP_ID,
    day: isNaN(configuredDay) ? 30 : Math.max(1, Math.min(31, configuredDay)),
    hour: isNaN(configuredHour) ? 20 : Math.max(0, Math.min(23, configuredHour)),
    timeZone: COURSE_REMINDER_TIME_ZONE,
    message: props.getProperty("COURSE_REMINDER_MESSAGE") ||
      "📢 【空空_公告】\n空空來了！空空來囉~~\n\n感謝老師們完成本月的陪伴服務，有您真好 ❤️\n\n空空小幫手在此提醒老師們抽空完成新月份預排、本月份課程登錄及核銷 ，感謝您的每一步撐出更多盼望❤️"
  };
}

function getMonthlyCourseReminderPreviewText(): string {
  const config = getMonthlyCourseReminderConfig();
  return [
    "【每月授課登記提醒－排程預覽】",
    "群組 ID：" + config.groupId,
    "日期：每月 " + config.day + " 日；若當月不足 " + config.day + " 日，改為當月最後一天",
    "時間：台灣時間 " + String(config.hour).padStart(2, "0") + ":00 所在時段（GAS 可能於該小時內執行）",
    "防重複：同一月份成功推播後不再重送",
    "訊息：\n" + config.message
  ].join("\n");
}

/**
 * 第一步：在 GAS 編輯器執行本函式，檢查記錄中的完整預覽。
 * 預覽有效 30 分鐘；確認內容無誤後才可執行 confirmInstallMonthlyCourseReminder。
 */
function previewMonthlyCourseReminderInstall() {
  const props = PropertiesService.getScriptProperties();
  const preview = getMonthlyCourseReminderPreviewText();
  props.setProperty(COURSE_REMINDER_PENDING_INSTALL_KEY, JSON.stringify({
    preview: preview,
    createdAt: new Date().getTime()
  }));
  Logger.log(preview);
  Logger.log("確認內容無誤後，請於 30 分鐘內執行 confirmInstallMonthlyCourseReminder()。此步驟不會推播訊息。");
  return preview;
}

/** 第二步：確認預覽後建立唯一的每日檢查觸發器。 */
function confirmInstallMonthlyCourseReminder() {
  const props = PropertiesService.getScriptProperties();
  const pendingText = props.getProperty(COURSE_REMINDER_PENDING_INSTALL_KEY);
  if (!pendingText) throw new Error("尚未預覽；請先執行 previewMonthlyCourseReminderInstall()。 ");

  const pending = JSON.parse(pendingText);
  const ageMs = new Date().getTime() - Number(pending.createdAt || 0);
  if (ageMs < 0 || ageMs > 30 * 60 * 1000) {
    props.deleteProperty(COURSE_REMINDER_PENDING_INSTALL_KEY);
    throw new Error("預覽已超過 30 分鐘，請重新預覽後再確認。");
  }
  if (pending.preview !== getMonthlyCourseReminderPreviewText()) {
    props.deleteProperty(COURSE_REMINDER_PENDING_INSTALL_KEY);
    throw new Error("提醒設定已變更，請重新預覽後再確認。");
  }

  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === COURSE_REMINDER_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const config = getMonthlyCourseReminderConfig();
  ScriptApp.newTrigger(COURSE_REMINDER_TRIGGER_FUNCTION)
    .timeBased()
    .atHour(config.hour)
    .everyDays(1)
    .inTimezone(config.timeZone)
    .create();

  props.deleteProperty(COURSE_REMINDER_PENDING_INSTALL_KEY);
  Logger.log("已建立每月授課登記提醒。每日於指定小時檢查，只有目標日期才會推播。");
  return { ok: true, preview: pending.preview };
}

/**
 * 測試第一步：預覽單次測試推播；有效 30 分鐘，不會發送 LINE。
 */
function previewMonthlyCourseReminderTest() {
  const config = getMonthlyCourseReminderConfig();
  const preview = [
    "【每月授課登記提醒－單次測試推播預覽】",
    "群組 ID：" + config.groupId,
    "注意：此為立即測試，不會寫入正式月份的已發送標記。",
    "訊息：\n" + config.message
  ].join("\n");
  PropertiesService.getScriptProperties().setProperty(COURSE_REMINDER_PENDING_TEST_KEY, JSON.stringify({
    preview: preview,
    groupId: config.groupId,
    message: config.message,
    createdAt: new Date().getTime()
  }));
  Logger.log(preview);
  Logger.log("確認群組與內容無誤後，請於 30 分鐘內執行 confirmSendMonthlyCourseReminderTest()。此步驟尚未推播。");
  return preview;
}

/**
 * 測試第二步：確認最近一次預覽後，立即推播一次；不影響正式月份防重複標記。
 */
function confirmSendMonthlyCourseReminderTest() {
  const props = PropertiesService.getScriptProperties();
  const pendingText = props.getProperty(COURSE_REMINDER_PENDING_TEST_KEY);
  if (!pendingText) throw new Error("尚未預覽測試訊息；請先執行 previewMonthlyCourseReminderTest()。");

  const pending = JSON.parse(pendingText);
  const ageMs = new Date().getTime() - Number(pending.createdAt || 0);
  const config = getMonthlyCourseReminderConfig();
  if (ageMs < 0 || ageMs > 30 * 60 * 1000) {
    props.deleteProperty(COURSE_REMINDER_PENDING_TEST_KEY);
    throw new Error("測試預覽已超過 30 分鐘，請重新預覽後再確認。");
  }
  if (pending.groupId !== config.groupId || pending.message !== config.message) {
    props.deleteProperty(COURSE_REMINDER_PENDING_TEST_KEY);
    throw new Error("群組或訊息內容已變更，請重新預覽後再確認。");
  }

  const response = LineClient.push(config.groupId, [{ type: "text", text: config.message }]);
  if (!response || response.getResponseCode() !== 200) {
    throw new Error("LINE 測試推播失敗，請查看執行記錄中的 LINE API 回應。");
  }
  props.deleteProperty(COURSE_REMINDER_PENDING_TEST_KEY);
  Logger.log("測試推播成功；未寫入正式月份已發送標記。群組：" + config.groupId);
  return { ok: true, sent: true, groupId: config.groupId };
}

/** 時間觸發器入口；每天檢查，僅在 min(設定日, 當月最後一天) 推播一次。 */
function runMonthlyCourseRegistrationReminder() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("提醒工作已有另一執行程序，略過本次。");
    return { ok: false, skipped: "locked" };
  }

  try {
    const config = getMonthlyCourseReminderConfig();
    const now = new Date();
    const year = Number(Utilities.formatDate(now, config.timeZone, "yyyy"));
    const month = Number(Utilities.formatDate(now, config.timeZone, "M"));
    const day = Number(Utilities.formatDate(now, config.timeZone, "d"));
    const hour = Number(Utilities.formatDate(now, config.timeZone, "H"));
    const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
    const targetDay = Math.min(config.day, lastDay);
    const monthKey = Utilities.formatDate(now, config.timeZone, "yyyy-MM");
    const props = PropertiesService.getScriptProperties();

    if (day !== targetDay || hour !== config.hour) {
      return { ok: true, skipped: "not_target_time", targetDay: targetDay };
    }
    if (props.getProperty(COURSE_REMINDER_LAST_SENT_KEY) === monthKey) {
      return { ok: true, skipped: "already_sent", month: monthKey };
    }

    const response = LineClient.push(config.groupId, [{ type: "text", text: config.message }]);
    if (!response || response.getResponseCode() !== 200) {
      throw new Error("LINE 群組提醒推播失敗；未寫入本月已發送標記。");
    }
    props.setProperty(COURSE_REMINDER_LAST_SENT_KEY, monthKey);
    Logger.log("每月授課登記提醒已推播：" + monthKey + "，群組：" + config.groupId);
    return { ok: true, sent: true, month: monthKey };
  } finally {
    lock.releaseLock();
  }
}

/** 唯讀狀態檢查，不建立、不刪除觸發器，也不發送 LINE。 */
function inspectMonthlyCourseReminder() {
  const handlers = ScriptApp.getProjectTriggers().map(function(trigger) {
    return trigger.getHandlerFunction();
  });
  const result = {
    preview: getMonthlyCourseReminderPreviewText(),
    installedTriggerCount: handlers.filter(function(name) {
      return name === COURSE_REMINDER_TRIGGER_FUNCTION;
    }).length,
    lastSentMonth: PropertiesService.getScriptProperties().getProperty(COURSE_REMINDER_LAST_SENT_KEY) || "尚未發送"
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
