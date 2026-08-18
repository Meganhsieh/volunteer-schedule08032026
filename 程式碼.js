// ================================================
// 主要進入點與核心函數
// 版本：v1.4
// v1.1：純讀取改用 getSheetData 快取；寫入後補 invalidateSheetData
//       管理者通知改為一次寄送（每筆申請配額 4 → 2）
// v1.2：新增系統參數 get/save（優先期時數、取消截止天數、自助取消月額度）
// v1.3：排班規則到期日（C欄，純顯示用）
// v1.4：移除公告到期日；移除公告「自動月份」模式；
//       修正 generateBinocularSlots 傳參不一致；新增 getQuotaPageData 合併查詢
// ================================================

function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : "index";
  var template;
  switch(page) {
    case "form":
      template = HtmlService.createTemplateFromFile("Form");
      template.activityName = e.parameter.activity || "";
      var act = getActivityConfig(e.parameter.activity || "");
      template.activityType = act ? act.type : "";
      break;
    case "query":
      template = HtmlService.createTemplateFromFile("Query");
      break;
    case "admin":
      template = HtmlService.createTemplateFromFile("Admin");
      break;
    default:
      template = HtmlService.createTemplateFromFile("Index");
  }
  return template.evaluate()
    .setTitle("志工排班系統")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function validateVolunteer(name, id) {
  var data = getSheetData(SHEET.ROSTER);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].trim() == id.trim() && data[i][1].trim() == name.trim()) {
      return {
        id:        data[i][0].trim(),
        name:      data[i][1].trim(),
        type:      data[i][2].trim(),
        teleCert:  data[i][3].trim(),
        venueCert: data[i][4].trim(),
        email:     (data[i][5] || "").trim()
      };
    }
  }
  return null;
}

function getVolunteerNameById(id) {
  var data = getSheetData(SHEET.ROSTER);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].trim() == id.trim()) return data[i][1].trim();
  }
  return null;
}

function getVolunteerById(id) {
  var data = getSheetData(SHEET.ROSTER);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].trim() == id.trim()) {
      return {
        id:        data[i][0].trim(),
        name:      data[i][1].trim(),
        type:      data[i][2].trim(),
        teleCert:  data[i][3].trim(),
        venueCert: data[i][4].trim(),
        email:     (data[i][5] || "").trim()
      };
    }
  }
  return null;
}

function processVolunteerRequest(name, volunteerId, activityName, scheduleSlot, operation, notes) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if (operation !== "申請排班") return "操作失敗：此系統目前僅開放申請排班，如需取消請聯繫管理人員。";
    var volunteer = validateVolunteer(name, volunteerId);
    if (!volunteer) return "申請失敗：姓名與編號不符。";
    var activity = getActivityConfig(activityName);
    if (!activity) return "申請失敗：找不到活動設定。";
    var result;
    if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
      result = processBinocularRequest(volunteer, activity, scheduleSlot, notes);
    } else {
      result = processSessionRequest(volunteer, activity, scheduleSlot, notes);
    }
    if (result.includes("成功")) {
      if (volunteer.email) sendVolunteerNotification(volunteer, activityName, scheduleSlot);
      sendAdminNotification(volunteer, activityName, scheduleSlot, result);
    }
    return result;
  } catch(e) {
    return "系統繁忙：" + e.toString();
  } finally {
    lock.releaseLock();
  }
}

function processVolunteerRequestBatch(name, volunteerId, activityName, slots, operation, notes) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if (operation !== "申請排班") return { results: [], summary: "操作失敗：此系統目前僅開放申請排班，如需取消請聯繫管理人員。" };
    var volunteer = validateVolunteer(name, volunteerId);
    if (!volunteer) return { results: [], summary: "申請失敗：姓名與編號不符。" };
    var activity = getActivityConfig(activityName);
    if (!activity) return { results: [], summary: "申請失敗：找不到活動設定。" };
    if (slots.length > 1) {
      var firstDate = slots[0].split(" ")[0];
      for (var i = 1; i < slots.length; i++) {
        if (slots[i].split(" ")[0] !== firstDate) return { results: [], summary: "申請失敗：僅允許同一天複選時段。" };
      }
    }
    var results = [], successCount = 0, failCount = 0;
    for (var j = 0; j < slots.length; j++) {
      var slot = slots[j], result;
      if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
        result = processBinocularRequest(volunteer, activity, slot, notes);
      } else {
        result = processSessionRequest(volunteer, activity, slot, notes);
      }
      results.push({ slot: slot, result: result });
      if (result.includes("成功")) { successCount++; } else { failCount++; }
    }
    var successSlots = results.filter(function(r) { return r.result.includes("成功"); }).map(function(r) { return r.slot; });
    if (successSlots.length > 0) {
      if (volunteer.email) sendVolunteerNotificationBatch(volunteer, activityName, successSlots);
      sendAdminNotificationBatch(volunteer, activityName, successSlots, activity);
    }
    return { results: results, summary: "成功 " + successCount + " 筆，失敗 " + failCount + " 筆。" };
  } catch(e) {
    return { results: [], summary: "系統繁忙：" + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ================================================
// 通知輔助函數
// ================================================

function getActivityCalendarLink(activityName) {
  var activity = getActivityConfig(activityName);
  if (!activity || !activity.calendarId) return "（未設定月曆）";
  return "https://calendar.google.com/calendar/embed?src=" +
    encodeURIComponent(activity.calendarId) + "&ctz=Asia%2FTaipei";
}

function getSlotInfo(activityName, scheduleSlot) {
  var activity = getActivityConfig(activityName);
  var defaultInfo = { current: 0, max: 0, status: "未知" };
  if (!activity) return defaultInfo;
  var slotParts = scheduleSlot.split(" ");
  if (slotParts.length < 2) return defaultInfo;
  var dateStr = slotParts[0], timeSlot = slotParts[1];
  if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
    var overviewData = getSheetData(SHEET.OVERVIEW);
    for (var i = 1; i < overviewData.length; i++) {
      if (overviewData[i][0] === activityName && overviewData[i][1] === dateStr && overviewData[i][2] === timeSlot) {
        var current = (overviewData[i][3] ? 1 : 0) + (overviewData[i][5] ? 1 : 0) +
                      (overviewData[i][7] ? 1 : 0) + (overviewData[i][9] ? 1 : 0);
        var max     = getBinocularMax(dateStr, timeSlot, activityName);
        var status  = current >= max ? "已滿" : current > 0 ? "部分排班" : "可排班";
        return { current: current, max: max, status: status };
      }
    }
    return { current: 0, max: getBinocularMax(dateStr, timeSlot, activityName), status: "可排班" };
  } else {
    var sessionConfig = getSessionConfig(activityName, dateStr, timeSlot);
    if (!sessionConfig) return defaultInfo;
    var gc = countSessionBookings(dateStr, timeSlot, activityName, VOLUNTEER_TYPE.GENERAL);
    var nc = countSessionBookings(dateStr, timeSlot, activityName, VOLUNTEER_TYPE.NEW);
    var total = gc + nc, maxT = sessionConfig.generalMax + sessionConfig.newMax;
    return { current: total, max: maxT, status: total >= maxT ? "已滿" : total > 0 ? "部分排班" : "可排班" };
  }
}

function buildSubject(volunteer, activityName, slots) {
  var dateLabel, timeLabel;
  if (typeof slots === "string") {
    var parts = slots.split(" "); dateLabel = parts[0]; timeLabel = parts.length >= 2 ? parts[1] : "";
  } else if (slots.length === 1) {
    var parts1 = slots[0].split(" "); dateLabel = parts1[0]; timeLabel = parts1.length >= 2 ? parts1[1] : "";
  } else {
    dateLabel = slots[0].split(" ")[0];
    timeLabel = slots.map(function(s) { var p = s.split(" "); return p.length >= 2 ? p[1] : s; }).join("、");
  }
  return "【志工排班系統】通知_" + volunteer.id + "_" + volunteer.name + "_申請_" + activityName + "_" + dateLabel + "_" + timeLabel;
}

function sendVolunteerNotification(volunteer, activityName, scheduleSlot) {
  try {
    var subject = buildSubject(volunteer, activityName, scheduleSlot);
    var calendarLink = getActivityCalendarLink(activityName);
    var body = volunteer.name + " 您好，\n\n您已成功完成以下排班申請：\n\n  活動：" + activityName + "\n  時段：" + scheduleSlot + "\n\n如需取消排班，請聯繫管理人員。\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + calendarLink + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(volunteer.email, subject, body);
  } catch(e) { Logger.log("發送志工通知失敗：" + e.toString()); }
}

function sendVolunteerNotificationBatch(volunteer, activityName, successSlots) {
  try {
    var subject = buildSubject(volunteer, activityName, successSlots);
    var calendarLink = getActivityCalendarLink(activityName);
    var body = volunteer.name + " 您好，\n\n您已成功完成以下排班申請：\n\n  活動：" + activityName + "\n  時段：\n";
    for (var i = 0; i < successSlots.length; i++) body += "    " + (i+1) + ". " + successSlots[i] + "\n";
    body += "\n如需取消排班，請聯繫管理人員。\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + calendarLink + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(volunteer.email, subject, body);
  } catch(e) { Logger.log("發送志工批次通知失敗：" + e.toString()); }
}

function sendAdminNotification(volunteer, activityName, scheduleSlot, result) {
  try {
    var adminEmails = getAdminEmails();
    if (adminEmails.length === 0) return;
    var subject = buildSubject(volunteer, activityName, scheduleSlot);
    var slotInfo = getSlotInfo(activityName, scheduleSlot);
    var calendarLink = getActivityCalendarLink(activityName);
    var body = "管理者您好，\n\n有志工完成排班申請，詳情如下：\n\n  志工姓名：" + volunteer.name + "\n  志工編號：" + volunteer.id + "\n  活動：" + activityName + "\n  時段：" + scheduleSlot + "\n  結果：" + result + "\n\n─── 時段目前狀態 ───\n  已報名：" + slotInfo.current + " / " + slotInfo.max + " 人\n  尚缺：" + Math.max(0, slotInfo.max - slotInfo.current) + " 人\n  狀態：" + slotInfo.status + "\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + calendarLink + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(adminEmails.join(","), subject, body);   // ★ 一次寄送
  } catch(e) { Logger.log("發送管理者通知失敗：" + e.toString()); }
}

function sendAdminNotificationBatch(volunteer, activityName, successSlots, activity) {
  try {
    var adminEmails = getAdminEmails();
    if (adminEmails.length === 0) return;
    var subject = buildSubject(volunteer, activityName, successSlots);
    var calendarLink = getActivityCalendarLink(activityName);
    var body = "管理者您好，\n\n有志工完成批次排班申請，詳情如下：\n\n  志工姓名：" + volunteer.name + "\n  志工編號：" + volunteer.id + "\n  活動：" + activityName + "\n\n─── 各時段狀態 ───\n";
    for (var k = 0; k < successSlots.length; k++) {
      var slotInfo = getSlotInfo(activityName, successSlots[k]);
      body += "  " + (k+1) + ". " + successSlots[k] + "\n     已報名：" + slotInfo.current + " / " + slotInfo.max + " 人　尚缺：" + Math.max(0, slotInfo.max - slotInfo.current) + " 人（" + slotInfo.status + "）\n";
    }
    body += "\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + calendarLink + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(adminEmails.join(","), subject, body);   // ★ 一次寄送
  } catch(e) { Logger.log("發送管理者批次通知失敗：" + e.toString()); }
}

function formatNow() {
  var d = new Date(), mm = d.getMonth()+1, dd = d.getDate(), hh = d.getHours(), mi = d.getMinutes();
  return d.getFullYear() + "/" + mm + "/" + dd + " " + (hh<10?"0"+hh:hh) + ":" + (mi<10?"0"+mi:mi);
}

function sendVolunteerCancelNotification(volunteer, activityName, scheduleSlot, adminEmail, cancelTime) {
  if (!volunteer.email) return;
  try {
    var calendarLink = getActivityCalendarLink(activityName);
    var parts = scheduleSlot.split(" ");
    var subject = "【志工排班系統】通知_" + volunteer.id + "_" + volunteer.name + "_取消_" + activityName + "_" + parts[0] + "_" + (parts[1]||"");
    var body = volunteer.name + " 您好，\n\n您的以下排班已由管理人員取消，請留意：\n\n  活動：" + activityName + "\n  時段：" + scheduleSlot + "\n  取消時間：" + cancelTime + "\n\n如有疑問，請直接聯繫管理人員。\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + calendarLink + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(volunteer.email, subject, body);
  } catch(e) { Logger.log("發送志工取消通知失敗：" + e.toString()); }
}

function sendAdminCancelNotification(volunteer, activityName, scheduleSlot, adminEmail, cancelTime) {
  try {
    var adminEmails = getAdminEmails();
    if (adminEmails.length === 0) return;
    var calendarLink = getActivityCalendarLink(activityName);
    var slotInfo = getSlotInfo(activityName, scheduleSlot);
    var parts = scheduleSlot.split(" ");
    var subject = "【志工排班系統】通知_" + volunteer.id + "_" + volunteer.name + "_取消_" + activityName + "_" + parts[0] + "_" + (parts[1]||"");
    var body = "管理者您好，\n\n已由管理者執行取消排班，詳情如下：\n\n  志工姓名：" + volunteer.name + "\n  志工編號：" + volunteer.id + "\n  活動：" + activityName + "\n  時段：" + scheduleSlot + "\n  操作者：" + adminEmail + "\n  取消時間：" + cancelTime + "\n\n─── 時段目前狀態（取消後）───\n  已報名：" + slotInfo.current + " / " + slotInfo.max + " 人\n  尚缺：" + Math.max(0, slotInfo.max - slotInfo.current) + " 人\n  狀態：" + slotInfo.status + "\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + calendarLink + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(adminEmails.join(","), subject, body);   // ★ 一次寄送
  } catch(e) { Logger.log("發送管理者取消通知失敗：" + e.toString()); }
}

function sendVolunteerReplaceOutNotification(oldVolunteer, activityName, scheduleSlot, adminEmail, opTime) {
  if (!oldVolunteer.email) return;
  try {
    var parts = scheduleSlot.split(" ");
    var subject = "【志工排班系統】通知_" + oldVolunteer.id + "_" + oldVolunteer.name + "_代班異動_" + activityName + "_" + parts[0] + "_" + (parts[1]||"");
    var body = oldVolunteer.name + " 您好，\n\n您的以下排班已由管理人員調整為代班，請留意：\n\n  活動：" + activityName + "\n  時段：" + scheduleSlot + "\n  異動時間：" + opTime + "\n\n如有疑問，請直接聯繫管理人員。\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + getActivityCalendarLink(activityName) + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(oldVolunteer.email, subject, body);
  } catch(e) { Logger.log("發送換人（原）通知失敗：" + e.toString()); }
}

function sendVolunteerReplaceInNotification(newVolunteer, activityName, scheduleSlot, adminEmail, opTime) {
  if (!newVolunteer.email) return;
  try {
    var parts = scheduleSlot.split(" ");
    var subject = "【志工排班系統】通知_" + newVolunteer.id + "_" + newVolunteer.name + "_代班接班_" + activityName + "_" + parts[0] + "_" + (parts[1]||"");
    var body = newVolunteer.name + " 您好，\n\n管理人員已將您安排代班，詳情如下：\n\n  活動：" + activityName + "\n  時段：" + scheduleSlot + "\n  安排時間：" + opTime + "\n\n請確認您的行程，如有疑問請聯繫管理人員。\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + getActivityCalendarLink(activityName) + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(newVolunteer.email, subject, body);
  } catch(e) { Logger.log("發送換人（新）通知失敗：" + e.toString()); }
}

function sendAdminReplaceNotification(oldVolunteer, newVolunteer, activityName, scheduleSlot, adminEmail, opTime) {
  try {
    var adminEmails = getAdminEmails();
    if (adminEmails.length === 0) return;
    var slotInfo = getSlotInfo(activityName, scheduleSlot);
    var parts = scheduleSlot.split(" ");
    var subject = "【志工排班系統】通知_代班異動_" + activityName + "_" + parts[0] + "_" + (parts[1]||"");
    var body = "管理者您好，\n\n已執行代班異動，詳情如下：\n\n  活動：" + activityName + "\n  時段：" + scheduleSlot + "\n  原志工：" + oldVolunteer.name + "（" + oldVolunteer.id + "）\n  代班志工：" + newVolunteer.name + "（" + newVolunteer.id + "）\n  操作者：" + adminEmail + "\n  操作時間：" + opTime + "\n\n─── 時段目前狀態 ───\n  已報名：" + slotInfo.current + " / " + slotInfo.max + " 人\n  狀態：" + slotInfo.status + "\n\n─────────────────────────────\n📆 " + activityName + " 排班月曆：\n" + getActivityCalendarLink(activityName) + "\n\n此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(adminEmails.join(","), subject, body);   // ★ 一次寄送
  } catch(e) { Logger.log("發送管理者換人通知失敗：" + e.toString()); }
}

// ================================================
// 活動設定函數
// ================================================

function getActivityConfig(activityName) {
  var data = getSheetData(SHEET.ACTIVITY);
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].trim() === activityName.trim()) {
      return { id: data[i][0], name: data[i][1], type: data[i][2], calendarId: data[i][3], showCover: data[i][4], openReg: data[i][5] };
    }
  }
  return null;
}

function getActiveActivities() {
  var data = getSheetData(SHEET.ACTIVITY);
  var activities = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][4] === "是") activities.push({ name: data[i][1], type: data[i][2], calendarId: data[i][3], openReg: data[i][5] });
  }
  return activities;
}

function getIndexData() {
  var actData = getSheetData(SHEET.ACTIVITY);
  var activities = [];
  for (var i = 1; i < actData.length; i++) {
    if (actData[i][1] && actData[i][4] === "是") {
      var actName = actData[i][1], actType = actData[i][2], openReg = actData[i][5];
      var openStatus = { open: true };
      if (openReg === "是") {
        try { openStatus = getActivityOpenTime(actName); } catch(e) { openStatus = { open: true }; }
      }
      activities.push({ name: actName, type: actType, calendarId: actData[i][3], openReg: openReg, openStatus: openStatus });
    }
  }
  // ★ v1.3：排班規則加入到期日過濾（C欄）；留空 = 永久顯示
  var ruleData = getSheetData("排班規則");
  var rules = {};
  for (var j = 1; j < ruleData.length; j++) {
    if (!ruleData[j][0].trim()) continue;
    if (isRuleExpired(ruleData[j][2])) continue;   // 已過期不顯示於首頁
    rules[ruleData[j][0].trim()] = ruleData[j][1].trim();
  }
  return { activities: activities, rules: rules };
}

// ================================================
// ★ v1.3：排班規則到期日（純顯示用，不影響排班功能）
// ================================================

/**
 * 判斷排班規則是否已過期
 * 留空 = 永久顯示；到期日當天仍顯示，隔天起隱藏
 */
function isRuleExpired(rawExpiry) {
  var s = rawExpiry ? rawExpiry.toString().trim() : "";
  if (!s) return false;
  var d = parseExpiryDate(s);
  if (!d) return false;
  d.setHours(23, 59, 59, 999);
  return new Date() > d;
}

function getVolunteerBookedSlots(volunteerId, activityName) {
  var booked = [], activity = getActivityConfig(activityName);
  if (!activity) return booked;
  if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
    var data = getSheetData(SHEET.OVERVIEW);
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== activityName) continue;
      if (data[i][4] == volunteerId || data[i][6] == volunteerId ||
          data[i][8] == volunteerId || data[i][10] == volunteerId) {
        booked.push(data[i][1] + " " + data[i][2]);
      }
    }
  } else {
    var rosterData = getSheetData(SHEET.SESSION_ROSTER);
    for (var j = 1; j < rosterData.length; j++) {
      if (rosterData[j][2] == volunteerId) booked.push(rosterData[j][0] + " " + rosterData[j][1]);
    }
  }
  return booked;
}

// ================================================
// 公告設定儲存
// ================================================

function saveAnnounceDate(dateStr, activityName, timeStr, targetMonths, monthQuotas) {
  var dateParts = dateStr.split("-");
  var slashDate = dateParts[0] + "/" + parseInt(dateParts[1]) + "/" + parseInt(dateParts[2]);
  var openTimeStr = slashDate + " " + (timeStr || "00:00");

  // 月份一律由前端傳入（v1.4 起已移除「自動」模式）
  // 保留 fallback：萬一沒傳，才用公告日後第 2 個月
  var monthsToSave = [];
  if (targetMonths && targetMonths.length > 0) {
    monthsToSave = targetMonths;
  } else {
    var dp = dateStr.split("-");
    var baseY = parseInt(dp[0]), baseM = parseInt(dp[1]) - 1;
    var autoMonth = new Date(baseY, baseM + 2, 1);
    monthsToSave = [autoMonth.getFullYear() + "/" + (autoMonth.getMonth() + 1)];
  }

  var sheet = getSheet(SHEET.ANNOUNCE);
  for (var i = 0; i < monthsToSave.length; i++) {
    sheet.appendRow([openTimeStr, activityName, monthsToSave[i], ""]);
  }
  invalidateSheetData(SHEET.ANNOUNCE);

  if (monthQuotas && monthQuotas.length > 0) {
    saveMonthQuotas(activityName, monthQuotas);
  }

  // ★ v1.4 修正：改傳 monthsToSave（原本傳 targetMonths，null 時兩邊各自重算，
  //    可能因時區導致公告設定寫入 10 月、時段卻產生在 9 月）
  var slotsAdded = generateBinocularSlots(dateStr, activityName, monthsToSave);

  var monthLabel = monthsToSave.join("、");
  return "公告設定已儲存！開放月份：" + monthLabel +
         "，新增 " + slotsAdded + " 個時段。\n" +
         "志工自 " + openTimeStr + " 起可開始申請。";
}

// ================================================
// 月份時段查詢（前台用）
// ================================================

function logFormResponse(name, id, activityName, slot, operation, notes, result) {
  getSheet(SHEET.FORM).appendRow([new Date(), name, id, activityName, slot, operation, notes, result]);
  invalidateSheetData(SHEET.FORM);
}

function getMonthSlots(activityName, year, month) {
  var activity = getActivityConfig(activityName);
  if (!activity) return {};
  var result = {};

  if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
    var targetMonth = year + "/" + parseInt(month);
    var openMonths  = getOpenMonths(activityName);
    var isMonthOpen = false;
    for (var om = 0; om < openMonths.length; om++) {
      if (openMonths[om] === targetMonth) { isMonthOpen = true; break; }
    }
    if (!isMonthOpen) return {};

    var specialCache    = buildSpecialDateCache(activityName);
    var monthQuotaCache = buildMonthQuotaCache();
    var now             = new Date();

    var data = getSheetData(SHEET.OVERVIEW);
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== activityName) continue;
      var dateStr = data[i][1];
      if (!dateStr) continue;
      var dp = dateStr.split("/");
      if (dp.length < 3) continue;
      if (parseInt(dp[0]) !== year || parseInt(dp[1]) !== parseInt(month)) continue;
      var timeStr   = padTime(data[i][2].split("-")[0].trim());
      var slotStart = new Date(dateToISO(dateStr) + "T" + timeStr + ":00");
      if (isNaN(slotStart.getTime()) || slotStart <= now) continue;

      var timeSlot   = data[i][2];
      var currentVol = (data[i][3] ? 1 : 0) + (data[i][5] ? 1 : 0) +
                       (data[i][7] ? 1 : 0) + (data[i][9] ? 1 : 0);
      var maxVol = getBinocularMaxCached(dateStr, timeSlot, activityName, specialCache, monthQuotaCache);
      var status = currentVol >= maxVol ? "已滿" : currentVol > 0 ? "部分排班" : "可排班";

      if (!result[dateStr]) result[dateStr] = [];
      result[dateStr].push(timeSlot + ":" + status);
    }

  } else {
    var rosterData = getSheetData(SHEET.SESSION_ROSTER);
    var countMap = {};
    for (var r = 1; r < rosterData.length; r++) {
      var key = rosterData[r][0] + "|" + rosterData[r][1];
      countMap[key] = (countMap[key] || 0) + 1;
    }
    var sessionData = getSheetData(SHEET.SESSION);
    var now2 = new Date();
    for (var j = 1; j < sessionData.length; j++) {
      if (sessionData[j][0] !== activityName) continue;
      var dateStr2 = sessionData[j][1].toString();
      var dp2 = dateStr2.split("/");
      if (dp2.length < 3) continue;
      if (parseInt(dp2[0]) !== year || parseInt(dp2[1]) !== parseInt(month)) continue;
      var odtRaw = (sessionData[j][6] || "").toString().trim();
      if (!odtRaw) continue;
      var odt = parseSessionOpenTime(odtRaw);
      if (!odt || now2 < odt) continue;
      var startTime = padTime(sessionData[j][2].toString());
      var endTime   = padTime(sessionData[j][3].toString());
      var timeSlot2  = startTime + "-" + endTime;
      var totalMax  = (parseInt(sessionData[j][4]) || 0) + (parseInt(sessionData[j][5]) || 0);
      var cur       = countMap[dateStr2 + "|" + timeSlot2] || 0;
      var slotStart2 = new Date(dateToISO(dateStr2) + "T" + startTime + ":00");
      if (isNaN(slotStart2.getTime()) || slotStart2 <= now2) continue;
      var status2 = cur >= totalMax ? "已滿" : cur > 0 ? "部分排班" : "可排班";
      if (!result[dateStr2]) result[dateStr2] = [];
      result[dateStr2].push(timeSlot2 + ":" + status2);
    }
  }
  return result;
}

function getCalendarData(activityName) {
  var result = { openStatus: null, monthSlots: null, pendingAnnounces: [], targetYear: 0, targetMonth: 0 };
  var activity = getActivityConfig(activityName);
  if (!activity) { result.openStatus = { open: true, openMonths: [] }; return result; }

  if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
    var announceCache = buildAnnounceCache(activityName);
    var openMonths    = getOpenMonthsFromCache(announceCache);

    if (openMonths.length > 0) {
      result.openStatus = { open: true, openMonths: openMonths };
    } else {
      var nextTime = getNextOpenTimeFromCache(announceCache);
      result.openStatus = { open: false, openTime: nextTime ? formatOpenDateTime(nextTime) : "", openMonths: [] };
      result.pendingAnnounces = getPendingAnnouncesFromCache(announceCache);
      return result;
    }

    var now = new Date();
    var currentKey = now.getFullYear() + "/" + (now.getMonth() + 1);
    var sorted = openMonths.slice().sort();
    var nearestMonth = null;
    for (var i = 0; i < sorted.length; i++) {
      if (sorted[i] >= currentKey) { nearestMonth = sorted[i]; break; }
    }
    if (!nearestMonth) nearestMonth = sorted[sorted.length - 1];

    if (nearestMonth) {
      var parts = nearestMonth.split("/");
      result.targetYear  = parseInt(parts[0]);
      result.targetMonth = parseInt(parts[1]);
      result.monthSlots  = getMonthSlots(activityName, result.targetYear, result.targetMonth);
    }
    result.pendingAnnounces = getPendingAnnouncesFromCache(announceCache);

  } else {
    var openStatus = getActivityOpenTime(activityName);
    result.openStatus = openStatus;
    if (!openStatus.open) {
      if (openStatus.pendingSessionTimes && openStatus.pendingSessionTimes.length > 0) {
        result.pendingAnnounces = openStatus.pendingSessionTimes;
      }
      return result;
    }

    var now3 = new Date();
    var data3 = getSheetData(SHEET.SESSION);
    var currentKey3 = now3.getFullYear() + "/" + (now3.getMonth() + 1);
    var candidates = [];
    for (var si2 = 1; si2 < data3.length; si2++) {
      if (data3[si2][0] !== activityName) continue;
      var odtRaw4 = (data3[si2][6] || "").toString().trim();
      if (!odtRaw4) { candidates.push("immediate"); continue; }
      var odt4 = parseSessionOpenTime(odtRaw4);
      if (!odt4 || now3 < odt4) continue;
      var dp4 = data3[si2][1].toString().split("/");
      if (dp4.length < 3) continue;
      var mk4 = dp4[0] + "/" + parseInt(dp4[1]);
      if (candidates.indexOf(mk4) < 0) candidates.push(mk4);
    }
    candidates = candidates.filter(function(c) { return c !== "immediate"; }).sort();

    var nearestSessionMonth = null;
    for (var ci3 = 0; ci3 < candidates.length; ci3++) {
      if (candidates[ci3] >= currentKey3) { nearestSessionMonth = candidates[ci3]; break; }
    }
    if (!nearestSessionMonth && candidates.length > 0) nearestSessionMonth = candidates[candidates.length - 1];

    if (nearestSessionMonth) {
      var smParts = nearestSessionMonth.split("/");
      result.targetYear  = parseInt(smParts[0]);
      result.targetMonth = parseInt(smParts[1]);
      result.monthSlots  = getMonthSlots(activityName, result.targetYear, result.targetMonth);
    }

    if (openStatus.pendingSessionTimes && openStatus.pendingSessionTimes.length > 0) {
      result.pendingAnnounces = openStatus.pendingSessionTimes;
    }
  }

  return result;
}

function getVolunteerAndCalendarData(volunteerId, activityName) {
  var result = { volunteerName: null, openStatus: null, monthSlots: null, pendingAnnounces: [], targetYear: 0, targetMonth: 0 };
  result.volunteerName = getVolunteerNameById(volunteerId);
  if (!result.volunteerName) return result;
  var openStatus = getActivityOpenTime(activityName);
  result.openStatus = openStatus;
  if (!openStatus.open) {
    result.pendingAnnounces = getPendingAnnounces(activityName);
    return result;
  }
  var openMonths = openStatus.openMonths || [];
  var nearestMonth = null;
  if (openMonths.length > 0) {
    var now = new Date();
    var currentKey = now.getFullYear() + "/" + (now.getMonth() + 1);
    var sorted = openMonths.slice().sort();
    for (var i = 0; i < sorted.length; i++) {
      if (sorted[i] >= currentKey) { nearestMonth = sorted[i]; break; }
    }
    if (!nearestMonth) nearestMonth = sorted[sorted.length - 1];
  }
  if (nearestMonth) {
    var parts = nearestMonth.split("/");
    result.targetYear  = parseInt(parts[0]);
    result.targetMonth = parseInt(parts[1]);
    result.monthSlots = getMonthSlots(activityName, result.targetYear, result.targetMonth);
  }
  result.pendingAnnounces = getPendingAnnounces(activityName);
  return result;
}

// ================================================
// 活動開放時間查詢
// ================================================

function getActivityOpenTime(activityName) {
  var activityConfig = getActivityConfig(activityName);
  if (!activityConfig) return { open: true, openMonths: [] };
  if (activityConfig.type === ACTIVITY_TYPE.BINOCULAR) {
    var openMonths = getOpenMonths(activityName);
    if (openMonths.length > 0) return { open: true, openMonths: openMonths };
    var nextTime = getNextOpenTime(activityName);
    if (nextTime) return { open: false, openTime: formatOpenDateTime(nextTime), openMonths: [] };
    return { open: false, openTime: "", openMonths: [] };
  } else {
    var data = getSheetData(SHEET.SESSION);
    var now2 = new Date(), hasAnyOpen = false, earliestPending = null;
    var pendingTimes = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0] !== activityName) continue;
      var odtRaw = (data[i][6] || "").toString().trim();
      if (!odtRaw) { hasAnyOpen = true; continue; }
      var odt = parseSessionOpenTime(odtRaw);
      if (!odt) { hasAnyOpen = true; continue; }
      if (now2 >= odt) {
        hasAnyOpen = true;
      } else {
        var dateStr = data[i][1].toString().trim();
        var dp = dateStr.split("/");
        var monthKey = dp[0] + "/" + parseInt(dp[1]);
        var openStr = formatOpenDateTime(odt);
        var found = false;
        for (var p = 0; p < pendingTimes.length; p++) {
          if (pendingTimes[p].openTime === openStr) {
            if (pendingTimes[p].months.indexOf(monthKey) < 0) pendingTimes[p].months.push(monthKey);
            found = true; break;
          }
        }
        if (!found) pendingTimes.push({ openTime: openStr, months: [monthKey] });
        if (!earliestPending || odt < earliestPending) earliestPending = odt;
      }
    }
    if (hasAnyOpen) return { open: true, openMonths: [], pendingSessionTimes: pendingTimes };
    if (earliestPending) return { open: false, openTime: formatOpenDateTime(earliestPending), openMonths: [], pendingSessionTimes: pendingTimes };
    return { open: false, openTime: "", openMonths: [], pendingSessionTimes: [] };
  }
}

function getPendingAnnounces(activityName) {
  var data = getSheetData(SHEET.ANNOUNCE);
  var now  = new Date();
  var pending = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][1] !== activityName) continue;
    if (isAnnounceExpired(data[i][4])) continue;   // ★ v1.2：已到期不再提示
    var raw = data[i][0].toString().trim();
    var openTime = parseOpenDateTime(raw);
    if (!openTime) continue;
    if (now >= openTime) continue;
    var m1 = (data[i][2] || "").toString().trim();
    var m2 = (data[i][3] || "").toString().trim();
    var months = [m1, m2].filter(function(m) { return m; });
    if (months.length === 0) continue;
    var found = false;
    for (var p = 0; p < pending.length; p++) {
      if (pending[p].openTime === raw) {
        months.forEach(function(m) { if (pending[p].months.indexOf(m) < 0) pending[p].months.push(m); });
        found = true; break;
      }
    }
    if (!found) pending.push({ openTime: formatOpenDateTime(openTime), months: months });
  }
  pending.sort(function(a, b) { return parseOpenDateTime(a.openTime) - parseOpenDateTime(b.openTime); });
  return pending;
}

function getNextOpenTimeForDisplay(activityName) {
  var nextTime = getNextOpenTime(activityName);
  if (!nextTime) return null;
  return formatOpenDateTime(nextTime);
}

// ================================================
// 管理者後台函數
// ================================================

function getAdminActivities() {
  var data = getSheetData(SHEET.ACTIVITY);
  var activities = [];
  for (var i = 1; i < data.length; i++) {
    activities.push({ id: data[i][0], name: data[i][1], type: data[i][2], calendarId: data[i][3], showCover: data[i][4], openReg: data[i][5] });
  }
  return activities;
}

function saveAdminActivities(activities) {
  var sheet = getSheet(SHEET.ACTIVITY), lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  activities.forEach(function(act, idx) {
    sheet.getRange(idx + 2, 1, 1, 6).setValues([[act.id, act.name, act.type, act.calendarId, act.showCover, act.openReg]]);
  });
  invalidateSheetData(SHEET.ACTIVITY);
  return "活動設定已儲存！";
}

function getAdminSessions() {
  var data = getSheetData(SHEET.SESSION);
  var sessions = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var padT = function(t) { var p = t.toString().split(":"); return (p[0].length < 2 ? "0"+p[0] : p[0]) + ":" + p[1]; };
    sessions.push({ activity: data[i][0], date: data[i][1], startTime: padT(data[i][2]), endTime: padT(data[i][3]), generalMax: data[i][4], newMax: data[i][5], openDateTime: (data[i][6] || "").toString().trim() });
  }
  return sessions;
}

function saveAdminSession(activity, date, startTime, endTime, generalMax, newMax, openDateTime) {
  var dateParts = date.split("-");
  if (dateParts.length === 3) date = dateParts[0] + "/" + parseInt(dateParts[1]) + "/" + parseInt(dateParts[2]);
  startTime = padTime(startTime); endTime = padTime(endTime);
  var openDTStr = "";
  if (openDateTime && openDateTime.trim()) {
    var raw = openDateTime.trim();
    if (raw.indexOf("T") >= 0) {
      var dp = raw.split("T"), datep = dp[0].split("-");
      openDTStr = datep[0] + "/" + parseInt(datep[1]) + "/" + parseInt(datep[2]) + " " + dp[1];
    } else { openDTStr = raw; }
  }
  var overviewSheet = getSheet(SHEET.OVERVIEW), sessionSheet = getSheet(SHEET.SESSION);
  sessionSheet.appendRow([activity, date, startTime, endTime, generalMax, newMax, openDTStr]);
  var timeSlot = startTime + "-" + endTime;
  var existing = overviewSheet.getDataRange().getDisplayValues();
  var found = false;
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][0] === activity && existing[i][1] === date && existing[i][2] === timeSlot) { found = true; break; }
  }
  if (!found) overviewSheet.appendRow([activity, date, timeSlot, "", "", "", "", "", "", "", "", "可排班"]);
  invalidateSheetData();
  return "場次已新增！";
}

function deleteAdminSession(activity, date, startTime, endTime) {
  function padT(t) { var p = t.toString().split(":"); return (p[0].length < 2 ? "0"+p[0] : p[0]) + ":" + p[1]; }
  startTime = padT(startTime); endTime = padT(endTime);
  var timeSlot = startTime + "-" + endTime;
  var sessionSheet = getSheet(SHEET.SESSION), sessionData = sessionSheet.getDataRange().getDisplayValues();
  for (var i = sessionData.length - 1; i >= 1; i--) {
    if (sessionData[i][0] === activity && sessionData[i][1] === date && padT(sessionData[i][2]) === startTime) { sessionSheet.deleteRow(i + 1); break; }
  }
  var overviewSheet = getSheet(SHEET.OVERVIEW), overviewData = overviewSheet.getDataRange().getDisplayValues();
  for (var j = overviewData.length - 1; j >= 1; j--) {
    if (overviewData[j][0] === activity && overviewData[j][1] === date && overviewData[j][2] === timeSlot) { overviewSheet.deleteRow(j + 1); break; }
  }
  invalidateSheetData();
  return "場次已刪除！";
}

function updateAdminSessionMax(activity, date, startTime, endTime, newGeneralMax, newNewMax, openDateTime) {
  function padT(t) { var p = t.toString().split(":"); return (p[0].length < 2 ? "0"+p[0] : p[0]) + ":" + p[1]; }
  startTime = padT(startTime); endTime = padT(endTime);
  var timeSlot = startTime + "-" + endTime;
  var generalCount = countSessionBookings(date, timeSlot, activity, VOLUNTEER_TYPE.GENERAL);
  var newCount     = countSessionBookings(date, timeSlot, activity, VOLUNTEER_TYPE.NEW);
  var gMax = parseInt(newGeneralMax) || 0, nMax = parseInt(newNewMax) || 0;
  if (gMax < generalCount) return "儲存失敗：一般志工上限（" + gMax + "）不可低於已報名人數（" + generalCount + "）。";
  if (nMax < newCount)     return "儲存失敗：新志工上限（" + nMax + "）不可低於已報名人數（" + newCount + "）。";
  var openDTStr = "";
  if (openDateTime && openDateTime.trim()) {
    var raw = openDateTime.trim();
    if (raw.indexOf("T") >= 0) { var dp = raw.split("T"), datep = dp[0].split("-"); openDTStr = datep[0] + "/" + parseInt(datep[1]) + "/" + parseInt(datep[2]) + " " + dp[1]; }
    else { openDTStr = raw; }
  }
  var sessionSheet = getSheet(SHEET.SESSION), sessionData = sessionSheet.getDataRange().getDisplayValues();
  var updated = false;
  for (var i = 1; i < sessionData.length; i++) {
    if (sessionData[i][0] === activity && sessionData[i][1] === date && padT(sessionData[i][2]) === startTime) {
      sessionSheet.getRange(i+1, 5).setValue(gMax); sessionSheet.getRange(i+1, 6).setValue(nMax); sessionSheet.getRange(i+1, 7).setValue(openDTStr);
      updated = true; break;
    }
  }
  if (!updated) return "儲存失敗：找不到對應場次。";
  invalidateSheetData(SHEET.SESSION);
  updateSessionOverviewStatus(activity, date, timeSlot, { generalMax: gMax, newMax: nMax });
  return "場次已更新！";
}

function verifyAdminPassword(email, password) {
  var cache = CacheService.getScriptCache(), cacheKey = "adminList", admins;
  var cached = cache.get(cacheKey);
  if (cached) { try { admins = JSON.parse(cached); } catch(e) { admins = null; } }
  if (!admins) {
    var data = getSheetData(SHEET.SYSTEM);
    admins = [];
    for (var i = 1; i <= ADMIN_MAX; i++) admins.push({ email: getSystemConfigFromData(data, "管理者"+i+"_Email"), password: getSystemConfigFromData(data, "管理者"+i+"_密碼") });
    cache.put(cacheKey, JSON.stringify(admins), 600);
  }
  for (var i = 0; i < admins.length; i++) {
    if (admins[i].email && admins[i].password && email.trim() === admins[i].email.trim() && password === admins[i].password) return true;
  }
  return false;
}

function getSystemConfigFromData(data, key) {
  for (var i = 1; i < data.length; i++) { if (data[i][0] === key) return data[i][1]; }
  return null;
}

var ADMIN_MAX = 10;

function getAdminEmails() {
  var emails = [];
  for (var i = 1; i <= ADMIN_MAX; i++) { var email = getSystemConfig("管理者"+i+"_Email"); if (email && email.trim()) emails.push(email.trim()); }
  return emails;
}

function getAdminList() {
  var admins = [];
  for (var i = 1; i <= ADMIN_MAX; i++) admins.push({ email: getSystemConfig("管理者"+i+"_Email") || '', password: getSystemConfig("管理者"+i+"_密碼") || '' });
  return admins;
}

function saveAdminList(admins) {
  var sheet = getSheet(SHEET.SYSTEM), data = sheet.getDataRange().getDisplayValues();
  for (var row = 1; row < data.length; row++) { if (data[row][0] && data[row][0].indexOf("管理者") === 0) sheet.getRange(row+1, 1, 1, 2).clearContent(); }
  for (var i = 0; i < admins.length; i++) {
    var num = i + 1;
    if (admins[i].email || admins[i].password) { sheet.appendRow(["管理者"+num+"_Email", admins[i].email]); sheet.appendRow(["管理者"+num+"_密碼", admins[i].password]); }
  }
  CacheService.getScriptCache().remove("adminList");
  invalidateSheetData(SHEET.SYSTEM);
  return "管理者帳號已儲存！";
}

/**
 * 後台用：回傳每個活動的規則內容 + 到期日
 * 格式：{ 活動名稱: { content, expiry, expiryForInput, expired } }
 */
function getActivityRules() {
  var data = getSheetData("排班規則"), rules = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0].trim()) continue;
    var expiryRaw = (data[i][2] || "").toString().trim();
    rules[data[i][0].trim()] = {
      content:        data[i][1].trim(),
      expiry:         expiryRaw,
      expiryForInput: expirySlashToDash(expiryRaw),
      expired:        isRuleExpired(expiryRaw)
    };
  }
  return rules;
}

function saveActivityRule(activityName, ruleContent, expiryDate) {
  var sheet = getSheet("排班規則"), data = sheet.getDataRange().getDisplayValues();
  var expiryStr = normalizeExpiryToSlash(expiryDate);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].trim() === activityName.trim()) {
      sheet.getRange(i+1, 2).setValue(ruleContent);
      sheet.getRange(i+1, 3).setValue(expiryStr);
      invalidateSheetData("排班規則");
      return "規則已儲存！";
    }
  }
  sheet.appendRow([activityName, ruleContent, expiryStr]);
  invalidateSheetData("排班規則");
  return "規則已儲存！";
}

function saveAdminEmail(email) {
  var sheet = getSheet(SHEET.SYSTEM), data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "管理者Email") {
      sheet.getRange(i+1, 2).setValue(email);
      invalidateSheetData(SHEET.SYSTEM);
      return "管理者 Email 已儲存！";
    }
  }
  sheet.appendRow(["管理者Email", email]);
  invalidateSheetData(SHEET.SYSTEM);
  return "管理者 Email 已儲存！";
}

function parseExpiryDate(str) {
  if (!str || !str.trim()) return null;
  var d = new Date(str.trim().replace(/\//g, "-"));
  return isNaN(d.getTime()) ? null : d;
}

function normalizeExpiryToSlash(str) {
  if (!str || !str.trim()) return "";
  var d = parseExpiryDate(str);
  if (!d) return str.trim();
  return d.getFullYear() + "/" + (d.getMonth()+1) + "/" + d.getDate();
}

function expirySlashToDash(str) {
  if (!str || !str.trim()) return "";
  return str.trim().replace(/\//g, "-");
}

function getRecentActivities() {
  var sheet = getSpreadsheet().getSheetByName(SHEET.RECENT);
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues(), items = [];
  for (var i = 1; i < data.length; i++) {
    var title = (data[i][0] || "").trim(), url = (data[i][1] || "").trim(), expiry = (data[i][2] || "").trim();
    if (title && url) items.push({ title: title, url: url, expiry: expiry, expiryForInput: expirySlashToDash(expiry) });
  }
  return items;
}

function getRecentActivitiesForIndex() {
  var all = getRecentActivities(), today = new Date();
  today.setHours(0, 0, 0, 0);
  return all.filter(function(item) {
    if (!item.expiry) return true;
    var exp = parseExpiryDate(item.expiry);
    if (!exp) return true;
    exp.setHours(0, 0, 0, 0);
    return exp >= today;
  });
}

function saveRecentActivities(items) {
  var ss = getSpreadsheet(), sheet = ss.getSheetByName(SHEET.RECENT);
  if (!sheet) { sheet = ss.insertSheet(SHEET.RECENT); sheet.getRange(1, 1, 1, 3).setValues([["標題", "連結", "到期日"]]); }
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow-1, 3).clearContent();
  var writeRow = 2;
  for (var n = 0; n < items.length && n < 10; n++) {
    var item = items[n];
    if (!item.title.trim() && !item.url.trim()) continue;
    sheet.getRange(writeRow, 1, 1, 3).setValues([[item.title.trim(), item.url.trim(), normalizeExpiryToSlash(item.expiry)]]);
    writeRow++;
  }
  invalidateSheetData(SHEET.RECENT);
  return "最新活動已儲存！";
}

function getIndexDataFull() {
  var base = getIndexData();
  base.recentActivities = getRecentActivitiesForIndex();
  base.vacancySummary = getVacancySummary(30);
  return base;
}

function getAdminAnnounces() {
  var data = getSheetData(SHEET.ANNOUNCE);
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      rowIndex: i+1,
      openTime: data[i][0].toString().trim(),
      activity: data[i][1].toString().trim(),
      month1:   data[i][2].toString().trim(),
      month2:   data[i][3].toString().trim()
    });
  }
  return list;
}

function deleteAdminAnnounce(rowIndex) {
  var sheet = getSheet(SHEET.ANNOUNCE);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return "找不到對應公告。";
  sheet.deleteRow(rowIndex);
  invalidateSheetData(SHEET.ANNOUNCE);
  return "公告已刪除。";
}

/**
 * ★ v1.2：更新單筆公告的到期日（E欄）
 * 留空 = 永久顯示。到期只影響提示顯示，不影響排班
 */

function getAdminVolunteers() {
  var data = getSheetData(SHEET.ROSTER);
  var volunteers = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    volunteers.push({ rowIndex: i+1, id: data[i][0].toString().trim(), name: data[i][1].toString().trim(), type: data[i][2].toString().trim(), teleCert: data[i][3].toString().trim(), venueCert: data[i][4].toString().trim(), email: (data[i][5]||"").toString().trim() });
  }
  return volunteers;
}

function addVolunteerToRoster(id, name, type, teleCert, venueCert, email) {
  var sheet = getSheet(SHEET.ROSTER), data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === id.toString().trim()) return "新增失敗：志工編號 " + id + " 已存在。";
  }
  sheet.appendRow([id.toString().trim(), name.toString().trim(), type.toString().trim(), teleCert.toString().trim(), venueCert.toString().trim(), (email||"").toString().trim()]);
  invalidateSheetData(SHEET.ROSTER);
  return "志工「" + name + "」已成功新增！";
}

function updateAdminVolunteer(rowIndex, id, name, type, teleCert, venueCert, email) {
  var sheet = getSheet(SHEET.ROSTER), data = sheet.getDataRange().getDisplayValues();
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return "更新失敗：列號超出範圍。";
  for (var i = 1; i < data.length; i++) {
    if (i+1 === rowIndex) continue;
    if (data[i][0].toString().trim() === id.toString().trim()) return "更新失敗：志工編號 " + id + " 已被其他志工使用。";
  }
  sheet.getRange(rowIndex, 1, 1, 6).setValues([[id.toString().trim(), name.toString().trim(), type.toString().trim(), teleCert.toString().trim(), venueCert.toString().trim(), (email||"").toString().trim()]]);
  invalidateSheetData(SHEET.ROSTER);
  return "志工「" + name + "」資料已更新！";
}

function deleteAdminVolunteer(rowIndex, volunteerId) {
  var sheet = getSheet(SHEET.ROSTER);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return "刪除失敗：列號超出範圍。";
  var volunteerName = getVolunteerNameById(volunteerId);
  sheet.deleteRow(rowIndex);
  invalidateSheetData(SHEET.ROSTER);
  return "志工「" + (volunteerName||volunteerId) + "」已刪除。";
}

function exportVolunteersCSV() {
  var data = getSheetData(SHEET.ROSTER);
  var rows = [["志工編號","姓名","志工類型","望遠鏡驗收","展場驗收","Email"]];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    rows.push([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5]||""]);
  }
  var csv = rows.map(function(row) {
    return row.map(function(cell) {
      var s = cell.toString();
      if (s.indexOf(",")>=0||s.indexOf('"')>=0||s.indexOf('\n')>=0) s='"'+s.replace(/"/g,'""')+'"';
      return s;
    }).join(",");
  }).join("\n");
  return "\uFEFF" + csv;
}

// ================================================
// ★ v1.2：系統參數設定（管理者後台可調）
// ================================================

function getSystemParams() {
  return {
    priorityHours:      getSystemConfig("優先期時數")   || "48",
    cancelDeadlineDays: getSystemConfig("取消截止天數") || "7",
    selfCancelQuota:    getSystemConfig("自助取消月額度") || "2"
  };
}

/**
 * ★ v1.4：「二觀優先排班人數設定」頁籤一次取得所有資料
 * 原本前端要發 3 次 google.script.run（配額清單、活動清單、參數），
 * 每次都有 1-2 秒固定延遲，排隊起來就是 3-6 秒。
 * 合併成 1 次後端呼叫，速度大幅改善。
 */
function getQuotaPageData() {
  var activities = [];
  var actData = getSheetData(SHEET.ACTIVITY);
  for (var i = 1; i < actData.length; i++) {
    if (actData[i][1] && actData[i][2] === ACTIVITY_TYPE.BINOCULAR) {
      activities.push(actData[i][1].toString().trim());
    }
  }
  return {
    quotas:     getAdminMonthQuotas(),
    activities: activities,
    params:     getSystemParams()
  };
}

function saveSystemParams(params) {
  var sheet = getSheet(SHEET.SYSTEM);
  var data  = sheet.getDataRange().getDisplayValues();

  function setParam(key, value) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
    }
    sheet.appendRow([key, value]);
    data.push([key, value]);
  }

  var ph = parseInt(params.priorityHours);
  var cd = parseInt(params.cancelDeadlineDays);
  var sq = parseInt(params.selfCancelQuota);

  setParam("優先期時數",     isNaN(ph) || ph < 0 ? 48 : ph);
  setParam("取消截止天數",   isNaN(cd) || cd < 0 ? 7  : cd);
  setParam("自助取消月額度", isNaN(sq) || sq < 0 ? 2  : sq);

  invalidateSheetData(SHEET.SYSTEM);
  return "排班參數已儲存！";
}

function clearAdminCache() {
  CacheService.getScriptCache().remove("adminList");
  invalidateSheetData();
  Logger.log("快取已清除");
}

function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    var events = json.events;
    if (events && events.length > 0) {
      var source = events[0].source;
      if (source && source.groupId) {
        getSheet(SHEET.SYSTEM).appendRow(["LINE_GROUP_ID", source.groupId]);
        invalidateSheetData(SHEET.SYSTEM);
      }
    }
  } catch(err) {
    Logger.log("doPost 錯誤：" + err.toString());
  }
}

// ================================================
// ★★★ v1.5：自助取消通知 ★★★
// ================================================

/**
 * 通知志工：您已自行取消排班
 */
function sendVolunteerSelfCancelNotification(volunteer, activityName, scheduleSlot, cancelTime, cancelInfo) {
  if (!volunteer.email) return;
  try {
    var parts = scheduleSlot.split(" ");
    var subject = "【志工排班系統】通知_" + volunteer.id + "_" + volunteer.name +
                  "_自助取消_" + activityName + "_" + parts[0] + "_" + (parts[1] || "");
    var body = volunteer.name + " 您好，\n\n" +
      "您已完成以下排班的取消，該時段已重新開放給其他志工登記：\n\n" +
      "  活動：" + activityName + "\n" +
      "  時段：" + scheduleSlot + "\n" +
      "  取消時間：" + cancelTime + "\n\n" +
      "─── 自助取消額度 ───\n" +
      "  本月已使用：" + cancelInfo.used + " / " + cancelInfo.max + " 次\n" +
      "  尚可取消：" + cancelInfo.remaining + " 次\n" +
      "  額度重置日：" + cancelInfo.resetDate + "\n\n" +
      "如需再次登記，請至排班系統重新申請。\n\n" +
      "─────────────────────────────\n" +
      "📆 " + activityName + " 排班月曆：\n" + getActivityCalendarLink(activityName) + "\n\n" +
      "此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(volunteer.email, subject, body);
  } catch(e) { Logger.log("發送志工自助取消通知失敗：" + e.toString()); }
}

/**
 * 通知管理者：有志工自助取消（一次寄送給全部管理者）
 */
function sendAdminSelfCancelNotification(volunteer, activityName, scheduleSlot, cancelTime) {
  try {
    var adminEmails = getAdminEmails();
    if (adminEmails.length === 0) return;
    var parts = scheduleSlot.split(" ");
    var slotInfo = getSlotInfo(activityName, scheduleSlot);
    var subject = "【志工排班系統】通知_" + volunteer.id + "_" + volunteer.name +
                  "_自助取消_" + activityName + "_" + parts[0] + "_" + (parts[1] || "");
    var body = "管理者您好，\n\n" +
      "有志工自行取消排班，該時段已釋出，詳情如下：\n\n" +
      "  志工姓名：" + volunteer.name + "\n" +
      "  志工編號：" + volunteer.id + "\n" +
      "  活動：" + activityName + "\n" +
      "  時段：" + scheduleSlot + "\n" +
      "  取消時間：" + cancelTime + "\n\n" +
      "─── 時段目前狀態（取消後）───\n" +
      "  已報名：" + slotInfo.current + " / " + slotInfo.max + " 人\n" +
      "  尚缺：" + Math.max(0, slotInfo.max - slotInfo.current) + " 人\n" +
      "  狀態：" + slotInfo.status + "\n\n" +
      "─────────────────────────────\n" +
      "📆 " + activityName + " 排班月曆：\n" + getActivityCalendarLink(activityName) + "\n\n" +
      "此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(adminEmails.join(","), subject, body);
  } catch(e) { Logger.log("發送管理者自助取消通知失敗：" + e.toString()); }
}

function migrateSessionRoster() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. 找到舊工作表
    var oldSheet = ss.getSheetByName("親子營排班");
    if (!oldSheet) {
      Browser.msgBox("找不到「親子營排班」工作表，請確認名稱。");
      return;
    }

    // 2. 改名為「場次排班」
    oldSheet.setName("場次排班");

    // 3. 在 A 欄前插入新欄
    oldSheet.insertColumnBefore(1);

    // 4. 填入標頭
    oldSheet.getRange(1, 1).setValue("活動名稱");

    // 5. 把所有資料列的 A 欄填入活動名稱
    //    ★ 把下面這個字串改成你「活動設定」B 欄裡的正式名稱（完全一樣）
    var activityName = "親子營";  // ← 改這裡

    var lastRow = oldSheet.getLastRow();
    if (lastRow > 1) {
      var range = oldSheet.getRange(2, 1, lastRow - 1, 1);
      var values = Array(lastRow - 1).fill([activityName]);
      range.setValues(values);
    }

    Browser.msgBox("遷移完成！共更新 " + (lastRow - 1) + " 筆資料。");
  }
