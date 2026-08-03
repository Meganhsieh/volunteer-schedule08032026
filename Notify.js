// ================================================
// Email 通知函數
// ================================================
function getLineToken() {
  return PropertiesService.getScriptProperties().getProperty("LINE_TOKEN") || "";
}

function getLineGroupId() {
  return PropertiesService.getScriptProperties().getProperty("LINE_GROUP_ID") || "";
}
/**
 * 每日定時執行：掃描未來一個月未排滿時段，通知管理者
 * sendLine: true = 同時推 LINE，false = 只發 Email
 */
function dailyNotifyCheck(sendLine) {

  var adminEmails = getAdminEmails();
  if (adminEmails.length === 0) {
    Logger.log("未設定管理者Email，略過通知。");
    return;
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var oneMonthLater = new Date(today);
  oneMonthLater.setMonth(oneMonthLater.getMonth() + 1);

  var overviewData  = getSheet(SHEET.OVERVIEW).getDataRange().getDisplayValues();
  var sessionData   = getSheet(SHEET.SESSION).getDataRange().getDisplayValues();
  var sessionRoster = getSheet(SHEET.SESSION_ROSTER).getDataRange().getDisplayValues();
  var actData       = getSheet(SHEET.ACTIVITY).getDataRange().getDisplayValues();

  var sessionCountMap = {};
  for (var r = 1; r < sessionRoster.length; r++) {
    var rKey = sessionRoster[r][0] + "|" + sessionRoster[r][1];
    sessionCountMap[rKey] = (sessionCountMap[rKey] || 0) + 1;
  }

  function getBinocularActivityName() {
    for (var a = 1; a < actData.length; a++) {
      if (actData[a][2] === ACTIVITY_TYPE.BINOCULAR) return actData[a][1];
    }
    return "望遠鏡二觀";
  }

  function getBinocularSlotInfo(actName, dateStr, timeSlot) {
    for (var i = 1; i < overviewData.length; i++) {
      if (overviewData[i][0] === actName &&
          overviewData[i][1] === dateStr &&
          overviewData[i][2] === timeSlot) {
        var current = (overviewData[i][3] ? 1 : 0) + (overviewData[i][5] ? 1 : 0);
        var status  = overviewData[i][7] || "可排班";
        return { current: current, status: status };
      }
    }
    return { current: 0, status: "可排班" };
  }

  function getSessionCurrentCount(dateStr, timeSlot) {
    return sessionCountMap[dateStr + "|" + timeSlot] || 0;
  }

  function generateBinocularSlotsForNotify(actName, fromDate, toDate) {
    var slots = [];
    var current = new Date(fromDate);
    while (current <= toDate) {
      var day = current.getDay();
      if (day !== 1) {
        var year    = current.getFullYear();
        var month   = current.getMonth() + 1;
        var date    = current.getDate();
        var dateStr = year + "/" + month + "/" + date;

        var slotDefs = (day === 6) ? BINOCULAR_SLOTS.SATURDAY
                     : (day === 0) ? BINOCULAR_SLOTS.SUNDAY
                     : BINOCULAR_SLOTS.WEEKDAY;

        slotDefs.forEach(function(slot) {
          var timeSlot = slot.start + "-" + slot.end;
          var max      = getBinocularMax(dateStr);
          var info     = getBinocularSlotInfo(actName, dateStr, timeSlot);
          var status   = info.current >= max ? "已滿" : info.status;
          slots.push({
            activity: actName,
            date:     dateStr,
            slot:     timeSlot,
            current:  info.current,
            max:      max,
            status:   status,
            shortage: max - info.current
          });
        });
      }
      current.setDate(current.getDate() + 1);
    }
    return slots;
  }

  var binocularActName = getBinocularActivityName();
  var allSlots = [];

  var binocularFull = generateBinocularSlotsForNotify(binocularActName, today, oneMonthLater);
  binocularFull.forEach(function(s) {
    if (s.status !== "已滿") allSlots.push(s);
  });

  for (var s3 = 1; s3 < sessionData.length; s3++) {
    if (!sessionData[s3][1]) continue;
    var slotD = new Date(sessionData[s3][1]);
    slotD.setHours(0, 0, 0, 0);
    if (slotD < today || slotD > oneMonthLater) continue;
    var actN   = sessionData[s3][0];
    var dateS  = sessionData[s3][1];
    var stTime = padTime(sessionData[s3][2]);
    var enTime = padTime(sessionData[s3][3]);
    var tSlot  = stTime + "-" + enTime;
    var maxV   = (parseInt(sessionData[s3][4]) || 0) + (parseInt(sessionData[s3][5]) || 0);
    var curV   = getSessionCurrentCount(dateS, tSlot);
    if (curV < maxV) {
      allSlots.push({
        activity: actN,
        date:     dateS,
        slot:     tSlot,
        current:  curV,
        max:      maxV,
        status:   curV > 0 ? "部分排班" : "可排班",
        shortage: maxV - curV
      });
    }
  }

  // 全部排滿
  if (allSlots.length === 0) {
    var fullSubject = "【志工排班系統】排班通知 " + formatDate(today);
    var fullBody = "您好，\n\n" +
      "🎉 未來一個月內所有時段均已排滿，目前無缺額！\n\n" +
      "─────────────────────────────\n" +
      "📆 排班月曆\n\n";
    for (var fc = 1; fc < actData.length; fc++) {
      var fcName = actData[fc][1];
      var fcId   = actData[fc][3];
      if (!fcName || !fcId) continue;
      fullBody += fcName + "：\n";
      fullBody += "https://calendar.google.com/calendar/embed?src=" +
        encodeURIComponent(fcId) + "&ctz=Asia%2FTaipei\n\n";
    }
    fullBody += "─────────────────────────────\n";
    fullBody += "⚙️ 管理者後台：\n";
    fullBody += ScriptApp.getService().getUrl() + "?page=admin\n\n";
    fullBody += "此信件由系統自動發送，請勿回覆。";
    MailApp.sendEmail(adminEmails.join(","), fullSubject, fullBody);
    if (sendLine) sendLineMessage("🎉 " + formatDate(today) + " 未來一個月所有時段均已排滿，目前無缺額！");
    Logger.log("所有時段已滿，已發送滿班通知。");
    return;
  }

  // 有缺額
  var subject = "【志工排班系統】排班缺額通知 " + formatDate(today);
  var body = "您好，\n\n以下為未來一個月內尚未排滿的時段，請協助安排：\n\n";

  allSlots.sort(function(a, b) {
    function parseDate(str) {
      var parts = str.split("/");
      return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    }
    var dateA = parseDate(a.date);
    var dateB = parseDate(b.date);
    if (dateA - dateB !== 0) return dateA - dateB;
    var timeA = a.slot.split("-")[0];
    var timeB = b.slot.split("-")[0];
    return timeA.localeCompare(timeB);
  });

  var allGrouped = {};
  allSlots.forEach(function(s) {
    if (!allGrouped[s.activity]) allGrouped[s.activity] = [];
    allGrouped[s.activity].push(s);
  });

  Object.keys(allGrouped).forEach(function(actName) {
    body += "【" + actName + "】\n";
    allGrouped[actName].forEach(function(s) {
      body += "  ✗ " + s.date + " " + s.slot +
              "　已排 " + s.current + "/" + s.max + " 人　缺 " + s.shortage + " 人" +
              "　（" + s.status + "）\n";
    });
    body += "\n";
  });

  body += "─────────────────────────────\n";
  body += "請登入排班系統進行人工安排。\n\n";
  body += "─────────────────────────────\n";
  body += "📆 排班月曆\n\n";
  var calLinksAdded = false;
  for (var ca = 1; ca < actData.length; ca++) {
    var caName = actData[ca][1];
    var caId   = actData[ca][3];
    if (!caName || !caId) continue;
    body += caName + "：\n";
    body += "https://calendar.google.com/calendar/embed?src=" +
      encodeURIComponent(caId) + "&ctz=Asia%2FTaipei\n\n";
    calLinksAdded = true;
  }
  if (!calLinksAdded) body += "（未設定活動月曆）\n\n";

  body += "─────────────────────────────\n";
  body += "⚙️ 管理者後台：\n";
  body += ScriptApp.getService().getUrl() + "?page=admin\n\n";
  body += "此信件由系統自動發送，請勿回覆。";

  var recipient = adminEmails.join(",");
  MailApp.sendEmail(recipient, subject, body);
  if (sendLine) sendLineMessage(buildLineMessage(allGrouped));
  Logger.log("通知Email已發送至：" + recipient);
}

/**
 * 每天執行：只發 Email，不推 LINE
 */
function dailyEmailOnly() {
  dailyNotifyCheck(false);
}

/**
 * 每週一執行：發 Email + 推 LINE
 */
function weeklyNotifyCheck() {
  var today = new Date();
  if (today.getDay() !== 1) return;
  dailyNotifyCheck(true);
}

/**
 * 設定每日定時觸發器
 */
function setupDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "dailyNotifyCheck") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("dailyEmailOnly")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  Logger.log("每日觸發器設定完成，每天早上8點執行。");
}

/**
 * 格式化日期為 YYYY/MM/DD
 */
function formatDate(date) {
  var y = date.getFullYear();
  var m = date.getMonth() + 1;
  var d = date.getDate();
  return y + "/" + (m < 10 ? "0" + m : m) + "/" + (d < 10 ? "0" + d : d);
}

function buildLineMessage(allGrouped) {
  var today = new Date();
  var dateStr = today.getFullYear() + "/" + (today.getMonth()+1) + "/" + today.getDate();
  var msg = "【志工排班缺額通知】" + dateStr + "\n\n";
  Object.keys(allGrouped).forEach(function(actName) {
    msg += "▌" + actName + "\n";
    allGrouped[actName].forEach(function(s) {
      msg += s.date + " " + s.slot + "\n";
      msg += "已排 " + s.current + "/" + s.max + " 人，缺 " + s.shortage + " 人\n";
    });
    msg += "\n";
  });
  msg += "請至排班系統安排志工。";
  return msg;
}

function sendLineMessage(message) {
  try {
    var token   = getLineToken();
    var groupId = getLineGroupId();
    if (!token || !groupId) {
      Logger.log("LINE 設定未完成，略過推播。");
      return;
    }
    var options = {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + token },
      payload: JSON.stringify({
        to: groupId,
        messages: [{ type: "text", text: message }]
      }),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
    Logger.log("LINE 推送結果：" + response.getContentText());
  } catch(e) {
    Logger.log("LINE 推送失敗：" + e.toString());
  }
}