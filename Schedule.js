// ================================================
// 排班處理函數
// 版本：v1.4
// v1.1：讀表快取（getSheetData）
// v1.2：優先期時數改為後台可設定；新增取消相關參數
// v1.4：移除公告到期日（改由「排班規則」表管理）
// ================================================

function padTime(t) {
  var parts = t.toString().split(":");
  return (parts[0].length < 2 ? "0" + parts[0] : parts[0]) + ":" + parts[1];
}

function dateToISO(dateStr) {
  var p = dateStr.toString().split("/");
  if (p.length !== 3) return dateStr;
  return p[0] + "-" +
    (p[1].length < 2 ? "0" + p[1] : p[1]) + "-" +
    (p[2].length < 2 ? "0" + p[2] : p[2]);
}

function parseOpenDateTime(raw) {
  if (!raw || !raw.trim()) return null;
  var s = raw.trim();
  var d;
  if (s.indexOf("T") >= 0) {
    d = new Date(s);
  } else if (s.indexOf(" ") >= 0) {
    var pp = s.split(" ");
    var dp = pp[0].split("/");
    var y  = dp[0];
    var m  = (dp[1].length < 2 ? "0" + dp[1] : dp[1]);
    var dd = (dp[2].length < 2 ? "0" + dp[2] : dp[2]);
    var tp = pp[1].split(":");
    var hh = (tp[0].length < 2 ? "0" + tp[0] : tp[0]);
    var mi = (tp[1].length < 2 ? "0" + tp[1] : tp[1]);
    d = new Date(y + "-" + m + "-" + dd + "T" + hh + ":" + mi + ":00");
  } else {
    d = new Date(s.replace(/\//g, "-"));
  }
  return isNaN(d.getTime()) ? null : d;
}

function formatOpenDateTime(d) {
  return d.getFullYear() + "/" + (d.getMonth()+1) + "/" + d.getDate() + " " +
    (d.getHours() < 10 ? "0" + d.getHours() : d.getHours()) + ":" +
    (d.getMinutes() < 10 ? "0" + d.getMinutes() : d.getMinutes());
}

function parseSessionOpenTime(raw) {
  if (!raw || !raw.trim()) return null;
  raw = raw.trim();
  var d;
  if (raw.indexOf("T") >= 0) {
    d = new Date(raw);
  } else if (raw.indexOf(" ") >= 0) {
    var pp = raw.split(" ");
    var dp = pp[0].split("/");
    var y  = dp[0];
    var mo = (dp[1].length < 2 ? "0" + dp[1] : dp[1]);
    var dd = (dp[2].length < 2 ? "0" + dp[2] : dp[2]);
    var tp = pp[1].split(":");
    var hh = (tp[0].length < 2 ? "0" + tp[0] : tp[0]);
    var mi = (tp[1].length < 2 ? "0" + tp[1] : tp[1]);
    d = new Date(y + "-" + mo + "-" + dd + "T" + hh + ":" + mi + ":00");
  } else {
    d = new Date(raw.replace(/\//g, "-"));
  }
  return isNaN(d.getTime()) ? null : d;
}

// ================================================
// ★ v1.2 新增：後台可調參數
// 全部存在「系統設定」工作表，管理者後台可修改
// ================================================

/**
 * 優先排班期時數（公告開放後多久內套用優先配額）
 * 僅適用二觀型。未設定時預設 48 小時
 */
function getPriorityHours() {
  var v = parseInt(getSystemConfig("優先期時數"));
  return (isNaN(v) || v < 0) ? 48 : v;
}

/**
 * 取消截止天數（值勤日前幾天內不可自助取消）
 * 未設定時預設 7 天。供自助取消功能使用
 */
function getCancelDeadlineDays() {
  var v = parseInt(getSystemConfig("取消截止天數"));
  return (isNaN(v) || v < 0) ? 7 : v;
}

/**
 * 每人每月自助取消額度（跨所有活動累計）
 * 未設定時預設 2 次。供自助取消功能使用
 */
function getSelfCancelQuota() {
  var v = parseInt(getSystemConfig("自助取消月額度"));
  return (isNaN(v) || v < 0) ? 2 : v;
}

// ================================================
// ★ 依申請月份找對應公告時間
// 若同一月份有多筆公告，取最新（最晚）的那筆
// ================================================
function getAnnounceTimeForMonth(activityName, yearMonth) {
  var data = getSheetData(SHEET.ANNOUNCE);
  var result = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== activityName) continue;
    var m1 = (data[i][2] || "").toString().trim();
    var m2 = (data[i][3] || "").toString().trim();
    if (m1 !== yearMonth && m2 !== yearMonth) continue;
    var openTime = parseOpenDateTime(data[i][0].toString().trim());
    if (!openTime) continue;
    if (!result || openTime > result) result = openTime;
  }
  return result;
}

// ================================================
// 公告設定：多批次月份開放邏輯
// ================================================

function buildAnnounceCache(activityName) {
  var data  = getSheetData(SHEET.ANNOUNCE);
  var now   = new Date();
  var cache = { open: [], pending: [], latestAnnounce: null };

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][1] !== activityName) continue;
    var raw      = data[i][0].toString().trim();
    var openTime = parseOpenDateTime(raw);
    if (!openTime) continue;
    var m1 = (data[i][2] || "").toString().trim();
    var m2 = (data[i][3] || "").toString().trim();

    if (now >= openTime) {
      if (m1) cache.open.push(m1);
      if (m2) cache.open.push(m2);
    } else {
      cache.pending.push({ raw: raw, openTime: openTime, months: [m1, m2].filter(function(m){ return m; }) });
    }
    if (!cache.latestAnnounce || openTime > cache.latestAnnounce) cache.latestAnnounce = openTime;
  }
  return cache;
}

function getOpenMonthsFromCache(cache) {
  var map = {};
  cache.open.forEach(function(m) { map[m] = true; });
  return Object.keys(map);
}

function getNextOpenTimeFromCache(cache) {
  if (!cache.pending.length) return null;
  var next = null;
  cache.pending.forEach(function(p) { if (!next || p.openTime < next) next = p.openTime; });
  return next;
}

function getPendingAnnouncesFromCache(cache) {
  var pending = [];
  cache.pending.forEach(function(p) {
    var found = false;
    for (var i = 0; i < pending.length; i++) {
      if (pending[i].openTime === p.raw) {
        p.months.forEach(function(m) { if (pending[i].months.indexOf(m) < 0) pending[i].months.push(m); });
        found = true; break;
      }
    }
    if (!found) pending.push({ openTime: formatOpenDateTime(p.openTime), months: p.months.slice() });
  });
  pending.sort(function(a, b) { return parseOpenDateTime(a.openTime) - parseOpenDateTime(b.openTime); });
  return pending;
}

function getOpenMonths(activityName) {
  var data = getSheetData(SHEET.ANNOUNCE);
  var now  = new Date();
  var openMonths = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][1] !== activityName) continue;
    var raw = data[i][0].toString().trim();
    var openTime = parseOpenDateTime(raw);
    if (!openTime || now < openTime) continue;
    // 注意：到期日不影響開放月份，避免缺額出現卻無法排班
    var m1 = (data[i][2] || "").toString().trim();
    var m2 = (data[i][3] || "").toString().trim();
    if (m1) openMonths[m1] = true;
    if (m2) openMonths[m2] = true;
  }
  return Object.keys(openMonths);
}

function getNextOpenTime(activityName) {
  var data = getSheetData(SHEET.ANNOUNCE);
  var now  = new Date(), next = null;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][1] !== activityName) continue;
    var openTime = parseOpenDateTime(data[i][0].toString().trim());
    if (!openTime || now >= openTime) continue;
    if (!next || openTime < next) next = openTime;
  }
  return next;
}

function checkMonthOpen(activityName, dateStr) {
  var dp    = dateStr.split("/");
  var month = dp[0] + "/" + parseInt(dp[1]);
  var openMonths = getOpenMonths(activityName);
  for (var i = 0; i < openMonths.length; i++) {
    if (openMonths[i] === month) return true;
  }
  var nextTime = getNextOpenTime(activityName);
  if (nextTime) return "申請失敗：" + month + " 尚未開放登記，最近開放時間為 " + formatOpenDateTime(nextTime) + "，請屆時再申請。";
  return "申請失敗：" + month + " 尚未開放登記，請聯繫管理人員。";
}

function getLatestAnnounceDate(activityName) {
  var data = getSheetData(SHEET.ANNOUNCE);
  var latest = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== activityName || !data[i][0]) continue;
    var d = parseOpenDateTime(data[i][0].toString().trim());
    if (d && (!latest || d > latest)) latest = d;
  }
  return latest;
}

// ================================================
// 月份人數設定
// ================================================

function getMonthQuota(activityName, yearMonth) {
  try {
    var data = getSheetData(SHEET.MONTH_QUOTA);
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() !== activityName) continue;
      if (data[i][1].toString().trim() !== yearMonth) continue;
      function parseQuota(val) {
        var s = val.toString().trim();
        if (s === '') return null;
        var n = parseInt(s);
        return isNaN(n) ? null : n;
      }
      return {
        weekdayMax:  parseInt(data[i][2]) || 1,
        holidayMax:  parseInt(data[i][3]) || 2,
        priorQuota: {
          general_pass: parseQuota(data[i][4]),
          general_fail: parseQuota(data[i][5]),
          new_pass:     parseQuota(data[i][6]),
          new_fail:     parseQuota(data[i][7])
        },
        normalQuota: {
          general_pass: parseQuota(data[i][8]),
          general_fail: parseQuota(data[i][9]),
          new_pass:     parseQuota(data[i][10]),
          new_fail:     parseQuota(data[i][11])
        }
      };
    }
  } catch(e) {
    Logger.log("月份人數設定讀取失敗：" + e.toString());
  }
  return null;
}

function getDefaultQuotas() {
  function parseQ(val) {
    if (!val || val.toString().trim() === '') return null;
    var n = parseInt(val);
    return isNaN(n) ? null : n;
  }
  return {
    days: parseInt(getSystemConfig("優先排班天數")) || 3,
    priorQuota: {
      general_pass: parseQ(getSystemConfig("預設優先配額_一般_通過")) !== null ? parseQ(getSystemConfig("預設優先配額_一般_通過")) : 4,
      general_fail: parseQ(getSystemConfig("預設優先配額_一般_未通過")) !== null ? parseQ(getSystemConfig("預設優先配額_一般_未通過")) : 4,
      new_pass:     parseQ(getSystemConfig("預設優先配額_新志工_通過")) !== null ? parseQ(getSystemConfig("預設優先配額_新志工_通過")) : 2,
      new_fail:     parseQ(getSystemConfig("預設優先配額_新志工_未通過")) !== null ? parseQ(getSystemConfig("預設優先配額_新志工_未通過")) : 2
    },
    normalQuota: {
      general_pass: parseQ(getSystemConfig("預設一般配額_一般_通過")),
      general_fail: parseQ(getSystemConfig("預設一般配額_一般_未通過")),
      new_pass:     parseQ(getSystemConfig("預設一般配額_新志工_通過")),
      new_fail:     parseQ(getSystemConfig("預設一般配額_新志工_未通過"))
    }
  };
}

function getQuotaKey(volunteer) {
  var isNew  = (volunteer.type === VOLUNTEER_TYPE.NEW);
  var isPass = (volunteer.teleCert === CERT.PASS);
  if (!isNew && isPass)  return 'general_pass';
  if (!isNew && !isPass) return 'general_fail';
  if (isNew  && isPass)  return 'new_pass';
  return 'new_fail';
}

function saveMonthQuotas(activityName, months) {
  if (!months || months.length === 0) return;
  var sheet = getSheet(SHEET.MONTH_QUOTA);
  var data  = sheet.getDataRange().getDisplayValues();
  months.forEach(function(m) {
    if (!m.yearMonth) return;
    function toVal(v) { return (v === null || v === undefined || v === '') ? '' : parseInt(v); }
    var row = [
      activityName, m.yearMonth,
      parseInt(m.weekdayMax) || 1, parseInt(m.holidayMax) || 2,
      toVal(m.prior_general_pass), toVal(m.prior_general_fail),
      toVal(m.prior_new_pass),     toVal(m.prior_new_fail),
      toVal(m.normal_general_pass), toVal(m.normal_general_fail),
      toVal(m.normal_new_pass),     toVal(m.normal_new_fail)
    ];
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === activityName && data[i][1].toString().trim() === m.yearMonth) {
        sheet.getRange(i + 1, 1, 1, 12).setValues([row]);
        found = true; break;
      }
    }
    if (!found) { sheet.appendRow(row); data.push(row); }
  });
  invalidateSheetData(SHEET.MONTH_QUOTA);
}

function getAdminMonthQuotas() {
  try {
    var data = getSheetData(SHEET.MONTH_QUOTA);
    var list = [];
    function pq(v) { var s = v ? v.toString().trim() : ''; return s === '' ? '' : parseInt(s); }
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1]) continue;
      list.push({
        rowIndex:   i + 1,
        activity:   data[i][0].toString().trim(),
        yearMonth:  data[i][1].toString().trim(),
        weekdayMax: parseInt(data[i][2]) || 1,
        holidayMax: parseInt(data[i][3]) || 2,
        prior_general_pass: pq(data[i][4]), prior_general_fail: pq(data[i][5]),
        prior_new_pass:     pq(data[i][6]), prior_new_fail:     pq(data[i][7]),
        normal_general_pass: pq(data[i][8]), normal_general_fail: pq(data[i][9]),
        normal_new_pass:     pq(data[i][10]), normal_new_fail:    pq(data[i][11])
      });
    }
    return list;
  } catch(e) {
    Logger.log("讀取月份人數設定失敗：" + e.toString());
    return [];
  }
}

function updateMonthQuota(rowIndex, activity, yearMonth, weekdayMax, holidayMax,
  prior_gp, prior_gf, prior_np, prior_nf,
  normal_gp, normal_gf, normal_np, normal_nf) {
  var sheet = getSheet(SHEET.MONTH_QUOTA);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return "找不到對應設定。";
  function toVal(v) { return (v === null || v === undefined || v === '') ? '' : parseInt(v); }
  sheet.getRange(rowIndex, 1, 1, 12).setValues([[
    activity, yearMonth,
    parseInt(weekdayMax) || 1, parseInt(holidayMax) || 2,
    toVal(prior_gp), toVal(prior_gf), toVal(prior_np), toVal(prior_nf),
    toVal(normal_gp), toVal(normal_gf), toVal(normal_np), toVal(normal_nf)
  ]]);
  invalidateSheetData(SHEET.MONTH_QUOTA);
  return "月份人數設定已更新！";
}

function deleteMonthQuota(rowIndex) {
  var sheet = getSheet(SHEET.MONTH_QUOTA);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return "找不到對應設定。";
  sheet.deleteRow(rowIndex);
  invalidateSheetData(SHEET.MONTH_QUOTA);
  return "已刪除月份人數設定。";
}

// ================================================
// 特殊日期設定
// ================================================

function getSpecialDateConfigs() {
  try {
    var data = getSheetData(SHEET.SPECIAL_DATE);
    var configs = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      configs.push({
        date:       data[i][0].toString().trim(),
        activity:   data[i][1].toString().trim(),
        type:       data[i][2].toString().trim(),
        timeSlot:   data[i][3].toString().trim(),
        maxPerSlot: parseInt(data[i][4]) || 2,
        enabled:    data[i][5].toString().trim() !== "否",
        note:       data[i][6].toString().trim()
      });
    }
    return configs;
  } catch(e) { return []; }
}

function getSpecialDateConfigsForDate(dateStr, activityName) {
  var configs = getSpecialDateConfigs();
  var result  = [];
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].date === dateStr &&
       (configs[i].activity === activityName || configs[i].activity === "")) {
      result.push(configs[i]);
    }
  }
  return result;
}

function isDateDisabled(dateStr, activityName) {
  var configs = getSpecialDateConfigsForDate(dateStr, activityName);
  if (configs.length === 0) return false;
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].enabled) return false;
  }
  return true;
}

// ================================================
// 暑假設定
// ================================================

function getSummerConfigs() {
  try {
    var data = getSheetData(SHEET.SUMMER);
    var configs = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || !data[i][1]) continue;
      configs.push({
        startDate:  data[i][0].toString().trim(),
        endDate:    data[i][1].toString().trim(),
        activity:   data[i][2].toString().trim(),
        maxPerSlot: parseInt(data[i][3]) || 2,
        slotGroup:  data[i][4].toString().trim() || "平日2時段",
        note:       data[i][5].toString().trim()
      });
    }
    return configs;
  } catch(e) { return []; }
}

function getSummerConfig(dateStr, activityName) {
  var configs = getSummerConfigs();
  var dateISO = dateToISO(dateStr);
  var d = new Date(dateISO);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  for (var i = 0; i < configs.length; i++) {
    var cfg = configs[i];
    if (cfg.activity && cfg.activity !== activityName) continue;
    var start = new Date(dateToISO(cfg.startDate));
    var end   = new Date(dateToISO(cfg.endDate));
    if (isNaN(start.getTime()) || isNaN(end.getTime())) continue;
    start.setHours(0, 0, 0, 0); end.setHours(0, 0, 0, 0);
    if (d >= start && d <= end) return cfg;
  }
  return null;
}

// ================================================
// 跨活動衝突檢查
// ================================================

function slotToMinutes(timeSlot) {
  var parts = timeSlot.split("-");
  if (parts.length < 2) return null;
  function toMin(t) {
    var p = padTime(t.trim()).split(":");
    return parseInt(p[0]) * 60 + parseInt(p[1]);
  }
  return { start: toMin(parts[0]), end: toMin(parts[1]) };
}

function isTimeOverlap(slotA, slotB) {
  var a = slotToMinutes(slotA);
  var b = slotToMinutes(slotB);
  if (!a || !b) return false;
  return a.start < b.end && b.start < a.end;
}

function checkCrossActivityConflict(volunteerId, currentActivityName, dateStr, currentTimeSlot) {
  var blocked = [], warning = [];
  var overviewData = getSheetData(SHEET.OVERVIEW);
  for (var i = 1; i < overviewData.length; i++) {
    var rowActivity = overviewData[i][0];
    if (rowActivity === currentActivityName) continue;
    if (overviewData[i][1] !== dateStr) continue;
    var matched = false;
    for (var v = 0; v < VOL_COLS.length; v++) {
      if (overviewData[i][VOL_COLS[v].id] == volunteerId) { matched = true; break; }
    }
    if (!matched) continue;
    var existingSlot = overviewData[i][2];
    var label = rowActivity + " " + existingSlot;
    if (currentTimeSlot && isTimeOverlap(currentTimeSlot, existingSlot)) {
      blocked.push(label);
    } else { warning.push(label); }
  }
  var rosterData = getSheetData(SHEET.SESSION_ROSTER);
  for (var j = 1; j < rosterData.length; j++) {
    if (rosterData[j][2] != volunteerId) continue;
    if (rosterData[j][0] !== dateStr) continue;
    var sessionActName = getSessionActivityName(dateStr, rosterData[j][1]);
    if (!sessionActName || sessionActName === currentActivityName) continue;
    var existingSlot2 = rosterData[j][1];
    var label2 = sessionActName + " " + existingSlot2;
    if (currentTimeSlot && isTimeOverlap(currentTimeSlot, existingSlot2)) {
      blocked.push(label2);
    } else { warning.push(label2); }
  }
  return { blocked: blocked, warning: warning };
}

function getSessionActivityName(dateStr, timeSlot) {
  var data = getSheetData(SHEET.SESSION);
  var normalizedStart = padTime(timeSlot.split("-")[0]);
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === dateStr && padTime(data[i][2]) === normalizedStart) return data[i][0];
  }
  var overviewData = getSheetData(SHEET.OVERVIEW);
  for (var j = 1; j < overviewData.length; j++) {
    if (overviewData[j][1] === dateStr && overviewData[j][2] === timeSlot) return overviewData[j][0];
  }
  return null;
}

function updateOverviewSheetCached(sheet, overviewData, activityName, slot, name, id, action, specialCache, monthQCache) {
  for (var i = 1; i < overviewData.length; i++) {
    var rowSlot = overviewData[i][1] + " " + overviewData[i][2];
    if (overviewData[i][0] !== activityName || rowSlot !== slot) continue;
    var row = i + 1;
    if (action === "add") {
      var filled = false;
      for (var v = 0; v < VOL_COLS.length; v++) {
        if (!overviewData[i][VOL_COLS[v].name]) {
          sheet.getRange(row, VOL_COLS[v].name + 1).setValue(name);
          sheet.getRange(row, VOL_COLS[v].id   + 1).setValue(id);
          filled = true; break;
        }
      }
      if (!filled) return;
    } else {
      for (var v2 = 0; v2 < VOL_COLS.length; v2++) {
        if (overviewData[i][VOL_COLS[v2].id] == id) {
          sheet.getRange(row, VOL_COLS[v2].name + 1, 1, 2).clearContent(); break;
        }
      }
    }
    var d = sheet.getRange(row, 1, 1, 12).getDisplayValues()[0];
    var count  = countVolunteers(d);
    var max    = getBinocularMaxCached(d[1], d[2], d[0], specialCache, monthQCache);
    var status = count >= max ? "已滿" : count > 0 ? "部分排班" : "可排班";
    sheet.getRange(row, STATUS_COL).setValue(status);
    break;
  }
}

function checkCrossActivityConflictCached(volunteerId, currentActivityName, dateStr, currentTimeSlot, overviewData) {
  var blocked = [], warning = [];
  for (var i = 1; i < overviewData.length; i++) {
    var rowActivity = overviewData[i][0];
    if (rowActivity === currentActivityName) continue;
    if (overviewData[i][1] !== dateStr) continue;
    var matched = false;
    for (var v = 0; v < VOL_COLS.length; v++) {
      if (overviewData[i][VOL_COLS[v].id] == volunteerId) { matched = true; break; }
    }
    if (!matched) continue;
    var existingSlot = overviewData[i][2];
    var label = rowActivity + " " + existingSlot;
    if (currentTimeSlot && isTimeOverlap(currentTimeSlot, existingSlot)) {
      blocked.push(label);
    } else { warning.push(label); }
  }
  var rosterData = getSheetData(SHEET.SESSION_ROSTER);
  for (var j = 1; j < rosterData.length; j++) {
    if (rosterData[j][2] != volunteerId) continue;
    if (rosterData[j][0] !== dateStr) continue;
    var sessionActName = getSessionActivityNameCached(dateStr, rosterData[j][1], overviewData);
    if (!sessionActName || sessionActName === currentActivityName) continue;
    var existingSlot2 = rosterData[j][1];
    var label2 = sessionActName + " " + existingSlot2;
    if (currentTimeSlot && isTimeOverlap(currentTimeSlot, existingSlot2)) {
      blocked.push(label2);
    } else { warning.push(label2); }
  }
  return { blocked: blocked, warning: warning };
}

function getSessionActivityNameCached(dateStr, timeSlot, overviewData) {
  var data = getSheetData(SHEET.SESSION);
  var normalizedStart = padTime(timeSlot.split("-")[0]);
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === dateStr && padTime(data[i][2].toString()) === normalizedStart) return data[i][0];
  }
  for (var j = 1; j < overviewData.length; j++) {
    if (overviewData[j][1] === dateStr && overviewData[j][2] === timeSlot) return overviewData[j][0];
  }
  return null;
}

// ================================================
// 二觀型排班
// ================================================

function processBinocularRequest(volunteer, activity, scheduleSlot, notes) {
  var slotParts0 = scheduleSlot.split(" ");
  var overviewSheet  = getSheet(SHEET.OVERVIEW);
  var overviewData   = getSheetData(SHEET.OVERVIEW);
  var specialCache   = buildSpecialDateCache(activity.name);
  var monthQCache    = buildMonthQuotaCache();
  var announceCache  = buildAnnounceCache(activity.name);

  var dp = slotParts0[0].split("/");
  var month = dp[0] + "/" + parseInt(dp[1]);
  var openMonths = getOpenMonthsFromCache(announceCache);
  var isOpen = false;
  for (var om = 0; om < openMonths.length; om++) { if (openMonths[om] === month) { isOpen = true; break; } }
  if (!isOpen) {
    var nextTime = getNextOpenTimeFromCache(announceCache);
    if (nextTime) return "申請失敗：" + month + " 尚未開放登記，最近開放時間為 " + formatOpenDateTime(nextTime) + "，請屆時再申請。";
    return "申請失敗：" + month + " 尚未開放登記，請聯繫管理人員。";
  }

  for (var i = 1; i < overviewData.length; i++) {
    var rowSlot = overviewData[i][1] + " " + overviewData[i][2];
    if (overviewData[i][0] !== activity.name || rowSlot !== scheduleSlot) continue;
    if (overviewData[i][4].toString().trim() == volunteer.id.toString().trim() ||
        overviewData[i][6].toString().trim() == volunteer.id.toString().trim() ||
        overviewData[i][8].toString().trim() == volunteer.id.toString().trim() ||
        overviewData[i][10].toString().trim() == volunteer.id.toString().trim())
      return "申請失敗：您已有排班。";
    var maxVol     = getBinocularMaxCached(overviewData[i][1], overviewData[i][2], activity.name, specialCache, monthQCache);
    var currentVol = (overviewData[i][3] ? 1 : 0) + (overviewData[i][5] ? 1 : 0) +
                     (overviewData[i][7] ? 1 : 0) + (overviewData[i][9] ? 1 : 0);
    if (currentVol >= maxVol) return "申請失敗：該時段已滿。";
    volunteer.activityName = activity.name;
    var quotaCheck = checkBinocularQuotaCached(volunteer, scheduleSlot, overviewData, monthQCache);
    if (quotaCheck !== true) return quotaCheck;
    var conflictResult = checkCrossActivityConflictCached(volunteer.id, activity.name, overviewData[i][1], overviewData[i][2], overviewData);
    if (conflictResult.blocked.length > 0) {
      return "申請失敗：您在 " + overviewData[i][1] + " 已有時間重疊的排班：\n" +
        conflictResult.blocked.map(function(c) { return "  • " + c; }).join("\n") + "\n請聯繫管理人員處理。";
    }
    updateOverviewSheetCached(overviewSheet, overviewData, activity.name, scheduleSlot, volunteer.name, volunteer.id, "add", specialCache, monthQCache);
    invalidateSheetData(SHEET.OVERVIEW);   // ★ 批次申請時確保下一筆讀到最新資料
    createCalendarEvent(activity.calendarId, scheduleSlot, volunteer.name);
    logFormResponse(volunteer.name, volunteer.id, activity.name, scheduleSlot, "申請排班", notes, "成功");
    if (conflictResult.warning.length > 0) {
      return "申請成功！已同步至日曆。\n⚠️ 提醒：您當天尚有其他排班（" + conflictResult.warning.join("、") + "），請注意時間安排。";
    }
    return "申請成功！已同步至日曆。";
  }
  return "申請失敗：找不到對應時段。";
}

// ★ v1.2：優先期時數改讀後台設定
function checkBinocularQuotaCached(volunteer, scheduleSlot, overviewData, monthQCache) {
  var slotParts = scheduleSlot.split(" ");
  var dateStr   = slotParts[0];
  var dp        = dateStr.split("/");
  var yearMonth = dp[0] + "/" + parseInt(dp[1]);
  var actName   = volunteer.activityName || "望遠鏡二觀";

  var monthQ = monthQCache[actName + "|" + yearMonth] ? {
    priorQuota: {
      general_pass: monthQCache[actName + "|" + yearMonth].prior_gp !== undefined ? monthQCache[actName + "|" + yearMonth].prior_gp : null,
      general_fail: monthQCache[actName + "|" + yearMonth].prior_gf !== undefined ? monthQCache[actName + "|" + yearMonth].prior_gf : null,
      new_pass:     monthQCache[actName + "|" + yearMonth].prior_np !== undefined ? monthQCache[actName + "|" + yearMonth].prior_np : null,
      new_fail:     monthQCache[actName + "|" + yearMonth].prior_nf !== undefined ? monthQCache[actName + "|" + yearMonth].prior_nf : null
    },
    normalQuota: {
      general_pass: monthQCache[actName + "|" + yearMonth].normal_gp !== undefined ? monthQCache[actName + "|" + yearMonth].normal_gp : null,
      general_fail: monthQCache[actName + "|" + yearMonth].normal_gf !== undefined ? monthQCache[actName + "|" + yearMonth].normal_gf : null,
      new_pass:     monthQCache[actName + "|" + yearMonth].normal_np !== undefined ? monthQCache[actName + "|" + yearMonth].normal_np : null,
      new_fail:     monthQCache[actName + "|" + yearMonth].normal_nf !== undefined ? monthQCache[actName + "|" + yearMonth].normal_nf : null
    }
  } : null;

  var defaults = getDefaultQuotas();
  var key      = getQuotaKey(volunteer);

  // ★ v1.2：優先期時數由後台設定（預設 48 小時）
  var priorHours   = getPriorityHours();
  var announceDate = getAnnounceTimeForMonth(actName, yearMonth);
  var now = new Date();
  var isPrior = announceDate && ((now - announceDate) <= priorHours * 60 * 60 * 1000);

  var quota;
  if (isPrior) {
    quota = (monthQ && monthQ.priorQuota[key] !== undefined && monthQ.priorQuota[key] !== null) ? monthQ.priorQuota[key] : defaults.priorQuota[key];
  } else {
    quota = (monthQ && monthQ.normalQuota[key] !== undefined && monthQ.normalQuota[key] !== null) ? monthQ.normalQuota[key] : defaults.normalQuota[key];
  }

  if (quota === null || quota === undefined || quota === '') return true;
  if (quota === 0) return "申請失敗：依本月排班規定，您目前暫不開放申請。";

  var slotDate = new Date(dateToISO(dateStr));
  var year = slotDate.getFullYear(), month = slotDate.getMonth();
  var bookedCount = 0;
  for (var i = 1; i < overviewData.length; i++) {
    if (overviewData[i][0] !== actName) continue;
    if (overviewData[i][4] == volunteer.id || overviewData[i][6] == volunteer.id ||
        overviewData[i][8] == volunteer.id || overviewData[i][10] == volunteer.id) {
      var d = new Date(dateToISO(overviewData[i][1]));
      if (d.getFullYear() === year && d.getMonth() === month) bookedCount++;
    }
  }

  if (bookedCount >= quota) {
    if (isPrior) {
      return "申請失敗：優先排班期間，您本月已達排班上限（" + quota + "個時段）。如需加排，請於公告開放 " + priorHours + " 小時後再申請。謝謝。";
    } else {
      return "申請失敗：您本月已達排班上限（" + quota + "個時段）。";
    }
  }
  return true;
}

// ================================================
// 場次型排班
// ================================================

function processSessionRequest(volunteer, activity, scheduleSlot, notes) {
  var openCheck = checkRegistrationOpen(activity.name, scheduleSlot);
  if (openCheck !== true) return openCheck;
  var sessionRosterSheet = getSheet(SHEET.SESSION_ROSTER);
  var slotParts     = scheduleSlot.split(" ");
  var dateStr       = slotParts[0];
  var timeSlot      = slotParts[1];
  var sessionConfig = getSessionConfig(activity.name, dateStr, timeSlot);
  if (!sessionConfig) return "申請失敗：找不到場次設定。";
  if (!sessionConfig.openDateTime || !sessionConfig.openDateTime.trim()) {
    return "申請失敗：此場次尚未設定開放時間，請聯繫管理人員。";
  }
  var odt = parseSessionOpenTime(sessionConfig.openDateTime);
  if (odt && new Date() < odt) {
    var y  = odt.getFullYear(), mo = odt.getMonth()+1, d = odt.getDate();
    var h  = odt.getHours(), mi = odt.getMinutes();
    var ts = y+"/"+mo+"/"+d+" "+(h<10?"0"+h:h)+":"+(mi<10?"0"+mi:mi);
    return "申請失敗：此場次尚未開放登記，開放時間為 " + ts + "，請屆時再申請。";
  }
  var rosterData = getSheetData(SHEET.SESSION_ROSTER);
  for (var i = 1; i < rosterData.length; i++) {
    if (rosterData[i][0] === dateStr && rosterData[i][1] === timeSlot && rosterData[i][2] === volunteer.id)
      return "申請失敗：您已報名此場次。";
  }
  var generalCount = countSessionBookings(dateStr, timeSlot, activity.name, VOLUNTEER_TYPE.GENERAL);
  var newCount     = countSessionBookings(dateStr, timeSlot, activity.name, VOLUNTEER_TYPE.NEW);
  if (volunteer.type === VOLUNTEER_TYPE.NEW) {
    if (newCount >= sessionConfig.newMax) return "申請失敗：該場次新志工名額已滿（" + sessionConfig.newMax + " 名）。";
  } else {
    if (generalCount >= sessionConfig.generalMax) return "申請失敗：該場次一般志工名額已滿（" + sessionConfig.generalMax + " 名）。";
  }
  var conflictResult = checkCrossActivityConflict(volunteer.id, activity.name, dateStr, timeSlot);
  if (conflictResult.blocked.length > 0) {
    return "申請失敗：您在 " + dateStr + " 已有時間重疊的排班：\n" +
      conflictResult.blocked.map(function(c) { return "  • " + c; }).join("\n") + "\n請聯繫管理人員處理。";
  }
  sessionRosterSheet.appendRow([dateStr, timeSlot, volunteer.id, volunteer.name, volunteer.type, new Date()]);
  invalidateSheetData(SHEET.SESSION_ROSTER);   // ★ 批次申請時確保下一筆讀到最新資料
  updateSessionOverviewStatus(activity.name, dateStr, timeSlot, sessionConfig);
  createCalendarEvent(activity.calendarId, scheduleSlot, volunteer.name);
  logFormResponse(volunteer.name, volunteer.id, activity.name, scheduleSlot, "申請排班", notes, "成功");
  if (conflictResult.warning.length > 0) {
    return "申請成功！已同步至日曆。\n⚠️ 提醒：您當天尚有其他排班（" + conflictResult.warning.join("、") + "），請注意時間安排。";
  }
  return "申請成功！已同步至日曆。";
}

// ================================================
// 場次輔助函數
// ================================================

function getSessionConfig(activityName, date, timeSlot) {
  var data = getSheetData(SHEET.SESSION);
  var normalizedStart = padTime(timeSlot.split("-")[0]);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === activityName && data[i][1] === date && padTime(data[i][2]) === normalizedStart)
      return { generalMax: parseInt(data[i][4]) || 0, newMax: parseInt(data[i][5]) || 0, openDateTime: (data[i][6] || "").toString().trim() };
  }
  return null;
}

function countSessionBookings(dateStr, timeSlot, activityName, volunteerType) {
  var data = getSheetData(SHEET.SESSION_ROSTER);
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === dateStr && data[i][1] === timeSlot && data[i][4] === volunteerType) count++;
  }
  return count;
}

function updateSessionOverviewStatus(activityName, dateStr, timeSlot, sessionConfig) {
  var overviewSheet = getSheet(SHEET.OVERVIEW);
  var data = getSheetData(SHEET.OVERVIEW);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === activityName && data[i][1] === dateStr && data[i][2] === timeSlot) {
      var gc    = countSessionBookings(dateStr, timeSlot, activityName, VOLUNTEER_TYPE.GENERAL);
      var nc    = countSessionBookings(dateStr, timeSlot, activityName, VOLUNTEER_TYPE.NEW);
      var total = gc + nc;
      var maxT  = sessionConfig.generalMax + sessionConfig.newMax;
      overviewSheet.getRange(i + 1, 8).setValue(total >= maxT ? "已滿" : total > 0 ? "部分排班" : "可排班");
      invalidateSheetData(SHEET.OVERVIEW);
      break;
    }
  }
}

function getBinocularMax(dateStr, timeSlot, activityName) {
  var actName = activityName || "望遠鏡二觀";
  var day     = new Date(dateToISO(dateStr)).getDay();
  var isWeekend = (day === 0 || day === 6);
  var specialConfigs = getSpecialDateConfigsForDate(dateStr, actName);
  if (specialConfigs.length > 0) {
    for (var s = 0; s < specialConfigs.length; s++) {
      if (specialConfigs[s].timeSlot === timeSlot) {
        if (!specialConfigs[s].enabled) return 0;
        return specialConfigs[s].maxPerSlot;
      }
    }
    return 0;
  }
  var dp = dateStr.split("/");
  var yearMonth = dp[0] + "/" + parseInt(dp[1]);
  var monthQuota = getMonthQuota(actName, yearMonth);
  if (monthQuota) return isWeekend ? monthQuota.holidayMax : monthQuota.weekdayMax;
  return isWeekend ? BINOCULAR_MAX.WEEKEND : BINOCULAR_MAX.WEEKDAY;
}

function getBinocularMaxCached(dateStr, timeSlot, activityName, specialCache, monthQuotaCache) {
  var actName   = activityName || "望遠鏡二觀";
  var day       = new Date(dateToISO(dateStr)).getDay();
  var isWeekend = (day === 0 || day === 6);
  var specials = specialCache[dateStr + "|" + actName] || [];
  if (specials.length > 0) {
    for (var s = 0; s < specials.length; s++) {
      if (specials[s].timeSlot === timeSlot) {
        if (!specials[s].enabled) return 0;
        return specials[s].maxPerSlot;
      }
    }
    return 0;
  }
  var dp = dateStr.split("/");
  var yearMonth = dp[0] + "/" + parseInt(dp[1]);
  var monthQuota = monthQuotaCache[actName + "|" + yearMonth];
  if (monthQuota) return isWeekend ? monthQuota.holidayMax : monthQuota.weekdayMax;
  return isWeekend ? BINOCULAR_MAX.WEEKEND : BINOCULAR_MAX.WEEKDAY;
}

function buildSpecialDateCache(activityName) {
  var cache   = {};
  var configs = getSpecialDateConfigs();
  for (var i = 0; i < configs.length; i++) {
    var c = configs[i];
    if (c.activity && c.activity !== activityName) continue;
    var key = c.date + "|" + activityName;
    if (!cache[key]) cache[key] = [];
    cache[key].push(c);
  }
  return cache;
}

function buildMonthQuotaCache() {
  var cache = {};
  try {
    var data = getSheetData(SHEET.MONTH_QUOTA);
    function pq(v) { var s = v !== undefined && v !== null ? v.toString().trim() : ''; return s === '' ? null : parseInt(s); }
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1]) continue;
      var key = data[i][0].toString().trim() + "|" + data[i][1].toString().trim();
      cache[key] = {
        weekdayMax: parseInt(data[i][2]) || 1, holidayMax: parseInt(data[i][3]) || 2,
        prior_gp: pq(data[i][4]), prior_gf: pq(data[i][5]), prior_np: pq(data[i][6]), prior_nf: pq(data[i][7]),
        normal_gp: pq(data[i][8]), normal_gf: pq(data[i][9]), normal_np: pq(data[i][10]), normal_nf: pq(data[i][11])
      };
    }
  } catch(e) { Logger.log("buildMonthQuotaCache 失敗：" + e.toString()); }
  return cache;
}

function getBinocularSlotDefs(dateStr, activityName) {
  var actName = activityName || "望遠鏡二觀";
  var day     = new Date(dateToISO(dateStr)).getDay();
  var specialConfigs = getSpecialDateConfigsForDate(dateStr, actName);
  if (specialConfigs.length > 0) {
    var slots = [];
    for (var s = 0; s < specialConfigs.length; s++) {
      if (!specialConfigs[s].enabled) continue;
      var ts    = specialConfigs[s].timeSlot;
      var parts = ts.split("-");
      if (parts.length === 2) slots.push({ start: parts[0].trim(), end: parts[1].trim(), label: ts });
    }
    return slots;
  }
  return day === 6 ? BINOCULAR_SLOTS.SATURDAY : day === 0 ? BINOCULAR_SLOTS.SUNDAY : BINOCULAR_SLOTS.WEEKDAY;
}

// ★ v1.2：優先期時數改讀後台設定
function checkBinocularQuota(volunteer, scheduleSlot) {
  var slotParts = scheduleSlot.split(" ");
  var dateStr   = slotParts[0];
  var dp        = dateStr.split("/");
  var yearMonth = dp[0] + "/" + parseInt(dp[1]);
  var actName   = volunteer.activityName || "望遠鏡二觀";
  var monthQ    = getMonthQuota(actName, yearMonth);
  var defaults  = getDefaultQuotas();
  var key       = getQuotaKey(volunteer);

  // ★ v1.2：優先期時數由後台設定（預設 48 小時）
  var priorHours   = getPriorityHours();
  var announceDate = getAnnounceTimeForMonth(actName, yearMonth);
  var now = new Date();
  var isPrior = announceDate && ((now - announceDate) <= priorHours * 60 * 60 * 1000);

  var quota;
  if (isPrior) {
    quota = (monthQ && monthQ.priorQuota[key] !== undefined && monthQ.priorQuota[key] !== '') ? monthQ.priorQuota[key] : defaults.priorQuota[key];
  } else {
    quota = (monthQ && monthQ.normalQuota[key] !== undefined && monthQ.normalQuota[key] !== '') ? monthQ.normalQuota[key] : defaults.normalQuota[key];
  }
  if (quota === null || quota === undefined || quota === '') return true;
  if (quota === 0) return "申請失敗：依本月排班規定，您目前暫不開放申請。";
  var slotDate    = new Date(dateToISO(dateStr));
  var bookedCount = countMonthlyBookings(volunteer.id, actName, slotDate.getFullYear(), slotDate.getMonth());
  if (bookedCount >= quota) {
    if (isPrior) {
      return "申請失敗：優先排班期間，您本月已達排班上限（" + quota + "個時段）。如需加排，請於公告開放 " + priorHours + " 小時後再申請。謝謝。";
    } else {
      return "申請失敗：您本月已達排班上限（" + quota + "個時段）。";
    }
  }
  return true;
}

function checkRegistrationOpen(activityName, scheduleSlot) {
  var activity = getActivityConfig(activityName);
  if (!activity) return true;
  if (activity.type === ACTIVITY_TYPE.BINOCULAR) return true;
  if (!scheduleSlot) return true;
  var parts    = scheduleSlot.split(" ");
  var dateStr  = parts[0];
  var timeSlot = parts[1];
  var startT   = padTime(timeSlot.split("-")[0]);
  var data = getSheetData(SHEET.SESSION);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== activityName || data[i][1] !== dateStr) continue;
    if (padTime(data[i][2].toString()) !== startT) continue;
    var odtRaw = (data[i][6] || "").toString().trim();
    if (!odtRaw) return true;
    var odt = parseOpenDateTime(odtRaw);
    if (!odt) return true;
    if (new Date() >= odt) return true;
    return "申請失敗：此場次尚未開放報名，開放時間為 " + formatOpenDateTime(odt) + "，請屆時再報名。";
  }
  return true;
}

function countWorkDays(startDate, endDate) {
  var count = 0, current = new Date(startDate);
  current.setHours(0, 0, 0, 0);
  while (current <= endDate) {
    if (current.getDay() !== 1) count++;
    current.setDate(current.getDate() + 1);
  }
  return count;
}

function countMonthlyBookings(volunteerId, activityName, year, month) {
  var data = getSheetData(SHEET.OVERVIEW);
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== activityName) continue;
    if (data[i][4] == volunteerId || data[i][6] == volunteerId ||
        data[i][8] == volunteerId || data[i][10] == volunteerId) {
      var d = new Date(dateToISO(data[i][1]));
      if (d.getFullYear() === year && d.getMonth() === month) count++;
    }
  }
  return count;
}

var VOL_COLS = [
  { name: 3, id: 4 },
  { name: 5, id: 6 },
  { name: 7, id: 8 },
  { name: 9, id: 10 }
];
var STATUS_COL = 12;

function countVolunteers(row) {
  var count = 0;
  for (var v = 0; v < VOL_COLS.length; v++) { if (row[VOL_COLS[v].name]) count++; }
  return count;
}

function updateOverviewSheet(sheet, activityName, slot, name, id, action) {
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    var rowSlot = data[i][1] + " " + data[i][2];
    if (data[i][0] !== activityName || rowSlot !== slot) continue;
    var row = i + 1;
    if (action === "add") {
      var filled = false;
      for (var v = 0; v < VOL_COLS.length; v++) {
        if (!data[i][VOL_COLS[v].name]) {
          sheet.getRange(row, VOL_COLS[v].name + 1).setValue(name);
          sheet.getRange(row, VOL_COLS[v].id   + 1).setValue(id);
          filled = true; break;
        }
      }
      if (!filled) return;
    } else {
      for (var v2 = 0; v2 < VOL_COLS.length; v2++) {
        if (data[i][VOL_COLS[v2].id] == id) {
          sheet.getRange(row, VOL_COLS[v2].name + 1, 1, 2).clearContent(); break;
        }
      }
    }
    updateSlotStatus(sheet, row, data[i][1], data[i][2]);
    invalidateSheetData(SHEET.OVERVIEW);
    break;
  }
}

function updateSlotStatus(sheet, row, dateStr, timeSlot) {
  var d = sheet.getRange(row, 1, 1, 12).getDisplayValues()[0];
  if (dateStr  === undefined) dateStr  = d[1];
  if (timeSlot === undefined) timeSlot = d[2];
  var count = countVolunteers(d);
  var activityName = d[0];
  var max    = getBinocularMax(dateStr, timeSlot, activityName);
  var status = count >= max ? "已滿" : count > 0 ? "部分排班" : "可排班";
  sheet.getRange(row, STATUS_COL).setValue(status);
  try {
    var cacheSheet = getSheet(SHEET.CACHE);
    var cacheData  = cacheSheet.getDataRange().getDisplayValues();
    for (var c = 1; c < cacheData.length; c++) {
      if (cacheData[c][0] === activityName && cacheData[c][1] === dateStr && cacheData[c][2] === timeSlot) {
        cacheSheet.getRange(c + 1, 4).setValue(status); break;
      }
    }
  } catch(e) {}
}

function getAvailableScheduleSlots(activityName) {
  var now  = new Date();
  var data = getSheetData(SHEET.OVERVIEW);
  var slots = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== activityName || data[i][11] === "已滿" || !data[i][1] || !data[i][2]) continue;
    var slotStart = new Date(dateToISO(data[i][1]) + "T" + padTime(data[i][2].split("-")[0].trim()) + ":00");
    if (isNaN(slotStart.getTime()) || slotStart <= now) continue;
    slots.push(data[i][1] + " " + data[i][2]);
  }
  return slots.length > 0 ? slots : ["目前無可用時段"];
}

function getAvailableSessionSlots(activityName) {
  var now        = new Date();
  var rosterData = getSheetData(SHEET.SESSION_ROSTER);
  var countMap   = {};
  for (var r = 1; r < rosterData.length; r++) {
    var k = rosterData[r][0] + "|" + rosterData[r][1];
    countMap[k] = (countMap[k] || 0) + 1;
  }
  var sessionData = getSheetData(SHEET.SESSION);
  var slots = [];
  for (var i = 1; i < sessionData.length; i++) {
    if (sessionData[i][0] !== activityName) continue;
    var date      = sessionData[i][1].toString();
    var startTime = padTime(sessionData[i][2].toString());
    var endTime   = padTime(sessionData[i][3].toString());
    var timeSlot  = startTime + "-" + endTime;
    var totalMax  = (parseInt(sessionData[i][4]) || 0) + (parseInt(sessionData[i][5]) || 0);
    var dp        = date.split("/");
    if (dp.length !== 3) continue;
    var dateISO   = dp[0] + "-" + (dp[1].length < 2 ? "0"+dp[1] : dp[1]) + "-" + (dp[2].length < 2 ? "0"+dp[2] : dp[2]);
    var slotStart = new Date(dateISO + "T" + startTime + ":00");
    if (isNaN(slotStart.getTime()) || slotStart <= now) continue;
    if ((countMap[date + "|" + timeSlot] || 0) >= totalMax) continue;
    slots.push(date + " " + timeSlot);
  }
  return slots.length > 0 ? slots : ["目前無可用場次"];
}

function onEdit(e) {
  var range = e.range, sheet = range.getSheet();
  if (sheet.getName() !== SHEET.OVERVIEW) return;
  var row = range.getRow(), col = range.getColumn();
  var idCols = [5, 7, 9, 11];
  if (idCols.indexOf(col) >= 0 && row > 1 && !e.value && e.oldValue) {
    var rowData  = sheet.getRange(row, 1, 1, 3).getDisplayValues()[0];
    var activity = getActivityConfig(rowData[0]);
    if (activity) deleteCalendarEvent(activity.calendarId, rowData[1] + " " + rowData[2], e.oldValue);
    sheet.getRange(row, col - 1).clearContent();
    updateSlotStatus(sheet, row, rowData[1], rowData[2]);
  }
}

function generateBinocularSlots(announceDate, activityName, targetMonths) {
  var overviewSheet = getSheet(SHEET.OVERVIEW);
  var cacheSheet    = getSheet(SHEET.CACHE);
  var monthList = [];
  if (targetMonths && targetMonths.length > 0) {
    for (var m = 0; m < targetMonths.length; m++) {
      var parts = targetMonths[m].split("/");
      if (parts.length === 2) monthList.push(new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1));
    }
  }
  if (monthList.length === 0) {
    var announce = new Date(announceDate);
    monthList.push(new Date(announce.getFullYear(), announce.getMonth() + 2, 1));
  }
  var existingData = overviewSheet.getDataRange().getDisplayValues();
  var existingSet  = {};
  for (var e = 1; e < existingData.length; e++)
    existingSet[existingData[e][0] + "|" + existingData[e][1] + "|" + existingData[e][2]] = true;
  var cacheData = cacheSheet.getDataRange().getDisplayValues();
  var cacheSet  = {};
  for (var c = 1; c < cacheData.length; c++)
    cacheSet[cacheData[c][0] + "|" + cacheData[c][1] + "|" + cacheData[c][2]] = true;
  var newOverviewRows = [], newCacheRows = [], slotsAdded = 0;
  monthList.forEach(function(monthStart) {
    var year = monthStart.getFullYear(), month = monthStart.getMonth();
    var daysInMonth = new Date(year, month + 1, 0).getDate();
    for (var d = 1; d <= daysInMonth; d++) {
      var day = new Date(year, month, d).getDay();
      if (day === 1) continue;
      var dateStr  = year + "/" + (month + 1) + "/" + d;
      var slotDefs = getBinocularSlotDefs(dateStr, activityName);
      if (!slotDefs || slotDefs.length === 0) continue;
      var weekdayNames = ["星期日","星期一","星期二","星期三","星期四","星期五","星期六"];
      var weekdayStr   = weekdayNames[new Date(year, month, d).getDay()];
      slotDefs.forEach(function(slot) {
        var timeSlot = slot.start + "-" + slot.end;
        var key      = activityName + "|" + dateStr + "|" + timeSlot;
        if (!existingSet[key]) {
          newOverviewRows.push([activityName, dateStr, timeSlot, "", "", "", "", "", "", "", "", "可排班", weekdayStr]);
          existingSet[key] = true;
          slotsAdded++;
        }
        if (!cacheSet[key]) {
          newCacheRows.push([activityName, dateStr, timeSlot, "可排班"]);
          cacheSet[key] = true;
        }
      });
    }
  });
  if (newOverviewRows.length > 0)
    overviewSheet.getRange(overviewSheet.getLastRow() + 1, 1, newOverviewRows.length, 13).setValues(newOverviewRows);
  if (newCacheRows.length > 0)
    cacheSheet.getRange(cacheSheet.getLastRow() + 1, 1, newCacheRows.length, 4).setValues(newCacheRows);
  invalidateSheetData(SHEET.OVERVIEW);
  return slotsAdded;
}

function getPersonalSchedule(name, volunteerId, startMonth, endMonth) {
  var volunteer = validateVolunteer(name, volunteerId);
  if (!volunteer) return null;
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var activeRecords = [], totalMinutes = 0;
  var overviewData = getSheetData(SHEET.OVERVIEW);
  for (var i = 1; i < overviewData.length; i++) {
    var isMySlot = false;
    for (var v = 0; v < VOL_COLS.length; v++) {
      if (overviewData[i][VOL_COLS[v].id] == volunteerId) { isMySlot = true; break; }
    }
    if (!isMySlot) continue;
    var d = new Date(dateToISO(overviewData[i][1]));
    if (d >= today) {
      var slot  = overviewData[i][2];
      var parts = slot.match(/(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
      var mins  = 0;
      if (parts) {
        var s  = parts[1].split(":"), e2 = parts[2].split(":");
        mins   = (parseInt(e2[0])*60+parseInt(e2[1])) - (parseInt(s[0])*60+parseInt(s[1]));
        totalMinutes += mins;
      }
      activeRecords.push({ activity: overviewData[i][0], activityType: ACTIVITY_TYPE.BINOCULAR, date: overviewData[i][1], slot: slot, hours: (mins/60).toFixed(1) });
    }
  }
  var rosterData = getSheetData(SHEET.SESSION_ROSTER);
  for (var j = 1; j < rosterData.length; j++) {
    if (rosterData[j][2] == volunteerId) {
      var d2 = new Date(dateToISO(rosterData[j][0]));
      if (d2 >= today) {
        var slot2  = rosterData[j][1];
        var parts2 = slot2.match(/(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
        var mins2  = 0;
        if (parts2) {
          var s2 = parts2[1].split(":"), e3 = parts2[2].split(":");
          mins2  = (parseInt(e3[0])*60+parseInt(e3[1])) - (parseInt(s2[0])*60+parseInt(s2[1]));
          totalMinutes += mins2;
        }
        var sessionActName = getSessionActivityName(rosterData[j][0], rosterData[j][1]) || "（場次型活動）";
        activeRecords.push({ activity: sessionActName, activityType: ACTIVITY_TYPE.SESSION, date: rosterData[j][0], slot: slot2, hours: (mins2/60).toFixed(1) });
      }
    }
  }
  activeRecords.sort(function(a, b) {
    var aDate = new Date(dateToISO(a.date)), bDate = new Date(dateToISO(b.date));
    if (aDate.getTime() !== bDate.getTime()) return aDate - bDate;
    var aTime = a.slot.split("-")[0].trim(), bTime = b.slot.split("-")[0].trim();
    return aTime > bTime ? 1 : aTime < bTime ? -1 : 0;
  });
  var logRecords = [];
  try {
    var formData = getSheet(SHEET.FORM).getDataRange().getDisplayValues();
    for (var k = 1; k < formData.length; k++) {
      if (formData[k][2].toString().trim() !== volunteerId.toString().trim()) continue;
      var slotStr = formData[k][4].toString().trim();
      if (!slotStr) continue;
      var opTimeStr = "", dt = new Date(formData[k][0]);
      if (!isNaN(dt.getTime())) {
        var mm = dt.getMonth()+1, dd = dt.getDate(), hh = dt.getHours(), mi = dt.getMinutes();
        opTimeStr = dt.getFullYear() + "/" + (mm<10?"0":"") + mm + "/" + (dd<10?"0":"") + dd + " " + (hh<10?"0":"") + hh + ":" + (mi<10?"0":"") + mi;
      } else { opTimeStr = formData[k][0].toString(); }
      logRecords.push({
        opTime: opTimeStr, rowIndex: k,
        activity: formData[k][3].toString().trim(), slot: slotStr,
        operation: formData[k][5].toString().trim(), result: formData[k][7].toString().trim(),
        notes: formData[k][6].toString().trim()
      });
    }
  } catch(e) { Logger.log("讀取表單回應失敗：" + e.toString()); }
  logRecords.sort(function(a, b) { return b.rowIndex - a.rowIndex; });
  // ★ v1.5：加上自助取消資訊
  var cancelInfo   = getSelfCancelInfo(volunteerId);
  var deadlineDays = cancelInfo.deadlineDays;
  activeRecords.forEach(function(r) {
    var d = daysUntilSlot(r.date);
    r.daysUntil = d;
    if (d !== null && d <= deadlineDays) {
      r.cancelable  = false;
      r.blockReason = 'deadline';
    } else {
      r.cancelable  = true;
      r.blockReason = null;
    }
  });

  return {
    name: volunteer.name,
    activeRecords: activeRecords,
    totalHours: (totalMinutes / 60).toFixed(1),
    logRecords: logRecords,
    cancelInfo: cancelInfo
  };
}

function getVolunteerHours(name, volunteerId, startMonth, endMonth) {
  var volunteer = validateVolunteer(name, volunteerId);
  if (!volunteer) return null;
  var start = startMonth.substring(0, 7), end = endMonth.substring(0, 7);
  var data = getSheetData("時數資料");
  var latestMonth = "";
  for (var i = data.length - 1; i >= 1; i--) {
    var raw = data[i][0].toString().trim();
    if (raw) { latestMonth = raw.replace(/\//g, "-").substring(0, 7); break; }
  }
  var monthlyData = [], grandTotalMins = 0;
  for (var i = 1; i < data.length; i++) {
    var rowMonth = data[i][0].toString().trim().replace(/\//g, "-").substring(0, 7);
    var rowId    = data[i][1].toString().trim();
    if (!rowMonth || rowId !== volunteerId.toString().trim()) continue;
    if (rowMonth < start || rowMonth > end) continue;
    var minsVal  = parseInt(data[i][5])   || 0;
    var hoursVal = parseFloat(data[i][6]) || 0;
    var dateStr  = data[i][3] ? data[i][3].toString().trim() : "";
    var countVal = parseInt(data[i][4])   || 0;
    monthlyData.push({ month: rowMonth.replace("-", "/"), dates: dateStr, count: countVal, mins: minsVal, hours: hoursVal.toFixed(1) });
    grandTotalMins += minsVal;
  }
  monthlyData.sort(function(a, b) { return a.month > b.month ? 1 : -1; });
  var latestDisplay = "無紀錄";
  if (latestMonth) {
    var parts = latestMonth.split("-");
    latestDisplay = parts[0] + "年" + parseInt(parts[1]) + "月";
  }
  return { name: volunteer.name, startMonth: start.replace("-", "/"), endMonth: end.replace("-", "/"), latestMonth: latestDisplay, monthlyData: monthlyData, grandTotalMins: grandTotalMins, grandTotalHours: (grandTotalMins / 60).toFixed(1) };
}

function checkDateConflictForVolunteer(volunteerId, currentActivityName, dateStr) {
  var result = checkCrossActivityConflict(volunteerId, currentActivityName, dateStr, null);
  if (result.blocked.length === 0 && result.warning.length === 0) return null;
  return result;
}

// ================================================
// 管理者取消 / 換人排班
// ================================================

function adminGetVolunteerSchedule(volunteerId) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var records = [];
  var overviewData = getSheetData(SHEET.OVERVIEW);
  for (var i = 1; i < overviewData.length; i++) {
    var row = overviewData[i], matchedVol = null;
    for (var v = 0; v < VOL_COLS.length; v++) {
      if (row[VOL_COLS[v].id] == volunteerId) { matchedVol = v; break; }
    }
    if (matchedVol === null) continue;
    var d = new Date(dateToISO(row[1]));
    if (d < today) continue;
    records.push({ activity: row[0], activityType: ACTIVITY_TYPE.BINOCULAR, date: row[1], slot: row[2], volunteerId: row[VOL_COLS[matchedVol].id], volunteerName: row[VOL_COLS[matchedVol].name] });
  }
  var rosterData = getSheetData(SHEET.SESSION_ROSTER);
  for (var j = 1; j < rosterData.length; j++) {
    var r = rosterData[j];
    if (r[2] != volunteerId) continue;
    var d2 = new Date(dateToISO(r[0]));
    if (d2 < today) continue;
    var actName = getSessionActivityName(r[0], r[1]) || "（場次型）";
    records.push({ activity: actName, activityType: ACTIVITY_TYPE.SESSION, date: r[0], slot: r[1], volunteerId: r[2], volunteerName: r[3] });
  }
  records.sort(function(a, b) {
    var da = new Date(dateToISO(a.date)), db = new Date(dateToISO(b.date));
    if (da.getTime() !== db.getTime()) return da - db;
    return a.slot > b.slot ? 1 : -1;
  });
  return records;
}

function adminGetSlotRoster(activityName, dateStr) {
  var records  = [], activity = getActivityConfig(activityName);
  if (!activity) return records;
  if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
    var overviewData = getSheetData(SHEET.OVERVIEW);
    for (var i = 1; i < overviewData.length; i++) {
      var row = overviewData[i];
      if (row[0] !== activityName || row[1] !== dateStr) continue;
      for (var v = 0; v < VOL_COLS.length; v++) {
        if (row[VOL_COLS[v].name] && row[VOL_COLS[v].id]) {
          records.push({ activity: activityName, activityType: ACTIVITY_TYPE.BINOCULAR, date: dateStr, slot: row[2], volunteerId: row[VOL_COLS[v].id], volunteerName: row[VOL_COLS[v].name] });
        }
      }
    }
  } else {
    var rosterData = getSheetData(SHEET.SESSION_ROSTER);
    for (var j = 1; j < rosterData.length; j++) {
      var r = rosterData[j];
      if (r[0] !== dateStr) continue;
      var actName = getSessionActivityName(r[0], r[1]);
      if (actName !== activityName) continue;
      records.push({ activity: activityName, activityType: ACTIVITY_TYPE.SESSION, date: dateStr, slot: r[1], volunteerId: r[2], volunteerName: r[3] });
    }
  }
  records.sort(function(a, b) { return a.slot > b.slot ? 1 : -1; });
  return records;
}

function adminCancelSchedule(activityName, scheduleSlot, volunteerId, adminEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    invalidateSheetData();   // ★ 批次取消時確保每筆都讀到最新資料
    var activity = getActivityConfig(activityName);
    if (!activity) return "取消失敗：找不到活動設定。";
    var volunteerName = getVolunteerNameById(volunteerId);
    if (!volunteerName) return "取消失敗：找不到志工編號 " + volunteerId + "。";
    var cancelTime = formatNow();
    var adminNote  = "管理者 " + (adminEmail || "未知") + " @ " + cancelTime;
    var slotParts = scheduleSlot.split(" ");
    var dateStr   = slotParts[0], timeSlot = slotParts[1];
    if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
      var overviewSheet = getSheet(SHEET.OVERVIEW);
      var data = getSheetData(SHEET.OVERVIEW);
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] !== activityName || data[i][1] !== dateStr || data[i][2] !== timeSlot) continue;
        var hasVol = false;
        for (var v = 0; v < VOL_COLS.length; v++) { if (data[i][VOL_COLS[v].id] == volunteerId) { hasVol = true; break; } }
        if (hasVol) {
          updateOverviewSheet(overviewSheet, activityName, scheduleSlot, volunteerName, volunteerId, "remove");
          deleteCalendarEvent(activity.calendarId, scheduleSlot, volunteerName);
          logFormResponse(volunteerName, volunteerId, activityName, scheduleSlot, "取消排班（管理者）", adminNote, "取消");
          found = true; break;
        }
      }
      if (!found) return "取消失敗：找不到對應排班紀錄。";
    } else {
      var rosterSheet = getSheet(SHEET.SESSION_ROSTER);
      var rosterData  = getSheetData(SHEET.SESSION_ROSTER);
      var deleted = false;
      for (var j = rosterData.length - 1; j >= 1; j--) {
        if (rosterData[j][0] === dateStr && rosterData[j][1] === timeSlot && rosterData[j][2] == volunteerId) {
          rosterSheet.deleteRow(j + 1);
          invalidateSheetData(SHEET.SESSION_ROSTER);
          deleted = true; break;
        }
      }
      if (!deleted) return "取消失敗：找不到對應排班紀錄。";
      var sessionConfig = getSessionConfig(activityName, dateStr, timeSlot);
      if (sessionConfig) updateSessionOverviewStatus(activityName, dateStr, timeSlot, sessionConfig);
      deleteCalendarEvent(activity.calendarId, scheduleSlot, volunteerName);
      logFormResponse(volunteerName, volunteerId, activityName, scheduleSlot, "取消排班（管理者）", adminNote, "取消");
    }
    var volunteerObj = getVolunteerById(volunteerId);
    if (volunteerObj) sendVolunteerCancelNotification(volunteerObj, activityName, scheduleSlot, adminEmail, cancelTime);
    sendAdminCancelNotification({ id: volunteerId, name: volunteerName, email: "" }, activityName, scheduleSlot, adminEmail, cancelTime);
    return "已成功取消「" + volunteerName + "」在 " + scheduleSlot + " 的排班。";
  } catch(e) { return "系統錯誤：" + e.toString(); }
  finally { lock.releaseLock(); }
}

function adminReplaceSchedule(activityName, scheduleSlot, oldVolunteerId, newVolunteerId, adminEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    invalidateSheetData();   // ★ 確保讀到最新資料
    var activity = getActivityConfig(activityName);
    if (!activity) return "換人失敗：找不到活動設定。";
    var oldName = getVolunteerNameById(oldVolunteerId);
    if (!oldName) return "換人失敗：找不到原志工編號 " + oldVolunteerId + "。";
    var newVolunteer = getVolunteerById(newVolunteerId);
    if (!newVolunteer) return "換人失敗：找不到代班志工編號 " + newVolunteerId + "。";
    if (oldVolunteerId == newVolunteerId) return "換人失敗：代班志工與原志工相同。";
    var opTime    = formatNow();
    var adminNote = "管理者 " + (adminEmail || "未知") + " 換人 " + oldVolunteerId + " → " + newVolunteerId + " @ " + opTime;
    var slotParts = scheduleSlot.split(" ");
    var dateStr   = slotParts[0], timeSlot = slotParts[1];
    var conflictResult = checkCrossActivityConflict(newVolunteer.id, activityName, dateStr, timeSlot);
    if (conflictResult.blocked.length > 0) {
      return "換人失敗：代班志工在 " + dateStr + " 已有時間重疊的排班：\n" + conflictResult.blocked.map(function(c) { return "  • " + c; }).join("\n");
    }
    if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
      var ov = getSheetData(SHEET.OVERVIEW);
      for (var ci = 1; ci < ov.length; ci++) {
        if (ov[ci][0] === activityName && ov[ci][1] === dateStr && ov[ci][2] === timeSlot) {
          for (var vc = 0; vc < VOL_COLS.length; vc++) {
            if (ov[ci][VOL_COLS[vc].id] == newVolunteer.id) return "換人失敗：代班志工已在此時段有排班。";
          }
        }
      }
      if (newVolunteer.teleCert !== CERT.PASS && newVolunteer.type === VOLUNTEER_TYPE.NEW) {
        var dow = new Date(dateToISO(dateStr)).getDay();
        if (dow >= 2 && dow <= 5) return "換人失敗：代班志工（新志工）尚未通過望遠鏡驗收，平日時段暫無法單獨值班。";
      }
      var quotaCheck = checkBinocularQuota(newVolunteer, scheduleSlot);
      if (quotaCheck !== true) return "換人失敗（代班志工配額）：" + quotaCheck;
      var overviewSheet = getSheet(SHEET.OVERVIEW);
      var data = getSheetData(SHEET.OVERVIEW);
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] !== activityName || data[i][1] !== dateStr || data[i][2] !== timeSlot) continue;
        var targetVol = null;
        for (var vr = 0; vr < VOL_COLS.length; vr++) { if (data[i][VOL_COLS[vr].id] == oldVolunteerId) { targetVol = vr; break; } }
        if (targetVol === null) continue;
        var row = i + 1;
        overviewSheet.getRange(row, VOL_COLS[targetVol].name + 1).setValue(newVolunteer.name);
        overviewSheet.getRange(row, VOL_COLS[targetVol].id   + 1).setValue(newVolunteer.id);
        invalidateSheetData(SHEET.OVERVIEW);
        found = true; break;
      }
      if (!found) return "換人失敗：找不到原志工的排班紀錄。";
    } else {
      var sr = getSheetData(SHEET.SESSION_ROSTER);
      for (var ci2 = 1; ci2 < sr.length; ci2++) {
        if (sr[ci2][0] === dateStr && sr[ci2][1] === timeSlot && sr[ci2][2] == newVolunteer.id) return "換人失敗：代班志工已在此場次有排班。";
      }
      var sessionConfig2 = getSessionConfig(activityName, dateStr, timeSlot);
      if (sessionConfig2) {
        var rsd = getSheetData(SHEET.SESSION_ROSTER);
        var genCount = 0, newCount = 0;
        for (var ni = 1; ni < rsd.length; ni++) {
          if (rsd[ni][0] !== dateStr || rsd[ni][1] !== timeSlot || rsd[ni][2] == oldVolunteerId) continue;
          if (rsd[ni][4] === VOLUNTEER_TYPE.NEW) { newCount++; } else { genCount++; }
        }
        if (newVolunteer.type === VOLUNTEER_TYPE.NEW) { if (newCount >= sessionConfig2.newMax) return "換人失敗：該場次新志工名額已滿。"; }
        else { if (genCount >= sessionConfig2.generalMax) return "換人失敗：該場次一般志工名額已滿。"; }
      }
      var rosterSheet = getSheet(SHEET.SESSION_ROSTER);
      var rosterData  = getSheetData(SHEET.SESSION_ROSTER);
      var foundRow    = -1;
      for (var j = rosterData.length - 1; j >= 1; j--) {
        if (rosterData[j][0] === dateStr && rosterData[j][1] === timeSlot && rosterData[j][2] == oldVolunteerId) { foundRow = j + 1; break; }
      }
      if (foundRow < 0) return "換人失敗：找不到原志工的排班紀錄。";
      rosterSheet.getRange(foundRow, 3).setValue(newVolunteer.id);
      rosterSheet.getRange(foundRow, 4).setValue(newVolunteer.name);
      rosterSheet.getRange(foundRow, 5).setValue(newVolunteer.type);
      invalidateSheetData(SHEET.SESSION_ROSTER);
    }
    deleteCalendarEvent(activity.calendarId, scheduleSlot, oldName);
    createCalendarEvent(activity.calendarId, scheduleSlot, newVolunteer.name);
    logFormResponse(oldName, oldVolunteerId, activityName, scheduleSlot, "代班異動（管理者）—原志工", adminNote, "取消");
    logFormResponse(newVolunteer.name, newVolunteer.id, activityName, scheduleSlot, "代班異動（管理者）—代班", adminNote, "成功");
    var oldVolunteerObj = getVolunteerById(oldVolunteerId);
    if (oldVolunteerObj) sendVolunteerReplaceOutNotification(oldVolunteerObj, activityName, scheduleSlot, adminEmail, opTime);
    sendVolunteerReplaceInNotification(newVolunteer, activityName, scheduleSlot, adminEmail, opTime);
    sendAdminReplaceNotification({ id: oldVolunteerId, name: oldName }, newVolunteer, activityName, scheduleSlot, adminEmail, opTime);
    return "已成功將「" + oldName + "」換為「" + newVolunteer.name + "」的代班。";
  } catch(e) { return "系統錯誤：" + e.toString(); }
  finally { lock.releaseLock(); }
}

function adminCancelScheduleBatch(items) {
  var results = [];
  for (var i = 0; i < items.length; i++) {
    var r = adminCancelSchedule(items[i].activityName, items[i].scheduleSlot, items[i].volunteerId, items[i].adminEmail || "");
    results.push({ item: items[i], result: r });
  }
  return results;
}

// ================================================
// 特殊日期設定 CRUD
// ================================================

function getAdminSpecialDates() { return getSpecialDateConfigs(); }

function saveAdminSpecialDate(dateStr, activityName, type, timeSlots, maxPerSlot, enabled, note) {
  var sheet = getSheet(SHEET.SPECIAL_DATE);
  if (dateStr.indexOf("-") >= 0) {
    var dp = dateStr.split("-");
    dateStr = dp[0] + "/" + parseInt(dp[1]) + "/" + parseInt(dp[2]);
  }
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === dateStr && data[i][1] === activityName) sheet.deleteRow(i + 1);
  }
  var addedCount = 0;
  if (!enabled || type === "不排班") {
    sheet.appendRow([dateStr, activityName, type, "", parseInt(maxPerSlot) || 0, "否", note || ""]);
    addedCount = 1;
  } else {
    for (var j = 0; j < timeSlots.length; j++) {
      if (!timeSlots[j]) continue;
      sheet.appendRow([dateStr, activityName, type, timeSlots[j], parseInt(maxPerSlot) || 2, "是", note || ""]);
      addedCount++;
    }
  }
  invalidateSheetData(SHEET.SPECIAL_DATE);
  return addedCount > 0 ? "特殊日期設定已儲存（" + dateStr + "，共 " + addedCount + " 個時段）！" : "未新增任何時段，請確認輸入。";
}

function updateAdminSpecialDate(dateStr, activityName, type, timeSlots, maxPerSlot, enabled, note) {
  return saveAdminSpecialDate(dateStr, activityName, type, timeSlots, maxPerSlot, enabled, note);
}

function deleteAdminSpecialDate(dateStr, activityName) {
  var sheet = getSheet(SHEET.SPECIAL_DATE);
  var data  = sheet.getDataRange().getDisplayValues();
  var count = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === dateStr && data[i][1] === activityName) { sheet.deleteRow(i + 1); count++; }
  }
  invalidateSheetData(SHEET.SPECIAL_DATE);
  return count > 0 ? "已刪除特殊日期設定（共 " + count + " 筆）。" : "找不到對應設定。";
}

// ================================================
// 暑假設定 CRUD
// ================================================

function getAdminSummerConfigs() { return getSummerConfigs(); }

function saveAdminSummerConfig(startDate, endDate, activityName, maxPerSlot, slotGroup, note) {
  function normDate(d) {
    if (!d) return "";
    if (d.indexOf("-") >= 0) { var dp = d.split("-"); return dp[0] + "/" + parseInt(dp[1]) + "/" + parseInt(dp[2]); }
    return d;
  }
  var sheet = getSheet(SHEET.SUMMER);
  sheet.appendRow([normDate(startDate), normDate(endDate), activityName, parseInt(maxPerSlot) || 2, slotGroup || "平日2時段", note || ""]);
  invalidateSheetData(SHEET.SUMMER);
  return "寒暑假設定已新增！";
}

function updateAdminSummerConfig(rowIndex, startDate, endDate, activityName, maxPerSlot, slotGroup, note) {
  function normDate(d) {
    if (!d) return "";
    if (d.indexOf("-") >= 0) { var dp = d.split("-"); return dp[0] + "/" + parseInt(dp[1]) + "/" + parseInt(dp[2]); }
    return d;
  }
  var sheet = getSheet(SHEET.SUMMER);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return "找不到對應設定。";
  sheet.getRange(rowIndex, 1, 1, 6).setValues([[normDate(startDate), normDate(endDate), activityName, parseInt(maxPerSlot) || 2, slotGroup || "平日2時段", note || ""]]);
  invalidateSheetData(SHEET.SUMMER);
  return "寒暑假設定已更新！";
}

function deleteAdminSummerConfig(rowIndex) {
  var sheet = getSheet(SHEET.SUMMER);
  if (rowIndex < 1 || rowIndex + 1 > sheet.getLastRow()) return "找不到對應設定。";
  sheet.deleteRow(rowIndex + 1);
  invalidateSheetData(SHEET.SUMMER);
  return "已刪除寒暑假設定。";
}

// ================================================
// 套用特殊日期到排班總覽
// ================================================

function applySpecialDatesToOverview(activityName) {
  var overviewSheet = getSheet(SHEET.OVERVIEW);
  var cacheSheet    = getSheet(SHEET.CACHE);
  var configs       = getSpecialDateConfigs();
  var activityConfigs = configs.filter(function(c) { return c.activity === activityName || c.activity === ""; });
  if (activityConfigs.length === 0) return "目前沒有設定任何特殊日期。";
  var dateGroups = {};
  activityConfigs.forEach(function(c) { if (!dateGroups[c.date]) dateGroups[c.date] = []; dateGroups[c.date].push(c); });
  var addedCount = 0, skippedCount = 0, removedCount = 0;
  Object.keys(dateGroups).forEach(function(dateStr) {
    var dayConfigs = dateGroups[dateStr];
    var isDisabled = dayConfigs.every(function(c) { return !c.enabled; });
    var overviewData = overviewSheet.getDataRange().getDisplayValues();
    var existingRows = [];
    for (var i = 1; i < overviewData.length; i++) {
      if (overviewData[i][0] === activityName && overviewData[i][1] === dateStr) {
        existingRows.push({ rowIndex: i + 1, timeSlot: overviewData[i][2], hasVolunteer: !!(overviewData[i][3] || overviewData[i][5]) });
      }
    }
    if (isDisabled) {
      for (var r = existingRows.length - 1; r >= 0; r--) {
        if (!existingRows[r].hasVolunteer) { overviewSheet.deleteRow(existingRows[r].rowIndex); removedCount++; }
        else { skippedCount++; }
      }
      return;
    }
    var expectedSlots = [];
    dayConfigs.forEach(function(c) { if (c.enabled && c.timeSlot) expectedSlots.push(c.timeSlot); });
    var existingSlotSet = {};
    existingRows.forEach(function(r) { existingSlotSet[r.timeSlot] = true; });
    for (var r2 = existingRows.length - 1; r2 >= 0; r2--) {
      var row = existingRows[r2];
      if (expectedSlots.indexOf(row.timeSlot) < 0) {
        if (!row.hasVolunteer) { overviewSheet.deleteRow(row.rowIndex); removedCount++; }
        else { skippedCount++; }
      }
    }
    expectedSlots.forEach(function(timeSlot) {
      if (!existingSlotSet[timeSlot]) {
        var dp2 = dateStr.split("/");
        var weekdayNames2 = ["星期日","星期一","星期二","星期三","星期四","星期五","星期六"];
        var weekdayStr2   = weekdayNames2[new Date(parseInt(dp2[0]), parseInt(dp2[1])-1, parseInt(dp2[2])).getDay()];
        overviewSheet.appendRow([activityName, dateStr, timeSlot, "", "", "", "", "", "", "", "", "可排班", weekdayStr2]);
        addedCount++;
      }
    });
  });
  try {
    var newOverviewData = overviewSheet.getDataRange().getDisplayValues();
    var cacheData = cacheSheet.getDataRange().getDisplayValues();
    var cacheSet  = {};
    for (var c2 = 1; c2 < cacheData.length; c2++) { cacheSet[cacheData[c2][0]+"|"+cacheData[c2][1]+"|"+cacheData[c2][2]] = true; }
    var newCacheRows = [];
    for (var n = 1; n < newOverviewData.length; n++) {
      if (newOverviewData[n][0] !== activityName) continue;
      var cKey = activityName+"|"+newOverviewData[n][1]+"|"+newOverviewData[n][2];
      if (!cacheSet[cKey]) newCacheRows.push([activityName, newOverviewData[n][1], newOverviewData[n][2], "可排班"]);
    }
    if (newCacheRows.length > 0) cacheSheet.getRange(cacheSheet.getLastRow()+1, 1, newCacheRows.length, 4).setValues(newCacheRows);
  } catch(e) { Logger.log("快取更新失敗：" + e.toString()); }
  invalidateSheetData(SHEET.OVERVIEW);
  var msg = "套用完成！";
  if (addedCount   > 0) msg += "\n✅ 新增 " + addedCount   + " 個時段";
  if (removedCount > 0) msg += "\n🗑 移除 " + removedCount + " 個空白時段";
  if (skippedCount > 0) msg += "\n⚠️ " + skippedCount + " 個時段已有志工排班，未異動";
  return msg;
}

// ================================================
// 缺額彙總（給首頁「缺額通知」折疊框使用）
// ================================================
function getVacancySummary(daysAhead) {
  daysAhead = daysAhead || 30;

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var endDate = new Date(today);
  endDate.setDate(endDate.getDate() + daysAhead);

  var overviewData  = getSheetData(SHEET.OVERVIEW);
  var sessionData   = getSheetData(SHEET.SESSION);
  var sessionRoster = getSheetData(SHEET.SESSION_ROSTER);
  var actData       = getSheetData(SHEET.ACTIVITY);

  var sessionCountMap = {};
  for (var r = 1; r < sessionRoster.length; r++) {
    var rKey = sessionRoster[r][0] + "|" + sessionRoster[r][1];
    sessionCountMap[rKey] = (sessionCountMap[rKey] || 0) + 1;
  }

  var specialCacheByAct = {};
  var monthQCache = buildMonthQuotaCache();
  var allSlots = [];

  // ── 二觀型 ──
  for (var a = 1; a < actData.length; a++) {
    var actName = actData[a][1], actType = actData[a][2];
    if (!actName || actType !== ACTIVITY_TYPE.BINOCULAR) continue;
    if (!specialCacheByAct[actName]) specialCacheByAct[actName] = buildSpecialDateCache(actName);
    var specialCache = specialCacheByAct[actName];

    var current = new Date(today);
    while (current <= endDate) {
      var day = current.getDay();
      if (day !== 1) {
        var dateStr = current.getFullYear() + "/" + (current.getMonth()+1) + "/" + current.getDate();
        var slotDefs = getBinocularSlotDefs(dateStr, actName);
        slotDefs.forEach(function(slot) {
          var timeSlot = slot.start + "-" + slot.end;
          var max = getBinocularMaxCached(dateStr, timeSlot, actName, specialCache, monthQCache);
          if (max <= 0) return;
          var cur = 0;
          for (var i = 1; i < overviewData.length; i++) {
            if (overviewData[i][0] === actName && overviewData[i][1] === dateStr && overviewData[i][2] === timeSlot) {
              cur = countVolunteers(overviewData[i]);
              break;
            }
          }
          var status = cur >= max ? "已滿" : cur > 0 ? "部分排班" : "可排班";
          if (status !== "已滿") {
            allSlots.push({ activity: actName, date: dateStr, slot: timeSlot, current: cur, max: max, shortage: max - cur });
          }
        });
      }
      current.setDate(current.getDate() + 1);
    }
  }

  // ── 場次型 ──
  for (var s = 1; s < sessionData.length; s++) {
    if (!sessionData[s][1]) continue;
    var slotD = new Date(dateToISO(sessionData[s][1]));
    slotD.setHours(0, 0, 0, 0);
    if (slotD < today || slotD > endDate) continue;
    var actN = sessionData[s][0], dateS = sessionData[s][1];
    var tSlot = padTime(sessionData[s][2]) + "-" + padTime(sessionData[s][3]);
    var maxV = (parseInt(sessionData[s][4]) || 0) + (parseInt(sessionData[s][5]) || 0);
    var curV = sessionCountMap[dateS + "|" + tSlot] || 0;
    if (curV < maxV) {
      allSlots.push({ activity: actN, date: dateS, slot: tSlot, current: curV, max: maxV, shortage: maxV - curV });
    }
  }

  allSlots.sort(function(x, y) {
    function pd(str) { var p = str.split("/"); return new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2])); }
    var dx = pd(x.date), dy = pd(y.date);
    return dx - dy !== 0 ? dx - dy : x.slot.split("-")[0].localeCompare(y.slot.split("-")[0]);
  });

  return allSlots;
}

// ================================================
// ★★★ v1.5：志工自助取消 ★★★
// 規則：
//   1. 每人每月 N 次（後台「自助取消月額度」，預設 2），跨所有活動累計
//   2. 以「操作時間」的日曆月計，每月 1 日自動重置
//   3. 值勤日前 N 天內（含當天）不可自助取消（後台「取消截止天數」，預設 7）
//   4. 取消後時段立即釋放、Calendar 事件移除、通知志工與管理者
//   5. 當月登記額度會回補（沿用原本「數現況」的計算方式）
// ================================================

/**
 * 距離班次還有幾天（以日曆天計，不跳過休館日）
 * 回傳 null 表示日期格式無法解析
 */
function daysUntilSlot(dateStr) {
  var d = new Date(dateToISO(dateStr));
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.round((d - today) / 86400000);
}

/**
 * 計算該志工「本月」已使用的自助取消次數
 * 只計 operation 為「取消排班（自助）」的紀錄，管理者取消不算
 */
function countSelfCancelThisMonth(volunteerId) {
  var count = 0;
  try {
    var data = getSheetData(SHEET.FORM);
    var now  = new Date();
    var y = now.getFullYear(), m = now.getMonth() + 1;
    var vid = volunteerId.toString().trim();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2].toString().trim() !== vid) continue;
      // ★ 用關鍵字比對，避免全形括號或文字微調造成漏算
      var op = data[i][5].toString();
      if (op.indexOf("取消") < 0 || op.indexOf("自助") < 0) continue;
      var ym = parseFormTimestampYM(data[i][0]);
      if (!ym) continue;
      if (ym.year === y && ym.month === m) count++;
    }
  } catch(e) {
    Logger.log("countSelfCancelThisMonth 失敗：" + e.toString());
  }
  return count;
}

/**
 * ★ v1.5.1 修正：解析「表單回應1」A欄的時間戳記，取出年與月
 *
 * getDisplayValues() 回傳的是顯示字串，例如「2026/8/3 下午 10:50:23」。
 * 直接用 new Date() 會因為中文的「上午/下午」而解析失敗（Invalid Date），
 * 導致取消次數永遠算成 0、額度形同虛設。
 * 這裡改用正規表示式直接抓開頭的年/月，不受時間格式影響。
 */
function parseFormTimestampYM(raw) {
  if (!raw) return null;
  var s = raw.toString().trim();
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})/);
  if (m) return { year: parseInt(m[1]), month: parseInt(m[2]) };
  // 後備：若儲存格為真正的日期物件
  var d = new Date(s);
  if (!isNaN(d.getTime())) return { year: d.getFullYear(), month: d.getMonth() + 1 };
  return null;
}

/**
 * 取得自助取消的額度狀態（給前端顯示用）
 */
function getSelfCancelInfo(volunteerId) {
  var max  = getSelfCancelQuota();
  var used = countSelfCancelThisMonth(volunteerId);
  var now  = new Date();
  var next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var contact = getSystemConfig("管理者1_Email");
  return {
    used:         used,
    max:          max,
    remaining:    Math.max(0, max - used),
    resetDate:    next.getFullYear() + "/" + (next.getMonth() + 1) + "/1",
    deadlineDays: getCancelDeadlineDays(),
    contactEmail: (contact && contact.trim()) ? contact.trim() : ""
  };
}

/**
 * 額度用完時顯示的提示文字
 */
function buildCancelContactHint() {
  var email = getSystemConfig("管理者1_Email");
  return (email && email.trim())
    ? "如需取消請來信 " + email.trim()
    : "如需取消請聯繫管理人員";
}

/**
 * 志工自助取消排班（主函式）
 * 回傳 { ok: bool, message: string, cancelInfo: {...} }
 */
function selfCancelSchedule(name, volunteerId, activityName, scheduleSlot) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    invalidateSheetData();

    // ── 1. 身分驗證 ──
    var volunteer = validateVolunteer(name, volunteerId);
    if (!volunteer) {
      return { ok: false, message: "取消失敗：姓名與編號不符。" };
    }

    var activity = getActivityConfig(activityName);
    if (!activity) {
      return { ok: false, message: "取消失敗：找不到活動設定。" };
    }

    var slotParts = scheduleSlot.split(" ");
    if (slotParts.length < 2) {
      return { ok: false, message: "取消失敗：時段格式錯誤。" };
    }
    var dateStr = slotParts[0], timeSlot = slotParts[1];

    // ── 2. 截止天數檢查（後端必驗，不可只靠前端）──
    var deadlineDays = getCancelDeadlineDays();
    var daysLeft = daysUntilSlot(dateStr);
    if (daysLeft === null) {
      return { ok: false, message: "取消失敗：日期格式錯誤。" };
    }
    if (daysLeft <= deadlineDays) {
      return {
        ok: false,
        message: "取消失敗：值勤日前 " + deadlineDays + " 天內不可自助取消。\n" + buildCancelContactHint() + "。",
        cancelInfo: getSelfCancelInfo(volunteerId)
      };
    }

    // ── 3. 額度檢查（後端必驗）──
    var info = getSelfCancelInfo(volunteerId);
    if (info.remaining <= 0) {
      return {
        ok: false,
        message: "取消失敗：本月自助取消額度已用完（上限 " + info.max + " 次）。\n" +
                 buildCancelContactHint() + "。\n額度將於 " + info.resetDate + " 重置。",
        cancelInfo: info
      };
    }

    var cancelTime = formatNow();
    var note = "志工自助取消 @ " + cancelTime;

    // ── 4. 釋放時段 ──
    if (activity.type === ACTIVITY_TYPE.BINOCULAR) {
      var overviewSheet = getSheet(SHEET.OVERVIEW);
      var data  = getSheetData(SHEET.OVERVIEW);
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] !== activityName || data[i][1] !== dateStr || data[i][2] !== timeSlot) continue;
        for (var v = 0; v < VOL_COLS.length; v++) {
          if (data[i][VOL_COLS[v].id] == volunteerId) { found = true; break; }
        }
        if (found) break;
      }
      if (!found) {
        return { ok: false, message: "取消失敗：找不到您在此時段的排班紀錄。", cancelInfo: info };
      }
      updateOverviewSheet(overviewSheet, activityName, scheduleSlot, volunteer.name, volunteer.id, "remove");
    } else {
      var rosterSheet = getSheet(SHEET.SESSION_ROSTER);
      var rosterData  = getSheetData(SHEET.SESSION_ROSTER);
      var deleted = false;
      for (var j = rosterData.length - 1; j >= 1; j--) {
        if (rosterData[j][0] === dateStr && rosterData[j][1] === timeSlot && rosterData[j][2] == volunteerId) {
          rosterSheet.deleteRow(j + 1);
          invalidateSheetData(SHEET.SESSION_ROSTER);
          deleted = true;
          break;
        }
      }
      if (!deleted) {
        return { ok: false, message: "取消失敗：找不到您在此場次的報名紀錄。", cancelInfo: info };
      }
      var sessionConfig = getSessionConfig(activityName, dateStr, timeSlot);
      if (sessionConfig) updateSessionOverviewStatus(activityName, dateStr, timeSlot, sessionConfig);
    }

    // ── 5. 移除 Calendar 事件 ──
    deleteCalendarEvent(activity.calendarId, scheduleSlot, volunteer.name);

    // ── 6. 寫入紀錄（額度以此計算，operation 字串不可更動）──
    logFormResponse(volunteer.name, volunteer.id, activityName, scheduleSlot, "取消排班（自助）", note, "取消");

    // ── 7. 通知 ──
    var newInfo = getSelfCancelInfo(volunteerId);
    try {
      if (volunteer.email) sendVolunteerSelfCancelNotification(volunteer, activityName, scheduleSlot, cancelTime, newInfo);
      sendAdminSelfCancelNotification(volunteer, activityName, scheduleSlot, cancelTime);
    } catch(e) {
      Logger.log("自助取消通知發送失敗：" + e.toString());
    }

    return {
      ok: true,
      message: "已成功取消 " + scheduleSlot + " 的排班。\n本月尚可自助取消 " + newInfo.remaining + " 次。",
      cancelInfo: newInfo
    };

  } catch(e) {
    return { ok: false, message: "系統錯誤：" + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 批次自助取消（前端限制最多勾選「剩餘額度」個）
 */
function selfCancelScheduleBatch(name, volunteerId, items) {
  var results = [], successCount = 0, failCount = 0;
  for (var i = 0; i < items.length; i++) {
    var r = selfCancelSchedule(name, volunteerId, items[i].activityName, items[i].scheduleSlot);
    results.push({ item: items[i], ok: r.ok, message: r.message });
    if (r.ok) { successCount++; } else { failCount++; }
  }
  return {
    results: results,
    successCount: successCount,
    failCount: failCount,
    cancelInfo: getSelfCancelInfo(volunteerId)
  };
}
