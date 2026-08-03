// ================================================
// 系統全域設定
// ================================================

var SHEET = {
  ROSTER:         "志工名單",
  ACTIVITY:       "活動設定",
  ANNOUNCE:       "公告設定",
  SESSION:        "場次設定",
  OVERVIEW:       "排班總覽",
  CACHE:          "可用時段簡表",
  FORM:           "表單回應1",
  SYSTEM:         "系統設定",
  SESSION_ROSTER: "親子營排班",
  RECENT:         "最新活動",
  SPECIAL_DATE:   "特殊日期設定",
  SUMMER:         "暑假設定",
  MONTH_QUOTA:    "月份人數設定"   // ← 新增這行
};

var ACTIVITY_TYPE = {
  BINOCULAR: "二觀型",
  SESSION:   "場次型"
};

var VOLUNTEER_TYPE = {
  NEW:     "新志工",
  GENERAL: "一般志工"
};

var CERT = {
  PASS: "通過",
  FAIL: "未通過"
};

var BINOCULAR_SLOTS = {
  WEEKDAY: [
    { label: "上午", start: "10:00", end: "12:00" },
    { label: "下午", start: "14:00", end: "16:00" }
  ],
  SATURDAY: [
    { label: "上午", start: "10:00", end: "12:00" },
    { label: "下午", start: "14:00", end: "16:00" },
    { label: "晚上", start: "19:00", end: "21:00" }
  ],
  SUNDAY: [
    { label: "上午", start: "10:00", end: "12:00" },
    { label: "下午", start: "14:00", end: "16:00" }
  ]
};

var BINOCULAR_MAX = {
  WEEKDAY: 1,
  WEEKEND: 2
};

var PRIORITY_QUOTA = {
  CERTIFIED:     4,
  NOT_CERTIFIED: 2
};

var WORK_DAYS = [0, 2, 3, 4, 5, 6];

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(sheetName) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("找不到工作表：" + sheetName);
  return sheet;
}

function getSystemConfig(key) {
  var data = getSheet(SHEET.SYSTEM).getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function getAdminEmail() {
  return getSystemConfig("管理者Email");
}