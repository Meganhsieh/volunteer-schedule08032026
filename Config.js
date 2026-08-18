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
  var data = getSheetData(SHEET.SYSTEM);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

// ================================================
// 單次執行內的讀表快取
// GAS 每次請求都是獨立執行環境,快取只存活於本次請求,
// 請求結束即消失,不會有跨使用者的髒資料問題。
// ================================================

var _sheetDataCache = {};

/**
 * 取得整張工作表的 DisplayValues(本次執行內只讀一次)
 */
function getSheetData(sheetName) {
  if (!_sheetDataCache.hasOwnProperty(sheetName)) {
    _sheetDataCache[sheetName] = getSheet(sheetName).getDataRange().getDisplayValues();
  }
  return _sheetDataCache[sheetName];
}

/**
 * 寫入工作表後呼叫,讓後續讀取拿到新資料
 * 不帶參數 = 清除全部
 */
function invalidateSheetData(sheetName) {
  if (sheetName) {
    delete _sheetDataCache[sheetName];
  } else {
    _sheetDataCache = {};
  }
}
// ================================================
// 單次執行內的讀表快取
// ================================================
var _sheetDataCache = {};

function getSheetData(sheetName) {
  if (!_sheetDataCache.hasOwnProperty(sheetName)) {
    _sheetDataCache[sheetName] = getSheet(sheetName).getDataRange().getDisplayValues();
  }
  return _sheetDataCache[sheetName];
}

function invalidateSheetData(sheetName) {
  if (sheetName) {
    delete _sheetDataCache[sheetName];
  } else {
    _sheetDataCache = {};
  }
}
