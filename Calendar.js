// ================================================
// Google Calendar 同步函數
// ================================================

/**
 * 建立日曆活動
 * event.setStatus(FREE) 確保事件不會在 Google Calendar 月檢視顯示為「忙碌」
 */
function createCalendarEvent(calendarId, scheduleSlot, name) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log("找不到日曆：" + calendarId);
      return;
    }
    var parts = scheduleSlot.match(/(\d{4}\/\d{1,2}\/\d{1,2}).*?(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
    if (!parts) {
      Logger.log("時段格式錯誤：" + scheduleSlot);
      return;
    }
    var start = new Date(parts[1] + " " + parts[2]);
    var end   = new Date(parts[1] + " " + parts[3]);
    var event = calendar.createEvent(name, start, end);
    event.setStatus(CalendarApp.EventStatus.FREE);  // 設為「有空」，避免月檢視顯示忙碌
  } catch(e) {
    Logger.log("建立日曆活動失敗：" + e.toString());
  }
}

/**
 * 刪除日曆活動
 */
function deleteCalendarEvent(calendarId, scheduleSlot, name) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return;
    var parts = scheduleSlot.match(/(\d{4}\/\d{1,2}\/\d{1,2}).*?(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
    if (!parts) return;
    var start  = new Date(parts[1] + " " + parts[2]);
    var end    = new Date(parts[1] + " " + parts[3]);
    var events = calendar.getEvents(start, end, {search: name});
    for (var i = 0; i < events.length; i++) {
      events[i].deleteEvent();
    }
  } catch(e) {
    Logger.log("刪除日曆活動失敗：" + e.toString());
  }
}
