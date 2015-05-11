function emailString(msg) {
  GmailApp.sendEmail("kduleba@google.com", "test", msg);
}

function onOpen() {  
  var menu = [ 
    {name: "Update from calendar", functionName: "updateSchedule"},
  ];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Run scripts", menu);
}

function getOneOnOne(event) {
  var status = event.getMyStatus();
  if (status == CalendarApp.GuestStatus.NO) return;
  if (!event.guestsCanSeeGuests()) return;

  var guests = event.getGuestList(true);
  
  var result;
  for (var i = 0; i < guests.length; i++) {
    var email = guests[i].getEmail();
    if (email == "kduleba@google.com") continue;
    if (email.indexOf("resource.calendar.google.com") != -1) continue;
    if (email.indexOf("group.calendar.google.com") != -1) continue;
    
    if (result) return;
    result = email;
  }
  
  if (result) return result.substr(0, result.indexOf("@"));
}

function getCalendarEntries(calendar, now, day, result) {
  var events = calendar.getEventsForDay(day);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getStartTime() <= now) continue;
    
    var guest = getOneOnOne(events[i]);
    if (guest) {
      if (!(guest in result)) {
        result[guest] = events[i].getStartTime();
      }
    }
  }
  return result;
}

function formatDate(x, timezone) {
  return Utilities.formatDate(x, timezone, "yyyy-MM-dd");
}

function updateProgressMeter(sheet, day_index, max_index) {
  var range = sheet.getRange(1, 1, 1, 10);
  var cell = range.getCell(1, 8);
  cell.setValue((day_index + 1) + " / " + max_index);
}


function getTodayDateForCalendar(calendar) {
  var today = new Date();  // in the script timezone.
  var timezone = calendar.getTimeZone();
  
  // Need to format and then parse back to apply the timezone.
  // No other way to do it!
  var formattedDate = Utilities.formatDate(today, timezone, "MMM dd, yyyy");
  return new Date(Date.parse(formattedDate));
}

function updateSchedule() {
  var calendar = CalendarApp.getDefaultCalendar();
  var timezone = calendar.getTimeZone();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("chat syncer");
  var lastRow = ss.getLastRow();
  var today = getTodayDateForCalendar(calendar);
  var now = new Date();
  var oneonones = {};
  var max_days = 60;

  for (var day_index = 0; day_index < max_days; day_index++) {
    updateProgressMeter(sheet, day_index, max_days);
    var day = new Date(today.getTime() + (day_index * 24 * 60 * 60 * 1000));
    oneonones = getCalendarEntries(calendar, now, day, oneonones);
  
    for (var row = sheet.getFrozenRows() + 1; row <= lastRow; row++) {
      var range = sheet.getRange(row, 1, 1, 4);
      var email = range.getValues()[0][0];
      if (email && (email in oneonones)) {
        var date = formatDate(oneonones[email], timezone);

        nextMeetingCell = range.getCell(1, 4);
        nextMeetingCell.setNumberFormat("yyyy-MM-dd");
        nextMeetingCell.setValue(date);
      }
    }
  }
}
