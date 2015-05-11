// The script tracks a kinds of calendar entries: 
// * meetings
// * tasks
// * other personal events
//
// Breakfast, lunch and dinner are considered work time to me, if they are
// present on the work calendar. I am using this time to think about work, or
// integrating better with my colleagues. If that's not the case for you, add a
// "personal:" prefix to the calendar entry name, or adjust the code below to
// do the right thing for you.

var whoAmI = "kduleba@google.com";
var labelName = "_calendar guardian";
var holidayCalendarOwner = "Corp Holidays (Switzerland)";

function advanceDate(today, day_index) {
  var date = new Date(today.getTime() + (day_index * 24 * 60 * 60 * 1000));
  
  date.setHours(0, 0, 0, 0);
  return date;
}

function cleanupLabel() {
  var label = GmailApp.getUserLabelByName(labelName);
  var threads = label.getThreads();
  for (var i = 1; i < threads.length; i++) {
    threads[i].moveToTrash();
  }
}

function getEventOwner(event) {
  var calendarId = event.getOriginalCalendarId();
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (calendar === null) return '?';
  return calendar.getName();
}

function isHoliday(event, duration_h) {
  if (duration_h < 4.0) return false;
  if (getEventOwner(event) == holidayCalendarOwner) return true;
  return false;
}

function isTeamSocialFunction(title) {
  var common_tasks = ['social gathering'];

  for (var i = 0; i < common_tasks.length; i++) {
    if (title == common_tasks[i]) return true;
  }
  
  return false;
}

function isTaskEvent(event, title) {
  var common_tasks = ['task:', 'breakfast', 'lunch', 'dinner'];

  for (var i = 0; i < common_tasks.length; i++) {
    if (title == common_tasks[i]) return true;
    if (title.indexOf(common_tasks[i] + ' ') == 0) return true;
    var match = title.lastIndexOf(' ' + common_tasks[i]);
    if (match != -1 && match == title.length - common_tasks[i].length - 1) return true;
  }
  
  return false;
}

function isObviousTaskEvent(event, title) {
  var common_tasks = ['task: email', 'task: email (dns)', 'task: daily goals', 'task: email (dns without asking)', 'task: core work (dns without asking)'];

  for (var i = 0; i < common_tasks.length; i++) {
    if (title == common_tasks[i]) return true;
  }
  
  return false;
}

function isPersonalEvent(title) {
  var common_tasks = ['personal', 'gym', 'gfit volleyball'];
  var common_tasks_prefix = ['traditional massage', 'haircuts ', 'personal: ', 'free stuff pick-up'];
  
  for (var i = 0; i < common_tasks.length; i++) {
    if (title == common_tasks[i]) return true;
  }

  for (var i = 0; i < common_tasks_prefix.length; i++) {
    if (title.indexOf(common_tasks_prefix[i]) == 0) return true;
  }

  return false;
}

function canIgnoreEvent(event, title, status, duration_h) {
  var common_tasks = ['wfh', 'busy', 'room', 'geocoding stand-up', 'busy (dns)', 'in transit', 'leaving the office', 'kduleba team biweekly slot', 'lunch', 'breakfast'];
  var common_tasks_prefix = ['flying ', 'flying:', 'conference: ', 'dns ', 'dns: ', 'room: ', 'block: '];

  if (status == CalendarApp.GuestStatus.NO) return true;
  if (duration_h >= 16.0) return true;
  if (duration_h < 0.1) return true;
  if (event.isAllDayEvent()) return true;

  if (title.length == 0) return true;

  for (var i = 0; i < common_tasks.length; i++) {
    if (title == common_tasks[i]) return true;
  }
  for (var i = 0; i < common_tasks_prefix.length; i++) {
    if (title.indexOf(common_tasks_prefix[i]) == 0) return true;
  }

  if (isHoliday(event, duration_h)) return true;

  return false;
}

function roundTimeDelta(start, end) {
  var x = (end - start) / 3600000.0;
  var h = Math.floor(x);
  var m = (x - h) * 60;
  if (m > 19 && m < 30) {
    return h + 0.5;
  }
  if (m > 49 && m < 51) {
    return h + 1;
  }
  return x;
}

function roundMeetingLength(event, day) {
  var start = event.getStartTime();
  var end = event.getEndTime();
  if (day > start) return 0;
  return roundTimeDelta(start, end);
}

function getMyStatus(event) {
  var status = event.getMyStatus();
  if (status != CalendarApp.GuestStatus.OWNER) return status;
  
  var guests = event.getGuestList(true);
  for (var i = 0; i < guests.length; i++) {
    if (guests[i].getEmail() == whoAmI) {
      return guests[i].getGuestStatus();
    }
  }
  return "?";
}

function getGuestStatus(event) {
  var guests = event.getGuestList(true);
  var declines = 0;
  var accepts = 0;
  var maybes = 0;
  for (var i = 0; i < guests.length; i++) {
    if (guests[i].getEmail() == whoAmI) continue;
    var status = guests[i].getGuestStatus();
    if (status == CalendarApp.GuestStatus.NO) {
      declines += 1;
    } else if (status == CalendarApp.GuestStatus.YES) {
      accepts += 1;
    } else {
      maybes += 1;
    }
  }
  
  if (declines == 0) return "";
  if (declines < 0.2 * (accepts + maybes)) {
    return "";
  }
  
  var result = "";
  if (accepts > 0) result += " YES: " + accepts;
  if (maybes > 0) result += " MAYBE: " + maybes;
  if (declines > 0) result += " <strong>NO: " + declines + "</strong>";
  return result;
}

function formatMeetingName(event, title, duration_h, status) {     
  if (status == "INVITED" || status == "MAYBE") {
    return '<strong>' + title + ' (' + status + ', ' + duration_h.toFixed(2) + ')</strong> ' + getGuestStatus(event);
  }
      
  return title + ' (' + duration_h.toFixed(2) + ') ' + getGuestStatus(event);
}

function getTodayDateForCalendar(calendar) {
  var today = new Date();  // in the script timezone.
  var timezone = calendar.getTimeZone();
  
  // Need to format and then parse back to apply the timezone.
  // No other way to do it!
  var formattedDate = Utilities.formatDate(today, timezone, "MMM dd, yyyy");
  return new Date(Date.parse(formattedDate));
}


function CalendarGuardian() {
  cleanupLabel();

  var calendar = CalendarApp.getDefaultCalendar();
  var start_time = new Date();
  var today = getTodayDateForCalendar(calendar);
  
  var weekly_total = 0;
  var weekly_total_h = 0.0;
  var biweekly_workload = 0.0;
  
  var work_week_days = 0.0;
  var email_lines = [];

  for (var day_index = 0; day_index < 14; day_index++) {
    var day = advanceDate(today, day_index);
    var events = calendar.getEventsForDay(day);
    
    var daily_total_meeting_h = 0.0;
    var events_num = 0;
    var meeting_list = [];
    var task_list = [];
    var personal_time = 0;
    var first_meeting_t = -1;
    var last_meeting_t = -1;

    for (var i = 0; i < events.length; i++) {
      var title = events[i].getTitle().toLowerCase();
      var duration_h = roundMeetingLength(events[i], day);
      var status = getMyStatus(events[i]);
     
      if (canIgnoreEvent(events[i], title, status, duration_h)) continue;

      if (events[i].getEndTime() > last_meeting_t) {
        last_meeting_t = events[i].getEndTime();
      }
      if (first_meeting_t == -1 || first_meeting_t > events[i].getStartTime()) {
        first_meeting_t = events[i].getStartTime();
      }

      if (title == "commute" || title == "airport commute") {
        personal_time += duration_h / 2;
        continue;
      }
      if (isPersonalEvent(title)) {
        personal_time += duration_h;
        continue;
      }
      if (isTeamSocialFunction(title)) {
        personal_time += duration_h / 2;
        continue;
      }

      var formatted_title = formatMeetingName(events[i], events[i].getTitle(), duration_h, status);
      
      if (isTaskEvent(events[i], title)) {
        if (!isObviousTaskEvent(events[i], title)) {
          task_list.push("[TASK] " + formatted_title);
        }
      } else {
        meeting_list.push(formatted_title);
        daily_total_meeting_h += duration_h;
        events_num += 1;
      }
    }

    var total_workload = roundTimeDelta(first_meeting_t, last_meeting_t) - personal_time;
    biweekly_workload += total_workload;
    email_lines.push(day.toDateString() + ' - meeting duration for ' + events_num + ' meetings is ' + daily_total_meeting_h.toFixed(2)
      + " , total workload is " + total_workload.toFixed(2));
    email_lines.push.apply(email_lines, meeting_list);
    email_lines.push.apply(email_lines, task_list);
    email_lines.push('');

    weekly_total += events_num;
    weekly_total_h += daily_total_meeting_h;

    if (events_num > 0 && daily_total_meeting_h > 0) work_week_days += 1.0;
  }

  email_lines.push('<br><br>' + 'Bi-weekly total: ' + weekly_total + ' meetings with duration ' + weekly_total_h.toFixed(2) +
    ', per-day average ' + (weekly_total_h / work_week_days).toFixed(2));
  email_lines.push('<br>Bi-weekly workload: ' + biweekly_workload);

  email_lines.push('');
  email_lines.push('Generation time: ' + Math.floor(((new Date()) - start_time) / 1000.0) + ' s');
  
  GmailApp.sendEmail(whoAmI, "calendar guardian", "", { htmlBody: email_lines.join("<br>") });

  cleanupLabel();
}
