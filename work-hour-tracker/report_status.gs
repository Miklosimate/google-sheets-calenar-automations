function createWeeklyProgressEvent() {
  var calendarId = 'calendar_id@group.calendar.google.com'; // Replace with your Calendar ID
  var calendar = CalendarApp.getCalendarById(calendarId);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Read Year, Month, and Target Hours from "Control" sheet
  var controlSheet = spreadsheet.getSheetByName("Control");
  if (!controlSheet) {
    SpreadsheetApp.getUi().alert("Control sheet not found. Please create one with Year in A1, Month in B1, and Target Hours in E1.");
    return;
  }

  var year = parseInt(controlSheet.getRange("A1").getValue(), 10);
  var month = parseInt(controlSheet.getRange("B1").getValue(), 10) - 1; // Convert to zero-based index (Jan = 0, Dec = 11)
  var targetHours = parseFloat(controlSheet.getRange("E1").getValue()); // Read target hours

  // Validate inputs
  if (isNaN(year) || isNaN(month) || month < 0 || month > 11 || isNaN(targetHours) || targetHours < 0) {
    SpreadsheetApp.getUi().alert("Invalid values in Control sheet. Please enter a valid Year (A1), Month (B1), and Target Hours in E1.");
    return;
  }

  var firstDay = new Date(year, month, 1);
  var lastDay = new Date(year, month + 1, 0); // Last day of the selected month

  // Expand to capture all weeks that have at least one day in the selected month
  var firstWeekStart = new Date(firstDay);
  firstWeekStart.setDate(firstDay.getDate() - firstDay.getDay()); // Move back to Sunday

  var lastWeekEnd = new Date(lastDay);
  lastWeekEnd.setDate(lastDay.getDate() + (6 - lastDay.getDay())); // Move forward to Saturday

  // Fetch events from the calendar (for full weeks)
  var events = calendar.getEvents(firstWeekStart, lastWeekEnd);
  
  var workedMinutesPerWeek = {};
  var aiMinutesPerWeek = {};

  // Process events and sum hours per week
  events.forEach(event => {
    var title = event.getTitle().toLowerCase();
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();
    var durationMinutes = (endTime - startTime) / (1000 * 60);

    var weekNumber = getWeekNumber(startTime);
    
    if (!workedMinutesPerWeek[weekNumber]) {
      workedMinutesPerWeek[weekNumber] = 0;
    }
    if (!aiMinutesPerWeek[weekNumber]) {
      aiMinutesPerWeek[weekNumber] = 0;
    }

    if (title.includes("munka")) { // Matches "munka", "Munka (43)", etc.
      workedMinutesPerWeek[weekNumber] += durationMinutes;
    }

    if (title.includes("ai")) { // Matches "AI", "AI Project", etc.
      aiMinutesPerWeek[weekNumber] += durationMinutes;
    }
  });

  // Fetch existing progress events
  var progressEvents = getExistingProgressEvents(calendar, firstWeekStart, lastWeekEnd);

  // Update progress events for every Friday
  for (var week in workedMinutesPerWeek) {
    var workedMinutes = workedMinutesPerWeek[week];
    var aiMinutes = aiMinutesPerWeek[week] || 0; // Default to 0 if no AI work
    var workedTimeHHMM = formatMinutesToHHMM(workedMinutes);
    var targetTimeHHMM = formatMinutesToHHMM(targetHours * 60);
    var aiTimeHHMM = formatMinutesToHHMM(aiMinutes);

    var fridayDate = getFridayOfWeek(year, month, parseInt(week, 10));
    if (!fridayDate) continue;

    var newEventTitle ="W:" + workedTimeHHMM + "/"  + "A:" + aiTimeHHMM + " [PROGRESS]";
    
    console.log("Updating Event: " + newEventTitle + " on " + fridayDate.toDateString());

    var existingEvent = progressEvents[fridayDate.toDateString()];
    if (existingEvent) {
      existingEvent.setTitle(newEventTitle);
    } else {
      var startTime = new Date(fridayDate);
      startTime.setHours(17, 0, 0); // 5:00 PM
      var endTime = new Date(fridayDate);
      endTime.setHours(18, 0, 0); // 6:00 PM

      calendar.createEvent(newEventTitle, startTime, endTime);
    }
  }
}

// 🔹 **Fetch existing progress events for modification**
function getExistingProgressEvents(calendar, firstWeekStart, lastWeekEnd) {
  var events = calendar.getEvents(firstWeekStart, lastWeekEnd);
  var progressEvents = {};

  events.forEach(event => {
    if (event.getTitle().includes("[PROGRESS]")) {
      var eventDate = event.getStartTime().toDateString();
      progressEvents[eventDate] = event;
    }
  });

  return progressEvents;
}

// 🔹 **Gets the week number of a given date**
function getWeekNumber(date) {
  var firstDayOfYear = new Date(date.getFullYear(), 0, 1);
  var pastDaysOfYear = (date - firstDayOfYear) / (1000 * 60 * 60 * 24);
  return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
}

// 🔹 **Finds the Friday of a given week within the year**
function getFridayOfWeek(year, month, weekNumber) {
    var firstDay = new Date(year, 0, 1); // First day of the year
    var firstFriday = new Date(firstDay);

    while (firstFriday.getDay() !== 5) {
        firstFriday.setDate(firstFriday.getDate() + 1); // Move to first Friday of the year
    }

    var targetFriday = new Date(firstFriday);
    targetFriday.setDate(firstFriday.getDate() + (weekNumber - 1) * 7); // Find correct Friday

    return targetFriday;
}

// 🔹 **Formats minutes to HH:MM**
function formatMinutesToHHMM(minutes) {
  var hh = Math.floor(minutes / 60);
  var mm = minutes % 60;
  return hh.toString().padStart(2, '0') + ":" + mm.toString().padStart(2, '0');
}