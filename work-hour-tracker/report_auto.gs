function createSheetAndFetchEvents() {
  var calendarId = 'YOUR_WORK_CALENDAR_ID@group.calendar.google.com'; // Replace with your Calendar ID
  var calendar = CalendarApp.getCalendarById(calendarId);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Read Year, Month, and Hourly Rate from "Control" sheet (A1 = Year, B1 = Month, C1 = Hourly Rate)
  var controlSheet = spreadsheet.getSheetByName("Control");
  if (!controlSheet) {
    SpreadsheetApp.getUi().alert("Control sheet not found. Please create one with Year in A1, Month in B1, and Hourly Rate in C1.");
    return;
  }
  
  var year = parseInt(controlSheet.getRange("A1").getValue(), 10);
  var month = parseInt(controlSheet.getRange("B1").getValue(), 10); // User inputs 1-12
  var hourlyRate = parseFloat(controlSheet.getRange("C1").getValue()); // Read Hourly Rate

  // Validate inputs
  if (isNaN(year) || isNaN(month) || month < 1 || month > 12 || isNaN(hourlyRate) || hourlyRate < 0) {
    SpreadsheetApp.getUi().alert("Invalid values in Control sheet. Please enter a valid Year (A1), Month (B1), and Hourly Rate (C1).");
    return;
  }

  month = month - 1; // Convert to zero-based index (Jan = 0, Dec = 11)
  var monthName = new Date(year, month, 1).toLocaleString('default', { month: 'short' });
  var sheetName = year + "_" + monthName;
  
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // If sheet exists, delete it first
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
  
  // Create new sheet
  sheet = spreadsheet.insertSheet(sheetName);
  
  // Add headers
  var headers = ["Date", "Start Time", "End Time", "Duration (HH:MM)", "Paid Time (HH:MM)", "Weekly Total (HH:MM)", "Event Title"];
  sheet.appendRow(headers);
  
  // Get first and last day of the selected month
  var firstDay = new Date(year, month, 1);
  var lastDay = new Date(year, month + 1, 0); // Last day of the month
  lastDay.setHours(23, 59, 59, 999); // Ensure last day's events are included

  // Fetch events (Inclusive of last day)
  var events = calendar.getEvents(firstDay, lastDay);
  
  // Filter events where the title CONTAINS "munka" (case-insensitive)
  var filteredEvents = events.filter(event => event.getTitle().toLowerCase().includes("munka"));

  // Organize events by date
  var eventMap = {};
  filteredEvents.forEach(event => {
    var dateKey = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (!eventMap[dateKey]) {
      eventMap[dateKey] = [];
    }
    eventMap[dateKey].push(event);
  });

  var totalMinutes = 0;
  var totalPaidMinutes = 0;
  var data = []; // Store data before inserting into the sheet (faster performance)

  var weekMinutes = 0; // Weekly total
  var weekEndDate = null;

  for (var day = 1; day <= lastDay.getDate(); day++) {
    var currentDate = new Date(year, month, day);
    var dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var dayOfWeek = currentDate.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

    weekEndDate = dateString; // Track last day of the week

    if (eventMap[dateString]) {
      eventMap[dateString].forEach(event => {
        var startTime = event.getStartTime();
        var endTime = event.getEndTime();
        var durationMinutes = (endTime - startTime) / (1000 * 60);
        var paidMinutes = durationMinutes;

        // Apply break deductions
        if (durationMinutes > 360 && durationMinutes <= 540) {
          paidMinutes -= 20;
        } else if (durationMinutes > 540) {
          paidMinutes -= 45;
        }
        paidMinutes = Math.max(0, paidMinutes);

        totalMinutes += durationMinutes;
        totalPaidMinutes += paidMinutes;
        weekMinutes += durationMinutes;

        var durationHHMM = formatMinutesToHHMM(durationMinutes);
        var paidTimeHHMM = formatMinutesToHHMM(paidMinutes);

        data.push([
          dateString,
          Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm"),
          Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm"),
          durationHHMM,
          paidTimeHHMM,
          "", // Placeholder for weekly total, filled at the end of the week
          event.getTitle()
        ]);
      });
    } else {
      data.push([dateString, "", "", "", "", "", ""]);
    }

    // If it's Sunday (last day of the week) or the last day of the month, add a weekly total
    if (dayOfWeek === 0 || day === lastDay.getDate()) {
      var weeklyTotal = formatMinutesToHHMM(weekMinutes);
      data[data.length - 1][5] = weeklyTotal; // Column index 5 for "Weekly Total (HH:MM)"
      weekMinutes = 0; // Reset weekly total
    }
  }

  // Insert all data at once for better performance
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }

  // Final calculations
  var totalPaidHours = totalPaidMinutes / 60;
  var totalEarnings = (totalPaidHours * hourlyRate) - 2500;
  var totalFormatted = formatMinutesToHHMM(totalMinutes);
  var totalPaidFormatted = formatMinutesToHHMM(totalPaidMinutes);

  // Add summary
  if (totalMinutes > 0) {
    sheet.appendRow(["", "", "", "", "", "", ""]);
    sheet.appendRow(["Total Hours Worked", "", "", totalFormatted, totalPaidFormatted, "", ""]);
    sheet.appendRow(["Total Pay (HUF)", "", "", "", totalEarnings.toFixed(2), "", ""]);
  }

  // Auto-resize columns for readability
  sheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert("Sheet '" + sheetName + "' created successfully.\nTotal Hours: " + totalFormatted + "\nTotal Paid Time: " + totalPaidFormatted + "\nTotal Pay: HUF " + totalEarnings.toFixed(2));
}

// Function to convert minutes to HH:MM format
function formatMinutesToHHMM(minutes) {
  var hh = Math.floor(minutes / 60);
  var mm = minutes % 60;
  return hh.toString().padStart(2, '0') + ":" + mm.toString().padStart(2, '0');
}
