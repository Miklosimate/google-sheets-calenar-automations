
/**
 * Google Apps Script: sync the Plan sheet into a Google Calendar.
 * Required sheets: Plan, Settings, Logs
 */

function setupMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Training Sync')
    .addItem('Test calendar access', 'testCalendarAccess')
    .addItem('Sync plan to calendar', 'syncPlanToCalendar')
    .addItem('Clear logs', 'clearLogs')
    .addToUi();
}

function onOpen() {
  setupMenu();
}

function getSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Settings');
  const values = sh.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const val = values[i][1];
    if (key) out[key] = val;
  }
  return out;
}

function getCalendar_() {
  const settings = getSettings_();
  const id = String(settings.calendar_id || 'primary').trim();
  const cal = id === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(id);
  if (!cal) throw new Error('Calendar not found for id: ' + id);
  return cal;
}

function ensureLogsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Logs');
  if (!sh) {
    sh = ss.insertSheet('Logs');
    sh.appendRow(['Timestamp','Level','Row','Date','Title','Message','EventId']);
  }
  return sh;
}

function log_(level, rowNumber, dateValue, title, message, eventId) {
  const sh = ensureLogsSheet_();
  sh.appendRow([new Date(), level, rowNumber || '', dateValue || '', title || '', message || '', eventId || '']);
}

function clearLogs() {
  const sh = ensureLogsSheet_();
  sh.clearContents();
  sh.appendRow(['Timestamp','Level','Row','Date','Title','Message','EventId']);
  log_('INFO', '', '', '', 'Logs cleared', '');
}

function testCalendarAccess() {
  try {
    const cal = getCalendar_();
    log_('INFO', '', '', '', 'Calendar access OK: ' + cal.getName(), '');
    SpreadsheetApp.getActive().toast('Calendar access OK: ' + cal.getName());
  } catch (err) {
    log_('ERROR', '', '', '', err.message, '');
    SpreadsheetApp.getUi().alert('Calendar access failed: ' + err.message);
    throw err;
  }
}

function syncPlanToCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings_();
  const tz = String(settings.timezone || Session.getScriptTimeZone() || 'Europe/Budapest');
  const planSheetName = String(settings.plan_sheet_name || 'Plan');
  const createRest = String(settings.create_rest_events || 'FALSE').toUpperCase() === 'TRUE';
  const onlyFuture = String(settings.only_sync_future_rows || 'TRUE').toUpperCase() === 'TRUE';
  const titlePrefix = String(settings.title_prefix || '');
  const sh = ss.getSheetByName(planSheetName);
  if (!sh) throw new Error('Plan sheet not found: ' + planSheetName);

  const cal = getCalendar_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const today = new Date();
  today.setHours(0,0,0,0);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rowNumber = r + 1;
    try {
      const syncFlag = row[idx['SyncToCalendar']];
      const type = String(row[idx['SessionType']] || '');
      if (!syncFlag) {
        sh.getRange(rowNumber, idx['LastSyncStatus'] + 1).setValue('Skipped: SyncToCalendar=FALSE');
        continue;
      }
      if (type === 'Rest' && !createRest) {
        sh.getRange(rowNumber, idx['LastSyncStatus'] + 1).setValue('Skipped: Rest event');
        continue;
      }

      const startDate = row[idx['StartDate']];
      const startTime = row[idx['StartTime']];
      const endDate = row[idx['EndDate']];
      const endTime = row[idx['EndTime']];
      const title = titlePrefix + String(row[idx['Title']] || '').trim();
      const notes = String(row[idx['Notes']] || '').trim();
      const hrMin = row[idx['HRMin']];
      const hrMax = row[idx['HRMax']];
      const dist = row[idx['DistanceKm']];
      const priority = String(row[idx['Priority']] || '').trim();

      if (!startDate) {
        sh.getRange(rowNumber, idx['LastSyncStatus'] + 1).setValue('Error: missing StartDate');
        log_('ERROR', rowNumber, '', title, 'Missing StartDate', '');
        continue;
      }

      const sDate = new Date(startDate);
      sDate.setHours(0,0,0,0);
      if (onlyFuture && sDate < today) {
        sh.getRange(rowNumber, idx['LastSyncStatus'] + 1).setValue('Skipped: past date');
        continue;
      }

      let startDt, endDt, isAllDay = false;
      if (!startTime || !endTime || type === 'Rest') {
        startDt = new Date(sDate);
        endDt = new Date(sDate);
        isAllDay = true;
      } else {
        startDt = combineDateAndTime_(startDate, startTime, tz);
        endDt = combineDateAndTime_(endDate || startDate, endTime, tz);
      }

      const descParts = [
        'Type: ' + type,
        dist ? 'Distance target: ' + dist + ' km' : '',
        (hrMin || hrMax) ? ('HR target: ' + (hrMin || '') + '-' + (hrMax || '') + ' bpm') : '',
        priority ? 'Priority: ' + priority : '',
        notes ? 'Notes: ' + notes : ''
      ].filter(Boolean);
      const description = descParts.join('\n');

      let eventId = String(row[idx['EventId']] || '').trim();
      let event = null;
      if (eventId) {
        try {
          event = cal.getEventById(eventId);
        } catch (e) {
          event = null;
        }
      }

      if (event) {
        if (isAllDay) {
          event.setTitle(title);
          event.setDescription(description);
        } else {
          event.setTitle(title);
          event.setTime(startDt, endDt);
          event.setDescription(description);
        }
        sh.getRange(rowNumber, idx['LastSyncStatus'] + 1).setValue('Updated');
        sh.getRange(rowNumber, idx['LastSyncAt'] + 1).setValue(new Date());
        log_('INFO', rowNumber, sDate, title, 'Updated existing event', event.getId());
      } else {
        if (isAllDay) {
          event = cal.createAllDayEvent(title, sDate, {description: description});
        } else {
          event = cal.createEvent(title, startDt, endDt, {description: description});
        }
        sh.getRange(rowNumber, idx['EventId'] + 1).setValue(event.getId());
        sh.getRange(rowNumber, idx['LastSyncStatus'] + 1).setValue('Created');
        sh.getRange(rowNumber, idx['LastSyncAt'] + 1).setValue(new Date());
        log_('INFO', rowNumber, sDate, title, 'Created new event', event.getId());
      }
    } catch (err) {
      sh.getRange(rowNumber, idx['LastSyncStatus'] + 1).setValue('Error');
      sh.getRange(rowNumber, idx['LastSyncAt'] + 1).setValue(new Date());
      log_('ERROR', rowNumber, row[idx['Date']], row[idx['Title']], err.message, row[idx['EventId']] || '');
    }
  }

  SpreadsheetApp.getActive().toast('Sync completed. Check Logs sheet.');
}

function combineDateAndTime_(dateValue, timeValue, tz) {
  const d = new Date(dateValue);
  let hh = 0, mm = 0;
  if (Object.prototype.toString.call(timeValue) === '[object Date]') {
    hh = timeValue.getHours();
    mm = timeValue.getMinutes();
  } else {
    const parts = String(timeValue).split(':');
    hh = parseInt(parts[0], 10);
    mm = parseInt(parts[1], 10);
  }
  d.setHours(hh, mm, 0, 0);
  return d;
}

