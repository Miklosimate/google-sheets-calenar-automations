
const SHEETS = {
  PLAN: 'Plan',
  SCHEDULE: 'Schedule',
  SETTINGS: 'Settings',
  LOGS: 'Logs'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Training Sync')
    .addItem('Test calendar access', 'testCalendarAccess')
    .addItem('Sync plan to calendar', 'syncPlanToCalendar')
    .addSeparator()
    .addItem('Clear logs', 'clearLogs')
    .addItem('Reset sync status', 'resetSyncStatus')
    .addToUi();
}

function testCalendarAccess() {
  const settings = getSettingsMap();
  const cal = CalendarApp.getCalendarById(settings.calendar_id || 'primary');
  if (!cal) {
    appendLog_('ERROR', 'TEST', '', '', 'Calendar not found', 'Check Settings!calendar_id');
    throw new Error('Calendar not found. Check Settings sheet.');
  }
  appendLog_('INFO', 'TEST', '', '', 'Calendar access OK', 'Connected to calendar: ' + cal.getName());
  SpreadsheetApp.getActive().toast('Calendar access OK: ' + cal.getName(), 'Training Sync', 5);
}

function syncPlanToCalendar() {
  const ss = SpreadsheetApp.getActive();
  const planSheet = ss.getSheetByName(SHEETS.PLAN);
  const scheduleSheet = ss.getSheetByName(SHEETS.SCHEDULE);
  if (!planSheet || !scheduleSheet) throw new Error('Missing required sheets.');

  const settings = getSettingsMap();
  const calendarId = settings.calendar_id || 'primary';
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) {
    appendLog_('ERROR', 'SYNC', '', '', 'Calendar not found', 'calendar_id=' + calendarId);
    throw new Error('Calendar not found: ' + calendarId);
  }

  const schedule = getScheduleMap_();
  const values = planSheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = indexMap_(headers);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rowNumber = r + 1;
    try {
      const enabled = toBool_(row[idx.enabled]);
      if (!enabled && toBool_(settings.skip_disabled_rows)) {
        updatePlanStatus_(planSheet, rowNumber, idx, '', '', 'skipped_disabled');
        continue;
      }

      const rawDate = row[idx.date];
      const date = parseDateCell_(rawDate);
      if (!date) {
        appendLog_('ERROR', 'ROW', rowNumber, row[idx.title], 'Invalid date', String(rawDate));
        updatePlanStatus_(planSheet, rowNumber, idx, '', '', 'invalid_date');
        continue;
      }

      const title = String(row[idx.title] || '').trim() || ('Training: ' + String(row[idx.category] || 'Session'));
      const weekday = String(row[idx.weekday] || weekdayName_(date)).trim();
      const useSchedule = toBool_(row[idx.use_schedule]);
      const overrideMode = normalizeMode_(row[idx.event_mode_override]);
      const overrideStart = row[idx.start_time_override];
      const overrideEnd = row[idx.end_time_override];
      const durationMin = Number(row[idx.duration_min] || 0);
      const location = String(row[idx.location] || '');
      const description = buildDescription_(headers, row);
      const scheduleRow = useSchedule ? (schedule[weekday.toLowerCase()] || null) : null;

      if (scheduleRow && !scheduleRow.enabled) {
        appendLog_('INFO', 'ROW', rowNumber, title, 'Skipped: schedule disabled', weekday);
        updatePlanStatus_(planSheet, rowNumber, idx, '', '', 'skipped_schedule_disabled');
        continue;
      }

      let finalMode = overrideMode || (scheduleRow ? scheduleRow.event_mode : '') || normalizeMode_(settings.default_event_mode) || 'ALL_DAY';
      let finalEventId = String(row[idx.event_id] || '').trim();

      if (toBool_(settings.overwrite_existing) && finalEventId) {
        deleteExistingEvent_(finalEventId);
        finalEventId = '';
      }

      let event;
      if (finalMode === 'ALL_DAY') {
        event = cal.createAllDayEvent(title, date, {
          description: description,
          location: location || (scheduleRow ? scheduleRow.location : '')
        });
      } else {
        const startMinutes = parseTimeToMinutes_(overrideStart) ??
          (scheduleRow ? parseTimeToMinutes_(scheduleRow.start_time_raw) : null) ??
          parseTimeToMinutes_(settings.default_start_time) ?? 17 * 60;

        const endMinutesDirect = parseTimeToMinutes_(overrideEnd) ??
          (scheduleRow ? parseTimeToMinutes_(scheduleRow.end_time_raw) : null) ??
          parseTimeToMinutes_(settings.default_end_time);

        let endMinutes = endMinutesDirect;
        if (endMinutes == null) {
          endMinutes = startMinutes + (durationMin > 0 ? durationMin : 60);
        }
        if (endMinutes <= startMinutes) {
          endMinutes = startMinutes + (durationMin > 0 ? durationMin : 60);
        }

        const start = combineDateAndMinutes_(date, startMinutes);
        const end = combineDateAndMinutes_(date, endMinutes);

        event = cal.createEvent(title, start, end, {
          description: description,
          location: location || (scheduleRow ? scheduleRow.location : '')
        });
      }

      const eventId = event.getId();
      updatePlanStatus_(planSheet, rowNumber, idx, eventId, new Date(), 'synced');
      appendLog_('INFO', 'SYNC', rowNumber, title, 'Synced', finalMode);
    } catch (err) {
      appendLog_('ERROR', 'SYNC', rowNumber, row[idx.title], err.message, String(err.stack || ''));
      updatePlanStatus_(planSheet, rowNumber, idx, '', new Date(), 'error');
    }
  }

  SpreadsheetApp.getActive().toast('Sync finished. Check Logs sheet.', 'Training Sync', 5);
}

function resetSyncStatus() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.PLAN);
  const values = sheet.getDataRange().getValues();
  const idx = indexMap_(values[0].map(String));
  for (let r = 1; r < values.length; r++) {
    sheet.getRange(r + 1, idx.event_id + 1).clearContent();
    sheet.getRange(r + 1, idx.last_sync + 1).clearContent();
    sheet.getRange(r + 1, idx.sync_status + 1).setValue('');
  }
  appendLog_('INFO', 'RESET', '', '', 'Sync columns cleared', '');
}

function clearLogs() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.LOGS);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  }
}

function getSettingsMap() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (!key) continue;
    map[key] = values[i][1];
  }
  return map;
}

function getScheduleMap_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SCHEDULE);
  const values = sheet.getDataRange().getValues();
  const map = {};
  const headers = indexMap_(values[0].map(String));
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const weekday = String(row[headers.weekday] || '').trim();
    if (!weekday) continue;
    map[weekday.toLowerCase()] = {
      enabled: toBool_(row[headers.enabled]),
      event_mode: normalizeMode_(row[headers.event_mode]) || 'ALL_DAY',
      start_time_raw: row[headers.start_time],
      end_time_raw: row[headers.end_time],
      location: String(row[headers.location] || ''),
      notes: String(row[headers.notes] || '')
    };
  }
  return map;
}

function buildDescription_(headers, row) {
  const idx = indexMap_(headers);
  return [
    'Category: ' + safe_(row[idx.category]),
    'Details: ' + safe_(row[idx.details]),
    'Duration: ' + safe_(row[idx.duration_min]) + ' min',
    'Distance target: ' + safe_(row[idx.distance_target]),
    'HR target: ' + safe_(row[idx.hr_target]),
    'Week: ' + safe_(row[idx.week])
  ].join('\n');
}

function deleteExistingEvent_(eventId) {
  try {
    const event = CalendarApp.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch (e) {
    // Ignore and recreate
  }
}

function updatePlanStatus_(sheet, rowNumber, idx, eventId, lastSync, status) {
  if (idx.event_id != null) sheet.getRange(rowNumber, idx.event_id + 1).setValue(eventId || '');
  if (idx.last_sync != null && lastSync) sheet.getRange(rowNumber, idx.last_sync + 1).setValue(lastSync);
  if (idx.sync_status != null) sheet.getRange(rowNumber, idx.sync_status + 1).setValue(status || '');
}

function appendLog_(level, action, rowNumber, title, message, details) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.LOGS);
  sheet.appendRow([new Date(), level, action, rowNumber || '', title || '', message || '', details || '']);
}

function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).trim()] = i);
  return map;
}

function normalizeMode_(value) {
  const s = String(value || '').trim().toUpperCase();
  if (s === 'ALL_DAY' || s === 'TIMED') return s;
  return '';
}

function toBool_(value) {
  if (typeof value === 'boolean') return value;
  const s = String(value || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1';
}

function parseDateCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  const s = String(value || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const parsed = new Date(s);
  if (!isNaN(parsed)) return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  return null;
}

function parseTimeToMinutes_(value) {
  if (value === null || value === undefined || value === '') return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value.getHours() * 60 + value.getMinutes();
  }

  if (typeof value === 'number') {
    return Math.round(value * 24 * 60);
  }

  const s = String(value).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return Number(m[1]) * 60 + Number(m[2]);
  return null;
}

function combineDateAndMinutes_(dateOnly, minutes) {
  const d = new Date(dateOnly.getFullYear(), dateOnly.getMonth(), dateOnly.getDate(), 0, 0, 0, 0);
  d.setMinutes(minutes);
  return d;
}

function weekdayName_(date) {
  return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][date.getDay()];
}

function safe_(value) {
  return value === null || value === undefined ? '' : String(value);
}
