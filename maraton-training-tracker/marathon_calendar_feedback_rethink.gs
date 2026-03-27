
/**
 * Simplified Marathon Plan <-> Google Calendar sync
 *
 * Calendar event text is intentionally minimal.
 * Edit only the lines inside [FEEDBACK] ... [/FEEDBACK].
 *
 * Examples that work:
 * [FEEDBACK]
 * status=done
 * km=12
 * duration_min=75
 * avg_hr=138
 * max_hr=152
 * avg_pace=5:42
 * rpe=6
 * notes=felt smooth
 * [/FEEDBACK]
 *
 * Skipped examples that work:
 * [FEEDBACK]
 * skipped
 * [/FEEDBACK]
 *
 * [FEEDBACK]
 * status=skipped
 * [/FEEDBACK]
 *
 * [FEEDBACK]
 * status: skipped
 * [/FEEDBACK]
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Training Feedback Sync')
    .addItem('Sync Plan -> Calendar', 'syncPlanToCalendar')
    .addItem('Sync Feedback <- Calendar', 'syncFeedbackFromCalendar')
    .addItem('Sync All', 'syncAll')
    .addSeparator()
    .addItem('Test Calendar Access', 'testCalendarAccess')
    .addItem('Reset Sync Status', 'resetSyncStatus')
    .addItem('Delete Prefixed Events', 'deletePrefixedEvents')
    .addToUi();
}

function getSheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function getSettings_() {
  var sh = getSheet_('Settings');
  var values = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] !== '') out[String(values[i][0]).trim()] = values[i][1];
  }
  return out;
}

function getCalendar_() {
  var settings = getSettings_();
  var calendarId = settings.calendar_id || 'primary';
  var cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error('Calendar not found for calendar_id=' + calendarId);
  return cal;
}

function logAction_(level, fnName, rowNumber, message) {
  var sh = getSheet_('Logs');
  sh.appendRow([new Date(), level, fnName, rowNumber || '', message || '']);
}

function getHeadersMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) map[String(headers[i]).trim()] = i;
  return map;
}

function rowToObject_(row, headers) {
  var obj = {};
  for (var k in headers) obj[k] = row[headers[k]];
  return obj;
}

function toBoolean_(v) {
  if (v === true || v === false) return v;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1';
}

function formatDate_(d) {
  if (!(d instanceof Date)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function cleanLeadingStatus_(title) {
  return String(title || '').replace(/^[\s✅❌⚠🟦🟩🟥]+/, '');
}

function isBlankish_(value) {
  var s = String(value || '').replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\u00A0/g, ' ').trim().toLowerCase();
  return s === '' || s === '-' || s === '...' || s === '…';
}

function normalizeActualStatus_(status) {
  var s = String(status || '').replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\u00A0/g, ' ').trim().toLowerCase();
  if (isBlankish_(s)) return '';
  if (s === 'done' || s === 'completed' || s === 'complete') return 'done';
  if (s === 'skip' || s === 'skipped') return 'skipped';
  if (s === 'modified' || s === 'change' || s === 'changed') return 'modified';
  if (s === 'planned' || s === 'plan') return 'planned';
  return s;
}

function canonicalFeedbackKey_(key) {
  var k = String(key || '')
    .toLowerCase()
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
    .replace(/["“”'`]/g, '')
    .replace(/\s+/g, '')
    .trim();

  var map = {
    status: 'status',
    km: 'actual_km',
    distance_km: 'actual_km',
    distancekm: 'actual_km',
    actual_km: 'actual_km',
    actualkm: 'actual_km',
    duration_min: 'actual_duration_min',
    durationmin: 'actual_duration_min',
    duration: 'actual_duration_min',
    minutes: 'actual_duration_min',
    actual_duration_min: 'actual_duration_min',
    avg_hr: 'actual_avg_hr',
    avghr: 'actual_avg_hr',
    actual_avg_hr: 'actual_avg_hr',
    max_hr: 'actual_max_hr',
    maxhr: 'actual_max_hr',
    actual_max_hr: 'actual_max_hr',
    avg_pace: 'actual_avg_pace',
    avgpace: 'actual_avg_pace',
    pace: 'actual_avg_pace',
    actual_avg_pace: 'actual_avg_pace',
    rpe: 'actual_rpe',
    actual_rpe: 'actual_rpe',
    notes: 'actual_notes',
    note: 'actual_notes',
    actual_notes: 'actual_notes'
  };

  return map[k] || '';
}

function skippedPayload_() {
  return {
    status: 'skipped',
    actual_km: '0',
    actual_duration_min: '0',
    actual_avg_hr: '0',
    actual_max_hr: '0',
    actual_avg_pace: '0',
    actual_rpe: '0',
    actual_notes: 'skipped'
  };
}

function hasMeaningfulActuals_(actuals) {
  if (!actuals) return false;
  var status = normalizeActualStatus_(actuals.status || '');
  if (status && status !== 'planned') return true;

  var keys = ['actual_km', 'actual_duration_min', 'actual_avg_hr', 'actual_max_hr', 'actual_avg_pace', 'actual_rpe', 'actual_notes'];
  for (var i = 0; i < keys.length; i++) {
    if (!isBlankish_(actuals[keys[i]])) return true;
  }
  return false;
}

function buildFeedbackSummary_(rowObj) {
  var out = [];
  if (!isBlankish_(rowObj.status)) out.push('Status: ' + rowObj.status);

  if (!isBlankish_(rowObj.actual_km)) {
    var deltaKm = '';
    if (!isBlankish_(rowObj.planned_km)) {
      var dk = Number(rowObj.actual_km) - Number(rowObj.planned_km);
      if (!isNaN(dk)) deltaKm = ' (' + (dk >= 0 ? '+' : '') + dk.toFixed(1) + ' vs plan)';
    }
    out.push('KM: ' + rowObj.actual_km + deltaKm);
  }

  if (!isBlankish_(rowObj.actual_duration_min)) {
    var deltaMin = '';
    if (!isBlankish_(rowObj.planned_duration_min)) {
      var dd = Number(rowObj.actual_duration_min) - Number(rowObj.planned_duration_min);
      if (!isNaN(dd)) deltaMin = ' (' + (dd >= 0 ? '+' : '') + dd + ' min vs plan)';
    }
    out.push('Duration: ' + rowObj.actual_duration_min + ' min' + deltaMin);
  }

  if (!isBlankish_(rowObj.actual_avg_hr)) out.push('Avg HR: ' + rowObj.actual_avg_hr);
  if (!isBlankish_(rowObj.actual_avg_pace)) out.push('Avg pace: ' + rowObj.actual_avg_pace);
  if (!isBlankish_(rowObj.actual_rpe)) out.push('RPE: ' + rowObj.actual_rpe);
  if (!isBlankish_(rowObj.actual_notes)) out.push('Notes: ' + rowObj.actual_notes);

  return out.join(' | ');
}

function buildScheduleMap_() {
  var sh = getSheet_('Schedule');
  var values = sh.getDataRange().getValues();
  var headers = getHeadersMap_(sh);
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var row = rowToObject_(values[i], headers);
    if (!row.weekday) continue;
    map[String(row.weekday)] = row;
  }
  return map;
}

function chooseEventMode_(rowObj, scheduleMap, settings) {
  var override = String(rowObj.event_mode_override || '').trim();
  if (override) return override.toUpperCase();

  if (String(settings.use_schedule).toUpperCase() === 'TRUE' && scheduleMap[rowObj.weekday]) {
    var mode = String(scheduleMap[rowObj.weekday].event_mode || '').trim();
    if (mode) return mode.toUpperCase();
  }

  return String(settings.default_event_mode || 'ALL_DAY').toUpperCase();
}

function chooseTimes_(rowObj, scheduleMap) {
  var start = String(rowObj.start_time_override || '').trim();
  var end = String(rowObj.end_time_override || '').trim();

  if ((!start || !end) && scheduleMap[rowObj.weekday]) {
    start = start || String(scheduleMap[rowObj.weekday].start_time || '').trim();
    end = end || String(scheduleMap[rowObj.weekday].end_time || '').trim();
  }

  return { start: start, end: end };
}

function parseTimeOnDate_(dateObj, hhmm) {
  var parts = String(hhmm).split(':');
  if (parts.length < 2) throw new Error('Invalid time value: ' + hhmm);

  var d = new Date(dateObj);
  d.setHours(parseInt(parts[0], 10), parseInt(parts[1], 10), 0, 0);
  return d;
}

function buildBaseTitle_(rowObj, settings) {
  var prefix = String(settings.title_prefix || '');
  return prefix + String(rowObj.title || rowObj.session_type || 'Training');
}

function buildEventTitle_(rowObj, settings) {
  var base = buildBaseTitle_(rowObj, settings);
  var status = String(rowObj.status || 'planned').toLowerCase();
  var suffix = [];

  if (status === 'done') {
    if (!isBlankish_(rowObj.actual_km)) suffix.push(rowObj.actual_km + ' km');
    if (!isBlankish_(rowObj.actual_avg_pace)) suffix.push(rowObj.actual_avg_pace);
    if (!isBlankish_(rowObj.actual_avg_hr)) suffix.push('HR ' + rowObj.actual_avg_hr);
    return '✅ ' + base + (suffix.length ? ' | ' + suffix.join(' | ') : '');
  }
  if (status === 'skipped') return '❌ ' + base;
  if (status === 'modified') return '⚠ ' + base;
  return base;
}

function buildFeedbackBlock_(rowObj, settings) {
  var startMarker = settings.feedback_block_start || '[FEEDBACK]';
  var endMarker = settings.feedback_block_end || '[/FEEDBACK]';

  var lines = [
    startMarker,
    'status=' + (rowObj.status || ''),
    'km=' + (rowObj.actual_km || ''),
    'duration_min=' + (rowObj.actual_duration_min || ''),
    'avg_hr=' + (rowObj.actual_avg_hr || ''),
    'max_hr=' + (rowObj.actual_max_hr || ''),
    'avg_pace=' + (rowObj.actual_avg_pace || ''),
    'rpe=' + (rowObj.actual_rpe || ''),
    'notes=' + (rowObj.actual_notes || ''),
    endMarker
  ];

  return lines.join('\n');
}

function buildEventDescription_(rowObj, settings, rowNumber) {
  var lines = [];
  lines.push('PLAN');
  lines.push('row=' + rowNumber);
  lines.push('date=' + formatDate_(rowObj.date));
  lines.push('title=' + (rowObj.title || ''));
  lines.push('type=' + (rowObj.session_type || ''));
  lines.push('planned_duration_min=' + (rowObj.planned_duration_min || ''));
  lines.push('planned_km=' + (rowObj.planned_km || ''));
  lines.push('planned_hr=' + ((rowObj.planned_hr_min || '') + (rowObj.planned_hr_max ? '-' + rowObj.planned_hr_max : '')));
  lines.push('planned_notes=' + (rowObj.planned_notes || ''));
  lines.push('');
  lines.push('Feedback: edit only the block below. You can use "=" or ":".');
  lines.push('Examples: status=done, km=12, avg_hr=138, skipped');
  lines.push('');
  lines.push(buildFeedbackBlock_(rowObj, settings));
  return lines.join('\n');
}

function parseFeedbackBlock_(description, settings) {
  description = String(description || '');

  var startMarker = settings.feedback_block_start || '[FEEDBACK]';
  var endMarker = settings.feedback_block_end || '[/FEEDBACK]';

  var start = description.indexOf(startMarker);
  var end = description.indexOf(endMarker);
  if (start === -1 || end === -1 || end <= start) return null;

  var body = description.substring(start + startMarker.length, end);
  if (!body) return null;

  body = body
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .trim();

  if (!body) return null;

  if (/(?:^|\n|\r)\s*skipped\s*(?:$|\n|\r)/i.test(body) || /status\s*[:=]\s*skipped/i.test(body)) {
    return skippedPayload_();
  }

  var out = {};
  var lines = body.split(/\r?\n/);

  for (var i = 0; i < lines.length; i++) {
    var line = String(lines[i] || '')
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/\u00A0/g, ' ')
      .trim();

    if (!line) continue;
    if (line.charAt(0) === '#') continue;

    if (line.indexOf('=') === -1 && line.indexOf(':') === -1) {
      var loneStatus = normalizeActualStatus_(line);
      if (loneStatus === 'skipped') return skippedPayload_();
      if (loneStatus) out.status = loneStatus;
      continue;
    }

    var match = line.match(/^(.+?)\s*[:=]\s*(.*)$/);
    if (!match) continue;

    var key = canonicalFeedbackKey_(match[1]);
    var value = String(match[2] || '').trim();
    if (!key) continue;

    if (key === 'status') {
      var status = normalizeActualStatus_(value);
      if (status === 'skipped') return skippedPayload_();
      if (status) out.status = status;
    } else {
      if (!isBlankish_(value)) out[key] = value;
    }
  }

  return out;
}

function mergeActualsWithExistingRow_(rowObj, actuals) {
  var merged = {
    status: rowObj.status || '',
    actual_km: rowObj.actual_km || '',
    actual_duration_min: rowObj.actual_duration_min || '',
    actual_avg_hr: rowObj.actual_avg_hr || '',
    actual_max_hr: rowObj.actual_max_hr || '',
    actual_avg_pace: rowObj.actual_avg_pace || '',
    actual_rpe: rowObj.actual_rpe || '',
    actual_notes: rowObj.actual_notes || ''
  };

  if (!actuals) return merged;
  if (normalizeActualStatus_(actuals.status || '') === 'skipped') return skippedPayload_();

  var keys = ['status', 'actual_km', 'actual_duration_min', 'actual_avg_hr', 'actual_max_hr', 'actual_avg_pace', 'actual_rpe', 'actual_notes'];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (typeof actuals[key] === 'undefined') continue;
    if (isBlankish_(actuals[key])) continue;

    if (key === 'status') {
      var normalized = normalizeActualStatus_(actuals[key]);
      if (normalized) merged.status = normalized;
    } else {
      merged[key] = actuals[key];
    }
  }

  if ((!merged.status || merged.status === 'planned') && hasMeaningfulActuals_(merged)) {
    merged.status = 'done';
  }

  return merged;
}

function setRowValuesFromActuals_(sheet, rowNumber, headers, actualMap) {
  function setIfPresent(field, value) {
    if (typeof headers[field] === 'undefined') return;
    if (typeof value === 'undefined') return;
    sheet.getRange(rowNumber, headers[field] + 1).setValue(value);
  }

  setIfPresent('status', actualMap.status || '');
  setIfPresent('actual_km', actualMap.actual_km || '');
  setIfPresent('actual_duration_min', actualMap.actual_duration_min || '');
  setIfPresent('actual_avg_hr', actualMap.actual_avg_hr || '');
  setIfPresent('actual_max_hr', actualMap.actual_max_hr || '');
  if (typeof headers.actual_avg_pace !== 'undefined') {
  var paceCell = sheet.getRange(rowNumber, headers.actual_avg_pace + 1);
  paceCell.setNumberFormat('@STRING@');
  paceCell.setValue(String(actualMap.actual_avg_pace || ''));
}
  setIfPresent('actual_rpe', actualMap.actual_rpe || '');
  setIfPresent('actual_notes', actualMap.actual_notes || '');
}

function getPlanRows_() {
  var sh = getSheet_('Plan');
  var values = sh.getDataRange().getValues();
  var headers = getHeadersMap_(sh);
  return { sheet: sh, values: values, headers: headers };
}

function syncPlanToCalendar() {
  var settings = getSettings_();
  var calendar = getCalendar_();
  var scheduleMap = buildScheduleMap_();
  var plan = getPlanRows_();
  var sh = plan.sheet;
  var headers = plan.headers;
  var values = plan.values;

  for (var r = 1; r < values.length; r++) {
    var rowNumber = r + 1;
    var rowObj = rowToObject_(values[r], headers);

    try {
      if (!toBoolean_(rowObj.enabled)) {
        logAction_('INFO', 'syncPlanToCalendar', rowNumber, 'Skipped disabled row');
        continue;
      }

      if (!(rowObj.date instanceof Date)) {
        logAction_('WARN', 'syncPlanToCalendar', rowNumber, 'Skipped row with invalid date');
        continue;
      }

      var title = buildEventTitle_(rowObj, settings);
      var desc = buildEventDescription_(rowObj, settings, rowNumber);
      var mode = chooseEventMode_(rowObj, scheduleMap, settings);
      var times = chooseTimes_(rowObj, scheduleMap);
      var event = null;

      if (rowObj.calendar_event_id) {
        event = calendar.getEventById(String(rowObj.calendar_event_id));
      }

      if (!event) {
        if (mode === 'TIMED') {
          if (!times.start || !times.end) throw new Error('TIMED event without start/end time');
          event = calendar.createEvent(title, parseTimeOnDate_(rowObj.date, times.start), parseTimeOnDate_(rowObj.date, times.end), { description: desc });
        } else {
          event = calendar.createAllDayEvent(title, rowObj.date, { description: desc });
        }
      } else {
        if (mode === 'TIMED') {
          if (!times.start || !times.end) throw new Error('TIMED event without start/end time');
          event.setTime(parseTimeOnDate_(rowObj.date, times.start), parseTimeOnDate_(rowObj.date, times.end));
        } else {
          event.setAllDayDate(rowObj.date);
        }
        event.setTitle(title);
        event.setDescription(desc);
      }

      try {
        event.setTag('sheet_row', String(rowNumber));
        event.setTag('plan_date', formatDate_(rowObj.date));
      } catch (e) {}

      if (typeof headers.calendar_event_id !== 'undefined') sh.getRange(rowNumber, headers.calendar_event_id + 1).setValue(event.getId());
      if (typeof headers.event_last_synced_at !== 'undefined') sh.getRange(rowNumber, headers.event_last_synced_at + 1).setValue(new Date());
      if (typeof headers.calendar_title_snapshot !== 'undefined') sh.getRange(rowNumber, headers.calendar_title_snapshot + 1).setValue(event.getTitle());
      if (typeof headers.feedback_summary !== 'undefined') sh.getRange(rowNumber, headers.feedback_summary + 1).setValue(buildFeedbackSummary_(rowObj));

      logAction_('INFO', 'syncPlanToCalendar', rowNumber, 'Synced ' + event.getTitle());
    } catch (e) {
      logAction_('ERROR', 'syncPlanToCalendar', rowNumber, e.message);
    }
  }
}

function syncFeedbackFromCalendar() {
  var settings = getSettings_();
  if (String(settings.feedback_from_calendar_enabled).toUpperCase() !== 'TRUE') {
    throw new Error('feedback_from_calendar_enabled is not TRUE in Settings');
  }

  var calendar = getCalendar_();
  var plan = getPlanRows_();
  var sh = plan.sheet;
  var headers = plan.headers;
  var values = sh.getDataRange().getValues();

  for (var r = 1; r < values.length; r++) {
    var rowNumber = r + 1;
    var rowObj = rowToObject_(values[r], headers);

    try {
      if (!toBoolean_(rowObj.enabled)) continue;
      if (!rowObj.calendar_event_id) continue;

      var event = calendar.getEventById(String(rowObj.calendar_event_id));
      if (!event) {
        logAction_('WARN', 'syncFeedbackFromCalendar', rowNumber, 'Event not found for stored event id');
        continue;
      }

      var rawDescription = event.getDescription();
      logAction_('INFO', 'syncFeedbackFromCalendar', rowNumber, 'RAW DESCRIPTION = ' + rawDescription);

      var actuals = parseFeedbackBlock_(rawDescription, settings);
      logAction_('INFO', 'syncFeedbackFromCalendar', rowNumber, 'Parsed FEEDBACK = ' + JSON.stringify(actuals));

      if (!actuals) {
        logAction_('INFO', 'syncFeedbackFromCalendar', rowNumber, 'No FEEDBACK block found');
        continue;
      }

      if (!hasMeaningfulActuals_(actuals)) {
        logAction_('INFO', 'syncFeedbackFromCalendar', rowNumber, 'Skipped import because no meaningful feedback. Parsed=' + JSON.stringify(actuals));
        continue;
      }

      var mergedActuals = mergeActualsWithExistingRow_(rowObj, actuals);
      logAction_('INFO', 'syncFeedbackFromCalendar', rowNumber, 'Merged result = ' + JSON.stringify(mergedActuals));

      setRowValuesFromActuals_(sh, rowNumber, headers, mergedActuals);

      var freshRow = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
      var freshObj = rowToObject_(freshRow, headers);
      var feedbackSummary = buildFeedbackSummary_(freshObj);

      if (typeof headers.feedback_summary !== 'undefined') sh.getRange(rowNumber, headers.feedback_summary + 1).setValue(feedbackSummary);
      if (typeof headers.event_last_synced_at !== 'undefined') sh.getRange(rowNumber, headers.event_last_synced_at + 1).setValue(new Date());

      event.setTitle(buildEventTitle_(freshObj, settings));
      event.setDescription(buildEventDescription_(freshObj, settings, rowNumber));

      if (typeof headers.calendar_title_snapshot !== 'undefined') sh.getRange(rowNumber, headers.calendar_title_snapshot + 1).setValue(event.getTitle());

      logAction_('INFO', 'syncFeedbackFromCalendar', rowNumber, 'Imported FEEDBACK block and refreshed event');
    } catch (e) {
      logAction_('ERROR', 'syncFeedbackFromCalendar', rowNumber, e.message);
    }
  }
}

function resetSyncStatus() {
  var plan = getPlanRows_();
  var sh = plan.sheet;
  var headers = plan.headers;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var colsToClear = ['calendar_event_id', 'event_last_synced_at', 'calendar_title_snapshot'];
  for (var i = 0; i < colsToClear.length; i++) {
    if (typeof headers[colsToClear[i]] === 'undefined') continue;
    var col = headers[colsToClear[i]] + 1;
    sh.getRange(2, col, lastRow - 1, 1).clearContent();
  }

  logAction_('INFO', 'resetSyncStatus', '', 'Cleared event ids and sync timestamps');
}

function deletePrefixedEvents() {
  var settings = getSettings_();
  var prefix = String(settings.delete_prefix || settings.title_prefix || '').trim();
  if (!prefix) throw new Error('delete_prefix is blank in Settings');

  var plan = getPlanRows_();
  var values = plan.values;
  var headers = plan.headers;
  var minDate = null;
  var maxDate = null;

  for (var r = 1; r < values.length; r++) {
    var d = values[r][headers.date];
    if (d instanceof Date) {
      if (!minDate || d < minDate) minDate = new Date(d);
      if (!maxDate || d > maxDate) maxDate = new Date(d);
    }
  }

  if (!minDate || !maxDate) throw new Error('No dates found in Plan');

  minDate.setDate(minDate.getDate() - 30);
  maxDate.setDate(maxDate.getDate() + 30);

  var calendar = getCalendar_();
  var events = calendar.getEvents(minDate, maxDate);
  var deleted = 0;

  for (var i = 0; i < events.length; i++) {
    var t = cleanLeadingStatus_(events[i].getTitle());
    if (t.indexOf(prefix) === 0) {
      events[i].deleteEvent();
      deleted++;
    }
  }

  resetSyncStatus();
  logAction_('INFO', 'deletePrefixedEvents', '', 'Deleted ' + deleted + ' events with prefix ' + prefix);
}

function testCalendarAccess() {
  var cal = getCalendar_();
  logAction_('INFO', 'testCalendarAccess', '', 'Calendar OK: ' + cal.getName());
}

function syncAll() {
  syncFeedbackFromCalendar();
  syncPlanToCalendar();
}
