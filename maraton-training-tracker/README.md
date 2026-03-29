# Marathon Training Tracker

Google Apps Script for syncing a marathon training plan between Google Sheets and Google Calendar.

The script file is: `marathon_calendar_feedback_fueling.gs`

## What It Does

- Syncs rows from `Plan` sheet to Calendar events (`syncPlanToCalendar`)
- Pulls workout feedback from event descriptions back into the `Plan` sheet (`syncFeedbackFromCalendar`)
- Keeps event title and description aligned with actual run data
- Supports all-day and timed events with per-day and per-row overrides
- Writes operational logs to `Logs`

## Custom Menu

When the spreadsheet opens, the script adds this menu:

- `Training Feedback Sync -> Sync Plan -> Calendar`
- `Training Feedback Sync -> Sync Feedback <- Calendar`
- `Training Feedback Sync -> Sync All`
- `Training Feedback Sync -> Test Calendar Access`
- `Training Feedback Sync -> Reset Sync Status`
- `Training Feedback Sync -> Delete Prefixed Events`

`Sync All` runs:
1. Feedback import from Calendar
2. Plan push to Calendar

## Required Sheets

You must have these sheet tabs:

- `Plan`
- `Schedule`
- `Settings`
- `Logs`

## Settings Sheet

`Settings` is read as key-value pairs:

- Column A: setting key
- Column B: value

Recommended keys:

- `calendar_id` (example: `primary` or your calendar ID)
- `use_schedule` (`TRUE`/`FALSE`)
- `default_event_mode` (`ALL_DAY` or `TIMED`)
- `title_prefix` (example: `Marathon | `)
- `feedback_from_calendar_enabled` (`TRUE` required for feedback import)
- `feedback_block_start` (default `[FEEDBACK]`)
- `feedback_block_end` (default `[/FEEDBACK]`)
- `delete_prefix` (used by `Delete Prefixed Events`)

If `feedback_block_start` and `feedback_block_end` are not set, the defaults above are used.

## Plan Sheet Columns

Minimum practical columns (exact header names):

- `enabled`
- `date`
- `weekday`
- `title`
- `session_type`
- `planned_duration_min`
- `planned_km`
- `planned_hr_min`
- `planned_hr_max`
- `planned_notes`
- `event_mode_override`
- `start_time_override`
- `end_time_override`
- `status`
- `actual_km`
- `actual_duration_min`
- `actual_avg_hr`
- `actual_max_hr`
- `actual_avg_pace`
- `actual_rpe`
- `actual_notes`
- `actual_fuel_notes`
- `calendar_event_id`
- `event_last_synced_at`
- `calendar_title_snapshot`
- `feedback_summary`

Notes:

- `enabled` should be `TRUE` for rows that should sync.
- `date` must be a valid date.
- Rows without `enabled=TRUE` or valid `date` are skipped.
- `calendar_event_id`, `event_last_synced_at`, and `calendar_title_snapshot` are maintained by the script.

## Schedule Sheet Columns

Used when `use_schedule=TRUE` and there is no row override:

- `weekday`
- `event_mode` (`ALL_DAY` or `TIMED`)
- `start_time` (for timed events, format `HH:mm`)
- `end_time` (for timed events, format `HH:mm`)

## Logs Sheet

Script appends logs as:

- Timestamp
- Level (`INFO`, `WARN`, `ERROR`)
- Function name
- Plan row number
- Message

## Event Title Rules

Calendar title is generated from `title_prefix + (title or session_type)` and status:

- `planned`: no icon
- `done`: starts with `✅` and may append km/pace/hr
- `skipped`: starts with `❌`
- `modified`: starts with `⚠`

## Event Description Format

Each event description contains:

- A `PLAN` metadata section
- An editable feedback section between markers

Only edit the feedback block between `[FEEDBACK]` and `[/FEEDBACK]`.

Supported feedback keys (case/spacing tolerant):

- `status`
- `km`
- `duration_min`
- `avg_hr`
- `max_hr`
- `avg_pace`
- `fuel` (or `fueling`)
- `rpe`
- `notes`

`=` and `:` separators are both supported.

### Feedback examples

```text
[FEEDBACK]
status=done
km=12
duration_min=75
avg_hr=138
max_hr=152
avg_pace=5:42
fuel=banana + dates
rpe=6
notes=felt smooth
[/FEEDBACK]
```

Skipped workout examples:

```text
[FEEDBACK]
skipped
[/FEEDBACK]
```

```text
[FEEDBACK]
status=skipped
[/FEEDBACK]
```

## Sync Behavior

### Plan -> Calendar (`syncPlanToCalendar`)

- Creates new events when `calendar_event_id` is empty
- Updates existing events when `calendar_event_id` exists
- Uses event mode priority:
1. `event_mode_override` on row
2. `Schedule.event_mode` for matching `weekday` (if `use_schedule=TRUE`)
3. `default_event_mode`
- For `TIMED` events, start/end are chosen from row override first, then `Schedule`

### Calendar -> Feedback (`syncFeedbackFromCalendar`)

- Requires `feedback_from_calendar_enabled=TRUE`
- Reads FEEDBACK block from each linked event description
- Merges parsed values into existing actual columns
- Auto-sets `status=done` if meaningful actual metrics exist and status is still planned/blank
- Rewrites event title and full description after importing

## Utility Actions

- `Test Calendar Access`: verifies `calendar_id` and logs calendar name
- `Reset Sync Status`: clears `calendar_event_id`, `event_last_synced_at`, `calendar_title_snapshot`
- `Delete Prefixed Events`: deletes calendar events whose title starts with `delete_prefix` (or `title_prefix`), within a date window around `Plan` dates, then resets sync status

## Setup Steps

1. Create/open a Google Sheet and add tabs: `Plan`, `Schedule`, `Settings`, `Logs`.
2. Add header row for `Plan` and `Schedule` with the column names above.
3. Fill `Settings` keys and values.
4. Open Extensions -> Apps Script and paste `marathon_calendar_feedback_fueling.gs`.
5. Save, reload the sheet, then use the `Training Feedback Sync` menu.
6. Run `Test Calendar Access` first.
7. Run `Sync Plan -> Calendar` to generate events.
8. Edit feedback blocks in Calendar events, then run `Sync Feedback <- Calendar`.

## Troubleshooting

- `Missing sheet: ...`: a required tab is missing or named differently.
- `Calendar not found`: `calendar_id` is invalid.
- `TIMED event without start/end time`: no valid times found from override or `Schedule`.
- `feedback_from_calendar_enabled is not TRUE`: enable it in `Settings` before importing feedback.

