# Google Sheets Calendar Automations

A collection of Google Apps Scripts for automating workflows with Google Sheets and Google Calendar.

<p><strong><span style="color: red;">Warning: These scripts are mostly vibe-coded for personal use and are not tested to be easily transferable. They were built to save me several hours of work, but feel free to use and modify them as you like.</span></strong></p>

## 🚀 Projects

### 📅 [Marathon Training Tracker](maraton-training-tracker/)
Two-way sync between a Google Sheets marathon plan and Google Calendar events.

recommended usage: Ask ChatGPT to re-write the plan part according to your needs and fitness level, then use the script to keep your calendar and sheet in sync, and to track your feedback on each workout.

**Features:**
- Plan -> Calendar sync (creates/updates all-day or timed training events)
- Calendar -> Sheet feedback import via editable `[FEEDBACK]` block
- Status-aware event titles (`planned`, `done`, `skipped`, `modified`)
- Per-row event overrides (mode, start/end time)
- Sync metadata and logging (`event ID`, sync timestamp, title snapshot, logs)
- Utility actions (calendar access test, reset sync state, bulk delete by prefix)

**Requires these sheets:**
- `Plan`
- `Schedule`
- `Settings`
- `Logs`

[→ Setup Guide](maraton-training-tracker/README.md)

---

### ⏱️ [Work Hour Tracker](work-hour-tracker/)
Track work hours from calendar events and generate detailed reports with earnings calculations.

**Features:**
- Automatic hour tracking from calendar events
- Monthly reports with hourly breakdowns
- Earnings calculations with break deductions
- Weekly progress tracking
- Apple Shortcut integration (trigger from calendar)

[→ Setup Guide](work-hour-tracker/README.md)

---

## 📋 Quick Start

1. Choose a project above
2. Follow the **Setup Guide** in its README
3. Upload your data to Google Sheets
4. Paste the Google Apps Script code
5. Authorize and run

## 🔧 Common Requirements

- Google account with Sheets and Calendar access
- A Google Sheet for your data
- Google Apps Script enabled
- Your calendar ID (optional, can use "primary")

## 📝 General Notes

- Each project is self-contained
- Scripts are independent per spreadsheet
- All code is open and customizable
- Refer to individual README files for detailed configuration

---

Choose a project above to get started!
