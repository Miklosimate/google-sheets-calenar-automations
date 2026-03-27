# Work Hour Tracker - Google Sheets & Calendar Integration

Automatically track your work hours from Google Calendar events and generate detailed reports with earnings calculations.

## Features

- 📅 Fetches work events from Google Calendar
- 📊 Generates monthly reports with hourly breakdowns
- 💰 Calculates paid time (with automatic break deductions)
- 💵 Computes total earnings based on hourly rate
- 📈 Weekly progress tracking with AI work separation
- ⚡ Apple Shortcut integration for one-click updates

## Setup Instructions

### 1. Create a Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Click **"+ Create"** → **"Blank spreadsheet"**
3. Rename it to "Work Hour Tracker"

### 2. Create Required Sheets

Your spreadsheet must have these sheets:

#### **Control Sheet**
The first sheet that manages configuration:

| Cell | Value |
|------|-------|
| A1 | Year (e.g., 2024) |
| B1 | Month (1-12) |
| C1 | Hourly Rate (in your currency) |
| E1 | Target Hours (weekly target, e.g., 40) |

Example:
```
A1: 2024
B1: 3
C1: 5000
E1: 40
```

#### **Additional Sheets**
- The script will automatically create sheets named `YYYY_MMM` (e.g., `2024_Mar`) for each month you generate reports for

### 3. Add the Google Apps Script

1. In your Google Sheet, click **"Extensions"** → **"Apps Script"**
2. Delete any existing code
3. Copy and paste ALL three script files from this folder:
   - `report_auto.gs` - Generates detailed monthly reports
   - `report_status.gs` - Creates weekly progress events
   - `linktrigger.gs` - Webhook endpoint for Apple Shortcut integration
4. Save the project (give it a name like "Work Hour Tracker")

### 4. Configure Your Calendar ID

1. Go to [Google Calendar Settings](https://calendar.google.com/calendar/u/0/r/settings)
2. Select the calendar you use for work events
3. Scroll to **"Calendar details"**
4. Copy the **Calendar ID**
5. In both `report_auto.gs` and `report_status.gs`, replace:
   - `YOUR_WORK_CALENDAR_ID@group.calendar.google.com` with your actual calendar ID

### 5. Authorize the Script

1. Click the **▶️ Run** button in Apps Script
2. Select `createSheetAndFetchEvents` from the dropdown
3. Authorize the script to access your Calendar and Sheets
4. If you see a security warning, click **"Advanced"** → **"Go to [Project Name] (unsafe)"**

### 6. Set Up Calendar Events

Tag your work events in Google Calendar with **"munka"** in the title:
- ✅ Good: "munka - Client Project"
- ✅ Good: "munka (43)"
- ❌ Bad: "Work session" (won't be tracked without "munka")

For AI work, add **"ai"** to the title:
- "AI Research"
- "AI munka"

### How to Generate Reports

#### Method 1: Manual (From Google Sheet)
1. Update the **Control** sheet with desired Year, Month, and Hourly Rate
2. Click the menu button (⋮) → Look for "Work Tracker" menu (if it appears)
3. Select "Generate Monthly Report"
4. A new sheet `YYYY_MMM` is created with all tracked hours

#### Method 2: Weekly Progress (Automatic)
The script creates weekly summary events on Friday:
- Example: "W:35:30/A:04:00 [PROGRESS]" means 35.5 hours worked, 4 hours on AI

### Apple Shortcut Integration

Use Apple Shortcuts to trigger the script with one tap from your phone:

#### Create the Shortcut

1. On iPhone/iPad, open the **Shortcuts** app
2. Tap **"Create Shortcut"**
3. Add these actions:

```
Ask for "Year" (number)
Ask for "Month" (number)
Ask for "Hourly Rate" (number)
Ask for "Target Hours" (number)

HTTP POST Request:
  URL: [YOUR_DEPLOYMENT_URL]
  
  Add Body:
    - year: Year
    - month: Month
    - hourlyRate: Hourly Rate
    - targetHours: Target Hours
    
  Headers:
    - Content-Type: application/x-www-form-urlencoded

Display the response
```

#### Get Your Deployment URL

1. In Apps Script, click **"Deploy"** → **"New deployment"**
2. Select **Type: "Web app"**
3. Configure:
   - **Execute as:** Your email
   - **Who has access:** "Anyone"
4. Click **"Deploy"**
5. Copy the generated URL (looks like: `https://script.google.com/macros/d/.../usercontent`)

#### Set Up Calendar Trigger (Optional)

To automatically run the shortcut when you open Google Calendar:

1. In Shortcuts app, go to **"Automation"**
2. Tap **"+"** and select **"Open app"**
3. Choose **Google Calendar**
4. Turn on **"Notify When Run"**
5. Select your "Work Hour Tracker" shortcut

Now whenever you open Google Calendar, the shortcut can optionally run!

## Feature Details

### Break Deductions
The script automatically deducts breaks from paid time:
- **6-9 hours:** 20 minutes deducted
- **9+ hours:** 45 minutes deducted

### Monthly Report
Shows:
- Date, start/end times
- Duration and paid duration
- Weekly totals
- Total earnings (paid hours × hourly rate)

### Weekly Progress Events
Creates Friday events showing:
- **W:** Total work hours for the week
- **A:** AI project hours (if tagged with "ai")

## Troubleshooting

**"Calendar not found" error:**
- Verify your Calendar ID is correct
- Make sure the calendar exists
- Check that you've authorized the script

**No events appearing:**
- Ensure events have "munka" in the title
- Check the date range (year/month in Control sheet)
- Verify the calendar is shared with your Google account

**Script won't run:**
- Confirm all 3 .gs files are pasted
- Check that year/month/hourly rate values are in Control sheet
- Re-authorize the script

**Shortcut not working:**
- Verify the deployment URL is correct
- Check internet connection
- Ensure the shortcut payload matches the URL requirements

## Security Notes

- The deployment URL is public, but it only triggers your script actions on your sheet
- Consider limiting automations to manual trigger if privacy is a concern
- Never share your Calendar ID with untrusted sources

---

For questions or improvements, refer to the individual script files for more details on functionality.
