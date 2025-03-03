# School Email to Calendar Integration

This Google Apps Script automatically extracts events from your school's emails and adds them to a shared Google Calendar, making it easier to keep track of important school dates and events.

## Features

- 🔍 Automatically searches for emails from your school's domain
- 📅 Extracts dates and times from email content
- 🗓️ Adds events to a shared Google Calendar
- 🏷️ Labels processed emails to avoid duplication
- ⏰ Can run daily or on-demand

## Setup Instructions

### 1. Create or Identify a Shared Google Calendar

1. Go to [Google Calendar](https://calendar.google.com)
2. Create a new calendar or use an existing one
3. In calendar settings, find the Calendar ID
   - It looks like: `something@group.calendar.google.com`
   - You'll need this ID for the script configuration

### 2. Set Up the Script

1. Go to [Google Apps Script](https://script.google.com)
2. Create a new project
3. Copy and paste the code from `school-calendar-script.js` into the editor
4. Modify the configuration variables at the top of the script:
   ```javascript
   const SEARCH_QUERY = "from:@yourschooldomain.org"; // Change to your school's email domain
   const TARGET_CALENDAR_ID = "your_calendar_id@group.calendar.google.com"; // Your calendar ID
   const DAYS_TO_LOOK_BACK = 7; // Adjust as needed
   const LABEL_NAME = "School-Processed"; // Label name for processed emails
   const SCHOOL_NAME_PREFIX = "School Event: "; // Customize the event prefix
   ```
5. Save the project with a meaningful name (e.g., "School Calendar Integration")

### 3. Test the Setup

1. Run the `testCalendarAccess` function first
   - Select it from the dropdown menu at the top of the editor
   - Click the Run button (▶️)
   - Grant the necessary permissions when prompted
   - Check the execution logs to confirm success
   - Verify a test event appears in your calendar

2. Run the main script manually
   - Select the `processSchoolEmails` function from the dropdown menu
   - Click the Run button (▶️)
   - Check the execution logs to see what happened
   - Look for events in your calendar

### 4. Set Up Automatic Daily Runs (Optional)

1. Run the `setUpDailyTrigger` function once
   - This will configure the script to run daily at 6 AM
   - You can modify the time in the script if needed

## Troubleshooting

### Common Issues

- **No events being added:**
  - Verify your school's email domain in the `SEARCH_QUERY`
  - Make sure you have emails from that domain with dates in them
  - Check the logs for details on what the script found

- **Calendar access issues:**
  - Confirm your Calendar ID is correct by running `testCalendarAccess`
  - Make sure you have edit permissions for the calendar

- **Date detection issues:**
  - The script looks for common date formats, but some schools may use unusual formats
  - Check the logs to see what content is being processed
  - You can add additional date patterns in the `datePatterns` array

### Getting Better Results

- For first-time use, consider increasing `DAYS_TO_LOOK_BACK` to catch older events
- Refine the `SEARCH_QUERY` if your school uses multiple email domains
- Check the execution logs for detailed information on what's being processed

## License

MIT License - Feel free to modify and share this script!

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
