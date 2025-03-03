/**
 * School Email to Google Calendar Integration
 * 
 * This script searches Gmail for emails from your school's domain,
 * extracts dates and event information, and adds them to a specified
 * shared Google Calendar.
 *
 * GitHub: https://github.com/YourUsername/school-calendar-integration
 * License: MIT
 */

// Configuration variables - MODIFY THESE TO MATCH YOUR NEEDS
const SEARCH_QUERY = "from:@yourschooldomain.org"; // Target emails from the school domain
const TARGET_CALENDAR_ID = "your_calendar_id@group.calendar.google.com"; // Your shared calendar ID
const DAYS_TO_LOOK_BACK = 7; // How many days of emails to search through
const LABEL_NAME = "School-Processed"; // Label to apply to processed emails
const SCHOOL_NAME_PREFIX = "School Event: "; // Prefix for calendar events

/**
 * Main function that orchestrates the email processing
 * This is the function you should run manually when needed
 */
function processSchoolEmails() {
  // Create the label if it doesn't exist
  let label = getOrCreateLabel(LABEL_NAME);
  
  Logger.log(`Starting search with query: "${SEARCH_QUERY}" for emails in the last ${DAYS_TO_LOOK_BACK} days`);
  
  // Get emails matching our search criteria
  let threads = getRelevantEmailThreads(SEARCH_QUERY, DAYS_TO_LOOK_BACK);
  
  if (threads.length === 0) {
    Logger.log("No emails found matching the search criteria.");
    return;
  }
  
  Logger.log(`Found ${threads.length} email threads to process.`);
  
  // Process each thread
  let processedCount = 0;
  let skippedCount = 0;
  let eventsCreated = 0;
  
  for (let thread of threads) {
    // Log the subject of the thread for debugging
    let firstMessage = thread.getMessages()[0];
    Logger.log(`Processing thread: "${firstMessage.getSubject()}" from ${firstMessage.getFrom()}`);
    
    // Skip already processed threads
    if (thread.getLabels().some(l => l.getName() === LABEL_NAME)) {
      Logger.log("  Skipping thread - already processed");
      skippedCount++;
      continue;
    }
    
    let messages = thread.getMessages();
    for (let message of messages) {
      let newEvents = processEmail(message);
      eventsCreated += newEvents;
    }
    
    // Mark thread as processed
    thread.addLabel(label);
    processedCount++;
  }
  
  Logger.log(`Processing complete: ${processedCount} threads processed, ${skippedCount} threads skipped, ${eventsCreated} calendar events created`);
}

/**
 * Gets email threads matching the search criteria within the specified time period
 */
function getRelevantEmailThreads(query, daysToLookBack) {
  let now = new Date();
  let pastDate = new Date(now.getTime() - (daysToLookBack * 24 * 60 * 60 * 1000));
  let formattedDate = Utilities.formatDate(pastDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  let fullQuery = `${query} after:${formattedDate}`;
  return GmailApp.search(fullQuery);
}

/**
 * Process an individual email message
 */
function processEmail(message) {
  let subject = message.getSubject();
  let body = message.getPlainBody();
  let from = message.getFrom();
  
  Logger.log(`Processing email: "${subject}" from ${from}`);
  
  // Log the first 200 characters of the body for debugging
  Logger.log(`Email preview: ${body.substring(0, 200).replace(/\n/g, ' ')}...`);
  
  // Extract event details
  let eventDetails = extractEventDetails(subject, body);
  
  if (eventDetails.length === 0) {
    Logger.log("No event details found in this email.");
    return 0;
  }
  
  Logger.log(`Found ${eventDetails.length} potential events in this email`);
  
  // Add events to calendar
  let eventsCreated = 0;
  for (let event of eventDetails) {
    Logger.log(`Attempting to create event: ${event.description} on ${event.date}`);
    let calEvent = addEventToCalendar(event, from, subject);
    if (calEvent) eventsCreated++;
  }
  
  return eventsCreated;
}

/**
 * Extract event details (date, time, description) from email content
 * This function uses various pattern matching techniques to identify dates and event information
 */
function extractEventDetails(subject, body) {
  let events = [];
  let combinedText = subject + "\n" + body;
  
  // Array of date patterns to look for
  let datePatterns = [
    // MM/DD/YYYY or MM/DD/YY
    /(\b\d{1,2}\/\d{1,2}\/\d{2,4}\b)/g,
    
    // Month DD, YYYY
    /(\b(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+\d{1,2}(?:st|nd|rd|th)?,\s+\d{4}\b)/gi,
    
    // Month DD
    /(\b(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+\d{1,2}(?:st|nd|rd|th)?\b)/gi,
    
    // MM-DD-YYYY
    /(\b\d{1,2}-\d{1,2}-\d{4}\b)/g,
    
    // YYYY-MM-DD (ISO format)
    /(\b\d{4}-\d{1,2}-\d{1,2}\b)/g
  ];
  
  // Time patterns
  let timePattern = /(\b(?:\d{1,2}:\d{2}\s*(?:am|pm|AM|PM)|(?:\d{1,2})\s*(?:am|pm|AM|PM))\b)/g;
  
  // Look for each date pattern
  for (let pattern of datePatterns) {
    let match;
    while ((match = pattern.exec(combinedText)) !== null) {
      let dateStr = match[1];
      let dateObj = parseDateString(dateStr);
      
      if (!dateObj) continue;
      
      // Look for times in nearby context (within 100 characters)
      let contextStart = Math.max(0, match.index - 100);
      let contextEnd = Math.min(combinedText.length, match.index + 100);
      let context = combinedText.substring(contextStart, contextEnd);
      
      let times = [];
      let timeMatch;
      while ((timeMatch = timePattern.exec(context)) !== null) {
        times.push(timeMatch[1]);
      }
      
      // Extract a description (up to 150 chars after the date)
      let descEnd = Math.min(combinedText.length, match.index + 150);
      let description = combinedText.substring(match.index, descEnd)
        .replace(dateStr, '')
        .trim()
        .split('\n')[0];
      
      // If we have times, create an event for each time
      if (times.length > 0) {
        for (let time of times) {
          let dateTime = parseDateTime(dateStr, time);
          events.push({
            date: dateObj,
            dateTime: dateTime,
            hasTime: true,
            description: description
          });
        }
      } else {
        // All-day event
        events.push({
          date: dateObj,
          dateTime: null,
          hasTime: false,
          description: description
        });
      }
    }
  }
  
  return events;
}

/**
 * Parse a date string into a Date object
 */
function parseDateString(dateStr) {
  try {
    // Handle different date formats
    
    // MM/DD/YYYY
    if (dateStr.match(/\d{1,2}\/\d{1,2}\/\d{4}/)) {
      let parts = dateStr.split('/');
      return new Date(parts[2], parts[0]-1, parts[1]);
    }
    
    // MM/DD/YY
    if (dateStr.match(/\d{1,2}\/\d{1,2}\/\d{2}/)) {
      let parts = dateStr.split('/');
      let year = parseInt(parts[2]) + 2000; // Assume 20xx
      return new Date(year, parts[0]-1, parts[1]);
    }
    
    // Month DD, YYYY
    if (dateStr.match(/[A-Za-z]+ \d{1,2}(?:st|nd|rd|th)?, \d{4}/)) {
      dateStr = dateStr.replace(/(\d{1,2})(?:st|nd|rd|th)/, '$1'); // Remove ordinals
      return new Date(dateStr);
    }
    
    // Month DD (current year)
    if (dateStr.match(/[A-Za-z]+ \d{1,2}(?:st|nd|rd|th)?/)) {
      dateStr = dateStr.replace(/(\d{1,2})(?:st|nd|rd|th)/, '$1'); // Remove ordinals
      let currentYear = new Date().getFullYear();
      return new Date(`${dateStr}, ${currentYear}`);
    }
    
    // Default: let JavaScript try to parse it
    return new Date(dateStr);
  } catch (e) {
    Logger.log(`Error parsing date string "${dateStr}": ${e}`);
    return null;
  }
}

/**
 * Parse a date and time string into a Date object
 */
function parseDateTime(dateStr, timeStr) {
  try {
    let dateObj = parseDateString(dateStr);
    if (!dateObj) return null;
    
    // Parse time (12:30pm or 12pm format)
    let hours = 0;
    let minutes = 0;
    let isPM = timeStr.toLowerCase().indexOf('pm') > -1;
    
    if (timeStr.indexOf(':') > -1) {
      // Format: 12:30pm
      let timeParts = timeStr.replace(/[^0-9:]/g, '').split(':');
      hours = parseInt(timeParts[0]);
      minutes = parseInt(timeParts[1]);
    } else {
      // Format: 12pm
      hours = parseInt(timeStr.replace(/[^0-9]/g, ''));
      minutes = 0;
    }
    
    // Adjust for PM
    if (isPM && hours < 12) {
      hours += 12;
    }
    // Adjust for 12AM
    if (!isPM && hours === 12) {
      hours = 0;
    }
    
    dateObj.setHours(hours, minutes, 0, 0);
    return dateObj;
  } catch (e) {
    Logger.log(`Error parsing date and time: ${e}`);
    return null;
  }
}

/**
 * Add an event to the specified calendar
 */
function addEventToCalendar(eventDetails, from, subject) {
  try {
    let calendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
    
    if (!calendar) {
      Logger.log(`Could not find calendar with ID: ${TARGET_CALENDAR_ID}`);
      return;
    }
    
    // Create title based on subject or description
    let title = SCHOOL_NAME_PREFIX + (eventDetails.description || subject);
    
    // Limit title length
    if (title.length > 100) {
      title = title.substring(0, 97) + "...";
    }
    
    // Create description with email source
    let description = `Event from email: "${subject}"\nFrom: ${from}\n\n${eventDetails.description || ''}`;
    
    // Add to calendar based on whether it has a specific time
    let calendarEvent;
    if (eventDetails.hasTime && eventDetails.dateTime) {
      // Create event with time (default duration: 1 hour)
      let endTime = new Date(eventDetails.dateTime.getTime() + 60 * 60 * 1000);
      calendarEvent = calendar.createEvent(title, eventDetails.dateTime, endTime, {
        description: description
      });
    } else {
      // Create all-day event
      calendarEvent = calendar.createAllDayEvent(title, eventDetails.date, {
        description: description
      });
    }
    
    Logger.log(`Created event: "${title}" on ${eventDetails.date}`);
    return calendarEvent;
  } catch (e) {
    Logger.log(`Error creating calendar event: ${e}`);
    return null;
  }
}

/**
 * Get or create a Gmail label
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

/**
 * Test function to verify calendar access
 * Run this first to make sure your calendar ID is correct
 */
function testCalendarAccess() {
  try {
    const calendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
    if (calendar) {
      Logger.log(`Successfully accessed calendar: ${calendar.getName()}`);
      
      // Create a test event
      const now = new Date();
      const testEvent = calendar.createEvent(
        "Test Event - Script Setup", 
        new Date(now.getTime() + 24 * 60 * 60 * 1000), // Tomorrow
        new Date(now.getTime() + 25 * 60 * 60 * 1000), // Tomorrow + 1 hour
        {description: "This is a test event to verify calendar access. You can delete this."}
      );
      
      Logger.log(`Test event created with ID: ${testEvent.getId()}`);
      return true;
    } else {
      Logger.log("Could not access the calendar - check the Calendar ID");
      return false;
    }
  } catch (e) {
    Logger.log(`Error accessing calendar: ${e}`);
    return false;
  }
}

/**
 * Utility function to set up a daily trigger for this script
 * Run this function manually once to set up the automation
 */
function setUpDailyTrigger() {
  // Delete any existing triggers
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "processSchoolEmails") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new daily trigger
  ScriptApp.newTrigger("processSchoolEmails")
    .timeBased()
    .everyDays(1)
    .atHour(6) // Run at 6 AM
    .create();
    
  Logger.log("Daily trigger set up successfully!");
}
