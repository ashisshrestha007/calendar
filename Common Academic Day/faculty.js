/**
 * APP SCRIPT 1: FACULTY CALENDAR EVENTS (FIXED VERSION)
 * Purpose: Create calendar events from Faculty sheet data, send emails, and mark as "Done"
 * 
 * SHEET STRUCTURE:
 * Faculty Sheet:
 *   Column A: Event Title
 *   Column B: Start Date (MM/DD/YYYY or any valid date format)
 *   Column C: Start Time (HH:MM AM/PM, H:MM AM/PM, HH:MM, or H:MM)
 *   Column D: Description
 *   Column E: Calendar ID (Google Calendar email)
 *   Column F: Guest Email
 *   Column G: Status (auto-filled with "Done")
 */

function createFacultyCalendarEvents() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const facultySheet = spreadsheet.getSheetByName("Faculty");
    
    // Validate sheet exists
    if (!facultySheet) {
      Logger.log("ERROR: Faculty sheet not found!");
      SpreadsheetApp.getUi().alert("ERROR: Faculty sheet not found!");
      return;
    }

    const data = facultySheet.getDataRange().getValues();
    let processedCount = 0;
    let errorCount = 0;

    // Headers in row 1 (index 0)
    // Start processing from row 2 (index 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      try {
        // Extract data from columns
        const eventTitle = row[0];        // Column A
        const startDate = row[1];         // Column B
        const startTime = row[2];         // Column C
        const description = row[3];       // Column D
        const calendarId = row[4];        // Column E
        const guestEmail = row[5];        // Column F
        const status = row[6];            // Column G

        // Skip if empty title or already "Done"
        if (!eventTitle || status === "Done") {
          continue;
        }

        // Log raw data for debugging
        Logger.log(`Row ${i + 1}: Title="${eventTitle}", Date="${startDate}", Time="${startTime}"`);

        // Validate data
        if (!validateEventData(eventTitle, startDate, startTime, calendarId, guestEmail, i + 1)) {
          errorCount++;
          continue;
        }

        // Create calendar event
        const eventId = createEvent(eventTitle, startDate, startTime, description, calendarId, guestEmail);

        if (eventId) {
          // Send confirmation email
          const emailSent = sendConfirmationEmail(guestEmail, eventTitle, startDate, startTime, description, calendarId);

          // Update status if email sent
          if (emailSent) {
            facultySheet.getRange(i + 1, 7).setValue("Done"); // Column G = row 7
            processedCount++;
            Logger.log(`✓ Row ${i + 1}: Event created and email sent for "${eventTitle}"`);
          } else {
            Logger.log(`✗ Row ${i + 1}: Event created but email failed for "${eventTitle}"`);
            errorCount++;
          }
        } else {
          Logger.log(`✗ Row ${i + 1}: Failed to create event for "${eventTitle}"`);
          errorCount++;
        }

      } catch (error) {
        Logger.log(`ERROR in row ${i + 1}: ${error.message}`);
        errorCount++;
      }
    }

    // Show summary
    const message = `Completed!\n\nProcessed: ${processedCount} events\nErrors: ${errorCount}`;
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    Logger.log(`FATAL ERROR: ${error.message}`);
    SpreadsheetApp.getUi().alert(`FATAL ERROR: ${error.message}`);
  }
}

/**
 * Validate event data before creating
 */
function validateEventData(title, date, time, calendarId, email, rowNum) {
  let isValid = true;
  let errors = [];

  // Validate title
  if (!title || title.toString().trim() === "") {
    errors.push("Event Title (Column A) is empty");
    isValid = false;
  }

  // Validate date
  if (!isValidDate(date)) {
    errors.push("Start Date (Column B) is invalid");
    isValid = false;
  }

  // Validate time
  if (!isValidTimeFormat(time)) {
    errors.push("Start Time (Column C) is invalid");
    isValid = false;
  }

  // Validate calendar ID
  if (!calendarId || calendarId.toString().trim() === "") {
    errors.push("Calendar ID (Column E) is empty");
    isValid = false;
  }

  // Validate email
  if (!isValidEmail(email)) {
    errors.push("Guest Email (Column F) is invalid");
    isValid = false;
  }

  if (!isValid) {
    Logger.log(`Row ${rowNum} - Validation errors: ${errors.join(", ")}`);
  }

  return isValid;
}

/**
 * Create calendar event - FIXED VERSION with better time handling
 */
function createEvent(title, date, time, description, calendarId, guestEmail) {
  try {
    // Parse date
    const eventDate = new Date(date);
    
    // Check if date is valid
    if (isNaN(eventDate.getTime())) {
      Logger.log(`Invalid date: ${date}`);
      return null;
    }

    // Parse time - handle both time objects and strings
    const timeInfo = parseTimeFlexible(time);
    if (!timeInfo) {
      Logger.log(`Failed to parse time: ${time}`);
      return null;
    }

    const { hours, minutes } = timeInfo;
    eventDate.setHours(hours, minutes, 0, 0);

    Logger.log(`Parsed event: ${title} at ${eventDate}`);

    // Validate event date is in future
    if (eventDate < new Date()) {
      Logger.log(`WARNING: Event date is in the past: ${eventDate}`);
    }

    // Get calendar
    const calendar = CalendarApp.getCalendarById(calendarId.toString().trim());
    if (!calendar) {
      Logger.log(`Calendar not found: ${calendarId}`);
      return null;
    }

    // Create event (1 hour duration)
    const endDate = new Date(eventDate.getTime() + 60 * 60 * 1000);
    const event = calendar.createEvent(title, eventDate, endDate);

    // Set description
    if (description) {
      event.setDescription(description.toString());
    }

    // Add guest
    if (isValidEmail(guestEmail)) {
      event.addGuest(guestEmail.toString().trim());
    }

    // Add email notifications
    // 1 week before (7 days = 10080 minutes)
    event.addEmailNotification(7 * 24 * 60);
    
    // 1 day before (1440 minutes)
    event.addEmailNotification(24 * 60);

    return event.getId();

  } catch (error) {
    Logger.log(`Error creating calendar event: ${error.message}`);
    return null;
  }
}

/**
 * Parse time flexibly - handles multiple formats
 */
function parseTimeFlexible(timeInput) {
  if (!timeInput) {
    Logger.log("Time input is empty");
    return null;
  }

  // Convert to string and trim
  let timeStr = timeInput.toString().trim();
  
  Logger.log(`Parsing time: "${timeStr}" (type: ${typeof timeInput})`);

  // Handle if it's already a Date object (Google Sheets time)
  if (timeInput instanceof Date) {
    return {
      hours: timeInput.getHours(),
      minutes: timeInput.getMinutes()
    };
  }

  // Try to extract time if it's a Date object
  if (typeof timeInput === 'object' && timeInput !== null) {
    try {
      const hours = timeInput.getHours ? timeInput.getHours() : null;
      const minutes = timeInput.getMinutes ? timeInput.getMinutes() : null;
      if (hours !== null && minutes !== null && !isNaN(hours) && !isNaN(minutes)) {
        Logger.log(`Extracted from object: ${hours}:${minutes}`);
        return { hours, minutes };
      }
    } catch (e) {
      Logger.log(`Could not extract from object: ${e.message}`);
    }
  }

  // Try parsing as string
  timeStr = timeStr.toUpperCase();

  // Remove extra spaces
  timeStr = timeStr.replace(/\s+/g, ' ');

  // Pattern 1: "HH:MM AM/PM" or "H:MM AM/PM"
  const ampmMatch = timeStr.match(/^(\d{1,2}):(\d{2})\s?(AM|PM)$/);
  if (ampmMatch) {
    let hours = parseInt(ampmMatch[1]);
    const minutes = parseInt(ampmMatch[2]);
    const period = ampmMatch[3];

    // Convert to 24-hour format
    if (period === 'PM' && hours !== 12) {
      hours += 12;
    }
    if (period === 'AM' && hours === 12) {
      hours = 0;
    }

    if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
      Logger.log(`Parsed AM/PM: ${hours}:${minutes}`);
      return { hours, minutes };
    }
  }

  // Pattern 2: "HH:MM" or "H:MM" (24-hour format)
  const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    const hours = parseInt(timeMatch[1]);
    const minutes = parseInt(timeMatch[2]);

    if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
      Logger.log(`Parsed 24-hour: ${hours}:${minutes}`);
      return { hours, minutes };
    }
  }

  // Pattern 3: Try decimal format (sometimes Google Sheets uses decimals)
  const decimalMatch = timeStr.match(/^(\d+\.?\d*)$/);
  if (decimalMatch) {
    const decimalTime = parseFloat(decimalMatch[1]);
    if (decimalTime >= 0 && decimalTime <= 24) {
      const hours = Math.floor(decimalTime);
      const minutes = Math.round((decimalTime - hours) * 60);
      if (minutes >= 0 && minutes <= 59) {
        Logger.log(`Parsed decimal: ${hours}:${minutes}`);
        return { hours, minutes };
      }
    }
  }

  Logger.log(`Could not parse time format: "${timeStr}"`);
  return null;
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailStr = email.toString().trim();
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(emailStr);
}

/**
 * Validate date format - more flexible
 */
function isValidDate(date) {
  if (!date) return false;
  
  // Try parsing as Date
  const parsed = new Date(date);
  
  // Check if valid date
  if (!isNaN(parsed.getTime())) {
    return true;
  }
  
  return false;
}

/**
 * Validate time format - now with multiple support
 */
function isValidTimeFormat(time) {
  const timeInfo = parseTimeFlexible(time);
  return timeInfo !== null;
}

/**
 * Send confirmation email to guest
 */
function sendConfirmationEmail(guestEmail, title, date, time, description, calendarId) {
  if (!isValidEmail(guestEmail)) {
    Logger.log(`Invalid email address: ${guestEmail}`);
    return false;
  }

  try {
    const subject = `Calendar Event Created: ${title}`;
    
    const dateObj = new Date(date);
    const dateStr = dateObj.toLocaleDateString('en-US', { 
      weekday: 'long', 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    });

    const body = `
Hello,

A calendar event has been created for you.

EVENT DETAILS
═══════════════════════════════════════
Event Title:    ${title}
Date:           ${dateStr}
Time:           ${time}
Description:    ${description || "N/A"}
Calendar ID:    ${calendarId}
═══════════════════════════════════════

NOTIFICATIONS
You will receive email reminders:
• 1 week before the scheduled date
• 1 day before the scheduled date

Please check your calendar for the event details.

Best regards,
Faculty Calendar Automation System
    `;

    GmailApp.sendEmail(guestEmail.toString().trim(), subject, body);
    return true;

  } catch (error) {
    Logger.log(`Error sending email: ${error.message}`);
    return false;
  }
}

/**
 * Create menu for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 Faculty Calendar')
    .addItem('Create Faculty Events', 'createFacultyCalendarEvents')
    .addItem('View Logs', 'viewLogs')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * View execution logs
 */
function viewLogs() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Logs are available in:\nExtensions → Apps Script → View execution logs');
}

/**
 * Show help information
 */
function showHelp() {
  const help = `
📅 FACULTY CALENDAR EVENTS - HELP

PURPOSE:
Create calendar events from Faculty sheet and send emails to guests.

REQUIREMENTS:
• Sheet named "Faculty" with columns A-G
• Column A: Event Title (required)
• Column B: Start Date - any common format (required)
  Examples: 05/20/2024, 5/20/2024, May 20, 2024, 2024-05-20
• Column C: Start Time - multiple formats supported (required)
  Examples: 10:00 AM, 10:00 PM, 14:00, 2:00 PM, 10:00, 14
• Column D: Description (optional)
• Column E: Calendar ID - your@gmail.com (required)
• Column F: Guest Email (required)
• Column G: Status (auto-filled with "Done")

HOW TO USE:
1. Fill Faculty sheet with event data
2. Click "📅 Faculty Calendar" → "Create Faculty Events"
3. Wait for completion message
4. Check column G for "Done" status
5. Check emails sent to guests
6. Verify events in Google Calendar

SUPPORTED TIME FORMATS:
✓ 10:00 AM / 10:00 PM
✓ 10:00 (24-hour format)
✓ 2:00 PM
✓ 14:00
✓ Any time format Google Sheets supports

SUPPORTED DATE FORMATS:
✓ 05/20/2024
✓ 5/20/2024
✓ May 20, 2024
✓ 2024-05-20
✓ Any standard date format

NOTES:
• Rows with Status = "Done" will be skipped
• Email notifications: 1 week before + 1 day before
• Confirmation emails sent to guest
• Check execution logs for errors and debugging info

For detailed help, see the documentation.
  `;
  
  SpreadsheetApp.getUi().alert(help);
}