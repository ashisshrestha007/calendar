/**
 * APP SCRIPT 2: RESIDENT PRESENTATIONS DISTRIBUTION
 * Purpose: Calculate student distribution across presentations and auto-schedule
 * 
 * SHEET STRUCTURE:
 * Faculty Sheet:
 *   Column A: Event Title
 *   Column B: Start Date (MM/DD/YYYY)
 *   Column C: Start Time (HH:MM AM/PM)
 *   Column D: Description (must contain "Resident Presentation")
 *   Column E: Calendar ID
 * 
 * StudentEmail Sheet:
 *   Column A: Student email addresses (one per row, starting row 2)
 *   NOTE: For your use case, this contains RESIDENT emails (81 residents)
 * 
 * The script will distribute all StudentEmail entries (residents) 
 * sequentially across the presentations
 */

function scheduleResidentPresentations() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get sheets
    const facultySheet = spreadsheet.getSheetByName("Faculty");
    const studentEmailSheet = spreadsheet.getSheetByName("StudentEmail");

    // Validate sheets exist
    if (!facultySheet) {
      const msg = "ERROR: Faculty sheet not found!";
      Logger.log(msg);
      SpreadsheetApp.getUi().alert(msg);
      return;
    }

    if (!studentEmailSheet) {
      const msg = "ERROR: StudentEmail sheet not found!";
      Logger.log(msg);
      SpreadsheetApp.getUi().alert(msg);
      return;
    }

    Logger.log("=== RESIDENT PRESENTATION SCHEDULER ===");
    Logger.log("Using StudentEmail sheet for distribution (contains resident emails)");

    // Step 1: Calculate distribution
    const distribution = calculateDistribution(facultySheet, studentEmailSheet);
    
    if (distribution.X === 0) {
      const msg = `No "Resident Presentation" entries found in Faculty column D.\n\nCheck that column D contains exactly "Resident Presentation" text.`;
      Logger.log(msg);
      SpreadsheetApp.getUi().alert(msg);
      return;
    }

    if (distribution.Y === 0) {
      const msg = "No emails found in StudentEmail sheet column A.";
      Logger.log(msg);
      SpreadsheetApp.getUi().alert(msg);
      return;
    }

    Logger.log(`\n=== DISTRIBUTION CALCULATED ===`);
    Logger.log(`X (Presentations) = ${distribution.X}`);
    Logger.log(`Y (Emails in StudentEmail) = ${distribution.Y}`);
    Logger.log(`Z (Per Presentation) = ${distribution.Z}`);

    // Step 2: Schedule presentations
    let scheduledCount = 0;
    const facultyData = facultySheet.getDataRange().getValues();
    let studentIndex = 0;

    Logger.log(`\n=== SCHEDULING PRESENTATIONS ===`);

    for (let i = 1; i < facultyData.length; i++) {
      const row = facultyData[i];
      const description = row[3];  // Column D

      // Check if this is a resident presentation
      if (!description || !description.toString().includes("Resident Presentation")) {
        continue;
      }

      try {
        // Extract presentation data
        const title = row[0];          // Column A
        const startDate = row[1];      // Column B
        const startTime = row[2];      // Column C
        const calendarId = row[4];     // Column E

        // Validate data
        if (!title || !startDate || !startTime || !calendarId) {
          Logger.log(`Row ${i + 1}: Skipped - missing required data`);
          continue;
        }

        // Get assigned students for this presentation
        const assignedStudents = getAssignedStudents(
          distribution.students,
          studentIndex,
          distribution.Z
        );
        studentIndex += distribution.Z;

        // Create presentation event
        createPresentationEvent(
          title,
          startDate,
          startTime,
          description,
          calendarId,
          assignedStudents
        );

        scheduledCount++;
        Logger.log(`✓ Row ${i + 1}: Presentation "${title}" scheduled with ${assignedStudents.length} attendees`);

      } catch (error) {
        Logger.log(`ERROR in row ${i + 1}: ${error.message}`);
      }
    }

    Logger.log(`\n=== SCHEDULING COMPLETE ===`);
    Logger.log(`Total presentations scheduled: ${scheduledCount}`);

    // Show summary
    const message = `Presentations Scheduled: ${scheduledCount}\n\nDistribution:\nPresentations (X): ${distribution.X}\nAttendees (Y): ${distribution.Y}\nPer Presentation (Z): ${distribution.Z}`;
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    Logger.log(`FATAL ERROR: ${error.message}`);
    SpreadsheetApp.getUi().alert(`FATAL ERROR: ${error.message}`);
  }
}

/**
 * STEP 1: Calculate distribution
 * X = count of "Resident Presentation" in Faculty column D
 * Y = count of emails in StudentEmail column A
 * Z = Y / X (rounded to integer)
 */
function calculateDistribution(facultySheet, studentEmailSheet) {
  Logger.log("\n=== STEP 1: COUNTING PRESENTATIONS ===");
  
  // Count presentations (X)
  const facultyData = facultySheet.getDataRange().getValues();
  let presentationCount = 0;

  for (let i = 1; i < facultyData.length; i++) {
    const description = facultyData[i][3];  // Column D
    
    if (description && description.toString().includes("Resident Presentation")) {
      presentationCount++;
      Logger.log(`Row ${i + 1}: "Resident Presentation" found`);
    }
  }

  Logger.log(`Total presentations found: ${presentationCount}`);

  Logger.log("\n=== STEP 2: COUNTING ATTENDEES (StudentEmail) ===");
  
  // Count student/attendee emails (Y)
  const studentData = studentEmailSheet.getDataRange().getValues();
  let studentCount = 0;
  const studentEmails = [];

  for (let i = 1; i < studentData.length; i++) {
    const email = studentData[i][0];  // Column A
    
    if (email && email.toString().trim() !== "") {
      const emailStr = email.toString().trim();
      
      // Validate email format
      if (isValidEmail(emailStr)) {
        studentEmails.push(emailStr);
        studentCount++;
        Logger.log(`Row ${i + 1}: ${emailStr}`);
      } else {
        Logger.log(`Row ${i + 1}: INVALID FORMAT - ${emailStr}`);
      }
    }
  }

  Logger.log(`Total valid emails found: ${studentCount}`);

  Logger.log("\n=== STEP 3: CALCULATE DISTRIBUTION ===");
  
  // Calculate distribution (Z = Y / X)
  const Z = presentationCount > 0 
    ? Math.round(studentCount / presentationCount)
    : 0;

  Logger.log(`Formula: Z = Y ÷ X`);
  Logger.log(`Z = ${studentCount} ÷ ${presentationCount} = ${Z}`);

  return {
    X: presentationCount,
    Y: studentCount,
    Z: Z,
    students: studentEmails
  };
}

/**
 * STEP 2: Get Z assigned students starting from studentIndex
 */
function getAssignedStudents(allStudents, startIndex, count) {
  const assigned = [];
  
  for (let i = 0; i < count && (startIndex + i) < allStudents.length; i++) {
    assigned.push(allStudents[startIndex + i]);
  }
  
  return assigned;
}

/**
 * STEP 3: Create calendar event for presentation
 */
function createPresentationEvent(title, date, time, description, calendarId, studentEmails) {
  try {
    // Parse date and time
    const eventDate = new Date(date);
    const timeInfo = parseTimeFlexible(time);
    
    if (!timeInfo) {
      Logger.log(`Failed to parse time: ${time}`);
      return;
    }

    const { hours, minutes } = timeInfo;
    eventDate.setHours(hours, minutes, 0, 0);

    // Get calendar
    const calendar = CalendarApp.getCalendarById(calendarId.toString().trim());
    if (!calendar) {
      Logger.log(`Calendar not found: ${calendarId}`);
      return;
    }

    // Create event (1 hour duration)
    const endDate = new Date(eventDate.getTime() + 60 * 60 * 1000);
    const event = calendar.createEvent(title, eventDate, endDate);

    // Set description with attendee details
    const attendeeInfo = `
PRESENTATION DETAILS
════════════════════════════════════════

${description}

ATTENDEES (${studentEmails.length}):
${studentEmails.map((email, idx) => `${idx + 1}. ${email}`).join('\n')}

════════════════════════════════════════
    `;

    event.setDescription(attendeeInfo);

    // Add student/attendee guests
    for (const email of studentEmails) {
      if (isValidEmail(email)) {
        event.addGuest(email);
      }
    }

    // Add email notifications
    // 1 week before (10080 minutes)
    event.addEmailNotification(7 * 24 * 60);
    
    // 1 day before (1440 minutes)
    event.addEmailNotification(24 * 60);

    Logger.log(`  Event created successfully`);

  } catch (error) {
    Logger.log(`Error creating presentation event: ${error.message}`);
  }
}

/**
 * Parse time flexibly - handles multiple formats
 */
function parseTimeFlexible(timeInput) {
  if (!timeInput) {
    return null;
  }

  let timeStr = timeInput.toString().trim();

  // Handle if it's already a Date object
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
        return { hours, minutes };
      }
    } catch (e) {
      // Continue to string parsing
    }
  }

  // Try parsing as string
  timeStr = timeStr.toUpperCase();
  timeStr = timeStr.replace(/\s+/g, ' ');

  // Pattern 1: "HH:MM AM/PM" or "H:MM AM/PM"
  const ampmMatch = timeStr.match(/^(\d{1,2}):(\d{2})\s?(AM|PM)$/);
  if (ampmMatch) {
    let hours = parseInt(ampmMatch[1]);
    const minutes = parseInt(ampmMatch[2]);
    const period = ampmMatch[3];

    if (period === 'PM' && hours !== 12) {
      hours += 12;
    }
    if (period === 'AM' && hours === 12) {
      hours = 0;
    }

    if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
      return { hours, minutes };
    }
  }

  // Pattern 2: "HH:MM" or "H:MM" (24-hour format)
  const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    const hours = parseInt(timeMatch[1]);
    const minutes = parseInt(timeMatch[2]);

    if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
      return { hours, minutes };
    }
  }

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
 * View distribution report
 */
function viewDistributionReport() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const facultySheet = spreadsheet.getSheetByName("Faculty");
    const studentEmailSheet = spreadsheet.getSheetByName("StudentEmail");

    if (!facultySheet || !studentEmailSheet) {
      SpreadsheetApp.getUi().alert("Required sheets not found!");
      return;
    }

    const distribution = calculateDistribution(facultySheet, studentEmailSheet);

    const report = `
PRESENTATION DISTRIBUTION REPORT
════════════════════════════════════════════════════════

Total Presentations (X):          ${distribution.X}
Total Attendees in StudentEmail (Y): ${distribution.Y}
Per Presentation (Z):             ${distribution.Z}

FORMULA:
Z = Y ÷ X = ${distribution.Y} ÷ ${distribution.X} = ${distribution.Z}

ATTENDEE EMAIL LIST:
${distribution.students.map((email, idx) => `${idx + 1}. ${email}`).join('\n')}

DISTRIBUTION BREAKDOWN:
${Array.from({length: distribution.X}, (_, i) => {
  const start = i * distribution.Z;
  const end = Math.min(start + distribution.Z, distribution.Y);
  const attendees = distribution.students.slice(start, end);
  return `Presentation ${i + 1} (${attendees.length} attendees): ${attendees.slice(0, 3).map(e => e.split('@')[0]).join(', ')}${attendees.length > 3 ? '...' : ''}`;
}).join('\n')}

════════════════════════════════════════════════════════
    `;

    SpreadsheetApp.getUi().alert(report);

  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

/**
 * Create menu for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Resident Presentations')
    .addItem('Schedule Presentations', 'scheduleResidentPresentations')
    .addItem('View Distribution Report', 'viewDistributionReport')
    .addSeparator()
    .addItem('View Logs', 'viewLogs')
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
🎓 PRESENTATION DISTRIBUTION - HELP

PURPOSE:
Calculate fair distribution of attendees (from StudentEmail) across presentations.

REQUIREMENTS:
Two sheets with proper structure:

1. Faculty Sheet:
   • Column A: Event Title
   • Column B: Start Date (MM/DD/YYYY)
   • Column C: Start Time (HH:MM AM/PM or 24-hour format)
   • Column D: Description (must contain "Resident Presentation")
   • Column E: Calendar ID - your@gmail.com

2. StudentEmail Sheet:
   • Column A: Email addresses to distribute (one per row, starting row 2)
   • NOTE: For your use case, this contains your resident emails

CALCULATION:
X = Count of "Resident Presentation" in Faculty column D
Y = Count of valid emails in StudentEmail column A
Z = Y ÷ X (rounded to nearest integer)

HOW TO USE:
1. Fill Faculty sheet with presentations
   (Mark column D with "Resident Presentation")
2. Ensure StudentEmail sheet has all attendee emails
3. Click "🎓 Resident Presentations" → "Schedule Presentations"
4. Wait for completion message
5. Check the distribution report
6. Verify events in Google Calendar

DISTRIBUTION EXAMPLE:
If Faculty has 3 presentations and StudentEmail has 81 emails:
X = 3, Y = 81, Z = 27
Result: Each presentation gets 27 attendees

NOTES:
• Attendees assigned sequentially
• Notifications: 1 week before + 1 day before
• Check execution logs for detailed info
• View Distribution Report to see assignments

For detailed help, see the documentation.
  `;
  
  SpreadsheetApp.getUi().alert(help);
}