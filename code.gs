/**
 * Exam Scheduler - Google Apps Script Backend
 * This script handles the web app deployment and data management
 */

// Serve the HTML page
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Exam Scheduler")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Get or create the spreadsheet for storing exam data
function getSpreadsheet() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let spreadsheetId = scriptProperties.getProperty("SPREADSHEET_ID");

    Logger.log("Looking for spreadsheet ID: " + spreadsheetId);

    if (!spreadsheetId) {
      Logger.log("No spreadsheet ID found, creating new spreadsheet");
      // Create new spreadsheet
      const ss = SpreadsheetApp.create("Exam Scheduler Data");
      spreadsheetId = ss.getId();
      Logger.log("Created new spreadsheet with ID: " + spreadsheetId);

      scriptProperties.setProperty("SPREADSHEET_ID", spreadsheetId);

      // Set up headers
      const sheet = ss.getActiveSheet();
      sheet.setName("Exams");
      sheet
        .getRange("A1:H1")
        .setValues([
          [
            "ID",
            "Exam Name",
            "Date",
            "Start Time",
            "End Time",
            "Venue",
            "Class",
            "Laptops Needed",
          ],
        ]);
      sheet.getRange("A1:H1").setFontWeight("bold");
      sheet.setFrozenRows(1);

      Logger.log("Spreadsheet setup complete");
      return ss;
    }

    Logger.log("Opening existing spreadsheet");
    const ss = SpreadsheetApp.openById(spreadsheetId);
    Logger.log("Spreadsheet opened successfully");
    return ss;
  } catch (error) {
    Logger.log("ERROR in getSpreadsheet: " + error.toString());
    Logger.log("Error stack: " + error.stack);
    throw error;
  }
}

// Get the exams sheet
function getExamsSheet() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName("Exams");

    if (!sheet) {
      Logger.log("Exams sheet not found, using active sheet");
      sheet = ss.getActiveSheet();
      sheet.setName("Exams");
    }

    return sheet;
  } catch (error) {
    Logger.log("ERROR in getExamsSheet: " + error.toString());
    throw error;
  }
}

// Venue capacity configuration (seats per venue)
var VENUE_CAPACITIES = {
  "Learning Hive 1": 20,
  "Learning Hive 2": 20,
  "Learning Hive 3": 20,
  "Learning Hive 4": 20,
  Hall: 80,
  ISH: 60,
  "Com Lab 1": 40,
  "Com Lab 1A": 20,
  "Com Lab 1B": 20,
  "Com Lab 2": 40,
  "Com Lab 2A": 20,
  "Com Lab 2B": 20,
};

// Return venue capacity data
function getVenueCapacities() {
  return VENUE_CAPACITIES;
}

// Check availability of all venues for a given date/time/seats
function checkAvailability(date, startTime, endTime, seatsRequired) {
  try {
    var reqStart = parseTimeToMinutes(startTime);
    var reqEnd = parseTimeToMinutes(endTime);
    var seats = parseInt(seatsRequired) || 0;

    if (!date || reqStart === null || reqEnd === null) {
      return {
        success: false,
        message: "Please provide a valid date, start time, and end time.",
      };
    }

    if (reqEnd <= reqStart) {
      return { success: false, message: "End time must be after start time." };
    }

    if (seats < 1) {
      return { success: false, message: "Seats required must be at least 1." };
    }

    var allExams = getExams();
    var venueNames = Object.keys(VENUE_CAPACITIES);
    var venues = [];

    for (var v = 0; v < venueNames.length; v++) {
      var venueName = venueNames[v];
      var capacity = VENUE_CAPACITIES[venueName];
      var conflictingExams = [];

      for (var i = 0; i < allExams.length; i++) {
        var exam = allExams[i];
        if (exam.date === date && exam.venue === venueName) {
          var examStart = parseTimeToMinutes(exam.time);
          var examEnd = parseTimeToMinutes(exam.endTime);

          if (
            examStart !== null &&
            examEnd !== null &&
            reqStart < examEnd &&
            reqEnd > examStart
          ) {
            conflictingExams.push({
              examName: exam.name,
              time: exam.time,
              endTime: exam.endTime,
              class: exam.class || "",
            });
          }
        }
      }

      venues.push({
        name: venueName,
        capacity: capacity,
        available: conflictingExams.length === 0,
        hasEnoughSeats: capacity >= seats,
        conflictingExams: conflictingExams,
      });
    }

    return JSON.parse(
      JSON.stringify({
        success: true,
        data: { venues: venues },
      }),
    );
  } catch (error) {
    Logger.log("Error in checkAvailability: " + error.toString());
    return {
      success: false,
      message: "Error checking availability: " + error.toString(),
    };
  }
}

// Check for venue/date/time clashes using actual start and end times
function checkForClashes(newExam) {
  try {
    Logger.log("Checking for clashes...");
    Logger.log(
      "New exam: " +
        newExam.name +
        " at " +
        newExam.venue +
        " on " +
        newExam.date +
        " from " +
        newExam.time +
        " to " +
        newExam.endTime,
    );

    const exams = getExams();
    const clashes = [];

    // Parse new exam times
    const newExamStart = parseTimeToMinutes(newExam.time);
    const newExamEnd = parseTimeToMinutes(newExam.endTime);

    if (newExamStart === null || newExamEnd === null) {
      Logger.log("Could not parse new exam times");
      return {
        hasClash: false,
        clashes: [],
        error: "Invalid time format",
      };
    }

    for (let i = 0; i < exams.length; i++) {
      const existingExam = exams[i];

      // Only check same date AND same venue (venue-only conflict checking)
      if (
        existingExam.date === newExam.date &&
        existingExam.venue === newExam.venue
      ) {
        Logger.log(
          "Checking against: " +
            existingExam.name +
            " from " +
            existingExam.time +
            " to " +
            existingExam.endTime,
        );

        const existingExamStart = parseTimeToMinutes(existingExam.time);
        const existingExamEnd = parseTimeToMinutes(existingExam.endTime);

        if (existingExamStart === null || existingExamEnd === null) {
          Logger.log("Could not parse existing exam times, skipping");
          continue;
        }

        Logger.log("New exam: " + newExamStart + "-" + newExamEnd + " mins");
        Logger.log(
          "Existing exam: " +
            existingExamStart +
            "-" +
            existingExamEnd +
            " mins",
        );

        // Standard overlap check: (StartA < EndB) AND (EndA > StartB)
        // Note: exact end/start matches are OK (back-to-back allowed)
        if (newExamStart < existingExamEnd && newExamEnd > existingExamStart) {
          Logger.log("CLASH DETECTED - Time overlap!");
          clashes.push({
            examName: existingExam.name,
            venue: existingExam.venue,
            date: existingExam.date,
            time: existingExam.time,
            endTime: existingExam.endTime,
            class: existingExam.class,
            overlapType: getOverlapType(
              newExamStart,
              newExamEnd,
              existingExamStart,
              existingExamEnd,
            ),
          });
        } else {
          Logger.log("No overlap - times are separate");
        }
      }
    }

    if (clashes.length > 0) {
      Logger.log("Total clashes found: " + clashes.length);
      return {
        hasClash: true,
        clashes: clashes,
      };
    }

    Logger.log("No clashes found");
    return {
      hasClash: false,
      clashes: [],
    };
  } catch (error) {
    Logger.log("Error checking for clashes: " + error.toString());
    return {
      hasClash: false,
      clashes: [],
      error: error.toString(),
    };
  }
}

// Helper function to parse time string (HH:MM) to minutes since midnight
function parseTimeToMinutes(timeString) {
  try {
    if (!timeString) return null;

    const parts = timeString.split(":");
    if (parts.length !== 2) return null;

    const hours = parseInt(parts[0]);
    const minutes = parseInt(parts[1]);

    if (isNaN(hours) || isNaN(minutes)) return null;

    return hours * 60 + minutes;
  } catch (error) {
    Logger.log("Error parsing time: " + error.toString());
    return null;
  }
}

// Helper function to determine overlap type
function getOverlapType(newStart, newEnd, existingStart, existingEnd) {
  if (newStart === existingStart) {
    return "exact";
  } else if (newStart < existingEnd && newStart >= existingStart) {
    return "starts-during";
  } else if (newEnd > existingStart && newEnd <= existingEnd) {
    return "ends-during";
  } else if (newStart < existingStart && newEnd > existingEnd) {
    return "encompasses";
  } else {
    return "overlaps";
  }
}

// Save exam to spreadsheet
function saveExam(examData) {
  try {
    Logger.log("=== saveExam called ===");
    Logger.log("Exam data: " + JSON.stringify(examData));

    // Check for clashes first
    const clashCheck = checkForClashes(examData);

    if (clashCheck.hasClash) {
      Logger.log("Cannot save - clash detected!");

      // Build detailed error message
      let clashMessage = "VENUE CLASH DETECTED!\n\n";
      clashMessage +=
        'The venue "' + examData.venue + '" is already booked for:\n\n';

      for (let i = 0; i < clashCheck.clashes.length; i++) {
        const clash = clashCheck.clashes[i];
        clashMessage += "• " + clash.examName + "\n";
        clashMessage += "  Date: " + clash.date + "\n";
        clashMessage += "  Time: " + clash.time + " - " + clash.endTime + "\n";
        clashMessage += "  Class: " + clash.class + "\n\n";
      }

      clashMessage +=
        "Please choose a different venue, date, or time that doesn't overlap.";

      return {
        success: false,
        clash: true,
        message: clashMessage,
        clashes: clashCheck.clashes,
      };
    }

    // No clash - proceed to save
    Logger.log("No clash - proceeding to save");
    const sheet = getExamsSheet();

    // Generate ID based on timestamp
    const id = new Date().getTime();

    // Append the new exam
    sheet.appendRow([
      id,
      examData.name,
      examData.date,
      examData.time,
      examData.endTime,
      examData.venue,
      examData.class,
      examData.laptopsNeeded,
    ]);

    Logger.log("Exam saved successfully");
    return {
      success: true,
      clash: false,
      message: "Exam scheduled successfully!",
    };
  } catch (error) {
    Logger.log("Error in saveExam: " + error.toString());
    return {
      success: false,
      clash: false,
      message: "Error saving exam: " + error.toString(),
    };
  }
}

// Get all exams from spreadsheet - Returns array of exam objects
function getExams() {
  Logger.log("=== getExams function started ===");

  // Initialize empty array as default
  const exams = [];

  try {
    Logger.log("Step 1: Getting spreadsheet...");
    const ss = getSpreadsheet();

    if (!ss) {
      Logger.log("ERROR: Spreadsheet is null, returning empty array");
      return JSON.parse(JSON.stringify(exams)); // Ensure proper serialization
    }

    Logger.log("Step 2: Getting sheet...");
    const sheet = ss.getSheetByName("Exams") || ss.getActiveSheet();

    if (!sheet) {
      Logger.log("ERROR: Sheet is null, returning empty array");
      return JSON.parse(JSON.stringify(exams)); // Ensure proper serialization
    }

    Logger.log("Sheet name: " + sheet.getName());

    Logger.log("Step 3: Getting data...");
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    Logger.log("Total rows (including header): " + data.length);

    // Skip header row
    if (data.length <= 1) {
      Logger.log("No data rows found, returning empty array");
      return JSON.parse(JSON.stringify(exams)); // Ensure proper serialization
    }

    Logger.log("Step 4: Processing exams...");

    for (let i = 1; i < data.length; i++) {
      try {
        // Skip empty rows
        if (!data[i][0] && !data[i][1]) {
          Logger.log("Skipping empty row " + (i + 1));
          continue;
        }

        // Format date properly - ensure it's a string
        let dateValue = "";
        if (data[i][2]) {
          if (data[i][2] instanceof Date) {
            dateValue = Utilities.formatDate(
              data[i][2],
              Session.getScriptTimeZone(),
              "yyyy-MM-dd",
            );
          } else {
            dateValue = String(data[i][2]);
          }
        }

        // Format start time properly - handle both string and Date object
        let timeValue = "";
        if (data[i][3]) {
          if (data[i][3] instanceof Date) {
            // Google Sheets stores time as Date object, extract just the time
            timeValue = Utilities.formatDate(
              data[i][3],
              Session.getScriptTimeZone(),
              "HH:mm",
            );
          } else if (typeof data[i][3] === "string") {
            // Already a string, use as-is
            timeValue = data[i][3];
          } else {
            timeValue = String(data[i][3]);
          }
        }

        // Format end time properly - handle both string and Date object
        let endTimeValue = "";
        if (data[i][4]) {
          if (data[i][4] instanceof Date) {
            endTimeValue = Utilities.formatDate(
              data[i][4],
              Session.getScriptTimeZone(),
              "HH:mm",
            );
          } else if (typeof data[i][4] === "string") {
            endTimeValue = data[i][4];
          } else {
            endTimeValue = String(data[i][4]);
          }
        }

        // Create plain object (not a special type)
        // Column indices: 0=ID, 1=Name, 2=Date, 3=StartTime, 4=EndTime, 5=Venue, 6=Class, 7=Laptops
        const exam = {
          id: Number(data[i][0]) || Date.now() + i,
          name: String(data[i][1] || "Unnamed Exam"),
          date: dateValue,
          time: timeValue,
          endTime: endTimeValue,
          venue: String(data[i][5] || ""),
          class: String(data[i][6] || ""),
          laptopsNeeded: parseInt(data[i][7]) || 0,
        };

        Logger.log(
          "Row " +
            (i + 1) +
            ": " +
            exam.name +
            " from " +
            exam.time +
            " to " +
            exam.endTime,
        );
        exams.push(exam);
      } catch (rowError) {
        Logger.log(
          "Error processing row " + (i + 1) + ": " + rowError.toString(),
        );
        // Continue with next row
      }
    }

    Logger.log("Step 5: Completed - found " + exams.length + " exams");
    Logger.log("=== getExams function completed successfully ===");
  } catch (error) {
    Logger.log("=== CRITICAL ERROR in getExams ===");
    Logger.log("Error message: " + error.toString());
    Logger.log("Error stack: " + error.stack);
    Logger.log("=== Returning empty array ===");
  }

  // Ensure proper serialization - parse and stringify to create plain objects
  const result = JSON.parse(JSON.stringify(exams));
  Logger.log("Returning result with " + result.length + " items");
  return result;
}

// Delete an exam by ID
function deleteExam(examId) {
  try {
    const sheet = getExamsSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == examId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Exam deleted successfully!" };
      }
    }

    return { success: false, message: "Exam not found" };
  } catch (error) {
    return {
      success: false,
      message: "Error deleting exam: " + error.toString(),
    };
  }
}

// Delete all exams from the spreadsheet
function deleteAllExams() {
  try {
    const sheet = getExamsSheet();
    const lastRow = sheet.getLastRow();

    // Check if there are any exams to delete (lastRow = 1 means only header)
    if (lastRow <= 1) {
      return {
        success: false,
        message: "No exams to delete",
      };
    }

    // Calculate number of rows to delete (all rows except header)
    const numRowsToDelete = lastRow - 1;

    // Efficient batch deletion: delete from row 2 onwards
    sheet.deleteRows(2, numRowsToDelete);

    return {
      success: true,
      message:
        "All exams deleted successfully! (" +
        numRowsToDelete +
        " exam" +
        (numRowsToDelete !== 1 ? "s" : "") +
        " removed)",
    };
  } catch (error) {
    return {
      success: false,
      message: "Error deleting exams: " + error.toString(),
    };
  }
}

// Get the spreadsheet URL for users to view
function getSpreadsheetUrl() {
  const ss = getSpreadsheet();
  return ss.getUrl();
}

// Test function - run this manually to check if everything works
function testGetExams() {
  Logger.log("=== TESTING getExams ===");

  try {
    const exams = getExams();
    Logger.log("SUCCESS! Got " + exams.length + " exams");
    Logger.log("Exams data: " + JSON.stringify(exams));
    return { success: true, count: exams.length, data: exams };
  } catch (error) {
    Logger.log("FAILED! Error: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// Initialize function - run this first to set up everything
function initializeApp() {
  Logger.log("=== INITIALIZING APP ===");

  try {
    Logger.log("Step 1: Creating/getting spreadsheet...");
    const ss = getSpreadsheet();
    Logger.log("Spreadsheet created/found: " + ss.getName());
    Logger.log("Spreadsheet URL: " + ss.getUrl());

    Logger.log("Step 2: Checking sheet...");
    const sheet = getExamsSheet();
    Logger.log("Sheet found: " + sheet.getName());

    Logger.log("Step 3: Testing data retrieval...");
    const exams = getExams();
    Logger.log("Retrieved " + exams.length + " exams");

    Logger.log("=== INITIALIZATION COMPLETE ===");
    Logger.log("Your app is ready to use!");
    Logger.log("Spreadsheet URL: " + ss.getUrl());

    return {
      success: true,
      message: "Initialization successful",
      spreadsheetUrl: ss.getUrl(),
      examCount: exams.length,
    };
  } catch (error) {
    Logger.log("=== INITIALIZATION FAILED ===");
    Logger.log("Error: " + error.toString());
    Logger.log("Stack: " + error.stack);
    return {
      success: false,
      error: error.toString(),
    };
  }
}

// Validate bulk exams for clashes and errors
// Input: Array of {rowIndex, name, date, time, endTime, venue, class, laptopsNeeded}
// Returns: {results: [{rowIndex, valid, errors[], clashesWithExisting[], clashesWithBatch[]}], summary: {...}}
function validateBulkExams(examDataArray) {
  Logger.log(
    "=== validateBulkExams called with " + examDataArray.length + " exams ===",
  );

  try {
    // Fetch existing exams once
    const existingExams = getExams();
    Logger.log("Fetched " + existingExams.length + " existing exams");

    const results = [];
    const summary = {
      totalRows: examDataArray.length,
      validRows: 0,
      invalidRows: 0,
      clashingRows: 0,
    };

    // Process each exam in the batch
    for (let i = 0; i < examDataArray.length; i++) {
      const exam = examDataArray[i];
      const result = {
        rowIndex: exam.rowIndex,
        valid: true,
        errors: [],
        clashesWithExisting: [],
        clashesWithBatch: [],
      };

      // Validate required fields
      if (!exam.name || exam.name.trim() === "") {
        result.errors.push("Subject is required");
        result.valid = false;
      }
      if (!exam.date || exam.date.trim() === "") {
        result.errors.push("Date is required");
        result.valid = false;
      }
      if (!exam.time || exam.time.trim() === "") {
        result.errors.push("Start time is required");
        result.valid = false;
      }
      if (!exam.endTime || exam.endTime.trim() === "") {
        result.errors.push("End time is required");
        result.valid = false;
      }
      if (!exam.venue || exam.venue.trim() === "") {
        result.errors.push("Venue is required");
        result.valid = false;
      }
      if (!exam.class || exam.class.trim() === "") {
        result.errors.push("Teaching group is required");
        result.valid = false;
      }

      // Validate time format and logic
      const startMinutes = parseTimeToMinutes(exam.time);
      const endMinutes = parseTimeToMinutes(exam.endTime);

      if (startMinutes === null && exam.time) {
        result.errors.push("Invalid start time format");
        result.valid = false;
      }
      if (endMinutes === null && exam.endTime) {
        result.errors.push("Invalid end time format");
        result.valid = false;
      }
      if (
        startMinutes !== null &&
        endMinutes !== null &&
        endMinutes <= startMinutes
      ) {
        result.errors.push("End time must be after start time");
        result.valid = false;
      }

      // Skip clash checking if basic validation failed
      if (!result.valid) {
        summary.invalidRows++;
        results.push(result);
        continue;
      }

      // Check clashes against existing exams
      for (let j = 0; j < existingExams.length; j++) {
        const existing = existingExams[j];

        if (existing.date === exam.date && existing.venue === exam.venue) {
          const existingStart = parseTimeToMinutes(existing.time);
          const existingEnd = parseTimeToMinutes(existing.endTime);

          if (existingStart !== null && existingEnd !== null) {
            // Check for overlap: (StartA < EndB) AND (EndA > StartB)
            if (startMinutes < existingEnd && endMinutes > existingStart) {
              result.clashesWithExisting.push({
                examName: existing.name,
                venue: existing.venue,
                date: existing.date,
                time: existing.time,
                endTime: existing.endTime,
                class: existing.class,
              });
            }
          }
        }
      }

      // Check clashes within the batch itself
      for (let k = 0; k < i; k++) {
        const otherExam = examDataArray[k];

        if (otherExam.date === exam.date && otherExam.venue === exam.venue) {
          const otherStart = parseTimeToMinutes(otherExam.time);
          const otherEnd = parseTimeToMinutes(otherExam.endTime);

          if (otherStart !== null && otherEnd !== null) {
            if (startMinutes < otherEnd && endMinutes > otherStart) {
              result.clashesWithBatch.push({
                rowIndex: otherExam.rowIndex,
                examName: otherExam.name,
                venue: otherExam.venue,
                date: otherExam.date,
                time: otherExam.time,
                endTime: otherExam.endTime,
                class: otherExam.class,
              });

              // Also mark the other exam as clashing with this one
              const otherResult = results.find(
                (r) => r.rowIndex === otherExam.rowIndex,
              );
              if (otherResult && otherResult.valid) {
                otherResult.clashesWithBatch.push({
                  rowIndex: exam.rowIndex,
                  examName: exam.name,
                  venue: exam.venue,
                  date: exam.date,
                  time: exam.time,
                  endTime: exam.endTime,
                  class: exam.class,
                });
              }
            }
          }
        }
      }

      // Mark as invalid if there are clashes
      if (
        result.clashesWithExisting.length > 0 ||
        result.clashesWithBatch.length > 0
      ) {
        result.valid = false;
        summary.clashingRows++;
      } else {
        summary.validRows++;
      }

      results.push(result);
    }

    // Recalculate valid count after batch clash updates
    summary.validRows = results.filter((r) => r.valid).length;
    summary.clashingRows = results.filter(
      (r) => r.clashesWithExisting.length > 0 || r.clashesWithBatch.length > 0,
    ).length;
    summary.invalidRows = results.filter((r) => r.errors.length > 0).length;

    Logger.log(
      "Validation complete. Valid: " +
        summary.validRows +
        ", Invalid: " +
        summary.invalidRows +
        ", Clashing: " +
        summary.clashingRows,
    );

    return {
      results: results,
      summary: summary,
    };
  } catch (error) {
    Logger.log("Error in validateBulkExams: " + error.toString());
    return {
      results: [],
      summary: { totalRows: 0, validRows: 0, invalidRows: 0, clashingRows: 0 },
      error: error.toString(),
    };
  }
}

// Save multiple exams in bulk
// Input: Array of validated exam objects
// Returns: {success, savedCount, failedCount, errors[]}
function saveBulkExams(examDataArray) {
  Logger.log(
    "=== saveBulkExams called with " + examDataArray.length + " exams ===",
  );

  try {
    const sheet = getExamsSheet();
    const baseTimestamp = new Date().getTime();
    const errors = [];
    let savedCount = 0;

    // Prepare all rows for batch insert
    const rowsToInsert = [];

    for (let i = 0; i < examDataArray.length; i++) {
      const exam = examDataArray[i];

      // Generate unique ID (timestamp + index to ensure uniqueness)
      const id = baseTimestamp + i;

      rowsToInsert.push([
        id,
        exam.name,
        exam.date,
        exam.time,
        exam.endTime,
        exam.venue,
        exam.class,
        exam.laptopsNeeded || 0,
      ]);
    }

    // Batch insert all rows at once for efficiency
    if (rowsToInsert.length > 0) {
      const lastRow = sheet.getLastRow();
      const startRow = lastRow + 1;
      const numRows = rowsToInsert.length;
      const numCols = 8;

      sheet.getRange(startRow, 1, numRows, numCols).setValues(rowsToInsert);
      savedCount = numRows;

      Logger.log("Successfully saved " + savedCount + " exams");
    }

    return {
      success: true,
      savedCount: savedCount,
      failedCount: errors.length,
      errors: errors,
    };
  } catch (error) {
    Logger.log("Error in saveBulkExams: " + error.toString());
    return {
      success: false,
      savedCount: 0,
      failedCount: examDataArray.length,
      errors: [error.toString()],
    };
  }
}

// Migration function - run once to add End Time column to existing data
function migrateToEndTimeSchema() {
  Logger.log("=== MIGRATION: Adding End Time column ===");

  try {
    const sheet = getExamsSheet();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    // Check if migration is needed by looking at header
    const headers = data[0];
    if (headers.length >= 8 && headers[4] === "End Time") {
      Logger.log("Migration already complete - End Time column exists");
      return { success: true, message: "Already migrated" };
    }

    // Insert new column E for End Time
    sheet.insertColumnAfter(4); // Insert after column D (Start Time)

    // Update header
    sheet.getRange("E1").setValue("End Time");
    sheet.getRange("A1:H1").setFontWeight("bold");

    // Calculate default end times for existing data (start time + 3 hours)
    const DEFAULT_DURATION_MINUTES = 180;

    for (let i = 2; i <= data.length; i++) {
      const startTimeCell = sheet.getRange(i, 4).getValue(); // Column D

      if (startTimeCell) {
        let startTimeMinutes = null;

        if (startTimeCell instanceof Date) {
          startTimeMinutes =
            startTimeCell.getHours() * 60 + startTimeCell.getMinutes();
        } else if (typeof startTimeCell === "string") {
          startTimeMinutes = parseTimeToMinutes(startTimeCell);
        }

        if (startTimeMinutes !== null) {
          const endTimeMinutes = startTimeMinutes + DEFAULT_DURATION_MINUTES;
          const endHours = Math.floor(endTimeMinutes / 60) % 24;
          const endMins = endTimeMinutes % 60;
          const endTimeString =
            String(endHours).padStart(2, "0") +
            ":" +
            String(endMins).padStart(2, "0");

          sheet.getRange(i, 5).setValue(endTimeString); // Column E
        }
      }
    }

    Logger.log("Migration complete!");
    return {
      success: true,
      message: "Migration successful - End Time column added",
    };
  } catch (error) {
    Logger.log("Migration error: " + error.toString());
    return { success: false, error: error.toString() };
  }
}
