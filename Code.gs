function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Cathedral High School Check-In')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function resolveStudent(inputStr, cache, studentDataOrGetter) {
  let cacheKey = "student_" + inputStr.toLowerCase();
  let cachedData = cache.get(cacheKey);

  if (cachedData) {
    const parsed = JSON.parse(cachedData);
    return {
      found: true,
      id: parsed.id,
      name: parsed.name,
      email: parsed.email || ""
    };
  }

  let studentDataRange = typeof studentDataOrGetter === 'function' ? studentDataOrGetter() : studentDataOrGetter;

  if (!studentDataRange._lookupMapExactId) {
    const mapExactId = new Map();
    const mapLower = new Map();
    for (let i = 0; i < studentDataRange.length; i++) {
      let rowId = String(studentDataRange[i][0]).trim();
      let rowName = String(studentDataRange[i][1]).trim();
      let rowEmail = String(studentDataRange[i][2]).trim();

      const data = { id: rowId, name: rowName, email: rowEmail };

      if (rowId && !mapExactId.has(rowId)) mapExactId.set(rowId, data);

      let lowerName = rowName.toLowerCase();
      let lowerEmail = rowEmail.toLowerCase();
      if (lowerName && !mapLower.has(lowerName)) mapLower.set(lowerName, data);
      if (lowerEmail && !mapLower.has(lowerEmail)) mapLower.set(lowerEmail, data);
    }
    Object.defineProperty(studentDataRange, '_lookupMapExactId', { value: mapExactId, enumerable: false });
    Object.defineProperty(studentDataRange, '_lookupMapLower', { value: mapLower, enumerable: false });
  }

  let match = studentDataRange._lookupMapExactId.get(inputStr);
  if (!match) {
    match = studentDataRange._lookupMapLower.get(inputStr.toLowerCase());
  }

  if (match) {
    cache.put(cacheKey, JSON.stringify(match), 21600);
    return {
      found: true,
      id: match.id,
      name: match.name,
      email: match.email
    };
  }

  return {
    found: false,
    id: "Manual/Unknown",
    name: inputStr,
    email: ""
  };
}

function processCheckIn(location, studentInput, manualTimeStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Log');

  if (!logSheet) {
    throw new Error("Make sure your tab is named exactly 'Log'.");
  }

  const userEmail = Session.getActiveUser().getEmail();

  // 1. Resolve Student Name and ID
  let inputStr = String(studentInput).trim();
  const cache = CacheService.getScriptCache();

  let studentResult = resolveStudent(inputStr, cache, () => {
    const studentSheet = ss.getSheetByName('Students');
    if (!studentSheet) {
      throw new Error("Make sure your tab is named exactly 'Students'.");
    }
    const lastRow = studentSheet.getLastRow();
    return lastRow > 1 ? studentSheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  });

  if (!studentResult.found) {
    throw new Error("Student Not Found. Please check the spelling or ID and try again.");
  }

  let studentName = studentResult.name;
  let studentId = studentResult.id;
  let studentEmail = studentResult.email;

  const lock = LockService.getScriptLock();
  // Wait for up to 30000 milliseconds for other processes to finish.
  if (!lock.tryLock(30000)) {
    throw new Error("System is currently busy due to high traffic. Please try again in a few seconds.");
  }

  try {

    // 2. Check for an active session within the last 1 hour
    const actualNow = new Date();
    const now = manualTimeStr ? new Date(manualTimeStr) : actualNow;
    const oneHourMs = 60 * 60 * 1000;
    const lastRow = logSheet.getLastRow();

    // Only look at the last 1000 rows to speed up check-ins/check-outs significantly.
    // If the sheet has fewer than 1000 rows, it stops at row 2 (header is row 1).
    const startRow = Math.max(2, lastRow - 999);
    const numRows = lastRow - startRow + 1;

    if (numRows > 0) {
      // getRange(row, column, numRows, numColumns)
      const logData = logSheet.getRange(startRow, 1, numRows, 8).getValues();

      for (let i = logData.length - 1; i >= 0; i--) {
        let row = logData[i];
        let checkOutTime = row[5];
        if (checkOutTime) continue;

        let rowId = String(row[2]).trim();
        if (rowId !== studentId) continue;

        let rowLocation = String(row[1]).trim();
        if (rowLocation !== location) continue;

        let checkInUser = String(row[7]).trim();
        if (checkInUser !== String(userEmail).trim()) continue;

        let checkInTime = new Date(row[0]);
        let timeDiffMs = now.getTime() - checkInTime.getTime();

        if (timeDiffMs >= 0 && timeDiffMs <= oneHourMs) {
          let durationMins = Math.round(timeDiffMs / 60000);
          const actualRowToUpdate = startRow + i;
          logSheet.getRange(actualRowToUpdate, 6).setValue(now);
          logSheet.getRange(actualRowToUpdate, 7).setValue(durationMins);
          logSheet.getRange(actualRowToUpdate, 8).setValue(userEmail);

          if (manualTimeStr) {
            logSheet.getRange(actualRowToUpdate, 10).setValue(actualNow); // Column I: Actual Check-Out Time
          }
          return { name: studentName, status: "out", time: durationMins };
        }
      }
    }

    // 3. Log a new check-in
    let rowData = [now, sanitizeForSheets(location), sanitizeForSheets(studentId), sanitizeForSheets(studentName), sanitizeForSheets(studentEmail), "", "", sanitizeForSheets(userEmail)];
    if (manualTimeStr) {
      rowData[8] = actualNow; // Column I: Actual Check-In Time
      rowData[9] = "";        // Column J: Actual Check-Out Time
    }
    logSheet.appendRow(rowData);
    return { name: studentName, status: "in" };
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/**
 * Sanitizes input to prevent formula/CSV injection when appending data to Google Sheets.
 * Escapes values that begin with =, +, -, or @ by prepending a single quote.
 * @param {*} value - The input to sanitize.
 * @returns {string|*} - The sanitized string, or the original value if not a string/number.
 */
function sanitizeForSheets(value) {
  if (value === null || value === undefined) {
    return "";
  }

  const strValue = String(value);
  if (/^[=+\-@]/.test(strValue)) {
    return "'" + strValue;
  }

  return strValue;
}

// NEW: Fetches all setup data in one fast call
function getSetupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Helper function to extract a specific column from a sheet
  const getColumnData = (sheetName, colIndex) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // Only header or empty
    const data = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
    const list = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) list.push(String(data[i][0]).trim());
    }
    return list;
  };

  // Cache all students in the background and build the list of names
  const studentSheet = ss.getSheetByName('Students');
  const studentNames = [];
  if (studentSheet) {
    const lastRow = studentSheet.getLastRow();
    const data = lastRow > 1 ? studentSheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
    const cache = CacheService.getScriptCache();
    let cacheBatch = {};
    let batchKeyCount = 0;

    for (let i = 0; i < data.length; i++) {
      let rowId = String(data[i][0]).trim();
      let rowName = String(data[i][1]).trim();
      let rowEmail = String(data[i][2]).trim();

      if (rowName) studentNames.push(rowName);

      if (rowId || rowName || rowEmail) {
        const studentDataStr = JSON.stringify({ id: rowId, name: rowName, email: rowEmail });
        if (rowId) {
          cacheBatch["student_" + rowId.toLowerCase()] = studentDataStr;
          batchKeyCount++;
        }
        if (rowName) {
          cacheBatch["student_" + rowName.toLowerCase()] = studentDataStr;
          batchKeyCount++;
        }
        if (rowEmail) {
          cacheBatch["student_" + rowEmail.toLowerCase()] = studentDataStr;
          batchKeyCount++;
        }

        // Put in batches of 500 keys to avoid hitting any limits
        if (batchKeyCount >= 500) {
          cache.putAll(cacheBatch, 21600); // 6 hours
          cacheBatch = {};
          batchKeyCount = 0;
        }
      }
    }
    if (batchKeyCount > 0) {
      cache.putAll(cacheBatch, 21600);
    }
  }

  return {
    scriptUrl: ScriptApp.getService().getUrl(),
    students: studentNames,
    locations: getColumnData('Locations', 0),   // Locations Col A (index 0)
    clubs: getColumnData('Clubs', 0),           // Clubs Col A (index 0)
    counselors: getColumnData('Counselors', 0), // Counselors Col A (index 0)
    lrc: getColumnData('LRC', 0),               // LRC Col A (index 0)
    ase: getColumnData('ASE', 0)                // ASE Col A (index 0)
  };
}

// NEW: Multi-check-in function
function processMultiCheckIn(location, studentInputs, manualTimeStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Log');
  const studentSheet = ss.getSheetByName('Students');

  if (!logSheet) throw new Error("Make sure your tab is named exactly 'Log'.");
  if (!studentSheet) throw new Error("Make sure your tab is named exactly 'Students'.");

  const userEmail = Session.getActiveUser().getEmail();
  const actualNow = new Date();
  const now = manualTimeStr ? new Date(manualTimeStr) : actualNow;
  const cache = CacheService.getScriptCache();

  let validStudents = [];
  let invalidInputs = [];
  let successfulCheckIns = [];
  let successfulCheckOuts = [];
  let errors = [];

  // 1. Resolve all students first
  const studentDataRange = studentSheet.getLastRow() > 1 ? studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 3).getValues() : [];

  for (let input of studentInputs) {
    let inputStr = String(input).trim();
    if (!inputStr) continue;

    let result = resolveStudent(inputStr, cache, studentDataRange);

    if (result.found) {
      validStudents.push({ input: inputStr, id: result.id, name: result.name, email: result.email });
    } else {
      invalidInputs.push(inputStr);
    }
  }

  if (validStudents.length === 0) {
    return {
      successes: [],
      checkouts: [],
      errors: invalidInputs.map(input => `${input} (Not Found)`)
    };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("System is currently busy due to high traffic. Please try again in a few seconds.");
  }

  try {
    const oneHourMs = 60 * 60 * 1000;
    const lastRow = logSheet.getLastRow();
    const startRow = Math.max(2, lastRow - 999);
    const numRows = lastRow - startRow + 1;
    let logData = [];
    if (numRows > 0) {
      logData = logSheet.getRange(startRow, 1, numRows, 7).getValues();
    }

    let newRowsToAppend = [];

    // Process each valid student
    for (let student of validStudents) {
      let isCheckout = false;

      // 2. Check for active session for this student
      if (logData.length > 0) {
        for (let i = logData.length - 1; i >= 0; i--) {
          let row = logData[i];
          let checkOutTime = row[5];
          if (checkOutTime) continue;

          let rowId = String(row[2]).trim();
          if (rowId !== student.id) continue;

          let rowLocation = String(row[1]).trim();
          if (rowLocation !== location) continue;

          let checkInUser = String(row[7]).trim();
          if (checkInUser !== String(userEmail).trim()) continue;

          let checkInTime = new Date(row[0]);
          let timeDiffMs = now.getTime() - checkInTime.getTime();

          if (timeDiffMs >= 0 && timeDiffMs <= oneHourMs) {
            let durationMins = Math.round(timeDiffMs / 60000);
            const actualRowToUpdate = startRow + i;

            // Check out this student by updating the specific row in place.
            // In a batched scenario, doing this individually is fine if not many checkouts.
            logSheet.getRange(actualRowToUpdate, 6).setValue(now);
            logSheet.getRange(actualRowToUpdate, 7).setValue(durationMins);
            logSheet.getRange(actualRowToUpdate, 8).setValue(userEmail);

            if (manualTimeStr) {
              logSheet.getRange(actualRowToUpdate, 10).setValue(actualNow); // Column I: Actual Check-Out Time
            }

            successfulCheckOuts.push(`${student.name} (${durationMins} min)`);
            isCheckout = true;

            // Update local logData so subsequent checkouts logic doesn't pick it up again if there are duplicates
            logData[i][5] = now;
            break;
          }
        }
      }

      // 3. Prepare new check-in row
      if (!isCheckout) {
        let rowData = [
          now,
          sanitizeForSheets(location),
          sanitizeForSheets(student.id),
          sanitizeForSheets(student.name),
          sanitizeForSheets(student.email),
          "",
          "",
          sanitizeForSheets(userEmail)
        ];

        if (manualTimeStr) {
          rowData.push(actualNow); // Column I: Actual Check-In Time
          rowData.push("");        // Column J: Actual Check-Out Time
        }

        newRowsToAppend.push(rowData);
        successfulCheckIns.push(student.name);
      }
    }

    // Append all new check-ins at once
    if (newRowsToAppend.length > 0) {
      const numCols = newRowsToAppend[0].length;
      logSheet.getRange(lastRow + 1, 1, newRowsToAppend.length, numCols).setValues(newRowsToAppend);
    }

    // Add invalid inputs as errors
    if (invalidInputs.length > 0) {
      errors = invalidInputs.map(input => `${input} (Not Found)`);
    }

    return { successes: successfulCheckIns, checkouts: successfulCheckOuts, errors: errors };
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// NEW: Admin script to backfill student emails into the new Column E
function backfillStudentEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Log');
  const studentSheet = ss.getSheetByName('Students');

  if (!logSheet) throw new Error("Make sure your tab is named exactly 'Log'.");
  if (!studentSheet) throw new Error("Make sure your tab is named exactly 'Students'.");

  // 1. Build a dictionary mapping Student ID to Student Email
  const studentDataRange = studentSheet.getLastRow() > 1
      ? studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 3).getValues()
      : [];

  const emailMap = {};
  for (let i = 0; i < studentDataRange.length; i++) {
    const studentId = String(studentDataRange[i][0]).trim();
    const studentEmail = String(studentDataRange[i][2]).trim();

    if (studentId) {
      emailMap[studentId] = studentEmail;
    }
  }

  // 2. Read the Log sheet data (specifically Student ID in Column C, which is index 2)
  const logLastRow = logSheet.getLastRow();
  if (logLastRow <= 1) return; // Nothing to backfill

  const logDataRange = logSheet.getRange(2, 3, logLastRow - 1, 1).getValues(); // Get only Column C (Student ID)

  // 3. Prepare the array to write back into Column E
  const emailsToWrite = [];
  for (let i = 0; i < logDataRange.length; i++) {
    const logStudentId = String(logDataRange[i][0]).trim();
    // Look up email, default to blank if not found as per user request
    const email = emailMap[logStudentId] || "";
    emailsToWrite.push([email]);
  }

  // 4. Batch write the emails to Column E (Column 5)
  if (emailsToWrite.length > 0) {
    logSheet.getRange(2, 5, emailsToWrite.length, 1).setValues(emailsToWrite);
  }
}
