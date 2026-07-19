function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Cathedral High School Check-In')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processCheckIn(location, studentInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Log');

  if (!logSheet) {
    throw new Error("Make sure your tab is named exactly 'Log'.");
  }

  const userEmail = Session.getActiveUser().getEmail();

  // 1. Resolve Student Name and ID
  let studentName = studentInput;
  let studentId = "Manual/Unknown";
  let inputStr = String(studentInput).trim();

  const cache = CacheService.getScriptCache();
  const cacheKey = "student_" + inputStr.toLowerCase();
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    const parsed = JSON.parse(cachedData);
    studentId = parsed.id;
    studentName = parsed.name;
  } else {
    const studentSheet = ss.getSheetByName('Students');
    if (!studentSheet) {
      throw new Error("Make sure your tab is named exactly 'Students'.");
    }
    const data = studentSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][0]).trim();
      let rowName = String(data[i][1]).trim();

      if (rowId === inputStr || rowName.toLowerCase() === inputStr.toLowerCase()) {
        studentId = rowId;
        studentName = rowName;
        cache.put(cacheKey, JSON.stringify({ id: studentId, name: studentName }), 21600); // 6 hours cache
        break;
      }
    }
  }

  const lock = LockService.getScriptLock();
  // Wait for up to 30000 milliseconds for other processes to finish.
  if (!lock.tryLock(30000)) {
    throw new Error("System is currently busy due to high traffic. Please try again in a few seconds.");
  }

  try {

    // 2. Check for an active session within the last 1 hour
    const now = new Date();
    const oneHourMs = 60 * 60 * 1000;
    const lastRow = logSheet.getLastRow();

    // Only look at the last 1000 rows to speed up check-ins/check-outs significantly.
    // If the sheet has fewer than 1000 rows, it stops at row 2 (header is row 1).
    const startRow = Math.max(2, lastRow - 999);
    const numRows = lastRow - startRow + 1;

    if (numRows > 0) {
      // getRange(row, column, numRows, numColumns)
      const logData = logSheet.getRange(startRow, 1, numRows, 5).getValues();

      for (let i = logData.length - 1; i >= 0; i--) {
        let row = logData[i];
        let checkOutTime = row[4];
        if (checkOutTime) continue;

        let rowId = String(row[2]).trim();
        if (rowId !== studentId) continue;

        let rowLocation = String(row[1]).trim();
        if (rowLocation !== location) continue;

        let checkInTime = new Date(row[0]);
        let timeDiffMs = now.getTime() - checkInTime.getTime();

        if (timeDiffMs <= oneHourMs) {
          let durationMins = Math.round(timeDiffMs / 60000);
          const actualRowToUpdate = startRow + i;
          logSheet.getRange(actualRowToUpdate, 5).setValue(now);
          logSheet.getRange(actualRowToUpdate, 6).setValue(durationMins);
          logSheet.getRange(actualRowToUpdate, 7).setValue(userEmail);
          return { name: studentName, status: "out", time: durationMins };
        }
      }
    }

    // 3. Log a new check-in
    logSheet.appendRow([now, location, studentId, studentName, "", "", userEmail]);
    return { name: studentName, status: "in" };
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function sanitizeForSheets(value) {
  if (typeof value === 'string' && /^[=+\-@]/.test(value)) {
    return "'" + value;
  }
  return value;
}

// NEW: Fetches all setup data in one fast call
function getSetupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Helper function to extract a specific column from a sheet
  const getColumnData = (sheetName, colIndex) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][colIndex]) list.push(String(data[i][colIndex]).trim());
    }
    return list;
  };

  // Cache all students in the background and build the list of names
  const studentSheet = ss.getSheetByName('Students');
  const studentNames = [];
  if (studentSheet) {
    const data = studentSheet.getDataRange().getValues();
    const cache = CacheService.getScriptCache();
    let cacheBatch = {};
    let batchKeyCount = 0;

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][0]).trim();
      let rowName = String(data[i][1]).trim();

      if (rowName) studentNames.push(rowName);

      if (rowId || rowName) {
        const studentDataStr = JSON.stringify({ id: rowId, name: rowName });
        if (rowId) {
          cacheBatch["student_" + rowId.toLowerCase()] = studentDataStr;
          batchKeyCount++;
        }
        if (rowName) {
          cacheBatch["student_" + rowName.toLowerCase()] = studentDataStr;
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
    students: studentNames,
    locations: getColumnData('Locations', 0),   // Locations Col A (index 0)
    clubs: getColumnData('Clubs', 0),           // Clubs Col A (index 0)
    counselors: getColumnData('Counselors', 0)  // Counselors Col A (index 0)
  };
}
