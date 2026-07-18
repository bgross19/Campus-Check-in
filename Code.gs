function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Campus Check-In')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processCheckIn(location, studentInput) {
  const lock = LockService.getScriptLock();
  // Wait for up to 30000 milliseconds for other processes to finish.
  if (!lock.tryLock(30000)) {
    throw new Error("System is currently busy due to high traffic. Please try again in a few seconds.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Log');
    const studentSheet = ss.getSheetByName('Students');

    if (!logSheet || !studentSheet) {
      throw new Error("Make sure your tabs are named exactly 'Log' and 'Students'.");
    }

    const userEmail = Session.getActiveUser().getEmail();

    // 1. Resolve Student Name and ID
    const data = studentSheet.getDataRange().getValues();
    let studentName = studentInput;
    let studentId = "Manual/Unknown";

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][0]).trim();
      let rowName = String(data[i][1]).trim();
      let inputStr = String(studentInput).trim();

      if (rowId === inputStr || rowName.toLowerCase() === inputStr.toLowerCase()) {
        studentId = rowId;
        studentName = rowName;
        break;
      }
    }

    // 2. Check for an active session within the last 1 hour
    const logData = logSheet.getDataRange().getValues();
    const now = new Date();
    const oneHourMs = 60 * 60 * 1000;

    for (let i = logData.length - 1; i > 0; i--) {
      let row = logData[i];
      let checkInTime = new Date(row[0]);
      let rowLocation = String(row[1]).trim();
      let rowId = String(row[2]).trim();
      let checkOutTime = row[4];

      if (rowId === studentId && rowLocation === location && !checkOutTime) {
        let timeDiffMs = now.getTime() - checkInTime.getTime();

        if (timeDiffMs <= oneHourMs) {
          let durationMins = Math.round(timeDiffMs / 60000);
          logSheet.getRange(i + 1, 5).setValue(now);
          logSheet.getRange(i + 1, 6).setValue(durationMins);
          logSheet.getRange(i + 1, 7).setValue(userEmail);
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

  return {
    students: getColumnData('Students', 1),     // Students Col B (index 1)
    locations: getColumnData('Locations', 0),   // Locations Col A (index 0)
    clubs: getColumnData('Clubs', 0),           // Clubs Col A (index 0)
    counselors: getColumnData('Counselors', 0)  // Counselors Col A (index 0)
  };
}
