const { performance } = require('perf_hooks');

// ==========================================
// Benchmark 1: processCheckIn Loop Optimization
// ==========================================

// Generate mock data
const logData = [["Header", "Location", "StudentId", "StudentName", "CheckOutTime", "Duration", "Email"]];
const now = new Date();
const location = "Library";
const studentId = "12345";
const oneHourMs = 60 * 60 * 1000;

for (let i = 0; i < 100000; i++) {
  // Mostly mismatched data to simulate worst-case for the original code
  // and best-case for short-circuiting.
  // row: [checkInTime, rowLocation, rowId, studentName, checkOutTime, ...]
  logData.push([
    new Date(now.getTime() - i * 1000).toISOString(),
    "Room " + (i % 10),
    String(10000 + i),
    "Student " + i,
    i % 2 === 0 ? new Date().toISOString() : ""
  ]);
}

// Add the target match near the beginning (end of array)
logData.push([
  new Date(now.getTime() - 1000).toISOString(),
  "Library",
  "12345",
  "Target Student",
  ""
]);

function originalLoop() {
  const start = performance.now();
  let found = 0;
  for (let i = logData.length - 1; i > 0; i--) {
    let row = logData[i];
    let checkInTime = new Date(row[0]);
    let rowLocation = String(row[1]).trim();
    let rowId = String(row[2]).trim();
    let checkOutTime = row[4];

    if (rowId === studentId && rowLocation === location && !checkOutTime) {
      let timeDiffMs = now.getTime() - checkInTime.getTime();

      if (timeDiffMs <= oneHourMs) {
        found++;
      }
    }
  }
  const end = performance.now();
  return { time: end - start, found };
}

function optimizedLoop() {
  const start = performance.now();
  let found = 0;
  for (let i = logData.length - 1; i > 0; i--) {
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
      found++;
    }
  }
  const end = performance.now();
  return { time: end - start, found };
}

// Warm up
originalLoop();
optimizedLoop();

// Measure
let origTimes = [];
let optTimes = [];
for (let i = 0; i < 10; i++) {
  origTimes.push(originalLoop().time);
  optTimes.push(optimizedLoop().time);
}

const avgOrig = origTimes.reduce((a, b) => a + b) / origTimes.length;
const avgOpt = optTimes.reduce((a, b) => a + b) / optTimes.length;

console.log("--- Benchmark 1: processCheckIn Loop ---");
console.log("Original average time (ms):", avgOrig.toFixed(2));
console.log("Optimized average time (ms):", avgOpt.toFixed(2));
console.log("Speedup:", (avgOrig / avgOpt).toFixed(2) + "x");
console.log("Improvement:", (((avgOrig - avgOpt) / avgOrig) * 100).toFixed(2) + "%\n");


// ==========================================
// Benchmark 2: getColumnData Optimization
// ==========================================

// In Google Apps Script, fetching getDataRange() returns a 2D array of all rows and columns.
// Fetching a single column with getRange(row, col, numRows, numCols) returns a 2D array
// with only one element per row. The main cost is data serialization and memory payload.

// To simulate this in Node.js, we will create a large dataset and measure the time
// to process the data array. Note that in Google Apps Script, the actual fetch time
// difference is much more significant due to the network/bridge payload.

const sheetRows = 5000;
const sheetCols = 20;
const targetColIndex = 0; // Simulate extracting column A

// Generate simulated full sheet data (getDataRange)
const fullSheetData = [];
for (let r = 0; r < sheetRows; r++) {
  const row = [];
  for (let c = 0; c < sheetCols; c++) {
    row.push(r === 0 ? `Header${c}` : `Row${r}_Col${c}`);
  }
  fullSheetData.push(row);
}

// Generate simulated single column data (getRange)
const singleColData = [];
for (let r = 1; r < sheetRows; r++) { // Skipping header row
  singleColData.push([`Row${r}_Col${targetColIndex}`]);
}

function originalGetColumnData() {
  const start = performance.now();

  // Simulated: const data = sheet.getDataRange().getValues();
  const data = fullSheetData;

  const list = [];
  // Loop starts at 1 to skip header
  for (let i = 1; i < data.length; i++) {
    if (data[i][targetColIndex]) {
      list.push(String(data[i][targetColIndex]).trim());
    }
  }

  const end = performance.now();
  return { time: end - start, items: list.length };
}

function optimizedGetColumnData() {
  const start = performance.now();

  // Simulated: const data = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  const data = singleColData;

  const list = [];
  // Loop starts at 0 since header was already skipped in the getRange call
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      list.push(String(data[i][0]).trim());
    }
  }

  const end = performance.now();
  return { time: end - start, items: list.length };
}

// Warm up
originalGetColumnData();
optimizedGetColumnData();

// Measure
let origColTimes = [];
let optColTimes = [];
for (let i = 0; i < 1000; i++) {
  origColTimes.push(originalGetColumnData().time);
  optColTimes.push(optimizedGetColumnData().time);
}

const avgColOrig = origColTimes.reduce((a, b) => a + b) / origColTimes.length;
const avgColOpt = optColTimes.reduce((a, b) => a + b) / optColTimes.length;

console.log("--- Benchmark 2: getColumnData Data Processing ---");
console.log("Note: This only simulates JS array processing. Actual Google Apps Script fetch payload savings will be much larger.");
console.log("Original average time (ms):", avgColOrig.toFixed(4));
console.log("Optimized average time (ms):", avgColOpt.toFixed(4));
console.log("Speedup:", (avgColOrig / avgColOpt).toFixed(2) + "x");
console.log("Improvement:", (((avgColOrig - avgColOpt) / avgColOrig) * 100).toFixed(2) + "%");
