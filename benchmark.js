const { performance } = require('perf_hooks');

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

console.log("Original average time (ms):", avgOrig.toFixed(2));
console.log("Optimized average time (ms):", avgOpt.toFixed(2));
console.log("Speedup:", (avgOrig / avgOpt).toFixed(2) + "x");
console.log("Improvement:", (((avgOrig - avgOpt) / avgOrig) * 100).toFixed(2) + "%");
