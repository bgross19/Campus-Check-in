const { performance } = require('perf_hooks');

const sheetRows = 2000;
const sheetCols = 26; // Simulated A-Z columns

const fullSheetData = [];
for (let r = 0; r < sheetRows; r++) {
  const row = [];
  for (let c = 0; c < sheetCols; c++) {
    row.push(r === 0 ? `Header${c}` : `Data_${r}_${c}`);
  }
  fullSheetData.push(row);
}

const twoColData = [];
for (let r = 1; r < sheetRows; r++) {
  twoColData.push([`Data_${r}_0`, `Data_${r}_1`]);
}

function original() {
  const start = performance.now();
  const data = fullSheetData; // simulate getDataRange()

  let names = [];
  for (let i = 1; i < data.length; i++) {
    let rowId = String(data[i][0]).trim();
    let rowName = String(data[i][1]).trim();
    if (rowName) names.push(rowName);
  }
  return { time: performance.now() - start, names: names.length };
}

function optimized() {
  const start = performance.now();
  const data = twoColData; // simulate getRange(2, 1, lastRow-1, 2)

  let names = [];
  for (let i = 0; i < data.length; i++) {
    let rowId = String(data[i][0]).trim();
    let rowName = String(data[i][1]).trim();
    if (rowName) names.push(rowName);
  }
  return { time: performance.now() - start, names: names.length };
}

original();
optimized();

let origTimes = [];
let optTimes = [];
for (let i = 0; i < 5000; i++) {
  origTimes.push(original().time);
  optTimes.push(optimized().time);
}

const avgOrig = origTimes.reduce((a, b) => a + b) / origTimes.length;
const avgOpt = optTimes.reduce((a, b) => a + b) / optTimes.length;

console.log("--- Benchmark: Students getDataRange ---");
console.log("Original average time (ms):", avgOrig.toFixed(4));
console.log("Optimized average time (ms):", avgOpt.toFixed(4));
console.log("Speedup:", (avgOrig / avgOpt).toFixed(2) + "x");
console.log("Improvement:", (((avgOrig - avgOpt) / avgOrig) * 100).toFixed(2) + "%");
