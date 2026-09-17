const { performance } = require('perf_hooks');

const numStudents = 2000;
const studentDataRange = [];
for (let i = 0; i < numStudents; i++) {
  studentDataRange.push([String(10000 + i), `Student Name ${i}`, `student${i}@example.com`]);
}

const mockCache = {
  get: function(key) { return null; },
  put: function(key, val, time) { }
};

function resolveStudentOriginal(inputStr, cache, studentDataOrGetter) {
  let cacheKey = "student_" + inputStr.toLowerCase();
  let cachedData = cache.get(cacheKey);

  if (cachedData) {
    const parsed = JSON.parse(cachedData);
    return { found: true, id: parsed.id, name: parsed.name, email: parsed.email || "" };
  }

  let dataRange = typeof studentDataOrGetter === 'function' ? studentDataOrGetter() : studentDataOrGetter;

  for (let i = 0; i < dataRange.length; i++) {
    let rowId = String(dataRange[i][0]).trim();
    let rowName = String(dataRange[i][1]).trim();
    let rowEmail = String(dataRange[i][2]).trim();

    if (rowId === inputStr || rowName.toLowerCase() === inputStr.toLowerCase() || (rowEmail && rowEmail.toLowerCase() === inputStr.toLowerCase())) {
      cache.put(cacheKey, JSON.stringify({ id: rowId, name: rowName, email: rowEmail }), 21600);
      return { found: true, id: rowId, name: rowName, email: rowEmail };
    }
  }

  return { found: false, id: "Manual/Unknown", name: inputStr, email: "" };
}

function resolveStudentOptimized(inputStr, cache, studentDataOrGetter) {
  let cacheKey = "student_" + inputStr.toLowerCase();
  let cachedData = cache.get(cacheKey);

  if (cachedData) {
    const parsed = JSON.parse(cachedData);
    return { found: true, id: parsed.id, name: parsed.name, email: parsed.email || "" };
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
    return { found: true, id: match.id, name: match.name, email: match.email };
  }

  return { found: false, id: "Manual/Unknown", name: inputStr, email: "" };
}


const inputs = [];
for(let i=0; i<40; i++) {
    inputs.push(`Student Name ${1000 + i}`);
}
for(let i=0; i<10; i++) {
    inputs.push(`Not Found ${i}`);
}

function benchmarkOriginal() {
  const start = performance.now();
  for (let input of inputs) {
    resolveStudentOriginal(input, mockCache, studentDataRange);
  }
  return performance.now() - start;
}

function benchmarkOptimized() {
  const start = performance.now();
  delete studentDataRange._lookupMapExactId;
  delete studentDataRange._lookupMapLower;
  for (let input of inputs) {
    resolveStudentOptimized(input, mockCache, studentDataRange);
  }
  return performance.now() - start;
}

// Warmup
benchmarkOriginal();
benchmarkOptimized();

let origTimes = [];
let optTimes = [];
for (let i = 0; i < 100; i++) {
  origTimes.push(benchmarkOriginal());
  optTimes.push(benchmarkOptimized());
}

const avgOrig = origTimes.reduce((a, b) => a + b) / origTimes.length;
const avgOpt = optTimes.reduce((a, b) => a + b) / optTimes.length;

console.log("--- Benchmark: MultiCheckIn resolveStudent ---");
console.log("Original average time (ms):", avgOrig.toFixed(4));
console.log("Optimized average time (ms):", avgOpt.toFixed(4));
console.log("Speedup:", (avgOrig / avgOpt).toFixed(2) + "x");
console.log("Improvement:", (((avgOrig - avgOpt) / avgOrig) * 100).toFixed(2) + "%");
