function resolveStudentOriginal(inputStr, cache, studentDataOrGetter) {
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

  for (let i = 0; i < studentDataRange.length; i++) {
    let rowId = String(studentDataRange[i][0]).trim();
    let rowName = String(studentDataRange[i][1]).trim();
    let rowEmail = String(studentDataRange[i][2]).trim();

    if (rowId === inputStr || rowName.toLowerCase() === inputStr.toLowerCase() || (rowEmail && rowEmail.toLowerCase() === inputStr.toLowerCase())) {
      cache.put(cacheKey, JSON.stringify({ id: rowId, name: rowName, email: rowEmail }), 21600);
      return {
        found: true,
        id: rowId,
        name: rowName,
        email: rowEmail
      };
    }
  }

  return {
    found: false,
    id: "Manual/Unknown",
    name: inputStr,
    email: ""
  };
}

function resolveStudentOptimized(inputStr, cache, studentDataOrGetter) {
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

      // We want to keep the FIRST match found in the array to mimic original loop logic
      if (rowId && !mapExactId.has(rowId)) mapExactId.set(rowId, data);

      let lowerName = rowName.toLowerCase();
      let lowerEmail = rowEmail.toLowerCase();
      if (lowerName && !mapLower.has(lowerName)) mapLower.set(lowerName, data);
      if (lowerEmail && !mapLower.has(lowerEmail)) mapLower.set(lowerEmail, data);
    }
    Object.defineProperty(studentDataRange, '_lookupMapExactId', { value: mapExactId, enumerable: false });
    Object.defineProperty(studentDataRange, '_lookupMapLower', { value: mapLower, enumerable: false });
  }

  // Original checked: rowId === inputStr
  let match = studentDataRange._lookupMapExactId.get(inputStr);

  // Or lowercased name / email match
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

const mockCache = {
  store: {},
  get: function(key) { return this.store[key]; },
  put: function(key, val, time) { this.store[key] = val; }
};

const mockData = [
  ["123", "Alice Smith", "alice@example.com"],
  ["aBc12X", "Bob Jones", "bob@example.com"],
  ["789", "Charlie Brown", "charlie@example.com"],
  ["dup_id", "Duplicate Name", "dup1@example.com"],
  ["999", "Duplicate Name", "dup2@example.com"]
];

function test(name, inputStr) {
  mockCache.store = {};
  const resOrig = resolveStudentOriginal(inputStr, mockCache, mockData);
  mockCache.store = {};
  const resOpt = resolveStudentOptimized(inputStr, mockCache, mockData);
  console.log(`Test: ${name}`);
  console.log(`Input: ${inputStr}`);
  console.log(`Orig:`, resOrig);
  console.log(`Opt:`, resOpt);
  console.log(`Matches:`, JSON.stringify(resOrig) === JSON.stringify(resOpt));
  console.log();
}

test("ID Exact Case Sensitive", "aBc12X");
test("ID Exact Wrong Case", "abc12x");
test("First Name Found For Duplicates", "Duplicate Name");
test("Email match", "dup2@example.com");
