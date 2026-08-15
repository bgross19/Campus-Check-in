const Code = require('fs').readFileSync('Code.gs', 'utf8');

// A very basic mock of Google Apps Script objects to test the fix
const mockLogSheet = {
  appendRow: function(row) {
    this.lastAppendedRow = row;
  }
};

const mockSpreadsheet = {
  getSheetByName: function(name) {
    if (name === 'Log') return mockLogSheet;
    return null;
  }
};

global.SpreadsheetApp = {
  getActiveSpreadsheet: () => mockSpreadsheet,
  flush: () => {}
};

global.Session = {
  getActiveUser: () => ({ getEmail: () => 'test@example.com' })
};

global.CacheService = {
  getScriptCache: () => ({
    get: () => null,
    put: () => {}
  })
};

global.LockService = {
  getScriptLock: () => ({
    tryLock: () => true,
    releaseLock: () => {}
  })
};

// Evaluate the functions from Code.gs
eval(Code);

// Test sanitizeForSheets function directly
console.log("Testing sanitizeForSheets directly:");
console.log("=SUM(A1:B1) ->", sanitizeForSheets("=SUM(A1:B1)"));
console.log("+1+2 ->", sanitizeForSheets("+1+2"));
console.log("-5 ->", sanitizeForSheets("-5"));
console.log("@here ->", sanitizeForSheets("@here"));
console.log("Safe string ->", sanitizeForSheets("Safe string"));
console.log("123 ->", sanitizeForSheets(123));

// Check if sanitizeForSheets works when it's not starting with a bad char
console.log("abc=SUM() ->", sanitizeForSheets("abc=SUM()"));

// Note: It's hard to mock all of processCheckIn completely without mock data for Students
console.log("\nFinished tests.");
