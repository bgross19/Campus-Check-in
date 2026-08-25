import re

with open("Code.gs", "r") as f:
    code = f.read()

# Replace the single checkin `now` logic which was missed.
replace_start = """  try {

    // 2. Check for an active session within the last 1 hour
    const now = new Date();
    const oneHourMs = 60 * 60 * 1000;
    const lastRow = logSheet.getLastRow();"""

replace_with = """  try {

    // 2. Check for an active session within the last 1 hour
    const actualNow = new Date();
    const now = manualTimeStr ? new Date(manualTimeStr) : actualNow;
    const oneHourMs = 60 * 60 * 1000;
    const lastRow = logSheet.getLastRow();"""

code = code.replace(replace_start, replace_with)

with open("Code.gs", "w") as f:
    f.write(code)
