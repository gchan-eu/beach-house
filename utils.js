// ---------- UTILITIES & HELPER FUNCTIONS ----------

function getColumnIndexByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index   = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error("Header '" + headerName + "' not found on sheet '" + sheet.getName() + "'.");
  }
  return index + 1; // convert 0-based to 1-based index
}


function sameDate(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}


function round2(num) {
  return Math.round(num * 100) / 100;
}


function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}


function resolveHelperFromPid(pid) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const peopleSheet = ss.getSheetByName("PEOPLE");

  if (!peopleSheet) return pid;

  const codeCol   = getColumnIndexByHeader(peopleSheet, "Code");
  const helperCol = getColumnIndexByHeader(peopleSheet, "Helper");

  const lastRow = peopleSheet.getLastRow();
  if (lastRow < 2) return pid;

  const data = peopleSheet.getRange(2, 1, lastRow - 1, peopleSheet.getLastColumn()).getValues();

  for (let row of data) {
    if (row[codeCol - 1] == pid) {
      return row[helperCol - 1];   // return Helper value
    }
  }

  return pid; // fallback
}

