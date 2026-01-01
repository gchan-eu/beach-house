// ---------- UTILITIES & HELPER FUNCTIONS ----------

/**
 * Returns the 1-based column index for a given header name in a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to search.
 * @param {string} headerName - The header name to find.
 * @returns {number} The 1-based column index.
 */
function getColumnIndexByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index   = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error("Header '" + headerName + "' not found on sheet '" + sheet.getName() + "'.");
  }
  return index + 1; // convert 0-based to 1-based index
}


/**
 * Checks if two dates are the same calendar day.
 * @param {Date} a - First date.
 * @param {Date} b - Second date.
 * @returns {boolean} True if same day, else false.
 */
function sameDate(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}


/**
 * Rounds a number to 2 decimal places.
 * @param {number} num - The number to round.
 * @returns {number} The rounded number.
 */
function round2(num) {
  return Math.round(num * 100) / 100;
}


/**
 * Formats a date as yyyy-MM-dd in the script's timezone.
 * @param {Date} d - The date to format.
 * @returns {string} The formatted date string.
 */
function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}


/**
 * Resolves a Helper name from a person ID (PID).
 * @param {string|number} pid - The person ID.
 * @returns {string|number} The Helper name or PID if not found.
 */
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


/**
 * Returns a mapping of person IDs to overnight stays in a date range.
 * @param {Date} startDate - The start date.
 * @param {Date} endDate - The end date.
 * @returns {Object} Mapping of PID to stays count.
 */
function getOvernightStaysByPid(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const overnightSheet = ss.getSheetByName("OVERNIGHT_STAYS");
  if (!overnightSheet) return {};

  const osPersonCol    = getColumnIndexByHeader(overnightSheet, "Person");
  const osPersonIdCol  = getColumnIndexByHeader(overnightSheet, "Person_ID");
  const osStartCol     = getColumnIndexByHeader(overnightSheet, "Start_Date");
  const osEndCol       = getColumnIndexByHeader(overnightSheet, "End_Date");
  const osCountCol     = getColumnIndexByHeader(overnightSheet, "Person_Count");

  const lastRow = overnightSheet.getLastRow();
  const staysByPid = {};

  if (lastRow > 1) {
    const data = overnightSheet.getRange(2, 1, lastRow - 1, overnightSheet.getLastColumn()).getValues();

    data.forEach(row => {
      const rowStart = row[osStartCol - 1];
      const rowEnd   = row[osEndCol - 1];

      if (!(rowStart instanceof Date) || !(rowEnd instanceof Date)) return;

      if (rowStart <= endDate && rowEnd >= startDate) {
        const overlapStart = new Date(Math.max(rowStart, startDate));
        const overlapEnd   = new Date(Math.min(rowEnd, endDate));

        const overlapDays = Math.floor((overlapEnd - overlapStart) / (1000 * 60 * 60 * 24)) + 1;

        if (overlapDays > 0) {
          const personId    = row[osPersonIdCol - 1];
          const personCount = Number(row[osCountCol - 1]);
          const stays = overlapDays * personCount;
          staysByPid[personId] = (staysByPid[personId] || 0) + stays;
        }
      }
    });
  }

  return staysByPid; // { PID: staysCount, ... }
}


/**
 * Computes share fractions for each person based on overnight stays in a period.
 * @param {Date} startDate - The start date.
 * @param {Date} endDate - The end date.
 * @returns {Object} Mapping of PID to share fraction.
 */
function computeSharesByOvernightStays(startDate, endDate) {
  const staysByPid = getOvernightStaysByPid(startDate, endDate);
  const pids = Object.keys(staysByPid);
  if (pids.length === 0) return {};

  const totalStays = pids.reduce((sum, pid) => sum + staysByPid[pid], 0);
  if (totalStays === 0) return {};

  const shares = {};
  pids.forEach(pid => {
    shares[pid] = staysByPid[pid] / totalStays;
  });

  return shares; // { PID: share }
}

