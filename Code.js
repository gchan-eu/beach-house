function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const name = sheet.getName();

  switch (name) {
    case "OWNERSHIP_SETS":
      handleOwnershipSet(e);
      break;

    case "OWNERSHIP_DETAILS":
      handleOwnerShare(e);
      break;

    case "OVERNIGHT_STAYS":
      handleOvernightStay(e);
      break;

    case "SPLIT_METHODS":
      handleSplitMethods(e);
      break;

    case "EXPENSES":
      handleExpenses(e);
      break;

    case "TRANSACTIONS":
      handleTransactions(e);
      break;

    // Add more sheets here
  }
}



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

