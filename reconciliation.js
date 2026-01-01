// ---------- CORE ENGINE: RECONCILE CHARGES ----------

function reconcileCharges(expensesSheet, triggerRow) {
  const headers = expensesSheet.getRange(1, 1, 1, expensesSheet.getLastColumn()).getValues()[0];

  const statusCol  = headers.indexOf("Status") + 1;
  const startCol   = headers.indexOf("Start_Date") + 1;
  const endCol     = headers.indexOf("End_Date") + 1;
  const typeCol    = headers.indexOf("Type") + 1;
  const lastRecCol = headers.indexOf("Last_Reconciliation_Date") + 1; // optional; may be 0

  const status      = expensesSheet.getRange(triggerRow, statusCol).getValue();
  const startDate   = expensesSheet.getRange(triggerRow, startCol).getValue();
  const endDate     = expensesSheet.getRange(triggerRow, endCol).getValue();
  const expenseType = expensesSheet.getRange(triggerRow, typeCol).getValue();

  if (!startDate || !endDate) {
    SpreadsheetApp.getActive().toast("Start_Date and End_Date must be set for reconciliation.", "❌ Error", 5);
    return;
  }

  if (!status || (!status.toString().startsWith("Provisionally Charged") && !status.toString().startsWith("Reconciled"))) {
    SpreadsheetApp.getActive().toast("Only expenses with Status 'Provisionally Charged' or 'Reconciled' can be reconciled.", "❌ Error", 5);
    return;
  }

  Logger.log('Recon group params: startDate=' + startDate + ', endDate=' + endDate + ', expenseType=' + expenseType);
  // Modified: Include both 'Provisionally Charged' and 'Reconciled' expenses in the group
  const groupRaw = getReconciliationGroup(expensesSheet, startDate, endDate, expenseType);
  Logger.log('Raw group length: ' + groupRaw.length);
  if (groupRaw.length === 0) {
    Logger.log('No rows matched in getReconciliationGroup.');
  } else {
    groupRaw.forEach(item => {
      let status = expensesSheet.getRange(item.rowIndex, statusCol).getValue();
      Logger.log('Raw group row ' + item.rowIndex + ' status: ' + status);
    });
  }
  const group = groupRaw.filter(item => {
    let status = expensesSheet.getRange(item.rowIndex, statusCol).getValue();
    if (!status) return false;
    status = String(status)
      .replace(/\u00A0/g, " ")   // remove non-breaking spaces
      .replace(/\s+/g, " ")      // collapse weird whitespace
      .trim()
      .toLowerCase();
    return status.startsWith("provisionally charged") || status.startsWith("reconciled");
  });
  Logger.log('Group length after filter: ' + group.length);
  if (group.length === 0) {
    SpreadsheetApp.getActive().toast("No provisionally charged or reconciled expenses found in this date window and expense type.", "ℹ️ Info", 5);
    return;
  }

  const totalCost = group.reduce((sum, item) => sum + item.amount, 0);
  if (totalCost === 0) {
    SpreadsheetApp.getActive().toast("Total cost for this reconciliation group is zero.", "ℹ️ Info", 5);
    return;
  }

  const shares = computeSharesByOvernightStays(startDate, endDate);
  const pids   = Object.keys(shares);
  if (pids.length === 0) {
    SpreadsheetApp.getActive().toast("No overnight stays found in this period. Nothing to reconcile.", "ℹ️ Info", 5);
    return;
  }

  const finalCostByPid = {};
  pids.forEach(pid => {
    finalCostByPid[pid] = round2(totalCost * shares[pid]);
  });

  const chargedSoFarByPid = getChargedSoFarForGroup(group);

  createReconciliationAdjustments(group, finalCostByPid, chargedSoFarByPid, startDate, endDate);

  const now = new Date();
  function formatDMY(d) {
    return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + String(d.getFullYear()).slice(-2);
  }
  const statusText = "Reconciled (" + formatDMY(now) + ")";

  group.forEach(item => {
    expensesSheet.getRange(item.rowIndex, statusCol).setValue(statusText);
    if (lastRecCol > 0) {
      expensesSheet.getRange(item.rowIndex, lastRecCol).setValue(now);
    }
  });

  SpreadsheetApp.getActive().toast("Reconciliation completed for this period and expense type.", "✅ Reconciled", 4);
}


function createReconciliationAdjustments(groupExpenses, finalCostByPid, chargedSoFarByPid, startDate, endDate) {
  const transactionsSheet = SpreadsheetApp.getActive().getSheetByName("TRANSACTIONS");
  const lastRow  = transactionsSheet.getLastRow();
  const headers  = transactionsSheet.getRange(1, 1, 1, transactionsSheet.getLastColumn()).getValues()[0];

  const idCol           = headers.indexOf("ID") + 1;
  const dateCol         = headers.indexOf("Date") + 1;
  const typeCol         = headers.indexOf("Type") + 1;
  const expenseTypeCol  = headers.indexOf("Expense_Type") + 1;
  const amountCol       = headers.indexOf("Amount") + 1;
  const personCol       = headers.indexOf("Person") + 1;
  const accountCol      = headers.indexOf("Account") + 1;
  const expenseIdCol    = headers.indexOf("Expense_ID") + 1;
  const noteCol         = headers.indexOf("Notes") + 1;

  // Find next transaction ID
  let nextTransactionId = 1;
  if (lastRow > 1) {
    const idRange = transactionsSheet.getRange(2, idCol, lastRow - 1).getValues().flat().filter(n => typeof n === "number");
    if (idRange.length > 0) nextTransactionId = Math.max(...idRange) + 1;
  }

  const now = new Date();
  const rowsToInsert = [];
  const pids = Object.keys(finalCostByPid);


  // Get primary expense type for linking
  const primaryExpenseId = groupExpenses[0].id; // for linking
  const expensesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EXPENSES");
  const expenseTypeHeaderCol = getColumnIndexByHeader(expensesSheet, "Type");
  const expenseIdHeaderCol = getColumnIndexByHeader(expensesSheet, "ID");
  let expenseTypeValue = "";
  for (let r = 2; r <= expensesSheet.getLastRow(); r++) {
    if (expensesSheet.getRange(r, expenseIdHeaderCol).getValue() == primaryExpenseId) {
      expenseTypeValue = expensesSheet.getRange(r, expenseTypeHeaderCol).getValue();
      break;
    }
  }

  // Prepare PEOPLE lookup for Account
  const peopleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PEOPLE");
  const codeCol = getColumnIndexByHeader(peopleSheet, "Code");
  const accountColPeople = getColumnIndexByHeader(peopleSheet, "Account_Number");
  const helperCol = getColumnIndexByHeader(peopleSheet, "Helper");
  const peopleData = peopleSheet.getRange(2, 1, peopleSheet.getLastRow() - 1, peopleSheet.getLastColumn()).getValues();

  function getAccountForPid(pid) {
    for (let i = 0; i < peopleData.length; i++) {
      if (String(peopleData[i][codeCol - 1]) === String(pid)) {
        return peopleData[i][accountColPeople - 1];
      }
    }
    return '';
  }

  // Calculate stays info for the period (using Days, not Total_Stays)
  function getDaysByPid(startDate, endDate) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const overnightSheet = ss.getSheetByName("OVERNIGHT_STAYS");
    if (!overnightSheet) return {};

    const osPersonIdCol  = getColumnIndexByHeader(overnightSheet, "Person_ID");
    const osStartCol     = getColumnIndexByHeader(overnightSheet, "Start_Date");
    const osEndCol       = getColumnIndexByHeader(overnightSheet, "End_Date");

    const lastRow = overnightSheet.getLastRow();
    const daysByPid = {};
    const today = new Date();
    today.setHours(0,0,0,0); // ignore time

    if (lastRow > 1) {
      const data = overnightSheet.getRange(2, 1, lastRow - 1, overnightSheet.getLastColumn()).getValues();

      data.forEach(row => {
        const rowStart = row[osStartCol - 1];
        const rowEnd   = row[osEndCol - 1];
        if (!(rowStart instanceof Date) || !(rowEnd instanceof Date)) return;

        // Calculate overlap between stay and reconciliation period, capped at today
        const overlapStart = new Date(Math.max(rowStart, startDate));
        let overlapEnd = new Date(Math.min(rowEnd, endDate, today));
        if (overlapEnd > today) overlapEnd = today;
        const overlapDays = Math.floor((overlapEnd - overlapStart) / (1000 * 60 * 60 * 24)) + 1;
        if (overlapStart <= overlapEnd && overlapDays > 0) {
          const personId = row[osPersonIdCol - 1];
          daysByPid[personId] = (daysByPid[personId] || 0) + overlapDays;
        }
      });
    }
    return daysByPid;
  }

  const daysByPid = getDaysByPid(startDate, endDate);
  const totalDays = Object.values(daysByPid).reduce((a, b) => a + b, 0);

  // Calculate group total charged so far
  const groupTotalCharged = Object.values(chargedSoFarByPid).reduce((a, b) => a + b, 0);
  // Calculate total days (sum of all person days)
  // Already calculated above: const totalDays = Object.values(daysByPid).reduce((a, b) => a + b, 0);

  // N-1 rounding on fair shares
  let fairShares = [];
  let runningFairShare = 0;
  const n = pids.length;
  for (let i = 0; i < n; i++) {
    const pid = pids[i];
    const personDays = daysByPid[pid] || 0;
    const shareFraction = totalDays !== 0 ? personDays / totalDays : 0;
    let fairShare;
    if (i < n - 1) {
      fairShare = round2(groupTotalCharged * shareFraction);
      runningFairShare += fairShare;
    } else {
      // Last person gets the remainder to ensure sum matches groupTotalCharged
      fairShare = round2(groupTotalCharged - runningFairShare);
    }
    fairShares.push(fairShare);
  }

  for (let i = 0; i < n; i++) {
    const pid = pids[i];
    const personDays = daysByPid[pid] || 0;
    const personName = resolveHelperFromPid(pid);
    const chargedSoFar = chargedSoFarByPid[personName] || 0;
    // Adjustment: fair share minus charged so far, rounded to 2 decimals
    const adj = round2(fairShares[i] - chargedSoFar);
    // Format period as dd/mm/yy
    function formatDMY(d) {
      return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + String(d.getFullYear()).slice(-2);
    }
    const today = new Date();
    today.setHours(0,0,0,0);
    let costLabel;
    if (today > endDate) {
      costLabel = "final cost";
    } else {
      costLabel = "cost (" + formatDMY(today) + ")";
    }
    const note =
      "Period: " + formatDMY(startDate) + " – " + formatDMY(endDate) +
      ", " + costLabel + ": " + fairShares[i].toFixed(2) +
      ", charged so far: " + chargedSoFar.toFixed(2) +
      ", adjustment: " + adj.toFixed(2) +
      " (" + personDays + "/" + totalDays + ").";

    const accountValue = getAccountForPid(pid);

    // Only create a row if adjustment is not zero
    if (adj !== 0) {
      const row = new Array(headers.length);
      row[idCol - 1]           = nextTransactionId++;
      row[dateCol - 1]         = now;
      row[typeCol - 1]         = "402 - Reconciliation";
      row[expenseTypeCol - 1]  = expenseTypeValue;
      row[amountCol - 1]       = adj; // negative: owes more, positive: refund
      row[personCol - 1]       = personName;
      row[expenseIdCol - 1]    = primaryExpenseId;
      row[noteCol - 1]         = note;
      row[accountCol - 1]      = accountValue;

      rowsToInsert.push(row);
    }
  }

  // Add refund adjustments for people who were provisionally charged but have zero share in the period
  Object.keys(chargedSoFarByPid).forEach(personName => {
    // Find pid for this personName (Helper)
    let pid = null;
    for (let i = 0; i < peopleData.length; i++) {
      if (resolveHelperFromPid(peopleData[i][codeCol - 1]) === personName) {
        pid = peopleData[i][codeCol - 1];
        break;
      }
    }
    if (!pid) return;
    if (pids.includes(pid)) return; // already handled above
    const chargedSoFar = chargedSoFarByPid[personName] || 0;
    if (chargedSoFar === 0) return;
    // Refund adjustment: fair share is zero, so adjustment is -chargedSoFar
    const adj = round2(0 - chargedSoFar);
    if (adj !== 0) {
      const today = new Date();
      today.setHours(0,0,0,0);
      function formatDMY(d) {
        return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + String(d.getFullYear()).slice(-2);
      }
      let costLabel;
      if (today > endDate) {
        costLabel = "final cost";
      } else {
        costLabel = "cost (" + formatDMY(today) + ")";
      }
      const note =
        "Period: " + formatDMY(startDate) + " – " + formatDMY(endDate) +
        ", " + costLabel + ": 0.00" +
        ", charged so far: " + chargedSoFar.toFixed(2) +
        ", adjustment: " + adj.toFixed(2) +
        " (0/" + totalDays + ").";
      const accountValue = getAccountForPid(pid);
      const row = new Array(headers.length);
      row[idCol - 1]           = nextTransactionId++;
      row[dateCol - 1]         = now;
      row[typeCol - 1]         = "402 - Reconciliation";
      row[expenseTypeCol - 1]  = expenseTypeValue;
      row[amountCol - 1]       = adj;
      row[personCol - 1]       = personName;
      row[expenseIdCol - 1]    = primaryExpenseId;
      row[noteCol - 1]         = note;
      row[accountCol - 1]      = accountValue;
      rowsToInsert.push(row);
    }
  });

  if (rowsToInsert.length > 0) {
    transactionsSheet.getRange(lastRow + 1, 1, rowsToInsert.length, headers.length).setValues(rowsToInsert);
  }
}


function getReconciliationGroup(expensesSheet, startDate, endDate, expenseType) {
  const values  = expensesSheet.getDataRange().getValues();
  const headers = values[0];

  const idIdx      = headers.indexOf("ID");
  const amountIdx  = headers.indexOf("Amount");
  const statusIdx  = headers.indexOf("Status");
  const startIdx   = headers.indexOf("Start_Date");
  const endIdx     = headers.indexOf("End_Date");
  const typeIdx    = headers.indexOf("Type");

  const result = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status   = row[statusIdx];
    const rowStart = row[startIdx];
    const rowEnd   = row[endIdx];
    const rowType  = row[typeIdx];
    const statusStr = status ? status.toString() : '';
    const startMatch = sameDate(rowStart, startDate);
    const endMatch = sameDate(rowEnd, endDate);
    const typeMatch = rowType === expenseType;
    Logger.log('ReconGroup row ' + (r+1) + ': status=' + statusStr + ', rowStart=' + rowStart + ', startMatch=' + startMatch + ', rowEnd=' + rowEnd + ', endMatch=' + endMatch + ', rowType=' + rowType + ', typeMatch=' + typeMatch);
    if (!status) continue;
    const statusNorm = statusStr.toLowerCase();
    if (!(statusNorm.startsWith("provisionally charged") || statusNorm.startsWith("reconciled"))) continue;
    if (!startMatch || !endMatch) continue;
    if (!typeMatch) continue;
    result.push({
      rowIndex: r + 1,
      id: row[idIdx],
      amount: Number(row[amountIdx]) || 0,
    });
  }

  return result;
}


function getChargedSoFarForGroup(groupExpenses) {
  const expenseIds = groupExpenses.map(e => e.id);
  if (expenseIds.length === 0) return {};

  const transactionsSheet = SpreadsheetApp.getActive().getSheetByName("TRANSACTIONS");
  const values  = transactionsSheet.getDataRange().getValues();
  const headers = values[0];

  const expenseIdIdx = headers.indexOf("Expense_ID");
  const personIdx    = headers.indexOf("Person");
  const amountIdx    = headers.indexOf("Amount");

  const charged = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const expId = row[expenseIdIdx];
    if (!expenseIds.includes(expId)) continue;

    const person = row[personIdx];
    const amount = Number(row[amountIdx]) || 0;

    if (!person) continue;
    if (!charged[person]) charged[person] = 0;
    charged[person] += amount;
  }

  return charged; // { Helper: totalAmount }
}

