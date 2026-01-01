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


// ---------- SHEET HANDLERS ----------

function handleOwnershipSet(e) {
  assignOwnershipSetId(e);
}

function handleOwnerShare(e) {
  assignOwnerShareId(e);
  // add more OWNER_SHARE features here
}

function handleOvernightStay(e) {
  assignOvernightStayId(e);
  // add more OVERNIGHT_STAYS features here
}

function handleSplitMethods(e) {
  assignSplitMethodId(e);
  // add more SPLIT_METHODS features here
}

function handleExpenses(e) {
  const sheet = e.source.getActiveSheet();
  const col = e.range.getColumn();
  const actionCol = getColumnIndexByHeader(sheet, "Action");

  // If user edited the Action column, DO NOT run assignExpenseId
  if (col === actionCol) {
    handleExpenseActions(e);
    return;
  }

  // Otherwise, normal behavior
  assignExpenseId(e);
}


function handleTransactions(e) {
  assignTransactionId(e);
  // add more TRANSACTIONS features here
}


// ---------- ID ASSIGNMENT HELPERS ----------

function assignOwnershipSetId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const triggerColumn = getColumnIndexByHeader(sheet, "Date");
  const idColumn      = getColumnIndexByHeader(sheet, "ID");
  const startNumber   = 100001;

  if (row === 1 || col !== triggerColumn) return;

  const idCell = sheet.getRange(row, idColumn);
  if (idCell.getValue()) return;

  const idRange = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  const existingIds = idRange.flat().filter(n => typeof n === "number");

  const nextId = existingIds.length > 0
    ? Math.max(...existingIds) + 1
    : startNumber;

  idCell.setValue(nextId);
}

function assignOwnerShareId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const triggerColumn = getColumnIndexByHeader(sheet, "Ownership_Set_ID");
  const idColumn      = getColumnIndexByHeader(sheet, "ID");
  const startNumber   = 100001;

  if (row === 1 || col !== triggerColumn) return;

  const idCell = sheet.getRange(row, idColumn);
  if (idCell.getValue()) return;

  const idRange = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  const existingIds = idRange.flat().filter(n => typeof n === "number");

  const nextId = existingIds.length > 0
    ? Math.max(...existingIds) + 1
    : startNumber;

  idCell.setValue(nextId);
}

function assignOvernightStayId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const triggerColumn = getColumnIndexByHeader(sheet, "Person");
  const idColumn      = getColumnIndexByHeader(sheet, "ID");
  const startNumber   = 100001;

  if (row === 1 || col !== triggerColumn) return;

  const idCell = sheet.getRange(row, idColumn);
  if (idCell.getValue()) return;

  const idRange = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  const existingIds = idRange.flat().filter(n => typeof n === "number");

  const nextId = existingIds.length > 0
    ? Math.max(...existingIds) + 1
    : startNumber;

  idCell.setValue(nextId);
}

function assignSplitMethodId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const triggerColumn = getColumnIndexByHeader(sheet, "Type");
  const idColumn      = getColumnIndexByHeader(sheet, "ID");
  const startNumber   = 101;

  if (row === 1 || col !== triggerColumn) return;

  const idCell = sheet.getRange(row, idColumn);
  if (idCell.getValue()) return;

  const idRange = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  const existingIds = idRange.flat().filter(n => typeof n === "number");

  const nextId = existingIds.length > 0
    ? Math.max(...existingIds) + 1
    : startNumber;

  idCell.setValue(nextId);
}

function assignExpenseId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const triggerColumn = getColumnIndexByHeader(sheet, "Date");
  const idColumn      = getColumnIndexByHeader(sheet, "ID");
  const statusColumn  = getColumnIndexByHeader(sheet, "Status");
  const startNumber   = 100001;

  if (row === 1 || col !== triggerColumn) return;

  const idCell = sheet.getRange(row, idColumn);
  const statusCell = sheet.getRange(row, statusColumn);

  // Assign ID if missing
  if (!idCell.getValue()) {
    const idRange = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
    const existingIds = idRange.flat().filter(n => typeof n === "number");

    const nextId = existingIds.length > 0
      ? Math.max(...existingIds) + 1
      : startNumber;

    idCell.setValue(nextId);
  }

  // Set Status to "Pending" if empty
  if (!statusCell.getValue()) {
    statusCell.setValue("Pending");
  }
}


function assignTransactionId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const triggerColumn = getColumnIndexByHeader(sheet, "Date");
  const idColumn      = getColumnIndexByHeader(sheet, "ID");
  const startNumber   = 1;

  if (row === 1 || col !== triggerColumn) return;

  const idCell = sheet.getRange(row, idColumn);
  if (idCell.getValue()) return;

  const idRange = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  const existingIds = idRange.flat().filter(n => typeof n === "number");

  const nextId = existingIds.length > 0
    ? Math.max(...existingIds) + 1
    : startNumber;

  idCell.setValue(nextId);
}


// ---------- EXPENSE ACTIONS ----------

function handleExpenseActions(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const ACTION_COLUMN = getColumnIndexByHeader(sheet, "Action");

  if (row === 1) return;              // ignore header
  if (col !== ACTION_COLUMN) return;  // only trigger on Action column

  runExpenseAction(e);
}


function runExpenseAction(e) {
  const sheet  = e.source.getActiveSheet();
  const row    = e.range.getRow();
  const action = e.range.getValue();

  Logger.log("Action received: '" + action + "'");

  switch (action) {
    case "Create Charges":
      createCharges(sheet, row);
      break;

    case "Create Provisional Charges":
      createProvisionalCharges(sheet, row);
      break;

    case "Reconcile Charges":
      reconcileCharges(sheet, row);
      break;

    case "Delete Charges":
      deleteChargesForExpense(sheet, row);
      break;
  }

  // Clear the action cell after running
  e.range.setValue("");
}



// ---------- CORE ENGINE: CREATE CHARGES ----------

function createCharges(expensesSheet, expenseRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Read EXPENSES row data ---
  const splitMethodIDCol  = getColumnIndexByHeader(expensesSheet, "Split_Method_ID");
  const splitMethodCol    = getColumnIndexByHeader(expensesSheet, "Split_Method_Type");
  const dateCol           = getColumnIndexByHeader(expensesSheet, "Date");
  const amountCol         = getColumnIndexByHeader(expensesSheet, "Amount");
  const expenseIdCol      = getColumnIndexByHeader(expensesSheet, "ID");
  const expenseTypeCol    = getColumnIndexByHeader(expensesSheet, "Type");
  const expenseStartCol   = getColumnIndexByHeader(expensesSheet, "Start_Date");
  const expenseEndCol     = getColumnIndexByHeader(expensesSheet, "End_Date");

  const splitMethodId   = Number(expensesSheet.getRange(expenseRow, splitMethodIDCol).getValue());
  const splitMethodType = Number(expensesSheet.getRange(expenseRow, splitMethodCol).getValue());

  if (![1, 2, 3, 4].includes(splitMethodType)) {
    SpreadsheetApp.getActive().toast("Unsupported split method type.", "❌ Error", 5);
    return;
  }

  const expenseDate   = expensesSheet.getRange(expenseRow, dateCol).getValue();
  const expenseAmount = Number(expensesSheet.getRange(expenseRow, amountCol).getValue());
  const expenseId     = expensesSheet.getRange(expenseRow, expenseIdCol).getValue();
  const expenseType   = expensesSheet.getRange(expenseRow, expenseTypeCol).getValue();
  const expenseStart  = expensesSheet.getRange(expenseRow, expenseStartCol).getValue();
  const expenseEnd    = expensesSheet.getRange(expenseRow, expenseEndCol).getValue();

  if (!expenseDate || !expenseAmount) {
    SpreadsheetApp.getActive().toast("Missing Date or Amount.", "❌ Cannot Create Charges", 5);
    return;
  }

  // --- Prevent duplicate charges ---
  const transactionsSheet = ss.getSheetByName("TRANSACTIONS");
  const tExpenseIdCol = getColumnIndexByHeader(transactionsSheet, "Expense_ID");

  const tLastRow = transactionsSheet.getLastRow();
  if (tLastRow > 1) {
    const tData = transactionsSheet.getRange(2, 1, tLastRow - 1, transactionsSheet.getLastColumn()).getValues();
    const duplicates = tData.some(row => row[tExpenseIdCol - 1] === expenseId);

    if (duplicates) {
      SpreadsheetApp.getActive().toast(
        "Charges already exist for Expense ID " + expenseId,
        "❌ Duplicate Charges",
        5
      );
      return;
    }
  }

  // --- Method 3 requires Start/End dates ---
  if (splitMethodType === 3) {
    if (!(expenseStart instanceof Date) || !(expenseEnd instanceof Date)) {
      SpreadsheetApp.getActive().toast("Method 3 requires Start/End dates.", "❌ Cannot Create Charges", 5);
      return;
    }
  }

  // --- 2. Method 3: Overnight Stays ---
  let overnightSplits = null;
  let totalStays = 0;

  if (splitMethodType === 3) {
    const overnightSheet = ss.getSheetByName("OVERNIGHT_STAYS");

    const osPersonCol    = getColumnIndexByHeader(overnightSheet, "Person");
    const osPersonIdCol  = getColumnIndexByHeader(overnightSheet, "Person_ID");
    const osStartCol     = getColumnIndexByHeader(overnightSheet, "Start_Date");
    const osEndCol       = getColumnIndexByHeader(overnightSheet, "End_Date");
    const osCountCol     = getColumnIndexByHeader(overnightSheet, "Person_Count");

    const lastRow = overnightSheet.getLastRow();
    const staysByPerson = {};
    const personNameById = {};

    if (lastRow > 1) {
      const data = overnightSheet.getRange(2, 1, lastRow - 1, overnightSheet.getLastColumn()).getValues();

      data.forEach(row => {
        const rowStart = row[osStartCol - 1];
        const rowEnd   = row[osEndCol - 1];

        if (!(rowStart instanceof Date) || !(rowEnd instanceof Date)) return;

        if (rowStart <= expenseEnd && rowEnd >= expenseStart) {
          const overlapStart = new Date(Math.max(rowStart, expenseStart));
          const overlapEnd   = new Date(Math.min(rowEnd, expenseEnd));

          const overlapDays = Math.floor((overlapEnd - overlapStart) / (1000 * 60 * 60 * 24)) + 1;

          if (overlapDays > 0) {
            const personId    = row[osPersonIdCol - 1];
            const personName  = row[osPersonCol - 1];
            const personCount = Number(row[osCountCol - 1]);

            const stays = overlapDays * personCount;

            staysByPerson[personId] = (staysByPerson[personId] || 0) + stays;
            personNameById[personId] = personName;
          }
        }
      });

      totalStays = Object.values(staysByPerson).reduce((a, b) => a + b, 0);
      if (totalStays === 0) {
        SpreadsheetApp.getActive().toast("No overnight stays found.", "❌ Cannot Create Charges", 5);
        return;
      }

      overnightSplits = Object.keys(staysByPerson).map(pid => ({
        personName: personNameById[pid],
        percentage: staysByPerson[pid] / totalStays,
        stays: staysByPerson[pid]
      }));
    }
  }

  // --- 3. Method 4: Custom Split ---
  let customSplits = null;

  if (splitMethodType === 4) {
    const splitMethodsSheet = ss.getSheetByName("SPLIT_METHODS");
    const smIdCol   = getColumnIndexByHeader(splitMethodsSheet, "ID");
    const smJsonCol = getColumnIndexByHeader(splitMethodsSheet, "JSON");

    const smLastRow = splitMethodsSheet.getLastRow();
    const smData = splitMethodsSheet.getRange(2, 1, smLastRow - 1, splitMethodsSheet.getLastColumn()).getValues();

    let jsonString = null;

    smData.forEach(row => {
      if (Number(row[smIdCol - 1]) === splitMethodId) {
        jsonString = row[smJsonCol - 1];
      }
    });

    if (!jsonString) {
      SpreadsheetApp.getActive().toast("No JSON found for custom split.", "❌ Cannot Create Charges", 5);
      return;
    }

    try {
      customSplits = JSON.parse(jsonString);
    } catch {
      SpreadsheetApp.getActive().toast("Invalid JSON for custom split.", "❌ Cannot Create Charges", 5);
      return;
    }
  }

  // --- 4. Ownership Set lookup (Methods 1 & 2 only) ---
  let applicableOwners = [];

  if (splitMethodType === 1 || splitMethodType === 2) {
    const ownershipSetsSheet = ss.getSheetByName("OWNERSHIP_SETS");
    const osetDateCol = getColumnIndexByHeader(ownershipSetsSheet, "Date");
    const osetIdCol   = getColumnIndexByHeader(ownershipSetsSheet, "ID");

    const osLastRow = ownershipSetsSheet.getLastRow();
    const osData = ownershipSetsSheet.getRange(2, 1, osLastRow - 1, ownershipSetsSheet.getLastColumn()).getValues();

    let chosenSetId = null;
    let chosenSetDate = null;

    osData.forEach(row => {
      const rowDate = row[osetDateCol - 1];
      const rowId   = row[osetIdCol - 1];

      if (rowDate instanceof Date && rowDate <= expenseDate) {
        if (!chosenSetDate || rowDate > chosenSetDate) {
          chosenSetDate = rowDate;
          chosenSetId   = rowId;
        }
      }
    });

    if (!chosenSetId) {
      SpreadsheetApp.getActive().toast("No Ownership Set found.", "❌ Cannot Create Charges", 5);
      return;
    }

    const ownershipDetailsSheet = ss.getSheetByName("OWNERSHIP_DETAILS");
    const odSetIdCol = getColumnIndexByHeader(ownershipDetailsSheet, "Ownership_Set_ID");
    const odOwnerCol = getColumnIndexByHeader(ownershipDetailsSheet, "Owner");
    const odPercCol  = getColumnIndexByHeader(ownershipDetailsSheet, "Percentage");

    const odLastRow = ownershipDetailsSheet.getLastRow();
    const odData = ownershipDetailsSheet.getRange(2, 1, odLastRow - 1, ownershipDetailsSheet.getLastColumn()).getValues();

    applicableOwners = odData.filter(row => row[odSetIdCol - 1] === chosenSetId);

    if (splitMethodType === 2) {
      const totalPercent = applicableOwners.reduce((sum, row) => sum + Number(row[odPercCol - 1] || 0), 0);
      if (totalPercent !== 100) {
        SpreadsheetApp.getActive().toast(
          "Ownership percentages do not sum to 100% (current: " + totalPercent + "%).",
          "❌ Cannot Create Charges",
          5
        );
        return;
      }
    }
  }

  // --- 5. Build ownersArray ---
  let ownersArray = [];

  if (splitMethodType === 1 || splitMethodType === 2) {
    const ownershipDetailsSheet = ss.getSheetByName("OWNERSHIP_DETAILS");
    const odOwnerCol = getColumnIndexByHeader(ownershipDetailsSheet, "Owner");
    const odPercCol  = getColumnIndexByHeader(ownershipDetailsSheet, "Percentage");
    const ownerCount = applicableOwners.length;

    ownersArray = applicableOwners.map(row => {
      const personName = row[odOwnerCol - 1];
      let percentage;

      if (splitMethodType === 1) {
        percentage = 1 / ownerCount;
      } else {
        percentage = Number(row[odPercCol - 1]) / 100;
      }

      return { personName, percentage };
    });
  }

  if (splitMethodType === 3) {
    ownersArray = overnightSplits;
  }

  if (splitMethodType === 4) {
    switch (customSplits.type) {
      case "percentage":
        ownersArray = customSplits.splits.map(s => ({
          personName: resolveHelperFromPid(s.pid),
          percentage: s.pct / 100
        }));
        break;

      case "fixed":
        const totalAmt = customSplits.splits.reduce((a, s) => a + s.amt, 0);
        ownersArray = customSplits.splits.map(s => ({
          personName: resolveHelperFromPid(s.pid),
          percentage: s.amt / totalAmt
        }));
        break;

      case "weights":
        const totalW = customSplits.splits.reduce((a, s) => a + s.w, 0);
        ownersArray = customSplits.splits.map(s => ({
          personName: resolveHelperFromPid(s.pid),
          percentage: s.w / totalW
        }));
        break;
    }
  }

  if (ownersArray.length === 0) {
    SpreadsheetApp.getActive().toast("No owners found.", "❌ Cannot Create Charges", 5);
    return;
  }

  // --- 6. Prepare TRANSACTIONS rows ---
  const tIdCol          = getColumnIndexByHeader(transactionsSheet, "ID");
  const tDateCol        = getColumnIndexByHeader(transactionsSheet, "Date");
  const tTypeCol        = getColumnIndexByHeader(transactionsSheet, "Type");
  const tExpenseTypeCol = getColumnIndexByHeader(transactionsSheet, "Expense_Type");
  const tAmountCol      = getColumnIndexByHeader(transactionsSheet, "Amount");
  const tPersonCol      = getColumnIndexByHeader(transactionsSheet, "Person");
  const tAccountCol     = getColumnIndexByHeader(transactionsSheet, "Account");
  const tNotesCol       = getColumnIndexByHeader(transactionsSheet, "Notes");

  let nextTransactionId = 1;

  if (tLastRow > 1) {
    const idRange = transactionsSheet.getRange(2, tIdCol, tLastRow - 1).getValues().flat().filter(n => typeof n === "number");
    if (idRange.length > 0) nextTransactionId = Math.max(...idRange) + 1;
  }

  const TYPE_VALUE = "401 - Charge";
  const rowsToAppend = [];

  // --- 7. N-1 rounding logic + Notes generation ---
  let runningTotal = 0;
  const ownerCount = ownersArray.length;

  ownersArray.forEach((owner, index) => {
    let chargeAmount;

    if (index < ownerCount - 1) {
      const raw = expenseAmount * owner.percentage;
      chargeAmount = Math.round(raw * 100) / 100;
      runningTotal += chargeAmount;
    } else {
      chargeAmount = Math.round((expenseAmount - runningTotal) * 100) / 100;
    }

    chargeAmount = chargeAmount * -1;

    // --- Build Notes ---
    let note = "";
    const pct = (owner.percentage * 100).toFixed(2);
    const base = expenseAmount.toFixed(2);

    if (splitMethodType === 1) {
      note = `${pct}% of ${base} based on equal split between ${ownerCount} owners.`;
    }

    if (splitMethodType === 2) {
      note = `${pct}% of ${base} based on ownership%.`;
    }

    if (splitMethodType === 3) {
      note = `${pct}% of ${base} based on ${owner.stays}/${totalStays} overnight stays.`;
    }

    if (splitMethodType === 4) {
      if (customSplits.type === "percentage") {
        const breakdown = customSplits.splits
          .map(s => `${s.pid} (${s.pct}%)`)
          .join(", ");
        note = `${pct}% of ${base} based on custom split: ${breakdown}.`;
      }

      if (customSplits.type === "fixed") {
        const breakdown = customSplits.splits
          .map(s => `${s.pid} (${s.amt} EUR)`)
          .join(", ");
        note = `${pct}% of ${base} based on custom fixed amounts: ${breakdown}.`;
      }

      if (customSplits.type === "weights") {
        const breakdown = customSplits.splits
          .map(s => `${s.pid} (weight ${s.w})`)
          .join(", ");
        note = `${pct}% of ${base} based on custom weights: ${breakdown}.`;
      }
    }

    // --- Build row ---
    const transactionRow = [];
    transactionRow[tIdCol - 1]          = nextTransactionId++;
    transactionRow[tDateCol - 1]        = expenseDate;
    transactionRow[tTypeCol - 1]        = TYPE_VALUE;
    transactionRow[tExpenseIdCol - 1]   = expenseId;
    transactionRow[tExpenseTypeCol - 1] = expenseType;
    transactionRow[tAmountCol - 1]      = chargeAmount;
    transactionRow[tPersonCol - 1]      = owner.personName;
    transactionRow[tNotesCol - 1]       = note;

    transactionRow[tAccountCol - 1] =
      '=IF(LEFT(INDIRECT("R[0]C[-1]", FALSE), FIND(" -", INDIRECT("R[0]C[-1]", FALSE))-1)="","",' +
      'VLOOKUP(LEFT(INDIRECT("R[0]C[-1]", FALSE), FIND(" -", INDIRECT("R[0]C[-1]", FALSE))-1), PEOPLE!A:D, 4, FALSE))';

    rowsToAppend.push(transactionRow);
  });

  // --- 8. Append rows ---
  if (rowsToAppend.length > 0) {
    const appendStartRow = transactionsSheet.getLastRow() + 1;
    const numCols        = transactionsSheet.getLastColumn();

    const normalizedRows = rowsToAppend.map(rowArr => {
      const row = new Array(numCols).fill(null);
      for (let i = 0; i < rowArr.length; i++) {
        if (rowArr[i] !== undefined) row[i] = rowArr[i];
      }
      return row;
    });

    transactionsSheet.getRange(appendStartRow, 1, normalizedRows.length, numCols).setValues(normalizedRows);
  }

  // --- 9. Sort by Date + ID ---
  const tLastRowFinal = transactionsSheet.getLastRow();
  const tLastColFinal = transactionsSheet.getLastColumn();

  if (tLastRowFinal > 1) {
    const dataRange = transactionsSheet.getRange(2, 1, tLastRowFinal - 1, tLastColFinal);
    dataRange.sort([
      { column: tDateCol, ascending: true },
      { column: tIdCol,   ascending: true }
    ]);
  }

  const statusCol = getColumnIndexByHeader(expensesSheet, "Status");
  // Only set to "Charged" if it is still empty or "Pending"
  const currentStatus = expensesSheet.getRange(expenseRow, statusCol).getValue();
  if (!currentStatus || currentStatus === "Pending") {
    expensesSheet.getRange(expenseRow, statusCol).setValue("Charged");
  }

  SpreadsheetApp.getActive().toast(
    "Charges created with notes for Expense ID " + expenseId,
    "✅ Success",
    4
  );
}


// ---------- CORE ENGINE: CREATE PROVISIONAL CHARGES ----------

function createProvisionalCharges(expensesSheet, expenseRow) {
  const statusCol   = getColumnIndexByHeader(expensesSheet, "Status");
  const startCol    = getColumnIndexByHeader(expensesSheet, "Start_Date");
  const endCol      = getColumnIndexByHeader(expensesSheet, "End_Date");
  const idCol       = getColumnIndexByHeader(expensesSheet, "ID");
  const amountCol   = getColumnIndexByHeader(expensesSheet, "Amount");

  const status    = expensesSheet.getRange(expenseRow, statusCol).getValue();
  const startDate = expensesSheet.getRange(expenseRow, startCol).getValue();
  const endDate   = expensesSheet.getRange(expenseRow, endCol).getValue();
  const expenseId = expensesSheet.getRange(expenseRow, idCol).getValue();
  const amount    = expensesSheet.getRange(expenseRow, amountCol).getValue();

  if (!expenseId || !amount) {
    SpreadsheetApp.getActive().toast("Expense must have ID and Amount before creating provisional charges.", "❌ Error", 5);
    return;
  }

  if (!startDate || !endDate) {
    SpreadsheetApp.getActive().toast("Start_Date and End_Date must be set for provisional charges.", "❌ Error", 5);
    return;
  }

  if (status && status.toString().startsWith("Provisionally Charged")) {
    SpreadsheetApp.getActive().toast("Provisional charges already exist for this expense.", "ℹ️ Info", 5);
    return;
  }

  if (status && status.toString().startsWith("Reconciled")) {
    SpreadsheetApp.getActive().toast("This expense has already been reconciled.", "❌ Error", 5);
    return;
  }

  // Reuse existing 1-phase charge logic.
  // It will create the transactions and (currently) set status to "Charged" if Pending.
  createCharges(expensesSheet, expenseRow);

  // Now overwrite status to indicate it's in 2-phase mode
  expensesSheet.getRange(expenseRow, statusCol).setValue("Provisionally Charged");

  SpreadsheetApp.getActive().toast(
    "Provisional charges created for Expense ID " + expenseId,
    "✅ Provisional",
    4
  );
}


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

  if (!status || !status.toString().startsWith("Provisionally Charged")) {
    SpreadsheetApp.getActive().toast("Only expenses with Status 'Provisionally Charged' can be reconciled.", "❌ Error", 5);
    return;
  }

  const group = getReconciliationGroup(expensesSheet, startDate, endDate, expenseType);
  if (group.length === 0) {
    SpreadsheetApp.getActive().toast("No provisionally charged expenses found in this date window and expense type.", "ℹ️ Info", 5);
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


// ---------- CORE ENGINE: DELETE CHARGES ----------

function deleteCharges(expensesSheet, expenseRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Read the Expense_ID from the EXPENSES row ---
  const expenseIdCol = getColumnIndexByHeader(expensesSheet, "ID");
  const expenseId    = expensesSheet.getRange(expenseRow, expenseIdCol).getValue();

  if (!expenseId) {
    SpreadsheetApp.getActive().toast(
      "Cannot delete charges: Expense row has no Expense ID.",
      "❌ Cannot Delete Charges",
      5
    );
    return;
  }

  // --- 2. Access TRANSACTIONS sheet ---
  const transactionsSheet = ss.getSheetByName("TRANSACTIONS");
  const tLastRow          = transactionsSheet.getLastRow();
  if (tLastRow < 2) {
    SpreadsheetApp.getActive().toast(
      "No transactions exist to delete.",
      "ℹ️ Information",
      4
    );
    return;
  }

  const tExpenseIdCol = getColumnIndexByHeader(transactionsSheet, "Expense_ID");

  // --- 3. Read all Transaction rows ---
  const tData = transactionsSheet
    .getRange(2, 1, tLastRow - 1, transactionsSheet.getLastColumn())
    .getValues();

  // --- 4. Find all rows where Expense_ID matches ---
  const rowsToDelete = [];
  tData.forEach((row, index) => {
    const rowExpenseId = row[tExpenseIdCol - 1];
    if (rowExpenseId === expenseId) {
      rowsToDelete.push(index + 2); // +2 because data starts at row 2
    }
  });

  if (rowsToDelete.length === 0) {
    SpreadsheetApp.getActive().toast(
      "No charges found for Expense ID " + expenseId + ".",
      "ℹ️ Information",
      4
    );
    return;
  }

  // --- 5. Delete rows bottom‑up to avoid shifting ---
  rowsToDelete.reverse().forEach(rowNum => {
    transactionsSheet.deleteRow(rowNum);
  });

  const statusCol = getColumnIndexByHeader(expensesSheet, "Status");
  expensesSheet.getRange(expenseRow, statusCol).setValue("Pending");

  SpreadsheetApp.getActive().toast(
    "Charges deleted for Expense ID " + expenseId + ".",
    "✅ Deleted",
    4
  );
}


// ---------- CORE ENGINE: DELETE CHARGES FOR EXPENSE ----------

function deleteChargesForExpense(expensesSheet, expenseRow) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const statusCol = getColumnIndexByHeader(expensesSheet, "Status");
  const idCol     = getColumnIndexByHeader(expensesSheet, "ID");

  const rawStatus = expensesSheet.getRange(expenseRow, statusCol).getValue();
  const status = String(rawStatus)
    .replace(/\u00A0/g, " ")   // remove non-breaking spaces
    .replace(/\s+/g, " ")      // collapse weird whitespace
    .trim()
    .toLowerCase();
  const expenseId = expensesSheet.getRange(expenseRow, idCol).getValue();

  if (!expenseId) {
    SpreadsheetApp.getActive().toast("Cannot delete charges: Expense has no ID.", "❌ Error", 5);
    return;
  }

  if (!status || status === "pending") {
    SpreadsheetApp.getActive().toast("This expense has no charges to delete.", "ℹ️ Info", 4);
    return;
  }

  if (status.startsWith("reconciled")) {
    SpreadsheetApp.getActive().toast(
      "This expense has already been reconciled. Reconciliation adjustments must not be deleted.\n" +
      "To correct distribution, run 'Reconcile Charges' again for the period.",
      "❌ Blocked",
      7
    );
    return;
  }

  if (status === "charged") {
    deleteCharges(expensesSheet, expenseRow)
    //expensesSheet.getRange(expenseRow, statusCol).setValue("Pending");
    //SpreadsheetApp.getActive().toast("Charges deleted for Expense ID " + expenseId + ".", "✅ Deleted", 4);
    return;
  }

  if (status.startsWith("provisionally charged")) {
    const transactionsSheet = ss.getSheetByName("TRANSACTIONS");
    const tLastRow  = transactionsSheet.getLastRow();
    if (tLastRow < 2) {
      SpreadsheetApp.getActive().toast("No transactions exist to delete.", "ℹ️ Info", 4);
      return;
    }

    const tExpenseIdCol = getColumnIndexByHeader(transactionsSheet, "Expense_ID");
    const tData = transactionsSheet.getRange(2, 1, tLastRow - 1, transactionsSheet.getLastColumn()).getValues();

    const rowsToDelete = [];
    tData.forEach((row, index) => {
      if (String(row[tExpenseIdCol - 1]) === String(expenseId)) {
        rowsToDelete.push(index + 2);
      }
    });

    if (rowsToDelete.length === 0) {
      SpreadsheetApp.getActive().toast("No charges found for this expense.", "ℹ️ Info", 4);
      return;
    }

    rowsToDelete.reverse().forEach(r => transactionsSheet.deleteRow(r));

    expensesSheet.getRange(expenseRow, statusCol).setValue("Pending");

    SpreadsheetApp.getActive().toast("Provisional charges deleted for Expense ID " + expenseId + ".", "✅ Deleted", 4);
    return;
  }

  SpreadsheetApp.getActive().toast("Unrecognized Status state; no charges deleted.", "ℹ️ Info", 4);
}


// ---------- UTILITIES & HELPER FUNCTIONS ----------

function getColumnIndexByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index   = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error("Header '" + headerName + "' not found on sheet '" + sheet.getName() + "'.");
  }
  return index + 1; // convert 0-based to 1-based index
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

    if (!status || !status.toString().startsWith("Provisionally Charged")) continue;
    if (!sameDate(rowStart, startDate) || !sameDate(rowEnd, endDate)) continue;
    if (rowType !== expenseType) continue;

    result.push({
      rowIndex: r + 1,
      id: row[idIdx],
      amount: Number(row[amountIdx]) || 0,
    });
  }

  return result;
}

function sameDate(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
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
