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
