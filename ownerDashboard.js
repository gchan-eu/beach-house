/**
 * Updates the OWNER_DASHBOARD sheet, including the Owner_Transactions table.
 */
function updateOwnerDashboard() {
  updateOwnerTransactions();
  updateOwnerOvernightStays();
}

/**
 * Updates the Owner_Transactions table in the OWNER_DASHBOARD sheet.
 * Table headers are at row 6. Transaction rows are filtered by owner account number (from B1).
 */
function updateOwnerTransactions() {
  Logger.log('updateOwnerTransactions() running');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("OWNER_DASHBOARD");
  const peopleSheet = ss.getSheetByName("PEOPLE");
  const transactionsSheet = ss.getSheetByName("TRANSACTIONS");
  const expensesSheet = ss.getSheetByName("EXPENSES");
  const expenseTypesSheet = ss.getSheetByName("EXPENSE_TYPES");

  if (!dashboardSheet || !peopleSheet || !transactionsSheet || !expensesSheet || !expenseTypesSheet) {
    SpreadsheetApp.getActive().toast("Required sheet missing.", "❌ Error", 5);
    return;
  }

  // Get owner name from B1
  const ownerName = dashboardSheet.getRange("B1").getValue();
  if (!ownerName) {
    SpreadsheetApp.getActive().toast("Owner name not set in B1", "❌ Error", 5);
    return;
  }

  // Get account number for owner
  const peopleData = peopleSheet.getDataRange().getValues();
  const peopleHeaders = peopleData[0];
  const nameCol = peopleHeaders.indexOf("Name");
  const accountCol = peopleHeaders.indexOf("Account_Number");
  let ownerAccount = null;
  peopleData.slice(1).forEach(row => {
    if (row[nameCol] == ownerName) {
      ownerAccount = row[accountCol];
    }
  });
  if (!ownerAccount) {
    SpreadsheetApp.getActive().toast("Owner account not found for " + ownerName, "❌ Error", 5);
    return;
  }

  // Get transactions for owner
  const tData = transactionsSheet.getDataRange().getValues();
  const tHeaders = tData[0];
  const tAccountCol = tHeaders.indexOf("Account");
  const tDateCol = tHeaders.indexOf("Date");
  const tTypeCol = tHeaders.indexOf("Type");
  const tAmountCol = tHeaders.indexOf("Amount");
  const tExpenseIdCol = tHeaders.indexOf("Expense_ID");
  const tNotesCol = tHeaders.indexOf("Notes");

  // Get expenses
  const expData = expensesSheet.getDataRange().getValues();
  const expHeaders = expData[0];
  const expIdCol = expHeaders.indexOf("ID");
  const expCodeCol = expHeaders.indexOf("Code");
  const expDateCol = expHeaders.indexOf("Date");
  const expNotesCol = expHeaders.indexOf("Notes");
  const expReceiptCol = expHeaders.indexOf("Receipt");
  // Get rich text values for the Receipt column only (excluding header)
  const expReceiptRichText = expensesSheet.getRange(2, expReceiptCol + 1, expData.length - 1, 1).getRichTextValues().map(row => row[0]);

  // Get expense types
  const expTypeData = expenseTypesSheet.getDataRange().getValues();
  const expTypeHeaders = expTypeData[0];
  const expTypeCodeCol = expTypeHeaders.indexOf("Code");
  const expTypeCatCol = expTypeHeaders.indexOf("Category");
  const expTypeDescCol = expTypeHeaders.indexOf("Description");

  // Build table headers (remove Expense_Date and Notes)
  const tableHeaders = [
    "Date", "Type", "Category", "Description", "Amount", "Expense_Notes", "Receipt"
  ];

  // Clear only the data rows of the Owner_Transactions table (not deleting rows)
  // Assume table header is at row 6, data starts at row 7
  const transactionsHeaderRow = 6;
  const transactionsDataStart = 7;
  // Find the current position of the overnight stays table header (in case it was pushed down previously)
  const overnightHeaders = [
    'Start_Date', 'End_date', 'Days', 'Person_Count', 'Stays', 'Notes'
  ];
  let overnightHeaderRow = null;
  let overnightTableRows = [];
  let maxRows = dashboardSheet.getLastRow();
  // Find the overnight stays table header and collect all its rows (header + data)
  for (let r = transactionsDataStart; r <= maxRows; r++) {
    const rowVals = dashboardSheet.getRange(r, 1, 1, overnightHeaders.length).getValues()[0];
    if (
      rowVals[0] === overnightHeaders[0] &&
      rowVals[1] === overnightHeaders[1] &&
      rowVals[2] === overnightHeaders[2] &&
      rowVals[3] === overnightHeaders[3] &&
      rowVals[4] === overnightHeaders[4] &&
      rowVals[5] === overnightHeaders[5]
    ) {
      overnightHeaderRow = r;
      // Collect all rows for the overnight table (header + data)
      overnightTableRows = [rowVals];
      // Collect data rows until a blank row or end of sheet
      for (let rr = r + 1; rr <= maxRows; rr++) {
        const dataVals = dashboardSheet.getRange(rr, 1, 1, overnightHeaders.length).getValues()[0];
        const isEmpty = dataVals.every(cell => cell === '' || cell === null);
        if (isEmpty) break;
        overnightTableRows.push(dataVals);
      }
      break;
    }
  }

  // Remove the overnight stays table (header + data rows) from the sheet
  if (overnightHeaderRow !== null) {
    dashboardSheet.deleteRows(overnightHeaderRow, overnightTableRows.length);
    maxRows = dashboardSheet.getLastRow();
  }

  // Calculate how many data rows are needed for the new transactions
  let ownerTxns = tData.slice(1).filter(row => row[tAccountCol] == ownerAccount);
  // Sort transactions by Date descending, then by ID descending
  const tIdCol = tHeaders.indexOf("ID");
  ownerTxns.sort((a, b) => {
    const dateA = new Date(a[tDateCol]);
    const dateB = new Date(b[tDateCol]);
    if (dateB - dateA !== 0) {
      return dateB - dateA;
    }
    // If dates are equal, sort by ID descending (assuming numeric or string comparison)
    const idA = a[tIdCol];
    const idB = b[tIdCol];
    if (idA == null && idB == null) return 0;
    if (idA == null) return 1;
    if (idB == null) return -1;
    // If IDs are numbers, compare numerically; else, compare as strings
    if (!isNaN(Number(idA)) && !isNaN(Number(idB))) {
      return Number(idB) - Number(idA);
    }
    return (idB + '').localeCompare(idA + '');
  });
  const numDataRows = ownerTxns.length;

  // Calculate and update sums in E1-E4
  let sumAll = 0;
  let sumDeposit = 0;
  let sumWithdrawal = 0;
  let sumChargeRecon = 0;
  ownerTxns.forEach(row => {
    const amount = parseFloat(row[tAmountCol]) || 0;
    const type = row[tTypeCol];
    sumAll += amount;
    if (type === "101 - Deposit") sumDeposit += amount;
    if (type === "201 - Withdrawal") sumWithdrawal += amount;
    if (type === "401 - Charge" || type === "402 - Reconciliation") sumChargeRecon += amount;
  });
  dashboardSheet.getRange("E1").setValue(sumAll);
  dashboardSheet.getRange("E2").setValue(sumDeposit);
  dashboardSheet.getRange("E3").setValue(sumWithdrawal);
  dashboardSheet.getRange("E4").setValue(sumChargeRecon);

  // Delete all data rows in the Owner_Transactions table (from row 7 up to the first empty row or end of sheet)
  let deleteStart = transactionsDataStart;
  let deleteEnd = deleteStart - 1;
  for (let r = transactionsDataStart; r <= maxRows; r++) {
    const rowVals = dashboardSheet.getRange(r, 1, 1, 1).getValues()[0];
    if (rowVals[0] === '' || rowVals[0] === null) {
      break;
    }
    deleteEnd = r;
  }
  if (deleteEnd >= deleteStart) {
    dashboardSheet.deleteRows(deleteStart, deleteEnd - deleteStart + 1);
  }

  // After deletion, the first available row for transactions data is transactionsDataStart
  // Insert enough rows for the new transactions
  if (ownerTxns.length > 0) {
    dashboardSheet.insertRowsBefore(transactionsDataStart, ownerTxns.length);
  }

  // After inserting, ensure there are exactly 3 blank rows between transactions and overnight table
  let gapStartRow = transactionsDataStart + ownerTxns.length;
  let gapRows = 0;
  maxRows = dashboardSheet.getLastRow();
  // Count how many blank rows currently exist after transactions
  while (gapStartRow + gapRows <= maxRows) {
    const rowVals = dashboardSheet.getRange(gapStartRow + gapRows, 1, 1, 1).getValues()[0];
    if (rowVals[0] === '' || rowVals[0] === null) {
      gapRows++;
    } else {
      break;
    }
  }
  if (gapRows < 3) {
    dashboardSheet.insertRowsBefore(gapStartRow + gapRows, 3 - gapRows);
  } else if (gapRows > 3) {
    dashboardSheet.deleteRows(gapStartRow, gapRows - 3);
  }

  // Insert the overnight table after the 3-row gap
  let overnightInsertRow = gapStartRow + 3;
  if (overnightTableRows.length > 0) {
    dashboardSheet.insertRowsBefore(overnightInsertRow, overnightTableRows.length);
    dashboardSheet.getRange(overnightInsertRow, 1, overnightTableRows.length, overnightHeaders.length).setValues(overnightTableRows);
  }

  // ownerTxns is already defined above, so do not redeclare it below

  // Filter transactions for owner (already sorted above)
  // ownerTxns = tData.slice(1).filter(row => row[tAccountCol] == ownerAccount);

  // Build table rows
  const tableRows = ownerTxns.map(row => {
    const txnDate = row[tDateCol];
    const txnType = row[tTypeCol];
    const txnAmount = row[tAmountCol];
    const txnExpenseId = row[tExpenseIdCol];
    const txnNotes = row[tNotesCol];

    // Find expense row index by ID
    let expRowIdx = expData.findIndex(r => r[expIdCol] == txnExpenseId);
    let expRow = expRowIdx >= 0 ? expData[expRowIdx] : null;
    let expCode = expRow ? expRow[expCodeCol] : "";
    let expNotes = expRow ? expRow[expNotesCol] : "";
    // Use rich text for receipt if available (expRowIdx-1 because expReceiptRichText starts from row 2)
    let expReceipt = "";
    if (expRowIdx > 0 && expReceiptRichText[expRowIdx - 1]) {
      const rich = expReceiptRichText[expRowIdx - 1];
       if (rich.getLinkUrl()) {
        expReceipt = rich.getLinkUrl();
      } else {
        expReceipt = rich.getText();
      }
    } else if (expRowIdx === 0 && expReceiptRichText[0]) {
      const rich = expReceiptRichText[0];
      if (rich.getLinkUrl()) {
        expReceipt = rich.getLinkUrl();
      } else {
        expReceipt = rich.getText();
      }
    } else if (expRow && typeof expReceiptCol === 'number') {
      expReceipt = expRow[expReceiptCol];
    }
    if (expReceipt && typeof expReceipt === "string" && expReceipt.trim() !== "") {
    }

    let expTypeRow = expTypeData.find(r => r[expTypeCodeCol] == expCode);
    let expCategory = expTypeRow ? expTypeRow[expTypeCatCol] : "";
    let expDescription = expTypeRow ? expTypeRow[expTypeDescCol] : "";

    let receiptCell = "";
    let isLink = false;
    if (expReceipt && typeof expReceipt === "string" && expReceipt.trim() !== "") {
      if (/^https?:\/\//i.test(expReceipt.trim())) {
        receiptCell = `=HYPERLINK("${expReceipt.trim()}", "View")`;
        isLink = true;
      } else {
        receiptCell = expReceipt;
      }
    }

    // Remove Notes from output row
    return [
      txnDate,
      txnType,
      expCategory,
      expDescription,
      txnAmount,
      expNotes,
      receiptCell
    ];
  });

  // Write table to dashboard, starting at row 6
  dashboardSheet.getRange(6, 1, 1, tableHeaders.length).setValues([tableHeaders]);
  if (tableRows.length > 0) {
    // Write all columns except Receipt with setValues (Notes removed, so Receipt is now col 6)
    const values = tableRows.map(row => row.map((cell, i) => (i === 6 ? null : cell)));
    dashboardSheet.getRange(7, 1, tableRows.length, tableHeaders.length).setValues(values);

    // Write Receipt column cell-by-cell: setFormula for links, setValue for plain text
    for (let i = 0; i < tableRows.length; i++) {
      const cellValue = tableRows[i][6]; // Receipt is now col 6
      const cell = dashboardSheet.getRange(7 + i, 7); // Column G
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('=HYPERLINK')) {
        cell.setFormula(cellValue);
      } else if (cellValue) {
        cell.setValue(cellValue);
      } else {
        cell.setValue('');
      }
    }

    // Add Notes as a note to the Amount column with custom logic
    for (let i = 0; i < tableRows.length; i++) {
      const typeValue = tableRows[i][1]; // Type column
      // Expense_Date is no longer a column, so get it from expRow
      let expenseDate = "";
      let expNotesVal = "";
      let expStartDate = "";
      let expEndDate = "";
      let expAmount = "";
      // Find expense row index by ID again
      const txnExpenseId = ownerTxns[i][tExpenseIdCol];
      let expRowIdx = expData.findIndex(r => r[expIdCol] == txnExpenseId);
      let expRow = expRowIdx >= 0 ? expData[expRowIdx] : null;
      if (expRow) {
        expenseDate = expRow[expDateCol];
        expNotesVal = expRow[expNotesCol];
        const expStartDateCol = expHeaders.indexOf("Start_Date");
        const expEndDateCol = expHeaders.indexOf("End_Date");
        if (expStartDateCol >= 0) expStartDate = expRow[expStartDateCol];
        if (expEndDateCol >= 0) expEndDate = expRow[expEndDateCol];
        const expAmountCol = expHeaders.indexOf("Amount");
        if (expAmountCol >= 0) expAmount = expRow[expAmountCol];
      }
      const expenseNotes = tableRows[i][5]; // Expense_Notes column
      // Notes column is deprecated, so get transactionNotes from ownerTxns
      const transactionNotes = ownerTxns[i][tNotesCol];
      let noteValue = '';
      function formatToDDMMYY(dateVal) {
        if (!dateVal) return '';
        let iso = '';
        if (dateVal instanceof Date) {
          iso = formatDate(dateVal); // yyyy-MM-dd
        } else if (typeof dateVal === 'string' && dateVal.match(/^\d{4}-\d{2}-\d{2}/)) {
          iso = dateVal;
        } else {
          return dateVal;
        }
        const [y, m, d] = iso.split('-');
        return `${d}/${m}/${y.slice(-2)}`;
      }
      if (typeValue === '401 - Charge' || typeValue === '402 - Reconciliation') {
        noteValue =
          'Expense Date: ' + (expenseDate ? formatToDDMMYY(expenseDate) : '') + '\n';
        // Format expense amount as euro currency (2 decimals)
        let formattedAmount = '';
        if (expAmount !== undefined && expAmount !== null && expAmount !== '') {
          let num = Number(expAmount);
          if (!isNaN(num)) {
            formattedAmount = num.toLocaleString('en-US', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 });
          } else {
            formattedAmount = expAmount;
          }
        }
        noteValue += 'Expense Amount: ' + formattedAmount + '\n';
        // Add Consumption Period if both Start_Date and End_Date are present
        if (expStartDate && expEndDate) {
          noteValue += 'Period: ' + formatToDDMMYY(expStartDate) + ' - ' + formatToDDMMYY(expEndDate) + '\n';
        }
        noteValue += 'Notes: ' + (transactionNotes ? transactionNotes : '');
      } else if (typeValue === '101 - Deposit') {
        noteValue = 'Notes: ' + (transactionNotes ? transactionNotes : '');
      } else {
        noteValue = transactionNotes ? transactionNotes : '';
      }
      const amountCell = dashboardSheet.getRange(7 + i, 5); // Amount column (E)
      if (noteValue && typeof noteValue === 'string' && noteValue.trim() !== '') {
        amountCell.setNote(noteValue);
      } else {
        amountCell.setNote('');
      }
    }
  }
}

/**
 * Updates the Overnight Stays table in the OWNER_DASHBOARD sheet, placing it dynamically below the transactions table.
 * Columns: Start_Date, End_date, Days, Person_Count, Stays, Notes
 */
function updateOwnerOvernightStays() {
  Logger.log('updateOwnerOvernightStays() running');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("OWNER_DASHBOARD");
  const peopleSheet = ss.getSheetByName("PEOPLE");
  const overnightSheet = ss.getSheetByName("OVERNIGHT_STAYS");
  if (!dashboardSheet || !peopleSheet || !overnightSheet) return;

  // Define headers for the prebuilt table
  const overnightHeaders = [
    'Start_Date', 'End_Date', 'Days', 'Person_Count', 'Stays', 'Notes'
  ];

  // Find the Owner_Overnight_Stays table header row in the dashboard sheet
  let headerRow = null;
  const scanRows = 200; // Scan first 200 rows to find the header, regardless of blank lines
  // Log the first 200 rows' values for columns 1-6 for debugging
  let debugRows = [];
  for (let r = 1; r <= scanRows; r++) {
    const rowVals = dashboardSheet.getRange(r, 1, 1, overnightHeaders.length).getValues()[0];
    debugRows.push({row: r, values: rowVals});
    if (
      rowVals[0] === overnightHeaders[0] &&
      rowVals[1] === overnightHeaders[1] &&
      rowVals[2] === overnightHeaders[2] &&
      rowVals[3] === overnightHeaders[3] &&
      rowVals[4] === overnightHeaders[4] &&
      rowVals[5] === overnightHeaders[5]
    ) {
      headerRow = r;
      Logger.log('Owner_Overnight_Stays header found at row: ' + headerRow);
      break;
    }
  }
  if (!headerRow) {
    Logger.log('No header row found for the Owner_Overnight_Stays table');
    // Log the first 200 rows for manual inspection
    debugRows.forEach(obj => {
      Logger.log('Row ' + obj.row + ': ' + JSON.stringify(obj.values));
    });
    return;
  }

  // Get owner name from B1
  const ownerName = dashboardSheet.getRange("B1").getValue();
  if (!ownerName) return;

  // Get People data
  const peopleData = peopleSheet.getDataRange().getValues();
  const peopleHeaders = peopleData[0];
  const nameCol = peopleHeaders.indexOf("Name");
  const codeCol = peopleHeaders.indexOf("Code");
  let personCode = null;
  peopleData.slice(1).forEach(row => {
    if (row[nameCol] == ownerName) {
      personCode = row[codeCol];
    }
  });
  if (!personCode) return;

  // Get Overnight_Stays data
  const overnightData = overnightSheet.getDataRange().getValues();
  const overnightHeadersRow = overnightData[0];
  const osPersonIdCol = overnightHeadersRow.indexOf("Person_ID");
  const osStartDateCol = overnightHeadersRow.indexOf("Start_Date");
  const osEndDateCol = overnightHeadersRow.indexOf("End_Date");
  const osDaysCol = overnightHeadersRow.indexOf("Days");
  const osPersonCountCol = overnightHeadersRow.indexOf("Person_Count");
  const osStaysCol = overnightHeadersRow.indexOf("Total_Stays");
  const osNotesCol = overnightHeadersRow.indexOf("Notes");

  // Filter rows for this person and sort by Start_Date descending (latest on top)
  let overnightRows = overnightData.slice(1)
    .filter(row => row[osPersonIdCol] == personCode)
    .sort((a, b) => {
      const dateA = new Date(a[osStartDateCol]);
      const dateB = new Date(b[osStartDateCol]);
      return dateB - dateA;
    })
    .map(row => [
      row[osStartDateCol],
      row[osEndDateCol],
      row[osDaysCol],
      row[osPersonCountCol],
      row[osStaysCol],
      row[osNotesCol]
    ]);


  // Delete all data rows under the Owner_Overnight_Stays table header (no footer logic, no empty rows)
  let dataStartRow = headerRow + 1;
  let lastRow = dashboardSheet.getLastRow();
  // Find the last non-empty row for this table (stop at first blank row)
  let dataEndRow = dataStartRow - 1;
  for (let r = dataStartRow; r <= lastRow; r++) {
    const rowVals = dashboardSheet.getRange(r, 1, 1, overnightHeaders.length).getValues()[0];
    const isEmpty = rowVals.every(cell => cell === '' || cell === null);
    if (isEmpty) break;
    dataEndRow = r;
  }
  // Delete all data rows (remove rows from the sheet)
  if (dataEndRow >= dataStartRow) {
    dashboardSheet.deleteRows(dataStartRow, dataEndRow - dataStartRow + 1);
  }
  // Insert new rows for new data if needed
  if (overnightRows.length > 0) {
    dashboardSheet.insertRowsBefore(dataStartRow, overnightRows.length);
    dashboardSheet.getRange(dataStartRow, 1, overnightRows.length, overnightHeaders.length).setValues(overnightRows);
  }
}
