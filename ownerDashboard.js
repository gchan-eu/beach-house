/**
 * Main entry point: refreshes the OWNER_DASHBOARD sheet based on filters.
 */
function updateOwnerDashboard() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName('OWNER_DASHBOARD');
  const txSheet = ss.getSheetByName('TRANSACTIONS');
  const staysSheet = ss.getSheetByName('OVERNIGHT_STAYS');

  if (!dash || !txSheet || !staysSheet) {
    throw new Error('One or more required sheets (OWNER_DASHBOARD, TRANSACTIONS, OVERNIGHT_STAYS) are missing.');
  }

  // 1. Read filters from control panel
  const owner = dash.getRange('B1').getValue();
  const startDate = dash.getRange('B2').getValue();
  const endDate = dash.getRange('B3').getValue();

  if (!owner || !startDate || !endDate) {
    // Clear outputs and exit gracefully if filters are incomplete
    clearDashboardTables_(dash);
    clearDashboardKPIs_(dash);
    return;
  }

  // 2. Get data from Transactions
  const txValues = txSheet.getDataRange().getValues();
  const txHeader = txValues.shift(); // remove header row

  // 3. Get data from OvernightStays
  const staysValues = staysSheet.getDataRange().getValues();
  const staysHeader = staysValues.shift(); // remove header row

  // 4. Filter Transactions for owner + date range
  const filteredTransactions = txValues.filter(row => {
    const date = row[0];     // A: Date
    const rowOwner = row[4]; // E: Owner
    if (!(date instanceof Date)) return false;
    return rowOwner === owner && date >= startDate && date <= endDate;
  });

  // 5. Filter Overnight Stays for owner + date range
  // We consider a stay relevant if it overlaps the [startDate, endDate] interval.
  const filteredStaysRaw = staysValues.filter(row => {
    const start = row[0];     // A: Start date
    const end = row[1];       // B: End date
    const rowOwner = row[2];  // C: Owner
    if (!(start instanceof Date) || !(end instanceof Date)) return false;
    const overlaps =
      rowOwner === owner &&
      end >= startDate &&
      start <= endDate;
    return overlaps;
  });

  // 6. Compute Days and Stays for each filtered stay
  const staysWithMetrics = filteredStaysRaw.map(row => {
    const start = row[0];        // Start date
    const end = row[1];          // End date
    const persons = row[3];      // D: Persons
    const notes = row[5] || '';  // F: Notes (optional)

    const msPerDay = 1000 * 60 * 60 * 24;
    const days = Math.round((end - start) / msPerDay) + 1;
    const stays = days * (persons || 0);

    return [start, end, days, persons, stays, notes];
  });

  // 7. Write data to dashboard

  clearDashboardTables_(dash);

  // Table 1: Transactions – write starting at row 7, columns A–F
  if (filteredTransactions.length > 0) {
    dash.getRange(7, 1, filteredTransactions.length, filteredTransactions[0].length)
      .setValues(filteredTransactions);
  }

  // Table 2: Stays – write starting at row 31, columns A–F
  if (staysWithMetrics.length > 0) {
    dash.getRange(31, 1, staysWithMetrics.length, staysWithMetrics[0].length)
      .setValues(staysWithMetrics);
  }

  // 8. Compute and write KPIs
  writeDashboardKPIs_(dash, filteredTransactions, staysWithMetrics);
}

/**
 * Clears the output areas for both tables on OWNER_DASHBOARD.
 */
function clearDashboardTables_(dash) {
  // Clear Transactions table area (A7:F2000)
  dash.getRange('A7:F2000').clearContent();

  // Clear Stays table area (A31:F2000)
  dash.getRange('A31:F2000').clearContent();
}

/**
 * Clears the KPI cells.
 */
function clearDashboardKPIs_(dash) {
  dash.getRange('E1:E3').clearContent();
}

/**
 * Computes and writes KPIs to OWNER_DASHBOARD.
 * KPIs:
 *  - E1: Total amount (sum of D in transactions)
 *  - E2: Total days (sum of C in staysWithMetrics)
 *  - E3: Total stays (sum of E in staysWithMetrics)
 */
function writeDashboardKPIs_(dash, transactions, staysWithMetrics) {
  // Total amount from transactions
  const amountIndex = 3; // D: Amount
  const totalAmount = transactions.reduce((sum, row) => {
    const val = Number(row[amountIndex]) || 0;
    return sum + val;
  }, 0);

  // Total days and stays from staysWithMetrics
  let totalDays = 0;
  let totalStays = 0;

  staysWithMetrics.forEach(row => {
    const days = Number(row[2]) || 0;  // C: Days
    const stays = Number(row[4]) || 0; // E: Stays
    totalDays += days;
    totalStays += stays;
  });

  dash.getRange('E1').setValue(totalAmount);
  dash.getRange('E2').setValue(totalDays);
  dash.getRange('E3').setValue(totalStays);
}

/**
 * Optional: auto-refresh when filters change (Owner, Start date, End date).
 * Attach this to an installable onEdit trigger if you want.
 */
function onEditOwnerDashboard_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'OWNER_DASHBOARD') return;

  const row = range.getRow();
  const col = range.getColumn();

  // If edit is in B1, B2, or B3, refresh dashboard
  if (row >= 1 && row <= 3 && col === 2) {
    updateOwnerDashboard();
  }
}
