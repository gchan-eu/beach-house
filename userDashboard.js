/**
 * Trigger: Updates the dashboard whenever the DASHBOARD sheet is selected.
 * To use, set this as an installable onSelectionChange trigger in the Apps Script UI.
 * (Simple triggers cannot modify other sheets, so this must be installable.)
 * @param {GoogleAppsScript.Events.SheetsOnSelectionChange} e
 */
function onSelectionChange(e) {
	if (!e) return;
	var sheet = e.range.getSheet();
	/**
	 * Updates the DASHBOARD sheet with wallet balances, deposits, charges, days, and stays for each person,
	 * filtered by date range if 'Date From' and 'Date To' are set in cells B1 and B2.
	 */
	function updateDashboard() {
		updateDashboard();
	}
}
/**
 * Updates the DASHBOARD sheet with wallet balances, deposits, charges, days, and stays for each person.
 * Assumes PEOPLE, TRANSACTIONS, and OVERNIGHT_STAYS sheets exist.
 */
function updateDashboard() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dashboardSheet = ss.getSheetByName("DASHBOARD");
	const peopleSheet = ss.getSheetByName("PEOPLE");
	const transactionsSheet = ss.getSheetByName("TRANSACTIONS");
	const overnightSheet = ss.getSheetByName("OVERNIGHT_STAYS");

	if (!dashboardSheet || !peopleSheet || !transactionsSheet || !overnightSheet) {
		SpreadsheetApp.getActive().toast("Required sheet missing.", "❌ Error", 5);
		return;
	}

	// Read date filter from dashboard sheet (B1: Date From, B2: Date To)
	let dateFrom = dashboardSheet.getRange("B1").getValue();
	let dateTo = dashboardSheet.getRange("B2").getValue();
	if (!(dateFrom instanceof Date)) dateFrom = null;
	if (!(dateTo instanceof Date)) dateTo = null;

	// Get people
	const peopleData = peopleSheet.getDataRange().getValues();
	const peopleHeaders = peopleData[0];
	const codeCol = peopleHeaders.indexOf("Code");
	const nameCol = peopleHeaders.indexOf("Helper") !== -1 ? peopleHeaders.indexOf("Helper") : peopleHeaders.indexOf("Name");
	const personList = peopleData.slice(1).map(row => ({
		code: row[codeCol],
		name: row[nameCol]
	}));

	// Get transactions
	const tData = transactionsSheet.getDataRange().getValues();
	const tHeaders = tData[0];
	const tPersonCol = tHeaders.indexOf("Person");
	const tAmountCol = tHeaders.indexOf("Amount");
	const tTypeCol = tHeaders.indexOf("Type");
	const tDateCol = tHeaders.indexOf("Date");

	// Get overnight stays
	const osData = overnightSheet.getDataRange().getValues();
	const osHeaders = osData[0];
	const osPersonIdCol = osHeaders.indexOf("Person_ID");
	const osPersonCol = osHeaders.indexOf("Person");
	const osStartCol = osHeaders.indexOf("Start_Date");
	const osEndCol = osHeaders.indexOf("End_Date");
	const osCountCol = osHeaders.indexOf("Person_Count");

	// Prepare dashboard rows
	const dashboardRows = [
		["Person", "Wallet Balance", "Total Deposited", "Total Charges", "Total Days", "Total Stays"]
	];

	personList.forEach(person => {
		// Transactions
		let totalDeposited = 0;
		let totalCharges = 0;
		let walletBalance = 0;
		tData.slice(1).forEach(row => {
			if (row[tPersonCol] == person.name) {
				const amt = Number(row[tAmountCol]) || 0;
				const type = row[tTypeCol] || "";
				const tDate = row[tDateCol];
				// Filter by date range if set
				if (dateFrom && tDate instanceof Date && tDate < dateFrom) return;
				if (dateTo && tDate instanceof Date && tDate > dateTo) return;
				if (String(type).toLowerCase().includes("deposit")) {
					totalDeposited += amt;
				} else if (String(type).toLowerCase().includes("charge") || String(type).toLowerCase().includes("reconciliation")) {
					totalCharges += amt;
				}
				walletBalance += amt;
			}
		});

		// Overnight stays
		let totalDays = 0;
		let totalStays = 0;
		osData.slice(1).forEach(row => {
			if (row[osPersonIdCol] == person.code || row[osPersonCol] == person.name) {
				const start = row[osStartCol];
				const end = row[osEndCol];
				const count = Number(row[osCountCol]) || 1;
				// Filter by date range if set
				if (dateFrom && start instanceof Date && end instanceof Date && end < dateFrom) return;
				if (dateTo && start instanceof Date && end instanceof Date && start > dateTo) return;
				if (start instanceof Date && end instanceof Date) {
					// Calculate overlap with filter
					let overlapStart = dateFrom && start < dateFrom ? dateFrom : start;
					let overlapEnd = dateTo && end > dateTo ? dateTo : end;
					const days = Math.floor((overlapEnd - overlapStart) / (1000 * 60 * 60 * 24)) + 1;
					if (days > 0) {
						totalDays += days * count;
						totalStays += count;
					}
				}
			}
		});

		dashboardRows.push([
			person.name,
			round2(walletBalance),
			round2(totalDeposited),
			round2(totalCharges),
			totalDays,
			totalStays
		]);
	});

	// Write to dashboard
	dashboardSheet.clearContents();
	dashboardSheet.getRange(1, 1, dashboardRows.length, dashboardRows[0].length).setValues(dashboardRows);
}
// ...existing code from dashboard.js...
