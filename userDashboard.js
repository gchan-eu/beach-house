/**
 * Updates the USER_DASHBOARD sheet's User_Dashboard table with account balances, deposits, withdrawals, charges, and days for each person.
 * Assumes PEOPLE, TRANSACTIONS, and OVERNIGHT_STAYS sheets exist.
 */
function updateUserDashboard() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dashboardSheet = ss.getSheetByName("USER_DASHBOARD");
	const peopleSheet = ss.getSheetByName("PEOPLE");
	const transactionsSheet = ss.getSheetByName("TRANSACTIONS");
	const overnightSheet = ss.getSheetByName("OVERNIGHT_STAYS");

	if (!dashboardSheet || !peopleSheet || !transactionsSheet || !overnightSheet) {
		SpreadsheetApp.getActive().toast("Required sheet missing.", "❌ Error", 5);
		return;
	}

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
	// const tDateCol = tHeaders.indexOf("Date"); // Not used for now

	// Get overnight stays
	const osData = overnightSheet.getDataRange().getValues();
	const osHeaders = osData[0];
	const osPersonIdCol = osHeaders.indexOf("Person_ID");
	const osPersonCol = osHeaders.indexOf("Person");
	const osStartCol = osHeaders.indexOf("Start_Date");
	const osEndCol = osHeaders.indexOf("End_Date");
	const osCountCol = osHeaders.indexOf("Person_Count");

	// Prepare dashboard rows for the User_Dashboard table
	const dashboardRows = [
		["Person", "Account_Balance", "Deposits", "Withdrawals", "Charges", "Days", "Days%", "Stays", "Stays%"]
	];


	// First, collect all totals for all people
	const personStats = personList.map(person => {
		let deposits = 0;
		let withdrawals = 0;
		let charges = 0;
		let accountBalance = 0;
		let totalDays = 0;
		let totalStays = 0;

		// Transactions
		tData.slice(1).forEach(row => {
			if (row[tPersonCol] == person.name) {
				const amt = Number(row[tAmountCol]) || 0;
				const type = String(row[tTypeCol] || "").toLowerCase();
				if (type.includes("deposit")) {
					deposits += amt;
				} else if (type.includes("withdrawal")) {
					withdrawals += amt;
				} else if (type.includes("charge") || type.includes("reconciliation")) {
					charges += amt;
				}
				accountBalance += amt;
			}
		});

		// Overnight stays (Days)
		osData.slice(1).forEach(row => {
			if (row[osPersonIdCol] == person.code || row[osPersonCol] == person.name) {
				const start = row[osStartCol];
				const end = row[osEndCol];
				const count = Number(row[osCountCol]) || 1;
				if (start instanceof Date && end instanceof Date) {
					const days = Math.floor((end - start) / (1000 * 60 * 60 * 24)) + 1;
					if (days > 0) {
						totalDays += days;
						totalStays += days * count;
					}
				}
			}
		});

		return {
			name: person.name,
			accountBalance,
			deposits,
			withdrawals,
			charges,
			totalDays,
			totalStays
		};
	});

	// Calculate grand totals
	const grandTotalDays = personStats.reduce((sum, p) => sum + p.totalDays, 0);
	const grandTotalStays = personStats.reduce((sum, p) => sum + p.totalStays, 0);

	// Add rows with percentage calculations
	personStats.forEach(p => {
		const pctDays = grandTotalDays > 0 ? Math.round((p.totalDays / grandTotalDays) * 10000) / 100 : 0;
		const pctStays = grandTotalStays > 0 ? Math.round((p.totalStays / grandTotalStays) * 10000) / 100 : 0;
		dashboardRows.push([
			p.name,
			round2(p.accountBalance),
			round2(p.deposits),
			round2(p.withdrawals),
			round2(p.charges),
			p.totalDays,
			pctDays,
			p.totalStays,
			pctStays
		]);
	});

	// Find the User_Dashboard table range (assume it starts at A1)
	dashboardSheet.getRange(1, 1, dashboardRows.length, dashboardRows[0].length).setValues(dashboardRows);
}

