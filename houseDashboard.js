/**
 * Updates the HOUSE_DASHBOARD sheet's House_Dashboard table with dynamic columns for each person and rows for each expense type.
 * Columns: Category, Type, Total, [person codes...]
 * Each row: one for each expense type (from EXPENSE_TYPES sheet)
 * Total: sum of charges and reconciliation for that type (from TRANSACTIONS)
 * Per-person columns: sum for that type and person (by code)
 */
function updateHouseDashboard() {
	try {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dashboardSheet = ss.getSheetByName("HOUSE_DASHBOARD");
	const peopleSheet = ss.getSheetByName("PEOPLE");
	const transactionsSheet = ss.getSheetByName("TRANSACTIONS");
	const expenseTypesSheet = ss.getSheetByName("EXPENSE_TYPES");

		if (!dashboardSheet || !peopleSheet || !transactionsSheet || !expenseTypesSheet) {
			SpreadsheetApp.getActive().toast("Required sheet missing.", "❌ Error", 5);
			Logger.log('Missing sheet: dashboardSheet=%s, peopleSheet=%s, transactionsSheet=%s, expenseTypesSheet=%s', !!dashboardSheet, !!peopleSheet, !!transactionsSheet, !!expenseTypesSheet);
			return;
		}

	// Get people codes
		const peopleData = peopleSheet.getDataRange().getValues();
		const peopleHeaders = peopleData[0];
		const codeCol = peopleHeaders.indexOf("Code");
		if (codeCol === -1) {
			Logger.log('Code column not found in PEOPLE sheet headers: %s', JSON.stringify(peopleHeaders));
			SpreadsheetApp.getActive().toast('Code column not found in PEOPLE sheet', '❌ Error', 5);
			return;
		}
		const personCodes = peopleData.slice(1).map(row => row[codeCol]).filter(Boolean);
		Logger.log('Person codes: %s', JSON.stringify(personCodes));

	// Get expense types
		const expTypeData = expenseTypesSheet.getDataRange().getValues();
		const expTypeHeaders = expTypeData[0];
		const catCol = expTypeHeaders.indexOf("Category");
		const descCol = expTypeHeaders.indexOf("Description");
		const expTypeCodeCol = expTypeHeaders.indexOf("Code");
		if (catCol === -1 || descCol === -1 || expTypeCodeCol === -1) {
			Logger.log('Missing columns in EXPENSE_TYPES headers: %s', JSON.stringify(expTypeHeaders));
			SpreadsheetApp.getActive().toast('Missing columns in EXPENSE_TYPES sheet', '❌ Error', 5);
			return;
		}
		const expTypes = expTypeData.slice(1).map(row => ({
			category: row[catCol],
			type: row[descCol],
			code: row[expTypeCodeCol]
		})).filter(row => row.category && row.type && row.code);
		Logger.log('Expense types: %s', JSON.stringify(expTypes));

	// Get transactions
		const tData = transactionsSheet.getDataRange().getValues();
		const tHeaders = tData[0];
		const tTypeCol = tHeaders.indexOf("Type");
		const tExpTypeCol = tHeaders.indexOf("Expense_Type");
		const tAmountCol = tHeaders.indexOf("Amount");
		const tPersonCol = tHeaders.indexOf("Person");
		if (tTypeCol === -1 || tExpTypeCol === -1 || tAmountCol === -1 || tPersonCol === -1) {
			Logger.log('Missing columns in TRANSACTIONS headers: %s', JSON.stringify(tHeaders));
			SpreadsheetApp.getActive().toast('Missing columns in TRANSACTIONS sheet', '❌ Error', 5);
			return;
		}

	// Build header row
		const headerRow = ["Category", "Type", "Total", ...personCodes];
		const dashboardRows = [headerRow];

	// For each expense type, build a row
		expTypes.forEach(expType => {
			// Filter transactions for this type (charges or reconciliation)
			const relevantTxns = tData.slice(1).filter(row => {
				const tType = String(row[tTypeCol] || "").toLowerCase();
				const tExpType = String(row[tExpTypeCol] || "");
				return (tType.includes("charge") || tType.includes("reconciliation")) && tExpType.startsWith(String(expType.code));
			});
			Logger.log('Expense type %s (%s): found %s relevant transactions', expType.type, expType.code, relevantTxns.length);

			// Total for this type
			const total = relevantTxns.reduce((sum, row) => sum + (Number(row[tAmountCol]) || 0), 0);

			// Per-person columns
			const perPerson = personCodes.map(code => {
				// Person column is like "AH - Soula Chantzopoulos", so match prefix
				const personTxns = relevantTxns.filter(row => String(row[tPersonCol] || "").trim().startsWith(code));
				Logger.log('Expense type %s, person %s: %s transactions', expType.code, code, personTxns.length);
				return personTxns.reduce((sum, row) => sum + (Number(row[tAmountCol]) || 0), 0);
			});

			dashboardRows.push([
				expType.category,
				expType.type,
				total,
				...perPerson
			]);
		});

		// Add footer row with totals
		if (dashboardRows.length > 1) {
			const numCols = dashboardRows[0].length;
			// Sum each column (skip first two columns: Category, Type)
			const totals = ["", "Total:"];
			for (let col = 2; col < numCols; col++) {
				let sum = 0;
				for (let row = 1; row < dashboardRows.length; row++) {
					sum += Number(dashboardRows[row][col]) || 0;
				}
				totals.push(sum);
			}
			dashboardRows.push(totals);
		}

		// Write to dashboard
		dashboardSheet.clearContents();
		dashboardSheet.getRange(1, 1, dashboardRows.length, dashboardRows[0].length).setValues(dashboardRows);
	} catch (err) {
		Logger.log('Error in updateHomeDashboard: %s', err && err.stack ? err.stack : err);
		SpreadsheetApp.getActive().toast('Error in updateHomeDashboard: ' + err, '❌ Error', 5);
	}
}
