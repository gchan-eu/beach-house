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

function handleUserDashboard(e) {
	updateUserDashboard();
}

function handleHouseDashboard(e) {
	updateHouseDashboard();
}

function handleOwnerDashboard(e) {
	updateOwnerDashboard();
}