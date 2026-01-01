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

