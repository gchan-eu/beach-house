// ---------- ID ASSIGNMENT HELPERS ----------

function assignId(sheet, row, col, triggerColumnName, idColumnName, startNumber) {
  const triggerColumn = getColumnIndexByHeader(sheet, triggerColumnName);
  const idColumn = getColumnIndexByHeader(sheet, idColumnName);
  if (row === 1 || col !== triggerColumn) return;
  const idCell = sheet.getRange(row, idColumn);
  if (idCell.getValue()) return;
  const idRange = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  const existingIds = idRange.flat().filter(n => typeof n === "number");
  const nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : startNumber;
  idCell.setValue(nextId);
}

function assignOwnershipSetId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  assignId(sheet, row, col, "Date", "ID", 100001);
}

function assignOwnerShareId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  assignId(sheet, row, col, "Ownership_Set_ID", "ID", 100001);
}

function assignOvernightStayId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  assignId(sheet, row, col, "Person", "ID", 100001);
}

function assignSplitMethodId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  assignId(sheet, row, col, "Type", "ID", 101);
}

function assignExpenseId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  assignId(sheet, row, col, "Date", "ID", 100001);
  // Also set Status to "Pending" if empty
  const statusColumn = getColumnIndexByHeader(sheet, "Status");
  const statusCell = sheet.getRange(row, statusColumn);
  if (!statusCell.getValue()) {
    statusCell.setValue("Pending");
  }
}

function assignTransactionId(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  assignId(sheet, row, col, "Date", "ID", 1);
}
