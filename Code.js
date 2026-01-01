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
