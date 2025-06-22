function clearSelected() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.getRange(GROUNDS_RANGE).clear()
  clearCheckbox();
}
