var OPTIONS = {};

function initMonthlyOptions() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  getOptionsData();

  var sheetName = 'Итог ' + OPTIONS.startDate.toLocaleDateString('ru-RU', {month: 'long', year: 'numeric'}).substr(2);
  var existingSheet = _ss.getSheetByName(sheetName);
  if (existingSheet)
    existingSheet.activate();
  else
    createNewSheet(sheetName, '#ffd966');
}

function createNewSheet(name, color) {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ss.insertSheet(name).setTabColor(color).setColumnWidth(1, 200).activate().setFrozenColumns(1);
}
