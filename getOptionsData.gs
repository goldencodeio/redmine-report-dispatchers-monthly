var OPTIONS = {};

function getOptionsData(reportType) {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  var optionsSheet = _ss.setActiveSheet(getOptionsSheet());

  var data = optionsSheet.getRange(1, 1, optionsSheet.getLastRow(), optionsSheet.getLastColumn()).getValues();
  data.forEach(function(row) {
    var key = row.shift();
    row = row.filter(function(a) {
      if (a === 0) return true;
      return a
    });
    OPTIONS[key] = row.length > 1 ? row : row[0];
  });

  OPTIONS.startDate.setHours(OPTIONS.startDate.getHours() - 1 * OPTIONS.startDate.getTimezoneOffset() / 60);
  OPTIONS.finalDate.setHours(OPTIONS.finalDate.getHours() - 1 * OPTIONS.finalDate.getTimezoneOffset() / 60);

  if (!Array.isArray(OPTIONS.performers)) OPTIONS.performers = [OPTIONS.performers];
  if (!Array.isArray(OPTIONS.performersBonus)) OPTIONS.performersBonus = [OPTIONS.performersBonus];
}

function getOptionsSheet() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === 'options')
      return sheets[i];
  }
  return null;
}
