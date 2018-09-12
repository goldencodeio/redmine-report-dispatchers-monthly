function createMonthlyReport() {
  initMonthlyOptions();
  initPeriodTable('#fff2cc');
  processReports();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.addMenu('GoldenCode Report', [
    {name: 'Создать Ежемесячный Отчёт', functionName: 'createMonthlyReport'}
  ]);
}

function createTrigger() {
  ScriptApp.newTrigger('createMonthlyReport')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
}

function deleteAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}
