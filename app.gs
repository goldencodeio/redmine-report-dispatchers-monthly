function createMonthlyReportOld() {
  initMonthlyOptions();
  initPeriodTable('#fff2cc');
  processReports();
}

function createMonthlyReport() {
  initMonthlyOptions();
  processNewTable();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.addMenu('GoldenCode Report', [
    {name: 'Создать Ежемесячный Отчёт', functionName: 'createMonthlyReport'},
    {name: 'Создать Ежемесячный Отчёт (старая форма)', functionName: 'createMonthlyReportOld'},
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
