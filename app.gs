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
