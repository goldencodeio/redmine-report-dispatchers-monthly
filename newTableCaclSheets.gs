function testcalc() {
  getOptionsData();
  var _ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetName = 'kek';
  var existingSheet = _ss.getSheetByName(sheetName);
  if (existingSheet)
    existingSheet.activate();
  else
    createNewSheet(sheetName, '#ffd966');
  createCalculationSheet();
}

function createCalculationSheet() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = _ss.getActiveSheet();
  
  var headers = [
  {
    startRow: "B1",
    endRow: "I1",
    name: 'РУЧНОЙ РАССЧЕТ',
    color: '#9BBB59',
  },
  {
    startRow: "J1",
    endRow: "L1",
    name: 'Штрафы',
    color: 'red',
  },
  ];
    
  var report = [
  {
    code: 'manual',
    name: 'Прием/передача смены дежурному',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежедневно',
  },
  {
    code: 'manual',
    name: 'Вовлеченность в работу',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежедневно',
  },
  {
    code: 'manual',
    name: '*Контроль соблюдения сроков выполнения задач',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежедневно',
  },
  {
    code: 'manual',
    name: '*Своевременное получение ОС по выполненным задачам',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежедневно',
  },
  {
    code: 'manual',
    name: '*Предоставление информации клиенту по текущим задачам, услугам, тарифам',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежедневно',
  },
  {
    code: 'manual',
    name: '*Прием, передача информации в отдел продаж о выявлен-х потребностях',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Еженедельно',
  },
  {
    code: 'manual',
    name: '*Выполнение задач, установленных непосредственным или высшим руководством ',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Еженедельно',
  },
  {
    code: 'manual',
    name: '*Создание не менее 3х вики статей',
    manual: true,
    coefficient: true,
    execPeriod: 'Ежемесячно',
    controlPeriod: 'Ежемесячно',
  },
    {
    code: 'fine',
    name: 'Соблюдение трудового распорядка',
    manual: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежедневно',
  },
  {
    code: 'fine',
    name: 'Соблюдение ежедневного перерыва',
    manual: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежемесячно',
  },
  {
    code: 'fine',
    name: 'Претензии',
    manual: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Ежемесячно',
  },
  ];
  
  days = getCurrentMonthDayCount(OPTIONS.startDate.getMonth() + 1, OPTIONS.startDate.getFullYear());
  header(sheet, headers, report);
  generateSideMonthlyPanel(sheet, days);
  weekend(sheet,days,report);
  merge(sheet, days, report);
  setBorders(days, report.length + 1);
  result(sheet, 5 + days, report);
}

function getCurrentMonthDayCount(x, y) {
  return 28 + ((x + Math.floor(x / 8)) % 2) + 2 % x + Math.floor((1 + (1 - (y % 4 + 2) % (y % 4 + 1)) * ((y % 100 + 2) % (y % 100 + 1)) + (1 - (y % 400 + 2) % (y % 400 + 1))) / x) + Math.floor(1/x) - Math.floor(((1 - (y % 4 + 2) % (y % 4 + 1)) * ((y % 100 + 2) % (y % 100 + 1)) + (1 - (y % 400 + 2) % (y % 400 + 1)))/x);
}

function generateSideMonthlyPanel(sheet) {
  sheet.getRange("A3")
    .setValue('Периодичность выполнения')
    .setBackground('#4DD0E1')
    .setFontWeight('bold')
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);
  
  sheet.getRange("A4")
    .setValue('Периодичность оценки')
    .setBackground('#4DD0E1')
    .setFontWeight('bold')
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);
  
  startDate = OPTIONS.startDate;
  row = 5;
  column = 1;
  for (i = 1; i <= days; i++) {
    range = sheet.getRange(row, column);
    date = new Date(startDate.getFullYear(), startDate.getMonth(), i);
    range.setValue(date);
    row++;
  }
}

function merge(sheet, days, reports) {
  column = 2;
  reports.forEach(function(report) {
    row = 5;
    if (report.controlPeriod === 'Ежемесячно') {
      sheet.getRange(row, column, days).merge().setBackground('white');
    } else if (report.controlPeriod === 'Еженедельно') {
      date = OPTIONS.startDate;
      i = 1;
      if (date.getDay() !== 1) {
        week = 6 - date.getDay();
        if (week > 0 && week < 5) {
          sheet.getRange(row, column, week).merge().setBackground('white');
          i += week - 1;
        }
      }
      for (; i <= days; i++) {
        idate = new Date(OPTIONS.startDate.getFullYear(), OPTIONS.startDate.getMonth(), i);
        if (idate.getDay() !== 1) continue;
        if (days - i < 5) {
          var ost = days - i + 1;
        }
        sheet.getRange(i + row - 1 , column, ost ? ost : 5).merge().setBackground('white');
        i += 2;
      }
    }
    
    column++;
  });
}

function weekend(sheet, days, reports) {
  column = 1;
  row = 5;
  date = OPTIONS.startDate;
  
  var i = 1;
  
  if (date.getDay() !== 6 && date.getDay() !== 0) {
    week = 6 - date.getDay();
    i += week - 1;
  } else if (date.getDay() === 0) {
    sheet.getRange(row, column, 1, reports.length + 1).setBackground('#cccccc');
    i += 5;
  } else if (date.getDay() === 6) {
    sheet.getRange(row, column, 2, reports.length + 1).setBackground('#cccccc');
    i += 6;
  }

  for (; i <= days; i+=7) {
    sheet.getRange(i + row, column, 2, reports.length + 1).setBackground('#cccccc');
  }
}

function setBorders(row, column) {
  initRow = 5;
  initColumn = 1;
  sheet.getRange(initRow, initColumn, row, column).setBorder(true, true, true, true, true, true);
 }

function result(sheet, row, reports) {
  var char = 'A';
  column = 2;
  reports.forEach(function(report) {
    char = getNextChar(char);
    if (report.controlPeriod !== 'Ежемесячно') {
      range = sheet.getRange(row, column);
      range.setValue("=IFERROR(AVERAGE(" + char + "5" + ":" + char + (4+days) + "))");
    }
    column++;
  });
}