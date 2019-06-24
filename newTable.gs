var REPORT = [
  {
    code: 'boss_rating_avg',
    name: 'Средняя оценка\nза создание задач',
    manual: false,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Еженедельно',
  },
  {
    code: 'feedback_rating_avg',
    name: 'Средняя оценка за отзв-ся задачи',
    manual: false,
    coefficient: true,
    execPeriod: 'Ежедневно',
    controlPeriod: 'Еженедельно',
  },
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

var HEADERS = [
  {
    startRow: "B1",
    endRow: "C1",
    name: 'АВТОМАТИЧЕСКИЙ РАССЧЕТ',
    color: 'yellow',
  },
  {
    startRow: "D1",
    endRow: "K1",
    name: 'РУЧНОЙ РАССЧЕТ',
    color: '#9BBB59',
  },
  {
    startRow: "L1",
    endRow: "N1",
    name: 'Штрафы',
    color: 'red',
  },
];
  
function processNewTable() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = _ss.getActiveSheet();
  
  performers = OPTIONS.performers;
  performers.forEach(function(user, i) {
    var userData = APIRequest('users', {query: [{key: 'name', value: user}]}).users[0];
    OPTIONS.performers[i] = userData;
  });
  
  header(sheet, HEADERS, REPORT);
  // getOptionsData();
  generateSidePanel(sheet);
  
  // Высота второй строки
  sheet.setRowHeights(2, 1, 100);
  //Закрепленная боковая панель
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 200);
  
  generationData(sheet);

  var calcSheetNames = [];  
  performers.forEach(function(user) {
    sheetName = "Рассчет " + user.firstname + " " + OPTIONS.startDate.toLocaleDateString('ru-RU', {month: 'long', year: 'numeric'}).substr(2);
    calcSheetNames.push(sheetName);
  });
  generateManualData(sheet, calcSheetNames);
  
  calcSheetNames.forEach(function(name) {
    sheet = _ss.getSheetByName(name);
    if (sheet)
      sheet.activate();
    else 
      s = createNewSheet(name, "white");
    createCalculationSheet();
  });
}
  
function header(sheet, headers, reports) {
  headers.forEach(function(head) {
    range = sheet.getRange(head.startRow + ":" + head.endRow);
    range.setBackground(head.color);
    sheet.getRange(head.startRow).setValue(head.name);
  })
  
  var row = 2;
  var col = 2;
  
  reports.forEach(function(report) {
    var colName = sheet.getRange(row, col);
    colName.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    colName.setValue(report.name);
    colName.setHorizontalAlignment("center");
    colName.setVerticalAlignment("middle");
    colName.setBorder(true, true, true, true, null, null);
  
    var execPeriod = sheet.getRange(row + 1, col);
    execPeriod.setValue(report.execPeriod);
    execPeriod.setBackground(getColorByPeriod(report.execPeriod));
    execPeriod.setHorizontalAlignment("center");
    execPeriod.setBorder(true, true, true, true, null, null);
    
  
    var controlPeriod = sheet.getRange(row + 2, col);
    controlPeriod.setValue(report.controlPeriod);
    controlPeriod.setBackground(getColorByPeriod(report.controlPeriod));
    controlPeriod.setHorizontalAlignment("center");
    controlPeriod.setBorder(true, true, true, true, null, null);
  
    col++;
  });
}

function generateSidePanel(sheet) {
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
  
  performers = OPTIONS.performers;
  names = [];
  performers.forEach(function(userData, i) {
    name = userData.firstname + ' ' + userData.lastname + ' (' + userData.login + ')';
    names.push(name);
  });
  
  row = 6;
  names.forEach(function(user) {
    sheet.getRange(row, 1)
      .setValue(user)
      .setBackground("#FFF2CC")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null);
    row++;
  });
    
  row++;
  sheet.getRange(row, 1)
    .setValue('Коэффициент:')
    .setFontWeight('bold')
    .setBackground('#4DD0E1')
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);
  
  row += 2;
  sheet.getRange(row, 1)
    .setValue('Расчет:')
    .setBackground('#4DD0E1')
    .setFontWeight('bold')
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);
    
  row++;
  names.forEach(function(user) {
    sheet.getRange(row, 1)
      .setValue(user)
      .setBackground("#FFF2CC")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null);
    row++;
  });
    
  row++;
  sheet.getRange(row, 1)
    .setValue('Итоговая оценка:')
    .setBackground('#4DD0E1')
    .setFontWeight('bold')
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);
    
  row++;
  names.forEach(function(user) {
    sheet.getRange(row, 1)
      .setValue(user)
      .setBackground("#FFF2CC")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null);
    row++;
  });
  
  row++;
  sheet.getRange(row, 1)
    .setValue('К выплате:')
    .setBackground('#4DD0E1')
    .setFontWeight('bold')
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);
    
  row++;
  names.forEach(function(user) {
    sheet.getRange(row, 1)
      .setValue(user)
      .setBackground("#FFF2CC")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null);
    row++;
  });
}

function generationData(sheet) {
  row = 6;
  column = 2;
  performers = OPTIONS.performers;
  
  var taskCount = 0;
  var counts = [];
  performers.forEach(function(user, i) {
    count = getUserReport('feedback_tasks', user);
    counts.push(count.length ? count.length : 0);
    taskCount += counts[i];
  });
    
  REPORT.forEach(function(report, index) {
    row = 6;
    startChar = 'A';
    
    performers.forEach(function(user, i) {
      range = sheet.getRange(row, column);
    
      if (!report.manual) {
        if (report.code !== 'feedback_rating_avg') {
          reportValue = getUserReport(report.code, user, i);
        
          if (Array.isArray(reportValue)) {
            if (Array.isArray(reportValue[0])) {
              range.setValue(reportValue[0].length + ' / ' + reportValue[1].length);
            } else {
              range.setValue(reportValue.length)
            }
          } else {
            range.setValue(reportValue);
          }
        } else {
          //range.setValue((102 * 5) / (reportValue.length / performers.length));
          range.setValue((counts[i] * 5) / (taskCount / counts.length))
        }
      }
      
      if (i % 2 === 0)
        range.setBackground(sheet.getRange(1, column).getBackground());

      range.setBorder(true, true, true, true, null, null).setHorizontalAlignment("center");
      row++;
    });
    column++;
  });

  generateCoeffitions(sheet, row);
}

function generateCoeffitions(sheet, row) {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();
  row++;
  column = 2;
  
  REPORT.forEach(function (report) {
    if (report.coefficient) {
      range = sheet.getRange(row, column++)
      range
        // .setValue(report.coefficient)
        .setBorder(true, true, true, true, null, null)
        .setFontWeight('bold')
        .setHorizontalAlignment("center");
    }
  });
  
  range = sheet.getRange(row + ':' + row);
  protec = range.protect()
    .setDescription("Coefficient Protect");
  
  protec.getEditors().forEach(function(user) {
    protec.removeEditor(user);
  });
  _ss.getEditors().forEach(function(user) {
    if (user.getEmail()) {
      protec.addEditor(user.getEmail())
    }
  });
  
  sheet.hideRows(row);
  generateCalculations(sheet, row);
}

function generateCalculations(sheet, row) {
  coefRow = row;
  dataRow = 6;
  row += 3;
  performers = OPTIONS.performers;
  performers.forEach(function(user, i) {
    column = 2;
    letter = 'A';
    REPORT.forEach(function(report) {
      if (report.coefficient) {
        range = sheet.getRange(row, column++);
        letter = getNextChar(letter);
        range
          .setBorder(true, true, true, true, null, null)
          .setValue('=SUM(' + letter + dataRow + '*' + letter + coefRow + ')')
          .setBackground(i % 2 === 0 ? '#e0f7fa' : 'white');
      }
    });
    
    row++;
    dataRow++;
  });
  
  generateFinal(sheet, row);
}

function generateFinal(sheet, row) {
  performers = OPTIONS.performers;
  dataRow = 6;
  calcRow = row - performers.length;
  row += 2;
  performers.forEach(function(user) {
    column = 2;
    range = sheet.getRange(row, column++);
    range.setValue('=SUM(B' + calcRow + ':K' + calcRow + ')-SUM(L' + dataRow + ':N' + dataRow + ')')
      .setBorder(true, true, true, true, null, null)
      .setFontWeight('bold');
    range = sheet.getRange(row, column);
    range.setValue('=SUM(C' + calcRow + ':K' + calcRow + ')');
    
    dataRow++;
    calcRow++;
    row++;
  });
  
  generatePay(sheet, row);
}

function generatePay(sheet, row) {
  finalRow = row - performers.length;
  row += 2;
  performers = OPTIONS.performers;
  performers.forEach(function(user) {
    column = 2;
    range = sheet.getRange(row, column);
    range.setValue('=SUM((7500*B' + finalRow++ + ')/5)');
    row++;
  });
}

function generateManualData(sheet, calcSheetNames) {
  row = 6;
  
  calcSheetNames.forEach(function(name) {
    calcColumn = 'A';
    column = 2;
    REPORT.forEach(function(report) {
      if (!report.manual) {
        column++;
        return;
      }
      calcRow = OPTIONS.finalDate.getDate() + 5;
      if (report.controlPeriod === 'Ежемесячно') calcRow = 5;
      
      range = sheet.getRange(row, column);
      calcColumn = getNextChar(calcColumn);
      range.setValue('=\'' + name + '\'!' + calcColumn + calcRow);
      
      column++;
    });
    row++;
  });
  
}

function getNextChar(c) {
  return String.fromCharCode(c.charCodeAt() + 1);
}

function isFloat(n){
    return Number(n) === n && n % 1 !== 0;
}

function changeFloat(n) {
  return String.replace(n.toFixed(2), '.', ',');
}

function getColorByPeriod(period) { 
  switch (period) {
    case 'Ежемесячно':
      return '#D9EAD3';
    case 'Ежедневно':
      return '#F2DCDB';
    case 'Еженедельно':
      return '#FFE599';
    default:
      return 'white';
  }
}
