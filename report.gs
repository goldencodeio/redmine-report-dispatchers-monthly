var REPORT = [
  {
    code: 'created_tasks',
    name: 'Задач заведенных\nоператором',
    manual: false
  },
  {
    code: 'feedback_tasks',
    name: 'Кол-во отзвонившихся\nзадач',
    manual: false
  },
  {
    code: 'boss_rating_avg',
    name: 'Средняя оценка\nза создание задач',
    manual: false
  },
  {
    code: 'feedback_rating_avg',
    name: 'Средняя оценка\nза отзв-ся задачи',
    manual: true
  },
  {
    code: 'work_rating_avg',
    name: 'Средняя оценка\nза работу оператора',
    manual: true
  },
  {
    code: 'involvement_rating_avg',
    name: 'Средняя оценка\nза вовлеченность\nв работу',
    manual: true
  },
  {
   code: 'penalty_feedback',
   name: 'Штраф за совершенную\nобр. связь',
   manual: true
  },
  {
   code: 'penalty_delays',
   name: 'Штраф за\nопоздания',
   manual: true
  },
  {
   code: 'penalty_algorithm_work',
   name: 'Штраф за\nнесоблюдение\nАлгоритма работы\nнад задачами',
   manual: true
  },
  {
   code: 'penalty_daily_break',
   name: 'Штраф за\nпревышение\nежедневного\nперерыва',
   manual: true
  },
  {
    code: 'total_rating',
    name: 'ВСЕГО',
    manual: true
  },
  {
    code: 'finally_bonus',
    name: 'КПЭ',
    manual: true
  }
];

function processReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowI = 2;
  var columnI = 2;

  // fix bug for increase quantity columns
  for (var i = 2; i <= 33; i++) {
    sheet.getRange(rowI, i).setBackground('#fff');
  }
  // end fix

  OPTIONS.performers.forEach(function(user, userIndex) {
    REPORT.forEach(function(report) {
      if (!report.manual) {
        var reportValue = getUserReport(report.code, user, userIndex);
        if ((Array.isArray(reportValue))) {
          var listUrl = '';
          if ((Array.isArray(reportValue[0]))) {
            reportValue[0].forEach(function(task) {
              listUrl += 'http://redmine.zolotoykod.ru/issues/' + task.id + '\n';
            });
            sheet.getRange(rowI, columnI++).setValue(reportValue[0].length + ' / '+ reportValue[1].length).setNote(listUrl);
          } else {
            reportValue.forEach(function(task) {
              listUrl += 'http://redmine.zolotoykod.ru/issues/' + task.id + '\n';
            });
            sheet.getRange(rowI, columnI++).setValue(reportValue.length).setNote(listUrl);
          }
        } else {
          sheet.getRange(rowI, columnI++).setValue(reportValue);
        }
      } else {
        switch (report.code) {
          case 'feedback_rating_avg':
            sheet.getRange(rowI, columnI).setFormula('=SUM((C' + rowI + '*5)/(B' + (OPTIONS.performers.length + 3) + '/' + OPTIONS.performers.length + '))');
            break;

          case 'work_rating_avg':
            sheet.getRange(rowI, columnI).setFormula('=SUM((D' + rowI + '+E' + rowI + ')/2)');
            break;

          case 'involvement_rating_avg':
            sheet.getRange(rowI, columnI).setFormula('=SUM(AG' + (rowI + 6) + ')');
            break;

          case 'total_rating':
            sheet.getRange(rowI, columnI).setFormula('=SUM((((0,65*F' + rowI + ')/5)+((0,35*G' + rowI + ')/5))*5)-SUM(H' + rowI + ':K' + rowI + ')');
            break;

          case 'finally_bonus':
            sheet.getRange(rowI, columnI).setFormula('=SUM((L' + rowI + '*' + OPTIONS.performersBonus[userIndex] +')/5)');
            break;
        }

        columnI++;
      }
    });

    columnI = 2;
    rowI++;
  });

  rowI++;
  sheet.getRange(rowI, columnI).setFormula('=SUM(C2:C' + (rowI - 2) + ')');

  rowI += 2;
  var rangeDaysA1 = sheet.getRange(rowI, columnI, 1, 31).getA1Notation();
  for (var i = 1; i <= 31; i++) {
    sheet.getRange(rowI, columnI++).setValue(i).setBackground('#aaa');
  }
  var countdaysA1 = sheet.getRange(rowI++, columnI).setFormula('=COUNTA(' + rangeDaysA1 + ')').setBackground('#ddd').getA1Notation();
  columnI = 2;

  OPTIONS.performers.forEach(function(user) {
    var rangeA1 = sheet.getRange(rowI, columnI, 1, 31).getA1Notation();
    columnI = 33;
    sheet.getRange(rowI++, columnI).setBackground('#aaa').setFormula('=SUM((5*COUNTA(' + rangeA1 + '))/' + countdaysA1 + ')');
    columnI = 2;
  });
}
