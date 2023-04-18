function processData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var teamSheet = ss.getSheetByName('TeamSheet');
  var sheets = ss.getSheets();
  teamSheet.getRange('A4:NB98').clearContent();
  teamSheet.getRange('A4:NB98').setFontColor('#000000');
  teamSheet.getRange('A4:NB98').setBackground('#FFFFFF');
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName() != 'TeamSheet') {
      var sheetName = sheet.getName();
      teamSheet.getRange(i + 4, 1).setValue(sheet.getRange('C1').getValue());
      var dates = sheet
        .getRange('A11:A375')
        .getValues()
        .flat()
        .map(function (date) {
          return date.toString();
        });
      var leaveTypes = sheet.getRange('D11:D375').getValues().flat();
      for (var j = 0; j < dates.length; j++) {
        var date = dates[j];
        var leaveType = leaveTypes[j];
        var employeeName = sheet.getRange('C1').getValue();
        if (leaveType == '' || leaveType == 'Weekend' || leaveType == 'Public Holiday') {
          continue;
        }
        var row = findEmployeeRow(employeeName, teamSheet);
        var col = findDateColumn(date, teamSheet);
        teamSheet.getRange(row, col).setValue(leaveType);
        teamSheet.getRange(row, col).setBackground(getColour(leaveType));
        if (getColour(leaveType) == '#0a52a8' || getColour(leaveType) == '#ff0100') {
          teamSheet.getRange(row, col).setFontColor('#FFFFFF');
        }
        console.log(`${sheetName} | Day: ${j + 1}`);
      }
    }
  }
}

function findEmployeeRow(employeeName, teamSheet) {
  var employees = teamSheet.getRange('A4:A98').getValues();
  for (var i = 0; i < employees.length; i++) {
    if (employees[i][0] == employeeName) {
      return i + 4;
    }
  }
}

function findDateColumn(date, teamSheet) {
  var dates = teamSheet.getRange('B1:NB1').getValues()[0];
  for (var i = 0; i < dates.length; i++) {
    if (dates[i] == date) {
      return i + 2;
    }
  }
}

function getColour(leaveType) {
  switch (leaveType) {
    case 'Vacation':
      return '#0a52a8';
    case 'Sick':
      return '#ffc7ce';
    case 'Casual Day':
      return '#ff0100';
    case 'Public Holiday':
      return '#92d050';
    case 'WFH':
      return '#ffe5a0';
    case 'Weekend':
      return '#ffff01';
    default:
      return '#FFFFFF';
  }
}
