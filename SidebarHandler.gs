function debugAttendance() {
  var updateReport = { date: new Date("2019-10-25"), attendants: {
    "אנה גוסר": { row: 9, dateColumn: 9, sheetName: "a-b" }
} };
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  
  var sheetsIndex = getSheetsForUpdate(sheets);
  drawReport(updateReport, sheetsIndex);
}

function debugFocus() {
  focusCell("a-b", 10, 10);
}

function processAttendance(response){
  Logger.log(response);
  
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  
  var sheetsIndex = getSheetsForUpdate(sheets);
  var updateReport = updateIndex(response, sheetsIndex);
  commitIndex(sheetsIndex);
  drawReport(updateReport, sheetsIndex);
}

function getSheetsForUpdate(sheets) {
  var index = {};
  sheets.forEach (function(sheet) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    var range = sheet.getRange(1, 1, lastRow + 1, lastCol);
    var name = sheet.getName();
    var values = range.getValues();
    
    index[name] = { sheet: sheet, range: range, values: values, dirty: false }
  });
  return index;
}

function updateIndex(response, index) {
  var date = new Date(response.date);
  var updateReport = { date: date, attendants: {} }
  
  for (var person in response.attendants) {
    Logger.log("Updating person: " + person);
    updateReport.attendants[person] = updatePerson(date, index, response.attendants[person]);
  };
  return updateReport;
}

function commitIndex(index) {
  for (var key in index) {
    var entry = index[key];
    if (!entry.dirty) { continue; }
    entry.range.setValues(entry.values);    
  }
}

function updatePerson(date, index, person) {
  sheetName = person.sheet;
  var column = 1 + parseInt(person.nameCol);
  Logger.log("Updating column: " + column);
  entry = index[sheetName];
  values = entry.values;
  var row = lastRow(column, values);
  Logger.log("Last row: " + row + " column: " + column);
  values[row][column] = formatOutputDate(date);
  entry.dirty = true;
  Logger.log("Updated row: " + row + " column: " + column + " with: " + values[row][column]);
  return { row: row, dateColumn: column, sheetName: sheetName };
}

function lastRow(column, values) {
  Logger.log("Counting lastRow for column: " + column);
  var i=0;
  for (i=values.length-1; i>1; i--) {
    if (values[i][column] != undefined && values[i][column] != "") { i++; break; }
  }
  
  Logger.log("lastRow returned i: " + i);
  return i;
}

function formatOutputDate(date) {
  return "" + date.getDate() + "." + (date.getMonth()+1) + "." + date.getFullYear();
}

function drawReport(updateReport, sheetIndex) {
  var updateReportView = createReportViewData(updateReport, sheetIndex);
  
  var t = HtmlService.createTemplateFromFile('AttendanceReport');
  t.updateReportView = updateReportView;
  html = t.evaluate()
          .setTitle('Attendance report');
  SpreadsheetApp.getUi()
      .showSidebar(html);

}

function createReportViewData(updateReport, sheetIndex){
  var reportView = { date: updateReport.date, attendants: {} }
  for (var person in updateReport.attendants) {
    Logger.log("Preparing to view person: " + person);
    var personReport = updateReport.attendants[person]
    var punchCol = personReport.dateColumn - 1;
    var punch = sheetIndex[personReport.sheetName].values[personReport.row][punchCol];
    reportView.attendants[person] = { sheet: personReport.sheetName, punchCol: punchCol, punchRow: personReport.row, punch: punch } ;
  };
  
  return reportView;
}

function focusCell(sheet, row, column) {
  var ss = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheet);
  sheet.setActiveRange(sheet.getRange(new Number(row) + 1, new Number(column) + 1));
}