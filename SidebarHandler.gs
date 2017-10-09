function debugAttendance() {
  response =  {
    date: "2017-10-08", 
    attendants:{
      "אלונה לוין":{
        nameCol:6, 
        sheet:"a-b"}, 
      "רוזה לנקין":{
        nameCol:1, 
        sheet:"c-d"}, 
      "קטי":{
        nameCol:5, 
        sheet:"c-d"}, 
      "רעיה זוזקין":{
        nameCol:3, 
        sheet:"c-d"}}};
  processAttendance(response);
}

function processAttendance(response){
  Logger.log(response);
  
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  
  var sheetsIndex = getSheetsForUpdate(sheets);
  updateIndex(response, sheetsIndex);
  commitIndex(sheetsIndex);
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
  
  for (var person in response.attendants) {
    Logger.log("Updating person: " + person);
    updatePerson(date, index, response.attendants[person]);
  };
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