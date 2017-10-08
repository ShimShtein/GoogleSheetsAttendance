 // The onOpen function is executed automatically every time a Spreadsheet is loaded
 function onOpen() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var menuEntries = [];
   menuEntries.push({name: "Fill attendance", functionName: "initForm"});

   ss.addMenu("Custom actions", menuEntries);
 }

function initForm() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  
  var namesIndex = indexSheets(sheets);
  setupSidebar(namesIndex);
  var i = 0;
}

function indexSheets(sheets) {
  var names = {};
  sheets.forEach (function(sheet) {
    var range = sheet.getDataRange().getValues();
    var col_offset = 0;
    if (isColEmpty(range, 0)) {
      col_offset = 1;
    }
    
    names_row = range[1];
    for(var i=0; i<(names_row.length+1-col_offset) / 2; i++) {
      var name = names_row[i*2+col_offset];
      if (name === undefined) { continue; }
      names[name] = { name: name, sheet: sheet, name_col: i*2+col_offset }
    }
  });
  return names;
}

function isColEmpty(range, col) {
  col = range.map(function(row) {
    return row[col];
  });
  var nonEmpty = col.reduce(function(acc, cell) {
    return acc || !(typeof(cell) === 'string' && cell.length === 0);
  }, false);
  return !nonEmpty;
}

function setupSidebar(names) {
  var t = HtmlService.createTemplateFromFile('AttendanceSidebar');
  t.namesIndex = names;
  html = t.evaluate()
          .setTitle('Attendance sidebar');
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function formatHtmlDate(date){
  return "" + date.getFullYear() + "-" + (date.getMonth()+1) + "-" +  Utilities.formatString('%02d', date.getDate());
}
