function onOpen(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "Assign Formulas", functionName: "assignFormulas"});
  menuEntries.push({name: "Refresh Data", functionName: "refreshData"});

  ss.addMenu("Scripts", menuEntries);
}

function assignFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; 

  var totalsRange = sheet.getRange(sheet.getLastRow() - 3, 3, 1, sheet.getLastColumn() - 2);
  updateFormula(totalsRange, "=SUMNUTRIENTS()");

  var servingRange = sheet.getRange(sheet.getLastRow(), 3, 1, sheet.getLastColumn() - 2);
  updateFormula(servingRange, "=INDIRECT(ADDRESS(ROW() -3, COLUMN())) / INDIRECT(ADDRESS(ROW() - 2, 2))");
}

function updateFormula(range, formula) {
  var formulas = range.getFormulas();

  for (var row in formulas) {
    for (var col in formulas[row]) {
      formulas[row][col] = formula;
    }
  }

  range.setFormulas(formulas);
  SpreadsheetApp.flush();
}

function refreshData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; 

  resetFormulas(sheet.getRange(sheet.getLastRow() - 3, 3, 1, sheet.getLastColumn()));
  // resetFormulas(sheet.getRange(sheet.getLastRow(), 3, 1, sheet.getLastColumn()));
}

function resetFormulas(range) {
  var formulas = range.getFormulas();

  range.clearContent();
  SpreadsheetApp.flush();
  range.setFormulas(formulas);
  SpreadsheetApp.flush();
}

function SUMNUTRIENTS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var cell = sheet.getCurrentCell();
  var colRange = sheet.getRange(2, cell.getColumn(), cell.getRowIndex() - 2);
  var data = colRange.getValues();
  var total = 0;

  for (var row in data) {
    // for (var col in data[row]) {
      // var col = 0;
      // Logger.log("row: %s, col: %s", row, col)
      var val = parseFloat(data[row][0]);

      if (isNaN(val))
        continue;

      // Logger.log("row: " + row + " col: " + col + " data: " + val);

      var scaleRange = sheet.getRange(parseInt(row) + 2, 2);
      // Logger.log("Scale row: %s", scaleRange.getRowIndex());
      var scale = parseFloat(scaleRange.getValue());

      // Logger.log("scale: " + scale);

      if (isNaN(scale))
        scale = 1;

      total += scale * val;
    // }
  }

  // Logger.log("total: " + total);

  return total;
}
