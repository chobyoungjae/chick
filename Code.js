function mergeSameValuesInColumnA() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, 1, lastRow, 1); // A열 전체
  var values = range.getValues();

  var startRow = 1;

  for (var i = 2; i <= lastRow + 1; i++) {
    if (i > lastRow || values[i - 1][0] !== values[i - 2][0]) {
      if (i - startRow > 1 && values[startRow - 1][0] !== "") {
        sheet.getRange(startRow, 1, i - startRow, 1).merge();
      }
      startRow = i;
    }
  }
}
