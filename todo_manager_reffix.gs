function refFix() {
  var startTime = new Date();
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
   
  var taskRange = sheet.getRange(1, 1, lastRow, 1);
  var frozenRange = sheet.getRange(1, 1, 1, lastColumn);
  for (var i = 1; i <= lastColumn; i++) {
    if (frozenRange.getCell(1, i).getValue() == "start") {
      var startRange = sheet.getRange(1, i, lastRow, 1);
    } else if (frozenRange.getCell(1, i).getValue() == "end") {
      var endRange = sheet.getRange(1, i, lastRow, 1);
    } else if (frozenRange.getCell(1, i).getValue() == "length") {
      var lengthRange = sheet.getRange(1, i, lastRow, 1);
    } else if (frozenRange.getCell(1, i).getValue() == "expected time") {
      var expectedTimeRange = sheet.getRange(1, i, lastRow, 1);
    } else if (frozenRange.getCell(1, i).getValue() == "discrepancy") {
      var discrepancyRange = sheet.getRange(1, i, lastRow, 1);
    }
  }
  
  var isNewDay = 1;
  for (var row = sheet.getFrozenRows() + 1; row <= lastRow; row++) {
    var taskCell = taskRange.getCell(row, 1);
    var lengthCell = lengthRange.getCell(row, 1);
    var startCell = startRange.getCell(row, 1);
    var endCell = endRange.getCell(row, 1);
    var expectedTimeCell = expectedTimeRange.getCell(row, 1);
    var discrepancyCell = discrepancyRange.getCell(row, 1);
    if (taskCell.getValue() == "") {
      isNewDay = 1;
    } else if (isNewDay) {
      if (startCell.getValue() == "") startCell.setValue("4:00");
      endCell.setFormula("=" + startCell.getA1Notation());
      isNewDay = 0;
    } else {
      startCell.setFormula(endCell.offset(-1, 0).getA1Notation());
      var t = "TIME(0," + lengthCell.getA1Notation() + ",0)+" + startCell.getA1Notation();
      endCell.setFormula("=TIME(HOUR(" + t + "),MINUTE(" + t + "),SECOND(" + t + "))");
      isNewDay = 0;
      
      // Does the actual time match the expected fixed time?
      if (expectedTimeCell.getValue() != "") {
        discrepancyCell.setFormula("=TO_TEXT(" + startCell.getA1Notation() + "-" + expectedTimeCell.getA1Notation() + ")");
        if (discrepancyCell.getValue() == "0:00:00") {
          startCell.setBackground("white");
        } else {
          startCell.setBackground("red");
        }
      } else {
        startCell.setBackground("white");
      }
    }
  }
}
