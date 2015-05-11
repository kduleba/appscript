var whoAmI = "kduleba@google.com";
var categoryDimensions = 3;

function onOpen() {  
  var menu = [ 
    {name: "Estimate Params", functionName: "estimateParams"},
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Run scripts", menu);
}

function AnalyzeEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mySheet = ss.getSheets()[0];  

  var currentDate = new Date();

  var starredEmails = GmailApp.search("is:starred in:inbox").length;
  var importantEmails = GmailApp.search("in:inbox is:important is:unread !is:starred").length;
  var unimportantEmails = GmailApp.search("in:inbox !is:important is:unread !is:starred").length;

  var firstRow = mySheet.getFrozenRows() + 1;
  mySheet.insertRows(firstRow);
  var cells = mySheet.getRange(firstRow, 1, 1, categoryDimensions + 1);
  cells.setValues([[currentDate, starredEmails, importantEmails, unimportantEmails]]);

  // TODO: this is not robust, it uses fixed frozen row value
  if (firstRow == 3) {
    mySheet.getRange(firstRow, categoryDimensions + 2, 1, 1).setFormula("=$B$2*B3+$C$2*C3+$D$2*D3");
  }
}

function Round(x) {
  return Math.round(x * 20) / 20.0
}

function evaluateParamError(dataColumns, workColumn, firstRow, lastRow, parameters, hourly) {
  var error = 0.0;
  var myEstimate = 0.0;
  for (var j = 1; j <= categoryDimensions; j++) {
    myEstimate += Number(dataColumns.getCell(lastRow, j).getValue()) * parameters[j];
  }

  for (var i = lastRow; i > firstRow; i--) {
    var nextEstimate = 0.0;
    for (var j = 1; j <= categoryDimensions; j++) {
      var nextVal = Number(dataColumns.getCell(i - 1, j).getValue());
      var currVal = Number(dataColumns.getCell(i, j).getValue());
      nextEstimate += nextVal * parameters[j];
      if (nextVal > currVal) {
        myEstimate += parameters[j] * (nextVal - currVal);
      }
    }

    var rowError = myEstimate + hourly - Number(workColumn.getCell(i, 1).getValue()) - nextEstimate;
    error += rowError * rowError;
    myEstimate = nextEstimate;
  }
  return error;
}

function displayParams(p) {
  return "" + Round(p[1]) + ", " + Round(p[2]) + ", " + Round(p[3]);
}

function readParameters(dataColumns) {
  var parameters = [0.0, 0.0, 0.0, 0.0];
  for (var i = 1; i <= categoryDimensions; i++) {
    parameters[i] = Number(dataColumns.getCell(2, i).getValue());
  }
  return parameters;
}

function storeParameters(p, dataColumns, hcell, h) {
  for (var i = 1; i <= categoryDimensions; i++) {
    dataColumns.getCell(2, i).setValue(p[i]);
  }
  hcell.setValue(h);
}

function estimateParams() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("log");
  var firstRow = sheet.getFrozenRows() + 1;
  var lastRow = sheet.getLastRow();
  var lastColumn = ss.getLastColumn();
  var email_lines = [];

  var frozenRange = sheet.getRange(1, 1, 1, lastColumn);
  for (var i = 1; i <= lastColumn; i++) {
    var colName = frozenRange.getCell(1, i).getValue();
    if (colName == "estimated time to clean up") {
      var cleanupTimeColumn = sheet.getRange(1, i, lastRow, 1);
    } else if (colName == "email time spent (manual entry)") {
      var workColumn = sheet.getRange(1, i, lastRow, 1);
    } else if (colName == "starred") {
      var dataColumns = sheet.getRange(1, i, lastRow, categoryDimensions);
    } else if (colName == "hourly_extra") {
      var hourlyCell = sheet.getRange(2, 1, 1, lastColumn).getCell(1, i);
    }
  }

  var elapsedTime = 0;

  var totalWork = Number(cleanupTimeColumn.getCell(firstRow, 1).getValue() - cleanupTimeColumn.getCell(lastRow, 1).getValue());
  for (var i = firstRow; i <= lastRow; i++) {
    totalWork += Number(workColumn.getCell(i, 1).getValue());  
  }
  var hourlyWorkAverage = Round(totalWork / (lastRow + 1 - firstRow));

  email_lines.push("hourly_extra: " + hourlyWorkAverage);
  email_lines.push("hours: " + (lastRow + 1 - firstRow));

  var parameters = readParameters(dataColumns);
  email_lines.push("params: " + displayParams(parameters));

  var error = evaluateParamError(dataColumns, workColumn, firstRow, lastRow, parameters, hourlyWorkAverage);

  for (var rep = 0; rep < 10; rep++) {
    var progress = false;
    for (var idx = 1; idx <= categoryDimensions; idx++) {
      var p = parameters[idx];

      parameters[idx] = p + 0.05;
      var newError1 = evaluateParamError(dataColumns, workColumn, firstRow, lastRow, parameters, hourlyWorkAverage);
      parameters[idx] = p - 0.05;
      var newError2 = evaluateParamError(dataColumns, workColumn, firstRow, lastRow, parameters, hourlyWorkAverage);

      if (newError1 < error && newError1 < newError2) {
        parameters[idx] = p + 0.05;
        error = newError1;
        progress = true;
      } else if (newError2 < error && newError2 < newError1) {
        parameters[idx] = p - 0.05;
        error = newError2;
        progress = true;
      } else {
        parameters[idx] = p;
      }
    }

    if (progress || rep == 0) {
      storeParameters(parameters, dataColumns, hourlyCell, hourlyWorkAverage);
    }
    if (!progress) break;
  }

  email_lines.push("params: " + displayParams(parameters));
  email_lines.push("final error: " + (error / (lastRow + 1 - firstRow)));

  GmailApp.sendEmail(whoAmI, "work estimates", "", { htmlBody: email_lines.join("<br>") });
}
