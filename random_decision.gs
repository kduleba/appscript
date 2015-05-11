function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function getColumn(values, numRows, colName) {
  var col = -1;
  var tab = new Array();
  
  for (var i = 0; i < values[0].length; i++) {
    if (values[0][i] == colName) col = i;
  }
  
  if (col == -1) return tab; 
  for (var i = 1; i < numRows; i++) {
    tab[i] = values[i][col];
  }
  return tab;
}
  
function getFrequency(values, numRows) {
  return getColumn(values, numRows, "frequency");

}

function generateTask() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var msg = "";
  var frequency = getFrequency(values, numRows);
  var tasks = getColumn(values, numRows, "activity");

  var prefixSum = new Array();
  
  var sum = 0;
  for (var i = 0; i < frequency.length; i++) {
    var left = 0;
    if (!isNaN(parseFloat(frequency[i]))) left += parseFloat(frequency[i]);
    if (left < 0) left = 0;
    sum += left;
    prefixSum[i] = sum;
  }
  
  var rnd = Math.random() * sum;
  var task = "?";
  for (var i = 0; i < prefixSum.length; i++) {
    if (rnd <= prefixSum[i]) {
      task = tasks[i];
      break;
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(task);
};
