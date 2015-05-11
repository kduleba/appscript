// This script, when put in a spreatsheet, will dump recent email interactions of a person that ended up in your mailbox.
// Input: username to process (A1 cell of the spreadsheet)

// Basic version: 287 (5.96 min)
// v1: 760 (6.07 min)

var PAGE_SIZE = 100;
var NUM_DAYS = 360;
var OLD_DATE = new Date((new Date()).getTime() - NUM_DAYS * 24 * 3600 * 1000);

function extractEmail(email) {
  if (!email) return;
  var beg = 0;
  var end = email.length;
  
  for (var i = 0; i < email.length; i++) {
    if (email[i] == '<') beg = i + 1;
    if (email[i] == '>') end = i;
  }
  
  email = email.substr(beg, end - beg);
  
  var at = email.length;
  for (var i = 0; i < email.length; i++) {
    if (email[i] == '@' || email[i] == '+') {
      at = i;
      break;
    }
  }
  
  return email.substr(0, at);
}

function getRow(sheet, peer, rowCache) {
  if (peer in rowCache) return rowCache[peer];
  return sheet.getLastRow() + 1;
}

function prePopulateCache(sheet) {
  var rowCache = {};
  var lastRow = sheet.getLastRow();
  for (var row = sheet.getFrozenRows() + 1; row <= lastRow; row++) {
    var range = sheet.getRange(row, 1, 1, 3);
    var email = range.getValues()[0][0];
    rowCache[email] = row;
  }
  return rowCache;
}

function getSize(toList) {
  var commas = 0;
  for (var i = 0; i < toList.length; i++) {
    if (toList[i] == ',') commas++;
  }
  return 1.0 / (1.0 + commas);
}

function processThreadMessages(sheet, rowCache, username, messages) {
  var demotion = Math.sqrt(messages.length);
  
  for (var i = 0; i < messages.length; i++) {
    sheet.getRange(1, 4).setValue((i + 1) + ' / ' + messages.length);
    
    if (messages[i].getDate() >= OLD_DATE) {
      var toList = messages[i].getTo();
      var ccList = messages[i].getCc();
      
      var w = 0;
      if (toList.indexOf(username) != -1) {
        var w = 1.0 / (getSize(toList) + getSize(ccList) / 4.0 + 1.0) / demotion;
      } else if (ccList.indexOf(username) != -1) { 
        var w = 1.0 / (getSize(toList) * 4 + getSize(ccList) + 1.0) / demotion;
      }
      
      if (w > 0) {
        var peer = extractEmail(messages[i].getFrom());
        if (peer) {
          var row = getRow(sheet, peer, rowCache);
          rowCache[peer] = row;
          var range = sheet.getRange(row, 1, 1, 3);
          var neww = range.getCell(1, 2).getValue() + w;
          range.getCell(1, 1).setValue(peer);
          range.getCell(1, 2).setValue(neww);
        }
      }
    }
  }
}

function processThreadBatch(threads, messages, index, username, sheet, rowCache) {
  var thread = threads[index];
  var date = thread.getLastMessageDate();
  if (date < OLD_DATE) return false;
  var alreadyProcessedDate = sheet.getRange(1, 3).getValue();
  if (alreadyProcessedDate > 0 && date.getTime() > alreadyProcessedDate) {
    sheet.getRange(1, 7).setValue(date);
    return rowCache;
  }
    
  sheet.getRange(1, 2).setValue(date);
  sheet.getRange(1, 3).setValue(date.getTime());

  var messages = messages[index];
  
  processThreadMessages(sheet, rowCache, username, messages);
  
  return rowCache;
}

function getDateRestrict(sheet) {
  var cell = sheet.getRange(1, 3).getValue();
  if (!cell) return;
  
  var alreadyProcessedDate = new Date(cell + 24 * 3600 * 1000);
  var yyyy = alreadyProcessedDate.getFullYear();
  var mm = alreadyProcessedDate.getMonth() + 1;
  var dd = alreadyProcessedDate.getDate();
  if (mm < 10) mm = "0" + mm;
  if (dd < 10) dd = "0" + dd;
  return yyyy + "/" + mm + "/" + dd;
}

function getPeers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");
  var username = sheet.getRange(1, 1).getValue();
  var startTime = (new Date()).getTime();
  
  var gmailSearchDateRestrict = getDateRestrict(sheet);
  var gmailSearch = 'to:' + username;
  if (gmailSearchDateRestrict) {
    gmailSearch = gmailSearch + " older:" + gmailSearchDateRestrict;
  }
  
  sheet.sort(2, false);

  var offset = 0;
  var page = null;
  var rowCache = prePopulateCache(sheet);
  while(!page || page.length == PAGE_SIZE) {    
    page = GmailApp.search(gmailSearch, offset, PAGE_SIZE);
    GmailApp.refreshThreads(page);
    offset += PAGE_SIZE;
    var msgCache = GmailApp.getMessagesForThreads(page);
    
    for (var i = 0; i < page.length; i++) {
      var timeElapsed = (((new Date()).getTime() - startTime) / 1000 / 60.0).toFixed(2);
      sheet.getRange(1, 5).setValue((offset - PAGE_SIZE + i) + ' (' + timeElapsed + ' min)');
      if (sheet.getRange(1, 1).getValue().length == 0) {
        sheet.getRange(1, 1).setValue("ABORTED");
        return;
      }      
      
      var newRowCache = processThreadBatch(page, msgCache, i, username, sheet, rowCache);
      if (!newRowCache) {
        page.pop();
        break;
      }
      rowCache = newRowCache;
    }
  }
  
  sheet.getRange(1, 1).setValue(username + " done");
};
