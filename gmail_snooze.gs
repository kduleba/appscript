var PAGE_SIZE = 100;  // Don't use more than 100 because thats the max you can write

function getLabelName(i) {
  return "Snooze/Snooze " + i + " days";
}

function setup() {
  GmailApp.createLabel("Snooze");
  for (var i = 1; i <= 7; ++i) {
    GmailApp.createLabel(getLabelName(i));
  }
}

function moveSnoozes() {
  var oldLabel, newLabel, page;
  var awakeLabel = GmailApp.getUserLabelByName("Snooze/Awake");
  var awoken = 0;
  
  for (var i = 1; i <= 7; ++i) {
    newLabel = oldLabel;
    oldLabel = GmailApp.getUserLabelByName(getLabelName(i));
    page = null;
    while(!page || page.length == PAGE_SIZE) {
      page = oldLabel.getThreads(0, PAGE_SIZE);
      if (page.length > 0) {
        if (newLabel) {
          newLabel.addToThreads(page);
        } else {
          GmailApp.moveThreadsToInbox(page);
          GmailApp.markThreadsImportant(page);
          awakeLabel.addToThreads(page);
          awoken = awoken + page.length;
        }     
        oldLabel.removeFromThreads(page);
      }  
    }
  }
  
  GmailApp.sendEmail("kduleba@google.com", "Snooze status", "works fine!\nAwoken threads: " + awoken);
}
