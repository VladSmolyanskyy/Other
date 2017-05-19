function VladFeedback() {

  var ss = SpreadsheetApp.getActiveSheet();
  var label = GmailApp.getUserLabelByName("Test");
  var threads = label.getThreads();

  for (var i=0; i<threads.length; i++)
  {
    var messages = threads[i].getMessages();
    var lin = threads[i].getPermalink();

    for (var j=0; j<messages.length; j++)
    {
//      var msg = messages[j].getBody();
      var msg = messages[j].getPlainBody();
      var sub = messages[j].getSubject();
      // var dat = messages[j].getDate();
      var fro = messages[j].getFrom();


      ss.appendRow([fro, sub, lin, msg])
    }
//      threads[i].removeLabel(label);
  }
}

function AddNewSheetWDate () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet(0);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // var newSheetName = new Date();
  var formattedDate = Utilities.formatDate(new Date(), "EST", "EEE, MMM d");
  sheet.setName(formattedDate)
  
}

//http://stackoverflow.com/questions/32890847/google-apps-scripts-extract-data-from-gmail-into-a-spreadsheet;



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Pull Emails', 'menuItem1')
      .addSeparator()
      .addItem('New Sheet w/Date', 'menuItem2')
      // .addSubMenu(ui.createMenu('Sub-menu')
      //     .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  VladFeedback();
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  AddNewSheetWDate();
}
