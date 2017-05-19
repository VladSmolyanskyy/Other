function VladFeedback() {

  var ss = SpreadsheetApp.getActiveSheet();
  var label = GmailApp.getUserLabelByName("Test");
  var threads = label.getThreads();
  for (var i=0; i<threads.length; i++)
  {
    var messages = threads[i].getMessages();
    for (var j=0; j<messages.length; j++)
    {
      var msg = messages[j].getPlainBody();
      var sub = messages[j].getSubject();
      var dat = messages[j].getDate();
      ss.appendRow([msg, sub, dat])
    }
  }
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('First item', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi()
  .alert('You clicked the first menu item!');
}
