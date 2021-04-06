/** @OnlyCurrentDoc */

function onOpen() {
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [
    {name: 'Set Up and Run New Titles Report', functionName: 'showSidebar'}
  ];
  spreadsheet.addMenu('Report', menuItems);
  const text = 'Activate "Set Up and Run New Titles Report" under the "Report" menu.';
  SpreadsheetApp.getUi().alert(text).CLOSE;
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile('Sidebar.html')
      .evaluate()
      .setTitle('WorldCat New Titles Report')
  SpreadsheetApp.getUi() 
      .showSidebar(template);    
}

function include(filename) {
 return HtmlService.createHtmlOutputFromFile(filename)
  .getContent();
}
