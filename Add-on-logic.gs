// Add-on logic
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createAddonMenu()
    .addItem("Begin Metadata Collection", "openSidebar")
    .addItem("Enter API key", "setKey")
    .addItem("Enter retmax", "setRetmax")
    .addToUi(); 
}

function onInstall() {
  onOpen();
}

function openSidebar() {
  var html = HtmlService
      .createTemplateFromFile('sidebar')
      .evaluate()
      .setTitle('SRA Metadata Collector');
  SpreadsheetApp.getUi().showSidebar(html);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
