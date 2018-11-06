import { DocConverter } from './DocConverter';

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Convert in Modall', 'convertInModal')
      .addItem('Convert in Sidebar', 'convertInSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function convertInModal() {
  var ui = HtmlService.createHtmlOutputFromFile('modal')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(800)
      .setHeight(625);
  DocumentApp.getUi().showModelessDialog(ui, 'Markdown');
}

function convertInSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Markdown');
  DocumentApp.getUi().showSidebar(ui);
}


function loadMarkdown() {
  return new DocConverter()
    .convertToMarkdown(DocumentApp.getActiveDocument());
}
