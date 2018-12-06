import { Doc2MdConverter } from './Doc2MdConverter';

function onOpen() {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Convert Doc to Markdown', 'convertDoc2MdInSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function convertDoc2MdInSidebar() {
  const ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Markdown - Doc2Md');
  DocumentApp.getUi().showSidebar(ui);
}

function loadMarkdown() {
  const document = DocumentApp.getActiveDocument();
  return new Doc2MdConverter()
    .convertToMarkdown(document);
}
