import { Doc2MdConverter } from './Doc2MdConverter';

function onOpen() {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Convert', 'convertDoc2MdInSidebar')
    .addItem('Help', 'showHelp')
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

function showHelp() {
  const html = HtmlService.createHtmlOutputFromFile('help')
    .setWidth(600)
    .setHeight(425);
  DocumentApp.getUi().showModalDialog(html, 'Doc2Md - Help');
}

function loadMarkdown() {
  const selectedElements = getSelectedElements();
  if (selectedElements.length) {
    return new Doc2MdConverter()
      .convertElementsToMarkdown(selectedElements);
  }
  return new Doc2MdConverter()
    .convertDocumentToMarkdown(DocumentApp.getActiveDocument());
}

function getSelectedElements() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  return selection
    ? selection.getSelectedElements().map(rangeElement => rangeElement.getElement())
    : [];
}
