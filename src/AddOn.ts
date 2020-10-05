import { Doc2MdConverter } from './Doc2MdConverter';

function onOpen() {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Open Doc2Md', 'openDoc2MdInSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function openDoc2MdInSidebar() {
  const ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Doc2Md');
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
    ? selection.getRangeElements().map(rangeElement => rangeElement.getElement())
    : [];
}
