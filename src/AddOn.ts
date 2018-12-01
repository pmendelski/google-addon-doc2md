import { DocConverter } from './DocConverter';

export const onOpen = () => {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Convert Doc to Markdown', 'convertDoc2MdInSidebar')
      .addToUi();
}

export const onInstall = () => {
  onOpen();
}

export const convertInModal = () => {
  const ui = HtmlService.createHtmlOutputFromFile('modal')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(800)
    .setHeight(625);
  DocumentApp.getUi().showModelessDialog(ui, 'Markdown');
}

export const convertDoc2MdInSidebar = () => {
  const ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Markdown');
  DocumentApp.getUi().showSidebar(ui);
}


export const loadMarkdown = () => {
  return new DocConverter()
    .convertToMarkdown(DocumentApp.getActiveDocument());
}
