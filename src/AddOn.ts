import { DocConverter } from './DocConverter';

export const onOpen = () => {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Convert in Modall', 'convertInModal')
      .addItem('Convert in Sidebar', 'convertInSidebar')
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

export const convertInSidebar = () => {
  const ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Markdown');
  DocumentApp.getUi().showSidebar(ui);
}


export const loadMarkdown = () => {
  return new DocConverter()
    .convertToMarkdown(DocumentApp.getActiveDocument());
}
