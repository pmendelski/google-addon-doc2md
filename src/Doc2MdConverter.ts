export interface DocConversionResult {
  title: string;
  markdown: string;
}

interface CompositeElement {
  getType(): GoogleAppsScript.Document.ElementType;
  getNumChildren(): number;
  getChild(childIndex: number): GoogleAppsScript.Document.Element;
}
export class Doc2MdConverter {
  private omittedElementTypes: GoogleAppsScript.Document.ElementType[] = [
    DocumentApp.ElementType.TABLE_OF_CONTENTS,
    DocumentApp.ElementType.HORIZONTAL_RULE,
    DocumentApp.ElementType.INLINE_DRAWING,
    DocumentApp.ElementType.UNSUPPORTED
  ];
  private listCounters: { [name: string] : number } = {};

  public convertToMarkdown(document: GoogleAppsScript.Document.Document): DocConversionResult {
    this.listCounters = {};
    const result = this.processCompositeElement(document.getBody());
    return {
      title: document.getName(),
      markdown: result
    };
  }

  private processCompositeElement(element: CompositeElement): string {
    if (this.omittedElementTypes.indexOf(element.getType()) >= 0) return '';
    return [...Array(element.getNumChildren())]
      .map((_, index) => element.getChild(index))
      .map(child => this.processElement(child))
      .filter(result => result.length > 0)
      .join('\n')
      .replace(/\n\n+/g, '\n\n')
      .trim();
  }

  private processElement(element: GoogleAppsScript.Document.Element): string {
    const type: GoogleAppsScript.Document.ElementType = element.getType();
    if (this.omittedElementTypes.indexOf(type) >= 0) return '';
    if (type == DocumentApp.ElementType.TABLE) {
      return this.processTable(element.asTable());
    } else if (type == DocumentApp.ElementType.PARAGRAPH) {
      return this.processParagraph(element.asParagraph());
    } else if (type == DocumentApp.ElementType.TEXT) {
      return this.processText(element.asText());
    } else if (type == DocumentApp.ElementType.LIST_ITEM) {
      return this.processList(element.asListItem());
    } else if (type == DocumentApp.ElementType.INLINE_IMAGE) {
      return this.processImage();
    } else {
      return this.unrecognizedElement(element);
    }
  }

  private unrecognizedElement(child: GoogleAppsScript.Document.Element) {
    return "(WARN_UNRECOGNIZED_ELEMENT: " + child.getType() + ")";
  }

  private processImage(): string {
    return `![WARN_IMG]()`;
  }

  private processParagraph(paragraph: GoogleAppsScript.Document.Paragraph): string {
    const prefix = this.processParagraphHeading(paragraph.getHeading());
    const text = this.processCompositeElement(paragraph);
    return text.length > 0 ? prefix + text : '\n';
  }

  private processList(list: GoogleAppsScript.Document.ListItem): string {
    const prefix = this.processListPrefix(list);
    const text = this.processText(list.editAsText());
    return text.length > 0 ? prefix + text : '';
  }

  private processParagraphHeading(heading: GoogleAppsScript.Document.ParagraphHeading): string {
    switch (heading) {
      case DocumentApp.ParagraphHeading.HEADING6: return "###### ";
      case DocumentApp.ParagraphHeading.HEADING5: return "##### ";
      case DocumentApp.ParagraphHeading.HEADING4: return "#### ";
      case DocumentApp.ParagraphHeading.HEADING3: return "### ";
      case DocumentApp.ParagraphHeading.HEADING2: return "## ";
      case DocumentApp.ParagraphHeading.HEADING1: return "# ";
      case DocumentApp.ParagraphHeading.SUBTITLE: return "## ";
      case DocumentApp.ParagraphHeading.TITLE: return "# ";
      default: return "";
    }
  }

  private processListPrefix(list: GoogleAppsScript.Document.ListItem): string {
    const level = list.getNestingLevel();
    const padding = [...Array(level)]
      .map(() => ' ')
      .join(' ');
    var glyph = list.getGlyphType();
    // Bullet list (<ul>):
    if (glyph === DocumentApp.GlyphType.BULLET
        || glyph === DocumentApp.GlyphType.HOLLOW_BULLET
        || glyph === DocumentApp.GlyphType.SQUARE_BULLET) {
      return padding + "* ";
    } else {
      // Ordered list (<ol>):
      const key = list.getListId() + '.' + list.getNestingLevel();
      const counter = this.listCounters[key] ? this.listCounters[key] + 1 : 1;
      this.listCounters[key] = counter;
      return padding + counter + ". ";
    }
  }

  private processText2(text: GoogleAppsScript.Document.Text): string {
    const indices = text.getTextAttributeIndices();
    let result = "text:" + text.getText() + "\n";
    let lastOffset = result.length;

    for (let i = indices.length - 1; i >= 0; i--) {
      let offset = indices[i];
      let value = text.getText().substring(offset, lastOffset);
      result += i + ": (" + offset + ", " + lastOffset + "): value: " + value + "\n";
      lastOffset = offset;
    }
    return result;
  }

  private processText(text: GoogleAppsScript.Document.Text): string {
    const indices = text.getTextAttributeIndices();
    let result = text.getText();
    let lastOffset = result.length;

    for (let i = indices.length - 1; i >= 0; i--) {
      let offset = indices[i];
      let url = text.getLinkUrl(offset);
      let font = text.getFontFamily(offset);
      let value = result.substring(offset, lastOffset);
      if (url) {
        while (i >= 1 && indices[i-1] == offset-1 && text.getLinkUrl(indices[i-1]) === url) {
          // detect links that are in multiple pieces because of errors on formatting:
          i -= 1;
          offset = indices[i];
        }
        value = '[' + result.substring(offset, lastOffset) + '](' + url + ')';
      } else if (font === 'COURIER_NEW') {
        while (i >= 1 && text.getFontFamily(indices[i-1]) === 'COURIER_NEW') {
          // detect fonts that are in multiple pieces because of errors on formatting:
          i-=1;
          offset = indices[i];
        }
        value = '`' + result.substring(offset, lastOffset) + '`';
      }
      if (text.isItalic(offset)) {
        value = '*' + value + '*';
      }
      if (text.isBold(offset)) {
        value = "**" + value + "**";
      }
      if (text.isUnderline(offset)) {
        value = "__" + value + "__";
      }
      result = result.substring(0, offset) + value + result.substring(lastOffset);
      lastOffset = offset;
    }
    return result;
  }

  private processTable(table: GoogleAppsScript.Document.Table): string {
    if (table.getNumRows() < 1) return '';
    const rows = table.getNumRows();
    const tableHeader = this.processTableFirstRow(table.getRow(0));
    const tableBody = [...Array(rows - 1)]
      .map((_, index) => this.processTableRow(table.getRow(index + 1)))
      .join('');
    return '\n' + tableHeader + tableBody + '\n';
  }

  private processTableFirstRow(row: GoogleAppsScript.Document.TableRow): string {
    const cells = row.getNumCells();
    const underline = [...Array(cells)]
      .map(() => '---')
      .join(' | ');
    return this.processTableRow(row)
      + "| " + underline + " |\n";
  }

  private processTableRow(row: GoogleAppsScript.Document.TableRow): string {
    const cells = row.getNumCells();
    const processed = [...Array(cells)]
      .map((_, index) => this.processCompositeElement(row.getCell(index)))
      .join(' | ');
    return '| ' + processed + ' |\n';
  }
}
