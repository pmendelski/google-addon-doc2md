export interface ConversionResult {
  markdown: string;
  warnings: Warning[];
}

export interface Warning {
  message: string,
  line: number
}

class Context {
  constructor(
    private markdown: string = '',
    private skipNewLine: boolean = false,
    private readonly warnings: Warning[] = [],
    private readonly footnotes: string[] = [],
    private readonly listCounters: { [name: string] : number } = {}
  ) {}

  public toConversionResult() {
    const footnotes = this.footnotes.join('\n');
    const md = (this.markdown.trim() + "\n\n" + footnotes).trim();
    return {
      markdown: md,
      warnings: this.warnings
    };
  }

  public nextListIndex(key: string): number {
    this.listCounters[key] = this.listCounters[key]
      ? this.listCounters[key] + 1
      : 1;
    return this.listCounters[key];
  }

  public isEmpty(): Boolean {
    return this.markdown.length === 0
      && this.warnings.length === 0;
  }

  public addMarkdown(md: string) {
    if (this.skipNewLine && md.startsWith('\n')) {
      md = md.substr(1, md.length);
    }
    this.skipNewLine = false;
    this.markdown += md;
    this.markdown = this.markdown.replace(/\n\n+/g, '\n\n');
  }

  public addFootnote(footnote: string) {
    const index = this.footnotes.length + 1;
    this.footnotes.push("[^" + index + "]: " + footnote);
    this.addMarkdown("[^" + index + "] ");
  }

  public getNextFootnoteIndex(): number {
    return this.footnotes.length + 1;
  }

  public addWarning(message: string) {
    const line = this.markdown.split('\n').length;
    const warning = { message, line };
    this.warnings.push(warning);
  }

  public skipNextNewLine() {
    this.skipNewLine = true;
  }
}

interface CompositeElement {
  getType(): GoogleAppsScript.Document.ElementType;
  getNumChildren(): number;
  getChild(childIndex: number): GoogleAppsScript.Document.Element;
}

export class Doc2MdConverter {
  private omittedElementTypes: GoogleAppsScript.Document.ElementType[] = [
    DocumentApp.ElementType.TABLE_OF_CONTENTS,
    DocumentApp.ElementType.INLINE_DRAWING,
    DocumentApp.ElementType.UNSUPPORTED
  ];

  public convertDocumentToMarkdown(document: GoogleAppsScript.Document.Document): ConversionResult {
    const ctx = new Context();
    this.processCompositeElement(document.getBody(), ctx)
    return ctx.toConversionResult();
  }

  public convertElementsToMarkdown(elements: GoogleAppsScript.Document.Element[]): ConversionResult {
    let ctx = new Context();
    for (let i = 0; i < elements.length; ++i) {
      this.processElement(elements[i], ctx);
    }
    return ctx.toConversionResult();
  }

  private processCompositeElement(element: CompositeElement, ctx: Context) {
    if (this.omittedElementTypes.indexOf(element.getType()) >= 0) return;
    for (let i = 0; i < element.getNumChildren(); ++i) {
      const child: GoogleAppsScript.Document.Element = element.getChild(i);
      this.processElement(child, ctx);
    }
  }

  private processElement(element: GoogleAppsScript.Document.Element, ctx: Context) {
    const type: GoogleAppsScript.Document.ElementType = element.getType();
    if (this.omittedElementTypes.indexOf(type) >= 0) return;
    if (type == DocumentApp.ElementType.TABLE) {
      this.processTable(element.asTable(), ctx);
    } else if (type == DocumentApp.ElementType.PARAGRAPH) {
      this.processParagraph(element.asParagraph(), ctx);
    } else if (type == DocumentApp.ElementType.TEXT) {
      this.processText(element.asText(), ctx);
    } else if (type == DocumentApp.ElementType.LIST_ITEM) {
      this.processList(element.asListItem(), ctx);
    } else if (type == DocumentApp.ElementType.INLINE_IMAGE) {
      this.processImage(element.asInlineImage(), ctx);
    } else if (type == DocumentApp.ElementType.FOOTNOTE) {
      this.processFootnote(element.asFootnote(), ctx);
    } else if (type == DocumentApp.ElementType.HORIZONTAL_RULE) {
      this.processHorizontalRule(ctx);
    } else {
      this.unrecognizedElement(element, ctx);
    }
  }

  private unrecognizedElement(child: GoogleAppsScript.Document.Element, ctx: Context) {
    ctx.addMarkdown("\n(WARN_UNRECOGNIZED_ELEMENT: " + child.getType() + ")\n");
    ctx.addWarning("Unrecognized element: " + child.getType());
  }

  private processImage(image: GoogleAppsScript.Document.InlineImage, ctx: Context) {
    const linkUrl = image.getLinkUrl();
    const imageMd = image.getAltTitle() !== null
      ? "![" + image.getAltTitle() + "](WARN_REPLACE_IMG_URL)"
      : "![](WARN_REPLACE_IMG_URL)";
    const md = linkUrl !== null && linkUrl.length > 0
      ? "[" + imageMd + "](" + linkUrl + ")"
      : imageMd;
    ctx.addMarkdown(md);
    ctx.addWarning("Image to replace");
  }

  private processFootnote(footnote: GoogleAppsScript.Document.Footnote, ctx: Context) {
    const footnoteContext = new Context();
    this.processCompositeElement(footnote.getFootnoteContents(), footnoteContext);
    ctx.addFootnote(footnoteContext.toConversionResult().markdown);
  }

  private processHorizontalRule(ctx: Context) {
    ctx.addMarkdown("\n\n---\n\n");
  }

  private processParagraph(paragraph: GoogleAppsScript.Document.Paragraph, ctx: Context) {
    this.processParagraphHeading(paragraph.getHeading(), ctx);
    this.processCompositeElement(paragraph, ctx);
  }

  private processList(list: GoogleAppsScript.Document.ListItem, ctx: Context) {
    this.processListPrefix(list, ctx);
    this.processCompositeElement(list, ctx);
  }

  private processParagraphHeading(heading: GoogleAppsScript.Document.ParagraphHeading, ctx: Context) {
    switch (heading) {
      case DocumentApp.ParagraphHeading.HEADING6: ctx.addMarkdown("\n\n###### "); break;
      case DocumentApp.ParagraphHeading.HEADING5: ctx.addMarkdown("\n\n##### "); break;
      case DocumentApp.ParagraphHeading.HEADING4: ctx.addMarkdown("\n\n#### "); break;
      case DocumentApp.ParagraphHeading.HEADING3: ctx.addMarkdown("\n\n### "); break;
      case DocumentApp.ParagraphHeading.HEADING2: ctx.addMarkdown("\n\n## "); break;
      case DocumentApp.ParagraphHeading.HEADING1: ctx.addMarkdown("\n\n# "); break;
      case DocumentApp.ParagraphHeading.SUBTITLE: ctx.addMarkdown("\n\n## "); break;
      case DocumentApp.ParagraphHeading.TITLE: ctx.addMarkdown("\n\n# "); break;
      default: ctx.addMarkdown("\n");
    }
  }

  private processListPrefix(list: GoogleAppsScript.Document.ListItem, ctx: Context) {
    const level = list.getNestingLevel();
    const padding = [...Array(level)]
      .map(() => ' ')
      .join(' ');
    var glyph = list.getGlyphType();
    // Bullet list (<ul>):
    if (glyph === DocumentApp.GlyphType.BULLET
        || glyph === DocumentApp.GlyphType.HOLLOW_BULLET
        || glyph === DocumentApp.GlyphType.SQUARE_BULLET) {
      ctx.addMarkdown("\n" + padding + "* ");
    } else {
      // Ordered list (<ol>):
      const key = list.getListId() + '.' + list.getNestingLevel();
      const index = ctx.nextListIndex(key);
      const prefix = padding + index + ". ";
      ctx.addMarkdown("\n" + padding + index + ". ");
    }
  }

  private processText(text: GoogleAppsScript.Document.Text, ctx: Context) {
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
      } else if (font === 'Courier New') {
        while (i >= 1 && text.getFontFamily(indices[i-1]) === 'Courier New') {
          // detect fonts that are in multiple pieces because of errors on formatting:
          i-=1;
          offset = indices[i];
        }
        value = '`' + result.substring(offset, lastOffset) + '`';
      }
      if (value.trim().length > 0) {
        if (text.isItalic(offset)) {
          value = '*' + value + '*';
        }
        if (text.isBold(offset)) {
          value = "**" + value + "**";
        }
        if (text.isStrikethrough(offset)) {
          value = "~~" + value + "~~";
        }
        if (!url && text.isUnderline(offset)) {
          value = "__" + value + "__";
        }
      }
      result = result.substring(0, offset) + value + result.substring(lastOffset);
      lastOffset = offset;
    }
    ctx.addMarkdown(result.trim());
  }

  private processTable(table: GoogleAppsScript.Document.Table, ctx: Context) {
    if (table.getNumRows() < 1) return;
    this.processTableFirstRow(table.getRow(0), ctx);
    for (let i = 1; i < table.getNumRows(); ++i) {
      this.processTableRow(table.getRow(i), ctx);
    }
  }

  private processTableFirstRow(row: GoogleAppsScript.Document.TableRow, ctx: Context) {
    const cells = row.getNumCells();
    this.processTableRow(row, ctx);
    const underline = [...Array(cells)]
      .map(() => '---')
      .join(' | ');
    ctx.addMarkdown("\n| " + underline + " |");
  }

  private processTableRow(row: GoogleAppsScript.Document.TableRow, ctx: Context) {
    ctx.addMarkdown('\n| ');
    for (let i = 0; i < row.getNumCells(); ++i) {
      ctx.skipNextNewLine();
      this.processCompositeElement(row.getCell(i), ctx);
      ctx.addMarkdown(' | ');
    }
  }
}
