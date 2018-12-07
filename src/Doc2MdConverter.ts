export interface ConversionResult {
  markdown: string;
  warnings: Warning[];
}

class ListCounters {
  constructor(
    private readonly counters: { [name: string] : number } = {}) {}

  public getCounter(key: string): number {
    return this.counters[key] || 0;
  }

  public incrementCounter(key: string): ListCounters {
    const copy: { [name: string] : number } = Object.keys(this.counters)
      .reduce((acc: { [name: string] : number }, i: string) => {
        acc[i] = this.counters[i];
        return acc;
      }, {});
    copy[key] = copy[key] ? copy[key] + 1 : 1;
    return new ListCounters(copy);
  }

  public merge(other: ListCounters): ListCounters {
    const copy: { [name: string] : number } = {};
    Object.keys(this.counters)
      .forEach((key: string) => {
        copy[key] = this.counters[key];
      });
    Object.keys(other.counters)
      .forEach((key: string) => {
        const max = copy[key] && copy[key] > other.counters[key] ? copy[key] : other.counters[key];
        copy[key] = max;
      });
    return new ListCounters(copy);
  }
}

class Warning {
  constructor(
    readonly message: string,
    readonly line: number = 1
  ) {}

  public shiftLine(delta: number): Warning {
    return new Warning(this.message, delta + this.line);
  }
}
class ProcessingResult {
  static empty(): ProcessingResult {
    return new ProcessingResult();
  }

  static newLine(): ProcessingResult {
    return new ProcessingResult('\n');
  }

  static fromContext(context: Context): ProcessingResult {
    return new ProcessingResult('', [], context.listCounters);
  }

  constructor(
    readonly markdown: string = '',
    readonly warnings: Warning[] = [],
    readonly listCounters: ListCounters = new ListCounters()
  ) {}

  public toContext(): Context {
    return new Context(this.listCounters);
  }

  public merge(other: ProcessingResult): ProcessingResult {
    const lines = this.markdown.split('\n').length - 1;
    const otherWarnings = other.warnings.map(warning => warning.shiftLine(lines))
    return new ProcessingResult(
      (this.markdown + other.markdown).replace(/\n\n+/g, '\n\n'),
      [...this.warnings, ...otherWarnings],
      this.listCounters.merge(other.listCounters)
    )
  }

  public mergeWithNewLine(other: ProcessingResult): ProcessingResult {
    if (this.isEmpty() && other.markdown == '\n') return this;
    if (this.isEmpty()) return other;
    if (this.markdown[this.markdown.length - 1] === '\n') return this.merge(other);
    return this.mergeMarkdown('\n').merge(other);
  }

  public mergeMarkdown(md: string): ProcessingResult {
    return new ProcessingResult(
      this.markdown + md,
      this.warnings,
      this.listCounters
    )
  }

  public isEmpty(): Boolean {
    return this.markdown.length === 0 && this.warnings.length === 0;
  }
}

class Context {
  constructor(
    readonly listCounters: ListCounters = new ListCounters()
  ) {}
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

  public convertDocumentToMarkdown(document: GoogleAppsScript.Document.Document): ConversionResult {
    const result = this.processCompositeElement(document.getBody(), new Context());
    return {
      markdown: result.markdown,
      warnings: result.warnings
    };
  }

  public convertElementsToMarkdown(elements: GoogleAppsScript.Document.Element[]): ConversionResult {
    let result = ProcessingResult.empty();
    for (let i = 0; i < elements.length; ++i) {
      const element: GoogleAppsScript.Document.Element = elements[i];
      const childResult: ProcessingResult = (element as any).getChild
        ? this.processCompositeElement(element as any as CompositeElement, result.toContext())
        : this.processElement(element, result.toContext());
      if (!childResult.isEmpty()) {
        result = result.mergeWithNewLine(childResult);
      }
    }
    return result;
  }

  private processCompositeElement(element: CompositeElement, initialContext: Context): ProcessingResult {
    if (this.omittedElementTypes.indexOf(element.getType()) >= 0) return ProcessingResult.empty();
    let result = ProcessingResult.empty();
    for (let i = 0; i < element.getNumChildren(); ++i) {
      const child: GoogleAppsScript.Document.Element = element.getChild(i);
      const childResult: ProcessingResult = this.processElement(child, result.toContext());
      if (!childResult.isEmpty()) {
        result = result.mergeWithNewLine(childResult);
      }
    }
    return result;
  }

  private processElement(element: GoogleAppsScript.Document.Element, ctx: Context): ProcessingResult {
    const type: GoogleAppsScript.Document.ElementType = element.getType();
    if (this.omittedElementTypes.indexOf(type) >= 0) return ProcessingResult.empty();
    if (type == DocumentApp.ElementType.TABLE) {
      return this.processTable(element.asTable(), ctx);
    } else if (type == DocumentApp.ElementType.PARAGRAPH) {
      return this.processParagraph(element.asParagraph(), ctx);
    } else if (type == DocumentApp.ElementType.TEXT) {
      return this.processText(element.asText());
    } else if (type == DocumentApp.ElementType.LIST_ITEM) {
      return this.processList(element.asListItem(), ctx);
    } else if (type == DocumentApp.ElementType.INLINE_IMAGE) {
      return this.processImage();
    } else {
      return this.unrecognizedElement(element);
    }
  }

  private unrecognizedElement(child: GoogleAppsScript.Document.Element): ProcessingResult {
    return new ProcessingResult(
      "(WARN_UNRECOGNIZED_ELEMENT: " + child.getType() + ")",
      [new Warning("Unrecognized element: " + child.getType())]
    );
  }

  private processImage(): ProcessingResult {
    return new ProcessingResult(
      "![WARN_REPLACE_IMG]()",
      [new Warning("Image to replace")]
    );
  }

  private processParagraph(paragraph: GoogleAppsScript.Document.Paragraph, context: Context): ProcessingResult {
    const prefix = this.processParagraphHeading(paragraph.getHeading());
    const content = this.processCompositeElement(paragraph, context);
    return content.isEmpty() ? ProcessingResult.newLine() : prefix.merge(content);
  }

  private processList(list: GoogleAppsScript.Document.ListItem, context: Context): ProcessingResult {
    const prefix = this.processListPrefix(list, context);
    const text = this.processText(list.editAsText());
    return text.isEmpty() ? ProcessingResult.empty() : prefix.merge(text);
  }

  private processParagraphHeading(heading: GoogleAppsScript.Document.ParagraphHeading): ProcessingResult {
    switch (heading) {
      case DocumentApp.ParagraphHeading.HEADING6: return new ProcessingResult("\n###### ");
      case DocumentApp.ParagraphHeading.HEADING5: return new ProcessingResult("\n##### ");
      case DocumentApp.ParagraphHeading.HEADING4: return new ProcessingResult("\n#### ");
      case DocumentApp.ParagraphHeading.HEADING3: return new ProcessingResult("\n### ");
      case DocumentApp.ParagraphHeading.HEADING2: return new ProcessingResult("\n## ");
      case DocumentApp.ParagraphHeading.HEADING1: return new ProcessingResult("\n# ");
      case DocumentApp.ParagraphHeading.SUBTITLE: return new ProcessingResult("\n## ");
      case DocumentApp.ParagraphHeading.TITLE: return new ProcessingResult("\n# ");
      default: return ProcessingResult.empty();
    }
  }

  private processListPrefix(list: GoogleAppsScript.Document.ListItem, context: Context): ProcessingResult {
    const level = list.getNestingLevel();
    const padding = [...Array(level)]
      .map(() => ' ')
      .join(' ');
    var glyph = list.getGlyphType();
    // Bullet list (<ul>):
    if (glyph === DocumentApp.GlyphType.BULLET
        || glyph === DocumentApp.GlyphType.HOLLOW_BULLET
        || glyph === DocumentApp.GlyphType.SQUARE_BULLET) {
      return new ProcessingResult(padding + "* ");
    } else {
      // Ordered list (<ol>):
      const key = list.getListId() + '.' + list.getNestingLevel();
      const counters = context.listCounters.incrementCounter(key);
      const prefix = padding + counters.getCounter(key) + ". ";
      return new ProcessingResult(prefix, [], counters);
    }
  }

  private processText(text: GoogleAppsScript.Document.Text): ProcessingResult {
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
      if (text.isItalic(offset)) {
        value = '*' + value + '*';
      }
      if (text.isBold(offset)) {
        value = "**" + value + "**";
      }
      if (!url && text.isUnderline(offset)) {
        value = "__" + value + "__";
      }
      result = result.substring(0, offset) + value + result.substring(lastOffset);
      lastOffset = offset;
    }
    return new ProcessingResult(result);
  }

  private processTable(table: GoogleAppsScript.Document.Table, ctx: Context): ProcessingResult {
    if (table.getNumRows() < 1) return new ProcessingResult();
    let result = this.processTableFirstRow(table.getRow(0), ctx);
    for (let i = 1; i < table.getNumRows(); ++i) {
      const rowResult = this.processTableRow(table.getRow(i), result.toContext());
      result = result.mergeWithNewLine(rowResult);
    }
    return result;
  }

  private processTableFirstRow(row: GoogleAppsScript.Document.TableRow, ctx: Context): ProcessingResult {
    const cells = row.getNumCells();
    const titleRow = this.processTableRow(row, ctx);
    const underline = [...Array(cells)]
      .map(() => '---')
      .join(' | ');
    return titleRow.mergeMarkdown("\n| " + underline + " |\n");
  }

  private processTableRow(row: GoogleAppsScript.Document.TableRow, initialContext: Context): ProcessingResult {
    let result = ProcessingResult.fromContext(initialContext);
    for (let i = 0; i < row.getNumCells(); ++i) {
      const cellResult = this.processCompositeElement(row.getCell(i), result.toContext());
      result = result
        .merge(cellResult)
        .mergeMarkdown(' | ');
    }
    return new ProcessingResult('| ').merge(result);
  }
}
