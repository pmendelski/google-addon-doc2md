export interface DocConversionResult {
  title: string;
  markdown: string;
}

export class DocConverter {
  public convertToMarkdown(document: GoogleAppsScript.Document.Document): DocConversionResult {
    const listCounters: { [name: string] : number } = {};
    let result = "";

    for (let i = 0; i < document.getBody().getNumChildren(); i++) {
      const child: GoogleAppsScript.Document.Element = document.getBody().getChild(i);
      const childMarkdown: string = this.processElement(child, listCounters);
      if (childMarkdown !== null && childMarkdown.length > 0) {
        result += childMarkdown + "\n\n";
      }
    }
    return {
      title: document.getName(),
      markdown: result
    };
  }

  private processElement(element: GoogleAppsScript.Document.Element, listCounters: { [name: string] : number } ): string {
    const type: GoogleAppsScript.Document.ElementType = element.getType();
    if (type == DocumentApp.ElementType.TABLE_OF_CONTENTS) {
      element.asTableOfContents()
    }
  }

  private processTableOfContents(element: GoogleAppsScript.Document.TableOfContents): string {
    return {
      text: "[[MD_TOC]]"
    };
  }

  private processElement2(element: GoogleAppsScript.Document.Element, listCounters: { [name: string] : number } ): string {
    // First, check for things that require no processing.
    if (element.getNumChildren()==0) {
      return null;
    }
    // Punt on TOC.
    if (element.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
      return {"text": "[[TOC]]"};
    }

    // Set up for real results.
    var result = {};
    var pOut = "";
    var textElements: string[] = [];

    // Handle Table elements. Pretty simple-minded now, but works for simple tables.
    // Note that Markdown does not process within block-level HTML, so it probably
    // doesn't make sense to add markup within tables.
    if (element.getType() === DocumentApp.ElementType.TABLE) {
      var table = element.as
      textElements.push("<table>\n");
      var nCols = element.getChild(0).getNumCells();
      for (var i = 0; i < element.getNumChildren(); i++) {
        textElements.push("  <tr>\n");
        // process this row
        for (var j = 0; j < nCols; j++) {
          textElements.push("    <td>" + element.getChild(i).getChild(j).getText() + "</td>\n");
        }
        textElements.push("  </tr>\n");
      }
      textElements.push("</table>\n");
    }

    // Process various types (ElementType).
    for (var i = 0; i < element.getNumChildren(); i++) {
      var t = element.getChild(i).getType();

      if (t === DocumentApp.ElementType.TABLE_ROW) {
        // do nothing: already handled TABLE_ROW
      } else if (t === DocumentApp.ElementType.TEXT) {
        var txt = element.getChild(i);
        pOut += txt.getText();
        textElements.push(txt);
      } else if (t === DocumentApp.ElementType.INLINE_IMAGE) {
        textElements.push('![INLINED_IMG](IMG_URL)');
      } else if (t === DocumentApp.ElementType.PAGE_BREAK) {
        // ignore
      } else if (t === DocumentApp.ElementType.HORIZONTAL_RULE) {
        textElements.push('* * *\n');
      } else if (t === DocumentApp.ElementType.FOOTNOTE) {
        textElements.push(' (NOTE: '+element.getChild(i).getFootnoteContents().getText()+')');
      } else {
        throw "Paragraph "+ index +" of type " + element.getType() + " has an unsupported child: "
        + t + " "+(element.getChild(i)["getText"] ? element.getChild(i).getText():'')+" index="+index;
      }
    }

    if (textElements.length==0) {
      // Isn't result empty now?
      return result;
    }

    // evb: Add source pretty too. (And abbreviations: src and srcp.)
    // process source code block:
    if (/^\s*---\s+srcp\s*$/.test(pOut) || /^\s*---\s+source pretty\s*$/.test(pOut)) {
      result.sourcePretty = "start";
    } else if (/^\s*---\s+src\s*$/.test(pOut) || /^\s*---\s+source code\s*$/.test(pOut)) {
      result.source = "start";
    } else if (/^\s*---\s+class\s+([^ ]+)\s*$/.test(pOut)) {
      result.inClass = "start";
      result.className = RegExp.$1;
    } else if (/^\s*---\s*$/.test(pOut)) {
      result.source = "end";
      result.sourcePretty = "end";
      result.inClass = "end";
    } else {

      prefix = this.findPrefix(inSrc, element, listCounters);

      var pOut = "";
      for (var i=0; i<textElements.length; i++) {
        pOut += this.processTextElement(inSrc, textElements[i]);
      }

      // replace Unicode quotation marks
      pOut = pOut.replace('\u201d', '"').replace('\u201c', '"');

      result.text = prefix+pOut;
    }

    return result;
  }

  // Add correct prefix to list items.
  private findPrefix(inSrc, element, listCounters): string {
    var prefix="";
    if (!inSrc) {
      if (element.getType()===DocumentApp.ElementType.PARAGRAPH) {
        var paragraphObj = element;
        switch (paragraphObj.getHeading()) {
          // Add a # for each heading level. No break, so we accumulate the right number.
          case DocumentApp.ParagraphHeading.HEADING6: prefix+="#";
          case DocumentApp.ParagraphHeading.HEADING5: prefix+="#";
          case DocumentApp.ParagraphHeading.HEADING4: prefix+="#";
          case DocumentApp.ParagraphHeading.HEADING3: prefix+="#";
          case DocumentApp.ParagraphHeading.HEADING2: prefix+="#";
          case DocumentApp.ParagraphHeading.HEADING1:
          case DocumentApp.ParagraphHeading.SUBTITLE:
          case DocumentApp.ParagraphHeading.TITLE: prefix+="# ";
          default:
        }
      } else if (element.getType()===DocumentApp.ElementType.LIST_ITEM) {
        var listItem = element;
        var nesting = listItem.getNestingLevel()
        for (var i=0; i<nesting; i++) {
          prefix += "    ";
        }
        var gt = listItem.getGlyphType();
        // Bullet list (<ul>):
        if (gt === DocumentApp.GlyphType.BULLET
            || gt === DocumentApp.GlyphType.HOLLOW_BULLET
            || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
          prefix += "* ";
        } else {
          // Ordered list (<ol>):
          var key = listItem.getListId() + '.' + listItem.getNestingLevel();
          var counter = listCounters[key] || 0;
          counter++;
          listCounters[key] = counter;
          prefix += counter+". ";
        }
      }
    }
    return prefix;
  }

  private processTextElement(inSrc: boolean, txt): string {
    if (typeof(txt) === 'string') {
      return txt;
    }

    var pOut = txt.getText();
    if (! txt.getTextAttributeIndices) {
      return pOut;
    }

    var attrs=txt.getTextAttributeIndices();
    var lastOff=pOut.length;

    for (var i=attrs.length-1; i>=0; i--) {
      var off=attrs[i];
      var url=txt.getLinkUrl(off);
      var font=txt.getFontFamily(off);
      if (url) {  // start of link
        if (i>=1 && attrs[i-1]==off-1 && txt.getLinkUrl(attrs[i-1])===url) {
          // detect links that are in multiple pieces because of errors on formatting:
          i-=1;
          off=attrs[i];
          url=txt.getLinkUrl(off);
        }
        pOut=pOut.substring(0, off)+'['+pOut.substring(off, lastOff)+']('+url+')'+pOut.substring(lastOff);
      } else if (font) {
        if (!inSrc && font===font.COURIER_NEW) {
          while (i>=1 && txt.getFontFamily(attrs[i-1]) && txt.getFontFamily(attrs[i-1])===font.COURIER_NEW) {
            // detect fonts that are in multiple pieces because of errors on formatting:
            i-=1;
            off=attrs[i];
          }
          pOut=pOut.substring(0, off)+'`'+pOut.substring(off, lastOff)+'`'+pOut.substring(lastOff);
        }
      }
      if (txt.isBold(off)) {
        var d1 = "**";
        var d2 = "**";
        if (txt.isItalic(off)) {
          // edbacher: changed this to handle bold italic properly.
          d1 = "**_";
          d2 = "_**";
        }
        pOut=pOut.substring(0, off)+d1+pOut.substring(off, lastOff)+d2+pOut.substring(lastOff);
      } else if (txt.isItalic(off)) {
        pOut=pOut.substring(0, off)+'*'+pOut.substring(off, lastOff)+'*'+pOut.substring(lastOff);
      }
      lastOff=off;
    }
    return pOut;
  }
}
