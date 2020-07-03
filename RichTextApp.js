/**
 * GitHub  https://github.com/tanaikech/RichTextApp<br>
 * Copy rich text in Document to a cell of Spreadsheet.<br>
 * @param {object} Object object
 * @return {string} Return copied text as a string value
 */
function DocumentToSpreadsheet(object) {
    return new RichTextApp(object).DocumentToSpreadsheet();
}

/**
 * Copy rich text in a cell of Spreadsheet to Document.<br>
 * @param {object} Object object
 * @return {string} Return copied text as a string value
 */
function SpreadsheetToDocument(object) {
    return new RichTextApp(object).SpreadsheetToDocument();
}

/**
 * Convert rich text in a cell to HTML format.<br>
 * @param {object} Object object
 * @return {string} Return HTML data as a string value
 */
function RichTextToHTMLForSpreadsheet(object) {
    return new RichTextApp(object).RichTextToHTMLForSpreadsheet();
}

/**
 * Auto resize font in a cell of Spreadsheet.<br>
 * @param {object} Object object
 * @return {object} Range source range
 */
function AutoResizeFontForSpreadsheet(object) {
    return new RichTextApp(object).AutoResizeFontForSpreadsheet();
}
;
(function(r) {
  var RichTextApp;
  RichTextApp = (function() {
    var autoResizeFontForSpreadsheet, convertRichTextToHTMLForSpreadsheet, convertRichTextToHTMLForSpreadsheetSingle, getRichTextFromDocument, getRichTextFromSpreadsheet, putRichTextToDocument, putRichTextToSpreadsheet, putTextStyleToObj;

    class RichTextApp {
      constructor(obj_) {
        if (!("range" in obj_)) {
          throw new Error("Set 'range' object.");
        }
        this.obj = obj_;
      }

      // --- methods --- begin
      DocumentToSpreadsheet() {
        var data;
        if (!(this.obj.hasOwnProperty("document"))) {
          throw new Error("Set 'document' object.");
        }
        data = getRichTextFromDocument.call(this);
        return putRichTextToSpreadsheet.call(this, data);
      }

      SpreadsheetToDocument() {
        var data, text;
        if (!(this.obj.hasOwnProperty("document"))) {
          throw new Error("Set 'document' object.");
        }
        [data, text] = getRichTextFromSpreadsheet.call(this);
        putRichTextToDocument.call(this, data);
        return text;
      }

      RichTextToHTMLForSpreadsheet() {
        return convertRichTextToHTMLForSpreadsheet.call(this);
      }

      AutoResizeFontForSpreadsheet() {
        return autoResizeFontForSpreadsheet.call(this);
      }

    };

    RichTextApp.name = "RichTextApp";

    // --- methods --- end
    autoResizeFontForSpreadsheet = function() {
      var c, columnObj, endColumn, endRow, fontSize, j, k, l, range, ratio, ref, ref1, ref2, ref3, ref4, resizedWidth, sheet, srcRange, ss, startColumn, startRow, tempRange, tempSheet, toLarge;
      if (!this.obj.hasOwnProperty("range")) {
        throw new Error("Error: Wrong range.");
      }
      range = this.obj.range;
      toLarge = this.obj.hasOwnProperty("toLarge") ? this.obj.toLarge : false;
      sheet = range.getSheet();
      ss = sheet.getParent();
      startColumn = range.getColumn();
      endColumn = range.getColumn() + range.getNumColumns();
      startRow = range.getRow();
      endRow = range.getRow() + range.getNumRows();
      columnObj = [];
      for (c = j = ref = startColumn, ref1 = endColumn; (ref <= ref1 ? j < ref1 : j > ref1); c = ref <= ref1 ? ++j : --j) {
        columnObj.push({
          column: c,
          width: sheet.getColumnWidth(c)
        });
      }
      tempSheet = ss.insertSheet("tempForAutoresizeFont");
      sheet.activate();
      tempRange = tempSheet.getRange("A1");
      for (r = k = ref2 = startRow, ref3 = endRow; (ref2 <= ref3 ? k < ref3 : k > ref3); r = ref2 <= ref3 ? ++k : --k) {
        for (c = l = 0, ref4 = columnObj.length; (0 <= ref4 ? l < ref4 : l > ref4); c = 0 <= ref4 ? ++l : --l) {
          srcRange = sheet.getRange(r, columnObj[c].column);
          tempSheet.setColumnWidth(1, columnObj[c].width);
          srcRange.copyTo(tempRange);
          tempSheet.autoResizeColumn(1);
          resizedWidth = tempSheet.getColumnWidth(1);
          tempSheet.setColumnWidth(1, columnObj[c].width);
          ratio = columnObj[c].width / resizedWidth;
          if (ratio > 1 && !toLarge) {
            continue;
          }
          fontSize = srcRange.getFontSize();
          srcRange.setFontSize(Math.ceil(fontSize * ratio));
        }
      }
      ss.deleteSheet(tempSheet);
      return range;
    };

    convertRichTextToHTMLForSpreadsheet = function() {
      var htmls, rt;
      rt = this.obj.range.getRichTextValues();
      htmls = rt.map((row) => {
        return row.map((col) => {
          return convertRichTextToHTMLForSpreadsheetSingle.call(this, col);
        });
      });
      if (htmls.length === 0) {
        throw new Error("Error: Wrong range.");
      } else if (htmls.length === 1 && htmls[0].length === 1) {
        return htmls[0][0];
      }
      return htmls;
    };

    convertRichTextToHTMLForSpreadsheetSingle = function(richTextValue) {
      return richTextValue.getRuns().reduce((s, r) => {
        var bold, fontFamily, fontSize, foregroundColor, italic, keys, obj, strikethrough, style, text, underline;
        text = r.getText().replace(/\n/g, "<br>").replace(/ /g, "&nbsp;");
        if (r.getLinkUrl()) {
          text = `<a href="${r.getLinkUrl()}">${text}<\/a>`;
        }
        style = r.getTextStyle();
        obj = {
          fontFamily: style.getFontFamily(),
          fontSize: style.getFontSize(),
          foregroundColor: style.getForegroundColor(),
          bold: style.isBold(),
          italic: style.isItalic(),
          strikethrough: style.isStrikethrough(),
          underline: style.isUnderline()
        };
        fontFamily = obj.fontFamily ? `font-family: '${obj.fontFamily}';` : "";
        fontSize = obj.fontSize ? `font-size: ${obj.fontSize * 1.333}px;` : "";
        foregroundColor = obj.foregroundColor ? `color: ${obj.foregroundColor};` : "";
        bold = obj.bold ? 'font-weight: bold;' : "";
        italic = obj.italic ? 'font-style: italic;' : "";
        strikethrough = obj.strikethrough ? 'text-decoration: line-through;' : "";
        underline = obj.underline ? 'text-decoration: underline;' : "";
        keys = [fontFamily, fontSize, foregroundColor, bold, italic, strikethrough, underline];
        if (keys.some((e) => {
          return e !== "";
        })) {
          s += `${keys.reduce((str, e) => {
            return str += e;
          }, '<span style="')}">${text}</span>`;
        } else {
          s += text;
        }
        return s;
      }, "");
    };

    putTextStyleToObj = function(c, style) {
      return {
        text: c.toString(),
        foregroundColor: style.getForegroundColor(),
        fontFamily: style.getFontFamily(),
        fontSize: style.getFontSize(),
        bold: style.isBold(),
        italic: style.isItalic(),
        strikethrough: style.isStrikethrough(),
        underline: style.isUnderline()
      };
    };

    getRichTextFromSpreadsheet = function() {
      var data, rt, temp, textData;
      rt = this.obj.range.getRichTextValue();
      textData = rt.getText();
      temp = [];
      data = Array.prototype.reduce.call(textData, (ar, c, offset) => {
        var end, style;
        end = offset + 1;
        style = rt.getTextStyle(offset, end);
        if (c.toString() === "\n" || offset === textData.length - 1) {
          if (c.toString() !== "\n" && offset === textData.length - 1) {
            temp.push(putTextStyleToObj.call(this, c, style));
          }
          ar.push(temp);
          temp = [];
        } else {
          temp.push(putTextStyleToObj.call(this, c, style));
        }
        return ar;
      }, []);
      return [data, textData];
    };

    putRichTextToDocument = function(data) {
      var body;
      body = this.obj.document.getBody();
      data.forEach((p) => {
        var para, text;
        para = body.appendParagraph("");
        text = para.editAsText();
        return p.forEach((e, i) => {
          return text.appendText(e.text).setForegroundColor(i, i, e.foregroundColor).setFontFamily(i, i, e.fontFamily).setFontSize(i, i, e.fontSize).setBold(i, i, e.bold).setItalic(i, i, e.italic).setStrikethrough(i, i, e.strikethrough).setUnderline(i, i, e.underline);
        });
      });
      return null;
    };

    getRichTextFromDocument = function(data) {
      var body, paragraphs;
      body = this.obj.document.getBody();
      paragraphs = body.getParagraphs();
      return paragraphs.reduce((ar, e) => {
        var styles, temp, text, textData;
        text = e.editAsText();
        textData = text.getText();
        styles = Array.prototype.map.call(textData, (_, offset) => {
          return {
            foregroundColor: text.getForegroundColor(offset) || "#000000",
            fontFamily: text.getFontFamily(offset) || "Arial",
            fontSize: text.getFontSize(offset) || 11,
            bold: text.isBold(offset) || false,
            italic: text.isItalic(offset) || false,
            strikethrough: text.isStrikethrough(offset) || false,
            underline: text.isUnderline(offset) || false
          };
        });
        if (text !== "" && styles.length > 0) {
          temp = {
            text: textData,
            styles: styles
          };
          ar.push(temp);
        }
        return ar;
      }, []);
    };

    putRichTextToSpreadsheet = function(data) {
      var end, richText, start, texts;
      texts = (data.map((e) => {
        return e.text;
      })).join("\n");
      richText = SpreadsheetApp.newRichTextValue().setText(texts);
      start = 0;
      end = 0;
      data.forEach((e) => {
        e.styles.forEach((f) => {
          var style;
          end = start + 1;
          style = SpreadsheetApp.newTextStyle().setBold(f.bold).setFontFamily(f.fontFamily).setFontSize(f.fontSize).setForegroundColor(f.foregroundColor).setItalic(f.italic).setStrikethrough(f.strikethrough).setUnderline(f.underline).build();
          richText.setTextStyle(start, end, style);
          return start += 1;
        });
        return start = end + 1;
      });
      this.obj.range.setRichTextValue(richText.build());
      return texts;
    };

    return RichTextApp;

  }).call(this);
  return r.RichTextApp = RichTextApp;
})(this);
