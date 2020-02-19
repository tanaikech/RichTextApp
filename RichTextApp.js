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
;
(function(r) {
  var RichTextApp;
  RichTextApp = (function() {
    var getRichTextFromDocument, getRichTextFromSpreadsheet, putRichTextToDocument, putRichTextToSpreadsheet, putTextStyleToObj;

    class RichTextApp {
      constructor(obj_) {
        if (!("range" in obj_) || !("document" in obj_)) {
          throw new Error("Set 'range' object and 'document' object.");
        }
        this.obj = obj_;
      }

      DocumentToSpreadsheet() {
        var data;
        data = getRichTextFromDocument.call(this);
        return putRichTextToSpreadsheet.call(this, data);
      }

      SpreadsheetToDocument() {
        var data, text;
        [data, text] = getRichTextFromSpreadsheet.call(this);
        putRichTextToDocument.call(this, data);
        return text;
      }

    };

    RichTextApp.name = "RichTextApp";

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
