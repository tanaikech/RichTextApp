# RichTextApp

<a name="top"></a>
[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENCE)

<a name="overview"></a>

# Overview

![](images/fig1.png)

**This is a GAS library for copying the rich text with the text styles from Google Document to Google Spreadsheet or from Google Spreadsheet to Google Document using Google Apps Script (GAS).**

<a name="description"></a>

## Description

Google Spreadsheet can use the rich text as the cell value. But I thought that it is difficult for me to directly edit the rich text in a cell. So I wanted to copy the rich text, that I edited at the Google Document, to the cell of Google Spreadsheet. But, unfortunately, when the rich text in Google Document is manually copied to a cell in Google Spreadsheet, the text style is removed. By this, only text values are copied. It seemed that this was the current specification. So in order to achieve above, I created this as a library.

This library can do as follows.

1. Rich text in Google Document can be copied to the cell of Google Spreadsheet with keeping the text style.

2. Rich text in the cell of Google Spreadsheet can be copied to Google Document with keeping the text style.

## Available text styles

In the current stage, the rich text of Google Spreadsheet can use the following text styles. [Ref1](https://developers.google.com/apps-script/reference/spreadsheet/text-style-builder), [Ref2](https://developers.google.com/apps-script/reference/spreadsheet/text-style)

- Font family of the text
- Font size of the text in points
- Font color of the text
- bold
- italic
- strikethrough
- underlined

# Library's project key

```
1Ka6RcIG6G9P8AhkJtjy1DAnHk1_rShFPWtdCQ2bxwMsrRE8cfn0cDXBv
```

# Methods

| Methods                                                               | Description                                     |
| :-------------------------------------------------------------------- | :---------------------------------------------- |
| [DocumentToSpreadsheet(object)](#documenttospreadsheet)               | Copy rich text from Document to Spreadsheet     |
| [SpreadsheetToDocument(object)](#spreadsheettodocument)               | Copy rich text from Spreadsheet to Document     |
| [RichTextToHTMLForSpreadsheet(object)](#richtexttohtmlforspreadsheet) | Convert rich texts in the cells to HTML format. |

<a name="usage"></a>

# Usage:

## 1. Install library

In order to use this library, please install this library as follows.

1. Create a GAS project.

   - You can use this library for the GAS project of both the standalone type and the container-bound script type.

1. [Install this library](https://developers.google.com/apps-script/guides/libraries).

   - Library's project key is **`1Ka6RcIG6G9P8AhkJtjy1DAnHk1_rShFPWtdCQ2bxwMsrRE8cfn0cDXBv`**.

### IMPORTANT

**This library uses V8 runtime. So please enable V8 at the script editor.**

### About scopes

This library use the scope of `https://www.googleapis.com/auth/spreadsheets`.

But in the following sample script, `https://www.googleapis.com/auth/documents` is also used.

## Methods

<a name="documenttospreadsheet"></a>

### `DocumentToSpreadsheet`

In this method, the text in Google Document can be put to the cell of Google Spreadsheet with keeping the text style.

#### Sample script

```javascript
function DocumentToSpreadsheet() {
  // From
  var doc = DocumentApp.openById("###");

  // To
  var ss = SpreadsheetApp.openById("###");
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange("A1");

  var res = RichTextApp.DocumentToSpreadsheet({ range: range, document: doc });
  console.log(res);
}
```

In this sample script, the text of Document is put to the cell "A1" of the 1st sheet in the Spreadsheet as the rich text.

<a name="spreadsheettodocument"></a>

### `SpreadsheetToDocument`

In this method, the text in the cell of Google Spreadsheet can be put to the Google Document with keeping the text style.

#### Sample script

```javascript
function SpreadsheetToDocument() {
  // From
  var ss = SpreadsheetApp.openById("###");
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange("A1");

  // To
  var doc = DocumentApp.openById("###");

  var res = RichTextApp.SpreadsheetToDocument({ range: range, document: doc });
  console.log(res);
}
```

In this sample script, the text of the cell "A1" of the 1st sheet in the Spreadsheet is put to the Document with the text style with the append method.

<a name="richtexttohtmlforspreadsheet"></a>

### `RichTextToHTMLForSpreadsheet`

In this method, the rich texts in the cells on Google Spreadsheet are converted to the HTML format.

#### Sample script

```javascript
function RichTextToHTMLForSpreadsheet() {
  var ss = SpreadsheetApp.openById("###");
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange("A1:A2");
  var res = RichTextApp_test.RichTextToHTMLForSpreadsheet({ range: range });
  console.log(res);
}
```

- When above demonstration image is converted to HTML format using this method, it can obtain this result. [https://jsfiddle.net/7e3mc10p/](https://jsfiddle.net/7e3mc10p/)

- When the range is only one cell, the string value of the HTML format is returned.
- When the range is the multiple cells, the 2 dimensional array including the HTML format is returned.

## Limitations

As the limitation, in the current stage, the table, list and images cannot be used with the rich text of Google Spreadsheet. So please use only the texts with the text style.

---

<a name="licence"></a>

# Licence

[MIT](LICENCE)

<a name="author"></a>

# Author

[Tanaike](https://tanaikech.github.io/about/)

If you have any questions and commissions for me, feel free to tell me.

<a name="updatehistory"></a>

# Update History

- v1.0.0 (February 19, 2020)

  1. Initial release.

- v1.1.0 (June 16, 2020)

  1. Add new method of `RichTextToHTMLForSpreadsheet`. The method of `RichTextToHTMLForSpreadsheet` can convert the rich texts in the cells to the HTML format.

[TOP](#top)
