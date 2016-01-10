# spreadsheet-writer

[![npm Version](https://badge.fury.io/js/spreadsheet-writer.png)](https://npmjs.org/package/spreadsheet-writer)

A spreadsheet writer, for excel compatible formats. (node.js)

Documentation and examples are lacking, but the code is all documented with JSDocs.

Supports:

* CSV (through SpreadsheetWriter.CsvSpreadsheetWriter)
* Excel XML (through SpreadsheetWriter.XmlSpreadsheetWriter)
* Handling very large spreadsheets without any impact on memory (everything is written directly to the stream)
* Styling
* Multiple worksheets
* Formulas, Dates, Booleans etc.
* Planning future support for XLSX

## Usage (Note: A much better example is needed...)

```javascript
var fs = require('fs');
var SpreadsheetWriter = require('spreadsheet-writer');

var writableStream = fs.createWriteStream('mysheet.xml');

var writer = new SpreadsheetWriter.XmlSpreadsheetWriter(writableStream);
writer.setEncoding('utf8');

writer.addStyle({
    font: {
        color: 'red'
    }
});

writer.beginFile();

writer.newWorksheet('Sample Sheet');
writer.addColumn(200, true);
writer.addColumn(100, true);
writer.addColumn(50, true);
writer.addColumn(50, true);
writer.addColumn(50, true);
writer.addColumn(50, true);

for (var i = 0; i < 100000; i++) {
    writer.beginRow();
    writer.addCell('test string', SpreadsheetWriter.Types.CellType.String, 0);
    writer.addCell(new Date());
    writer.addCell(true);
    writer.addCell(Math.rand() * 10);
    writer.addCell(Math.rand() * 20);
    writer.addFormulaCell('=RC[-1]+RC[-2]', SpreadsheetWriter.Types.CellFormulaPlaceholder.Null, SpreadsheetWriter.Types.CellType.Number, 0);
}

writer.endFile();

writer.on('finish', function () {

    writableStream.end();

});

```

## Supported file formats

There are currently two writers supported:

* SpreadsheetWriter.XmlSpreadsheetWriter
* SpreadsheetWriter.CsvSpreadsheetWriter

## Common writer functions

* setEncoding: (encoding)
* getEncoding: ()
* mimeType: ()
* fileExtension: ()
* addStyle: (style)
* addColumn: (width, autoFitWidth)
* beginFile: ()
* newWorksheet: (worksheetName)
* beginRow: ()
* addCell: (data, type, styleIndex, mergeAcross, mergeDown)
* addFormulaCell: (formula, dataPlaceholder, type, styleIndex, mergeAcross, mergeDown)
* addRichTextCell: (xml, styleIndex, mergeAcross, mergeDown)
* endFile: ()

## Styles

Object `SpreadsheetCellStyle`:

| property           | type                |
|--------------------|---------------------|
| `parentStyleIndex` | `int?`              |
| `numberFormat`     | `String?`           |
| `alignment`        | `Alignment = None`  |
| `borders`          | `Array.<Border>?`   |
| `interior`         | `Interior?`         |
| `font`             | `Font?`             |

Object `Alignment`:

| property           | type                |
|--------------------|---------------------|
| `horizontal`       | `HorizontalAlignment = Automatic`  |
| `vertical`         | `VerticalAlignment = Automatic`    |
| `indent`           | `int = 0`                          |
| `readingOrder`     | `HorizontalReadingOrder = Context` |
| `rotate`           | `Number = 0`                       |
| `shrinkToFit`      | `boolean = false`                  |
| `verticalText`     | `boolean = false`                  |
| `wrapText`         | `boolean = false`                  |

Object `Interior`:

| property           | type                |
|--------------------|---------------------|
| `color`            | `String?`  |
| `pattern`          | `InteriorPattern = None`    |
| `patternColor`     | `String?`   |

Object `Font`:

| property           | type                |
|--------------------|---------------------|
| `bold`            | `boolean = false`  |
| `color`            | `String?`  |
| `fontName`            | `String?`  |
| `italic`            | `boolean = false`  |
| `outline`            | `boolean = false`  |
| `shadow`            | `boolean = false`  |
| `size`            | `Number = 0`  |
| `strikeThrough`            | `boolean = false`  |
| `underline`            | `FontUnderline = None`  |
| `verticalAlign`            | `FontVerticalAlign = None`  |
| `charset`            | `int = 0`  |
| `family`            | `FontFamily = Automatic`  |

Object `Border`:

| property           | type                |
|--------------------|---------------------|
| `position`            | `BorderPosition`  |
| `color`            | `String`  |
| `lineStyle`            | `BorderLineStyle = None`  |
| `weight`            | `Number = 0`  |

Enum `VerticalAlignment`:
  `Automatic`, `Top`, `Center`, `Bottom`, `Justify`, `Distributed`, `JustifyDistributed`

Enum `BorderLineStyle`:
    `None`, `Continuous`, `Dash`, `Dot`, `DashDot`, `DashDotDot`, `SlantDashDot`, `Double: 'Double'`

Enum `BorderPosition`:
    `Left`, `Top`, `Right`, `Bottom`, `DiagonalLeft`, `DiagonalRight`

Enum `FontFamily`:
    `Automatic`, `Decorative`, `Modern`, `Roman`, `Script`, `Swiss`

Enum `FontUnderline`:
    `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`

Enum `FontVerticalAlign`:
    `None`, `Subscript`, `Superscript`

Enum `HorizontalAlignment`:
    `Automatic`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterAcrossSelection`, `Distributed`, `JustifyDistributed`

Enum `HorizontalReadingOrder`:
    `Context`, `RightToLeft`, `LeftToRight`

Enum `InteriorPattern`:
    `None`, `Solid`, `Gray75`, `Gray50`, `Gray25`, `Gray125`, `Gray0625`, `HorzStripe`, `VertStripe`, `ReverseDiagStripe`, `DiagStripe`, `DiagCross`, `ThickDiagCross`, `ThinHorzStripe`, `ThinVertStripe`, `ThinReverseDiagStripe`, `ThinDiagStripe`, `ThinHorzCross`, `ThinDiagCross`

## Other types

Enum `NumberFormats`:
  `Automatic`, `General`, `GeneralNumber`, `GeneralDate`, `LongDate`, `MediumDate`, `ShortDate`, `LongTime`, `MediumTime`, `ShortTime`, `Currency`, `EuroCurrency`, `Fixed`, `Standard`, `Percent`, `Scientific`, `YesNo`, `TrueFalse`, `OnOff`, `Number0`, `Number0_00`
};

Enum `CellType`:
    `Number`, `DateTime`, `Boolean`, `String`, `Error`

Enum `CellFormulaPlaceholder`:
    `Null`, `DivisionByZero`, `InvalidValue`, `InvalidReference`, `InvalidNameReference`, `NotANumber`, `NotAvailable`, `CircularReference`

## Excel Number Formats

These are quite annoying to define. The specifications are not very clear for beginners, so I'll try to explain it shortly:

A `NumberFormat` consists of 3 sections: `POSITIVE;NEGATIVE;ZERO`.
Respectively, the first section defines the format for a positive number, the second defines the format for a negative value, and the 3rd is for a zero number.

Of course, they are not always all relevant. For date formats, for example, there are only positive values.

Now if we have omitted the negative format in this manner `#.00` then it will default to a `-POSITIVE`, in this case `-#.00`.
But if we omitted it in this manner `#.00;;\Z\e\r\o`, then a negative number will show as an empty value!

Notes about date formats:

`:` should display a localized time separator. Excel throws errors for those if used not between two time specifiers. To always show a `:` you need to escape it.
`/` should display a localized date separator. Excel throws errors for those if used not between two time specifiers. To always show a `\` you need to escape it.

Not all codes behave according to specs. `N`/`Nn` should represent minutes, but they throw errors. `m`/`mm` are used instead.
`m` or `mm` represent minutes only when excel detects that there was an `h`/`H`/`hh`/`HH` before it. Otherwise it will print months.

There's a utility function that takes a standard date format and generates a `NumberFormat` for Excel. You can call it like this:
```javascript
var myNumberFormat = SpreadsheetWriter.Utils.excelNumberFormatForDateFormat('dd/MM/yyyy", at "H:mm');
console.log(myNumberFormat); // Prints 'dd/mm/yyyy", at "h:mm;@'

writer.addStyle({
    numberFormat: myNumberFormat,
    borders: [
        { color: '#0000ff', position: SpreadsheetWriter.Types.BorderPosition.Top, weight: 2, lineStyle: SpreadsheetWriter.Types.BorderLineStyle.Dash },
        { color: '#0000ff', position: SpreadsheetWriter.Types.BorderPosition.Bottom, weight: 2, lineStyle: SpreadsheetWriter.Types.BorderLineStyle.Dash }
    ],
    font: {
        color: '#00ff00'
    }
});

```

Sources:
This is the closest to a spec that I managed to find: https://msdn.microsoft.com/en-us/library/office/gg251755.aspx

## Contributing

If you have anything to contribute, or functionality that you luck - you are more than welcome to participate in this!  
If anyone wishes to contribute unit tests - that also would be great :-)

## Me
* Hi! I am Daniel Cohen Gindi. Or in short- Daniel.
* danielgindi@gmail.com is my email address.
* That's all you need to know.

## Help

If you want to buy me a beer, you are very welcome to
[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=G6CELS3E997ZE)
 Thanks :-)

## License

All the code here is under MIT license. Which means you could do virtually anything with the code.
I will appreciate it very much if you keep an attribution where appropriate.

    The MIT License (MIT)

    Copyright (c) 2013 Daniel Cohen Gindi (danielgindi@gmail.com)

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.



