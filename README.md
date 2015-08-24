# node-spreadsheet-writer

[![npm Version](https://badge.fury.io/js/spreadsheet-writer.png)](https://npmjs.org/package/spreadsheet-writer)

A spreadsheet writer, for excel compatible formats. (node.js) 

Supprts:
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

for (var i = 0; i < 100000; i++) {
    writer.beginRow();
    writer.addCell('test string', SpreadsheetWriter.Types.CellType.String, 0);
    writer.addCell(new Date());
    writer.addCell(true);
}

writer.endFile();

writer.on('finish', function () {

    writableStream.end();

});

```


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



