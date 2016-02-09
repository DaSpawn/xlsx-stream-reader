# xlsx-stream-reader
======

Memory efficinet minimalist streaming XLSX reader that can handle piped 
streams as input. Events are emmited while reading the stream.


Example
-------
```javascript
const XlsxStreamReader = require("xlsx-stream-reader");

var workBookReader = new XlsxStreamReader();
workBookReader.on('error', function (error) {
    throw(error);
});
workBookReader.on('sharedStrings', function () {
    // do not need to do anything with these, 
    // cached and used when processing worksheets
    console.log(workBookReader.workBookSharedStrings);
});

workBookReader.on('styles', function () {
    // do not need to do anything with these
    // but not currently handled in any other way
    console.log(workBookReader.workBookStyles);
});

workBookReader.on('worksheet', function (workSheetReader) {
    if (workSheetReader.id > 1){
        // we only want first sheet
        workSheetReader.skip();
        return; 
    }
    // if we do not listen for rows we will only get end event
    // and have infor about the sheet like row count
    workSheetReader.on('row', function (row) {
        if (row.attributes.r == 1){
            // do something with row 1 like save as column names
        }else{
            // do something with remaining rows
        }
    });
    workSheetReader.on('end', function () {
        console.log(workSheetReader.rowCount);
    });

    // call process after registering handlers
    workSheetReader.process();
});
workBookReader.on('end', function () {
    // end of workbook reached
});

fs.createReadStream(fileName).pipe(workBookReader);

```


Beta Warning
-------
This module is currently in use on a live internal business system for product 
management. That being said this should still be considered beta. More usage 
and input from users will be needed due to the numerous differences/incompatibilities/flukes 
I have already run into with XLSX files.


Limitations
-------
The row reader currently returns formulas and not stored values (if available). 
As time permits the row handler will be more capable but was enough for currrent 
purposes (loading values from large worksheets fast)


Inspiration
-----------
Need a simple XLSX file streaming reader to handle large excel sheets but only 
one available/compatible was by guyonroche/exceljs. The stream reader module at 
the time was unfinished/unusable and rewrite attempts exposed column shifting I 
could not solve


More Information
-----------
Events are emmited as pertinent parts of the workbook and worksheet are receieved 
in the stream. Theoretically you could pause the input stream if events are being 
receieved too fast but this has not been tested

Events can potentially (even though I have not seen it) be receieved out of order, 
if you receive a worksheet end event while still receieving rows be sure to make sure 
your number of rows receieved equals the `workSheetReader.rowCount` 

Theoretically you could process an excel sheet as it is being uploaded, depending 
on the sheet type, but untried (I encountered some XLSX files that have a different 
zip format that requires having the entire file to read the archive contents properly), 
but still probably better to save temp first and read streasm from there.

Currently if the zip archive does not have the shared strings at the begining of the 
archive an error will be thrown.  This will be solved by piping the input stream for 
each sheet into a temp file untill the shared string are encountered and processed, 
then re-read the temp worksheets processing in the shared strings.  This allows me 
to keep the input as piped streams, and still handle the situation with xlsx archives out 
of order.  Since the sheets are normally much smaller than the entire archive/shared 
strings table, the temp files should remain small and compressed, but all of this 
will be seamless/automatic


API Information
-----------
#### new XlsxStreamReader()

Create a new XlsxStreamReader object (workBookReader). After attaching handlers you 
can `pipe()` your input stream into the reader to begin processing

#### Event: 'error'

* {Error Object}

Emitted if there was an error in processing (may not catch all errors, 
some may be thrown depending on where the error happened)

#### Event: 'end'

Emmitted once the XLSX zip parser has closed and all sheets have been processed

#### Event: 'sharedStrings'

After the workbook shared strings have been parsed this event is emmited. Shared strings 
are available via {Array} `workBookReader.workBookSharedStrings`.  

#### Event: 'styles'

After the workbook styles have been parsed this event is emmited. Styles are available 
via {Array} `workBookReader.workBookStyles`

#### Event: 'worksheet'

* `options` {Object} XlsxStreamReaderWorkSheet object (workSheetReader)

Emmitted when a worksheet is reached. The sheet number is availble via 
{Number} `workSheetReader.id`. You can either process or skip at this point, 
but you must do one for the processing to the next sheet to continue/finish.

Once event is recieved you can attach worksheet on handlers (end, row) then you 
would `workSheetReader.process()`. If you do not want to process a sheet and instead 
want to skip entirely, you would `workSheetReader.skip()` without attaching any handlers.

#### Worksheet Event: 'end'

Emmitted once the end of the worksheet has been reached. The row count is 
available via {Number} `workSheetReader.rowCount`

#### Worksheet Event: 'row'

* `options` {Object} Row object

Emmitted on every row encountered in the worksheet. for more details on what 
is in the row object, see the [Row class][msdnRows] on MSDN.


References
-----------
* [Working with sheets (Open XML SDK)][msdnSheets]
* [Row class][msdnRows]
* [ExcelJS][ExcelJS]


Used Modules
-----------
* [Path][modPath]
* [Util][modUtil]
* [Stream][modStream]
* [Sax][modSax]
* [Unzip2][modUnzip2]


Authors
-----------
Written by [Brian Taber](https://github.com/DaSpawn) 

[![DaSpawn's Gratipay][gratipay-image-daspawn]][gratipay-url-daspawn]


License
-----------
[MIT](LICENSE)



[gratipay-url-daspawn]: https://gratipay.com/~DaSpawn
[gratipay-image-daspawn]: https://img.shields.io/gratipay/team/daspawn.svg
[msdnRows]: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.row.aspx
[msdnSheets]: https://msdn.microsoft.com/EN-US/library/office/gg278309.aspx
[ExcelJS]: https://github.com/guyonroche/exceljs

[modPath]: https://nodejs.org/api/path.html
[modStream]: https://nodejs.org/api/stream.html
[modUtil]: https://nodejs.org/api/util.html
[modSax]: https://github.com/isaacs/sax-js
[modUnzip2]: https://github.com/glebdmitriew/node-unzip-2