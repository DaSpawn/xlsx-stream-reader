const fs = require('fs')
const path = require('path')

const fileName = path.resolve(__dirname, 'example4.xlsx')

const XlsxStreamReader = require('../index');

var workBookReader = new XlsxStreamReader();
workBookReader.on('error', function (error) {
    throw(error);
});
workBookReader.on('sharedStrings', function () {
    // do not need to do anything with these, 
    // cached and used when processing worksheets
    console.log(workBookReader.workBookSharedStrings);
});

// workBookReader.on('styles', function () {
//     // do not need to do anything with these
//     // but not currently handled in any other way
//     console.log(workBookReader.workBookStyles);
// });

workBookReader.on('worksheet', function (workSheetReader) {
    if (workSheetReader.id > 1){
        // we only want first sheet
        workSheetReader.skip();
        return; 
    }
    // print worksheet name
    console.log(workSheetReader.name);

    // if we do not listen for rows we will only get end event
    // and have infor about the sheet like row count
    workSheetReader.on('row', function (row) {
        if (row.attributes.r == 1){
            // do something with row 1 like save as column names
        }else{
            // second param to forEach colNum is very important as
            // null columns are not defined in the array, ie sparse array
            row.values.forEach(function(rowVal, colNum){
                // do something with row values
            });
        }
    });
    // workSheetReader.on('end', function () {
    //     console.log(workSheetReader.rowCount);
    // });

    // call process after registering handlers
    workSheetReader.process();
});
// workBookReader.on('end', function () {
//     // end of workbook reached
// });

fs.createReadStream(fileName).pipe(workBookReader);