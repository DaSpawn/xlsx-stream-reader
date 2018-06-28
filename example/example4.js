const fs = require('fs')
const path = require('path')

const fileName = path.resolve(__dirname, 'example4.xlsx')
// const fileName = path.resolve(__dirname, 'example5.xlsx')

const XlsxStreamReader = require('../index')

var workBookReader = new XlsxStreamReader({
  verbose: false,
  formatting: true
})

workBookReader.on('error', function (error) {
  throw (error)
})
workBookReader.on('sharedStrings', function () {
  // do not need to do anything with these,
  // cached and used when processing worksheets
  // console.log(workBookReader.workBookSharedStrings);
})

workBookReader.on('worksheet', function (workSheetReader) {
  // if (workSheetReader.id > 1){
  //     // we only want first sheet
  //     workSheetReader.skip();
  //     return;
  // }

  // if we do not listen for rows we will only get end event
  // and have infor about the sheet like row count
  workSheetReader.on('row', function (row) {
    console.log(row.values)
  })

  // call process after registering handlers
  workSheetReader.process()
})

fs.createReadStream(fileName).pipe(workBookReader)
