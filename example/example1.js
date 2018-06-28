/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 *
 * example1
 *
 */

'use strict'

const fs = require('fs')
const XlsxStreamReader = require('../')

var workBookReader = new XlsxStreamReader()
workBookReader.on('error', function (error) {
  throw (error)
})

workBookReader.on('worksheet', function (workSheetReader) {
  if (workSheetReader.id > 1) {
    // we only want first sheet
    console.log('Skip Worksheet:', workSheetReader.id)
    workSheetReader.skip()
    return
  }
  console.log('Worksheet:', workSheetReader.id)

  workSheetReader.on('row', function (row) {
    row.values.forEach(function (rowVal, colNum) {
      console.log('RowNum', row.attributes.r, 'colNum', colNum, 'rowValLen', rowVal.length, 'rowVal', "'" + rowVal + "'")
    })
  })

  workSheetReader.on('end', function () {
    console.log('Worksheet', workSheetReader.id, 'rowCount:', workSheetReader.rowCount)
  })

  // call process after registering handlers
  workSheetReader.process()
})
workBookReader.on('end', function () {
  console.log('finished!')
})

fs.createReadStream('example/example1.xlsx').pipe(workBookReader)
