/* global describe, it */

const XlsxStreamReader = require('../index')
const fs = require('fs')
const assert = require('assert')
const path = require('path')

describe('The xslx stream parser', function () {
  it('parses large files', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'big.xlsx')).pipe(workBookReader)
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(workSheetReader.rowCount === 80000)
        done()
      })
      workSheetReader.process()
    })
  })
  it.only('parses dates', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'import.xlsx')).pipe(workBookReader)
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(workSheetReader.rowCount === 80000)
        done()
      })
      workSheetReader.on('row', function (r) {
        console.log(r)
      })
      workSheetReader.process()
    })
  })
  it('catches zip format errors', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'notanxlsx')).pipe(workBookReader)
    workBookReader.on('error', function (err) {
      assert(err.message === 'invalid signature: 0x6d612069')
      done()
    })
  })
})
