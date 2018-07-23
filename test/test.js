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
  it('supports predefined formats', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'predefined_formats.xlsx')).pipe(workBookReader)
    const rows = []
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(rows[1][4] === '9/27/86')
        assert(rows[1][8] === '20064')
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
  })
  it('supports custom formats', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'import.xlsx')).pipe(workBookReader)
    const rows = []
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(rows[1][2] === '27/09/1986')
        assert(rows[1][3] === '20064')
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
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
