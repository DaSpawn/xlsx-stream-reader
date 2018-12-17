/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict'

const Path = require('path')

const XlsxStreamReaderWorkBook = require(Path.join(__dirname, 'workbook'))

module.exports = XlsxStreamReader

function XlsxStreamReader (userOptions) {
  if (!(this instanceof XlsxStreamReader)) return new XlsxStreamReader()
  const defaults = {
    verbose: true,
    formatting: true,
    returnFormats: false
  }
  const saxOptions = {
    saxStrict: true,
    saxTrim: true,
    saxNormalize: true,
    saxPosition: true,
    saxStrictEntities: true
  }
  const options = Object.assign(saxOptions, defaults, userOptions)
  Object.defineProperty(this, 'options', {
    value: options,
    writable: true,
    enumerable: true
  })

  return new XlsxStreamReaderWorkBook(this.options)
}
