/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict';

const Path = require('path');

const XlsxStreamReaderWorkBook = require(Path.join(__dirname, 'workbook'));

module.exports = XlsxStreamReader;

function XlsxStreamReader(){
	if (!(this instanceof XlsxStreamReader)) return new XlsxStreamReader();

	Object.defineProperty(this, 'options', {    
		value: {
			saxStrict: true,
			saxTrim: false,
			saxPosition: true,
			saxStrictEntities: true,
			saxNormalize: true
		},
		writable: true,
		enumerable: true
	});

	return new XlsxStreamReaderWorkBook(this.options);
}