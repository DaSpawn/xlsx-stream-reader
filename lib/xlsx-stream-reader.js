/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict';

const Path = require('path');

const XlsxStreamReaderWorkBook = require(Path.join(__dirname, 'workbook'));

module.exports = XlsxStreamReader;

function XlsxStreamReader(options){
	if (!(this instanceof XlsxStreamReader)) return new XlsxStreamReader();

	if (!options || typeof options !== 'object') {
		options = {}
	}

	if (typeof options.verbose === 'undefined') {
		options.verbose = true
	}

	if (typeof options.formatting === 'undefined') {
		options.formatting = true
	}

	Object.defineProperty(this, 'options', {    
		value: {
			saxStrict: true,
			saxTrim: false,
			saxNormalize: true,
			saxPosition: true,
			saxStrictEntities: true,
			verbose: options.verbose,
			formatting: options.formatting
		},
		writable: true,
		enumerable: true
	});

	return new XlsxStreamReaderWorkBook(this.options);
}