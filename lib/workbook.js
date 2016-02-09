/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict';

const Sax = require('sax');
const Path = require('path');
const Util = require('util');
const Unzip2 = require('unzip2');
const Stream = require('stream');

const XlsxStreamReaderWorkSheet = require(Path.join(__dirname, 'worksheet'));

module.exports = XlsxStreamReaderWorkBook;

function XlsxStreamReaderWorkBook(options){
	var self = this;

	if (!(this instanceof XlsxStreamReaderWorkBook)) return new XlsxStreamReaderWorkBook(options);

	Object.defineProperties(this, { 
		'options': {
			value: options,
			writable: true,
			enumerable: true
		},
		'write': {
			value: function(){ return; },
		},
		'end': {
			value: function(){ return; },
		},		
		'workBookSharedStrings': {
			value: [],
			writable: true,
			enumerable: false
		},
		'workBookStyles': {
			value: [],
			writable: true,
			enumerable: false
		},
		'pipeMode': {
			value: false,
			writable: true,
			enumerable: false
		}
	});

	self._handleWorkBookStream();
}
Util.inherits(XlsxStreamReaderWorkBook, Stream);

XlsxStreamReaderWorkBook.prototype._handleWorkBookStream = function(){
	var self = this;

	self.on('pipe', function (srcPipe) {
		srcPipe.pipe(Unzip2.Parse())
		.on('entry', function (entry) {
			switch (entry.path){
				case "_rels/.rels":
				case "xl/workbook.xml":
				case "xl/_rels/workbook.xml.rels":
					entry.autodrain();
					break;
				case "xl/sharedStrings.xml":
					self._parseSax(entry, self._parseSharedStrings, function(){
						 self.emit('sharedStrings');
					});
					break;
				case "xl/styles.xml":
					self._parseSax(entry, self._parseStyles, function(){
						 self.emit('styles');
					});
					break;
				default:
					if (entry.path.match(/xl\/worksheets\/sheet\d+\.xml/)) {
						var match = entry.path.match(/xl\/worksheets\/sheet(\d+)\.xml/)
						var sheetNo = match[1];

						var workSheet = new XlsxStreamReaderWorkSheet(self, sheetNo, entry);

						self.emit('worksheet',workSheet);
					} else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+\.xml.rels/)) {
						var match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)\.xml.rels/)
						var sheetNo = match[1];
						console.log("_parseHyperlinks",sheetNo)
					} else {
						entry.autodrain();
					}
					break;
			}
		})
		.on('close', function (entry) {
			self.emit('end');
		});
	});
}

XlsxStreamReaderWorkBook.prototype._parseSax = function(entryStream, entryHandler, endHandler){
	var self = this;

	var isErred = false;

	var tmpNode = []
	var tmpNodeEmit = false;

	var saxOptions = {
		trim: self.options.saxTrim,
		position: self.options.saxPosition,
		strictEntities: self.options.saxStrictEntities
	}
	
	var saxStream = Sax.createStream(self.options.saxStrict, saxOptions)
	
	entryStream.on('end', function(node) {
		if (!isErred) setImmediate(endHandler);
	});

	saxStream.on('error', function (error) {
		isErred = true;

		self.emit('error',error);
	});

	saxStream.on('opentag', function(node) {
		if (Object.keys(node.attributes).length == 0){
			delete(node.attributes);
		}
		if (node.isSelfClosing){
			if (tmpNode.length > 0){
				entryHandler.call(self, tmpNode);
				tmpNode = [];
			}
			tmpNodeEmit = true;
		}
		delete(node.isSelfClosing);
		tmpNode.push(node);
	});

	saxStream.on('text', function (text) {
		tmpNodeEmit = true;
		tmpNode.push(text);
	});

	saxStream.on('closetag', function (nodeName) {
		if (tmpNodeEmit){
			entryHandler.call(self, tmpNode);
			tmpNodeEmit = false;
			tmpNode = [];
		}
		tmpNode.splice(-1,1);
	});

	try{
		entryStream.pipe(saxStream);
	}catch(error){
		self.emit('error',error);
	}
}

XlsxStreamReaderWorkBook.prototype._getSharedString = function(stringIndex){
	var self = this;

	if (stringIndex > self.workBookSharedStrings.length){
		console.log("WARNING: missing shared string:",stringIndex, self.workBookSharedStrings.length)
	}
	return self.workBookSharedStrings[stringIndex];
}

XlsxStreamReaderWorkBook.prototype._parseSharedStrings = function(nodeData){
	var self = this;

	var nodeObjValue = nodeData.pop();
	var nodeObjName = nodeData.pop();

	if (nodeObjName && nodeObjName.name == 't'){
		self.workBookSharedStrings.push(nodeObjValue);
	}else{
		if (nodeObjValue && typeof nodeObjValue == 'object' && nodeObjValue.hasOwnProperty('name') && nodeObjValue.name == 't'){
			self.workBookSharedStrings.push("");
		}
	}
}

XlsxStreamReaderWorkBook.prototype._parseStyles = function(nodeData){
	var self = this;

	nodeData.forEach(function(data){
		self.workBookStyles.push(data);
	});
}