/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict';

const Path = require('path');
const Fs = require('fs');
const Temp = require('temp');
const Stream = require('stream');
const Unzip2 = require('unzip2');
const Sax = require('sax');
const Util = require('util');
const Zlib = require('zlib');

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
		'workBookInfo': {
			value: {
				sheetRelationships: {},
				sheetRelationshipsNames: {}
			},
			writable: true,
			enumerable: false
		},
		'parsedWorkBookInfo': {
			value: false,
			writable: true,
			enumerable: false
		},
		'parsedWorkBookRels': {
			value: false,
			writable: true,
			enumerable: false
		},
		'parsedSharedStrings': {
			value: false,
			writable: true,
			enumerable: false
		},
		'waitSharedStrings': {
			value: false,
			writable: true,
			enumerable: false
		},
		'waitingWorkSheets': {
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
		},
		'abortBook': {
			value: false,
			writable: true
		}
	});

	Temp.track();

	self._handleWorkBookStream();
}
Util.inherits(XlsxStreamReaderWorkBook, Stream);

XlsxStreamReaderWorkBook.prototype._handleWorkBookStream = function(){
	var self = this;
	var match;

	self.on('pipe', function (srcPipe) {
    		srcPipe.pipe(Unzip2.Parse())
    		.on('error', function(err){
			self.emit('error', err);
    		})
		.on('entry', function (entry) {
			if (self.abortBook){
				entry.autodrain();
				return;
			}
			switch (entry.path){
				case "xl/workbook.xml":
					self._parseSax(entry, self._parseWorkBookInfo, function(){
						self.parsedWorkBookInfo = true;
						self.emit('workBookInfo');
					});
					break;
				case "xl/_rels/workbook.xml.rels":
					self._parseSax(entry, self._parseWorkBookRels, function(){
						self.parsedWorkBookRels = true;
						self.emit('workBookRels');
					});
					break;
				case "_rels/.rels":
					entry.autodrain();
					break;
				case "xl/sharedStrings.xml":
					self._parseSax(entry, self._parseSharedStrings, function(){
						self.parsedSharedStrings = true
						self.emit('sharedStrings');
					});
					break;
				case "xl/styles.xml":
					self._parseSax(entry, self._parseStyles, function(){
						 self.emit('styles');
					});
					break;
				default:
					if (match = entry.path.match(/xl\/(worksheets\/sheet(\d+)\.xml)/)) {
						var sheetPath = match[1];
						var sheetNo = match[2];

						if (self.parsedWorkBookInfo == false
							|| self.parsedWorkBookRels == false
							|| self.parsedSharedStrings == false
							|| self.waitingWorkSheets.length > 0){

							var stream = Temp.createWriteStream();

							self.waitingWorkSheets.push({sheetNo: sheetNo, name: entry.path, path: stream.path, sheetPath: sheetPath})

							entry.pipe(stream);

						}else{
							var name = self._getSheetName(sheetPath);
							var workSheet = new XlsxStreamReaderWorkSheet(self, name, sheetNo, entry);

							self.emit('worksheet',workSheet);
						}
					} else if (match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)\.xml.rels/)) {
						var sheetNo = match[1];
						entry.autodrain();
					} else {
						entry.autodrain();
					}
					break;
			}
		})
		.on('close', function (entry) {
			if (self.waitingWorkSheets.length > 0){
				var currentBook = 0;
				var processBooks = function(){
					var sheetInfo = self.waitingWorkSheets[currentBook];
					var entry = Fs.createReadStream(sheetInfo.path);
					var name = self._getSheetName(sheetInfo.sheetPath);
					var workSheet = new XlsxStreamReaderWorkSheet(self, name, sheetInfo.sheetNo, entry);

					workSheet.on('end', function(node) {
						++currentBook;
						if (currentBook == self.waitingWorkSheets.length){
							Temp.cleanupSync();
							setImmediate(self.emit.bind(self),'end');
						}else{
							setImmediate(processBooks);
						}
					})

					setImmediate(self.emit.bind(self),'worksheet',workSheet);
				}
				setImmediate(processBooks);
			}else{
				setImmediate(self.emit.bind(self),'end');
			}
		});
	});
};

XlsxStreamReaderWorkBook.prototype.abort = function(){
	var self = this;

	self.abortBook = true;
};

XlsxStreamReaderWorkBook.prototype._parseSax = function(entryStream, entryHandler, endHandler){
	var self = this;

	var isErred = false;

	var tmpNode = [];
	var tmpNodeEmit = false;

	var saxOptions = {
		trim: self.options.saxTrim,
		position: self.options.saxPosition,
		strictEntities: self.options.saxStrictEntities
	};

	var saxStream = Sax.createStream(self.options.saxStrict, saxOptions);

	entryStream.on('end', function(node) {
		if (self.abortBook) return;
		if (!isErred) setImmediate(endHandler);
	});

	saxStream.on('error', function (error) {
		if (self.abortBook) return;
		isErred = true;

		self.emit('error',error);
	});

	saxStream.on('opentag', function(node) {
		if (self.abortBook) return;
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
		if (self.abortBook) return;
		tmpNodeEmit = true;
		tmpNode.push(text);
	});

	saxStream.on('closetag', function (nodeName) {
		if (self.abortBook) return;
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
};

XlsxStreamReaderWorkBook.prototype._getSharedString = function(stringIndex){
	var self = this;

	if (stringIndex > self.workBookSharedStrings.length){
		self.emit('error',"missing shared string: " + stringIndex);
		return;
	}
	return self.workBookSharedStrings[stringIndex];
};

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
};

XlsxStreamReaderWorkBook.prototype._parseStyles = function(nodeData){
	var self = this;

	nodeData.forEach(function(data){
		self.workBookStyles.push(data);
	});
};

XlsxStreamReaderWorkBook.prototype._parseWorkBookInfo = function(nodeData){
	var self = this;

	nodeData.forEach(function(data){
		if (data.name == 'sheet'){
			self.workBookInfo.sheetRelationshipsNames[data.attributes['r:id']] = data.attributes.name;
		}
	});
};

XlsxStreamReaderWorkBook.prototype._parseWorkBookRels = function(nodeData){
	var self = this;

	nodeData.forEach(function(data){
		if (data.name == 'Relationship'){
			self.workBookInfo.sheetRelationships[data.attributes.Target] = data.attributes.Id;
		}
	});
};

XlsxStreamReaderWorkBook.prototype._getSheetName = function(sheetPath){
	var self = this;

	return self.workBookInfo.sheetRelationshipsNames[
		self.workBookInfo.sheetRelationships[sheetPath]
	];
};
