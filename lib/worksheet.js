/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict';

const Util = require('util');
const Stream = require('stream');

module.exports = XlsxStreamReaderWorkSheet;

function XlsxStreamReaderWorkSheet(workBook, workSheetId, workSheetStream){
	var self = this;

	if (!(this instanceof XlsxStreamReaderWorkSheet)) return new XlsxStreamReaderWorkSheet(workBook, workSheetId, workSheetStream);

	Object.defineProperties(this, { 
		'id': {
			value: workSheetId,
			enumerable: true,
		},
		'workBook': {
			value: workBook,
		},
		'name': {
			value: 'sheet' + workSheetId,
			enumerable: true,
		},
		'options': {
			value: workBook.options,
			writable: true,
			enumerable: true
		},
		'workSheetStream': {
			value: workSheetStream
		},
		'write': {
			value: function(){ return; },
		},
		'end': {
			value: function(){ return; },
		},
		'rowCount': {
			value: 0,
			enumerable: true,
			writable: true
		},
		'sheetData': {
			value: {},
			enumerable: true,
			writable: true
		},
		'pendingCell': {
			value: {},
			enumerable: true,
			writable: true
		},
		'inRows': {
			value: false,
			writable: true
		},
		'workingRow': {
			value: {},
			writable: true
		},
		'abortSheet': {
			value: false,
			writable: true
		},
	});

	self._handleWorkSheetStream();
}
Util.inherits(XlsxStreamReaderWorkSheet, Stream);

XlsxStreamReaderWorkSheet.prototype._handleWorkSheetStream = function(){
	var self = this;

	self.on('pipe', function (srcPipe) {
		self.workBook._parseSax.call(self, srcPipe, self._handleWorkSheetNode, function(){
			 self.emit('end');
		});
	});
}

XlsxStreamReaderWorkSheet.prototype.getColumnNumber = function(columnName){
	var self = this;

	var i = columnName.search(/\d/);
	var colNum = 0;
	columnName = +columnName.replace(/\D/g, function(letter) {
		colNum += (parseInt(letter, 36) - 9) * Math.pow(26, --i);
		return '';
	});

	return colNum;
}

XlsxStreamReaderWorkSheet.prototype.getColumnName = function(columnNumber){
	var self = this;

	if (!columnNumber) return;
	
	var columnName = "";
	var dividend = parseInt(columnNumber);
	var modulo = 0;
	while (dividend > 0){
		modulo = (dividend - 1) % 26;
		columnName = String.fromCharCode(65 + modulo).ToString() + columnName;
		dividend = ((dividend - modulo) / 26);
	}
	return columnName;
}

XlsxStreamReaderWorkSheet.prototype.process = function(){
	var self = this;

	self.workSheetStream.pipe(self);
}

XlsxStreamReaderWorkSheet.prototype.skip = function(){
	var self = this;

	self.workSheetStream.autodrain();
}

XlsxStreamReaderWorkSheet.prototype.abort = function(){
	var self = this;

	self.abortSheet = true;
}

XlsxStreamReaderWorkSheet.prototype._handleWorkSheetNode = function(nodeData){
	var self = this;

	if (self.abortSheet){
		return;
	}

	self.sheetData['cols'] = [];

	switch(nodeData[0].name){
		case 'worksheet':
		case 'sheetPr':
		case 'pageSetUpPr':
			return;

		case 'printOptions':
		case 'pageMargins':
		case 'pageSetup':
			self.inRows = false;
			if (Object.keys(self.workingRow).length > 0){
				delete(self.workingRow.name);
				self.emit('row',self.workingRow);
				self.workingRow = {};
			}			
			break;

		case 'cols':
			return;

		case 'col':
			delete(nodeData[0].name);
			self.sheetData['cols'].push(nodeData[0]);
			return;

		case 'sheetData':
			self.inRows = true;

			nodeData.shift();

		case 'row':
			if (Object.keys(self.workingRow).length > 0){
				delete(self.workingRow.name);
				self.emit('row',self.workingRow);
				self.workingRow = {};
			}

			++self.rowCount;

			self.workingRow = nodeData.shift();
			self.workingRow.values = [];
			break;
	}

	if (self.inRows == true){
		var workingCell = nodeData.shift();
		var workingPart = nodeData.shift();
		var workingVal = nodeData.shift();

		if (!workingCell){
			return;
		}

		if(workingCell.name == 'c'){
			var cellNum = self.getColumnNumber(workingCell.attributes.r)

			//ST_CellType
			switch (workingCell.attributes.t){
				case 's':
					// shared string
					var index = parseInt(workingVal);
					workingVal = self.workBook._getSharedString(index);

					self.workingRow.values[cellNum] = workingVal || "";

					workingCell = {};
					break;
				case 'str':
					// string (formula)
					self.pendingCell = workingCell;
					break;
				case 'b':
					// boolean
				case 'n':
					// number
				case 'e':
					// error
				case 'inlineStr':
					// inline string
				default:
					self.workingRow.values[cellNum] = workingVal || "";

					workingCell = {};					
			}
		}
		if(workingCell.name == 'v'){
			var cellNum = self.getColumnNumber(self.pendingCell.attributes.r)

			self.pendingCell = {};

			self.workingRow.values[cellNum] = workingPart || "";
		}
	}else{
		if (self.sheetData[nodeData[0].name]){
			if (!Array.isArray(self.sheetData[nodeData[0].name])){
				self.sheetData[nodeData[0].name] = [self.sheetData[nodeData[0].name]];
			}
			self.sheetData[nodeData[0].name].push(nodeData);
		}else{
			self.sheetData[nodeData[0].name] = nodeData;
		}
	}
}