/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict';

const Util = require('util');
const Stream = require('stream');

module.exports = XlsxStreamReaderWorkSheet;

function XlsxStreamReaderWorkSheet(workBook, sheetName, workSheetId, workSheetStream){
	var self = this;

	if (!(this instanceof XlsxStreamReaderWorkSheet)) return new XlsxStreamReaderWorkSheet(workBook, sheetName, workSheetId, workSheetStream);

	Object.defineProperties(this, { 
		'id': {
			value: workSheetId,
			enumerable: true,
		},
		'workBook': {
			value: workBook,
		},
		'name': {
			value: sheetName,
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
		self.workBook._parseXML.call(self, srcPipe, self._handleWorkSheetNode, function(){
			if (self.workingRow.name){
				delete(self.workingRow.name);
				self.emit('row',self.workingRow);
				self.workingRow = {};
			}
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
		columnName = String.fromCharCode(65 + modulo).toString() + columnName;
		dividend = Math.floor(((dividend - modulo) / 26));
	}
	return columnName;
}

XlsxStreamReaderWorkSheet.prototype.process = function(){
	var self = this;

	self.workSheetStream.pipe(self);
}

XlsxStreamReaderWorkSheet.prototype.skip = function(){
	var self = this;

	if (self.workSheetStream instanceof Stream){
		setImmediate(self.emit.bind(self),'end');
	}else{
		self.workSheetStream.autodrain();
	}
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
			if (self.workingRow.name){
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
			if (self.workingRow.name){
				delete(self.workingRow.name);
				self.emit('row',self.workingRow);
				self.workingRow = {};
			}

			++self.rowCount;

			self.workingRow = nodeData.shift() || {};
			if (typeof self.workingRow !== 'object') {
				self.workingRow = {}
			}
			self.workingRow.values = [];
			self.workingRow.formulas = [];
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

			if (workingPart && workingPart.name && workingPart.name == "f"){
				self.workingRow.formulas[cellNum] = workingVal;
			}

			//ST_CellType
			switch (workingCell.attributes.t){
				case 's':
					// shared string
					var index = parseInt(workingVal);
					workingVal = self.workBook._getSharedString(index);

					self.workingRow.values[cellNum] = workingVal || "";

					workingCell = {};
					break;
				case 'inlineStr':
					// inline string
					self.workingRow.values[cellNum] = nodeData.shift() || "";

					workingCell = {};	
					break;
				case 'str':
					// string (formula)
				case 'b':
					// boolean
				case 'n':
					// number
        		case 'e':
					// error
				default:
					if (self.options.formatting) {
						var numFmts = self.workBook.workBookStyles.filter(function (item) {
							return item.name === "numFmt"
						});
						if (numFmts.length > 0) {
							var cellXfs_index = self.workBook.workBookStyles.findIndex(function (item) {
								return item.name === "cellXfs"
							});
							var xfs = self.workBook.workBookStyles.filter(function (item, index) {
								return item.name === "xf" && index > cellXfs_index
							});
							var formatId = xfs[workingCell.attributes.s];
							if (formatId) { // if formatId exists in workbookStyles
								formatId = formatId.attributes.numFmtId
								var format = numFmts.find(function (item) {
									return item.attributes.numFmtId === formatId
								});
		
								if (format && format.attributes && isCodedDate(format.attributes.formatCode.toLowerCase())) {
									var date = new Date((+workingVal - (25567 + 2)) * 86400 * 1000);
									if (date.toString() !== 'Invalid Date') {
										var dd = date.getDate()
										var mm = (date.getMonth() + 1)
										if (dd < 10) {
											dd = '0' + dd;
										} 
										if (mm < 10) {
											mm = '0' + mm;
										} 
										workingVal = dd + '/' + mm + '/' + date.getFullYear();
									}
								}
							}
						} else if (!isNaN(parseFloat(workingVal))) { // this is number
							workingVal = parseFloat(parseFloat(workingVal).toFixed(2)); // parse to float or int, but only with maximum of 2 decimals
						}
					}

					self.workingRow.values[cellNum] = workingVal || "";

					workingCell = {};					
			}
		}
		if(workingCell.name == 'v'){
			self.workingRow.values[cellNum] = workingPart || "";
		}
	}else{
		if (self.sheetData[nodeData[0].name]){
			if (!Array.isArray(self.sheetData[nodeData[0].name])){
				self.sheetData[nodeData[0].name] = [self.sheetData[nodeData[0].name]];
			}
			self.sheetData[nodeData[0].name].push(nodeData);
		}else{
			if (nodeData[0].name) {
				self.sheetData[nodeData[0].name] = nodeData;
			}
		}
	}
}

function isCodedDate (formatCode) {
	// recognise coded dates
	return ~formatCode.indexOf('d') &&
		~formatCode.indexOf('m') &&
		~formatCode.indexOf('yy')
}
