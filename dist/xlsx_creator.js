(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else if(typeof exports === 'object')
		exports["SET THIS"] = factory();
	else
		root["SET THIS"] = factory();
})(this, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

/**
    app name xlsx_creator
 */
function depCheck() {
	try {
		var dep1 = XLSX;
	} catch (error) {
		throw new Error("The XLSX.js full library is required to use the generateXLSX class.");
	}
}
depCheck();

var generateXLSX = function () {
	function generateXLSX() {
		_classCallCheck(this, generateXLSX);

		this.sheets = [];
	}

	_createClass(generateXLSX, [{
		key: '_datenum',
		value: function _datenum(v, date1904) {
			if (date1904) {
				v += 1462;
			};
			var epoch = Date.parse(v);
			return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
		}
	}, {
		key: '_arrayToSheet',
		value: function _arrayToSheet(data, opts) {
			//data is an array of arrays
			//each index of inner array is the row / column placement
			var ws = {};
			var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
			for (var R = 0; R != data.length; ++R) {
				for (var C = 0; C != data[R].length; ++C) {
					if (range.s.r > R) {
						range.s.r = R;
					}
					if (range.s.c > C) {
						range.s.c = C;
					}
					if (range.e.r < R) {
						range.e.r = R;
					};
					if (range.e.c < C) {
						range.e.c = C;
					};

					var cell = {
						v: data[R][C]
					};
					if (cell.v == null) {
						continue;
					};
					var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

					if (typeof cell.v === 'number') {
						cell.t = 'n';
					} else if (typeof cell.v === 'boolean') {
						cell.t = 'b';
					} else if (cell.v instanceof Date) {
						cell.t = 'n';cell.z = XLSX.SSF._table[14];
						cell.v = datenum(cell.v);
					} else {
						cell.t = 's';
					}

					ws[cell_ref] = cell;
				}
			}
			if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
			return ws;
		}
	}, {
		key: '_wbToArrayBurrfer',
		value: function _wbToArrayBurrfer(wb) {
			var buf = new ArrayBuffer(wb.length);
			var view = new Uint8Array(buf);
			for (var i = 0; i != wb.length; ++i) {
				view[i] = wb.charCodeAt(i) & 0xFF;
			};
			return buf;
		}
		/**
   * data is an array of arrays
   * each inner array represents a row in the worksheet
   * @param {string} sheetName 
   * @param {array} data 
   */

	}, {
		key: 'addSheet',
		value: function addSheet(sheetName, data) {
			this.sheets.push({
				sheetName: sheetName,
				sheetData: data
			});
		}
		/**
   * Call to generate xlsx file
   * returns a binary blob
   * blob you can 
   * save single file with FileSaver.js,
   * add to zip file with jszip.js
   * @param {string} workBookName 
   * @returns {Blob}
   */

	}, {
		key: 'generateWorkbook',
		value: function generateWorkbook() {
			var _this = this;

			var workBook = {
				SheetNames: [],
				Sheets: {}
			};
			this.sheets.forEach(function (sheet) {
				workBook.SheetNames.push(sheet.sheetName);
				workBook.Sheets[sheet.sheetName] = _this._arrayToSheet(sheet.sheetData);
			});

			var preppedWB = XLSX.write(workBook, { bookType: 'xlsx', bookSST: false, type: 'binary' });
			return new Blob([this._wbToArrayBurrfer(preppedWB)], { type: "application/octet-stream" });
		}
	}]);

	return generateXLSX;
}();

/* harmony default export */ __webpack_exports__["default"] = (generateXLSX);

/***/ })
/******/ ]);
});