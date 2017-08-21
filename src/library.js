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

export default class generateXLSX {
	constructor() {
		this.sheets = [];
	}
	_datenum(v, date1904) {
		if(date1904) {
			v+=1462
		};
		let epoch = Date.parse(v);
		return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
	}
	_arrayToSheet(data, opts) {
		//data is an array of arrays
		//each index of inner array is the row / column placement
		let ws = {};
		let range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
		for(let R = 0; R != data.length; ++R) {
			for(let C = 0; C != data[R].length; ++C) {
				if(range.s.r > R) {
					range.s.r = R;
				}
				if(range.s.c > C) {
					range.s.c = C;
				}
				if(range.e.r < R) {
					range.e.r = R;
				};
				if(range.e.c < C) {
					range.e.c = C;
				};

				let cell = {
					v: data[R][C]
				};
				if(cell.v == null) {
					continue;
				};
				let cell_ref = XLSX.utils.encode_cell({c:C,r:R});
				
				if(typeof cell.v === 'number') {
					cell.t = 'n';
				} else if(typeof cell.v === 'boolean') {
					cell.t = 'b';
				} else if(cell.v instanceof Date) {
					cell.t = 'n'; cell.z = XLSX.SSF._table[14];
					cell.v = datenum(cell.v);
				} else {
					cell.t = 's';
				}
				
				ws[cell_ref] = cell;
			}
		}
		if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
		return ws;
	}
	_wbToArrayBurrfer(wb) {
		var buf = new ArrayBuffer(wb.length);
		var view = new Uint8Array(buf);
		for (var i=0; i!=wb.length; ++i) {
			view[i] = wb.charCodeAt(i) & 0xFF
		};
		return buf;
    }
    /**
     * data is an array of arrays
     * each inner array represents a row in the worksheet
     * @param {string} sheetName 
     * @param {array} data 
     */
	addSheet(sheetName, data) {
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
	generateWorkbook() {
		let workBook = {
			SheetNames: [],
			Sheets: {}
		};
		this.sheets.forEach(sheet => {
			workBook.SheetNames.push(sheet.sheetName);
			workBook.Sheets[sheet.sheetName] = this._arrayToSheet(sheet.sheetData);
		});

		let preppedWB = XLSX.write(workBook, {bookType:'xlsx', bookSST:false, type: 'binary'});
		return new Blob([this._wbToArrayBurrfer(preppedWB)],{type:"application/octet-stream"});
	}
}
