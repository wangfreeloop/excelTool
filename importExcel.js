const excel = require('exceljs');

exports.importExcel = async (fileArr) => {
	const workbook = new excel.Workbook();
	let dataArray = [];
	for(let i = 0;i < fileArr.length;i++) {
		await workbook.xlsx.readFile(fileArr[i]).then(function(){
			let keys = [];
			const worksheet = workbook.getWorksheet(1);
			const row = worksheet.getRow(2);
			worksheet.eachRow(function(row, rowNumber) {
				if(rowNumber == 1){}else if(rowNumber == 2) {
					keys = row.values;
				}
				else {
					let rowDict = cellValueToDict2(keys, row);
					dataArray.push(rowDict);
				}
			});
			// console.log(JSON.stringify(dataArray));
		});
	}

	return dataArray;
};

function cellValueToDict2(keys,row){
	let data = {};
	row.eachCell(function(cell, colNumber){
		var value = cell.value;
		if(typeof value == "object") value = value.text;
		data[keys[colNumber]]  = value;
	});
	return data;
}
