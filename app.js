const fs = require('fs');
const join = require('path').join;
const exportExcel = require('./exportExcel');
const importExcel = require('./importExcel');

function findSync(startPath) {
	let result=[];
	function finder(path) {
		let files=fs.readdirSync(path);
		files.forEach((val,index) => {
			let fPath=join(path,val);
			let stats=fs.statSync(fPath);
			if(stats.isDirectory()) finder(fPath);
			if(stats.isFile()) result.push(fPath);
		});
	}
	finder(startPath);
	return result;
}
let fileNames=findSync('./import');
let fileArr = [];
fileNames.forEach(function(item, ind) {
	let str = item.substr(10, item.length);
	fileArr.push(str);
});

(async function() {
	const dataArray = await importExcel.importExcel(fileNames);
	await exportExcel.exportExcel(dataArray);
})();

