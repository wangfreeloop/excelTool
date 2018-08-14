const excel = require('exceljs');

var fills = {
	solid: {type: "pattern", pattern:"solid",fgColor:{argb:"FF996633"}}
};
exports.exportExcel = async (dataArray) => {
	const workbook = new excel.Workbook();
	let worksheet1 = workbook.addWorksheet('客户');
	let worksheet2 = workbook.addWorksheet('Sheet竞品');
	let worksheet3 = workbook.addWorksheet('Sheet负责人');
	let worksheet4 = workbook.addWorksheet('Sheet客户来源');
	let worksheet5 = workbook.addWorksheet('Sheet客户状态');
	let worksheet6 = workbook.addWorksheet('Sheet跟进状态');
	worksheet2.addRow(['竞品']);
	worksheet2.addRow(['校管家']);
	worksheet2.addRow(['校宝']);
	worksheet2.addRow(['小麦助教']);
	worksheet2.addRow(['校盈易']);
	worksheet2.addRow(['学邦']);
	worksheet2.addRow(['天校']);
	worksheet2.addRow(['校360']);
	worksheet2.addRow(['校风云']);
	worksheet2.getCell('A1').fill = fills.solid;
	worksheet2.getCell('A1').font = {color:{argb:'FFFFFFFF'},size:14};

	worksheet3.addRow(['负责人']);
	worksheet3.addRow(['Tom']);
	worksheet3.addRow(['Eva']);
	worksheet3.addRow(['Anna']);
	worksheet3.addRow(['Viola']);
	worksheet3.addRow(['Sara']);
	worksheet3.addRow(['Nimo']);
	worksheet3.addRow(['Lisa']);
	worksheet3.addRow(['Daisy']);
	worksheet3.addRow(['Frank']);
	worksheet3.addRow(['Bill']);
	worksheet3.getCell('A1').fill = fills.solid;
	worksheet3.getCell('A1').font = {color:{argb:'FFFFFFFF'},size:14};

	worksheet4.addRow(['客户来源']);
	worksheet4.addRow(['网络']);
	worksheet4.addRow(['广告']);
	worksheet4.addRow(['搜索引擎']);
	worksheet4.addRow(['客户介绍']);
	worksheet4.addRow(['公司官网']);
	worksheet4.addRow(['独立开发']);
	worksheet4.getCell('A1').fill = fills.solid;
	worksheet4.getCell('A1').font = {color:{argb:'FFFFFFFF'},size:14};

	worksheet5.addRow(['客户状态']);
	worksheet5.addRow(['A(已演示，本月可成交)']);
	worksheet5.addRow(['B(本月可演示，可冲刺)']);
	worksheet5.addRow(['C(有意向了解)']);
	worksheet5.addRow(['无意向']);
	worksheet5.getCell('A1').fill = fills.solid;
	worksheet5.getCell('A1').font = {color:{argb:'FFFFFFFF'},size:14};

	worksheet6.addRow(['跟进状态']);
	worksheet6.addRow(['未联系']);
	worksheet6.addRow(['丢公海']);
	worksheet6.addRow(['练习中']);
	worksheet6.addRow(['预约演示']);
	worksheet6.addRow(['空号、错号']);
	worksheet6.addRow(['已演示']);
	worksheet6.addRow(['已拜访']);
	worksheet6.addRow(['成交']);
	worksheet6.getCell('A1').fill = fills.solid;
	worksheet6.getCell('A1').font = {color:{argb:'FFFFFFFF'},size:14};

	worksheet1.addRow(["机构名称","机构类型","客户名称(必填)","客户来源","客户姓名","手机","客户状态","跟进状态","下次跟进时间","姓名1","手机1","竞品","机构地址","备注","负责人","联系人姓名","联系人电话","联系人手机","电话1","电话2","电话3"]);
	for(var i = 0;i < 26;i++){
		var str = String.fromCharCode(65+i) + '1';
		worksheet1.getCell(str).fill = fills.solid;
		worksheet1.getCell(str).font = {color:{argb:'FFFFFFFF'},size:14};
	}
	for(var p in dataArray){
		var phoneStr = '';
		// console.log("!!!!!!!!!!", dataArray[p].联系电话);
		if(dataArray[p].联系电话){
			phoneStr = dataArray[p].联系电话;
			phoneStr = phoneStr.replace(/[\t]/g, "");
			phoneStr = phoneStr.substr(0,phoneStr.length-1);
			arr = phoneStr.split(';');
		}
		switch (arr.length){
			case 1:
				worksheet1.addRow([dataArray[p].公司名称,"",dataArray[p].法定代表人,"",dataArray[p].法定代表人,arr[0],"","","",dataArray[p].法定代表人,"","",dataArray[p].地址,"邮箱："+dataArray[p].邮箱+"\n经营范围:"+dataArray[p].经营范围,"",dataArray[p].法定代表人,"","","","",""]);
				break;
			case  2:
				worksheet1.addRow([dataArray[p].公司名称,"",dataArray[p].法定代表人,"",dataArray[p].法定代表人,arr[0],"","","",dataArray[p].法定代表人,"","",dataArray[p].地址,"邮箱："+dataArray[p].邮箱+"\n经营范围:"+dataArray[p].经营范围,"",dataArray[p].法定代表人,arr[1],"","","",""]);
				break;
			case 3:
				worksheet1.addRow([dataArray[p].公司名称,"",dataArray[p].法定代表人,"",dataArray[p].法定代表人,arr[0],"","","",dataArray[p].法定代表人,"","",dataArray[p].地址,"邮箱："+dataArray[p].邮箱+"\n经营范围:"+dataArray[p].经营范围,"",dataArray[p].法定代表人,arr[1],"",arr[2],"",""]);
				break;
			default:
				break;
		}

		}
	// 	worksheet1.addRow([dataArray[p].公司名称,"机构类型",dataArray[p].法定代表人,"客户来源",dataArray[p].法定代表人,"手机","客户状态","跟进状态","下次跟进时间",dataArray[p].法定代表人,"手机1","竞品","机构地址","备注","负责人",dataArray[p].法定代表人,"联系人电话","联系人手机","电话1","电话2","电话3"]);
	// }
	workbook.xlsx.writeFile('./export/result.xlsx').then(function(){
	});
};
