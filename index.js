import "babel-polyfill"
import XLSX from 'xlsx'
import fs from 'fs'
import path from 'path'

const base = '/Users/xinHonglin/Downloads/src/xl'
const files = fs.readdirSync(base)
let xlsxFiles = []
if(!!files) {
	let proxy = []
	proxy = files.map(item => {
		if(item.indexOf('.xls') >= 0) {
			return item
		}
	})
	for(let i=0,len=proxy.length; i<len; i++) {
		if(!!proxy[i]) {
			xlsxFiles.push(proxy[i])
		}
	}
}

const norepeatObj = {}
const allData = xlsxFiles.reduce((prev, next) => prev.concat(dealXlsx(next)), [])
output(allData)

function dealXlsx(file) {
	console.log('deal-file:::::', file)
	console.log('deal=workbook')
	let workbook = null
	try {
		workbook = XLSX.readFile(`${base}/${file}`);
	} catch(err) {
		if(err) return []
	}
	console.log('workbook')
	if(!workbook) return []
	console.log('sheetNames')
	const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
	if(!sheetNames) return []
	console.log('worksheet')
	const worksheet = workbook.Sheets[sheetNames[0]];
	if(!worksheet) return []
	const dataa =XLSX.utils.sheet_to_json(worksheet);
	const _data = []
	dataa.forEach(item => {
		if(!item['姓名']) return false
		if(item['姓名'] == '刘晶' && !norepeatObj[item['日期']]) {
			norepeatObj[item['日期']] = true
			_data.push(item)
		}
	})
	// fs.open(path.resolve(__dirname, `./js/${file}.js`), "w+", '0666',function(e,fd){
	//     if(e) throw e;
	//     fs.writeFile(fd, JSON.stringify(_data), 'utf8', function (err) {
	// 		if (err) throw err;
	// 		console.log('文件写入成功');
	// 	});
	// });
	console.log('...-DONE:::::', file)
	return _data
}

function output(_data) {
	console.log('output')
	const _headers = ['部门班组', '姓名', '日期',	'上午上班', '上班描述'	, '下午下班', '下班描述', 	'考勤情况', '异常原因']

	let headers = _headers
        // 为 _headers 添加对应的单元格位置
        // [ { v: 'id', position: 'A1' },
        //   { v: 'name', position: 'B1' },
        //   { v: 'age', position: 'C1' },
        //   { v: 'country', position: 'D1' },
        //   { v: 'remark', position: 'E1' } ]
        .map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 }))
        // 转换成 worksheet 需要的结构
        // { A1: { v: 'id' },
        //   B1: { v: 'name' },
        //   C1: { v: 'age' },
        //   D1: { v: 'country' },
        //   E1: { v: 'remark' } }
        .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});

    const data = _data
              // 匹配 headers 的位置，生成对应的单元格数据
              // [ [ { v: '1', position: 'A2' },
              //     { v: 'test1', position: 'B2' },
              //     { v: '30', position: 'C2' },
              //     { v: 'China', position: 'D2' },
              //     { v: 'hello', position: 'E2' } ],
              //   [ { v: '2', position: 'A3' },
              //     { v: 'test2', position: 'B3' },
              //     { v: '20', position: 'C3' },
              //     { v: 'America', position: 'D3' },
              //     { v: 'world', position: 'E3' } ],
              //   [ { v: '3', position: 'A4' },
              //     { v: 'test3', position: 'B4' },
              //     { v: '18', position: 'C4' },
              //     { v: 'Unkonw', position: 'D4' },
              //     { v: '???', position: 'E4' } ] ]
              .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) })))
              // 对刚才的结果进行降维处理（二维数组变成一维数组）
              // [ { v: '1', position: 'A2' },
              //   { v: 'test1', position: 'B2' },
              //   { v: '30', position: 'C2' },
              //   { v: 'China', position: 'D2' },
              //   { v: 'hello', position: 'E2' },
              //   { v: '2', position: 'A3' },
              //   { v: 'test2', position: 'B3' },
              //   { v: '20', position: 'C3' },
              //   { v: 'America', position: 'D3' },
              //   { v: 'world', position: 'E3' },
              //   { v: '3', position: 'A4' },
              //   { v: 'test3', position: 'B4' },
              //   { v: '18', position: 'C4' },
              //   { v: 'Unkonw', position: 'D4' },
              //   { v: '???', position: 'E4' } ]
              .reduce((prev, next) => prev.concat(next))
              // 转换成 worksheet 需要的结构
              //   { A2: { v: '1' },
              //     B2: { v: 'test1' },
              //     C2: { v: '30' },
              //     D2: { v: 'China' },
              //     E2: { v: 'hello' },
              //     A3: { v: '2' },
              //     B3: { v: 'test2' },
              //     C3: { v: '20' },
              //     D3: { v: 'America' },
              //     E3: { v: 'world' },
              //     A4: { v: '3' },
              //     B4: { v: 'test3' },
              //     C4: { v: '18' },
              //     D4: { v: 'Unkonw' },
              //     E4: { v: '???' } }
              .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
	// 合并 headers 和 data
	const output = Object.assign({}, headers, data);
	// 获取所有单元格的位置
	const outputPos = Object.keys(output);
	// 计算出范围
	const ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
	// 构建 workbook 对象
	const wb = {
	    SheetNames: ['mySheet'],
	    Sheets: {
	        'mySheet': Object.assign({}, output, { '!ref': ref })
	    }
	};
	console.log('write')
	// 导出 Excel
	XLSX.writeFile(wb, 'output.xlsx');
}

















