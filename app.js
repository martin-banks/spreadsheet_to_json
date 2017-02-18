const XL = require('xlsx')
const fs = require('fs')

const filename = 'sample.xlsx'

console.log('starting app')

let workbook = XL.readFile(__dirname + '/spreadsheets_in/' + filename)
let sheetnames = Object.keys(workbook.Sheets)
let firstSheet = workbook.Sheets[sheetnames[0]]

const columnHeader = {
	headerRow: '1',
	headerColumn: 'A',
	columns: 'abcdefghijklmnopqrstuvwxyz'.toUpperCase().split('')
}

console.log(firstSheet)


const cellIsHeaderRow = cell => {
	return cell[1] === columnHeader.headerRow
}

const cellIsHeaderColumn = cell => {
	return cell[0] === columnHeader.headerColumn
}

const getColumnHeaderValues = sheet => {
	let cells = Object.keys(sheet)
	let headerCells = cells.filter( cell => cellIsHeaderRow(cell) )
	return headerCells.reduce( (headerColumn, cell) => {
		headerColumn[cell] = sheet[cell].w
		return headerColumn
	}, {})
}

const getRowHeaderValues = sheet => {
	let cells = Object.keys(sheet)
	let headerCells = cells.filter( cell => cellIsHeaderColumn(cell) )
	return headerCells.reduce( (headerRow, cell) => {
		headerRow[cell] = sheet[cell].w
		return headerRow
	},{})
}

const getRowData = sheet => {
	let cells = Object.keys(sheet)
	let dataCellKeys = cells.filter( cell => {
		let isNotHeader = !cellIsHeaderColumn(cell) && !cellIsHeaderRow(cell) && cell[0] !== '!'
		return isNotHeader ? true : false
	})
	console.log('\ndataCellKeys', dataCellKeys)
	let rowHeads = Object.keys(getRowHeaderValues(sheet)).map(row => getRowHeaderValues(sheet)[row])
	let columnHeads = Object.keys(getColumnHeaderValues(sheet)).map(column => getColumnHeaderValues(sheet)[column])
	console.log('rows', rowHeads, 'columns', columnHeads)


	let update = dataCellKeys.reduce( (data, cell, i) => {
		console.log('\ndata', data)
		console.log('cell', parseInt(cell[1]))
		let rowKey = rowHeads
		let columnKey = columnHeads[i]
		console.log(
			'rowKey', rowKey,
			'\ncolumnKey', columnKey
		)
		if( !!data[rowKey] ){
			data[rowKey][columnKey] = sheet[cell].w	
		} else {
			data[rowKey] = {
				[columnKey]: sheet[cell].w
			}
		}

		return data
	}, {} )

	return update
}


const jsonData = {
	headers: {
		columns: getColumnHeaderValues(firstSheet),
		rows: getRowHeaderValues(firstSheet)
	},
	data: getRowData(firstSheet)
}
console.log('\njsonData:', JSON.stringify(jsonData,'utf8', 2))

fs.writeFile(__dirname +'/json_out/test.json', JSON.stringify(jsonData, 'utf8', '\t'), err => !!err ? console.log(err) : true)