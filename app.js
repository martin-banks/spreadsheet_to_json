const XL = require('xlsx')
const fs = require('fs')

const filename = 'sample.xlsx'

console.log('starting app')

let workbook = XL.readFile(__dirname + '/spreadsheets_in/' + filename)
let sheetnames = Object.keys(workbook.Sheets)
let firstSheet = workbook.Sheets[sheetnames[0]]
console.log(firstSheet)

const columnHeader = {
	headerRow: '1',
	headerColumn: 'A',
	columns: 'abcdefghijklmnopqrstuvwxyz'.toUpperCase().split('')
}

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


	let rowHeads = getRowHeaderValues(sheet)
	let columnHeads = getColumnHeaderValues(sheet)

	/*let rowHeads = Object.keys(getRowHeaderValues(sheet)).reduce(
		(update, row) => {
			console.log(getRowHeaderValues(sheet))
			let value = getRowHeaderValues(sheet)[row]
			update[row] = value
			return update
		}, {} 
	)*/

	/*let columnHeads = Object.keys(getColumnHeaderValues(sheet)).map(
		column => getColumnHeaderValues(sheet)[column])*/

	/*let columnHeads = Object.keys(getColumnHeaderValues(sheet)).reduce(
		(update, column) => {
			let value = getColumnHeaderValues(sheet)[column]
			update[column] = value
			return update
		}, {} 
	)*/

	
	console.log('\nrows', rowHeads, '\ncolumns', columnHeads)


	let dataToReturn = dataCellKeys.reduce( (data, cell, i) => {
		console.log('\n\ndata', data)
		console.log('cell', cell)
		let cellKey = {
			letter: cell.split('').filter( character => !parseInt(character) ).join(''),
			number: parseInt(cell.split('').filter( character => !!parseInt(character) ).join(''))
		}
		console.log('cell key', cellKey)
		let rowKey = Object.keys(rowHeads)
			.map( row => {
				let headAsArray = row.split('')
				let headNumbers = headAsArray.filter( node => !!parseInt(node) )
				return parseInt(headNumbers.join(''))
			})
			.filter( number => {
				return number === cellKey.number
			})
			.join('')

		let columnKey = Object.keys(columnHeads)
			.map( column => column.split('')[0] )
			.filter( col => col === cellKey.letter )
			.join('')

		console.log(
			'rowKey', rowKey,
			'\ncolumnKey', columnKey
		)
		console.log('contructing json', columnHeads, '\n')
		if( !!data[rowHeads[ 'A' + cellKey.number ]] ){
			data[rowHeads[ 'A' + cellKey.number ]][columnHeads[columnKey + '1']] = sheet[cell].w	
		} else {
			data[rowHeads[ 'A' + cellKey.number ]] = {
				[columnHeads[columnKey + '1']]: sheet[cell].w
			}
		}

		return data
	}, {} )

	return dataToReturn
}



const jsonData = {
	headers: {
		columns: getColumnHeaderValues(firstSheet),
		rows: getRowHeaderValues(firstSheet)
	},
	data: getRowData(firstSheet)
}
console.log('\njsonData:', JSON.stringify(jsonData,'utf8', 2))


fs.writeFile(__dirname +'/json_out/test.json', 
	JSON.stringify(jsonData, 'utf8', '\t'), 
	err => !!err ? console.log(err) : true
)