import XLSX from 'xlsx'
import path from 'path'

const file = path.resolve('./ejemplos/ejemplo-output.xlsx')
const wb = XLSX.readFile(file)
const sheetName = wb.SheetNames[0]
const ws = wb.Sheets[sheetName]
const json = XLSX.utils.sheet_to_json(ws, { defval: '' })
console.log('Sheet:', sheetName)
console.log('Rows:', json.length)
console.log('First rows:', JSON.stringify(json.slice(0,5), null, 2))
