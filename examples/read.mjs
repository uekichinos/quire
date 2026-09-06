// Run with:  node examples/read.mjs   (after `pnpm build`)
// Writes a workbook, reads it straight back, and prints what came out.
import { createWorkbook, readWorkbook } from '../dist/index.js'

const wb = createWorkbook()
const s = wb.addWorksheet('Sales', { autoFilter: 'A1:D1' })
s.addRow(['Product', 'Units', 'Revenue', 'Updated'], { style: { font: { bold: true } } })
s.addRow(['Widget', 1200, { value: 15003.4, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 1)])
s.addRow(['Gadget', 340, { value: 2450, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 3)])
s.setCell('C4', { formula: 'SUM(C2:C3)', result: 17453.4 })
s.merge('A6:D6')
s.setCell('A6', 'end of report')

const back = readWorkbook(wb.xlsx(), { styles: true })

console.log('sheets:', back.sheetNames)
const sheet = back.sheet('Sales')
console.log('dimension:', sheet.dimension, ' merges:', sheet.merges)
for (const row of sheet.rows()) {
  console.log(
    row.map((c) => (c ? `${c.value}${c.formula ? ` (=${c.formula})` : ''} [${c.type}]` : '·')),
  )
}

console.log('\nA1 style:', JSON.stringify(sheet.cell('A1').style))
console.log('C2 style:', JSON.stringify(sheet.cell('C2').style))
