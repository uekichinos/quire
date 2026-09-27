/**
 * Basic example — a single sheet: typed cells (string/number/date/formula),
 * simple header styling, column widths, write + read back.
 *
 * Run: node examples/1-basic.mjs   (after `pnpm build`)
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { createWorkbook, readWorkbook } from '../dist/index.js'

const wb = createWorkbook()
const sheet = wb.addWorksheet('Report')

const header = { font: { bold: true, color: 'FFFFFF' }, fill: '2F5597' }
sheet.addRow([
  { value: 'Product', style: header },
  { value: 'Units', style: header },
  { value: 'Unit Price', style: header },
  { value: 'Restocked', style: header },
])

sheet.addRow(['Widget', 120, { value: 4.5, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 1)])
sheet.addRow(['Gadget', 45, { value: 12.0, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 3)])
sheet.addRow(['Gizmo', 300, { value: 1.25, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 10)])

sheet.addRow(
  ['Total units', { value: { formula: 'SUM(B2:B4)', result: 465 } }],
  { style: { font: { bold: true }, border: { top: { style: 'thin' } } } },
)

sheet.setColumn(1, { width: 16 })
sheet.setColumn(4, { width: 14 })

const bytes = wb.xlsx()

mkdirSync(new URL('./output/', import.meta.url), { recursive: true })
writeFileSync(new URL('./output/1-basic.xlsx', import.meta.url), bytes)

// read it back to prove the round trip
const back = readWorkbook(bytes)
console.log('Sheet names:', back.sheetNames)
for (const row of back.sheet('Report').rows()) {
  console.log(row.map((c) => c?.value))
}
