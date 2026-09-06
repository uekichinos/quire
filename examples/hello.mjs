// Run with:  node examples/hello.mjs   (after `pnpm build`)
// Writes examples/hello.xlsx — open it in Excel / LibreOffice / Numbers to sanity-check.
import { writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createWorkbook } from '../dist/index.js'

const wb = createWorkbook()

const header = {
  font: { bold: true, color: 'FFFFFF' },
  fill: '2F5597',
  align: { horizontal: 'center' },
}
const money = { numFmt: '#,##0.00' }

const sales = wb.addWorksheet('Sales', {
  columns: [{ width: 16 }, { width: 10 }, { width: 14 }, { width: 18 }, { width: 10 }],
  freeze: { ySplit: 2 },
  autoFilter: 'A2:E2',
})

sales.merge('A1:E1')
sales.setCell('A1', 'Q1 sales report', {
  font: { bold: true, size: 14 },
  align: { horizontal: 'center' },
})
sales.setRow(1, { height: 22 })

sales.addRow(
  [
    { value: 'Product', style: header },
    { value: 'Units', style: header },
    { value: 'Revenue', style: header },
    { value: 'Updated', style: header },
    { value: 'In stock', style: header },
  ],
  { height: 18 },
)

sales.addRow(['Widget', 1200, { value: 15003.4, style: money }, new Date(2024, 2, 1, 9, 30), true])
sales.addRow(['Gadget', 340, { value: 2450, style: money }, new Date(2024, 2, 3, 14, 5), false])
sales.addRow(['Doohickey', 0, { value: 0, style: money }, new Date(2024, 1, 29), false])

sales.addRow(
  [
    'Total',
    { value: { formula: 'SUM(B3:B5)', result: 1540 } },
    { value: { formula: 'SUM(C3:C5)', result: 17453.4 }, style: money },
  ],
  { style: { font: { bold: true }, border: { top: { style: 'thin' } } } },
)

const notes = wb.addWorksheet('Notes')
notes.addRow(['Text with entities: a & b < c > "d"'])
notes.addRow(['Unicode: café · 你好 · 📊'])
notes.setColumn(1, { width: 40 })
notes.setCell('A4', 'A wrapped note placed with setCell', { align: { wrapText: true } })

const out = fileURLToPath(new URL('./hello.xlsx', import.meta.url))
writeFileSync(out, wb.xlsx())
console.log('wrote', out)
