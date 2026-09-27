// Demo for data bars and print setup — for a look in Excel (and Print Preview).
// Run with: node examples/review-databar-printsetup.mjs   (after `pnpm build`)
import { writeFileSync } from 'node:fs'
import { createWorkbook } from '../dist/index.js'

const wb = createWorkbook()
const s = wb.addWorksheet('Review', { columns: [{ width: 16 }, { width: 16 }] })

s.addRow(['Label', 'Value'])
s.addRow(['A', 20])
s.addRow(['B', 55])
s.addRow(['C', 90])
s.addRow(['D', 40])

// Data bar on the Value column.
s.addConditionalFormat('B2:B5', { type: 'dataBar', color: '638EC6' })

// Print setup: landscape, fit to 1 page wide, custom left margin, explicit print area.
s.setPageSetup({
  orientation: 'landscape',
  fitToWidth: 1,
  fitToHeight: 0,
  margins: { left: 0.5 },
  printArea: 'A1:B5',
})

const bytes = wb.xlsx()
const path = new URL('./review-databar-printsetup.xlsx', import.meta.url)
writeFileSync(path, bytes)
console.log('wrote', path.pathname.replace(/^\/([A-Za-z]:)/, '$1'))
