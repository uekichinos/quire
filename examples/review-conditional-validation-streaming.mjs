// Demo for extended data validation, conditional formatting, and the streaming
// writer — for a look in Excel. Run with:
//   node examples/review-conditional-validation-streaming.mjs   (after `pnpm build`)
import { writeFileSync } from 'node:fs'
import { createStreamingWorkbook } from '../dist/index.js'

const wb = createStreamingWorkbook()
const s = wb.addWorksheet('Review', { columns: [{ width: 14 }, { width: 22 }, { width: 14 }] })

// Extended data validation.
s.setDataValidation('A2:A20', { type: 'whole', operator: 'between', value: [0, 100] })
s.setDataValidation('B2:B20', { type: 'date', operator: 'greaterThan', value: new Date(2024, 0, 1) })
s.setDataValidation('C2:C20', { type: 'list', list: ['Open', 'In Progress', 'Done'] })

// Conditional formatting.
s.addConditionalFormat('A2:A20', {
  type: 'cellIs',
  operator: 'greaterThan',
  formula: 80,
  style: { font: { bold: true, color: 'FF0000' }, fill: 'FFF2CC' },
})
s.addConditionalFormat('D2:D20', { type: 'colorScale', colors: ['FF0000', 'FFFF00', '00FF00'] })

s.addRow(['Score (0-100)', 'Due date', 'Status', 'Heat'])
s.addRow([95, new Date(2024, 5, 1), 'Done', 90])
s.addRow([40, new Date(2024, 6, 1), 'Open', 40])
s.addRow([70, new Date(2024, 7, 1), 'In Progress', 65])

const bytes = await wb.finish()
const path = new URL('./review-conditional-validation-streaming.xlsx', import.meta.url)
writeFileSync(path, bytes)
console.log('wrote', path.pathname.replace(/^\/([A-Za-z]:)/, '$1'), `(${bytes.length} bytes, via the streaming writer)`)
