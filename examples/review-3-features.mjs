// Demo workbook for data validation, rich text, and comments — for a look in Excel.
// Run with: node examples/review-3-features.mjs   (after `pnpm build`)
import { writeFileSync } from 'node:fs'
import { createWorkbook } from '../dist/index.js'

const wb = createWorkbook()
const s = wb.addWorksheet('Review')

s.addRow([
  { value: 'Task', style: { font: { bold: true }, fill: 'D9D9D9' } },
  { value: 'Status', style: { font: { bold: true }, fill: 'D9D9D9' } },
  { value: 'Notes', style: { font: { bold: true }, fill: 'D9D9D9' } },
])

// Rich text: multiple formatted runs in one cell.
s.setCell('A2', [
  { text: 'Order ' },
  { text: '#1234', font: { bold: true, color: 'FF0000' } },
  { text: ' — ' },
  { text: 'shipped', font: { italic: true, color: '008000' } },
])
s.setCell('A3', [
  { text: 'Plain prefix, ' },
  { text: 'Arial 14pt suffix', font: { name: 'Arial', size: 14 } },
])

// Data validation: dropdown list on the Status column.
s.setDataValidation('B2:B20', {
  list: ['Open', 'In Progress', 'Done'],
  promptTitle: 'Status',
  promptMessage: 'Pick one of the three values',
})
s.setCell('B2', 'Open')
s.setCell('B3', 'Done')

// Comments: on a populated cell, and on a totally empty one.
s.setComment('A2', 'This is the record we discussed on the call.', { author: 'Reviewer' })
s.setComment('C2', 'Second reviewer note.', { author: 'QA' })
s.setComment('D5', 'Comment on an empty cell — should still show a red triangle here.')

s.setColumn(1, { width: 30 })
s.setColumn(2, { width: 16 })
s.setColumn(3, { width: 20 })

const bytes = wb.xlsx()
const path = new URL('./review-3-features.xlsx', import.meta.url)
writeFileSync(path, bytes)
console.log('wrote', path.pathname.replace(/^\/([A-Za-z]:)/, '$1'))
