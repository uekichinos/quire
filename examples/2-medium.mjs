/**
 * Medium example — a small sales report: a merged title, freeze panes,
 * auto-filter, a data-validation dropdown, a conditional-formatting rule,
 * hyperlinks, a defined name, and print setup.
 *
 * Run: node examples/2-medium.mjs   (after `pnpm build`)
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { createWorkbook, readWorkbook } from '../dist/index.js'

const wb = createWorkbook()
const sheet = wb.addWorksheet('Sales', { freeze: { ySplit: 2 }, autoFilter: 'A2:E2' })

sheet.merge('A1:E1')
sheet.setCell('A1', 'Q2 Sales Report', { font: { bold: true, size: 14 } })

const header = { font: { bold: true, color: 'FFFFFF' }, fill: '2F5597' }
sheet.addRow([
  { value: 'Product', style: header },
  { value: 'Revenue', style: header },
  { value: 'Status', style: header },
  { value: 'Owner', style: header },
  { value: 'Notes', style: header },
])

sheet.addRow([
  'Widget',
  { value: 15003.4, style: { numFmt: '#,##0.00' } },
  'Open',
  { hyperlink: 'mailto:jane@example.com', text: 'Jane' },
  'Restocking',
])
sheet.addRow([
  'Gadget',
  { value: 2450, style: { numFmt: '#,##0.00' } },
  'Done',
  { hyperlink: 'mailto:sam@example.com', text: 'Sam' },
  '',
])
sheet.addRow([
  'Gizmo',
  { value: 38999.99, style: { numFmt: '#,##0.00' } },
  'In Progress',
  { hyperlink: 'mailto:jane@example.com', text: 'Jane' },
  'Large order — confirm stock',
])

sheet.setDataValidation('C3:C100', { type: 'list', list: ['Open', 'In Progress', 'Done'] })
sheet.addConditionalFormat('B3:B100', {
  type: 'cellIs',
  operator: 'greaterThan',
  formula: 10000,
  style: { font: { bold: true, color: '006100' }, fill: 'C6EFCE' },
})

wb.defineName('SalesRange', 'Sales', 'B3:B5')

sheet.setColumn(1, { width: 14 })
sheet.setColumn(2, { width: 14 })
sheet.setColumn(3, { width: 14 })
sheet.setColumn(4, { width: 18 })
sheet.setColumn(5, { width: 26 })

sheet.setPageSetup({ orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 })

const bytes = wb.xlsx()

mkdirSync(new URL('./output/', import.meta.url), { recursive: true })
writeFileSync(new URL('./output/2-medium.xlsx', import.meta.url), bytes)

// read it back and sanity-check the metadata quire attaches, not just values
const back = readWorkbook(bytes)
const back_sheet = back.sheet('Sales')
console.log('Merges:', back_sheet.merges)
console.log('Data validations:', back_sheet.dataValidations)
console.log('Conditional formats:', back_sheet.conditionalFormats)
console.log('Defined names:', back.definedNames)
console.log('B3 hyperlink? no —', back_sheet.cell('B3').hyperlink, '  D3 hyperlink:', back_sheet.cell('D3').hyperlink)
