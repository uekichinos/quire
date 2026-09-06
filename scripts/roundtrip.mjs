// CI gate. Two directions, both through headless LibreOffice:
//   1. quire writes .xlsx  → LibreOffice reads it (convert to CSV), values survive
//   2. LibreOffice writes .xlsx (from CSV) → quire reads it, values survive
// Fails loudly on any LibreOffice error or missing data.
//
//   node scripts/roundtrip.mjs
//
// Requires `libreoffice-calc` (or `soffice`) on PATH.
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, writeFileSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createWorkbook, readWorkbook } from '../dist/index.js'

const dir = mkdtempSync(join(tmpdir(), 'quire-'))
const xlsx = join(dir, 'roundtrip.xlsx')

const wb = createWorkbook()
const s = wb.addWorksheet('Sales', { freeze: { ySplit: 2 }, autoFilter: 'A2:D2' })
s.setColumn(1, { width: 20 })
s.merge('A1:D1')
s.setCell('A1', 'Q1 sales report', { font: { bold: true, size: 14 } })
s.addRow(['Product', 'Units', 'Revenue', 'Updated'], { style: { font: { bold: true } } })
s.addRow(['Widget', 1200, { value: 15003.4, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 1)])
s.addRow(['Gadget', 340, { value: 2450, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 3)])
s.addRow(['Total', { value: { formula: 'SUM(B3:B4)', result: 1540 } }])
writeFileSync(xlsx, wb.xlsx())

const bin = process.env.SOFFICE_BIN || 'soffice'
let out
try {
  out = execFileSync(bin, ['--headless', '--convert-to', 'csv', '--outdir', dir, xlsx], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  })
} catch (err) {
  console.error('LibreOffice conversion failed:\n', err.stderr || err.message)
  process.exit(1)
}
console.log(out.trim())

const csvName = readdirSync(dir).find((f) => f.endsWith('.csv'))
if (!csvName) {
  console.error('no CSV produced — LibreOffice rejected the file')
  process.exit(1)
}
const csv = readFileSync(join(dir, csvName), 'utf8')
console.log('--- direction 1: LibreOffice read our .xlsx ---\n' + csv)

const required = ['Product', 'Widget', '1200', '15003.4', 'Gadget', 'Total']
const missing = required.filter((token) => !csv.includes(token))
if (missing.length) {
  console.error('CSV is missing expected values:', missing)
  process.exit(1)
}

// --- direction 2: LibreOffice writes an .xlsx, quire reads it ------------------
const srcCsv = join(dir, 'src.csv')
writeFileSync(srcCsv, 'Name,Qty,Price\nAlpha,3,1.5\nBravo,10,2\nCharlie,0,9.99\n')
try {
  execFileSync(bin, ['--headless', '--convert-to', 'xlsx', '--outdir', dir, srcCsv], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  })
} catch (err) {
  console.error('LibreOffice csv→xlsx failed:\n', err.stderr || err.message)
  process.exit(1)
}
const loXlsx = readdirSync(dir).find((f) => f === 'src.xlsx')
if (!loXlsx) {
  console.error('LibreOffice produced no src.xlsx')
  process.exit(1)
}
const wb = readWorkbook(readFileSync(join(dir, loXlsx)))
const rows = wb.sheets[0].values()
console.log('--- direction 2: quire read LibreOffice .xlsx ---')
console.log(rows)
const flat = JSON.stringify(rows)
for (const token of ['"Name"', '"Alpha"', '3', '1.5', '"Charlie"', '9.99']) {
  if (!flat.includes(token)) {
    console.error('quire read of LibreOffice output is missing:', token)
    process.exit(1)
  }
}

console.log('\nroundtrip OK (both directions)')
