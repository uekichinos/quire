# @uekichinos/quire

[![Socket Badge](https://badge.socket.dev/npm/package/@uekichinos/quire/0.2.0)](https://socket.dev/npm/package/@uekichinos/quire/overview/0.1.0)

A small, dependency-light `.xlsx` **writer** for Node and the browser.

- **One runtime dependency** — [`fflate`](https://github.com/101arrowz/fflate) for zip packaging
- **Writer only** — no OOXML reader, so none of the parser-side CVE surface
  (zip bombs, path traversal, prototype pollution) that affects reader libraries
- Worksheets · typed cells (string / number / boolean / `Date` / formula) ·
  a de-duplicated style pool · merged cells · column widths · freeze panes ·
  auto-filter · row heights
- Returns a `Uint8Array` (or `Blob`); identical in Node and the browser
- Deterministic output (stable bytes for stable input)

## Install

```bash
npm install @uekichinos/quire
```

## Usage

```js
import { createWorkbook } from '@uekichinos/quire'

const wb = createWorkbook()
const sheet = wb.addWorksheet('Sales')

const header = { font: { bold: true, color: 'FFFFFF' }, fill: '2F5597' }
sheet.addRow(
  [
    { value: 'Product', style: header },
    { value: 'Revenue', style: header },
    { value: 'Updated', style: header },
  ],
)

sheet.addRow(['Widget', { value: 15003.4, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 1)])
sheet.addRow(['Gadget', { value: 2450, style: { numFmt: '#,##0.00' } }, new Date(2024, 2, 3)])

sheet.addRow(
  ['Total', { value: { formula: 'SUM(B2:B3)', result: 17453.4 } }],
  { style: { font: { bold: true }, border: { top: { style: 'thin' } } } },
)

// layout
sheet.merge('A1:C1')
sheet.setColumn(1, { width: 24 })
sheet.freeze({ ySplit: 1 })
sheet.autoFilter('A1:C1')
sheet.addRow(['tall row'], { height: 30 })
sheet.setCell('E1', 'note', { align: { wrapText: true } })

// …or up front
wb.addWorksheet('Q2', { columns: [{ width: 20 }], freeze: { ySplit: 1 }, autoFilter: 'A1:C1' })

const bytes = wb.xlsx() // Uint8Array

// Node
import { writeFileSync } from 'node:fs'
writeFileSync('sales.xlsx', bytes)

// Browser
const url = URL.createObjectURL(wb.blob())
```

## API

### `createWorkbook(): Workbook`

| `Workbook` | |
|---|---|
| `addWorksheet(name, options?)` | `options`: `{ columns?, freeze?, autoFilter? }`. Name: 1–31 chars, unique, no `\ / ? * [ ] :` |
| `xlsx()` → `Uint8Array` | deterministic |
| `blob()` → `Blob` | `spreadsheetml.sheet` mime |

| `Worksheet` (all chainable) | |
|---|---|
| `addRow(values, { style?, height? })` | `values`: `CellInput[]` — a bare value or `{ value, style }` |
| `setCell(ref, value, style?)` | write to any A1 reference |
| `setRow(i, { style?, height? })` | style/size a row even with no cells |
| `setColumn(i, { width?, hidden?, style? })` | 1-based; column style resolves **column < row < cell** |
| `merge(range)` | `'A1:C1'`; rejects single-cell / backwards / overlapping |
| `freeze({ xSplit?, ySplit? })` | frozen panes; `freeze({})` clears |
| `autoFilter(range)` | header filter dropdowns |

**Cell values:** `string · number · boolean · Date · { formula, result? } · null`.
Dates use the 1900 serial system (default format `yyyy-mm-dd`). `null` /
`undefined` cells are skipped but keep column position. Non-finite numbers and
pre-1900 dates throw.

**`CellStyle`:** `font` (`name`, `size`, `bold`, `italic`, `underline`, `color`)
· `fill` (`'RRGGBB'` / `'AARRGGBB'`) · `align` (`horizontal`, `vertical`,
`wrapText`, `indent`) · `border` (`top`/`right`/`bottom`/`left`/`all` →
`{ style, color? }`) · `numFmt` (format code or built-in id). Every distinct
style is interned once. Cell style merges over row style over column style, one
nested level deep.

Run `node examples/hello.mjs` (after `pnpm build`) for a full example.

## Reading (`0.2.0`)

```js
import { readWorkbook } from '@uekichinos/quire'

const wb = readWorkbook(bytes)          // Uint8Array | ArrayBuffer

wb.sheetNames                            // ['Sales', …]
const s = wb.sheet('Sales')             // by name or index

s.dimension                              // { rows, cols }
s.merges                                 // ['A1:C1', …]
s.cell('B2')                             // { ref, row, col, type, value, formula? }
s.values()                               // (string|number|boolean|Date|null)[][]

for (const row of s.rows()) {            // sparse: row[col-1]
  console.log(row.map((c) => c?.value))
}
```

`ReadCell.type` is `'string' | 'number' | 'boolean' | 'date' | 'formula' |
'error' | 'empty'`. Numbers with a date format become `Date` (opt out with
`readWorkbook(bytes, { dates: false })`). Formula cells carry both `.formula`
(no leading `=`) and `.value` (the cached result). `{ sheets: [name|index] }`
loads a subset.

**Deliberately strict and small.** The XML is parsed by a ~200-line in-house
tokenizer (not a dependency) that rejects `<!DOCTYPE>`, `<!ENTITY>`, `<![CDATA[>`
and unknown entities outright — the DOCTYPE / entity-expansion class that
accounts for most XML-parser CVEs does not apply. Only an allow-list of parts is
extracted, with uncompressed-size caps enforced before and after inflation.
Still: parse untrusted uploads inside a worker with an overall time/memory limit.

**Not read** (yet): cell styles/formatting on read (opt-in, planned for `0.3.0`),
data validation, conditional formatting, images, charts, pivot tables,
hyperlinks, defined names. Reading is aimed at files from mainstream tools
(Excel, Google Sheets, LibreOffice, `openpyxl`, `exceljs`, quire) — not corrupt
files or every vendor quirk. See [`PLAN-READER.md`](./PLAN-READER.md).

## Scope

`quire` is **not** a drop-in ExcelJS replacement. The **writer** covers the
common "export a styled spreadsheet" case; the **reader** covers "import the
values from a normal `.xlsx`". Out of scope for both (see [`PLAN.md`](./PLAN.md)
and [`PLAN-READER.md`](./PLAN-READER.md)): images, charts, pivot tables, rich
text, a streaming mode, `.xls` / `.xlsb`.

## Performance

In-memory. Writing ~100k rows × 5 cols takes ~1 s; reading is comparable. For
millions of rows, chunk across sheets or wait for a streaming mode.

## License

MIT © [uekichinos](https://www.npmjs.com/~uekichinos). Portions adapted from
[ExcelJS](https://github.com/exceljs/exceljs) (MIT) — see [`NOTICE`](./NOTICE).
