# @uekichinos/quire

[![Socket Badge](https://badge.socket.dev/npm/package/@uekichinos/quire/0.6.0)](https://socket.dev/npm/package/@uekichinos/quire/overview/0.6.0)

A small, dependency-light `.xlsx` **reader + writer** for Node and the browser.

- **One runtime dependency** — [`fflate`](https://github.com/101arrowz/fflate) for zip packaging
- **Write:** worksheets · typed cells (string / number / boolean / `Date` / formula /
  hyperlink / rich text) · a de-duplicated style pool · merged cells · column widths ·
  freeze panes · auto-filter · row heights · data validation · conditional formatting
  (highlight rules, colour scales, data bars) · comments · print setup (orientation,
  fit-to-page, margins, print area). Deterministic `Uint8Array` / `Blob` output, plus a
  low-memory [streaming writer](#streaming-writer) for very large sheets
- **Read:** values, dates, formulas (+ cached result), merges, hyperlinks, comments,
  rich text, data validation, conditional formatting, print setup — via an in-house
  strict XML tokenizer, not a dependency, so the DOCTYPE / entity-expansion class
  behind most XML-parser CVEs doesn't apply. A [`readRows`](#streaming-reader) mode
  streams a sheet row-by-row for lower peak memory on large files
- Identical in Node and the browser

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
| `addWorksheet(name, options?)` | `options`: `{ columns?, freeze?, autoFilter?, outline? }`. Name: 1–31 chars, unique, no `\ / ? * [ ] :` |
| `defineName(name, sheetName, range, options?)` | named range, e.g. `wb.defineName('SalesRange', 'Sales', 'A1:B10')`; `{ scope: 'SheetName' }` limits visibility to one sheet |
| `protect(options?)` | structural workbook protection (no password) — `{ lockStructure? (default true), lockWindows? }` |
| `xlsx()` → `Uint8Array` | deterministic |
| `blob()` → `Blob` | `spreadsheetml.sheet` mime |

| `Worksheet` (all chainable) | |
|---|---|
| `addRow(values, { style?, height?, hidden?, outlineLevel? })` | `values`: `CellInput[]` — a bare value or `{ value, style }` |
| `setCell(ref, value, style?)` | write to any A1 reference |
| `setRow(i, { style?, height?, hidden?, outlineLevel? })` | style/size a row even with no cells |
| `setColumn(i, { width?, hidden?, style?, outlineLevel? })` | 1-based; column style resolves **column < row < cell** |
| `merge(range)` | `'A1:C1'`; rejects single-cell / backwards / overlapping |
| `freeze({ xSplit?, ySplit? })` | frozen panes; `freeze({})` clears |
| `autoFilter(range)` | header filter dropdowns |
| `setDataValidation(range, rule, options?)` | dropdown, or a numeric/date/time/text-length/custom-formula rule |
| `addConditionalFormat(range, rule)` | highlight matching cells, a colour scale, a data bar, an icon set, or top/bottom N |
| `setComment(ref, text, { author? })` | a cell comment/note, independent of the cell's value |
| `setPageSetup(options)` | orientation, paper size, fit-to-page/scale, margins, print area |
| `protect(options?)` | structural sheet protection (no password) — see [Protection](#protection) |

**Cell values:** `string · number · boolean · Date · { formula, result? } ·
{ hyperlink, text?, tooltip? } · RichTextRun[] · null`. Dates use the 1900
serial system (default format `yyyy-mm-dd`). `null` / `undefined` cells are
skipped but keep column position. Non-finite numbers and pre-1900 dates throw.

**Hyperlinks:** `{ hyperlink: 'https://…', text?, tooltip? }` as a cell value —
`text` (defaults to the URL itself) becomes the cell's string, `hyperlink` any
URI (`http(s)://`, `mailto:`, …) up to 2,079 characters. Re-setting the cell to
a non-hyperlink value drops it. `sheet.setCell('A1', { hyperlink, text }, style)`
combines it with a style.

**Defined names:** `wb.defineName(name, sheetName, range, options?)` —
`sheetName` must already be added, `range` a single cell (`'A1'`) or a range
(`'A1:C3'`). Names follow Excel's identifier rules (start with a letter/`_`/
`\`, then letters, digits, `_` or `.`; can't read as a cell reference; can't
use the reserved `_xlnm.` prefix) and must be unique (case-insensitively).
Workbook-scoped (visible everywhere) by default; pass `{ scope: 'SheetName' }`
(which must also already exist) to limit visibility to that one sheet — Excel
then requires the sheet-qualified form (`SheetName!SalesRange`) from any other
sheet.

**Outline (grouping):** `addRow(values, { hidden?, outlineLevel? })` /
`setRow(i, { hidden?, outlineLevel? })` / `setColumn(i, { hidden?,
outlineLevel? })` — `outlineLevel` is `0`–`7`, matching Excel's own grouping
depth. `addWorksheet(name, { outline: { summaryBelow?, summaryRight? } })`
controls which side the summary row/column sits on (both default `true`,
Excel's own default).

**Row/column default styles:** `sheet.setRow(i, { style })` / `setColumn(i, {
style })` now also mark the row/column itself with that default style (not
just each populated cell), so empty cells in a styled row or column show the
right formatting in Excel too.

**Rich text:** a cell value can be `RichTextRun[]` — `[{ text, font? }, …]` —
for multiple differently-formatted spans in one string, e.g. a bold word
mid-sentence. `font` is a subset of `CellStyle.font` (`name`, `size`, `bold`,
`italic`, `underline`, `color`) applied per run; the cell's own style (fill,
border, align, numFmt) still applies normally. Must have at least one run.

**Data validation:** `sheet.setDataValidation(range, rule, options?)` restricts
a range to one rule:

```js
// dropdown from fixed choices, or a range reference
sheet.setDataValidation('C2:C100', { type: 'list', list: ['Open', 'In Progress', 'Done'] })
sheet.setDataValidation('D2:D100', { type: 'list', list: 'Lookup!$A$1:$A$5' })

// numeric / date / text-length rules — `value` takes a pair for between/notBetween
sheet.setDataValidation('E2:E100', { type: 'whole', operator: 'between', value: [0, 100] })
sheet.setDataValidation('F2:F100', { type: 'date', operator: 'greaterThan', value: new Date(2024, 0, 1) })
sheet.setDataValidation('G2:G100', { type: 'textLength', operator: 'lessThanOrEqual', value: 50 })
sheet.setDataValidation('H2:H100', { type: 'time', operator: 'greaterThan', value: '09:00' })

// a custom formula (without `=`) that must evaluate truthy — no operator/value
sheet.setDataValidation('I2:I100', { type: 'custom', formula: 'MOD(I2,2)=0' })

// shared options (2nd arg for list, 3rd arg otherwise) — allowBlank, prompt/error text
sheet.setDataValidation('C2:C100', { type: 'list', list: ['Open', 'Done'] }, {
  promptTitle: 'Status', promptMessage: 'Pick one', allowBlank: false,
})
```
For `list`, none of the fixed choices may contain a comma, and the joined list
must fit Excel's 255-character limit. For `whole`/`decimal`/`date`/`time`/
`textLength`, `operator` is one of `between · notBetween · equal · notEqual ·
greaterThan · lessThan · greaterThanOrEqual · lessThanOrEqual`.

**Conditional formatting:** `sheet.addConditionalFormat(range, rule)` — three
rule kinds:

```js
// highlight cells matching a condition
sheet.addConditionalFormat('B2:B100', {
  type: 'cellIs',
  operator: 'greaterThan',
  formula: 100,
  style: { font: { bold: true, color: 'FF0000' }, fill: 'FFF2CC' },
})

// 2- or 3-colour scale (thresholds are the range's own min/mid/max)
sheet.addConditionalFormat('C2:C100', { type: 'colorScale', colors: ['FF0000', 'FFFF00', '00FF00'] })

// data bar (length scales between the range's own min and max)
sheet.addConditionalFormat('D2:D100', { type: 'dataBar', color: '638EC6' })

// icon set (evenly-spaced percentile thresholds) — one of 17 standard Excel names
sheet.addConditionalFormat('E2:E100', { type: 'iconSet', iconSet: '3TrafficLights1' })

// top/bottom N — `percent` interprets rank as a percentage instead of a count
sheet.addConditionalFormat('F2:F100', {
  type: 'top10', rank: 10, bottom: false, style: { fill: 'C6EFCE' },
})
```
`cellIs` styling is limited to `font.bold`/`italic`/`color` and `fill` (Excel's
differential-format record, shared by `top10`); `operator` takes the same set
as data validation, with `formula` as a pair for `between`/`notBetween`.

**Comments:** `sheet.setComment(ref, text, { author? })` attaches a note to a
cell, independent of whatever value (or no value) is there. Uses the classic
comment format (`xl/commentsN.xml` + a VML drawing for the indicator/popup),
not modern threaded comments.

**Print setup:** `sheet.setPageSetup(options)`:

```js
sheet.setPageSetup({
  orientation: 'landscape',
  paperSize: 9,              // Excel paper-size code — 9 = A4, 1 = Letter
  fitToWidth: 1,              // fit to 1 page wide
  fitToHeight: 0,              // any number of pages tall
  margins: { left: 0.5 },     // inches; unset sides fall back to Excel's own defaults
  printArea: 'A1:F50',
})
```
`printArea` is written as the sheet's own `_xlnm.Print_Area` workbook-level
defined name. In the **streaming writer**, `setPageSetup` must be called
before that sheet's first `addRow()`, same as `setColumn`/`freeze`.

### Protection

`sheet.protect(options?)` / `wb.protect(options?)` add **structural** sheet/
workbook protection — **no password support**: Excel's legacy password hash
must be bit-exact, and there's no way to verify that against a real Excel
instance here, so protection is lock-toggle-only.

```js
// every cell is implicitly locked once the sheet is protected — unlock the ones
// that should stay editable via CellStyle.protection
sheet.setCell('A1', 'Editable', { protection: { locked: false } })
sheet.setCell('B1', { formula: 'SECRET()' }, { protection: { hidden: true } }) // hides the formula bar

sheet.protect({ allowSort: true, allowAutoFilter: true }) // everything else stays at Excel's own defaults
wb.protect({ lockWindows: true }) // lockStructure defaults to true
```
The API is deliberately **positive** ("allow X"), matching Excel's own Protect
Sheet dialog defaults, rather than Excel's native XML attributes (which mix
"default allowed" and "default disallowed" polarity depending on the flag) —
quire translates internally so you never have to think about which way a
given flag defaults:

| `SheetProtectionOptions` | default | |
|---|---|---|
| `allowSelectLockedCells` / `allowSelectUnlockedCells` | `true` | selecting locked/unlocked cells |
| `allowFormatCells` / `allowFormatColumns` / `allowFormatRows` | `false` | formatting |
| `allowInsertColumns` / `allowInsertRows` / `allowDeleteColumns` / `allowDeleteRows` | `false` | structure edits |
| `allowSort` / `allowAutoFilter` | `false` | sorting / the autofilter dropdowns |

`WorkbookProtectionOptions` is `{ lockStructure? (default `true`), lockWindows?
(default `false`) }` — locking the sheet list (add/remove/rename/reorder/hide)
and/or the workbook window.

**`CellStyle`:** `font` (`name`, `size`, `bold`, `italic`, `underline`, `color`)
· `fill` (`'RRGGBB'` / `'AARRGGBB'`) · `align` (`horizontal`, `vertical`,
`wrapText`, `indent`) · `border` (`top`/`right`/`bottom`/`left`/`all` →
`{ style, color? }`) · `numFmt` (format code or built-in id) · `protection`
(`{ locked?, hidden? }` — see [Protection](#protection)). Every distinct
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
wb.definedNames                          // [{ name, sheetName?, range?, refersTo, hidden? }]

for (const row of s.rows()) {            // sparse: row[col-1]
  console.log(row.map((c) => c?.value))
}
```

`ReadCell.type` is `'string' | 'number' | 'boolean' | 'date' | 'formula' |
'error' | 'empty'`. Numbers with a date format become `Date` (opt out with
`readWorkbook(bytes, { dates: false })`) — formula cells get the same
treatment from their cached result. A string cell with multiple
differently-formatted spans comes back as `.value: ReadRichTextRun[]` instead
of a plain string. Every cell also carries `.numFmt` (the resolved format
code) when it has one, `.hyperlink` (`{ target, tooltip? }`) when it carries
one — external targets resolved via the worksheet's own relationships,
in-workbook links via `location` — and `.comment` (`{ text, author }`) when it
has a note, whether or not the cell itself has a value. Formula cells carry
both `.formula` (no leading `=`) and `.value` (the cached result). `{ sheets:
[name|index] }` loads a subset.

`readWorkbookAsync(bytes, options?)` returns the same `ReadWorkbook` but yields
to the event loop between sheets, so a large import doesn't block. `s.values({
ragged: true })` keeps each row's own length instead of padding to the sheet
width.

**Limits.** `readWorkbook(bytes, { limits })` caps `maxTotalBytes` (100 MB),
`maxPartBytes` (50 MB), `maxCells` (5 000 000) and `maxSheets` (256); the
archive is inflated part-by-part and aborted mid-stream when a cap trips. Over
the limit throws `QuireError`.

**Styles on read** — `readWorkbook(bytes, { styles: true })` resolves a
`ReadStyle` onto each cell (`font`, `fill`, `border`, `align`, `numFmt`), mirror
of the writer's `CellStyle`. `rgb` colours resolve exactly; indexed colours use
the standard palette; theme colours are read from `xl/theme` (with the 0/1 swap
and an approximate tint). Off by default — most imports only need values.
Under the same option, `sheet.columnStyles` / `sheet.rowStyles` (`Map<number,
ReadStyle>`, 1-based) expose column/row **default** styles — what an empty
cell (no `<c>` element at all) in that column/row would look like.

**Defined names** — `wb.definedNames` lists user-defined named ranges
(`_xlnm._FilterDatabase` and other Excel-internal names are excluded). Each
entry resolves `sheetName`/`range` when the reference is a simple single-sheet
range; anything else (a formula, a multi-area reference, a named constant)
keeps `refersTo` with `sheetName`/`range` left `undefined`. `scope` resolves to
the scoping sheet's name for a sheet-scoped name, `undefined` for a
workbook-scoped one.

**Data validation** — `sheet.dataValidations` lists `list` / `whole` /
`decimal` / `date` / `time` / `textLength` / `custom` rules (every
`ST_DataValidationType` except `none`, which is skipped). A `list` rule
resolves `list` (inline choices) or `formula` (range reference); `custom`
resolves `formula`; `time` resolves `operator` and `values` as `'HH:MM:SS'`
strings; the rest resolve `operator` and `values` (numbers, or `Date`s for
`type: 'date'` — one entry, or two for `between`/`notBetween`).

**Outline (grouping)** — `sheet.columnInfo` / `sheet.rowInfo` (`Map<number,
{ width?/height?, hidden?, outlineLevel? }>`, 1-based) resolve `<col>`/`<row>`
layout facts, whether or not the column/row also carries a default style.

**Conditional formatting** — `sheet.conditionalFormats` lists `cellIs`,
`colorScale`, `dataBar`, `iconSet` and `top10` rules as `{ ref, rule }`.
`cellIs`/`top10` resolve `style` (a `ReadStyle` with only
`font.bold`/`italic`/`color` and `fill` populated — the differential-format's
own limited scope) under `{ styles: true }`; `cellIs` also resolves `operator`
and `formula` (one or two, as written); `top10` also resolves `rank`/
`percent`/`bottom`. `colorScale` resolves `colors`; `dataBar` resolves
`color`; `iconSet` resolves `iconSet` (the standard Excel name).

**Print setup** — `sheet.pageSetup` resolves `orientation`, `paperSize`,
`fitToWidth`/`fitToHeight`, `scale`, `margins`, and `printArea` (from the
sheet's own `_xlnm.Print_Area` defined name) — `undefined` when the sheet has
neither a `<pageSetup>`/`<pageMargins>` block nor a print area.

**Protection** — `sheet.protection` / `wb.protection` resolve the same
positive ("allow"/`lockStructure`/`lockWindows`) shape the writer's
`sheet.protect()`/`wb.protect()` take, `undefined` when not protected. A
cell's `ReadStyle.protection` (`{ locked?, hidden? }`) resolves under
`{ styles: true }`.

**Deliberately strict and small.** The XML is parsed by a ~200-line in-house
tokenizer (not a dependency) that rejects `<!DOCTYPE>`, `<!ENTITY>`, `<![CDATA[>`
and unknown entities outright — the DOCTYPE / entity-expansion class that
accounts for most XML-parser CVEs does not apply. Only an allow-list of parts is
extracted, with uncompressed-size caps enforced before and after inflation.
Still: parse untrusted uploads inside a worker with an overall time/memory limit.

**Not read**: images, charts, pivot tables, threaded (modern) comments,
password-protected sheets/workbooks (protection state reads fine — the
password itself isn't verified or reproduced, since quire has no way to
verify Excel's legacy password hash against a real Excel instance). Reading is
aimed at files from mainstream tools (Excel, Google Sheets, LibreOffice,
`openpyxl`, `exceljs`, quire) — not corrupt files or every vendor quirk. See
[`PLAN-READER.md`](./PLAN-READER.md).

## Streaming reader

`readWorkbook()` materialises every cell of every loaded sheet in memory.
`readRows()` streams one sheet's rows through a callback instead, so peak
memory scales with row *width*, not row *count*:

```js
import { readRows } from '@uekichinos/quire'

const meta = readRows(bytes, 'Sales', (cells, row) => {
  db.insert(row, cells.map((c) => c?.value))
})

meta.dimension        // { rows, cols }
meta.dataValidations  // sheet-level metadata is still resolved up front
meta.pageSetup
```
`cells[i]` is `undefined` for a blank cell, matching `ReadWorksheet.rows()`.
Sheet-level metadata (`merges`, `columnStyles`/`rowStyles`, `dataValidations`,
`conditionalFormats`, `pageSetup`) comes back in the return value alongside
the stream. Two honest limits: **hyperlinks and comments aren't attached** to
streamed rows (resolving them needs a second pass over the sheet's cells,
which this mode specifically avoids holding — use `readWorkbook` if you need
them), and it's still **one synchronous parse pass** under the hood, same as
`readWorkbook` for a single sheet — this reduces memory, not blocking; there's
no async/streaming variant, since the hand-rolled tokenizer isn't
interruptible mid-parse.

## Scope

`quire` is **not** a drop-in ExcelJS replacement. The **writer** covers the
common "export a styled spreadsheet" case; the **reader** covers "import the
values from a normal `.xlsx`". Out of scope for both (see [`PLAN.md`](./PLAN.md)
and [`PLAN-READER.md`](./PLAN-READER.md)): images, charts, pivot tables,
`.xls` / `.xlsb`.

## Streaming writer

`createWorkbook()`'s `xlsx()` holds every cell in memory until you call it —
fine up to a few hundred thousand rows, wasteful well beyond that.
`createStreamingWorkbook()` renders and DEFLATEs each row as it's added instead,
so memory stays flat regardless of row count:

```js
import { createStreamingWorkbook } from '@uekichinos/quire'

const wb = createStreamingWorkbook()
const sheet = wb.addWorksheet('Big', { columns: [{ width: 20 }], freeze: { ySplit: 1 } })

sheet.addRow(['ID', 'Name', 'Created'])
for (let i = 1; i <= 2_000_000; i++) {
  sheet.addRow([i, `Row ${i}`, new Date(2024, 0, 1)])
}

const bytes = await wb.finish() // Uint8Array — same as xlsx(), just async
```

Everything from the buffered writer works — hyperlinks, rich text, merges,
freeze panes, auto-filter, data validation, conditional formatting, comments,
print setup, outline (grouping), protection, `defineName` — with one
constraint: it's **one pass, forward-only**. There's no `setCell`/`setRow`
random access, and `setColumn`/`freeze`/`setPageSetup` must be called before
that sheet's first `addRow()` (they render into the header, which is flushed
immediately). `merge`/`autoFilter`/`setDataValidation`/`addConditionalFormat`/
`setComment`/`protect()` (sheet and workbook) can be called any time before
`finish()`. Multiple sheets can be written interleaved or in any order.

## Performance

In-memory (buffered) writing of ~100k rows × 5 cols takes ~1 s; reading is
comparable. Past that, or for an unbounded/very large row count, use the
[streaming writer](#streaming-writer) and/or the [streaming reader](#streaming-reader).

## License

MIT © [uekichinos](https://www.npmjs.com/~uekichinos). Portions adapted from
[ExcelJS](https://github.com/exceljs/exceljs) (MIT) — see [`NOTICE`](./NOTICE).
