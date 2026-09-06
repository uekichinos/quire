# Changelog

All notable changes to `@uekichinos/quire` are documented here.

## [Unreleased] — reading (targeting 0.2.0)

### R1 — tokenizer, safe unzip, values
- **`readWorkbook(bytes | ArrayBuffer, options?)`** → `ReadWorkbook`
  (`sheetNames`, `sheets`, `sheet(name | index)`, `date1904`)
- `ReadWorksheet`: `name`, `dimension`, `merges`, `cell(ref)`, `rows()` (sparse),
  `toArray()` / `values()` (rectangular)
- `ReadCell`: `{ ref, row, col, type, value, formula? }` — string (shared +
  inline), number, boolean, formula (+ cached value), error, empty
- `src/xml-read.ts` — a strict ~200-line XML tokenizer written in place of a
  dependency after a CVE review of `fast-xml-parser` (13 advisories, most in the
  DOCTYPE / entity-expansion class this reader is exposed to). It **rejects**
  `<!DOCTYPE>` / `<!ENTITY>` / `<![CDATA[>` / unknown entities on sight, has no
  entity expansion, caps nesting depth, and guards object keys against
  `__proto__` pollution
- `src/unzip.ts` — extracts only an allow-list of parts, enforces
  total / per-part uncompressed-size caps from the ZIP directory before *and*
  after inflation, ignores traversal (`..`) entries
- quire stays at **one runtime dependency** (`fflate`)
- 36 new tests, incl. write → read round-trips and a hostile-input suite

## [0.1.0] - 2026-09-06

First release. A dependency-light `.xlsx` **writer** — one runtime dependency
(`fflate`), no OOXML reader.

### Workbook / worksheet
- `createWorkbook()` → chainable `Workbook` / `Worksheet` builder
- `addWorksheet(name, { columns, freeze, autoFilter })` with Excel name-rule
  validation (1–31 chars, unique case-insensitively, no `\ / ? * [ ] :`)
- `xlsx()` → deterministic `Uint8Array`; `blob()` → `Blob`
- shared-string table de-duplicated across the workbook, correct
  `count` / `uniqueCount`; `dimension` computed

### Cells
- `addRow(values, { style, height })`, `setCell(ref, value, style?)`
- string, number, boolean, `Date` (1900 serial system, phantom leap-day
  handled), and formula cells (`{ formula, result? }` → `<f>` + optional `<v>`,
  `t="str"` / `t="b"` for string / boolean results)
- `null` / `undefined` cells skipped, column position preserved
- non-finite numbers and invalid / pre-1900 dates rejected eagerly

### Styling
- de-duplicated style pool: fonts, fills, borders, number formats, cell formats
- `CellStyle`: `font` (name/size/bold/italic/underline/color), `fill`,
  `align` (h/v, wrapText, indent), `border` (per edge or `all`),
  `numFmt` (format code or built-in id; custom codes assigned from 164)
- resolution order **column < row < cell**, merged one nested level deep

### Layout
- `merge(range)` with single-cell / backwards-range / overlap guards
- `setColumn(i, { width, hidden, style })`, `setRow(i, { style, height })`
- `freeze({ xSplit, ySplit })`, `autoFilter(range)` (+ the workbook
  `_xlnm._FilterDatabase` defined name so Excel doesn't flag the file dirty)
- worksheet child-element order follows the CT_Worksheet schema

### Quality
- 100k rows × 5 cols serialises in ~1.1 s / ~3 MB output
- 118 tests; verified against Excel (Windows) at every phase
- `LICENSE` + `NOTICE` cover the MIT portions adapted from ExcelJS

### Not included (by design)
Reading `.xlsx`, images, charts, pivot tables, data validation, conditional
formatting, rich text, a streaming writer, `.xls` / `.xlsb`. See `PLAN.md`.
