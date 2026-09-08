# Changelog

All notable changes to `@uekichinos/quire` are documented here.

## [0.3.0] - 2026-09-08

Cell styles on read, plus a round of writer/reader hardening. Still **one
runtime dependency** (`fflate`).

### Hardening
- **`xlsx()` is now idempotent** — the shared-string table and style pool are
  built fresh per serialise call instead of being accumulated on the worksheet,
  so `wb.xlsx()` (or `blob()`) called twice returns byte-identical output
- **Large / tiny numbers** are written as plain decimals — `numToXml()` expands
  JavaScript's exponent notation (`1e21`, `1e-7`) so Excel never sees `1e+21`
- **Streaming unzip** — `src/unzip.ts` inflates part-by-part with a running byte
  budget and aborts mid-stream when a cap is hit, instead of decompressing the
  whole archive first (zip-bomb resistance)
- **Read caps** — `readWorkbook(bytes, { limits })` accepts `maxCells`
  (default 5 000 000) and `maxSheets` (default 256) on top of the existing
  byte caps; over-limit input throws `QuireError`
- `readWorkbookAsync(bytes, options?)` — same result, yields to the event loop
  between sheets so a large import doesn't block the main thread
- `ReadCell.numFmt` — the resolved number-format code is now on every cell,
  not only under `{ styles: true }`
- `ReadWorksheet.values({ ragged: true })` — rows keep their own length instead
  of being padded to the sheet width
- All thrown errors now derive from a single `QuireError` base (`XlsxReadError`,
  `XmlError` included); exported for `instanceof` checks
- New tests: `numToXml` unit coverage, `xlsx()` idempotency, a real-world-quirk
  fixture corpus (Excel `mc:AlternateContent`, openpyxl BOM, absolute rel
  targets), a seeded XML-tokenizer fuzz suite, and a happy-dom browser
  round-trip (202 total)

### Styles on read
- `readWorkbook(bytes, { styles: true })` resolves a `ReadStyle` onto each
  `ReadCell` — `font` (name / size / bold / italic / underline / colour),
  `fill`, `border` (per edge), `align`, `numFmt` — mirroring the writer's
  `CellStyle` so styles round-trip
- `src/style-read.ts` parses `styles.xml` fonts / fills / borders / `cellXfs`
  and resolves colours: `rgb` exactly, `indexed` via the standard palette,
  `theme` from `xl/theme/*` (the 0/1 & 2/3 swap, approximate HSL tint)
- built-in number-format ids resolve to their codes; `id 0` (`General`) and
  the sheet-default font/fill/border (`id 0`) are not surfaced as explicit style
- `xl/theme/*` added to the extraction allow-list; still **one runtime
  dependency**
- off by default (values-only path is unchanged and faster)
- 11 new tests: write → read style round-trips, indexed & theme colours (183 total)

## [0.2.0] - 2026-09-06

Adds `.xlsx` **reading**. Still **one runtime dependency** (`fflate`).

### `readWorkbook(bytes | ArrayBuffer, options?)`
- → `ReadWorkbook`: `sheetNames`, `sheets`, `sheet(name | index)`, `date1904`
- `ReadWorksheet`: `name`, `dimension`, `merges`, `cell(ref)`, `rows()` (sparse),
  `toArray()` / `values()` (rectangular)
- `ReadCell`: `{ ref, row, col, type, value, formula? }`
- Cell types: string (shared + inline), number, **date** (numeric cells with a
  date number-format → `Date`; `{ dates: false }` to opt out), boolean,
  **formula** (with `.formula` + cached `.value`), error, empty
- `serialToDate()` — inverse of `dateToSerial`, incl. the 1900 phantom leap day
  and 1904 mode; `isDateNumFmt()` classifies built-in and custom format codes
- `{ sheets: [name | index] }` loads a subset; positional fallback for writers
  that omit `r` attributes

### Security
- `src/xml-read.ts` — a strict ~200-line XML tokenizer, written **instead of a
  dependency** after a CVE review of `fast-xml-parser` (13 advisories, most in
  the DOCTYPE / entity-expansion class a reader is exposed to). It **rejects**
  `<!DOCTYPE>` / `<!ENTITY>` / `<![CDATA[>` / unknown entities outright, does no
  entity expansion, caps nesting depth, and guards keys against `__proto__`
  pollution
- `src/unzip.ts` — extracts only an allow-list of parts, enforces total and
  per-part uncompressed-size caps against the ZIP directory before *and* after
  inflation, ignores `..` traversal entries
- CI now round-trips **both directions** through LibreOffice (quire writes → LO
  reads; LO writes → quire reads)
- 54 new tests incl. write → read round-trips and a hostile-input suite (172 total)

### Not read (yet)
Cell styles on read (`0.3.0`), data validation, conditional formatting, images,
charts, pivot tables, hyperlinks, defined names, `.xls` / `.xlsb`.

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
