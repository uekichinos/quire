# Changelog

All notable changes to `@uekichinos/quire` are documented here.

## [Unreleased]

### Streaming reader
- `readRows(bytes, sheetNameOrIndex, onRow, options?)` — streams one sheet's
  rows through `onRow` as each is parsed, instead of materialising the whole
  sheet in memory like `readWorkbook` does. Peak memory scales with row
  *width*, not row *count*. Sheet-level metadata (`merges`, `columnStyles`/
  `rowStyles`, `dataValidations`, `conditionalFormats`, `pageSetup`) still
  comes back in the return value
- Two documented limits: hyperlinks/comments aren't attached to streamed rows
  (resolving them needs a second pass over the cells, which this mode avoids
  holding); still one synchronous parse pass under the hood (lower memory,
  not lower blocking) — the hand-rolled tokenizer isn't interruptible

### Print setup (read + write)
- Write: `sheet.setPageSetup({ orientation?, paperSize?, fitToWidth?,
  fitToHeight?, scale?, margins?, printArea? })` — `<sheetPr><pageSetUpPr
  fitToPage="1"/></sheetPr>` emitted only when fit-to-page is used; margins
  default to Excel's own (0.7/0.7/0.75/0.75/0.3/0.3in); `printArea` becomes
  the sheet's `_xlnm.Print_Area` workbook-level defined name. In the
  streaming writer, must be called before that sheet's first `addRow()`
- Read: `sheet.pageSetup` resolves the same shape from `<pageSetup>`/
  `<pageMargins>` plus the sheet's own `_xlnm.Print_Area`

### Conditional formatting — data bars
- Write: `addConditionalFormat(range, { type: 'dataBar', color })`
- Read: `sheet.conditionalFormats` now also resolves `dataBar` rules

### Streaming writer
- `createStreamingWorkbook()` — a low-memory writer for very large sheets:
  each row is rendered and DEFLATEd (via `fflate`'s streaming `Zip`/
  `ZipDeflate`) as soon as it's added, instead of being held in memory (as
  cell objects, then as one big XML string) for the workbook's lifetime
- Supports everything the buffered writer does — hyperlinks, rich text,
  merges, freeze panes, auto-filter, data validation, conditional formatting,
  comments, `defineName` — one pass, forward-only: no `setCell`/`setRow`
  random access; `setColumn`/`freeze` must be called before that sheet's
  first `addRow()` (they render into the header, flushed immediately).
  Multiple sheets can be written interleaved or in any order
- `finish(): Promise<Uint8Array>` — same byte format `xlsx()` produces
- Shares its cell/row/tail rendering with the buffered writer via a set of
  extracted pure functions (`renderCellXml`, `colsXmlOf`,
  `dataValidationsXmlOf`, `conditionalFormatsXmlOf`, `commentsXmlOf`,
  `vmlDrawingXmlOf`, `buildRelsAndRefsOf`) so the two can't drift on what a
  given value renders as — verified by a same-input-same-output test

### Data validation (read + write) — extended beyond lists
- Write: `sheet.setDataValidation(range, rule, options?)` — `rule` is now a
  discriminated union: `{ type: 'list', list }` (unchanged), or `{ type:
  'whole' | 'decimal' | 'textLength' | 'date', operator, value }` with
  `value` a single number/`Date` or a `[min, max]` pair for
  `between`/`notBetween`; shared options (`allowBlank`, prompt/error text)
  moved to a separate 3rd argument
- Read: `sheet.dataValidations` now also surfaces `whole`/`decimal`/`date`/
  `textLength` rules (previously `list` only), resolving `operator` and
  `values` (numbers, or `Date`s for `type: 'date'`)

### Conditional formatting (read + write)
- Write: `sheet.addConditionalFormat(range, rule)` — `{ type: 'cellIs',
  operator, formula, style }` (font bold/italic/color + fill, via a new
  differential-format ("dxf") pool in `styles.xml`) or `{ type: 'colorScale',
  colors }` (2 or 3 stops, thresholds at the range's own min/mid/max)
- Read: `sheet.conditionalFormats` lists `cellIs` and `colorScale` rules as
  `{ ref, rule }`; other rule kinds (data bars, icon sets, top/bottom N,
  custom formula) are skipped. `cellIs.style` resolves under
  `{ styles: true }` via a new `dxfStyle` pool in `style-read.ts`

### Rich text (read + write)
- A cell value can be `RichTextRun[]` — `[{ text, font? }, …]` — for multiple
  differently-formatted spans in one string. Interned through the same
  shared-string table as plain strings (deduped by run content)
- Read: a shared-string entry with more than one run, or one formatted run,
  comes back as `ReadRichTextRun[]`; a single unformatted run still collapses
  to a plain `string` so the common case is unaffected

### Comments (read + write)
- Write: `sheet.setComment(ref, text, { author? })` — classic (legacy)
  comments: `xl/commentsN.xml` + a VML drawing (`xl/drawings/vmlDrawingN.vml`)
  for the indicator triangle/popup, wired via the worksheet's own rels and
  `<legacyDrawing>`. Independent of the cell's value — works on an empty cell
- Read: `ReadCell.comment` → `{ text, author }`, resolved via the worksheet's
  `comments`-typed relationship; also fixed to attach to a cell with no `<c>`
  element at all instead of silently dropping (same fix applied to hyperlinks)
- `xl/commentsN.xml` added to the read allow-list

## [0.4.0] - 2026-09-26

Hyperlinks, defined names, and column/row default styles — read and write.
Plus a formula/date-read fix. Still **one runtime dependency** (`fflate`).

### Defined names (read + write)
- Write: `wb.defineName(name, sheetName, range)` — a workbook-scoped named
  range (single cell or `A1:C3`-style range). Validated against Excel's
  identifier rules (can't look like a cell reference, can't use the
  `_xlnm.` prefix) and de-duplicated case-insensitively
- Read: `wb.definedNames` lists them, excluding Excel-internal names like
  `_xlnm._FilterDatabase`; each entry resolves `sheetName`/`range` when the
  reference is a simple single-sheet range, else keeps `refersTo` raw

### Column/row default styles
- Writer: `setRow(i, { style })` now also stamps the `<row customFormat="1"
  s="…">` attributes (not just each populated cell's own `s`), so empty
  cells in a styled row show the right formatting in Excel too — closes a
  gap where the style was invisible on any cell without its own `<c>`
- Reader: `sheet.columnStyles` / `sheet.rowStyles` (`Map<number, ReadStyle>`,
  1-based) resolve `<col style>` / `<row customFormat s>` under `{ styles:
  true }` — what an empty cell in that column/row would look like

### Hyperlinks (read + write)
- Write: `{ hyperlink, text?, tooltip? }` as a cell value — `text` defaults to
  the URL, `tooltip` is the hover text. Emits a `<hyperlinks>` block in the
  worksheet plus a `_rels/sheetN.xml.rels` part (external `TargetMode`);
  re-setting the cell to a non-hyperlink value drops it. Targets capped at
  2,079 characters (Excel's practical limit)
- Read: `ReadCell.hyperlink` → `{ target, tooltip? }`, resolved from the
  worksheet's own relationships (external) or a bare `location` attribute
  (in-workbook links); a range `ref` attaches to its top-left cell; a dangling
  relationship id is ignored rather than thrown
- `xl/worksheets/_rels/*.xml.rels` added to the read allow-list

### Fixed
- **Formula cells with a date number-format now read back as `Date`** — the
  cached result was always coerced to a plain number regardless of style,
  unlike the equivalent non-formula numeric cell

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
