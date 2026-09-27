# @uekichinos/quire — reader plan (for review)

Adds `.xlsx` **reading** to quire, so it fully replaces `exceljs` for this
codebase (write **and** import). Lands as `0.2.0`.

Status: **not started.** XML-parser decision made (§2: hand-rolled, 0 deps).
5 open questions in §9 — answer those, then build Phase R1. ~2.5 weeks.

---

## 1. Goal & non-goals

### Goal

Parse `.xlsx` files produced by **mainstream tools** — Excel, Google Sheets,
LibreOffice Calc, `openpyxl`, `exceljs`, and quire itself — into a data model.
Read-first; the model lines up with quire's writer so write → read → write
round-trips.

### Non-goals (v1)

| Out | Why |
|---|---|
| `.xls` (BIFF), `.xlsb` | different formats |
| Corrupt / malformed-file recovery | best-effort parse, clear errors, no heroics |
| Streaming / constant-memory read of GB files | v1 loads into memory (with caps) |
| **Formula evaluation** | we return the formula string + the cached `<v>`, we don't compute |
| Images, charts, pivot tables, VBA, external workbook links | ignored, not errored |
| Every historical / vendor quirk | documented gaps; fix case-by-case when a real file breaks |

If a file needs something out of scope → use `@zurmokeeper/exceljs` for that
import path.

---

## 2. Dependencies — stays at **one** (`fflate`)

| Runtime | Role |
|---|---|
| `fflate` | already present — `Unzip` streaming API for size-capped decompression |

### XML parsing: hand-rolled minimal tokenizer (`src/xml-read.ts`, 0 deps)

Decided after a CVE review of the candidates:

- **`fast-xml-parser` — rejected.** 13 GitHub advisories (2023–2026), the majority
  in **DOCTYPE / entity-expansion on untrusted input** — exactly this threat
  model — including a critical (CVE-2026-25896) and repeated *"incomplete fix"*
  follow-ups (CVE-2026-33036, -73569). v5 also grew to 6 transitive deps of
  brand-new sub-packages.
- **`sax` / `saxes` / `txml`** — clean advisory records; kept as the fallback
  (`sax@1.6`, 0 deps, isaacs, longest clean history) if we ever decide not to
  own the tokenizer.
- **Hand-rolled — chosen.** OOXML parts are machine-generated, always
  well-formed, and never contain `<!DOCTYPE>` or custom entities — only the 5
  predefined ones. We accept a *tiny strict grammar* and **hard-error on anything
  unexpected**, which is a smaller attack surface than a general parser with its
  entity machinery. Keeps quire at **1 dependency**, aligned with the whole
  reason quire exists.

**Grammar accepted:** `<?xml …?>` decl, elements, attributes (both quote
styles), text, self-closing tags, the 5 predefined entities + `&#nn;` / `&#xnn;`
numeric char refs. **Rejected on sight:** `<!DOCTYPE`, `<!ENTITY`, `<![CDATA[`
(not used in the parts we read), processing instructions, anything malformed.

**Tokenizer self-hardening** (so it doesn't grow its own CVE):

- character scanning, **no regex** on the hot path → no ReDoS
- **iterative with an explicit stack** + nesting-depth cap → no stack overflow
- numeric char refs bounded to valid code points; no recursive entity expansion
- attribute keys and element names pass an `__proto__` / `constructor` /
  `prototype` guard before becoming object keys
- input already size-capped by the zip layer

Net: quire stays at **1 runtime dep**. (exceljs: ~50.)

---

## 3. Security — the whole reason reading is the risky half

Reading parses attacker-controlled input. Each mitigation is a test in Phase R4.

| Threat | Mitigation |
|---|---|
| **Decompression bomb** | `fflate` `Unzip` (streaming) with a running byte counter; abort past `limits.maxUncompressedBytes` (default 100 MB) and a per-entry cap. Reject before fully inflating. |
| **Path traversal** | only ever read a fixed allowlist of part names (`xl/worksheets/sheetN.xml`, …); entries with `..`, absolute paths, or backslashes are ignored. Nothing is written to disk. |
| **XXE / billion-laughs** | our tokenizer hard-errors on `<!DOCTYPE` / `<!ENTITY`; only the 5 predefined entities and bounded numeric char refs are decoded; nothing recursive. |
| **Prototype pollution** | element / attribute names pass an `__proto__` / `constructor` / `prototype` guard before becoming object keys; id→value maps use `Object.create(null)`. (The exact class that hit exceljs `deepMerge` and `fast-xml-parser` CVE-2023-26920.) |
| **Tokenizer ReDoS / stack overflow** | no regex on the hot path; iterative parse with an explicit stack and a nesting-depth cap. |
| **Resource exhaustion** | caps on sheets (`maxSheets` 256), total cells (`maxCells` 5 M), shared-string count, single-string length. O(1) id lookups (arrays / Maps) — no O(n²) resolution. |
| **Caller's job (documented)** | still run untrusted uploads inside a worker with an overall wall-clock + RSS limit. No parser removes that need. |

---

## 4. API

```ts
import { readWorkbook } from '@uekichinos/quire'

const wb = readWorkbook(bytes, options?)   // bytes: Uint8Array | ArrayBuffer

wb.sheetNames                 // string[]  (workbook order)
wb.sheets                     // ReadWorksheet[]
wb.sheet('Sales')             // ReadWorksheet | undefined   (name or index)

const s = wb.sheet(0)!
s.name
s.dimension                   // { rows: number, cols: number }
s.merges                      // string[]  e.g. ['A1:C1']
s.cell('B2')                  // ReadCell | undefined
for (const row of s.rows()) { // sparse: only rows that exist
  for (const cell of row) { … }  // ReadCell, indexed by column
}
s.toArray()                   // ReadCell[][], rectangular, undefined in the holes
s.values()                    // (string|number|boolean|Date|null)[][] — just values
```

```ts
interface ReadCell {
  ref: string                 // 'B2'
  row: number; col: number    // 1-based
  type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'error' | 'empty'
  value: string | number | boolean | Date | null
  formula?: string            // without the leading '='
  numFmt?: string             // resolved format code, when known
  style?: ReadStyle           // only when { styles: true }
}
```

```ts
interface ReadOptions {
  styles?: boolean            // resolve per-cell font/fill/border. Default: false (faster)
  dates?: boolean             // number + date-format → Date. Default: true
  sheets?: (string | number)[]  // subset to load. Default: all
  date1904?: boolean          // override; auto-detected from workbook.xml otherwise
  limits?: { maxUncompressedBytes?: number; maxCells?: number; maxSheets?: number }
}
```

Design choices:

- **Values-only by default; `{ styles: true }` opts into style resolution.** Most
  imports want values; parsing `styles.xml` and resolving every `xf` is extra
  work you shouldn't pay for by default.
- **Sparse iteration**, with `toArray()` / `values()` for the rectangular view.
- **Numbers stay numbers**; a number becomes a `Date` only when its cell format
  is a date format (built-in ids 14–22 / 45–47, or a custom code whose tokens
  are date/time). `{ dates: false }` disables.
- **Formulas** return `formula` + the cached `value` from `<v>`. No evaluation.
- **Errors** (`#DIV/0!`, …) → `type: 'error'`, `value` is the error string.
- Shared **and** inline strings both surface as `type: 'string'`.

Open question: this accessor API, **or** a flat snapshot
`readWorkbook(bytes) -> { sheets: { name, rows: Value[][] }[] }` (simpler, less
capable). Leaning accessor.

---

## 5. Parts parsed

| Part | Read | Ignore |
|---|---|---|
| `xl/workbook.xml` | sheet names + order, `date1904`, (optionally) defined names | views, calcPr |
| `xl/_rels/workbook.xml.rels` | rId → sheet file | other rels |
| `xl/sharedStrings.xml` | `<si>` → string (concatenate `<r>` runs; keep `xml:space`) | run formatting, phonetic |
| `xl/styles.xml` | `numFmts`, `cellXfs → numFmtId`; + fonts/fills/borders when `styles:true` | dxfs, tableStyles |
| `xl/worksheets/sheetN.xml` | `sheetData` (`r`, `t`, `s`, `v`, `f`), `mergeCells`, `dimension`; + `cols` when `styles:true` | drawings, pageSetup, sheetPr |
| media / charts / pivot / vba / drawings | — | everything |

---

## 6. Tricky bits (each gets a test)

1. **`<si>` run concatenation** — `<t>` or many `<r><t>`; concat, respect
   `xml:space="preserve"`, never trim.
2. **Date detection** — build `s (xf index) → numFmtId → formatCode`, classify.
   Built-in 14–22 / 45–47 are date/time; custom codes via a token scan
   (`y m d h s` outside quotes/`[...]`).
3. **Serial → Date** — inverse of the writer's `dateToSerial`, incl. the 1900
   phantom leap day and 1904 mode. Extend `datetime.ts`.
4. **Missing `r` attributes** — some writers omit `<c r>` / `<row r>` and rely on
   order. v1 **handles it** with a positional fallback (cheap, big compat win).
5. **Empty styled cells** — `<c r="C1" s="2"/>` → `type:'empty'`, `value:null`,
   `style` present.
6. **Booleans** — `t="b"`, `<v>` is `1` / `0`.
7. **Case / BOM** — match known part names case-insensitively; tolerate a BOM.
8. **Number precision** — `parseFloat` and leave it; don't "correct" it.
9. **Huge shared-string tables** — array index, O(1).

---

## 7. Testing

1. **Round-trip vs the writer** — `readWorkbook(createWorkbook()…xlsx())`
   reproduces values / types / merges (and styles in R3). Both halves are ours,
   so this is free and thorough.
2. **Multi-tool fixture corpus** — small `.xlsx` committed, one each from Excel,
   Google Sheets, LibreOffice, `openpyxl`, `exceljs`; assert known cells.
3. **Security fixtures** — zip bomb, `../` entries, `<!DOCTYPE>` billion-laughs,
   `__proto__` in a sheet name / defined name → rejected or neutralised, and
   `Object.prototype` stays clean.
4. **Malformed** — truncated zip, missing `workbook.xml`, sheet → missing rId,
   junk `<v>` → clear errors, never a crash.
5. **CI** — extend the LibreOffice job: write with quire → convert with
   LibreOffice → read the result back with quire → compare.

---

## 8. Phases (~2.5 weeks)

### R1 — tokenizer + plumbing + values ✅ done
`src/xml-read.ts` (strict tokenizer + hardening tests), `src/unzip.ts` (allow-list
+ size caps + traversal guard), `src/read.ts`. `readWorkbook()` → accessor model.
String (shared + inline), number, boolean, **formula (+ cached value)**, error,
and empty cells; `merges`; `dimension`; positional fallback for missing `r`;
`{ sheets: [...] }` subset. 36 tests incl. round-trip + hostile input.
(Formulas / errors / inline / merges were pulled forward from R2.)

### R2 — dates ✅ done
`serialToDate()` (1900 + phantom leap day, 1904); `isDateNumFmt()` (built-in ids
+ custom-code token scan, ignoring quoted/bracketed literals); `styles.xml`
`numFmts` + `cellXfs → numFmtId` map; numeric cells with a date format → `Date`;
`{ dates: false }` opt-out; `'date'` added to `ReadCellType`. Round-trips vs the
writer.

### R4 — ship ✅ done — **published as `0.2.0`**
`examples/read.mjs`; README reader section; `CHANGELOG` `0.2.0`; CI round-trips
**both directions** through LibreOffice (quire writes → LO reads; LO writes →
quire reads). 172 tests.

### R3 — styles on read ✅ done (unreleased, ships as `0.3.0`)
`src/style-read.ts`: parses `styles.xml` fonts / fills / borders / `cellXfs`;
`{ styles: true }` resolves a `ReadStyle` (font / fill / border / align / numFmt)
per cell. Colours: `rgb` exact, `indexed` via the standard palette, `theme` from
`xl/theme/*` with the 0/1 & 2/3 swap and an approximate tint. Sheet-default ids
(0) not surfaced. Round-trips vs the writer. 11 tests, 183 total.
**Deferred to a later minor:** column- and row-level styles on read.

### R5 — hardening ✅ done (unreleased, ships with `0.3.0`)
Fixes found in a self-review of the reader/writer:
- **writer** — `xlsx()` made idempotent (`sst`/`pool` built per serialise, not
  accumulated on the worksheet); `numToXml()` expands JS exponent notation so
  large / tiny numbers write as plain decimals
- **unzip** — streaming inflate (`Unzip` + `UnzipInflate`) with a running byte
  budget, aborts mid-archive on a cap instead of decompressing everything first
- **read limits** — `{ limits: { maxCells, maxSheets } }` on top of the byte
  caps; `readWorkbookAsync()` yields between sheets; `values({ ragged: true })`
- **errors** — single `QuireError` base for every throw (`XlsxReadError`,
  `XmlError` extend it), all exported
- **`ReadCell.numFmt`** surfaced unconditionally (not only under `{ styles }`)
- **tests** — `numToXml` unit, idempotency, real-world-quirk fixture corpus,
  seeded tokenizer fuzz suite, happy-dom browser round-trip. 202 total.

### R6 — hyperlinks ✅ done (unreleased)
Write: `{ hyperlink, text?, tooltip? }` cell value → `<hyperlinks>` +
`_rels/sheetN.xml.rels` (external `TargetMode`), 2,079-char cap. Read:
`ReadCell.hyperlink` resolved via the worksheet's own relationships or a bare
`location` (in-workbook links); range refs attach to the top-left cell;
dangling relationship ids are ignored. Also fixed: formula cells with a date
number-format now read back as `Date` (previously always coerced to a plain
number). 13 new tests.
**Deferred:** defined names / named ranges, column/row-level styles on read,
rich text runs.

### R7 — defined names + column/row default styles ✅ done (unreleased)
Write: `wb.defineName(name, sheetName, range)` — workbook-scoped only, Excel
identifier rules enforced, de-duplicated case-insensitively; emits into
`<definedNames>` alongside the existing internal `_xlnm._FilterDatabase`
entries. Read: `wb.definedNames`, excluding `_xlnm.*` names; resolves
`sheetName`/`range` for a simple single-sheet reference, else raw `refersTo`.
Also: the writer now stamps `<row customFormat="1" s="…">` when a row has a
style (previously only baked into each populated cell's own `s`, invisible to
any cell without a `<c>` element) — closing the gap that made column/row
default styles unreadable from quire's own output. Reader gains
`sheet.columnStyles` / `sheet.rowStyles` (`Map<number, ReadStyle>`) resolving
`<col style>` / `<row customFormat s>` under `{ styles: true }`, bounded to
16,384 column-style writes per sheet regardless of how many `<col>` elements a
hostile file contains. 12 new tests.
**Still deferred:** sheet-scoped defined names, rich text runs.

### R8 — data validation, rich text, comments ✅ done (unreleased)
Three additions picked as the "most used or interesting" of the remaining
exceljs-parity gaps:
- **Data validation** — write `sheet.setDataValidation(range, { list, ... })`
  (`type="list"` only: inline choices or a range reference); read
  `sheet.dataValidations`, filtered to `type="list"` rules
- **Rich text** — a cell value can be `RichTextRun[]`, interned through the
  same shared-string table as plain strings; reading collapses a single
  unformatted run back to a plain string so the common case is unaffected,
  but keeps a single *formatted* run as `ReadRichTextRun[]` (it carries data a
  plain string can't)
- **Comments** — write emits classic (legacy) comments: `xl/commentsN.xml` +
  a VML drawing part for the indicator/popup box, wired through the
  worksheet's own rels + `<legacyDrawing>`; multiple distinct authors are
  interned per sheet. Read resolves `ReadCell.comment` via the worksheet's
  `comments`-typed relationship (found by type, not by a fixed part name,
  since real files don't necessarily number `commentsN.xml` by sheet index)

**Bug fixed along the way:** a hyperlink or comment on a cell with *no* `<c>`
element at all (a genuinely empty cell) was silently dropped on read — both
now materialise a placeholder `empty` cell so the annotation isn't lost.

**Deferred:** non-`list` validation types, conditional formatting, threaded
(modern) comments, sheet-scoped defined names.

### R9 — extended data validation, conditional formatting, streaming writer ✅ done (unreleased)
Picked as the next "most used or interesting" three from the remaining gap:
- **Extended data validation** — `setDataValidation(range, rule, options?)`'s
  `rule` became a discriminated union (`list` unchanged; `whole`/`decimal`/
  `textLength`/`date` add `operator` + `value`, a single number/`Date` or a
  `[min, max]` pair for `between`/`notBetween`). Read resolves the same set,
  `values` (plural) replacing the single-formula shape
- **Conditional formatting** — `addConditionalFormat(range, rule)`: `cellIs`
  (highlight on a condition, styled via a new dxf pool in `styles.xml`
  mirroring cellXfs but simpler — no "id 0 is default" convention, inline
  font/fill instead of an indexed lookup) and `colorScale` (2/3-stop gradient,
  inline colours, no dxf). Read adds a `dxfStyle` pool to `style-read.ts` and
  resolves `sheet.conditionalFormats`, skipping unsupported rule kinds
  (dataBar, iconSet, top10, custom formula)
- **Streaming writer** (`src/streaming.ts`, `createStreamingWorkbook()`) —
  the biggest of the three: renders and DEFLATEs each row via fflate's
  streaming `Zip`/`ZipDeflate` as it's added, instead of holding the whole
  sheet in memory. One pass, forward-only (no `setCell`/`setRow`;
  `setColumn`/`freeze` must precede the first `addRow()` on that sheet since
  they render into the already-flushed header) but otherwise full parity
  with the buffered writer — hyperlinks, rich text, merges, data validation,
  conditional formatting, comments, `defineName` all supported, since none of
  them depend on row buffering (they're collected as metadata and rendered
  into the tail after all rows are flushed). Shares its cell/row/tail
  rendering with the buffered writer via extracted pure functions
  (`renderCellXml`, `colsXmlOf`, `dataValidationsXmlOf`,
  `conditionalFormatsXmlOf`, `commentsXmlOf`, `vmlDrawingXmlOf`,
  `buildRelsAndRefsOf`) — a pure refactor verified behavior-identical before
  the streaming class was built on top of it, so the two writers can't drift

**Deferred:** everything R8 deferred, plus sheet-scoped conditional formats,
a streaming *reader*.

### R10 — data bars, print setup, streaming reader ✅ done (unreleased)
The next "most used or interesting" three — the streaming *reader* R9 had
just deferred, plus two smaller, contained additions:
- **Data bars** — `addConditionalFormat(range, { type: 'dataBar', color })`,
  read/write. Extends R9's conditional-formatting union with a third variant;
  no new pool needed (colour is inline in the rule, like `colorScale`)
- **Print setup** — `setPageSetup({ orientation?, paperSize?, fitToWidth?,
  fitToHeight?, scale?, margins?, printArea? })`. `<sheetPr>` (fit-to-page
  flag) needed a new *header*-position slot ahead of `<dimension>`;
  `<pageMargins>`/`<pageSetup>` a new *tail* slot between `<hyperlinks>` and
  `<legacyDrawing>`. `printArea` reuses the existing defined-names machinery
  (a `_xlnm.Print_Area` entry per sheet, same shape as autoFilter's
  `_xlnm._FilterDatabase`) rather than inventing new plumbing
- **Streaming reader** (`readRows()`) — the honest scope, worked out
  carefully: the hand-rolled tokenizer (`xml-read.ts`) is a single
  synchronous pass with no pause/resume capability, so a *true* async,
  incrementally-decompressing streaming reader isn't achievable without
  rewriting the tokenizer into a resumable coroutine — out of scope for this
  round. What *is* achievable and still genuinely valuable: `parseSheet` now
  accepts an `onRow` callback and, when set, never populates the whole-sheet
  `cells` map — each row is buffered, handed to the callback, and discarded,
  so peak memory scales with row width instead of row count. Still one
  synchronous parse pass (documented plainly, same honesty bar as
  `readWorkbookAsync`'s single-sheet caveat) — this is a memory optimization,
  not a concurrency one. Hyperlinks/comments are skipped in this mode since
  resolving them needs a second pass over the sheet's cells, which is exactly
  what streaming mode avoids holding

**Deferred:** a true async/incremental streaming reader (would need a
resumable tokenizer), sheet-scoped conditional formats, data-bar gradient
customization (min/max always the range's own, no custom `cfvo` thresholds).

---

## 9. Questions for you

- ✅ **XML parsing** — decided: hand-rolled tokenizer, 0 deps, `sax` as the
  fallback (see §2).

Still open:

1. **Styles on read** — opt-in via `{ styles: true }` (recommended) or always?
2. **API** — the `sheet.cell()/rows()` accessor, or a flat
   `{ sheets: [{ name, rows }] }` snapshot?
3. **Version** — land as `0.2.0`?
4. ✅ **Defined names / named ranges** — decided: a later minor, shipped in
   R7 (workbook-scoped only; sheet-scoped still deferred).
5. ✅ **Hyperlinks on read** — decided: a later minor, shipped in R6.
