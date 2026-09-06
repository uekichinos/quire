# @uekichinos/quire — design plan (for review)

A small, dependency-light **`.xlsx` writer** to replace `exceljs` for the common
"export a styled spreadsheet" case.

Status: **scaffold created, nothing implemented.** Review this document, then I
build Phase 1.

---

## 1. Name

**`@uekichinos/quire`** — a *quire* is a set of sheets of paper. One evocative
word, on brand with `browser-gate` / `counter` / `sentinel` / `stash`, and it
literally means "sheets".

Alternatives if you'd rather: `@uekichinos/xlsx` (best for npm discoverability),
`@uekichinos/ledger`, `@uekichinos/folio`. Trivial to rename now — say the word.

---

## 2. Goal & non-goals

### Goal

Generate valid `.xlsx` files that open cleanly in Excel, LibreOffice, Numbers and
Google Sheets, covering what most apps need for a data export:

- Multiple worksheets
- Typed cells: string, number, boolean, `Date`, formula (with optional cached result), `null`
- Cell styling via a **deduplicated style pool**: font (name/size/bold/italic/underline/colour), solid fill, borders per edge, alignment (h/v, wrap, indent), number format (code or built-in id)
- Merged cell ranges
- Column widths + hidden columns + per-column default style
- Row heights + per-row default style
- Freeze panes, auto-filter over a range
- Workbook metadata (title, creator, created date)
- Output as `Uint8Array` and `Blob`; identical behaviour in Node and browser

### Non-goals (v1 — explicitly out)

| Out | Why |
|---|---|
| **Reading `.xlsx`** | The entire reason to build this. A robust reader is a multi-month, never-finished job. Use `@zurmokeeper/exceljs` if you need to parse. |
| Images, charts, drawings | Large spec surface; rare in exports |
| Pivot tables | Very large spec surface |
| Data validation, conditional formatting | Deferred to a later minor; not hard, just scope |
| Rich text (multi-run cells) | Deferred |
| Streaming writer (constant-memory) | Separate, harder mode; v1 builds in memory |
| `.xls` (BIFF) / `.xlsb` | Different formats entirely |
| Themes / custom colour schemes | We emit a fixed default theme |
| Shared formulas, array formulas | v1 writes each formula independently |

If a consumer needs an out-of-scope feature, that is a signal to use ExcelJS,
not to grow `quire`.

---

## 3. Public API

Full types live in `src/index.ts` already. Shape:

```ts
import { createWorkbook } from '@uekichinos/quire'

const wb = createWorkbook()

const sheet = wb.addWorksheet('Sales', {
  columns: [{ width: 24 }, { width: 12, style: { numFmt: '#,##0.00' } }],
  freeze: { ySplit: 1 },
})

sheet.addRow(['Product', 'Revenue'], { style: { font: { bold: true } } })
sheet.addRow(['Widget', 15003.4])
sheet.addRow(['Gadget', 2450])
sheet.setCell('B5', { formula: 'SUM(B2:B4)' }, { font: { bold: true } })
sheet.merge('A7:B7')
sheet.setCell('A7', new Date(), { numFmt: 'yyyy-mm-dd' })

const bytes = wb.xlsx()          // Uint8Array
// browser:  wb.blob()  ->  Blob
// node:     fs.writeFileSync('out.xlsx', bytes)
```

Design choices:

- **Builder pattern, chainable.** No getters into a live model; you build then
  serialise. Keeps the model write-optimised and the code small.
- **Styles are plain objects, deduped internally.** Callers never touch style
  indices. Passing the same style object (by value) twice reuses one `cellXfs`
  entry.
- **Colours** accept `'RRGGBB'` or `'AARRGGBB'`; normalised to ARGB internally.
- **Errors are thrown eagerly** — bad ref (`'A0'`), unknown border style, NaN
  cell value — so you never ship a "needs repair" file silently.

Open question: builder vs. a declarative `writeXlsx({ sheets: [...] })` one-shot.
I lean builder (matches ExcelJS muscle memory, easier partial writes). Tell me if
you want the declarative form instead or as well.

---

## 4. Architecture

```
caller → Workbook/Worksheet builder (in-memory model)
       → StylePool.intern()  (dedupe fonts/fills/borders/numFmts/xfs → indices)
       → serialisers (one per XML part, plain string templating)
       → fflate.zipSync()  (STORE for tiny parts, DEFLATE for sheets)
       → Uint8Array
```

- **No XML library.** Writing XML is string concatenation + a single
  `escapeXml()` that also strips characters illegal in XML 1.0
  (`\x00-\x08 \x0B \x0C \x0E-\x1F`).
- **StylePool** is the one piece of real cleverness: four sub-pools
  (`numFmts`, `fonts`, `fills`, `borders`) each keyed by a canonical string, plus
  a `cellXfs` pool keyed by the tuple of sub-pool indices + alignment. Every cell
  stores a single `s` index.
- **A1 <-> {row,col}** conversion and column-letter maths adapted from ExcelJS
  (`lib/utils/col-cache.js`) — small, pure, battle-tested.
- **Dates**: `Date` → Excel serial number (days since 1899-12-30, accounting for
  the fictitious 1900-02-29). 1900 date system only in v1.

---

## 5. XML parts emitted

| Part | Always? | Notes |
|---|---|---|
| `[Content_Types].xml` | yes | overrides per sheet |
| `_rels/.rels` | yes | → workbook, docProps |
| `xl/workbook.xml` | yes | sheet list, defined names (for auto-filter), `calcPr` |
| `xl/_rels/workbook.xml.rels` | yes | → sheets, styles, sharedStrings |
| `xl/styles.xml` | yes | the style pool |
| `xl/sharedStrings.xml` | if any string cells | v1 uses shared strings (smaller for typical exports); inline-string fallback behind an option |
| `xl/worksheets/sheetN.xml` | one per sheet | `dimension`, `cols`, `sheetData`, `mergeCells`, `autoFilter`, `sheetViews`/`pane` |
| `docProps/core.xml`, `docProps/app.xml` | yes | Excel is happier with them present |

Deliberately **not** emitting: `theme/theme1.xml` (Excel supplies a default when
absent for our feature set — to be verified in Phase 1), `calcChain.xml`
(optional; we set `fullCalcOnLoad`).

---

## 6. Module layout

```
src/
  index.ts            public API + types (exists)
  workbook.ts         Workbook builder
  worksheet.ts        Worksheet builder + in-memory row/cell store
  style-pool.ts       dedup pools → indices
  serialize/
    content-types.ts
    rels.ts
    workbook-xml.ts
    styles-xml.ts
    shared-strings.ts
    worksheet-xml.ts
    doc-props.ts
  xml.ts              escapeXml, attr helpers
  address.ts          A1 <-> {row,col}, colLetter(n) / colNumber(s)   [adapted from ExcelJS]
  datetime.ts         Date <-> serial, 1900 leap quirk                [adapted from ExcelJS]
  numfmt.ts           built-in id table, code validation              [adapted from ExcelJS]
  zip.ts              thin fflate wrapper (choose STORE vs DEFLATE)
  types.ts            shared internal types
  __tests__/
```

Target: **~1,500–2,000 LOC**, ~15 files.

---

## 7. Dependencies

| Runtime | Why | Size |
|---|---|---|
| `fflate` ^0.8 | zip container | ~30 kB min, 0 deps |

That's the whole tree. (`fflate` is bundled into the IIFE `<script>` build; kept
external for ESM/CJS.)

Dev: `tsup`, `typescript`, `vitest` — same as every `@uekichinos/*` package. CI
also installs a headless LibreOffice for the round-trip test (see §9).

vs. `exceljs@4.4.0`: 9 direct / ~50 transitive runtime deps → **1 / 1**.

---

## 8. What we take from ExcelJS (MIT)

Adapted (copied + trimmed + retyped), with attribution in `NOTICE`:

- column/address helpers — `lib/utils/col-cache.js`
- date ↔ serial incl. 1900 leap-year handling
- the built-in number-format id table

Reference only (read to learn the XML shapes, then write our own leaner
serialiser): `lib/xlsx/xform/**`, `lib/doc/workbook.js`, `worksheet.js`.

Their `spec/` fixtures are reused as an edge-case checklist.

`NOTICE` and the MIT text are already in place. `package.json` `license` stays
`MIT`; we add a `LICENSE` file in Phase 1.

---

## 9. Milestones

### Phase 1 — "hello workbook" ✅ done

- `zip.ts` (fflate wrapper, fixed mtime), `xml.ts` (escape + strip illegal chars),
  `address.ts` (A1 ↔ row/col, adapted from ExcelJS)
- `workbook.ts`: `createWorkbook()` → chainable `Workbook` / `Worksheet`;
  `addRow` with string / number / boolean cells; `null`/`undefined` skipped
- `serialize.ts`: all mandatory parts + shared strings (deduped, correct
  `count`/`uniqueCount`), `dimension` computed, doc props
- `xlsx()` → `Uint8Array` (deterministic), `blob()` → `Blob`
- 50 unit + structural tests (unzip output, assert on the XML)
- `examples/hello.mjs`
- **Gate — OUTSTANDING:** open `examples/hello.xlsx` in Excel (Windows),
  LibreOffice and Google Sheets and confirm no "needs repair". No LibreOffice on
  the dev box; needs a manual check before Phase 2. CI LibreOffice job lands in
  Phase 4.

### Phase 2 — types & styles ✅ done (pending Excel spot-check)

- `datetime.ts` (Date → serial, 1900 + phantom leap day), `numfmt.ts` (built-in
  id table, `DEFAULT_DATE_FORMAT`)
- `Date` and formula cells (`{ formula, result? }`, `t="str"`/`t="b"`)
- `style.ts` (`CellStyle` shape, colour normalise, merge) + `style-pool.ts`
  (dedup fonts/fills/borders/numFmts/cellXfs → indices, renders `styles.xml`)
- `addRow(values, { style })` row-level; `{ value, style }` per-cell; cell
  merges over row one level deep
- `worksheet.setCell(ref, value, style?)`
- 36 new tests (86 total)
- **Gate:** open the updated `examples/hello.xlsx` (styled header, money format,
  dates, formulas, top border, wrapText) in Excel — confirm no repair prompt.
  Column-per-style deferred to Phase 3 (arrives with column widths).

### Phase 3 — layout features ✅ done (pending Excel spot-check)

- `merge(range)` (guards: single-cell, backwards, overlap)
- `setColumn(i, { width, hidden, style })`; column style resolves as
  column < row < cell
- `freeze({ xSplit, ySplit })`; `autoFilter(range)` + `_FilterDatabase`
  defined name
- row height via `addRow(..., { height })` / `setRow(i, opts)`
- `addWorksheet(name, { columns, freeze, autoFilter })`
- schema-correct child order
- 20 new tests (106 total)
- **Gate:** open updated `examples/hello.xlsx` (merged title, frozen 2 rows,
  column widths, filter on the header, tall rows) in Excel.
- Still deferred to a later minor: `blob()` streaming, inline-string option.

### Phase 4 — polish & docs ✅ done

- README: real API reference, scope + performance notes
- `CHANGELOG` collapsed into `[0.1.0]`; `package.json` → `0.1.0`
- `LICENSE` (MIT) added; `NOTICE` covers the adapted ExcelJS portions
- `robustness.test.ts` — error paths (bad refs, NaN, invalid/pre-1900 dates,
  bad colours), hostile-input escaping/stripping across every part, ZIP
  signature, determinism, a linear-scaling guard
- fixed an O(n²) row-key dedup found by the scale test (100k rows: 9s → 1.1s)
- `scripts/roundtrip.mjs` + a CI job that installs LibreOffice, converts the
  output and asserts the values survived
- **118 tests** + 3 planned

**Ready for `0.1.0`.** Remaining before publish: §12 setup (git init, GitHub
repo, `NPM_TOKEN`).

### Later minors (not scheduled)

data validation, conditional formatting, rich text, `date1904`, a streaming
writer, then — only if genuinely needed — a *minimal* reader for round-tripping
our own output.

---

## 10. Testing strategy

1. **Unit** — `address`, `datetime`, `numfmt`, `style-pool` dedup, `escapeXml`.
2. **Structural** — unzip our output in the test, assert on the XML strings
   (parts present, `dimension` correct, `s` indices resolve, shared-string
   table).
3. **Round-trip in CI** — GitHub Actions installs `libreoffice-calc`, converts
   every fixture to CSV, diffs against expected. If LibreOffice accepts it, Excel
   will.
4. **Golden files** — commit `.xlsx` fixtures; fail on unexpected byte-level
   diff (after normalising zip mtime).
5. **Manual matrix once per phase** — open key fixtures in Excel (Win),
   LibreOffice, Numbers, Google Sheets; record results in the PR.

---

## 11. Risks / the hard parts

| Risk | Mitigation |
|---|---|
| "File needs repair" from a subtly wrong part | LibreOffice round-trip gate from Phase 1; build parts incrementally |
| Style-pool indexing bugs → wrong or missing formatting | dedicated dedup unit tests; structural test resolves every `s` index |
| Number-format built-in ids (0–163 implicit) vs custom (≥164) | port ExcelJS's table; only emit `<numFmt>` for custom codes |
| Column width unit (character-based, not px) | document clearly; provide the ExcelJS formula; don't pretend it's pixels |
| 1900 leap-year / timezone off-by-one on dates | convert in UTC; unit tests around 1900-02-28/29, 1901-01-01, epoch |
| XML-illegal control chars in user strings → corrupt file | central `escapeXml` strips them; test with ` `, emoji, CJK |
| Big workbooks blow memory | documented limitation; Phase 4 perf smoke test; streaming writer is a later, separate mode |
| Scope creep toward "just add reading" | non-goals table above is the contract |

---

## 12. Setup prerequisites (before first publish)

`packages/quire/` is scaffolded and joins the pnpm workspace automatically. Like
the sibling packages it is its **own git repo + npm package**, so before
`0.1.0`:

- [ ] `git init` in `packages/quire`, first commit, create `github.com/uekichinos/quire`
- [ ] add the `NPM_TOKEN` repo secret (granular, "bypass 2FA", scoped `@uekichinos/*`)
- [ ] `pnpm install` at the monorepo root to link `fflate`
- [ ] confirm the package name (`quire` vs `xlsx` vs …)

CI (`ci.yml`) and the tag-triggered publish (`publish.yml`) are already copied
from `sentinel`.

---

## 13. Questions for you

1. **Name** — `quire`, or `xlsx` / `ledger` / `folio`?
2. **API** — builder (proposed) or a one-shot `writeXlsx({...})`, or both?
3. **Scope** — anything in the non-goals table you actually need in `0.1.0`
   (most likely candidates: data validation, conditional formatting)?
4. **Shared vs inline strings as the default** — shared (smaller files) proposed;
   inline is simpler and diff-friendlier. Preference?
5. **Node floor** — keep `>=18`? (needs `Uint8Array`, `TextEncoder` — both fine
   on 18.)
6. Do you want a `LICENSE` file added now (MIT) to match, or is `package.json`
   `license` enough as elsewhere in the monorepo?
