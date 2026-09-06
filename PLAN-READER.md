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

---

## 9. Questions for you

- ✅ **XML parsing** — decided: hand-rolled tokenizer, 0 deps, `sax` as the
  fallback (see §2).

Still open:

1. **Styles on read** — opt-in via `{ styles: true }` (recommended) or always?
2. **API** — the `sheet.cell()/rows()` accessor, or a flat
   `{ sheets: [{ name, rows }] }` snapshot?
3. **Version** — land as `0.2.0`?
4. **Defined names / named ranges** — surface them, or skip for v1?
5. **Hyperlinks on read** — v1 or a later minor?
