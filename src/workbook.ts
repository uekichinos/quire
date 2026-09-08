import { QuireError } from './errors'
import { parseRef, toRef } from './address'
import { dateToSerial } from './datetime'
import {
  appXml,
  contentTypesXml,
  coreXml,
  rootRelsXml,
  sharedStringsXml,
  workbookRelsXml,
  workbookXml,
  worksheetXml,
} from './serialize'
import { StylePool } from './style-pool'
import { isEmptyStyle, mergeStyle, type CellStyle } from './style'
import { numToXml } from './number'
import { escapeText } from './xml'
import { zipParts } from './zip'

/** A formula cell: an A1 formula without the leading `=`, plus an optional cached result. */
export interface FormulaValue {
  formula: string
  result?: string | number | boolean
}

/** Scalar values a cell can hold. */
export type CellScalar = string | number | boolean | Date | FormulaValue | null | undefined

/** A cell may be given as a bare value, or as `{ value, style }` for per-cell formatting. */
export type CellInput = CellScalar | { value: CellScalar; style?: CellStyle }

export interface AddRowOptions {
  /** Default style for every cell in the row (individual cell styles win). */
  style?: CellStyle
  /** Row height in points. */
  height?: number
}

export interface RowOptions {
  style?: CellStyle
  height?: number
}

export interface ColumnSpec {
  /** Width in Excel "character" units (roughly the count of `0` glyphs that fit). */
  width?: number
  hidden?: boolean
  /** Default style for the whole column (row and cell styles win over it). */
  style?: CellStyle
}

export interface FreezeOptions {
  /** Number of columns to freeze on the left. */
  xSplit?: number
  /** Number of rows to freeze on the top. */
  ySplit?: number
}

export interface WorksheetOptions {
  /** Column specs applied to columns 1..N in order. */
  columns?: ColumnSpec[]
  freeze?: FreezeOptions
  /** Range for the filter dropdowns, e.g. `'A1:E1'`. */
  autoFilter?: string
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const INVALID_NAME_CHARS = /[\\/?*[\]:]/
const MAX_COL = 16384
const RANGE_RE = /^([A-Z]+[1-9][0-9]*):([A-Z]+[1-9][0-9]*)$/

function validateSheetName(name: string, taken: readonly string[]): void {
  if (typeof name !== 'string' || name.length === 0 || name.length > 31) {
    throw new QuireError('worksheet name must be 1–31 characters')
  }
  if (INVALID_NAME_CHARS.test(name)) {
    throw new QuireError('worksheet name cannot contain \\ / ? * [ ] :')
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new QuireError('worksheet name cannot start or end with an apostrophe')
  }
  if (taken.some((t) => t.toLowerCase() === name.toLowerCase())) {
    throw new QuireError(`duplicate worksheet name "${name}"`)
  }
}

function parseRange(range: string): { top: number; left: number; bottom: number; right: number } {
  const m = RANGE_RE.exec(range)
  if (!m) throw new QuireError(`invalid range "${range}" (expected e.g. "A1:C3")`)
  const a = parseRef(m[1]!)
  const b = parseRef(m[2]!)
  if (a.row > b.row || a.col > b.col) {
    throw new QuireError(`range "${range}" must go from top-left to bottom-right`)
  }
  return { top: a.row, left: a.col, bottom: b.row, right: b.col }
}

function isFormula(v: unknown): v is FormulaValue {
  return typeof v === 'object' && v !== null && typeof (v as FormulaValue).formula === 'string'
}

function isStyledCell(v: unknown): v is { value: CellScalar; style?: CellStyle } {
  return (
    typeof v === 'object' && v !== null && !(v instanceof Date) && 'value' in v && !isFormula(v)
  )
}

/** Running, de-duplicated table of cell strings, shared across the workbook. */
class SharedStrings {
  readonly list: string[] = []
  total = 0
  private index = new Map<string, number>()

  intern(value: string): number {
    this.total += 1
    let i = this.index.get(value)
    if (i === undefined) {
      i = this.list.length
      this.list.push(value)
      this.index.set(value, i)
    }
    return i
  }
}

type StoredScalar = string | number | boolean | Date | FormulaValue

interface StoredCell {
  value: StoredScalar
  style?: CellStyle
}

export interface Worksheet {
  readonly name: string
  /** Append a row of values; the row index is assigned automatically. Chainable. */
  addRow(values: CellInput[], options?: AddRowOptions): this
  /** Set a single cell by A1 reference. Chainable. */
  setCell(ref: string, value: CellScalar, style?: CellStyle): this
  /** Style/height for an existing or future row (1-based). Chainable. */
  setRow(row: number, options: RowOptions): this
  /** Width / hidden / default style for a column (1-based). Chainable. */
  setColumn(index: number, spec: ColumnSpec): this
  /** Merge a cell range, e.g. `'A1:C1'`. Put the value in the top-left cell. Chainable. */
  merge(range: string): this
  /** Freeze rows and/or columns. Chainable. */
  freeze(options: FreezeOptions): this
  /** Add filter dropdowns over a header range, e.g. `'A1:E1'`. Chainable. */
  autoFilter(range: string): this
}

class QuireWorksheet implements Worksheet {
  readonly name: string
  private rows = new Map<number, Map<number, StoredCell>>()
  private rowMeta = new Map<number, RowOptions>()
  private columns = new Map<number, ColumnSpec>()
  private merges: string[] = []
  private frozen: FreezeOptions | null = null
  private filterRange: string | null = null
  private nextRow = 1
  private maxRow = 0
  private maxCol = 0

  // The workbook passes a fresh shared-strings table + style pool into
  // `serialize()` each time, so `xlsx()` is pure and repeatable.
  private sst!: SharedStrings
  private pool!: StylePool

  constructor(name: string, options: WorksheetOptions = {}) {
    this.name = name
    options.columns?.forEach((spec, i) => this.setColumn(i + 1, spec))
    if (options.freeze) this.freeze(options.freeze)
    if (options.autoFilter) this.autoFilter(options.autoFilter)
  }

  addRow(values: CellInput[], options: AddRowOptions = {}): this {
    if (!Array.isArray(values)) {
      throw new QuireError('addRow expects an array of values')
    }
    const r = this.nextRow++
    if (options.style || options.height != null) {
      this.setRow(r, { style: options.style, height: options.height })
    }
    values.forEach((input, i) => {
      if (isStyledCell(input)) this.put(r, i + 1, input.value, input.style)
      else this.put(r, i + 1, input as CellScalar)
    })
    return this
  }

  setCell(ref: string, value: CellScalar, style?: CellStyle): this {
    const { row, col } = parseRef(ref)
    this.put(row, col, value, style)
    this.nextRow = Math.max(this.nextRow, row + 1)
    return this
  }

  setRow(row: number, options: RowOptions): this {
    if (!Number.isInteger(row) || row < 1) {
      throw new QuireError('row index must be a positive integer')
    }
    if (options.height != null && !(options.height > 0)) {
      throw new QuireError('row height must be a positive number')
    }
    const meta: RowOptions = { ...this.rowMeta.get(row) }
    if (options.style && !isEmptyStyle(options.style)) meta.style = options.style
    if (options.height != null) meta.height = options.height
    this.rowMeta.set(row, meta)
    this.nextRow = Math.max(this.nextRow, row + 1)
    return this
  }

  setColumn(index: number, spec: ColumnSpec): this {
    if (!Number.isInteger(index) || index < 1 || index > MAX_COL) {
      throw new QuireError(`column index must be between 1 and ${MAX_COL}`)
    }
    if (spec.width != null && !(spec.width >= 0)) {
      throw new QuireError('column width must be a non-negative number')
    }
    this.columns.set(index, { ...this.columns.get(index), ...spec })
    return this
  }

  merge(range: string): this {
    const { top, left, bottom, right } = parseRange(range)
    if (top === bottom && left === right) {
      throw new QuireError(`cannot merge a single cell ("${range}")`)
    }
    const norm = `${toRef(top, left)}:${toRef(bottom, right)}`
    for (const existing of this.merges) {
      const e = parseRange(existing)
      const overlaps = left <= e.right && right >= e.left && top <= e.bottom && bottom >= e.top
      if (overlaps) {
        throw new QuireError(`merge "${norm}" overlaps existing merge "${existing}"`)
      }
    }
    this.merges.push(norm)
    this.maxRow = Math.max(this.maxRow, bottom)
    this.maxCol = Math.max(this.maxCol, right)
    return this
  }

  freeze(options: FreezeOptions): this {
    const x = options.xSplit ?? 0
    const y = options.ySplit ?? 0
    if (!Number.isInteger(x) || !Number.isInteger(y) || x < 0 || y < 0) {
      throw new QuireError('freeze splits must be non-negative integers')
    }
    this.frozen = x === 0 && y === 0 ? null : { xSplit: x, ySplit: y }
    return this
  }

  autoFilter(range: string): this {
    parseRange(range)
    this.filterRange = range
    return this
  }

  /** @internal — the header range this sheet filters, for the workbook defined name. */
  getFilterRange(): string | undefined {
    return this.filterRange ?? undefined
  }

  private put(row: number, col: number, value: CellScalar, style?: CellStyle): void {
    if (value === null || value === undefined) return
    const stored: StoredScalar = value
    let line = this.rows.get(row)
    if (!line) {
      line = new Map()
      this.rows.set(row, line)
    }
    line.set(col, { value: stored, style })
    this.maxRow = Math.max(this.maxRow, row)
    this.maxCol = Math.max(this.maxCol, col)
  }

  /** @internal */
  serialize(sst: SharedStrings, pool: StylePool): string {
    this.sst = sst
    this.pool = pool
    const rowNumbers = [...new Set([...this.rows.keys(), ...this.rowMeta.keys()])].sort(
      (a, b) => a - b,
    )

    const rows = rowNumbers
      .map((r) => {
        const line = this.rows.get(r)
        const meta = this.rowMeta.get(r)
        const rowStyle = meta?.style
        const cells = line
          ? [...line.keys()]
              .sort((a, b) => a - b)
              .map((c) => this.renderCell(toRef(r, c), line.get(c)!, c, rowStyle))
              .join('')
          : ''
        const attrs = [`r="${r}"`]
        if (meta?.height != null) attrs.push(`ht="${meta.height}"`, 'customHeight="1"')
        return `<row ${attrs.join(' ')}>${cells}</row>`
      })
      .join('')

    return worksheetXml({
      dimension: this.maxRow > 0 ? `A1:${toRef(this.maxRow, this.maxCol)}` : 'A1',
      sheetViews: this.frozen ? sheetViewsXml(this.frozen) : undefined,
      cols: this.colsXml(),
      rows,
      autoFilter: this.filterRange ? `<autoFilter ref="${this.filterRange}"/>` : undefined,
      mergeCells: this.merges.length
        ? `<mergeCells count="${this.merges.length}">` +
          this.merges.map((r) => `<mergeCell ref="${r}"/>`).join('') +
          `</mergeCells>`
        : undefined,
    })
  }

  private colsXml(): string | undefined {
    if (this.columns.size === 0) return undefined
    const cols = [...this.columns.keys()]
      .sort((a, b) => a - b)
      .map((idx) => {
        const spec = this.columns.get(idx)!
        const attrs = [`min="${idx}"`, `max="${idx}"`]
        if (spec.width != null) attrs.push(`width="${spec.width}"`, 'customWidth="1"')
        if (spec.hidden) attrs.push('hidden="1"')
        if (spec.style && !isEmptyStyle(spec.style)) {
          attrs.push(`style="${this.pool.intern(spec.style)}"`)
        }
        return `<col ${attrs.join(' ')}/>`
      })
      .join('')
    return `<cols>${cols}</cols>`
  }

  private renderCell(
    ref: string,
    cell: StoredCell,
    col: number,
    rowStyle: CellStyle | undefined,
  ): string {
    const colStyle = this.columns.get(col)?.style
    const style = mergeStyle(mergeStyle(colStyle, rowStyle), cell.style)
    const { value } = cell

    if (isFormula(value)) {
      const s = this.pool.intern(style)
      const sAttr = s ? ` s="${s}"` : ''
      const f = `<f>${escapeText(value.formula.replace(/^=/, ''))}</f>`
      const res = value.result
      if (res === undefined) return `<c r="${ref}"${sAttr}>${f}</c>`
      if (typeof res === 'string') {
        return `<c r="${ref}"${sAttr} t="str">${f}<v>${escapeText(res)}</v></c>`
      }
      if (typeof res === 'boolean') {
        return `<c r="${ref}"${sAttr} t="b">${f}<v>${res ? 1 : 0}</v></c>`
      }
      if (!Number.isFinite(res)) {
        throw new QuireError(`formula result for ${ref} is not a finite number`)
      }
      return `<c r="${ref}"${sAttr}>${f}<v>${numToXml(res)}</v></c>`
    }

    if (value instanceof Date) {
      const s = this.pool.intern(style, { isDate: true })
      return `<c r="${ref}" s="${s}"><v>${numToXml(dateToSerial(value))}</v></c>`
    }

    const s = this.pool.intern(style)
    const sAttr = s ? ` s="${s}"` : ''

    if (typeof value === 'number') {
      if (!Number.isFinite(value)) {
        throw new QuireError(`cell ${ref} is not a finite number`)
      }
      return `<c r="${ref}"${sAttr}><v>${numToXml(value)}</v></c>`
    }
    if (typeof value === 'boolean') {
      return `<c r="${ref}"${sAttr} t="b"><v>${value ? 1 : 0}</v></c>`
    }
    return `<c r="${ref}"${sAttr} t="s"><v>${this.sst.intern(value)}</v></c>`
  }
}

function sheetViewsXml(f: FreezeOptions): string {
  const x = f.xSplit ?? 0
  const y = f.ySplit ?? 0
  const topLeft = toRef(y + 1, x + 1)
  const activePane = x > 0 && y > 0 ? 'bottomRight' : x > 0 ? 'topRight' : 'bottomLeft'
  const paneAttrs = [
    x > 0 ? `xSplit="${x}"` : '',
    y > 0 ? `ySplit="${y}"` : '',
    `topLeftCell="${topLeft}"`,
    `activePane="${activePane}"`,
    'state="frozen"',
  ]
    .filter(Boolean)
    .join(' ')
  return (
    '<sheetViews><sheetView workbookViewId="0">' +
    `<pane ${paneAttrs}/>` +
    `<selection pane="${activePane}" activeCell="${topLeft}" sqref="${topLeft}"/>` +
    '</sheetView></sheetViews>'
  )
}

export interface Workbook {
  /** Add a worksheet. Names: 1–31 chars, unique (case-insensitive), no `\ / ? * [ ] :`. */
  addWorksheet(name: string, options?: WorksheetOptions): Worksheet
  /** Serialise to an in-memory `.xlsx` byte array. */
  xlsx(): Uint8Array
  /** Serialise to a `Blob` (browser convenience). */
  blob(): Blob
}

class QuireWorkbook implements Workbook {
  private sheets: QuireWorksheet[] = []

  addWorksheet(name: string, options: WorksheetOptions = {}): Worksheet {
    validateSheetName(
      name,
      this.sheets.map((s) => s.name),
    )
    const sheet = new QuireWorksheet(name, options)
    this.sheets.push(sheet)
    return sheet
  }

  xlsx(): Uint8Array {
    if (this.sheets.length === 0) {
      throw new QuireError('a workbook needs at least one worksheet')
    }

    // Fresh per call — `xlsx()` / `blob()` are pure and repeatable.
    const sst = new SharedStrings()
    const pool = new StylePool()

    // Serialise sheets first — this populates the shared-string table and style pool.
    const sheetXmls = this.sheets.map((s) => s.serialize(sst, pool))
    const hasStrings = sst.list.length > 0
    const created = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')

    const parts: Record<string, string> = {
      '[Content_Types].xml': contentTypesXml(this.sheets.length, hasStrings),
      '_rels/.rels': rootRelsXml(),
      'docProps/core.xml': coreXml(created),
      'docProps/app.xml': appXml(),
      'xl/workbook.xml': workbookXml(
        this.sheets.map((s) => s.name),
        this.sheets.map((s) => s.getFilterRange()),
      ),
      'xl/_rels/workbook.xml.rels': workbookRelsXml(this.sheets.length, hasStrings),
      'xl/styles.xml': pool.toXml(),
    }
    sheetXmls.forEach((xml, i) => {
      parts[`xl/worksheets/sheet${i + 1}.xml`] = xml
    })
    if (hasStrings) {
      parts['xl/sharedStrings.xml'] = sharedStringsXml(sst.list, sst.total)
    }

    return zipParts(parts)
  }

  blob(): Blob {
    return new Blob([this.xlsx() as BlobPart], { type: XLSX_MIME })
  }
}

/** Create a new, empty workbook. */
export function createWorkbook(): Workbook {
  return new QuireWorkbook()
}
