import { colLetter, parseRef } from './address'
import { serialToDate } from './datetime'
import { QuireError } from './errors'
import { builtinNumFmtCode, isDateNumFmt } from './numfmt'
import { parseStyleSheet, type ReadStyle, type StyleSheet } from './style-read'
import { DEFAULT_LIMITS, extractParts, XlsxReadError, type UnzipLimits } from './unzip'
import { parseXml } from './xml-read'

export type { ReadStyle, ReadFont, ReadBorderEdge } from './style-read'

export type ReadCellType =
  | 'string'
  | 'number'
  | 'boolean'
  | 'date'
  | 'formula'
  | 'error'
  | 'empty'

export interface ReadCell {
  /** A1 reference, e.g. `'B2'`. */
  ref: string
  row: number
  col: number
  type: ReadCellType
  /** The cell's value. For `formula` cells this is the cached result. */
  value: string | number | boolean | Date | null
  /** Present when the cell contains a formula (without the leading `=`). */
  formula?: string
  /** The cell's number-format code, when it has a non-`General` format. */
  numFmt?: string
  /** Resolved cell style — only when `readWorkbook(bytes, { styles: true })`. */
  style?: ReadStyle
}

/** Caps that guard against resource exhaustion when parsing untrusted files. */
export interface ReadLimits extends UnzipLimits {
  /** Max total number of cells across all loaded sheets. Default 5,000,000. */
  maxCells: number
  /** Max number of worksheets. Default 256. */
  maxSheets: number
}

export interface ReadOptions {
  /** Restrict to these sheets (by name or 0-based index). Default: all. */
  sheets?: (string | number)[]
  /** Convert numeric cells with a date number-format into `Date` objects. Default: `true`. */
  dates?: boolean
  /** Resolve per-cell styles (font / fill / border / alignment / numFmt). Default: `false`. */
  styles?: boolean
  limits?: Partial<ReadLimits>
}

const DEFAULT_READ_LIMITS: ReadLimits = {
  ...DEFAULT_LIMITS,
  maxCells: 5_000_000,
  maxSheets: 256,
}

export interface ValuesOptions {
  /** Trim each row to its own last populated cell instead of padding to `maxCol`. */
  ragged?: boolean
}

export interface ReadWorksheet {
  readonly name: string
  /** `{ rows, cols }` — from `<dimension>` when present, else the populated extent. */
  readonly dimension: { rows: number; cols: number }
  /** Merged ranges, e.g. `['A1:C1']`. */
  readonly merges: readonly string[]
  /** One cell by A1 reference, or `undefined` if empty/out of range. */
  cell(ref: string): ReadCell | undefined
  /** Iterate the populated rows; each is a sparse array indexed by `col - 1`. */
  rows(): IterableIterator<ReadCell[]>
  /** Rectangular `ReadCell[][]` (row 1..maxRow × col 1..maxCol), holes `undefined`. */
  toArray(): ReadCell[][]
  /** Row/column grid of values — `null` for empty cells. Rectangular unless `{ ragged: true }`. */
  values(options?: ValuesOptions): (string | number | boolean | Date | null)[][]
}

export interface ReadWorkbook {
  readonly sheetNames: string[]
  readonly sheets: ReadWorksheet[]
  readonly date1904: boolean
  sheet(nameOrIndex: string | number): ReadWorksheet | undefined
}

/* -------------------------------------------------------------------------- */
/*  Part parsers                                                               */
/* -------------------------------------------------------------------------- */

interface SheetRef {
  name: string
  rid: string
}

function parseWorkbook(xml: string): { sheets: SheetRef[]; date1904: boolean } {
  const sheets: SheetRef[] = []
  let date1904 = false
  parseXml(xml, {
    onOpen(name, attrs) {
      if (name === 'sheet' || name.endsWith(':sheet')) {
        const rid = attrs['r:id'] ?? attrs['relationships:id'] ?? attrs.id ?? ''
        sheets.push({ name: attrs.name ?? '', rid })
      } else if (name === 'workbookPr' || name.endsWith(':workbookPr')) {
        const v = attrs.date1904
        date1904 = v === '1' || v === 'true'
      }
    },
  })
  return { sheets, date1904 }
}

function parseWorkbookRels(xml: string): Map<string, string> {
  const rels = new Map<string, string>()
  parseXml(xml, {
    onOpen(name, attrs) {
      if (name === 'Relationship' && attrs.Id && attrs.Target) {
        let target = attrs.Target.replace(/\\/g, '/')
        if (target.includes('..')) return // path traversal — ignore
        target = target.replace(/^\/?xl\//, '').replace(/^\//, '')
        rels.set(attrs.Id, `xl/${target}`.toLowerCase())
      }
    },
  })
  return rels
}

function parseSharedStrings(xml: string): string[] {
  const list: string[] = []
  let current: string[] | null = null
  let inT = 0
  parseXml(xml, {
    onOpen(name) {
      if (name === 'si') current = []
      else if (name === 't') inT++
    },
    onText(text) {
      if (inT > 0 && current) current.push(text)
    },
    onClose(name) {
      if (name === 't') inT = Math.max(0, inT - 1)
      else if (name === 'si') {
        list.push((current ?? []).join(''))
        current = null
      }
    },
  })
  return list
}

interface RawSheet {
  cells: Map<number, Map<number, ReadCell>>
  maxRow: number
  maxCol: number
  dimRef?: string
  merges: string[]
}

interface SheetParseCtx {
  sst: string[]
  styles: StyleSheet | null
  date1904: boolean
  dates: boolean
  withStyles: boolean
  maxCells: number
  /** Shared running cell count across all sheets. */
  counter: { n: number }
}

function numFmtIdOf(styleIdx: number, styles: StyleSheet | null): number | undefined {
  if (!styles || styleIdx < 0) return undefined
  return styles.xfNumFmtId[styleIdx]
}

function numFmtCodeOf(styleIdx: number, styles: StyleSheet | null): string | undefined {
  const id = numFmtIdOf(styleIdx, styles)
  if (id === undefined || id === 0) return undefined
  return styles!.customCode.get(id) ?? builtinNumFmtCode(id)
}

function isDateStyle(styleIdx: number, styles: StyleSheet | null): boolean {
  const id = numFmtIdOf(styleIdx, styles)
  if (id === undefined) return false
  return isDateNumFmt(id, styles!.customCode.get(id))
}

function parseSheet(xml: string, ctx: SheetParseCtx): RawSheet {
  const { sst, styles, date1904, dates, withStyles, counter, maxCells } = ctx
  const cells = new Map<number, Map<number, ReadCell>>()
  const merges: string[] = []
  let maxRow = 0
  let maxCol = 0
  let dimRef: string | undefined

  let curRow = 0
  let lastRow = 0
  let curCol = 0
  let cRef = ''
  let cType = ''
  let cStyle = -1
  let hasFormula = false
  let vBuf = ''
  let fBuf = ''
  let isBuf = ''
  let inV = false
  let inF = false
  let inIs = false
  let inIsT = 0

  const materialise = (): void => {
    if (cRef) {
      const p = parseRef(cRef)
      curRow = p.row
      curCol = p.col
    } else {
      curCol += 1
    }
    const ref = cRef || `${colLetter(curCol)}${curRow}`

    let type: ReadCellType
    let value: string | number | boolean | Date | null

    if (hasFormula) {
      type = 'formula'
      if (cType === 'str') value = vBuf
      else if (cType === 'b') value = vBuf === '1'
      else if (cType === 'e') value = vBuf
      else value = vBuf === '' ? null : Number(vBuf)
    } else if (cType === 's') {
      const idx = Number(vBuf)
      value = Number.isInteger(idx) && idx >= 0 && idx < sst.length ? sst[idx]! : ''
      type = 'string'
    } else if (cType === 'inlineStr') {
      value = isBuf
      type = 'string'
    } else if (cType === 'str') {
      value = vBuf
      type = 'string'
    } else if (cType === 'b') {
      value = vBuf === '1'
      type = 'boolean'
    } else if (cType === 'e') {
      value = vBuf
      type = 'error'
    } else if (vBuf === '') {
      value = null
      type = 'empty'
    } else {
      const num = Number(vBuf)
      if (dates && cStyle >= 0 && Number.isFinite(num) && isDateStyle(cStyle, styles)) {
        value = serialToDate(num, date1904)
        type = 'date'
      } else {
        value = num
        type = 'number'
      }
    }

    const cell: ReadCell = { ref, row: curRow, col: curCol, type, value }
    if (hasFormula && fBuf) cell.formula = fBuf.replace(/^=/, '')
    const numFmt = numFmtCodeOf(cStyle, styles)
    if (numFmt) cell.numFmt = numFmt
    if (withStyles && styles && cStyle >= 0) {
      const st = styles.xfStyle[cStyle]
      if (st && Object.keys(st).length) cell.style = st
    }

    if (++counter.n > maxCells) {
      throw new QuireError(`workbook has more than ${maxCells} cells (the read limit)`)
    }

    let line = cells.get(curRow)
    if (!line) {
      line = new Map()
      cells.set(curRow, line)
    }
    line.set(curCol, cell)
    maxRow = Math.max(maxRow, curRow)
    maxCol = Math.max(maxCol, curCol)
  }

  parseXml(xml, {
    onOpen(name, attrs) {
      switch (name) {
        case 'dimension':
          dimRef = attrs.ref
          break
        case 'mergeCell':
          if (attrs.ref) merges.push(attrs.ref)
          break
        case 'row':
          curRow = attrs.r ? Number(attrs.r) : lastRow + 1
          curCol = 0
          break
        case 'c': {
          cRef = attrs.r ?? ''
          cType = attrs.t ?? ''
          const s = Number(attrs.s)
          cStyle = Number.isInteger(s) && s >= 0 ? s : -1
          hasFormula = false
          vBuf = ''
          fBuf = ''
          isBuf = ''
          break
        }
        case 'v':
          inV = true
          break
        case 'f':
          inF = true
          hasFormula = true
          break
        case 'is':
          inIs = true
          break
        case 't':
          if (inIs) inIsT++
          break
      }
    },
    onText(text) {
      if (inV) vBuf += text
      else if (inF) fBuf += text
      else if (inIs && inIsT > 0) isBuf += text
    },
    onClose(name) {
      switch (name) {
        case 'v':
          inV = false
          break
        case 'f':
          inF = false
          break
        case 't':
          if (inIs) inIsT = Math.max(0, inIsT - 1)
          break
        case 'is':
          inIs = false
          break
        case 'c':
          materialise()
          break
        case 'row':
          lastRow = curRow
          break
      }
    },
  })

  return { cells, maxRow, maxCol, dimRef, merges }
}

/* -------------------------------------------------------------------------- */
/*  Model                                                                      */
/* -------------------------------------------------------------------------- */

class Worksheet implements ReadWorksheet {
  readonly dimension: { rows: number; cols: number }
  readonly merges: readonly string[]

  constructor(
    readonly name: string,
    private readonly cells: Map<number, Map<number, ReadCell>>,
    private readonly maxRow: number,
    private readonly maxCol: number,
    merges: string[],
    dimRef?: string,
  ) {
    this.merges = merges
    if (dimRef && dimRef.includes(':')) {
      const [, br] = dimRef.split(':')
      const p = parseRef(br!)
      this.dimension = { rows: p.row, cols: p.col }
    } else {
      this.dimension = { rows: maxRow, cols: maxCol }
    }
  }

  cell(ref: string): ReadCell | undefined {
    const { row, col } = parseRef(ref)
    return this.cells.get(row)?.get(col)
  }

  *rows(): IterableIterator<ReadCell[]> {
    const nums = [...this.cells.keys()].sort((a, b) => a - b)
    for (const r of nums) {
      const line = this.cells.get(r)!
      const width = Math.max(...line.keys())
      const arr: ReadCell[] = new Array(width)
      for (const [c, cell] of line) arr[c - 1] = cell
      yield arr
    }
  }

  toArray(): ReadCell[][] {
    const out: ReadCell[][] = []
    for (let r = 1; r <= this.maxRow; r++) {
      const line = this.cells.get(r)
      const arr: ReadCell[] = new Array(this.maxCol)
      if (line) for (const [c, cell] of line) arr[c - 1] = cell
      out.push(arr)
    }
    return out
  }

  values(options: ValuesOptions = {}): (string | number | boolean | Date | null)[][] {
    return this.toArray().map((row) => {
      const width = options.ragged ? lastIndex(row) + 1 : this.maxCol
      const out: (string | number | boolean | Date | null)[] = new Array(width).fill(null)
      for (let c = 0; c < width; c++) if (row[c]) out[c] = row[c]!.value
      return out
    })
  }
}

function lastIndex(row: readonly unknown[]): number {
  for (let i = row.length - 1; i >= 0; i--) if (row[i] !== undefined) return i
  return -1
}

/* -------------------------------------------------------------------------- */
/*  Entry point                                                                */
/* -------------------------------------------------------------------------- */

interface Prepared {
  sheetRefs: SheetRef[]
  date1904: boolean
  /** Build one worksheet, or `null` if the `sheets` filter excludes it. */
  buildSheet(sr: SheetRef, index: number): Worksheet | null
}

function prepare(input: Uint8Array | ArrayBuffer, options: ReadOptions): Prepared {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  const limits: ReadLimits = { ...DEFAULT_READ_LIMITS, ...options.limits }
  const parts = extractParts(bytes, limits)

  const { sheets: sheetRefs, date1904 } = parseWorkbook(parts.get('xl/workbook.xml')!)
  if (sheetRefs.length > limits.maxSheets) {
    throw new XlsxReadError(`workbook has ${sheetRefs.length} sheets (limit ${limits.maxSheets})`)
  }

  const rels = parts.has('xl/_rels/workbook.xml.rels')
    ? parseWorkbookRels(parts.get('xl/_rels/workbook.xml.rels')!)
    : new Map<string, string>()
  const sst = parts.has('xl/sharedstrings.xml')
    ? parseSharedStrings(parts.get('xl/sharedstrings.xml')!)
    : []

  const withStyles = options.styles === true
  let themeXml: string | undefined
  if (withStyles) {
    for (const [name, xml] of parts) {
      if (name.startsWith('xl/theme/')) {
        themeXml = xml
        break
      }
    }
  }
  const styles = parts.has('xl/styles.xml')
    ? parseStyleSheet(parts.get('xl/styles.xml')!, { withStyles, themeXml })
    : null

  const ctx: SheetParseCtx = {
    sst,
    styles,
    date1904,
    dates: options.dates !== false,
    withStyles,
    maxCells: limits.maxCells,
    counter: { n: 0 },
  }

  const wanted = options.sheets
  const shouldLoad = (name: string, index: number): boolean =>
    !wanted || wanted.some((w) => (typeof w === 'number' ? w === index : w === name))

  return {
    sheetRefs,
    date1904,
    buildSheet(sr, index) {
      if (!shouldLoad(sr.name, index)) return null
      const target = rels.get(sr.rid) ?? `xl/worksheets/sheet${index + 1}.xml`
      const xml = parts.get(target.toLowerCase())
      if (!xml) {
        throw new XlsxReadError(`worksheet part "${target}" for sheet "${sr.name}" is missing`)
      }
      const raw = parseSheet(xml, ctx)
      return new Worksheet(sr.name, raw.cells, raw.maxRow, raw.maxCol, raw.merges, raw.dimRef)
    },
  }
}

function assemble(built: Worksheet[], date1904: boolean): ReadWorkbook {
  return {
    sheetNames: built.map((s) => s.name),
    sheets: built,
    date1904,
    sheet(nameOrIndex) {
      return typeof nameOrIndex === 'number'
        ? built[nameOrIndex]
        : built.find((s) => s.name === nameOrIndex)
    },
  }
}

/**
 * Parses an `.xlsx` byte array into a read model.
 *
 * @example
 * const wb = readWorkbook(bytes)
 * for (const row of wb.sheet('Sales')!.rows()) {
 *   console.log(row.map((c) => c?.value))
 * }
 */
export function readWorkbook(
  input: Uint8Array | ArrayBuffer,
  options: ReadOptions = {},
): ReadWorkbook {
  const { sheetRefs, date1904, buildSheet } = prepare(input, options)
  const built: Worksheet[] = []
  sheetRefs.forEach((sr, i) => {
    const w = buildSheet(sr, i)
    if (w) built.push(w)
  })
  return assemble(built, date1904)
}

/**
 * Like {@link readWorkbook} but yields to the event loop between worksheets, so
 * a large multi-sheet import doesn't monopolise the thread in one tick. A single
 * huge sheet still parses in one synchronous step — run those in a worker.
 */
export async function readWorkbookAsync(
  input: Uint8Array | ArrayBuffer,
  options: ReadOptions = {},
): Promise<ReadWorkbook> {
  const { sheetRefs, date1904, buildSheet } = prepare(input, options)
  const built: Worksheet[] = []
  for (let i = 0; i < sheetRefs.length; i++) {
    const w = buildSheet(sheetRefs[i]!, i)
    if (w) built.push(w)
    await Promise.resolve()
  }
  return assemble(built, date1904)
}
