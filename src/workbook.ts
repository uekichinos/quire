import { QuireError } from './errors'
import { parseRef, toRef } from './address'
import { dateToSerial } from './datetime'
import {
  appXml,
  commentsXml,
  contentTypesXml,
  coreXml,
  type DefinedNameEntry,
  REL_COMMENTS,
  REL_HYPERLINK,
  REL_VML,
  rootRelsXml,
  sharedStringsXml,
  vmlDrawingXml,
  workbookRelsXml,
  workbookXml,
  worksheetRelsXml,
  worksheetXml,
  type WorksheetRelEntry,
} from './serialize'
import { StylePool } from './style-pool'
import { isEmptyStyle, mergeStyle, toArgb, type CellStyle, type FontStyle } from './style'
import { numToXml } from './number'
import { attr, escapeText } from './xml'
import { zipParts } from './zip'

/** A formula cell: an A1 formula without the leading `=`, plus an optional cached result. */
export interface FormulaValue {
  formula: string
  result?: string | number | boolean
}

/**
 * A hyperlink cell: `hyperlink` is the target URI (`http(s)://`, `mailto:`, …);
 * `text` is what's displayed (defaults to `hyperlink` itself); `tooltip` is the
 * hover text Excel shows.
 */
export interface HyperlinkValue {
  hyperlink: string
  text?: string
  tooltip?: string
}

/** One formatted span of a rich-text cell. `font` overrides only what it sets. */
export interface RichTextRun {
  text: string
  font?: Pick<FontStyle, 'name' | 'size' | 'bold' | 'italic' | 'underline' | 'color'>
}

/** A cell value made of multiple differently-formatted runs, e.g. a bold word mid-sentence. */
export type RichText = RichTextRun[]

/** Scalar values a cell can hold. */
export type CellScalar =
  | string
  | number
  | boolean
  | Date
  | FormulaValue
  | HyperlinkValue
  | RichText
  | null
  | undefined

/** A cell may be given as a bare value, or as `{ value, style }` for per-cell formatting. */
export type CellInput = CellScalar | { value: CellScalar; style?: CellStyle }

export interface AddRowOptions {
  /** Default style for every cell in the row (individual cell styles win). */
  style?: CellStyle
  /** Row height in points. */
  height?: number
  hidden?: boolean
  /** Outline (grouping) level, `0`–`7`. */
  outlineLevel?: number
}

export interface RowOptions {
  style?: CellStyle
  height?: number
  hidden?: boolean
  /** Outline (grouping) level, `0`–`7`. */
  outlineLevel?: number
}

export interface ColumnSpec {
  /** Width in Excel "character" units (roughly the count of `0` glyphs that fit). */
  width?: number
  hidden?: boolean
  /** Default style for the whole column (row and cell styles win over it). */
  style?: CellStyle
  /** Outline (grouping) level, `0`–`7`. */
  outlineLevel?: number
}

/** Which side the outline's summary row/column sits on. Both default to `true` (Excel's own default). */
export interface OutlineOptions {
  /** Summary row is below its detail rows, not above. */
  summaryBelow?: boolean
  /** Summary column is to the right of its detail columns, not the left. */
  summaryRight?: boolean
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
  /** Direction of row/column outline (grouping) summaries. */
  outline?: OutlineOptions
}

export interface DefineNameOptions {
  /** Limits visibility to this sheet instead of the whole workbook. Must already exist. */
  scope?: string
}

/**
 * Structural sheet protection — what's allowed once the sheet is protected. Cell-level
 * `locked`/`hidden` (via `CellStyle.protection`) decide which cells that applies to; every
 * cell is implicitly locked until styled otherwise. No password support: Excel's legacy
 * password hash needs to be bit-exact and can't be verified against a real Excel instance
 * in this environment, so protection here is structural-only.
 */
export interface SheetProtectionOptions {
  /** Selecting locked cells. Default `true` (Excel's own default). */
  allowSelectLockedCells?: boolean
  /** Selecting unlocked cells. Default `true`. */
  allowSelectUnlockedCells?: boolean
  /** Formatting cells. Default `false`. */
  allowFormatCells?: boolean
  /** Formatting columns (width, hide/unhide). Default `false`. */
  allowFormatColumns?: boolean
  /** Formatting rows (height, hide/unhide). Default `false`. */
  allowFormatRows?: boolean
  /** Inserting columns. Default `false`. */
  allowInsertColumns?: boolean
  /** Inserting rows. Default `false`. */
  allowInsertRows?: boolean
  /** Deleting columns. Default `false`. */
  allowDeleteColumns?: boolean
  /** Deleting rows. Default `false`. */
  allowDeleteRows?: boolean
  /** Sorting. Default `false`. */
  allowSort?: boolean
  /** Using the autofilter dropdowns. Default `false`. */
  allowAutoFilter?: boolean
}

/**
 * Structural workbook protection — locks the sheet list and/or the workbook window. No
 * password support (see {@link SheetProtectionOptions}).
 */
export interface WorkbookProtectionOptions {
  /** Locks the sheet structure (add/remove/rename/reorder/hide/unhide sheets). Default `true` — Excel's own default once workbook protection is turned on. */
  lockStructure?: boolean
  /** Locks the workbook window's size/position. Default `false`. */
  lockWindows?: boolean
}

export interface PageSetupOptions {
  orientation?: 'portrait' | 'landscape'
  /** Excel paper-size code, e.g. `1` = Letter, `9` = A4. */
  paperSize?: number
  /** Fit the printout to this many pages wide/tall (`0` = don't constrain that dimension). Wins over `scale`. */
  fitToWidth?: number
  fitToHeight?: number
  /** Zoom percentage. Ignored when `fitToWidth`/`fitToHeight` are set. */
  scale?: number
  /** Inches. Unset fields fall back to Excel's own defaults (0.7/0.7/0.75/0.75/0.3/0.3). */
  margins?: { left?: number; right?: number; top?: number; bottom?: number; header?: number; footer?: number }
  /** Repeats this range as the sheet's print area (a workbook-level `_xlnm.Print_Area` defined name). */
  printArea?: string
}

/** Comparison used by the numeric/date/text-length validation rules. `between`/`notBetween` take a 2-tuple `value`. */
export type ValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual'

/** A data-validation rule for `setDataValidation`. */
export type DataValidationRule =
  | {
      type: 'list'
      /** Fixed choices (`['Open', 'Done']`) — none may contain a comma — or a range reference (`'Lookup!A1:A5'`). */
      list: string[] | string
    }
  | { type: 'whole' | 'decimal' | 'textLength'; operator: ValidationOperator; value: number | [number, number] }
  | { type: 'date'; operator: ValidationOperator; value: Date | [Date, Date] }
  | {
      type: 'time'
      operator: ValidationOperator
      /** `'HH:MM'` or `'HH:MM:SS'`, or a pair for `between`/`notBetween`. */
      value: string | [string, string]
    }
  | {
      type: 'custom'
      /** An Excel formula (without `=`) that must evaluate truthy for the cell to be valid. */
      formula: string
    }

export interface DataValidationOptions {
  /** Allow the cell to be left empty. Default `true`. */
  allowBlank?: boolean
  promptTitle?: string
  promptMessage?: string
  errorTitle?: string
  errorMessage?: string
}

export interface CommentOptions {
  author?: string
}

/** A conditional-formatting rule for `addConditionalFormat`. */
export type ConditionalFormatRule =
  | {
      type: 'cellIs'
      operator: ValidationOperator
      /** One value, or two for `between`/`notBetween`. */
      formula: (string | number) | [string | number, string | number]
      /** Formatting applied when the rule matches. */
      style: { font?: Pick<FontStyle, 'bold' | 'italic' | 'color'>; fill?: string }
    }
  | {
      type: 'colorScale'
      /** 2 (min→max) or 3 (min→mid→max) stop colours; thresholds are the range's own min/mid/max. */
      colors: [string, string] | [string, string, string]
    }
  | {
      type: 'dataBar'
      /** Bar colour; length scales between the range's own min and max. */
      color: string
    }
  | {
      type: 'iconSet'
      /** Standard Excel icon-set name; thresholds are evenly-spaced percentiles. */
      iconSet: IconSetName
    }
  | {
      type: 'top10'
      /** How many (or, with `percent`, what percentage of) cells to highlight. */
      rank: number
      /** Interpret `rank` as a percentage instead of a count. Default `false`. */
      percent?: boolean
      /** Highlight the bottom instead of the top. Default `false`. */
      bottom?: boolean
      /** Formatting applied to matching cells. */
      style: { font?: Pick<FontStyle, 'bold' | 'italic' | 'color'>; fill?: string }
    }

/** Standard Excel icon-set names (`ST_IconSetType`). */
export type IconSetName =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const INVALID_NAME_CHARS = /[\\/?*[\]:]/
export const MAX_COL = 16384
const RANGE_RE = /^([A-Z]+[1-9][0-9]*):([A-Z]+[1-9][0-9]*)$/
// Excel's practical hyperlink-target limit (2019+; older versions cap at 255).
export const MAX_HYPERLINK_LEN = 2079
// Excel identifier rules: starts with a letter/underscore/backslash, then word chars or dots.
const DEFINED_NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.]*$/
// Rejects names that read as a plain cell reference (Excel disallows these too).
const CELL_LIKE_NAME_RE = /^[A-Za-z]{1,3}[0-9]{1,7}$/

export function validateSheetName(name: string, taken: readonly string[]): void {
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

export function parseRange(range: string): { top: number; left: number; bottom: number; right: number } {
  const m = RANGE_RE.exec(range)
  if (!m) throw new QuireError(`invalid range "${range}" (expected e.g. "A1:C3")`)
  const a = parseRef(m[1]!)
  const b = parseRef(m[2]!)
  if (a.row > b.row || a.col > b.col) {
    throw new QuireError(`range "${range}" must go from top-left to bottom-right`)
  }
  return { top: a.row, left: a.col, bottom: b.row, right: b.col }
}

export function validateDefinedName(name: string, taken: ReadonlySet<string>): void {
  if (typeof name !== 'string' || name.length === 0 || name.length > 255) {
    throw new QuireError('defined name must be 1–255 characters')
  }
  if (!DEFINED_NAME_RE.test(name)) {
    throw new QuireError(
      `defined name "${name}" must start with a letter, "_" or "\\", and contain only letters, digits, "_" and "."`,
    )
  }
  if (CELL_LIKE_NAME_RE.test(name)) {
    throw new QuireError(`defined name "${name}" looks like a cell reference, which Excel disallows`)
  }
  if (name.toLowerCase().startsWith('_xlnm.')) {
    throw new QuireError(`defined name "${name}" uses the "_xlnm." prefix, reserved for Excel`)
  }
  if (taken.has(name.toLowerCase())) {
    throw new QuireError(`duplicate defined name "${name}"`)
  }
}

/** Accepts a single cell ("A1") or a range ("A1:C3"); throws on anything else. */
export function validateNamedRangeTarget(range: string): void {
  if (range.includes(':')) parseRange(range)
  else parseRef(range)
}

export function validateConditionalFormatRule(range: string, rule: ConditionalFormatRule): void {
  if (rule.type === 'colorScale' && rule.colors.length !== 2 && rule.colors.length !== 3) {
    throw new QuireError(`colorScale for "${range}" needs 2 or 3 colours`)
  }
  if (rule.type === 'cellIs') {
    const isBetween = rule.operator === 'between' || rule.operator === 'notBetween'
    const isPair = Array.isArray(rule.formula)
    if (isBetween && !isPair) {
      throw new QuireError(`conditional format "${rule.operator}" for "${range}" needs a [min, max] formula`)
    }
    if (!isBetween && isPair) {
      throw new QuireError(`conditional format "${rule.operator}" for "${range}" takes a single formula, not a pair`)
    }
  }
  if (rule.type === 'top10' && (!Number.isInteger(rule.rank) || rule.rank < 1 || (rule.percent && rule.rank > 100))) {
    throw new QuireError(`top10 rank for "${range}" must be a positive integer (and ≤100 when percent is set)`)
  }
}

export function validateOutlineLevel(level: number | undefined): void {
  if (level != null && (!Number.isInteger(level) || level < 0 || level > 7)) {
    throw new QuireError('outlineLevel must be an integer between 0 and 7')
  }
}

export function validatePageSetup(options: PageSetupOptions): void {
  if (options.printArea) validateNamedRangeTarget(options.printArea)
  if (options.fitToWidth != null && (!Number.isInteger(options.fitToWidth) || options.fitToWidth < 0)) {
    throw new QuireError('fitToWidth must be a non-negative integer')
  }
  if (options.fitToHeight != null && (!Number.isInteger(options.fitToHeight) || options.fitToHeight < 0)) {
    throw new QuireError('fitToHeight must be a non-negative integer')
  }
  if (options.scale != null && !(options.scale > 0)) {
    throw new QuireError('scale must be a positive number')
  }
  if (options.paperSize != null && (!Number.isInteger(options.paperSize) || options.paperSize < 1)) {
    throw new QuireError('paperSize must be a positive integer')
  }
}

/** `formula1`/`formula2` text for a validation rule — dates become their Excel serial. */
/** `'HH:MM'` / `'HH:MM:SS'` → fraction of a day (`13:30` → `0.5625`), Excel's time-serial form. */
export function timeToFraction(time: string): number {
  const m = /^([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}(?:\.[0-9]+)?))?$/.exec(time.trim())
  if (!m) throw new QuireError(`invalid time "${time}" (expected "HH:MM" or "HH:MM:SS")`)
  const [, h, mins, secs] = m
  const totalSeconds = Number(h) * 3600 + Number(mins) * 60 + Number(secs ?? 0)
  return totalSeconds / 86400
}

export function renderValidationFormulas(rule: DataValidationRule): string[] {
  if (rule.type === 'list') {
    return [Array.isArray(rule.list) ? `"${rule.list.join(',')}"` : rule.list]
  }
  if (rule.type === 'custom') return [rule.formula]
  if (rule.type === 'time') {
    const toFormula = (v: string): string => numToXml(timeToFraction(v))
    return Array.isArray(rule.value) ? rule.value.map(toFormula) : [toFormula(rule.value)]
  }
  const toFormula = (v: number | Date): string =>
    numToXml(v instanceof Date ? dateToSerial(v) : v)
  return Array.isArray(rule.value) ? rule.value.map(toFormula) : [toFormula(rule.value)]
}

export function isFormula(v: unknown): v is FormulaValue {
  return typeof v === 'object' && v !== null && typeof (v as FormulaValue).formula === 'string'
}

export function isStyledCell(v: unknown): v is { value: CellScalar; style?: CellStyle } {
  return (
    typeof v === 'object' && v !== null && !(v instanceof Date) && 'value' in v && !isFormula(v)
  )
}

export function isHyperlinkValue(v: unknown): v is HyperlinkValue {
  return (
    typeof v === 'object' && v !== null && typeof (v as HyperlinkValue).hyperlink === 'string'
  )
}

export function isRichText(v: unknown): v is RichText {
  return (
    Array.isArray(v) &&
    v.every((r) => typeof r === 'object' && r !== null && typeof (r as RichTextRun).text === 'string')
  )
}

/** Running, de-duplicated table of cell strings (plain or rich text), shared across the workbook. */
export class SharedStrings {
  readonly list: (string | RichText)[] = []
  total = 0
  private index = new Map<string, number>()

  intern(value: string | RichText): number {
    this.total += 1
    const key = typeof value === 'string' ? `s:${value}` : `r:${JSON.stringify(value)}`
    let i = this.index.get(key)
    if (i === undefined) {
      i = this.list.length
      this.list.push(value)
      this.index.set(key, i)
    }
    return i
  }
}

export type StoredScalar = string | number | boolean | Date | FormulaValue | RichText

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
  /** Restrict a range to a dropdown, or a numeric/date/text-length rule. Chainable. */
  setDataValidation(range: string, rule: DataValidationRule, options?: DataValidationOptions): this
  /** Attach a comment/note to a cell, independent of its value. Chainable. */
  setComment(ref: string, text: string, options?: CommentOptions): this
  /** Highlight a range that matches a condition, or apply a colour scale. Chainable. */
  addConditionalFormat(range: string, rule: ConditionalFormatRule): this
  /** Page orientation, paper size, fit-to-page/scale, margins, and print area. Chainable. */
  setPageSetup(options: PageSetupOptions): this
  /**
   * Protect the sheet structurally (no password — see {@link SheetProtectionOptions}). Cells
   * are locked by default; use `CellStyle.protection` to unlock specific ones. Chainable.
   */
  protect(options?: SheetProtectionOptions): this
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
  private hyperlinks = new Map<string, { target: string; tooltip?: string }>()
  private validations: { range: string; rule: DataValidationRule; options: DataValidationOptions }[] = []
  private conditionalFormats: { range: string; rule: ConditionalFormatRule }[] = []
  private pageSetup: PageSetupOptions | null = null
  private outline: OutlineOptions | undefined
  private protection: SheetProtectionOptions | null = null
  private comments = new Map<string, { text: string; author: string }>()
  // Rebuilt fresh in serialize(): the sheet's single _rels/sheetN.xml.rels file
  // (hyperlinks + comments/vml, if any) and the legacyDrawing reference to it.
  private rels: { id: string; type: string; target: string; external?: boolean }[] = []

  // The workbook passes a fresh shared-strings table + style pool into
  // `serialize()` each time, so `xlsx()` is pure and repeatable.
  private sst!: SharedStrings
  private pool!: StylePool

  constructor(name: string, options: WorksheetOptions = {}) {
    this.name = name
    options.columns?.forEach((spec, i) => this.setColumn(i + 1, spec))
    if (options.freeze) this.freeze(options.freeze)
    if (options.autoFilter) this.autoFilter(options.autoFilter)
    this.outline = options.outline
  }

  addRow(values: CellInput[], options: AddRowOptions = {}): this {
    if (!Array.isArray(values)) {
      throw new QuireError('addRow expects an array of values')
    }
    const r = this.nextRow++
    if (options.style || options.height != null || options.hidden != null || options.outlineLevel != null) {
      this.setRow(r, {
        style: options.style,
        height: options.height,
        hidden: options.hidden,
        outlineLevel: options.outlineLevel,
      })
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
    validateOutlineLevel(options.outlineLevel)
    const meta: RowOptions = { ...this.rowMeta.get(row) }
    if (options.style && !isEmptyStyle(options.style)) meta.style = options.style
    if (options.height != null) meta.height = options.height
    if (options.hidden != null) meta.hidden = options.hidden
    if (options.outlineLevel != null) meta.outlineLevel = options.outlineLevel
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
    validateOutlineLevel(spec.outlineLevel)
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

  setPageSetup(options: PageSetupOptions): this {
    validatePageSetup(options)
    this.pageSetup = options
    return this
  }

  /** @internal — this sheet's print area, for the workbook defined name. */
  getPrintArea(): string | undefined {
    return this.pageSetup?.printArea
  }

  protect(options: SheetProtectionOptions = {}): this {
    this.protection = options
    return this
  }

  setDataValidation(
    range: string,
    rule: DataValidationRule,
    options: DataValidationOptions = {},
  ): this {
    validateNamedRangeTarget(range)
    if (rule.type === 'list') {
      if (Array.isArray(rule.list)) {
        if (rule.list.length === 0) {
          throw new QuireError(`data validation for "${range}" needs at least one choice`)
        }
        for (const item of rule.list) {
          if (item.includes(',')) {
            throw new QuireError(
              `data validation choice "${item}" for "${range}" cannot contain a comma (Excel's inline-list separator)`,
            )
          }
        }
        if (rule.list.join(',').length > 255) {
          throw new QuireError(`data validation inline list for "${range}" exceeds Excel's 255-character limit`)
        }
      } else if (!rule.list) {
        throw new QuireError(`data validation for "${range}" needs a list or a range reference`)
      }
    } else if (rule.type === 'custom') {
      if (!rule.formula) {
        throw new QuireError(`data validation for "${range}" needs a non-empty formula`)
      }
    } else {
      const isBetween = rule.operator === 'between' || rule.operator === 'notBetween'
      const isPair = Array.isArray(rule.value)
      if (isBetween && !isPair) {
        throw new QuireError(`data validation "${rule.operator}" for "${range}" needs a [min, max] value`)
      }
      if (!isBetween && isPair) {
        throw new QuireError(`data validation "${rule.operator}" for "${range}" takes a single value, not a pair`)
      }
    }
    this.validations.push({ range, rule, options })
    return this
  }

  addConditionalFormat(range: string, rule: ConditionalFormatRule): this {
    validateNamedRangeTarget(range)
    validateConditionalFormatRule(range, rule)
    this.conditionalFormats.push({ range, rule })
    return this
  }

  setComment(ref: string, text: string, options: CommentOptions = {}): this {
    parseRef(ref)
    if (typeof text !== 'string' || text.length === 0) {
      throw new QuireError(`comment text for ${ref} must be a non-empty string`)
    }
    this.comments.set(ref, { text, author: options.author ?? '' })
    return this
  }

  private put(row: number, col: number, value: CellScalar, style?: CellStyle): void {
    if (value === null || value === undefined) return
    const ref = toRef(row, col)
    let stored: StoredScalar
    if (isHyperlinkValue(value)) {
      if (!value.hyperlink) {
        throw new QuireError(`hyperlink target for ${ref} must be a non-empty string`)
      }
      if (value.hyperlink.length > MAX_HYPERLINK_LEN) {
        throw new QuireError(`hyperlink target for ${ref} exceeds ${MAX_HYPERLINK_LEN} characters`)
      }
      this.hyperlinks.set(ref, { target: value.hyperlink, tooltip: value.tooltip })
      stored = value.text ?? value.hyperlink
    } else {
      this.hyperlinks.delete(ref)
      if (isRichText(value) && value.length === 0) {
        throw new QuireError(`rich text value for ${ref} must have at least one run`)
      }
      stored = value
    }
    let line = this.rows.get(row)
    if (!line) {
      line = new Map()
      this.rows.set(row, line)
    }
    line.set(col, { value: stored, style })
    this.maxRow = Math.max(this.maxRow, row)
    this.maxCol = Math.max(this.maxCol, col)
  }

  /** @internal — the `_rels/sheetN.xml.rels` part for this sheet, if it needs one. */
  relsXml(): string | undefined {
    return this.rels.length ? worksheetRelsXml(this.rels) : undefined
  }

  /** @internal */
  hasComments(): boolean {
    return this.comments.size > 0
  }

  /** @internal — this sheet's `xl/commentsN.xml` part, if it has any comments. */
  commentsXmlPart(): string | undefined {
    return commentsXmlOf(this.comments)
  }

  /** @internal — this sheet's `xl/drawings/vmlDrawingN.vml` part, if it has any comments. */
  vmlDrawingXmlPart(): string | undefined {
    return vmlDrawingXmlOf(this.comments)
  }

  /** @internal */
  serialize(sst: SharedStrings, pool: StylePool, sheetIndex: number): string {
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
        if (meta?.hidden) attrs.push('hidden="1"')
        if (meta?.outlineLevel) attrs.push(`outlineLevel="${meta.outlineLevel}"`)
        // Also carried on the row itself (not just baked into each cell's own `s`) so
        // cells with no `<c>` element at all still show the row's default style.
        if (rowStyle && !isEmptyStyle(rowStyle)) {
          attrs.push(`s="${this.pool.intern(rowStyle)}"`, 'customFormat="1"')
        }
        return `<row ${attrs.join(' ')}>${cells}</row>`
      })
      .join('')

    const { hyperlinksXml, legacyDrawing } = this.buildRelsAndRefs(sheetIndex)

    const sheetPr = sheetPrXmlOf(this.pageSetup, this.outline)
    const { pageMargins, pageSetup } = pageSetupXmlOf(this.pageSetup)

    return worksheetXml({
      sheetPr,
      dimension: this.maxRow > 0 ? `A1:${toRef(this.maxRow, this.maxCol)}` : 'A1',
      sheetViews: this.frozen ? sheetViewsXml(this.frozen) : undefined,
      cols: this.colsXml(),
      rows,
      sheetProtection: sheetProtectionXmlOf(this.protection),
      autoFilter: this.filterRange ? `<autoFilter ref="${this.filterRange}"/>` : undefined,
      mergeCells: this.merges.length
        ? `<mergeCells count="${this.merges.length}">` +
          this.merges.map((r) => `<mergeCell ref="${r}"/>`).join('') +
          `</mergeCells>`
        : undefined,
      conditionalFormatting: this.conditionalFormatsXml(),
      dataValidations: this.dataValidationsXml(),
      hyperlinks: hyperlinksXml,
      pageMargins,
      pageSetup,
      legacyDrawing,
    })
  }

  private dataValidationsXml(): string | undefined {
    return dataValidationsXmlOf(this.validations)
  }

  private conditionalFormatsXml(): string | undefined {
    return conditionalFormatsXmlOf(this.conditionalFormats, this.pool)
  }

  /**
   * Rebuilt fresh every `serialize()` call so repeated `xlsx()`/`blob()` calls agree: fills
   * `this.rels` (the sheet's one `_rels/sheetN.xml.rels`) and returns the two worksheet-body
   * fragments that reference those rel ids (`<hyperlinks>` and `<legacyDrawing>`).
   */
  private buildRelsAndRefs(sheetIndex: number): {
    hyperlinksXml?: string
    legacyDrawing?: string
  } {
    const built = buildRelsAndRefsOf(this.hyperlinks, this.comments.size > 0, sheetIndex)
    this.rels = built.rels
    return { hyperlinksXml: built.hyperlinksXml, legacyDrawing: built.legacyDrawing }
  }

  private colsXml(): string | undefined {
    return colsXmlOf(this.columns, this.pool)
  }

  private renderCell(
    ref: string,
    cell: StoredCell,
    col: number,
    rowStyle: CellStyle | undefined,
  ): string {
    const colStyle = this.columns.get(col)?.style
    const style = mergeStyle(mergeStyle(colStyle, rowStyle), cell.style)
    return renderCellXml(ref, cell.value, style, this.sst, this.pool)
  }
}

/** Renders `<cols>…</cols>`. Shared with the streaming writer. */
export function colsXmlOf(columns: Map<number, ColumnSpec>, pool: StylePool): string | undefined {
  if (columns.size === 0) return undefined
  const cols = [...columns.keys()]
    .sort((a, b) => a - b)
    .map((idx) => {
      const spec = columns.get(idx)!
      const attrs = [`min="${idx}"`, `max="${idx}"`]
      if (spec.width != null) attrs.push(`width="${spec.width}"`, 'customWidth="1"')
      if (spec.hidden) attrs.push('hidden="1"')
      if (spec.outlineLevel) attrs.push(`outlineLevel="${spec.outlineLevel}"`)
      if (spec.style && !isEmptyStyle(spec.style)) {
        attrs.push(`style="${pool.intern(spec.style)}"`)
      }
      return `<col ${attrs.join(' ')}/>`
    })
    .join('')
  return `<cols>${cols}</cols>`
}

/**
 * Renders one `<c>` element. Shared between the buffered writer above and the
 * streaming writer (`streaming.ts`) so the two never drift on cell semantics.
 */
export function renderCellXml(
  ref: string,
  value: StoredScalar,
  style: CellStyle | undefined,
  sst: SharedStrings,
  pool: StylePool,
): string {
  if (isFormula(value)) {
    const s = pool.intern(style)
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
    const s = pool.intern(style, { isDate: true })
    return `<c r="${ref}" s="${s}"><v>${numToXml(dateToSerial(value))}</v></c>`
  }

  const s = pool.intern(style)
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
  return `<c r="${ref}"${sAttr} t="s"><v>${sst.intern(value)}</v></c>`
}

/** Renders `<dataValidations>…</dataValidations>`. Shared with the streaming writer. */
export function dataValidationsXmlOf(
  validations: readonly { range: string; rule: DataValidationRule; options: DataValidationOptions }[],
): string | undefined {
  if (validations.length === 0) return undefined
  const items = validations.map(({ range, rule, options }) => {
    const attrs =
      `type="${rule.type}" allowBlank="${options.allowBlank === false ? 0 : 1}"` +
      (rule.type === 'list' || rule.type === 'custom' ? '' : attr('operator', rule.operator)) +
      ' showInputMessage="1" showErrorMessage="1"' +
      attr('sqref', range) +
      attr('promptTitle', options.promptTitle) +
      attr('prompt', options.promptMessage) +
      attr('errorTitle', options.errorTitle) +
      attr('error', options.errorMessage)
    const formulaXml = renderValidationFormulas(rule)
      .map((f, i) => `<formula${i + 1}>${escapeText(f)}</formula${i + 1}>`)
      .join('')
    return `<dataValidation ${attrs}>${formulaXml}</dataValidation>`
  })
  return `<dataValidations count="${items.length}">${items.join('')}</dataValidations>`
}

const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }

/**
 * Renders the `<sheetPr>` (fit-to-page flag, must precede `<dimension>`) and the
 * `<pageMargins>`/`<pageSetup>` tail fragments (must precede `<legacyDrawing>`). Shared
 * with the streaming writer.
 */
export function pageSetupXmlOf(
  setup: PageSetupOptions | null,
): { pageMargins?: string; pageSetup?: string } {
  if (!setup) return {}
  const m = { ...DEFAULT_MARGINS, ...setup.margins }
  const pageMargins =
    `<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}"` +
    ` header="${m.header}" footer="${m.footer}"/>`

  const psAttrs =
    attr('orientation', setup.orientation) +
    attr('paperSize', setup.paperSize) +
    attr('fitToWidth', setup.fitToWidth) +
    attr('fitToHeight', setup.fitToHeight) +
    (setup.scale != null && setup.fitToWidth == null && setup.fitToHeight == null
      ? attr('scale', setup.scale)
      : '')
  const pageSetup = psAttrs ? `<pageSetup${psAttrs}/>` : undefined

  return { pageMargins, pageSetup }
}

/**
 * Renders `<sheetPr>` — the fit-to-page flag and/or the outline (grouping) direction —
 * or `undefined` if neither is set. Must precede `<dimension>`. Shared with the streaming
 * writer.
 */
export function sheetPrXmlOf(
  pageSetup: PageSetupOptions | null,
  outline: OutlineOptions | undefined,
): string | undefined {
  const outlinePr =
    outline?.summaryBelow !== undefined || outline?.summaryRight !== undefined
      ? `<outlinePr${outline.summaryBelow === false ? ' summaryBelow="0"' : ''}` +
        `${outline.summaryRight === false ? ' summaryRight="0"' : ''}/>`
      : ''
  const pageSetUpPr =
    pageSetup && (pageSetup.fitToWidth != null || pageSetup.fitToHeight != null)
      ? '<pageSetUpPr fitToPage="1"/>'
      : ''
  const inner = outlinePr + pageSetUpPr
  return inner ? `<sheetPr>${inner}</sheetPr>` : undefined
}

/**
 * Renders `<sheetProtection>`, or `undefined` if the sheet isn't protected. Must follow
 * `</sheetData>` and precede `<autoFilter>`. Shared with the streaming writer.
 */
export function sheetProtectionXmlOf(options: SheetProtectionOptions | null): string | undefined {
  if (!options) return undefined
  const attrs = ['sheet="1"']
  if (options.allowSelectLockedCells === false) attrs.push('selectLockedCells="1"')
  if (options.allowSelectUnlockedCells === false) attrs.push('selectUnlockedCells="1"')
  if (options.allowFormatCells) attrs.push('formatCells="0"')
  if (options.allowFormatColumns) attrs.push('formatColumns="0"')
  if (options.allowFormatRows) attrs.push('formatRows="0"')
  if (options.allowInsertColumns) attrs.push('insertColumns="0"')
  if (options.allowInsertRows) attrs.push('insertRows="0"')
  if (options.allowDeleteColumns) attrs.push('deleteColumns="0"')
  if (options.allowDeleteRows) attrs.push('deleteRows="0"')
  if (options.allowSort) attrs.push('sort="0"')
  if (options.allowAutoFilter) attrs.push('autoFilter="0"')
  return `<sheetProtection ${attrs.join(' ')}/>`
}

/**
 * Renders `<workbookProtection>`, or `undefined` if the workbook isn't protected. Must
 * precede `<sheets>`. Shared with the streaming writer.
 */
export function workbookProtectionXmlOf(options: WorkbookProtectionOptions | null): string | undefined {
  if (!options) return undefined
  const attrs: string[] = []
  if (options.lockStructure ?? true) attrs.push('lockStructure="1"')
  if (options.lockWindows) attrs.push('lockWindows="1"')
  return `<workbookProtection${attrs.length ? ' ' + attrs.join(' ') : ''}/>`
}

/** Renders one or more `<conditionalFormatting>` blocks. Shared with the streaming writer. */
export function conditionalFormatsXmlOf(
  conditionalFormats: readonly { range: string; rule: ConditionalFormatRule }[],
  pool: StylePool,
): string | undefined {
  if (conditionalFormats.length === 0) return undefined
  const items = conditionalFormats.map(({ range, rule }, i) => {
    const priority = i + 1
    if (rule.type === 'cellIs') {
      const dxfId = pool.internDxf(rule.style)
      const formulas = Array.isArray(rule.formula) ? rule.formula : [rule.formula]
      const formulaXml = formulas.map((f) => `<formula>${escapeText(String(f))}</formula>`).join('')
      return (
        `<conditionalFormatting${attr('sqref', range)}>` +
        `<cfRule type="cellIs" dxfId="${dxfId}" priority="${priority}"${attr('operator', rule.operator)}>` +
        `${formulaXml}</cfRule></conditionalFormatting>`
      )
    }
    if (rule.type === 'colorScale') {
      const stops = rule.colors.length === 2 ? (['min', 'max'] as const) : (['min', 'percentile', 'max'] as const)
      const cfvoXml = stops.map((t) => `<cfvo type="${t}"${t === 'percentile' ? ' val="50"' : ''}/>`).join('')
      const colorXml = rule.colors.map((c) => `<color rgb="${toArgb(c)}"/>`).join('')
      return (
        `<conditionalFormatting${attr('sqref', range)}>` +
        `<cfRule type="colorScale" priority="${priority}"><colorScale>${cfvoXml}${colorXml}</colorScale></cfRule>` +
        `</conditionalFormatting>`
      )
    }
    if (rule.type === 'dataBar') {
      return (
        `<conditionalFormatting${attr('sqref', range)}>` +
        `<cfRule type="dataBar" priority="${priority}"><dataBar>` +
        `<cfvo type="min"/><cfvo type="max"/><color rgb="${toArgb(rule.color)}"/>` +
        `</dataBar></cfRule></conditionalFormatting>`
      )
    }
    if (rule.type === 'iconSet') {
      const n = Number(rule.iconSet[0])
      const step = 100 / n
      const cfvoXml = Array.from(
        { length: n },
        (_, i) => `<cfvo type="percent" val="${Math.round(i * step)}"/>`,
      ).join('')
      return (
        `<conditionalFormatting${attr('sqref', range)}>` +
        `<cfRule type="iconSet" priority="${priority}"><iconSet iconSet="${rule.iconSet}">` +
        `${cfvoXml}</iconSet></cfRule></conditionalFormatting>`
      )
    }
    // top10
    const dxfId = pool.internDxf(rule.style)
    const attrs =
      `type="top10" dxfId="${dxfId}" priority="${priority}" rank="${rule.rank}"` +
      (rule.percent ? ' percent="1"' : '') +
      (rule.bottom ? ' bottom="1"' : '')
    return `<conditionalFormatting${attr('sqref', range)}><cfRule ${attrs}/></conditionalFormatting>`
  })
  return items.join('')
}

/** This sheet's `xl/commentsN.xml` part, if it has any comments. Shared with the streaming writer. */
export function commentsXmlOf(comments: ReadonlyMap<string, { text: string; author: string }>): string | undefined {
  if (comments.size === 0) return undefined
  const authorIds = new Map<string, number>()
  const authors: string[] = []
  const items = [...comments.entries()].map(([ref, c]) => {
    let id = authorIds.get(c.author)
    if (id === undefined) {
      id = authors.length
      authors.push(c.author)
      authorIds.set(c.author, id)
    }
    return { ref, authorId: id, text: c.text }
  })
  return commentsXml(authors, items)
}

/** This sheet's `xl/drawings/vmlDrawingN.vml` part, if it has any comments. Shared with the streaming writer. */
export function vmlDrawingXmlOf(comments: ReadonlyMap<string, { text: string; author: string }>): string | undefined {
  if (comments.size === 0) return undefined
  const shapes = [...comments.keys()].map((ref, i) => {
    const { row, col } = parseRef(ref)
    return { id: i, row: row - 1, col: col - 1 }
  })
  return vmlDrawingXml(shapes)
}

/**
 * Builds the sheet's `_rels/sheetN.xml.rels` relationship list and the two worksheet-body
 * fragments that reference those rel ids (`<hyperlinks>` and `<legacyDrawing>`). Shared with
 * the streaming writer.
 */
export function buildRelsAndRefsOf(
  hyperlinks: ReadonlyMap<string, { target: string; tooltip?: string }>,
  hasComments: boolean,
  sheetIndex: number,
): { hyperlinksXml?: string; legacyDrawing?: string; rels: WorksheetRelEntry[] } {
  const rels: WorksheetRelEntry[] = []
  let n = 1
  const nextId = (): string => `rId${n++}`

  let hyperlinksXml: string | undefined
  if (hyperlinks.size > 0) {
    const items: string[] = []
    for (const [ref, hl] of hyperlinks) {
      const id = nextId()
      rels.push({ id, type: REL_HYPERLINK, target: hl.target, external: true })
      items.push(`<hyperlink r:id="${id}"${attr('ref', ref)}${attr('tooltip', hl.tooltip)}/>`)
    }
    hyperlinksXml = `<hyperlinks>${items.join('')}</hyperlinks>`
  }

  let legacyDrawing: string | undefined
  if (hasComments) {
    const commentsId = nextId()
    rels.push({ id: commentsId, type: REL_COMMENTS, target: `../comments${sheetIndex}.xml` })
    const vmlId = nextId()
    rels.push({ id: vmlId, type: REL_VML, target: `../drawings/vmlDrawing${sheetIndex}.vml` })
    legacyDrawing = `<legacyDrawing r:id="${vmlId}"/>`
  }

  return { hyperlinksXml, legacyDrawing, rels }
}

export function sheetViewsXml(f: FreezeOptions): string {
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
  /**
   * Define a named range, e.g. `wb.defineName('SalesRange', 'Sales', 'A1:B10')`. `sheetName`
   * must already have been added. Names follow Excel's identifier rules (letters/digits/`_`/
   * `.`, can't look like a cell reference) and must be unique. Workbook-scoped (visible
   * everywhere) by default; pass `{ scope: 'SheetName' }` to limit visibility to one sheet
   * (which must also already exist) — Excel then requires the sheet-qualified form
   * (`SheetName!SalesRange`) from any other sheet.
   */
  defineName(name: string, sheetName: string, range: string, options?: DefineNameOptions): this
  /**
   * Protect the workbook structurally — locks the sheet list and/or window (no password —
   * see {@link WorkbookProtectionOptions}). Chainable.
   */
  protect(options?: WorkbookProtectionOptions): this
  /** Serialise to an in-memory `.xlsx` byte array. */
  xlsx(): Uint8Array
  /** Serialise to a `Blob` (browser convenience). */
  blob(): Blob
}

class QuireWorkbook implements Workbook {
  private sheets: QuireWorksheet[] = []
  private definedNames: DefinedNameEntry[] = []
  private definedNameKeys = new Set<string>()
  private protection: WorkbookProtectionOptions | null = null

  addWorksheet(name: string, options: WorksheetOptions = {}): Worksheet {
    validateSheetName(
      name,
      this.sheets.map((s) => s.name),
    )
    const sheet = new QuireWorksheet(name, options)
    this.sheets.push(sheet)
    return sheet
  }

  defineName(name: string, sheetName: string, range: string, options: DefineNameOptions = {}): this {
    validateDefinedName(name, this.definedNameKeys)
    if (!this.sheets.some((s) => s.name === sheetName)) {
      throw new QuireError(`defineName: no worksheet named "${sheetName}"`)
    }
    let localSheetId: number | undefined
    if (options.scope !== undefined) {
      localSheetId = this.sheets.findIndex((s) => s.name === options.scope)
      if (localSheetId < 0) {
        throw new QuireError(`defineName: no worksheet named "${options.scope}" to scope "${name}" to`)
      }
    }
    validateNamedRangeTarget(range)
    this.definedNameKeys.add(name.toLowerCase())
    this.definedNames.push({ name, sheetName, range, localSheetId })
    return this
  }

  protect(options: WorkbookProtectionOptions = {}): this {
    this.protection = options
    return this
  }

  xlsx(): Uint8Array {
    if (this.sheets.length === 0) {
      throw new QuireError('a workbook needs at least one worksheet')
    }

    // Fresh per call — `xlsx()` / `blob()` are pure and repeatable.
    const sst = new SharedStrings()
    const pool = new StylePool()

    // Serialise sheets first — this populates the shared-string table and style pool.
    const sheetXmls = this.sheets.map((s, i) => s.serialize(sst, pool, i + 1))
    const hasStrings = sst.list.length > 0
    const created = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
    const commentSheets = this.sheets
      .map((s, i) => (s.hasComments() ? i + 1 : -1))
      .filter((n) => n > 0)

    const parts: Record<string, string> = {
      '[Content_Types].xml': contentTypesXml(this.sheets.length, hasStrings, commentSheets),
      '_rels/.rels': rootRelsXml(),
      'docProps/core.xml': coreXml(created),
      'docProps/app.xml': appXml(),
      'xl/workbook.xml': workbookXml(
        this.sheets.map((s) => s.name),
        this.sheets.map((s) => s.getFilterRange()),
        this.definedNames,
        this.sheets.map((s) => s.getPrintArea()),
        workbookProtectionXmlOf(this.protection),
      ),
      'xl/_rels/workbook.xml.rels': workbookRelsXml(this.sheets.length, hasStrings),
      'xl/styles.xml': pool.toXml(),
    }
    sheetXmls.forEach((xml, i) => {
      parts[`xl/worksheets/sheet${i + 1}.xml`] = xml
    })
    this.sheets.forEach((s, i) => {
      const rels = s.relsXml()
      if (rels) parts[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = rels
      const commentsPart = s.commentsXmlPart()
      if (commentsPart) parts[`xl/comments${i + 1}.xml`] = commentsPart
      const vmlPart = s.vmlDrawingXmlPart()
      if (vmlPart) parts[`xl/drawings/vmlDrawing${i + 1}.vml`] = vmlPart
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
