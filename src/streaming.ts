/**
 * A low-memory `.xlsx` writer for very large sheets: rows are rendered and DEFLATEd as soon
 * as they're added instead of being held (as cell objects, then as one big XML string) for
 * the whole workbook's lifetime. One pass, forward-only — there is no `setCell`/`setRow`
 * random access, and column widths/freeze panes must be set before the first `addRow()`.
 *
 * Shares its cell/row/tail rendering with the buffered writer (`workbook.ts`) via a set of
 * exported pure functions, so the two never drift on what a given value renders as.
 */
import { Zip, ZipDeflate, strToU8 } from 'fflate'
import { parseRef, toRef } from './address'
import { QuireError } from './errors'
import {
  appXml,
  contentTypesXml,
  coreXml,
  type DefinedNameEntry,
  NS_MAIN,
  NS_R,
  rootRelsXml,
  sharedStringsXml,
  workbookRelsXml,
  workbookXml,
  worksheetRelsXml,
} from './serialize'
import { StylePool } from './style-pool'
import { isEmptyStyle, mergeStyle, type CellStyle } from './style'
import { XML_DECLARATION } from './xml'
import {
  buildRelsAndRefsOf,
  colsXmlOf,
  commentsXmlOf,
  conditionalFormatsXmlOf,
  dataValidationsXmlOf,
  isHyperlinkValue,
  isRichText,
  isStyledCell,
  MAX_COL,
  MAX_HYPERLINK_LEN,
  pageSetupXmlOf,
  parseRange,
  renderCellXml,
  renderValidationFormulas,
  SharedStrings,
  sheetViewsXml,
  validateDefinedName,
  validateNamedRangeTarget,
  validatePageSetup,
  validateSheetName,
  vmlDrawingXmlOf,
  type AddRowOptions,
  type CellInput,
  type CellScalar,
  type ColumnSpec,
  type CommentOptions,
  type ConditionalFormatRule,
  type DataValidationOptions,
  type DataValidationRule,
  type FreezeOptions,
  type PageSetupOptions,
  type StoredScalar,
  type WorksheetOptions,
} from './workbook'

export interface StreamingWorksheet {
  readonly name: string
  /**
   * Append a row and flush it immediately. Rows are 1-based and assigned in call order —
   * there is no going back to an earlier row.
   */
  addRow(values: CellInput[], options?: AddRowOptions): this
  /** Must be called before the first `addRow()` — column metadata precedes the row data. */
  setColumn(index: number, spec: ColumnSpec): this
  /** Must be called before the first `addRow()` — frozen panes precede the row data. */
  freeze(options: FreezeOptions): this
  /** Add filter dropdowns over a header range. Can be called any time before `finish()`. */
  autoFilter(range: string): this
  /** Merge a cell range. Can be called any time before `finish()`. */
  merge(range: string): this
  /** Restrict a range to a dropdown, or a numeric/date/text-length rule. */
  setDataValidation(range: string, rule: DataValidationRule, options?: DataValidationOptions): this
  /** Attach a comment/note to a cell, independent of its value. */
  setComment(ref: string, text: string, options?: CommentOptions): this
  /** Highlight a range that matches a condition, or apply a colour scale. */
  addConditionalFormat(range: string, rule: ConditionalFormatRule): this
  /**
   * Page orientation, paper size, fit-to-page/scale, margins, and print area. Must be called
   * before the first `addRow()` — it renders the `<sheetPr>` fit-to-page flag into the header.
   */
  setPageSetup(options: PageSetupOptions): this
}

export interface StreamingWorkbook {
  addWorksheet(name: string, options?: WorksheetOptions): StreamingWorksheet
  /** Define a workbook-scoped named range. `sheetName` must already have been added. */
  defineName(name: string, sheetName: string, range: string): this
  /** Finalises every sheet and resolves the complete `.xlsx` bytes. */
  finish(): Promise<Uint8Array>
}

class StreamingQuireWorksheet implements StreamingWorksheet {
  private columns = new Map<number, ColumnSpec>()
  private merges: string[] = []
  private frozen: FreezeOptions | null = null
  private filterRange: string | null = null
  private hyperlinks = new Map<string, { target: string; tooltip?: string }>()
  private comments = new Map<string, { text: string; author: string }>()
  private validations: { range: string; rule: DataValidationRule; options: DataValidationOptions }[] = []
  private conditionalFormats: { range: string; rule: ConditionalFormatRule }[] = []
  private pageSetup: PageSetupOptions | null = null
  private nextRow = 1
  private headerWritten = false
  private finished = false

  constructor(
    readonly name: string,
    options: WorksheetOptions,
    private readonly sheetIndex: number,
    private readonly sst: SharedStrings,
    private readonly pool: StylePool,
    private readonly file: ZipDeflate,
  ) {
    options.columns?.forEach((spec, i) => this.setColumn(i + 1, spec))
    if (options.freeze) this.freeze(options.freeze)
    if (options.autoFilter) this.autoFilter(options.autoFilter)
  }

  setColumn(index: number, spec: ColumnSpec): this {
    if (this.headerWritten) {
      throw new QuireError('setColumn must be called before the first addRow() in streaming mode')
    }
    if (!Number.isInteger(index) || index < 1 || index > MAX_COL) {
      throw new QuireError(`column index must be between 1 and ${MAX_COL}`)
    }
    if (spec.width != null && !(spec.width >= 0)) {
      throw new QuireError('column width must be a non-negative number')
    }
    this.columns.set(index, { ...this.columns.get(index), ...spec })
    return this
  }

  freeze(options: FreezeOptions): this {
    if (this.headerWritten) {
      throw new QuireError('freeze must be called before the first addRow() in streaming mode')
    }
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

  merge(range: string): this {
    const { top, left, bottom, right } = parseRange(range)
    if (top === bottom && left === right) {
      throw new QuireError(`cannot merge a single cell ("${range}")`)
    }
    const norm = `${toRef(top, left)}:${toRef(bottom, right)}`
    for (const existing of this.merges) {
      const e = parseRange(existing)
      const overlaps = left <= e.right && right >= e.left && top <= e.bottom && bottom >= e.top
      if (overlaps) throw new QuireError(`merge "${norm}" overlaps existing merge "${existing}"`)
    }
    this.merges.push(norm)
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

  setPageSetup(options: PageSetupOptions): this {
    if (this.headerWritten) {
      throw new QuireError('setPageSetup must be called before the first addRow() in streaming mode')
    }
    validatePageSetup(options)
    this.pageSetup = options
    return this
  }

  /** @internal — this sheet's print area, for the workbook defined name. */
  getPrintArea(): string | undefined {
    return this.pageSetup?.printArea
  }

  addRow(values: CellInput[], options: AddRowOptions = {}): this {
    if (this.finished) throw new QuireError(`worksheet "${this.name}" has already been finished`)
    if (!Array.isArray(values)) throw new QuireError('addRow expects an array of values')
    this.ensureHeader()

    const r = this.nextRow++
    const rowStyle = options.style
    const cellsXml: string[] = []

    values.forEach((input, i) => {
      const col = i + 1
      const cellStyle = isStyledCell(input) ? input.style : undefined
      const rawValue = (isStyledCell(input) ? input.value : input) as CellScalar
      if (rawValue === null || rawValue === undefined) return

      const ref = toRef(r, col)
      let stored: StoredScalar
      if (isHyperlinkValue(rawValue)) {
        if (!rawValue.hyperlink) {
          throw new QuireError(`hyperlink target for ${ref} must be a non-empty string`)
        }
        if (rawValue.hyperlink.length > MAX_HYPERLINK_LEN) {
          throw new QuireError(`hyperlink target for ${ref} exceeds ${MAX_HYPERLINK_LEN} characters`)
        }
        this.hyperlinks.set(ref, { target: rawValue.hyperlink, tooltip: rawValue.tooltip })
        stored = rawValue.text ?? rawValue.hyperlink
      } else {
        if (isRichText(rawValue) && rawValue.length === 0) {
          throw new QuireError(`rich text value for ${ref} must have at least one run`)
        }
        stored = rawValue
      }

      const colStyle = this.columns.get(col)?.style
      const style = mergeStyle(mergeStyle(colStyle, rowStyle), cellStyle)
      cellsXml.push(renderCellXml(ref, stored, style, this.sst, this.pool))
    })

    const attrs = [`r="${r}"`]
    if (options.height != null) attrs.push(`ht="${options.height}"`, 'customHeight="1"')
    if (rowStyle && !isEmptyStyle(rowStyle)) {
      attrs.push(`s="${this.pool.intern(rowStyle)}"`, 'customFormat="1"')
    }
    this.file.push(strToU8(`<row ${attrs.join(' ')}>${cellsXml.join('')}</row>`), false)
    return this
  }

  private ensureHeader(): void {
    if (this.headerWritten) return
    this.headerWritten = true
    const { sheetPr } = pageSetupXmlOf(this.pageSetup)
    const header =
      `${XML_DECLARATION}\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
      (sheetPr ?? '') +
      (this.frozen ? sheetViewsXml(this.frozen) : '') +
      (colsXmlOf(this.columns, this.pool) ?? '') +
      '<sheetData>'
    this.file.push(strToU8(header), false)
  }

  /** @internal — flushes the tail and ends this sheet's file stream. */
  finalize(): { relsXml?: string; commentsXml?: string; vmlXml?: string } {
    this.ensureHeader() // an all-empty sheet still needs a valid, closed <worksheet>
    this.finished = true

    const { hyperlinksXml, legacyDrawing, rels } = buildRelsAndRefsOf(
      this.hyperlinks,
      this.comments.size > 0,
      this.sheetIndex,
    )
    const { pageMargins, pageSetup } = pageSetupXmlOf(this.pageSetup)

    const tail =
      '</sheetData>' +
      (this.filterRange ? `<autoFilter ref="${this.filterRange}"/>` : '') +
      (this.merges.length
        ? `<mergeCells count="${this.merges.length}">${this.merges
            .map((r) => `<mergeCell ref="${r}"/>`)
            .join('')}</mergeCells>`
        : '') +
      (conditionalFormatsXmlOf(this.conditionalFormats, this.pool) ?? '') +
      (dataValidationsXmlOf(this.validations) ?? '') +
      (hyperlinksXml ?? '') +
      (pageMargins ?? '') +
      (pageSetup ?? '') +
      (legacyDrawing ?? '') +
      '</worksheet>'
    this.file.push(strToU8(tail), true)

    return {
      relsXml: rels.length ? worksheetRelsXml(rels) : undefined,
      commentsXml: commentsXmlOf(this.comments),
      vmlXml: vmlDrawingXmlOf(this.comments),
    }
  }
}

class StreamingQuireWorkbook implements StreamingWorkbook {
  private sheets: StreamingQuireWorksheet[] = []
  private definedNames: DefinedNameEntry[] = []
  private definedNameKeys = new Set<string>()
  private sst = new SharedStrings()
  private pool = new StylePool()
  private chunks: Uint8Array[] = []
  private zip = new Zip((err, data) => {
    if (err) throw err
    this.chunks.push(data)
  })

  addWorksheet(name: string, options: WorksheetOptions = {}): StreamingWorksheet {
    validateSheetName(
      name,
      this.sheets.map((s) => s.name),
    )
    const index = this.sheets.length + 1
    const file = new ZipDeflate(`xl/worksheets/sheet${index}.xml`)
    this.zip.add(file)
    const sheet = new StreamingQuireWorksheet(name, options, index, this.sst, this.pool, file)
    this.sheets.push(sheet)
    return sheet
  }

  defineName(name: string, sheetName: string, range: string): this {
    validateDefinedName(name, this.definedNameKeys)
    if (!this.sheets.some((s) => s.name === sheetName)) {
      throw new QuireError(`defineName: no worksheet named "${sheetName}"`)
    }
    validateNamedRangeTarget(range)
    this.definedNameKeys.add(name.toLowerCase())
    this.definedNames.push({ name, sheetName, range })
    return this
  }

  async finish(): Promise<Uint8Array> {
    if (this.sheets.length === 0) {
      throw new QuireError('a workbook needs at least one worksheet')
    }

    const finalized = this.sheets.map((s) => s.finalize())
    const hasStrings = this.sst.list.length > 0
    const created = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
    const commentSheets = finalized
      .map((f, i) => (f.commentsXml ? i + 1 : -1))
      .filter((n) => n > 0)

    finalized.forEach((f, i) => {
      if (f.relsXml) this.pushWholePart(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, f.relsXml)
      if (f.commentsXml) this.pushWholePart(`xl/comments${i + 1}.xml`, f.commentsXml)
      if (f.vmlXml) this.pushWholePart(`xl/drawings/vmlDrawing${i + 1}.vml`, f.vmlXml)
    })

    this.pushWholePart(
      '[Content_Types].xml',
      contentTypesXml(this.sheets.length, hasStrings, commentSheets),
    )
    this.pushWholePart('_rels/.rels', rootRelsXml())
    this.pushWholePart('docProps/core.xml', coreXml(created))
    this.pushWholePart('docProps/app.xml', appXml())
    this.pushWholePart(
      'xl/workbook.xml',
      workbookXml(
        this.sheets.map((s) => s.name),
        [],
        this.definedNames,
        this.sheets.map((s) => s.getPrintArea()),
      ),
    )
    this.pushWholePart('xl/_rels/workbook.xml.rels', workbookRelsXml(this.sheets.length, hasStrings))
    this.pushWholePart('xl/styles.xml', this.pool.toXml())
    if (hasStrings) {
      this.pushWholePart('xl/sharedStrings.xml', sharedStringsXml(this.sst.list, this.sst.total))
    }

    this.zip.end()

    let total = 0
    for (const c of this.chunks) total += c.length
    const out = new Uint8Array(total)
    let offset = 0
    for (const c of this.chunks) {
      out.set(c, offset)
      offset += c.length
    }
    return out
  }

  private pushWholePart(name: string, content: string): void {
    const file = new ZipDeflate(name)
    this.zip.add(file)
    file.push(strToU8(content), true)
  }
}

/**
 * Creates a low-memory streaming workbook: rows are DEFLATEd and released as soon as
 * they're added, instead of being held in memory for the whole workbook's lifetime like
 * {@link createWorkbook}'s buffered writer. One pass, forward-only per sheet — no
 * `setCell`/`setRow` random access; column widths and freeze panes must be set before the
 * first `addRow()` on that sheet.
 *
 * @example
 * const wb = createStreamingWorkbook()
 * const sheet = wb.addWorksheet('Big')
 * for (let i = 0; i < 1_000_000; i++) sheet.addRow([i, i * 2])
 * const bytes = await wb.finish()
 */
export function createStreamingWorkbook(): StreamingWorkbook {
  return new StreamingQuireWorkbook()
}
