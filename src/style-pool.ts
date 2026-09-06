import { escapeXml } from './xml'
import { builtinNumFmtId, DEFAULT_DATE_FORMAT, FIRST_CUSTOM_NUMFMT_ID } from './numfmt'
import {
  assertBorderStyle,
  isEmptyStyle,
  toArgb,
  type AlignStyle,
  type BorderEdge,
  type BorderStyleInput,
  type CellStyle,
  type FontStyle,
} from './style'

interface ResolvedAlign {
  horizontal?: string
  vertical?: string // 'middle' → 'center' for OOXML
  wrapText?: boolean
  indent?: number
}

interface ResolvedXf {
  numFmtId: number
  fontId: number
  fillId: number
  borderId: number
  align?: ResolvedAlign
}

const V_MAP: Record<string, string> = { top: 'top', middle: 'center', bottom: 'bottom' }

/**
 * Deduplicates every distinct font / fill / border / number-format / cell-format
 * and hands back a single `cellXfs` index per style. `styles.xml` is rendered
 * from the interned pools.
 */
export class StylePool {
  private numFmts: { id: number; code: string }[] = []
  private numFmtByCode = new Map<string, number>()
  private nextCustomId = FIRST_CUSTOM_NUMFMT_ID

  private fonts: string[] = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>']
  private fontByKey = new Map<string, number>([['', 0]])

  // fill 0 = none, fill 1 = gray125 are reserved by Excel.
  private fills: string[] = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ]
  private fillByKey = new Map<string, number>()

  private borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>']
  private borderByKey = new Map<string, number>([['', 0]])

  private xfs: ResolvedXf[] = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }]
  private xfByKey = new Map<string, number>([['0|0|0|0|', 0]])

  /** cellXfs index for a style. `0` for an empty/undefined non-date style. */
  intern(style: CellStyle | undefined, opts: { isDate?: boolean } = {}): number {
    if (!opts.isDate && isEmptyStyle(style)) return 0

    const numFmtId = this.resolveNumFmt(style?.numFmt, opts.isDate)
    const fontId = this.resolveFont(style?.font)
    const fillId = this.resolveFill(style?.fill)
    const borderId = this.resolveBorder(style?.border)
    const align = this.resolveAlign(style?.align)

    const key = `${numFmtId}|${fontId}|${fillId}|${borderId}|${align ? alignKey(align) : ''}`
    let idx = this.xfByKey.get(key)
    if (idx === undefined) {
      idx = this.xfs.length
      this.xfs.push({ numFmtId, fontId, fillId, borderId, align })
      this.xfByKey.set(key, idx)
    }
    return idx
  }

  /* --- resolvers ------------------------------------------------------------ */

  private resolveNumFmt(code: string | number | undefined, isDate?: boolean): number {
    if (code === undefined) return isDate ? this.resolveNumFmt(DEFAULT_DATE_FORMAT) : 0
    if (typeof code === 'number') {
      if (!Number.isInteger(code) || code < 0) {
        throw new Error(`@uekichinos/quire: numFmt id must be a non-negative integer`)
      }
      return code
    }
    const builtin = builtinNumFmtId(code)
    if (builtin !== undefined) return builtin
    let id = this.numFmtByCode.get(code)
    if (id === undefined) {
      id = this.nextCustomId++
      this.numFmtByCode.set(code, id)
      this.numFmts.push({ id, code })
    }
    return id
  }

  private resolveFont(font: FontStyle | undefined): number {
    if (!font || Object.keys(font).length === 0) return 0
    const key = JSON.stringify([
      font.name ?? null,
      font.size ?? null,
      !!font.bold,
      !!font.italic,
      !!font.underline,
      font.color ? toArgb(font.color) : null,
    ])
    let idx = this.fontByKey.get(key)
    if (idx === undefined) {
      idx = this.fonts.length
      this.fonts.push(renderFont(font))
      this.fontByKey.set(key, idx)
    }
    return idx
  }

  private resolveFill(fill: string | undefined): number {
    if (!fill) return 0
    const argb = toArgb(fill)
    let idx = this.fillByKey.get(argb)
    if (idx === undefined) {
      idx = this.fills.length
      this.fills.push(
        `<fill><patternFill patternType="solid"><fgColor rgb="${argb}"/><bgColor indexed="64"/></patternFill></fill>`,
      )
      this.fillByKey.set(argb, idx)
    }
    return idx
  }

  private resolveBorder(border: BorderStyleInput | undefined): number {
    if (!border || Object.keys(border).length === 0) return 0
    const edges = {
      left: border.left ?? border.all,
      right: border.right ?? border.all,
      top: border.top ?? border.all,
      bottom: border.bottom ?? border.all,
    }
    const key = JSON.stringify(
      (['left', 'right', 'top', 'bottom'] as const).map((side) => {
        const e = edges[side]
        if (!e?.style) return null
        assertBorderStyle(e.style)
        return [e.style, e.color ? toArgb(e.color) : 'FF000000']
      }),
    )
    if (key === '[null,null,null,null]') return 0
    let idx = this.borderByKey.get(key)
    if (idx === undefined) {
      idx = this.borders.length
      this.borders.push(renderBorder(edges))
      this.borderByKey.set(key, idx)
    }
    return idx
  }

  private resolveAlign(align: AlignStyle | undefined): ResolvedAlign | undefined {
    if (!align || Object.keys(align).length === 0) return undefined
    const out: ResolvedAlign = {}
    if (align.horizontal) out.horizontal = align.horizontal
    if (align.vertical) out.vertical = V_MAP[align.vertical]
    if (align.wrapText) out.wrapText = true
    if (align.indent && align.indent > 0) out.indent = Math.floor(align.indent)
    return Object.keys(out).length ? out : undefined
  }

  /* --- output ------------------------------------------------------------- */

  toXml(): string {
    const parts: string[] = []

    if (this.numFmts.length) {
      parts.push(
        `<numFmts count="${this.numFmts.length}">` +
          this.numFmts
            .map((n) => `<numFmt numFmtId="${n.id}" formatCode="${escapeXml(n.code)}"/>`)
            .join('') +
          `</numFmts>`,
      )
    }
    parts.push(`<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`)
    parts.push(`<fills count="${this.fills.length}">${this.fills.join('')}</fills>`)
    parts.push(`<borders count="${this.borders.length}">${this.borders.join('')}</borders>`)
    parts.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    parts.push(
      `<cellXfs count="${this.xfs.length}">` + this.xfs.map(renderXf).join('') + `</cellXfs>`,
    )
    parts.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      parts.join('') +
      '</styleSheet>'
    )
  }
}

/* --- fragment renderers -------------------------------------------------- */

function renderFont(font: FontStyle): string {
  const bits: string[] = []
  if (font.bold) bits.push('<b/>')
  if (font.italic) bits.push('<i/>')
  if (font.underline) bits.push('<u/>')
  bits.push(`<sz val="${font.size ?? 11}"/>`)
  if (font.color) bits.push(`<color rgb="${toArgb(font.color)}"/>`)
  bits.push(`<name val="${escapeXml(font.name ?? 'Calibri')}"/>`)
  bits.push('<family val="2"/>')
  return `<font>${bits.join('')}</font>`
}

function renderBorder(edges: Record<'left' | 'right' | 'top' | 'bottom', BorderEdge | undefined>): string {
  const edge = (name: string, e: BorderEdge | undefined): string => {
    if (!e?.style) return `<${name}/>`
    const color = e.color ? toArgb(e.color) : 'FF000000'
    return `<${name} style="${e.style}"><color rgb="${color}"/></${name}>`
  }
  return (
    `<border>${edge('left', edges.left)}${edge('right', edges.right)}` +
    `${edge('top', edges.top)}${edge('bottom', edges.bottom)}<diagonal/></border>`
  )
}

function renderXf(xf: ResolvedXf): string {
  const attrs = [
    `numFmtId="${xf.numFmtId}"`,
    `fontId="${xf.fontId}"`,
    `fillId="${xf.fillId}"`,
    `borderId="${xf.borderId}"`,
    'xfId="0"',
  ]
  if (xf.numFmtId !== 0) attrs.push('applyNumberFormat="1"')
  if (xf.fontId !== 0) attrs.push('applyFont="1"')
  if (xf.fillId !== 0) attrs.push('applyFill="1"')
  if (xf.borderId !== 0) attrs.push('applyBorder="1"')
  if (xf.align) attrs.push('applyAlignment="1"')

  if (!xf.align) return `<xf ${attrs.join(' ')}/>`

  const a: string[] = []
  if (xf.align.horizontal) a.push(`horizontal="${xf.align.horizontal}"`)
  if (xf.align.vertical) a.push(`vertical="${xf.align.vertical}"`)
  if (xf.align.wrapText) a.push('wrapText="1"')
  if (xf.align.indent) a.push(`indent="${xf.align.indent}"`)
  return `<xf ${attrs.join(' ')}><alignment ${a.join(' ')}/></xf>`
}

function alignKey(a: ResolvedAlign): string {
  return `${a.horizontal ?? ''}~${a.vertical ?? ''}~${a.wrapText ? 1 : 0}~${a.indent ?? 0}`
}
