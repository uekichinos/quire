import { parseXml } from './xml-read'

/* -------------------------------------------------------------------------- */
/*  Public shapes (mirror the writer's CellStyle so styles round-trip)         */
/* -------------------------------------------------------------------------- */

export interface ReadFont {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  /** 8-digit ARGB, e.g. `'FFFF0000'`. */
  color?: string
}

export interface ReadBorderEdge {
  style?: string
  color?: string
}

export interface ReadStyle {
  font?: ReadFont
  /** Solid fill colour as 8-digit ARGB. */
  fill?: string
  border?: {
    top?: ReadBorderEdge
    right?: ReadBorderEdge
    bottom?: ReadBorderEdge
    left?: ReadBorderEdge
  }
  align?: {
    horizontal?: string
    vertical?: string
    wrapText?: boolean
    indent?: number
  }
  /** Resolved number-format code (`'0.00'`, `'yyyy-mm-dd'`, …). */
  numFmt?: string
}

export interface StyleSheet {
  /** `cellXfs[i]` → numFmtId — always populated (cheap; used for date detection). */
  xfNumFmtId: number[]
  /** Custom format codes by id (≥ 164). */
  customCode: Map<number, string>
  /** `cellXfs[i]` → resolved `ReadStyle` — only when `withStyles` was set. */
  xfStyle: ReadStyle[]
}

/* -------------------------------------------------------------------------- */
/*  Colour resolution                                                          */
/* -------------------------------------------------------------------------- */

/** Legacy indexed colour palette (ECMA-376 §18.8.27). Undefined = "not standard". */
const INDEXED: (string | undefined)[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

/** `styles.xml` theme index → key in the theme's `<clrScheme>`, with the 0/1 & 2/3 swap. */
const THEME_ORDER = [
  'lt1', 'dk1', 'lt2', 'dk2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
] as const

interface ColorRef {
  rgb?: string
  indexed?: number
  theme?: number
  tint?: number
  auto?: boolean
}

function colorRef(attrs: Record<string, string>): ColorRef | undefined {
  const tint = attrs.tint !== undefined ? Number(attrs.tint) : undefined
  if (attrs.rgb) return { rgb: attrs.rgb.toUpperCase(), tint }
  if (attrs.indexed !== undefined) return { indexed: Number(attrs.indexed), tint }
  if (attrs.theme !== undefined) return { theme: Number(attrs.theme), tint }
  if (attrs.auto === '1') return { auto: true }
  return undefined
}

function applyTint(argb: string, tint: number): string {
  const alpha = argb.slice(0, 2)
  const ch = (hex: string): number => {
    const c = parseInt(hex, 16)
    const v = tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint
    return Math.max(0, Math.min(255, Math.round(v)))
  }
  return (
    alpha +
    [ch(argb.slice(2, 4)), ch(argb.slice(4, 6)), ch(argb.slice(6, 8))]
      .map((x) => x.toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  )
}

class ColorResolver {
  /** ARGB by THEME_ORDER index; `''` when the theme didn't define it. */
  constructor(private readonly theme: string[]) {}

  resolve(ref: ColorRef | undefined): string | undefined {
    if (!ref || ref.auto) return undefined
    let base: string | undefined

    if (ref.rgb) {
      base = ref.rgb.length === 6 ? `FF${ref.rgb}` : ref.rgb
    } else if (ref.indexed !== undefined) {
      if (ref.indexed === 64 || ref.indexed === 65) return undefined // system fg/bg
      const c = INDEXED[ref.indexed]
      base = c ? `FF${c}` : undefined
    } else if (ref.theme !== undefined) {
      base = this.theme[ref.theme] || undefined
    }

    if (!base) return undefined
    return ref.tint ? applyTint(base, ref.tint) : base
  }
}

function localName(name: string): string {
  const i = name.indexOf(':')
  return i === -1 ? name : name.slice(i + 1)
}

/** Reads a theme part's `<clrScheme>` into ARGB entries ordered by THEME_ORDER. */
function parseTheme(xml: string | undefined): string[] {
  if (!xml) return []
  const scheme: Record<string, string> = Object.create(null)
  let key: string | null = null
  const keys = new Set<string>([...THEME_ORDER, 'dk1', 'lt1', 'dk2', 'lt2'])

  parseXml(xml, {
    onOpen(name, attrs) {
      const ln = localName(name)
      if (keys.has(ln)) {
        key = ln
      } else if (key && (ln === 'srgbClr' || ln === 'sysClr')) {
        const val = ln === 'sysClr' ? (attrs.lastClr ?? attrs.val) : attrs.val
        if (val && !(key in scheme)) scheme[key] = `FF${val.toUpperCase()}`
        key = null
      }
    },
    onClose(name) {
      if (localName(name) === key) key = null
    },
  })

  return THEME_ORDER.map((k) => scheme[k] ?? '')
}

/* -------------------------------------------------------------------------- */
/*  styles.xml                                                                 */
/* -------------------------------------------------------------------------- */

const BUILTIN_CODE: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0',
  48: '##0.0E+0', 49: '@',
}

export interface ParseStyleOptions {
  withStyles: boolean
  themeXml?: string
}

export function parseStyleSheet(xml: string, options: ParseStyleOptions): StyleSheet {
  const customCode = new Map<number, string>()
  const xfNumFmtId: number[] = []
  const xfStyle: ReadStyle[] = []

  const fonts: ReadFont[] = []
  const fills: (string | undefined)[] = []
  const borders: NonNullable<ReadStyle['border']>[] = []

  const colors = new ColorResolver(parseTheme(options.themeXml))

  // section flags
  let sec: '' | 'fonts' | 'fills' | 'borders' | 'cellXfs' = ''
  let curFont: ReadFont | null = null
  let curFill: string | undefined
  let curFillIsSolid = false
  let curBorder: NonNullable<ReadStyle['border']> | null = null
  let curEdge: keyof NonNullable<ReadStyle['border']> | null = null
  let curXf: { style: ReadStyle; applyNumberFormat: boolean } | null = null

  const codeFor = (id: number): string | undefined => customCode.get(id) ?? BUILTIN_CODE[id]

  parseXml(xml, {
    onOpen(name, attrs) {
      const ln = localName(name)
      switch (ln) {
        case 'numFmt': {
          const id = Number(attrs.numFmtId)
          if (Number.isInteger(id) && attrs.formatCode !== undefined) {
            customCode.set(id, attrs.formatCode)
          }
          return
        }
        case 'fonts':
          sec = 'fonts'
          return
        case 'fills':
          sec = 'fills'
          return
        case 'borders':
          sec = 'borders'
          return
        case 'cellXfs':
          sec = 'cellXfs'
          return
      }

      if (!options.withStyles) {
        if (sec === 'cellXfs' && ln === 'xf') xfNumFmtId.push(Number(attrs.numFmtId ?? 0) || 0)
        return
      }

      switch (sec) {
        case 'fonts':
          if (ln === 'font') curFont = {}
          else if (curFont) {
            if (ln === 'b') curFont.bold = true
            else if (ln === 'i') curFont.italic = true
            else if (ln === 'u') curFont.underline = true
            else if (ln === 'sz' && attrs.val) curFont.size = Number(attrs.val)
            else if (ln === 'name' && attrs.val) curFont.name = attrs.val
            else if (ln === 'color') {
              const c = colors.resolve(colorRef(attrs))
              if (c) curFont.color = c
            }
          }
          break
        case 'fills':
          if (ln === 'fill') {
            curFill = undefined
            curFillIsSolid = false
          } else if (ln === 'patternFill') {
            curFillIsSolid = attrs.patternType === 'solid'
          } else if (ln === 'fgColor' && curFillIsSolid) {
            curFill = colors.resolve(colorRef(attrs))
          }
          break
        case 'borders':
          if (ln === 'border') curBorder = {}
          else if (curBorder) {
            if (ln === 'left' || ln === 'right' || ln === 'top' || ln === 'bottom') {
              curEdge = ln
              if (attrs.style) curBorder[ln] = { style: attrs.style }
            } else if (ln === 'color' && curEdge && curBorder[curEdge]) {
              const c = colors.resolve(colorRef(attrs))
              if (c) curBorder[curEdge]!.color = c
            }
          }
          break
        case 'cellXfs':
          if (ln === 'xf') {
            const numFmtId = Number(attrs.numFmtId ?? 0) || 0
            xfNumFmtId.push(numFmtId)
            const style: ReadStyle = {}
            const fontId = Number(attrs.fontId ?? 0)
            const fillId = Number(attrs.fillId ?? 0)
            const borderId = Number(attrs.borderId ?? 0)
            // id 0 is the sheet default — only surface an explicitly non-default one
            if (fontId > 0 && attrs.applyFont !== '0') {
              const f = fonts[fontId]
              if (f && Object.keys(f).length) style.font = f
            }
            if (fillId > 0 && attrs.applyFill !== '0') {
              const fill = fills[fillId]
              if (fill) style.fill = fill
            }
            if (borderId > 0 && attrs.applyBorder !== '0') {
              const b = borders[borderId]
              if (b && Object.keys(b).length) style.border = b
            }
            const code = codeFor(numFmtId)
            if (numFmtId !== 0 && code) style.numFmt = code
            curXf = { style, applyNumberFormat: attrs.applyNumberFormat !== '0' }
          } else if (ln === 'alignment' && curXf) {
            const a: NonNullable<ReadStyle['align']> = {}
            if (attrs.horizontal) a.horizontal = attrs.horizontal
            if (attrs.vertical) a.vertical = attrs.vertical
            if (attrs.wrapText === '1') a.wrapText = true
            if (attrs.indent) a.indent = Number(attrs.indent)
            if (Object.keys(a).length) curXf.style.align = a
          }
          break
      }
    },
    onClose(name) {
      const ln = localName(name)
      if (ln === 'fonts' || ln === 'fills' || ln === 'borders' || ln === 'cellXfs') {
        sec = ''
        return
      }
      if (!options.withStyles) return

      if (ln === 'font' && curFont) {
        fonts.push(curFont)
        curFont = null
      } else if (ln === 'fill') {
        fills.push(curFill)
      } else if (ln === 'border' && curBorder) {
        borders.push(curBorder)
        curBorder = null
        curEdge = null
      } else if ((ln === 'left' || ln === 'right' || ln === 'top' || ln === 'bottom') && curEdge === ln) {
        curEdge = null
      } else if (ln === 'xf' && curXf) {
        xfStyle.push(curXf.style)
        curXf = null
      }
    },
  })

  return { xfNumFmtId, customCode, xfStyle }
}
