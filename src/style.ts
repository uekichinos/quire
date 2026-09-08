/** Public cell-style shape and the helpers for normalising / merging it. */
import { QuireError } from './errors'

export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'

export interface BorderEdge {
  style?: BorderStyle
  /** `'RRGGBB'` or `'AARRGGBB'`. Default black. */
  color?: string
}

export interface FontStyle {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  /** `'RRGGBB'` or `'AARRGGBB'`. */
  color?: string
}

export interface AlignStyle {
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify'
  vertical?: 'top' | 'middle' | 'bottom'
  wrapText?: boolean
  indent?: number
}

export interface BorderStyleInput {
  top?: BorderEdge
  right?: BorderEdge
  bottom?: BorderEdge
  left?: BorderEdge
  /** Shorthand — applied to every edge that has no explicit value. */
  all?: BorderEdge
}

export interface CellStyle {
  font?: FontStyle
  /** Solid fill colour, `'RRGGBB'` or `'AARRGGBB'`. */
  fill?: string
  align?: AlignStyle
  border?: BorderStyleInput
  /** A format code (`'0.00'`, `'#,##0'`, `'yyyy-mm-dd'`) or a built-in id. */
  numFmt?: string | number
}

/** `'RRGGBB'` / `'AARRGGBB'` / `'#RRGGBB'` → upper-case 8-digit ARGB. */
export function toArgb(color: string): string {
  let c = color.trim().replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{6}$/.test(c)) c = `FF${c}`
  if (!/^[0-9A-F]{8}$/.test(c)) {
    throw new QuireError(`invalid colour "${color}" (expected RRGGBB or AARRGGBB)`)
  }
  return c
}

const VALID_BORDER_STYLES = new Set<BorderStyle>([
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
])

export function assertBorderStyle(style: string): asserts style is BorderStyle {
  if (!VALID_BORDER_STYLES.has(style as BorderStyle)) {
    throw new QuireError(`unknown border style "${style}"`)
  }
}

/**
 * Merges `override` on top of `base`, one nested level deep for `font` /
 * `align` / `border`. Used for cell-over-row-over-column resolution.
 */
export function mergeStyle(
  base: CellStyle | undefined,
  override: CellStyle | undefined,
): CellStyle | undefined {
  if (!base) return override
  if (!override) return base
  return {
    numFmt: override.numFmt ?? base.numFmt,
    fill: override.fill ?? base.fill,
    font: base.font || override.font ? { ...base.font, ...override.font } : undefined,
    align: base.align || override.align ? { ...base.align, ...override.align } : undefined,
    border:
      base.border || override.border ? { ...base.border, ...override.border } : undefined,
  }
}

export function isEmptyStyle(style: CellStyle | undefined): boolean {
  if (!style) return true
  return (
    style.numFmt === undefined &&
    style.fill === undefined &&
    (!style.font || Object.keys(style.font).length === 0) &&
    (!style.align || Object.keys(style.align).length === 0) &&
    (!style.border || Object.keys(style.border).length === 0)
  )
}
