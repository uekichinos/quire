/**
 * Number-format ids. Ids `0`–`163` are built into Excel and must **not** be
 * redeclared in `styles.xml`; custom format codes are assigned ids from `164`.
 *
 * The built-in code → id map below covers the formats Excel actually reuses;
 * anything not listed is treated as custom.
 */

/** Applied to `Date` cells that carry no explicit `numFmt`. */
export const DEFAULT_DATE_FORMAT = 'yyyy-mm-dd'

/** First id available for custom format codes. */
export const FIRST_CUSTOM_NUMFMT_ID = 164

const BUILTIN: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
}

const BUILTIN_BY_CODE = new Map<string, number>(
  Object.entries(BUILTIN).map(([id, code]) => [code, Number(id)]),
)

/** Returns the built-in id for a format code, or `undefined` if it is custom. */
export function builtinNumFmtId(code: string): number | undefined {
  return BUILTIN_BY_CODE.get(code)
}

/** Built-in number-format ids that render a date and/or time. */
const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47])

/**
 * Whether a number-format id / code produces a date or time — used on **read**
 * to decide if a numeric cell should become a `Date`.
 *
 * Built-in ids use a fixed table. For a custom code, scans for a date/time
 * token (`y m d h s`) that is not inside a `"…"` literal, a `[…]` section
 * (colours, locales, `[h]` elapsed), or `\`-escaped.
 */
export function isDateNumFmt(id: number, code?: string): boolean {
  if (id === 0) return false
  if (BUILTIN_DATE_IDS.has(id)) return true
  if (id < FIRST_CUSTOM_NUMFMT_ID || !code) return false

  let inQuote = false
  let inBracket = false
  for (let i = 0; i < code.length; i++) {
    const c = code[i]!
    if (inQuote) {
      if (c === '"') inQuote = false
      continue
    }
    if (inBracket) {
      if (c === ']') inBracket = false
      continue
    }
    if (c === '"') inQuote = true
    else if (c === '[') inBracket = true
    else if (c === '\\') i++ // skip the escaped char
    else if (c === 'y' || c === 'Y' || c === 'd' || c === 'D') return true
    else if (c === 'm' || c === 'M' || c === 'h' || c === 'H' || c === 's' || c === 'S') {
      return true
    }
  }
  return false
}
