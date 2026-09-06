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
