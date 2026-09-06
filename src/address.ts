/**
 * A1-notation helpers. Column and row numbers are **1-based** throughout,
 * matching the OOXML on-disk format.
 *
 * Adapted from ExcelJS `lib/utils/col-cache.js` (MIT).
 */

const A = 'A'.charCodeAt(0)

/** `1 → "A"`, `26 → "Z"`, `27 → "AA"`, `702 → "ZZ"`, `703 → "AAA"`. */
export function colLetter(col: number): string {
  if (!Number.isInteger(col) || col < 1) {
    throw new Error(`@uekichinos/quire: column number must be a positive integer, got ${col}`)
  }
  let n = col
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(A + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** `"A" → 1`, `"Z" → 26`, `"AA" → 27`. */
export function colNumber(letters: string): number {
  if (!/^[A-Z]+$/.test(letters)) {
    throw new Error(`@uekichinos/quire: invalid column "${letters}"`)
  }
  let n = 0
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - A + 1)
  }
  return n
}

const REF_RE = /^([A-Z]+)([1-9][0-9]*)$/

export interface CellAddress {
  row: number
  col: number
}

/** `"B3" → { row: 3, col: 2 }`. Throws on anything else (`"B0"`, `"$B$3"`, `"3B"`). */
export function parseRef(ref: string): CellAddress {
  const m = REF_RE.exec(ref)
  if (!m) throw new Error(`@uekichinos/quire: invalid cell reference "${ref}"`)
  return { col: colNumber(m[1]!), row: Number(m[2]) }
}

/** `(3, 2) → "B3"`. */
export function toRef(row: number, col: number): string {
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(`@uekichinos/quire: row number must be a positive integer, got ${row}`)
  }
  return colLetter(col) + row
}

/** `"A1:C3"` from two corner addresses (already normalised by the caller). */
export function toRange(top: number, left: number, bottom: number, right: number): string {
  return `${toRef(top, left)}:${toRef(bottom, right)}`
}
