import { describe, expect, it } from 'vitest'
import { colLetter, colNumber, parseRef, toRange, toRef } from '../address'

describe('colLetter', () => {
  it.each([
    [1, 'A'],
    [26, 'Z'],
    [27, 'AA'],
    [52, 'AZ'],
    [53, 'BA'],
    [702, 'ZZ'],
    [703, 'AAA'],
    [16384, 'XFD'], // Excel's last column
  ])('%i -> %s', (n, s) => {
    expect(colLetter(n)).toBe(s)
  })

  it('rejects non-positive or non-integer input', () => {
    expect(() => colLetter(0)).toThrow()
    expect(() => colLetter(-1)).toThrow()
    expect(() => colLetter(1.5)).toThrow()
  })
})

describe('colNumber', () => {
  it.each([
    ['A', 1],
    ['Z', 26],
    ['AA', 27],
    ['ZZ', 702],
    ['AAA', 703],
    ['XFD', 16384],
  ])('%s -> %i', (s, n) => {
    expect(colNumber(s)).toBe(n)
  })

  it('round-trips with colLetter', () => {
    for (const n of [1, 5, 26, 27, 100, 702, 703, 16384]) {
      expect(colNumber(colLetter(n))).toBe(n)
    }
  })

  it('rejects non A-Z input', () => {
    expect(() => colNumber('a1')).toThrow()
    expect(() => colNumber('')).toThrow()
  })
})

describe('parseRef / toRef', () => {
  it('parses A1-style references', () => {
    expect(parseRef('A1')).toEqual({ row: 1, col: 1 })
    expect(parseRef('B3')).toEqual({ row: 3, col: 2 })
    expect(parseRef('AA100')).toEqual({ row: 100, col: 27 })
  })

  it('rejects malformed references', () => {
    for (const bad of ['A0', '$B$3', '3B', 'B', '1', 'b3', 'A 1', '']) {
      expect(() => parseRef(bad), bad).toThrow()
    }
  })

  it('toRef is the inverse of parseRef', () => {
    for (const ref of ['A1', 'B3', 'Z26', 'AA27', 'XFD1048576']) {
      const { row, col } = parseRef(ref)
      expect(toRef(row, col)).toBe(ref)
    }
  })

  it('toRef rejects a non-positive row', () => {
    expect(() => toRef(0, 1)).toThrow()
  })
})

describe('toRange', () => {
  it('joins two corners', () => {
    expect(toRange(1, 1, 3, 2)).toBe('A1:B3')
  })
})
