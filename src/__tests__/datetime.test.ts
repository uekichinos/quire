import { describe, expect, it } from 'vitest'
import { dateToSerial } from '../datetime'

// Local-time constructors so the result is timezone-independent.
const d = (y: number, m: number, day: number, h = 0, min = 0, s = 0) =>
  new Date(y, m - 1, day, h, min, s)

describe('dateToSerial', () => {
  it.each([
    [d(1900, 1, 1), 1],
    [d(1900, 2, 28), 59],
    // 1900-02-29 does not exist; JS normalises it to 1900-03-01 → serial 61
    [d(1900, 3, 1), 61],
    [d(1901, 1, 1), 367],
    [d(2020, 1, 1), 43831],
    [d(2024, 2, 29), 45351],
  ])('%s → %i', (date, serial) => {
    expect(dateToSerial(date)).toBe(serial)
  })

  it('encodes the time of day as the fraction', () => {
    expect(dateToSerial(d(2020, 1, 1, 6, 0, 0))).toBeCloseTo(43831.25, 6)
    expect(dateToSerial(d(2020, 1, 1, 12, 0, 0))).toBeCloseTo(43831.5, 6)
    expect(dateToSerial(d(2020, 1, 1, 18, 0, 0))).toBeCloseTo(43831.75, 6)
  })

  it('throws for an invalid Date', () => {
    expect(() => dateToSerial(new Date('nope'))).toThrow(/invalid Date/)
  })

  it('throws for dates before 1900-01-01', () => {
    expect(() => dateToSerial(d(1899, 12, 31))).toThrow(/1900/)
  })
})
