import { describe, expect, it } from 'vitest'
import { dateToSerial, serialToDate } from '../datetime'

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

describe('serialToDate', () => {
  it.each([
    [1, d(1900, 1, 1)],
    [59, d(1900, 2, 28)],
    [61, d(1900, 3, 1)],
    [367, d(1901, 1, 1)],
    [43831, d(2020, 1, 1)],
    [45351, d(2024, 2, 29)],
  ])('%i → %s', (serial, date) => {
    expect(serialToDate(serial).getTime()).toBe(date.getTime())
  })

  it('decodes the time fraction', () => {
    expect(serialToDate(43831.25).getTime()).toBe(d(2020, 1, 1, 6).getTime())
    expect(serialToDate(43831.5).getTime()).toBe(d(2020, 1, 1, 12).getTime())
  })

  it('is the inverse of dateToSerial', () => {
    for (const dt of [d(1900, 1, 1), d(2000, 2, 29), d(2024, 3, 1, 9, 30), d(2099, 12, 31)]) {
      expect(serialToDate(dateToSerial(dt)).getTime()).toBe(dt.getTime())
    }
  })

  it('supports the 1904 date system', () => {
    // 1904 serial 0 = 1904-01-01
    expect(serialToDate(0, true).getTime()).toBe(d(1904, 1, 1).getTime())
    expect(serialToDate(367, true).getTime()).toBe(d(1905, 1, 2).getTime())
  })

  it('throws on a non-finite serial', () => {
    expect(() => serialToDate(NaN)).toThrow(/serial/)
  })
})
