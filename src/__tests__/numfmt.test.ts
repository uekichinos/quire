import { describe, expect, it } from 'vitest'
import { builtinNumFmtId, isDateNumFmt } from '../numfmt'

describe('builtinNumFmtId', () => {
  it('maps known codes to their ids', () => {
    expect(builtinNumFmtId('General')).toBe(0)
    expect(builtinNumFmtId('0.00')).toBe(2)
    expect(builtinNumFmtId('0%')).toBe(9)
  })
  it('returns undefined for a custom code', () => {
    expect(builtinNumFmtId('yyyy-mm-dd')).toBeUndefined()
  })
})

describe('isDateNumFmt', () => {
  it('recognises the built-in date/time ids', () => {
    for (const id of [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]) {
      expect(isDateNumFmt(id), `id ${id}`).toBe(true)
    }
  })

  it('rejects General and numeric built-ins', () => {
    for (const id of [0, 1, 2, 3, 4, 9, 10, 11, 37, 38, 49]) {
      expect(isDateNumFmt(id), `id ${id}`).toBe(false)
    }
  })

  it('scans custom codes for date/time tokens', () => {
    expect(isDateNumFmt(164, 'yyyy-mm-dd')).toBe(true)
    expect(isDateNumFmt(165, 'd/m/yy h:mm')).toBe(true)
    expect(isDateNumFmt(166, '[h]:mm:ss')).toBe(true)
    expect(isDateNumFmt(167, 'mmm yyyy')).toBe(true)
  })

  it('is not fooled by date letters inside quoted literals or brackets', () => {
    expect(isDateNumFmt(164, '0.00" days"')).toBe(false)
    expect(isDateNumFmt(165, '#,##0" m"')).toBe(false)
    expect(isDateNumFmt(166, '[Red]0.00')).toBe(false)
    expect(isDateNumFmt(167, '[$-409]#,##0')).toBe(false)
    expect(isDateNumFmt(168, '0"h"')).toBe(false)
  })

  it('needs a code for custom ids', () => {
    expect(isDateNumFmt(200)).toBe(false)
  })
})
