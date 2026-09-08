import { describe, expect, it } from 'vitest'
import { numToXml } from '../number'

describe('numToXml', () => {
  it('passes plain decimals through', () => {
    expect(numToXml(0)).toBe('0')
    expect(numToXml(42)).toBe('42')
    expect(numToXml(-3.5)).toBe('-3.5')
    expect(numToXml(0.1 + 0.2)).toBe('0.30000000000000004')
  })

  it('normalises -0 to 0', () => {
    expect(numToXml(-0)).toBe('0')
  })

  it('expands large exponents to full integers', () => {
    expect(numToXml(1e21)).toBe('1000000000000000000000')
    expect(numToXml(-1e21)).toBe('-1000000000000000000000')
    expect(numToXml(1.23e21)).toBe('1230000000000000000000')
    expect(numToXml(5e30)).toBe('5' + '0'.repeat(30))
  })

  it('expands small exponents to full fractions', () => {
    expect(numToXml(1e-7)).toBe('0.0000001')
    expect(numToXml(1.5e-7)).toBe('0.00000015')
    expect(numToXml(-2e-10)).toBe('-0.0000000002')
    expect(numToXml(1e-21)).toBe('0.' + '0'.repeat(20) + '1')
  })

  it('round-trips through Number()', () => {
    for (const n of [1e21, 1e-7, 1.23e21, 9.99e-9, 7e15, -4.2e18]) {
      expect(Number(numToXml(n))).toBe(n)
    }
  })
})
