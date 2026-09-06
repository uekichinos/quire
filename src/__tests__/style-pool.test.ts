import { describe, expect, it } from 'vitest'
import { StylePool } from '../style-pool'
import { toArgb } from '../style'

describe('toArgb', () => {
  it('pads 6-digit hex with an opaque alpha', () => {
    expect(toArgb('ff0000')).toBe('FFFF0000')
    expect(toArgb('#00FF00')).toBe('FF00FF00')
  })
  it('passes 8-digit ARGB through, upper-cased', () => {
    expect(toArgb('80ffffff')).toBe('80FFFFFF')
  })
  it('rejects nonsense', () => {
    expect(() => toArgb('red')).toThrow()
    expect(() => toArgb('12345')).toThrow()
  })
})

describe('StylePool.intern', () => {
  it('returns 0 for an empty or undefined non-date style', () => {
    const p = new StylePool()
    expect(p.intern(undefined)).toBe(0)
    expect(p.intern({})).toBe(0)
    expect(p.intern({ font: {} })).toBe(0)
  })

  it('de-duplicates identical styles', () => {
    const p = new StylePool()
    const a = p.intern({ font: { bold: true }, fill: 'FFFF00' })
    const b = p.intern({ font: { bold: true }, fill: 'FFFF00' })
    expect(a).toBe(b)
    expect(a).toBeGreaterThan(0)
  })

  it('separates distinct styles', () => {
    const p = new StylePool()
    const bold = p.intern({ font: { bold: true } })
    const italic = p.intern({ font: { italic: true } })
    const red = p.intern({ fill: 'FF0000' })
    expect(new Set([bold, italic, red]).size).toBe(3)
  })

  it('shares sub-pools across cell formats', () => {
    const p = new StylePool()
    p.intern({ font: { bold: true }, fill: 'FF0000' })
    p.intern({ font: { bold: true }, fill: '00FF00' }) // same font, different fill
    const xml = p.toXml()
    // 1 default + 1 bold font
    expect(xml).toContain('<fonts count="2">')
    // 2 reserved + 2 custom fills
    expect(xml).toContain('<fills count="4">')
  })

  it('assigns custom number-format ids from 164 and emits <numFmt>', () => {
    const p = new StylePool()
    p.intern({ numFmt: 'yyyy-mm-dd' })
    p.intern({ numFmt: '0.000"kg"' })
    const xml = p.toXml()
    expect(xml).toContain('<numFmts count="2">')
    expect(xml).toContain('numFmtId="164" formatCode="yyyy-mm-dd"')
    expect(xml).toContain('numFmtId="165" formatCode="0.000&quot;kg&quot;"')
  })

  it('uses built-in ids without emitting <numFmt>', () => {
    const p = new StylePool()
    const s = p.intern({ numFmt: '0.00' }) // built-in id 2
    const xml = p.toXml()
    expect(xml).not.toContain('<numFmts')
    expect(xml).toContain(`<cellXfs count="2">`)
    expect(xml).toMatch(/<xf numFmtId="2"[^>]*applyNumberFormat="1"/)
    expect(s).toBe(1)
  })

  it('gives Date cells the default date format when none is set', () => {
    const p = new StylePool()
    const s = p.intern(undefined, { isDate: true })
    expect(s).toBeGreaterThan(0)
    expect(p.toXml()).toContain('formatCode="yyyy-mm-dd"')
  })

  it('renders alignment', () => {
    const p = new StylePool()
    p.intern({ align: { horizontal: 'center', vertical: 'middle', wrapText: true } })
    const xml = p.toXml()
    expect(xml).toContain('applyAlignment="1"')
    expect(xml).toContain('<alignment horizontal="center" vertical="center" wrapText="1"/>')
  })

  it('renders borders, defaulting the colour to black', () => {
    const p = new StylePool()
    p.intern({ border: { all: { style: 'thin' } } })
    const xml = p.toXml()
    expect(xml).toContain('<borders count="2">')
    expect(xml).toContain(
      '<left style="thin"><color rgb="FF000000"/></left><right style="thin">',
    )
  })

  it('rejects an unknown border style', () => {
    const p = new StylePool()
    // @ts-expect-error — testing runtime guard
    expect(() => p.intern({ border: { all: { style: 'zigzag' } } })).toThrow(/border style/)
  })

  it('always emits the two reserved fills and the empty border', () => {
    const xml = new StylePool().toXml()
    expect(xml).toContain('<fills count="2">')
    expect(xml).toContain('patternType="none"')
    expect(xml).toContain('patternType="gray125"')
    expect(xml).toContain('<borders count="1">')
    expect(xml).toContain('<cellXfs count="1">')
  })
})
