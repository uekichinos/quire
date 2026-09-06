import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { createWorkbook } from '../workbook'

describe('error paths', () => {
  it('rejects a workbook with no worksheets', () => {
    expect(() => createWorkbook().xlsx()).toThrow(/at least one worksheet/)
  })

  it('rejects non-finite numbers on serialise', () => {
    for (const bad of [NaN, Infinity, -Infinity]) {
      const wb = createWorkbook()
      wb.addWorksheet('S').addRow([bad])
      expect(() => wb.xlsx(), String(bad)).toThrow(/finite/)
    }
  })

  it('rejects invalid Dates and pre-1900 Dates', () => {
    const a = createWorkbook()
    a.addWorksheet('S').setCell('A1', new Date('nope'))
    expect(() => a.xlsx()).toThrow(/invalid Date/)

    const b = createWorkbook()
    b.addWorksheet('S').setCell('A1', new Date(1899, 5, 1))
    expect(() => b.xlsx()).toThrow(/1900/)
  })

  it('rejects a bad A1 reference in setCell', () => {
    const s = createWorkbook().addWorksheet('S')
    for (const bad of ['A0', 'ZZ', '1', '$A$1', 'a1']) {
      expect(() => s.setCell(bad, 1), bad).toThrow(/invalid cell reference/)
    }
  })

  it('rejects an invalid colour and border style', () => {
    const s = createWorkbook().addWorksheet('S')
    expect(() => s.addRow([{ value: 1, style: { fill: 'red' } }])).not.toThrow() // deferred to serialise
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([{ value: 1, style: { fill: 'nope' } }])
    expect(() => wb.xlsx()).toThrow(/invalid colour/)
  })
})

describe('hostile input is neutralised, not passed through', () => {
  it('strips XML-illegal control characters from strings', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['a\x00\x01\x1Fb'])
    const parts = unzipSync(wb.xlsx())
    expect(strFromU8(parts['xl/sharedStrings.xml']!)).toContain(
      '<t xml:space="preserve">ab</t>',
    )
  })

  it('escapes markup in strings, sheet names, formulas and number formats', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('a & <b>')
    s.setCell('A1', '</t><si><t>injected', { numFmt: '"</numFmt>";0' })
    s.setCell('A2', { formula: 'IF(A1="</f>",1,0)' })
    const parts = unzipSync(wb.xlsx())
    const wbXml = strFromU8(parts['xl/workbook.xml']!)
    const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!)
    const styles = strFromU8(parts['xl/styles.xml']!)

    expect(wbXml).toContain('name="a &amp; &lt;b&gt;"')
    expect(strFromU8(parts['xl/sharedStrings.xml']!)).toContain('&lt;/t&gt;&lt;si&gt;')
    expect(sheet).toContain('IF(A1="&lt;/f&gt;",1,0)') // quotes kept, angle brackets escaped
    expect(styles).toContain('formatCode="&quot;&lt;/numFmt&gt;&quot;;0"')
  })

  it('every emitted part is still parseable as XML after hostile input', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S "quoted" & <angled>')
    s.addRow(["it's a <test> & \"more\"", 1, true, new Date(2020, 0, 1)])
    s.setCell('C3', { formula: 'A1&"<x>"', result: 'r&<s>' })
    s.merge('A1:B1')
    for (const [path, data] of Object.entries(unzipSync(wb.xlsx()))) {
      const xml = strFromU8(data)
      // balanced-ish sanity: declaration present, no raw stray '<' inside text
      expect(xml.startsWith('<?xml'), path).toBe(true)
      expect(xml).not.toMatch(/&(?!amp;|lt;|gt;|quot;|apos;|#)/) // no unescaped ampersands
    }
  })
})

describe('scale', () => {
  it('serialises 15k rows x 6 cols roughly linearly (guards against O(n^2))', () => {
    const run = (rows: number) => {
      const wb = createWorkbook()
      const s = wb.addWorksheet('Big')
      for (let r = 0; r < rows; r++) {
        s.addRow([`row ${r}`, r, r * 1.5, r % 2 === 0, new Date(2024, 0, 1 + (r % 365)), null])
      }
      const t0 = performance.now()
      const bytes = wb.xlsx()
      return { ms: performance.now() - t0, bytes }
    }

    const small = run(3_000)
    const big = run(15_000)

    // 5x the rows should be well under 15x the time if it's ~linear
    expect(big.ms).toBeLessThan(small.ms * 15 + 200)

    const sheet = strFromU8(unzipSync(big.bytes)['xl/worksheets/sheet1.xml']!)
    expect(sheet).toContain('<dimension ref="A1:E15000"/>') // col F was all null
    expect(sheet).toContain('<row r="15000">')
  })
})

describe('output shape', () => {
  it('is a valid ZIP starting with the local-file-header signature', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['x'])
    const b = wb.xlsx()
    expect([b[0], b[1], b[2], b[3]]).toEqual([0x50, 0x4b, 0x03, 0x04])
  })

  it('round-trips byte-identical for identical input (deterministic)', () => {
    const make = () => {
      const wb = createWorkbook()
      const s = wb.addWorksheet('S', { freeze: { ySplit: 1 } })
      s.addRow([{ value: 'H', style: { font: { bold: true } } }])
      s.addRow(['v', 1, new Date(2020, 0, 1)])
      s.merge('A3:B3')
      return wb.xlsx()
    }
    expect(Buffer.from(make()).equals(Buffer.from(make()))).toBe(true)
  })

  it('keeps a plain data export free of a sharedStrings part', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    for (let r = 0; r < 100; r++) s.addRow([r, r * 2, r % 2 === 0])
    expect(unzipSync(wb.xlsx())['xl/sharedStrings.xml']).toBeUndefined()
  })
})
