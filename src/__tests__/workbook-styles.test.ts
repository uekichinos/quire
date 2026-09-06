import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { createWorkbook } from '../workbook'

function build(fn: (wb: ReturnType<typeof createWorkbook>) => void): Record<string, string> {
  const wb = createWorkbook()
  fn(wb)
  const entries = unzipSync(wb.xlsx())
  const out: Record<string, string> = {}
  for (const [path, data] of Object.entries(entries)) out[path] = strFromU8(data)
  return out
}

describe('styling', () => {
  it('applies a cell style and references it via s=', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow([{ value: 'Header', style: { font: { bold: true } } }])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toMatch(/<c r="A1" s="1" t="s">/)
    expect(parts['xl/styles.xml']).toContain('<b/>')
  })

  it('applies a row default style to bare cells', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow(['a', 'b'], { style: { fill: 'FFFF00' } })
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toMatch(/<c r="A1" s="1"/)
    expect(sheet).toMatch(/<c r="B1" s="1"/)
    expect(parts['xl/styles.xml']).toContain('<fgColor rgb="FFFFFF00"/>')
  })

  it('lets a cell style override the row style, merged one level deep', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow(
        [{ value: 'x', style: { font: { italic: true } } }],
        { style: { font: { bold: true }, fill: 'EEEEEE' } },
      )
    })
    const styles = parts['xl/styles.xml']!
    // the effective font has BOTH bold (from row) and italic (from cell)
    expect(styles).toMatch(/<font><b\/><i\/>/)
    expect(styles).toContain('<fgColor rgb="FFEEEEEE"/>')
  })

  it('de-duplicates repeated styles into one cellXfs entry', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow([
        { value: 1, style: { numFmt: '#,##0.00' } },
        { value: 2, style: { numFmt: '#,##0.00' } },
      ])
    })
    // default xf + one shared "#,##0.00" xf
    expect(parts['xl/styles.xml']).toContain('<cellXfs count="2">')
  })

  it('shares the style pool across worksheets', () => {
    const parts = build((wb) => {
      wb.addWorksheet('A').addRow([{ value: 1, style: { font: { bold: true } } }])
      wb.addWorksheet('B').addRow([{ value: 2, style: { font: { bold: true } } }])
    })
    expect(parts['xl/styles.xml']).toContain('<cellXfs count="2">')
    expect(parts['xl/worksheets/sheet1.xml']).toContain('s="1"')
    expect(parts['xl/worksheets/sheet2.xml']).toContain('s="1"')
  })
})

describe('Date cells', () => {
  it('writes a serial number and a date number-format', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow([new Date(2020, 0, 1)])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toMatch(/<c r="A1" s="1"><v>43831<\/v><\/c>/)
    expect(parts['xl/styles.xml']).toContain('formatCode="yyyy-mm-dd"')
  })

  it('respects an explicit numFmt on a Date cell', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow([
        { value: new Date(2020, 0, 1, 9, 30), style: { numFmt: 'yyyy-mm-dd hh:mm' } },
      ])
    })
    expect(parts['xl/styles.xml']).toContain('formatCode="yyyy-mm-dd hh:mm"')
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<v>43831.395833333336</v>')
  })
})

describe('formula cells', () => {
  it('writes <f> with a cached numeric result', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow([1]).addRow([2]).addRow([{ value: { formula: 'SUM(A1:A2)', result: 3 } }])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<c r="A3"><f>SUM(A1:A2)</f><v>3</v></c>',
    )
  })

  it('strips a leading = and omits <v> when there is no result', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').setCell('B1', { formula: '=A1*2' })
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<c r="B1"><f>A1*2</f></c>')
  })

  it('tags a string result with t="str" and a boolean with t="b"', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setCell('A1', { formula: 'CONCAT("a","b")', result: 'ab' })
      s.setCell('A2', { formula: 'A1="ab"', result: true })
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<c r="A1" t="str"><f>CONCAT("a","b")</f><v>ab</v></c>')
    expect(sheet).toContain('<c r="A2" t="b"><f>A1="ab"</f><v>1</v></c>')
  })

  it('carries a style on a formula cell', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').setCell(
        'A1',
        { formula: 'SUM(B:B)', result: 10 },
        { font: { bold: true } },
      )
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<c r="A1" s="1"><f>SUM(B:B)</f>')
  })
})

describe('setCell', () => {
  it('places a value at an arbitrary A1 reference and extends the dimension', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').setCell('C5', 'here')
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<row r="5">')
    expect(sheet).toContain('<c r="C5" t="s">')
    expect(sheet).toContain('<dimension ref="A1:C5"/>')
  })

  it('interleaves with addRow without colliding', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['r1'])
      s.setCell('B3', 'r3b')
      s.addRow(['r4']) // next auto row is 4, after the setCell at row 3
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<row r="1">')
    expect(sheet).toContain('<row r="3">')
    expect(sheet).toContain('<row r="4">')
  })
})
