import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { createWorkbook } from '../workbook'
import { readWorkbook } from '../read'

/** Build a workbook, zip it, and return the parts as decoded strings. */
function build(fn: (wb: ReturnType<typeof createWorkbook>) => void): Record<string, string> {
  const wb = createWorkbook()
  fn(wb)
  const bytes = wb.xlsx()
  const entries = unzipSync(bytes)
  const out: Record<string, string> = {}
  for (const [path, data] of Object.entries(entries)) out[path] = strFromU8(data)
  return out
}

describe('createWorkbook — structure', () => {
  it('produces the mandatory OOXML parts', () => {
    const parts = build((wb) => {
      wb.addWorksheet('Sheet1').addRow(['Hello'])
    })
    expect(Object.keys(parts).sort()).toEqual(
      [
        '[Content_Types].xml',
        '_rels/.rels',
        'docProps/app.xml',
        'docProps/core.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/sharedStrings.xml',
        'xl/styles.xml',
        'xl/workbook.xml',
        'xl/worksheets/sheet1.xml',
      ].sort(),
    )
  })

  it('every part starts with the XML declaration', () => {
    const parts = build((wb) => wb.addWorksheet('S').addRow([1]))
    for (const [path, xml] of Object.entries(parts)) {
      expect(xml.startsWith('<?xml version="1.0"'), path).toBe(true)
    }
  })

  it('lists the worksheet in workbook.xml with a matching relationship', () => {
    const parts = build((wb) => wb.addWorksheet('Sales').addRow([1]))
    expect(parts['xl/workbook.xml']).toContain(
      '<sheet name="Sales" sheetId="1" r:id="rId1"/>',
    )
    expect(parts['xl/_rels/workbook.xml.rels']).toContain(
      'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"',
    )
  })

  it('declares content types for every part including sheets and sharedStrings', () => {
    const parts = build((wb) => wb.addWorksheet('S').addRow(['x']))
    const ct = parts['[Content_Types].xml']!
    expect(ct).toContain('PartName="/xl/worksheets/sheet1.xml"')
    expect(ct).toContain('PartName="/xl/sharedStrings.xml"')
    expect(ct).toContain('PartName="/xl/workbook.xml"')
  })
})

describe('createWorkbook — cells', () => {
  it('writes strings as shared, numbers bare, booleans as t="b"', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow(['a', 42, true, false, 3.5])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>')
    expect(sheet).toContain('<c r="B1"><v>42</v></c>')
    expect(sheet).toContain('<c r="C1" t="b"><v>1</v></c>')
    expect(sheet).toContain('<c r="D1" t="b"><v>0</v></c>')
    expect(sheet).toContain('<c r="E1"><v>3.5</v></c>')
  })

  it('skips null / undefined cells but keeps column position', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow(['a', null, 'c', undefined, 'e'])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>')
    expect(sheet).not.toContain('r="B1"')
    expect(sheet).toContain('<c r="C1" t="s"><v>1</v></c>')
    expect(sheet).not.toContain('r="D1"')
    expect(sheet).toContain('<c r="E1" t="s"><v>2</v></c>')
  })

  it('assigns increasing row numbers', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['r1']).addRow(['r2']).addRow(['r3'])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<row r="1">')
    expect(sheet).toContain('<row r="2">')
    expect(sheet).toContain('<row r="3">')
  })

  it('computes the dimension from the populated range', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['a', 'b', 'c'])
      s.addRow(['d'])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<dimension ref="A1:C2"/>')
  })

  it('escapes cell text', () => {
    const parts = build((wb) => wb.addWorksheet('S').addRow(['a & <b>']))
    expect(parts['xl/sharedStrings.xml']).toContain(
      '<t xml:space="preserve">a &amp; &lt;b&gt;</t>',
    )
  })

  it('rejects NaN / Infinity', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([NaN])
    expect(() => wb.xlsx()).toThrow(/finite/)
  })
})

describe('createWorkbook — shared strings', () => {
  it('de-duplicates repeated strings', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['dup', 'dup'])
      s.addRow(['dup', 'other'])
    })
    expect(parts['xl/sharedStrings.xml']).toContain('count="4" uniqueCount="2"')
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect((sheet.match(/<v>0<\/v>/g) ?? []).length).toBe(3) // three "dup" cells
    expect(sheet).toContain('<c r="B2" t="s"><v>1</v></c>')
  })

  it('omits sharedStrings.xml entirely when there are no string cells', () => {
    const parts = build((wb) => wb.addWorksheet('S').addRow([1, 2, 3]))
    expect(parts['xl/sharedStrings.xml']).toBeUndefined()
    expect(parts['[Content_Types].xml']).not.toContain('sharedStrings.xml')
    expect(parts['xl/_rels/workbook.xml.rels']).not.toContain('sharedStrings.xml')
  })

  it('shares the table across worksheets', () => {
    const parts = build((wb) => {
      wb.addWorksheet('One').addRow(['common'])
      wb.addWorksheet('Two').addRow(['common'])
    })
    expect(parts['xl/sharedStrings.xml']).toContain('count="2" uniqueCount="1"')
    expect(parts['xl/worksheets/sheet2.xml']).toContain('<v>0</v>')
  })
})

describe('createWorkbook — worksheets', () => {
  it('supports multiple sheets with their own parts and rels', () => {
    const parts = build((wb) => {
      wb.addWorksheet('A').addRow([1])
      wb.addWorksheet('B').addRow([2])
      wb.addWorksheet('C').addRow([3])
    })
    expect(parts['xl/worksheets/sheet3.xml']).toBeDefined()
    expect(parts['xl/workbook.xml']).toContain('sheetId="3" r:id="rId3"')
    expect(parts['xl/_rels/workbook.xml.rels']).toContain('Target="worksheets/sheet3.xml"')
    // styles relationship comes after the three sheets
    expect(parts['xl/_rels/workbook.xml.rels']).toContain('Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"')
  })

  it('rejects invalid, over-long, and duplicate names', () => {
    const wb = createWorkbook()
    wb.addWorksheet('Data')
    expect(() => wb.addWorksheet('')).toThrow(/1–31/)
    expect(() => wb.addWorksheet('x'.repeat(32))).toThrow(/1–31/)
    expect(() => wb.addWorksheet('a/b')).toThrow(/\\ \/ \? \* \[ \] :/)
    expect(() => wb.addWorksheet('DATA')).toThrow(/duplicate/i)
  })

  it('throws when serialising a workbook with no worksheets', () => {
    expect(() => createWorkbook().xlsx()).toThrow(/at least one worksheet/)
  })
})

describe('createWorkbook — output', () => {
  it('xlsx() returns a non-trivial Uint8Array with a ZIP signature', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['hi'])
    const bytes = wb.xlsx()
    expect(bytes).toBeInstanceOf(Uint8Array)
    expect(bytes.length).toBeGreaterThan(400)
    expect([bytes[0], bytes[1]]).toEqual([0x50, 0x4b]) // "PK"
  })

  it('is deterministic for identical input', () => {
    const make = () => {
      const wb = createWorkbook()
      wb.addWorksheet('S').addRow(['a', 1])
      return wb.xlsx()
    }
    expect(Array.from(make())).toEqual(Array.from(make()))
  })

  it('xlsx() is idempotent — repeated calls on one workbook match', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['dup', 'dup', 'x']).addRow([{ value: 1, style: { font: { bold: true } } }])

    const first = wb.xlsx()
    wb.blob() // also serialises internally
    const third = wb.xlsx()

    expect(Array.from(third)).toEqual(Array.from(first))
    const sst = strFromU8(unzipSync(third)['xl/sharedStrings.xml']!)
    expect(sst).toContain('count="3" uniqueCount="2"')
  })

  it('writes large / tiny magnitudes as plain decimals, round-trips them', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([1e21, 1e-7, 1.23e21])
    const sheet = strFromU8(unzipSync(wb.xlsx())['xl/worksheets/sheet1.xml']!)
    expect(sheet).toContain('<v>1000000000000000000000</v>')
    expect(sheet).toContain('<v>0.0000001</v>')
    expect(sheet).not.toMatch(/<v>[^<]*e[+-]/i)

    const back = readWorkbook(wb.xlsx()).sheet('S')!.values()[0]
    expect(back).toEqual([1e21, 1e-7, 1.23e21])
  })

  it('blob() carries the spreadsheet mime type', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([1])
    const blob = wb.blob()
    expect(blob.type).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    expect(blob.size).toBeGreaterThan(400)
  })
})
