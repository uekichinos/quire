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

describe('createWorkbook — hyperlinks', () => {
  it('writes a <hyperlinks> block plus a worksheet rels part', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow([{ hyperlink: 'https://example.com', text: 'Example', tooltip: 'go there' }])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<hyperlinks><hyperlink r:id="rId1" ref="A1" tooltip="go there"/></hyperlinks>',
    )
    expect(parts['xl/worksheets/sheet1.xml']).toContain('xmlns:r=')
    expect(parts['xl/worksheets/_rels/sheet1.xml.rels']).toContain(
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>',
    )
    // the cell text itself is a normal shared string
    expect(parts['xl/sharedStrings.xml']).toContain('Example')
  })

  it('defaults the displayed text to the target URL when omitted', () => {
    const parts = build((wb) => wb.addWorksheet('S').setCell('A1', { hyperlink: 'mailto:a@b.com' }))
    expect(parts['xl/sharedStrings.xml']).toContain('mailto:a@b.com')
  })

  it('omits the rels part and <hyperlinks> block when there are none', () => {
    const parts = build((wb) => wb.addWorksheet('S').addRow(['plain']))
    expect(parts['xl/worksheets/_rels/sheet1.xml.rels']).toBeUndefined()
    expect(parts['xl/worksheets/sheet1.xml']).not.toContain('<hyperlinks>')
  })

  it('overwriting a cell drops its hyperlink', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setCell('A1', { hyperlink: 'https://example.com' })
      s.setCell('A1', 'plain again')
    })
    expect(parts['xl/worksheets/sheet1.xml']).not.toContain('<hyperlinks>')
    expect(parts['xl/worksheets/_rels/sheet1.xml.rels']).toBeUndefined()
  })

  it('rejects an empty or over-long hyperlink target', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    expect(() => s.setCell('A1', { hyperlink: '' })).toThrow(/non-empty/)
    expect(() => s.setCell('A1', { hyperlink: 'https://x.com/' + 'a'.repeat(2100) })).toThrow(
      /exceeds/,
    )
  })

  it('round-trips through readWorkbook, including with a per-cell style', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow([{ hyperlink: 'https://example.com', text: 'Example', tooltip: 'hi' }])
    s.setCell('B1', { hyperlink: 'https://y.com', text: 'Y' }, { font: { bold: true } })
    const back = readWorkbook(wb.xlsx())
    const sheet = back.sheet('S')!
    expect(sheet.cell('A1')).toMatchObject({
      value: 'Example',
      hyperlink: { target: 'https://example.com', tooltip: 'hi' },
    })
    expect(sheet.cell('B1')).toMatchObject({
      value: 'Y',
      hyperlink: { target: 'https://y.com' },
    })
  })

  it('is idempotent across repeated xlsx() calls', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([{ hyperlink: 'https://example.com', text: 'Example' }])
    const first = wb.xlsx()
    const second = wb.xlsx()
    expect(Array.from(second)).toEqual(Array.from(first))
  })
})

describe('createWorkbook — defined names', () => {
  it('writes a workbook-scoped named range', () => {
    const parts = build((wb) => {
      wb.addWorksheet('Sales').addRow(['h'])
      wb.defineName('SalesRange', 'Sales', 'A1:B10')
    })
    expect(parts['xl/workbook.xml']).toContain(
      `<definedName name="SalesRange">'Sales'!$A$1:$B$10</definedName>`,
    )
  })

  it('accepts a single-cell target', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow(['h'])
      wb.defineName('Anchor', 'S', 'A1')
    })
    expect(parts['xl/workbook.xml']).toContain(
      `<definedName name="Anchor">'S'!$A$1</definedName>`,
    )
  })

  it('coexists with the internal autoFilter defined name', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['h']).autoFilter('A1:A1')
      wb.defineName('MyRange', 'S', 'A1')
    })
    expect(parts['xl/workbook.xml']).toContain('_xlnm._FilterDatabase')
    expect(parts['xl/workbook.xml']).toContain('name="MyRange"')
  })

  it('rejects invalid, duplicate, reserved and unresolvable names', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['h'])
    expect(() => wb.defineName('A1', 'S', 'A1')).toThrow(/cell reference/)
    expect(() => wb.defineName('has space', 'S', 'A1')).toThrow(/must start with/)
    expect(() => wb.defineName('_xlnm.Foo', 'S', 'A1')).toThrow(/reserved/)
    expect(() => wb.defineName('Good', 'NoSuchSheet', 'A1')).toThrow(/no worksheet/)
    expect(() => wb.defineName('Good', 'S', 'not a ref')).toThrow(/invalid cell reference/)
    wb.defineName('Dup', 'S', 'A1')
    expect(() => wb.defineName('Dup', 'S', 'B1')).toThrow(/duplicate/)
    expect(() => wb.defineName('dup', 'S', 'B1')).toThrow(/duplicate/) // case-insensitive
  })

  it('round-trips through readWorkbook', () => {
    const wb = createWorkbook()
    wb.addWorksheet('Sales').addRow(['h'])
    wb.defineName('SalesRange', 'Sales', 'A1:B10')
    const back = readWorkbook(wb.xlsx())
    expect(back.definedNames).toEqual([
      { name: 'SalesRange', sheetName: 'Sales', range: 'A1:B10', refersTo: "'Sales'!$A$1:$B$10", hidden: undefined },
    ])
  })
})

describe('createWorkbook — row default style (empty cells)', () => {
  it('emits s + customFormat on a row with a style but no populated cells', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setRow(2, { style: { fill: 'FF00FF00' } })
      s.addRow(['a']).addRow(['b']).addRow(['c'])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toMatch(/<row r="2" s="\d+" customFormat="1">/)
  })

  it('round-trips into ReadWorksheet.rowStyles under { styles: true }', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.setRow(2, { style: { font: { bold: true } } })
    s.addRow(['a']).addRow(['b']).addRow(['c'])
    const sheet = readWorkbook(wb.xlsx(), { styles: true }).sheet('S')!
    expect(sheet.rowStyles.get(2)).toMatchObject({ font: { bold: true } })
  })
})

describe('createWorkbook — data validation', () => {
  it('writes an inline list as a quoted comma-separated formula1', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setDataValidation('A2:A100', { type: 'list', list: ['Open', 'In Progress', 'Done'] })
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2:A100">' +
        '<formula1>"Open,In Progress,Done"</formula1></dataValidation>',
    )
  })

  it('writes a range reference unquoted, and prompt/error attrs when given', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setDataValidation(
        'B1:B10',
        { type: 'list', list: 'Lookup!$A$1:$A$5' },
        {
          allowBlank: false,
          promptTitle: 'Pick',
          promptMessage: 'Choose one',
          errorTitle: 'Bad',
          errorMessage: 'Not allowed',
        },
      )
    })
    const xml = parts['xl/worksheets/sheet1.xml']!
    expect(xml).toContain('allowBlank="0"')
    expect(xml).toContain('promptTitle="Pick"')
    expect(xml).toContain('prompt="Choose one"')
    expect(xml).toContain('errorTitle="Bad"')
    expect(xml).toContain('error="Not allowed"')
    expect(xml).toContain('<formula1>Lookup!$A$1:$A$5</formula1>')
  })

  it('rejects a comma in a choice, an empty list, and an over-long inline list', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    expect(() => s.setDataValidation('A1', { type: 'list', list: ['a,b'] })).toThrow(/comma/)
    expect(() => s.setDataValidation('A1', { type: 'list', list: [] })).toThrow(/at least one choice/)
    expect(() => s.setDataValidation('A1', { type: 'list', list: ['x'.repeat(260)] })).toThrow(/255-character/)
    expect(() => s.setDataValidation('A1', { type: 'list', list: '' })).toThrow(/needs a list/)
  })

  it('writes a whole-number "between" rule with two formulas', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setDataValidation('A1:A10', { type: 'whole', operator: 'between', value: [1, 100] })
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<dataValidation type="whole" allowBlank="1" operator="between" showInputMessage="1" showErrorMessage="1" sqref="A1:A10">' +
        '<formula1>1</formula1><formula2>100</formula2></dataValidation>',
    )
  })

  it('writes a date rule as Excel serials', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setDataValidation('A1', { type: 'date', operator: 'greaterThan', value: new Date(2024, 0, 1) })
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<dataValidation type="date" allowBlank="1" operator="greaterThan" showInputMessage="1" showErrorMessage="1" sqref="A1">' +
        '<formula1>45292</formula1></dataValidation>',
    )
  })

  it('rejects a pair value for a non-between operator and a single value for between', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    expect(() =>
      s.setDataValidation('A1', { type: 'whole', operator: 'greaterThan', value: [1, 2] }),
    ).toThrow(/takes a single value/)
    expect(() =>
      s.setDataValidation('A1', { type: 'whole', operator: 'between', value: 1 as never }),
    ).toThrow(/needs a \[min, max\]/)
  })

  it('round-trips a list rule through readWorkbook', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    s.setDataValidation('A2:A100', { type: 'list', list: ['Open', 'Done'] }, { promptMessage: 'Pick one' })
    const sheet = readWorkbook(wb.xlsx()).sheet('S')!
    expect(sheet.dataValidations).toEqual([
      {
        ref: 'A2:A100',
        type: 'list',
        allowBlank: true,
        promptTitle: undefined,
        promptMessage: 'Pick one',
        errorTitle: undefined,
        errorMessage: undefined,
        list: ['Open', 'Done'],
      },
    ])
  })

  it('round-trips a decimal "between" rule with resolved numeric values', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    s.setDataValidation('B1:B10', { type: 'decimal', operator: 'between', value: [0.5, 9.5] })
    const rule = readWorkbook(wb.xlsx()).sheet('S')!.dataValidations[0]!
    expect(rule).toMatchObject({ type: 'decimal', operator: 'between', values: [0.5, 9.5] })
  })

  it('round-trips a date rule back into Date objects', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    s.setDataValidation('C1', { type: 'date', operator: 'greaterThan', value: new Date(2024, 0, 1) })
    const rule = readWorkbook(wb.xlsx()).sheet('S')!.dataValidations[0]!
    expect(rule.type).toBe('date')
    expect((rule as { values: Date[] }).values[0]).toEqual(new Date(2024, 0, 1))
  })
})

describe('createWorkbook — conditional formatting', () => {
  it('writes a cellIs rule with a dxf and round-trips it', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow([50])
    s.addConditionalFormat('A1:A100', {
      type: 'cellIs',
      operator: 'greaterThan',
      formula: 100,
      style: { font: { bold: true, color: 'FF0000' }, fill: 'FFFF00' },
    })
    const bytes = wb.xlsx()
    const parts = unzipSync(bytes)
    const sheetXml = strFromU8(parts['xl/worksheets/sheet1.xml']!)
    expect(sheetXml).toContain('<conditionalFormatting sqref="A1:A100">')
    expect(sheetXml).toMatch(/<cfRule type="cellIs" dxfId="\d+" priority="1" operator="greaterThan"><formula>100<\/formula><\/cfRule>/)
    expect(strFromU8(parts['xl/styles.xml']!)).toContain('<dxfs count="1">')

    const sheet = readWorkbook(bytes, { styles: true }).sheet('S')!
    expect(sheet.conditionalFormats).toEqual([
      {
        ref: 'A1:A100',
        rule: {
          type: 'cellIs',
          operator: 'greaterThan',
          formula: ['100'],
          style: { font: { bold: true, color: 'FFFF0000' }, fill: 'FFFFFF00' },
        },
      },
    ])
  })

  it('writes a 3-stop colorScale rule and round-trips it', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow([1])
    s.addConditionalFormat('A1:A100', { type: 'colorScale', colors: ['FF0000', 'FFFF00', '00FF00'] })
    const bytes = wb.xlsx()
    const sheet = readWorkbook(bytes).sheet('S')!
    expect(sheet.conditionalFormats).toEqual([
      {
        ref: 'A1:A100',
        rule: { type: 'colorScale', colors: ['FFFF0000', 'FFFFFF00', 'FF00FF00'] },
      },
    ])
  })

  it('rejects a colorScale with the wrong number of colours and a between rule missing a pair', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    expect(() =>
      s.addConditionalFormat('A1', { type: 'colorScale', colors: ['FF0000'] as never }),
    ).toThrow(/2 or 3 colours/)
    expect(() =>
      s.addConditionalFormat('A1', { type: 'cellIs', operator: 'between', formula: 5, style: {} }),
    ).toThrow(/needs a \[min, max\] formula/)
  })

  it('writes a dataBar rule and round-trips it', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow([1])
    s.addConditionalFormat('A1:A100', { type: 'dataBar', color: '638EC6' })
    const bytes = wb.xlsx()
    const parts = unzipSync(bytes)
    expect(strFromU8(parts['xl/worksheets/sheet1.xml']!)).toContain(
      '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/>' +
        '<color rgb="FF638EC6"/></dataBar></cfRule>',
    )
    const sheet = readWorkbook(bytes).sheet('S')!
    expect(sheet.conditionalFormats).toEqual([
      { ref: 'A1:A100', rule: { type: 'dataBar', color: 'FF638EC6' } },
    ])
  })
})

describe('createWorkbook — rich text', () => {
  it('writes runs with per-run <rPr> into a shared <si>', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow([
        [{ text: 'Order ' }, { text: '#1234', font: { bold: true, color: 'FF0000' } }],
      ])
    })
    expect(parts['xl/sharedStrings.xml']).toContain(
      '<si><r><t xml:space="preserve">Order </t></r>' +
        '<r><rPr><b/><color rgb="FFFF0000"/></rPr><t xml:space="preserve">#1234</t></r></si>',
    )
  })

  it('de-duplicates identical rich text values', () => {
    const rich = [{ text: 'a' }, { text: 'b', font: { italic: true } }]
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow([rich])
      s.addRow([[...rich]]) // structurally identical, different array instance
    })
    expect(parts['xl/sharedStrings.xml']).toContain('uniqueCount="1"')
  })

  it('rejects an empty run array', () => {
    const wb = createWorkbook()
    expect(() => wb.addWorksheet('S').setCell('A1', [])).toThrow(/at least one run/)
  })

  it('round-trips through readWorkbook', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([
      [{ text: 'Order ' }, { text: '#1234', font: { bold: true, color: 'FF0000' } }, { text: ' shipped' }],
    ])
    const cell = readWorkbook(wb.xlsx()).sheet('S')!.cell('A1')!
    expect(cell.value).toEqual([
      { text: 'Order ' },
      { text: '#1234', font: { bold: true, color: 'FFFF0000' } },
      { text: ' shipped' },
    ])
  })

  it('a plain string cell is unaffected — still round-trips as a bare string', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['plain'])
    const cell = readWorkbook(wb.xlsx()).sheet('S')!.cell('A1')!
    expect(cell.value).toBe('plain')
  })
})

describe('createWorkbook — comments', () => {
  it('writes commentsN.xml, vmlDrawingN.vml, the sheet rels, and content types', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setComment('A1', 'Hello', { author: 'Alice' })
    })
    expect(parts['xl/comments1.xml']).toContain(
      '<author>Alice</author>',
    )
    expect(parts['xl/comments1.xml']).toContain(
      '<comment ref="A1" authorId="0"><text><r><t xml:space="preserve">Hello</t></r></text></comment>',
    )
    expect(parts['xl/drawings/vmlDrawing1.vml']).toContain('<x:Row>0</x:Row><x:Column>0</x:Column>')
    expect(parts['xl/worksheets/sheet1.xml']).toMatch(/<legacyDrawing r:id="rId\d+"\/>/)
    expect(parts['xl/worksheets/_rels/sheet1.xml.rels']).toContain('relationships/comments')
    expect(parts['xl/worksheets/_rels/sheet1.xml.rels']).toContain('relationships/vmlDrawing')
    expect(parts['[Content_Types].xml']).toContain('/xl/comments1.xml')
    expect(parts['[Content_Types].xml']).toContain('Extension="vml"')
  })

  it('omits comment parts entirely when there are none', () => {
    const parts = build((wb) => wb.addWorksheet('S').addRow(['x']))
    expect(parts['xl/comments1.xml']).toBeUndefined()
    expect(parts['xl/drawings/vmlDrawing1.vml']).toBeUndefined()
    expect(parts['xl/worksheets/_rels/sheet1.xml.rels']).toBeUndefined()
  })

  it('keys authors per distinct name', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setComment('A1', 'first', { author: 'Alice' })
      s.setComment('B1', 'second', { author: 'Bob' })
      s.setComment('C1', 'third', { author: 'Alice' })
    })
    const xml = parts['xl/comments1.xml']!
    expect(xml).toContain('<authors><author>Alice</author><author>Bob</author></authors>')
    expect(xml).toMatch(/<comment ref="A1" authorId="0">/)
    expect(xml).toMatch(/<comment ref="B1" authorId="1">/)
    expect(xml).toMatch(/<comment ref="C1" authorId="0">/)
  })

  it('rejects empty comment text', () => {
    const wb = createWorkbook()
    expect(() => wb.addWorksheet('S').setComment('A1', '')).toThrow(/non-empty/)
  })

  it('round-trips, including a comment on a cell with no value at all', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    s.setComment('A1', 'on a value', { author: 'A' })
    s.setComment('Z9', 'on nothing')
    const sheet = readWorkbook(wb.xlsx()).sheet('S')!
    expect(sheet.cell('A1')!.comment).toEqual({ text: 'on a value', author: 'A' })
    expect(sheet.cell('Z9')).toMatchObject({ type: 'empty', value: null, comment: { text: 'on nothing', author: '' } })
  })

  it('is idempotent across repeated xlsx() calls', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').setComment('A1', 'hi')
    expect(Array.from(wb.xlsx())).toEqual(Array.from(wb.xlsx()))
  })
})

describe('createWorkbook — print setup', () => {
  it('writes sheetPr fitToPage only when fitToWidth/fitToHeight are set', () => {
    const withFit = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setPageSetup({ fitToWidth: 1, fitToHeight: 0 })
    })
    expect(withFit['xl/worksheets/sheet1.xml']).toContain(
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
    )

    const noFit = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setPageSetup({ orientation: 'landscape' })
    })
    expect(noFit['xl/worksheets/sheet1.xml']).not.toContain('sheetPr')
  })

  it('defaults margins and merges partial overrides over Excel\'s own defaults', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x'])
      s.setPageSetup({ margins: { left: 0.2, header: 0.1 } })
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<pageMargins left="0.2" right="0.7" top="0.75" bottom="0.75" header="0.1" footer="0.3"/>',
    )
  })

  it('writes the print area as a workbook-level _xlnm.Print_Area defined name', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('Sales')
      s.addRow(['x'])
      s.setPageSetup({ printArea: 'A1:C10' })
    })
    expect(parts['xl/workbook.xml']).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="0">'Sales'!$A$1:$C$10</definedName>`,
    )
  })

  it('rejects invalid fitToWidth/fitToHeight/scale/paperSize/printArea', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    expect(() => s.setPageSetup({ fitToWidth: -1 })).toThrow(/fitToWidth/)
    expect(() => s.setPageSetup({ fitToHeight: 1.5 })).toThrow(/fitToHeight/)
    expect(() => s.setPageSetup({ scale: 0 })).toThrow(/scale/)
    expect(() => s.setPageSetup({ paperSize: 0 })).toThrow(/paperSize/)
    expect(() => s.setPageSetup({ printArea: 'not a ref' })).toThrow(/invalid/)
  })

  it('round-trips through readWorkbook', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    s.setPageSetup({ orientation: 'landscape', paperSize: 9, printArea: 'A1:B5' })
    const sheet = readWorkbook(wb.xlsx()).sheet('S')!
    expect(sheet.pageSetup).toMatchObject({ orientation: 'landscape', paperSize: 9, printArea: 'A1:B5' })
  })

  it('a sheet without setPageSetup has no pageSetup on read', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['x'])
    expect(readWorkbook(wb.xlsx()).sheet('S')!.pageSetup).toBeUndefined()
  })
})
