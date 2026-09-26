import { zipSync, strToU8 } from 'fflate'
import { describe, expect, it } from 'vitest'
import { createWorkbook } from '../workbook'
import { readWorkbook, readWorkbookAsync } from '../read'
import { XlsxReadError } from '../unzip'

/** Build an .xlsx with the writer, read it back. */
function roundTrip(fn: (wb: ReturnType<typeof createWorkbook>) => void) {
  const wb = createWorkbook()
  fn(wb)
  return readWorkbook(wb.xlsx())
}

/** Hand-assemble a minimal .xlsx from raw part strings, for edge cases the writer won't emit. */
function makeXlsx(parts: Record<string, string>): Uint8Array {
  const z: Record<string, Uint8Array> = {}
  for (const [k, v] of Object.entries(parts)) z[k] = strToU8(v)
  return zipSync(z)
}

const CT = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="x"/><Default Extension="xml" ContentType="x"/></Types>`
const RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="xl/workbook.xml"/></Relationships>`

describe('readWorkbook — round-trips the writer', () => {
  it('values, types and multiple sheets', () => {
    const wb = roundTrip((w) => {
      w.addWorksheet('One').addRow(['a', 1, true, false]).addRow([2.5, '', 'x'])
      w.addWorksheet('Two').addRow(['only'])
    })
    expect(wb.sheetNames).toEqual(['One', 'Two'])
    const one = wb.sheet('One')!
    // values() is rectangular to maxCol; an empty-string cell round-trips as ''
    expect(one.values()).toEqual([
      ['a', 1, true, false],
      [2.5, '', 'x', null],
    ])
    expect(one.cell('B1')).toMatchObject({ type: 'number', value: 1, ref: 'B1', row: 1, col: 2 })
    expect(one.cell('C1')).toMatchObject({ type: 'boolean', value: true })
    expect(wb.sheet(1)!.values()).toEqual([['only']])
  })

  it('de-duplicated strings resolve correctly', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow(['dup', 'dup'])
      s.addRow(['dup', 'other'])
    })
    expect(wb.sheet('S')!.values()).toEqual([
      ['dup', 'dup'],
      ['dup', 'other'],
    ])
  })

  it('formula cells expose formula + cached value', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow([1]).addRow([2])
      s.setCell('A3', { formula: 'SUM(A1:A2)', result: 3 })
      s.setCell('B1', { formula: '=A1&"x"', result: '1x' })
    })
    const s = wb.sheet('S')!
    expect(s.cell('A3')).toMatchObject({ type: 'formula', formula: 'SUM(A1:A2)', value: 3 })
    expect(s.cell('B1')).toMatchObject({ type: 'formula', formula: 'A1&"x"', value: '1x' })
  })

  it('merged ranges are reported', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.setCell('A1', 'title')
      s.merge('A1:C1')
    })
    expect(wb.sheet('S')!.merges).toEqual(['A1:C1'])
  })

  it('dimension comes from the file', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow(['a', 'b', 'c'])
      s.addRow(['d'])
    })
    expect(wb.sheet('S')!.dimension).toEqual({ rows: 2, cols: 3 })
  })

  it('escaped and unicode text survives the trip', () => {
    const wb = roundTrip((w) => {
      w.addWorksheet('S').addRow(['a & <b> "c"', 'café 你好 📊'])
    })
    expect(wb.sheet('S')!.values()).toEqual([['a & <b> "c"', 'café 你好 📊']])
  })

  it('sheets option loads a subset', () => {
    const wb = createWorkbook()
    wb.addWorksheet('A').addRow([1])
    wb.addWorksheet('B').addRow([2])
    wb.addWorksheet('C').addRow([3])
    const back = readWorkbook(wb.xlsx(), { sheets: ['A', 2] })
    expect(back.sheetNames).toEqual(['A', 'C'])
  })

  it('reads Date cells back as Date objects', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow([new Date(2024, 2, 1), 'not a date', 42])
      s.addRow([{ value: new Date(2024, 2, 1, 9, 30), style: { numFmt: 'yyyy-mm-dd hh:mm' } }])
    })
    const s = wb.sheet('S')!
    const a1 = s.cell('A1')!
    expect(a1.type).toBe('date')
    expect(a1.value).toBeInstanceOf(Date)
    expect((a1.value as Date).getTime()).toBe(new Date(2024, 2, 1).getTime())
    expect(s.cell('B1')!.type).toBe('string')
    expect(s.cell('C1')!.type).toBe('number')
    expect((s.cell('A2')!.value as Date).getTime()).toBe(new Date(2024, 2, 1, 9, 30).getTime())
  })

  it('{ dates: false } keeps date cells as numbers', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([new Date(2024, 2, 1)])
    const back = readWorkbook(wb.xlsx(), { dates: false })
    const c = back.sheet('S')!.cell('A1')!
    expect(c.type).toBe('number')
    expect(c.value).toBe(45352)
  })

  it('accepts an ArrayBuffer', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['x'])
    const buf = wb.xlsx().buffer
    expect(readWorkbook(buf as ArrayBuffer).sheet('S')!.cell('A1')!.value).toBe('x')
  })

  it('surfaces numFmt on every cell (not just under { styles: true })', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow([{ value: 1.5, style: { numFmt: '#,##0.00' } }, 42])
      s.addRow([{ value: 0.25, style: { numFmt: '0.00%' } }]) // built-in id 10
    })
    const s = wb.sheet('S')!
    expect(s.cell('A1')!.numFmt).toBe('#,##0.00')
    expect(s.cell('A2')!.numFmt).toBe('0.00%')
    expect(s.cell('B1')!.numFmt).toBeUndefined() // General
  })

  it('values({ ragged: true }) trims each row to its own width', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow(['a', 'b', 'c'])
      s.addRow(['d'])
    })
    const s = wb.sheet('S')!
    expect(s.values()).toEqual([
      ['a', 'b', 'c'],
      ['d', null, null],
    ])
    expect(s.values({ ragged: true })).toEqual([['a', 'b', 'c'], ['d']])
  })

  it('readWorkbookAsync returns the same model', async () => {
    const wb = createWorkbook()
    wb.addWorksheet('A').addRow([1, 2])
    wb.addWorksheet('B').addRow(['x'])
    const back = await readWorkbookAsync(wb.xlsx())
    expect(back.sheetNames).toEqual(['A', 'B'])
    expect(back.sheet('A')!.values()).toEqual([[1, 2]])
  })

  it('enforces maxCells', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    for (let i = 0; i < 50; i++) s.addRow([i, i, i, i])
    expect(() => readWorkbook(wb.xlsx(), { limits: { maxCells: 100 } })).toThrow(/cells/)
  })

  it('enforces maxSheets', () => {
    const wb = createWorkbook()
    for (let i = 0; i < 5; i++) wb.addWorksheet(`S${i}`).addRow([i])
    expect(() => readWorkbook(wb.xlsx(), { limits: { maxSheets: 3 } })).toThrow(/sheets/)
  })
})

describe('readWorkbook — files the writer never emits', () => {
  it('inline strings (t="inlineStr")', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>hi there</t></is></c></row></sheetData></worksheet>`,
    })
    expect(readWorkbook(bytes).sheet('S')!.cell('A1')).toMatchObject({
      type: 'string',
      value: 'hi there',
    })
  })

  it('cells without an r attribute fall back to position', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row><c><v>10</v></c><c><v>20</v></c></row><row><c><v>30</v></c></row></sheetData></worksheet>`,
    })
    expect(readWorkbook(bytes).sheet('S')!.values()).toEqual([
      [10, 20],
      [30, null],
    ])
  })

  it('error cells (t="e")', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row></sheetData></worksheet>`,
    })
    expect(readWorkbook(bytes).sheet('S')!.cell('A1')).toMatchObject({
      type: 'error',
      value: '#DIV/0!',
    })
  })
})

describe('readWorkbook — malformed / hostile input', () => {
  it('rejects a non-zip', () => {
    expect(() => readWorkbook(strToU8('not a zip'))).toThrow(XlsxReadError)
  })

  it('rejects a zip with no workbook.xml', () => {
    expect(() => readWorkbook(makeXlsx({ 'foo.txt': 'bar' }))).toThrow(/workbook\.xml is missing/)
  })

  it('rejects a DOCTYPE inside a part (billion-laughs guard)', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;">]><workbook><sheets/></workbook>`,
    })
    expect(() => readWorkbook(bytes)).toThrow(/DOCTYPE/)
  })

  it('ignores archive entries outside the allow-list and with traversal', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      '../../evil.sh': 'rm -rf /',
      'xl/media/image1.png': 'not really a png',
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
    })
    const wb = readWorkbook(bytes)
    expect(wb.sheet('S')!.cell('A1')!.value).toBe(1)
  })

  it('enforces the total-size limit', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('Big')
    for (let i = 0; i < 2000; i++) s.addRow([`row ${i}`, i])
    expect(() => readWorkbook(wb.xlsx(), { limits: { maxTotalBytes: 1000 } })).toThrow(
      /limit/,
    )
  })

  it('does not pollute Object.prototype from a hostile sheet name', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="__proto__" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/></worksheet>`,
    })
    const wb = readWorkbook(bytes)
    expect(wb.sheetNames).toEqual(['__proto__'])
    expect(({} as Record<string, unknown>).polluted).toBeUndefined()
  })
})

describe('readWorkbook — hyperlinks', () => {
  it('round-trips a hyperlink written by the writer, tooltip included', () => {
    const wb = roundTrip((w) => {
      w.addWorksheet('S').addRow([
        { hyperlink: 'https://example.com', text: 'Example', tooltip: 'hi' },
      ])
    })
    expect(wb.sheet('S')!.cell('A1')).toMatchObject({
      value: 'Example',
      hyperlink: { target: 'https://example.com', tooltip: 'hi' },
    })
  })

  it('resolves an internal link via the location attribute (no relationship)', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/></hyperlinks></worksheet>`,
      'xl/sharedStrings.xml': `<?xml version="1.0"?><sst><si><t>jump</t></si></sst>`,
    })
    const cell = readWorkbook(bytes).sheet('S')!.cell('A1')!
    expect(cell.hyperlink).toEqual({ target: 'Sheet2!A1', tooltip: undefined })
  })

  it('ignores a dangling relationship id without throwing', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" r:id="rIdMissing"/></hyperlinks></worksheet>`,
      'xl/sharedStrings.xml': `<?xml version="1.0"?><sst><si><t>x</t></si></sst>`,
    })
    const cell = readWorkbook(bytes).sheet('S')!.cell('A1')!
    expect(cell.hyperlink).toBeUndefined()
  })

  it('attaches a range hyperlink to its top-left cell only', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData><hyperlinks><hyperlink ref="A1:B1" r:id="rId1"/></hyperlinks></worksheet>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="https://example.com" TargetMode="External"/></Relationships>`,
      'xl/sharedStrings.xml': `<?xml version="1.0"?><sst><si><t>a</t></si><si><t>b</t></si></sst>`,
    })
    const sheet = readWorkbook(bytes).sheet('S')!
    expect(sheet.cell('A1')!.hyperlink).toEqual({ target: 'https://example.com', tooltip: undefined })
    expect(sheet.cell('B1')!.hyperlink).toBeUndefined()
  })
})

describe('readWorkbook — defined names', () => {
  it('excludes the internal _xlnm._FilterDatabase name', () => {
    const wb = roundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow(['h']).autoFilter('A1:A1')
    })
    expect(wb.definedNames).toEqual([])
  })

  it('resolves a quoted sheet name with an escaped apostrophe', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="Bob's Sheet" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Foo">'Bob''s Sheet'!$A$1</definedName></definedNames></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/></worksheet>`,
    })
    expect(readWorkbook(bytes).definedNames).toEqual([
      { name: 'Foo', sheetName: "Bob's Sheet", range: 'A1', refersTo: "'Bob''s Sheet'!$A$1", hidden: undefined },
    ])
  })

  it('keeps an unresolvable reference with sheetName/range undefined', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Pi">3.14159</definedName><definedName name="Multi">Sheet1!$A$1,Sheet2!$B$2</definedName></definedNames></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/></worksheet>`,
    })
    const names = readWorkbook(bytes).definedNames
    expect(names).toEqual([
      { name: 'Pi', sheetName: undefined, range: undefined, refersTo: '3.14159', hidden: undefined },
      { name: 'Multi', sheetName: undefined, range: undefined, refersTo: 'Sheet1!$A$1,Sheet2!$B$2', hidden: undefined },
    ])
  })
})

describe('readWorkbook — column/row default styles', () => {
  const stylesXml = `<?xml version="1.0"?><styleSheet xmlns="x"><fonts count="1"><font/></fonts><fills count="3"><fill/><fill/><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="3"><xf/><xf fontId="0"/><xf fillId="2" applyFill="1"/></cellXfs></styleSheet>`

  it('resolves <col style> and <row customFormat s> only under { styles: true }', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/styles.xml': stylesXml,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><cols><col min="2" max="4" style="2"/></cols><sheetData><row r="5" s="2" customFormat="1"/><row r="6"><c r="A6"><v>1</v></c></row></sheetData></worksheet>`,
    })
    const sheet = readWorkbook(bytes, { styles: true }).sheet('S')!
    expect(sheet.columnStyles.get(2)).toEqual({ fill: 'FFFF0000' })
    expect(sheet.columnStyles.get(3)).toEqual({ fill: 'FFFF0000' })
    expect(sheet.columnStyles.get(4)).toEqual({ fill: 'FFFF0000' })
    expect(sheet.columnStyles.get(5)).toBeUndefined()
    expect(sheet.columnStyles.get(1)).toBeUndefined()
    expect(sheet.rowStyles.get(5)).toEqual({ fill: 'FFFF0000' })
    expect(sheet.rowStyles.get(6)).toBeUndefined()

    const withoutStyles = readWorkbook(bytes).sheet('S')!
    expect(withoutStyles.columnStyles.size).toBe(0)
    expect(withoutStyles.rowStyles.size).toBe(0)
  })

  it('ignores a row s="N" without customFormat="1" (not a real row default)', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/styles.xml': stylesXml,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1" s="2"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
    })
    const sheet = readWorkbook(bytes, { styles: true }).sheet('S')!
    expect(sheet.rowStyles.size).toBe(0)
  })
})

describe('readWorkbook — formula cells with a date format', () => {
  it('converts a formula result to a Date when its style is a date format', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/styles.xml': `<?xml version="1.0"?><styleSheet xmlns="x"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><f>TODAY()</f><v>45900</v></c></row></sheetData></worksheet>`,
    })
    const cell = readWorkbook(bytes).sheet('S')!.cell('A1')!
    expect(cell.type).toBe('formula')
    expect(cell.value).toBeInstanceOf(Date)
    expect((cell.value as Date).getFullYear()).toBe(2025)
    expect(cell.formula).toBe('TODAY()')
  })

  it('leaves a plain-numeric-format formula result as a number', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><f>SUM(1,2)</f><v>3</v></c></row></sheetData></worksheet>`,
    })
    const cell = readWorkbook(bytes).sheet('S')!.cell('A1')!
    expect(cell.value).toBe(3)
  })
})
