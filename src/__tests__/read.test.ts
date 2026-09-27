import { zipSync, strToU8 } from 'fflate'
import { describe, expect, it } from 'vitest'
import { createWorkbook } from '../workbook'
import { readWorkbook, readWorkbookAsync, readRows } from '../read'
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

describe('readWorkbook — rich text', () => {
  it('a single formatted run does NOT collapse to a plain string', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': `<?xml version="1.0"?><sst><si><r><rPr><b/></rPr><t>bold only</t></r></si></sst>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`,
    })
    const value = readWorkbook(bytes).sheet('S')!.cell('A1')!.value
    expect(value).toEqual([{ text: 'bold only', font: { bold: true } }])
  })

  it('a single unformatted run DOES collapse to a plain string', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': `<?xml version="1.0"?><sst><si><r><t>plain</t></r></si></sst>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`,
    })
    expect(readWorkbook(bytes).sheet('S')!.cell('A1')!.value).toBe('plain')
  })

  it('reads an indexed/theme colour and a rFont name on a run', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': `<?xml version="1.0"?><sst><si><r><rPr><sz val="14"/><rFont val="Arial"/><color rgb="ff0000"/></rPr><t>a</t></r><r><t>b</t></r></si></sst>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`,
    })
    const value = readWorkbook(bytes).sheet('S')!.cell('A1')!.value
    expect(value).toEqual([
      { text: 'a', font: { size: 14, name: 'Arial', color: 'FFFF0000' } },
      { text: 'b' },
    ])
  })
})

describe('readWorkbook — data validation', () => {
  it('surfaces list/whole/decimal/date/textLength/time/custom rules, ignoring type="none"', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/><dataValidations count="3">` +
        `<dataValidation type="none" sqref="A1"/>` +
        `<dataValidation type="whole" operator="greaterThan" sqref="B1"><formula1>0</formula1></dataValidation>` +
        `<dataValidation type="list" sqref="C1:C5" allowBlank="0"><formula1>"X,Y"</formula1></dataValidation>` +
        `</dataValidations></worksheet>`,
    })
    const dv = readWorkbook(bytes).sheet('S')!.dataValidations
    expect(dv).toEqual([
      { ref: 'B1', type: 'whole', operator: 'greaterThan', values: [0], allowBlank: false, promptTitle: undefined, promptMessage: undefined, errorTitle: undefined, errorMessage: undefined },
      { ref: 'C1:C5', type: 'list', allowBlank: false, promptTitle: undefined, promptMessage: undefined, errorTitle: undefined, errorMessage: undefined, list: ['X', 'Y'] },
    ])
  })
})

describe('readWorkbook — conditional formatting', () => {
  it('surfaces cellIs and colorScale, skips unsupported rule types, across multiple rules per block', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/>` +
        `<conditionalFormatting sqref="A1:A10">` +
        `<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/></dataBar></cfRule>` +
        `<cfRule type="cellIs" dxfId="0" priority="2" operator="lessThan"><formula>0</formula></cfRule>` +
        `</conditionalFormatting>` +
        `<conditionalFormatting sqref="B1:B10"><cfRule type="colorScale" priority="3">` +
        `<colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000"/><color rgb="00FF00"/></colorScale>` +
        `</cfRule></conditionalFormatting></worksheet>`,
    })
    const cf = readWorkbook(bytes).sheet('S')!.conditionalFormats
    expect(cf).toEqual([
      { ref: 'A1:A10', rule: { type: 'cellIs', operator: 'lessThan', formula: ['0'] } },
      { ref: 'B1:B10', rule: { type: 'colorScale', colors: ['FFFF0000', 'FF00FF00'] } },
    ])
  })

  it('resolves the dxf style only under { styles: true }', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/styles.xml': `<?xml version="1.0"?><styleSheet xmlns="x"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><font><b/><color rgb="FF0000"/></font><fill><patternFill><bgColor rgb="FFFF00"/></patternFill></fill></dxf></dxfs></styleSheet>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/><conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting></worksheet>`,
    })
    const withStyles = readWorkbook(bytes, { styles: true }).sheet('S')!.conditionalFormats[0]!
    expect(withStyles.rule).toEqual({
      type: 'cellIs',
      operator: 'greaterThan',
      formula: ['0'],
      style: { font: { bold: true, color: 'FFFF0000' }, fill: 'FFFFFF00' },
    })
    const withoutStyles = readWorkbook(bytes).sheet('S')!.conditionalFormats[0]!
    expect((withoutStyles.rule as { style?: unknown }).style).toBeUndefined()
  })
})

describe('readWorkbook — comments', () => {
  it('round-trips text and author via the sheet rels comments relationship', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><legacyDrawing r:id="rId2"/></worksheet>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`,
      'xl/comments1.xml': `<?xml version="1.0"?><comments xmlns="x"><authors><author>Alice</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>hi there</t></r></text></comment></commentList></comments>`,
    })
    const cell = readWorkbook(bytes).sheet('S')!.cell('A1')!
    expect(cell.comment).toEqual({ text: 'hi there', author: 'Alice' })
  })

  it('ignores a comments-typed relationship whose target part is missing, without throwing', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>`,
    })
    const cell = readWorkbook(bytes).sheet('S')!.cell('A1')!
    expect(cell.comment).toBeUndefined()
    expect(cell.value).toBe(1)
  })
})

describe('readWorkbook — data bars', () => {
  it('resolves a dataBar rule', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/>` +
        `<conditionalFormatting sqref="A1:A10"><cfRule type="dataBar" priority="1">` +
        `<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="638EC6"/></dataBar>` +
        `</cfRule></conditionalFormatting></worksheet>`,
    })
    const cf = readWorkbook(bytes).sheet('S')!.conditionalFormats
    expect(cf).toEqual([{ ref: 'A1:A10', rule: { type: 'dataBar', color: 'FF638EC6' } }])
  })
})

describe('readWorkbook — print setup', () => {
  it('resolves pageSetup, pageMargins, and the print area', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>` +
        `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'S'!$A$1:$C$10</definedName></definedNames></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/>` +
        `<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
        `<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/></worksheet>`,
    })
    const sheet = readWorkbook(bytes).sheet('S')!
    expect(sheet.pageSetup).toEqual({
      orientation: 'landscape',
      paperSize: 9,
      fitToWidth: 1,
      fitToHeight: 0,
      scale: undefined,
      margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      printArea: 'A1:C10',
    })
  })

  it('a sheet with neither pageSetup nor a print area has pageSetup undefined', () => {
    const bytes = makeXlsx({
      '[Content_Types].xml': CT,
      '_rels/.rels': RELS,
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet><sheetData/></worksheet>`,
    })
    expect(readWorkbook(bytes).sheet('S')!.pageSetup).toBeUndefined()
  })
})

describe('readRows — streaming', () => {
  it('delivers every row via onRow and matches readWorkbook values', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('Sales')
    s.addRow(['ID', 'Name'])
    for (let i = 1; i <= 20; i++) s.addRow([i, `Row ${i}`])
    const bytes = wb.xlsx()

    const seen: { row: number; values: unknown[] }[] = []
    const meta = readRows(bytes, 'Sales', (cells, row) => {
      seen.push({ row, values: cells.map((c) => c?.value) })
    })

    expect(seen).toHaveLength(21)
    expect(seen[0]).toEqual({ row: 1, values: ['ID', 'Name'] })
    expect(seen[20]).toEqual({ row: 21, values: [20, 'Row 20'] })
    expect(meta.sheetName).toBe('Sales')
    expect(meta.dimension).toEqual({ rows: 21, cols: 2 })
  })

  it('selects the sheet by index as well as by name', () => {
    const wb = createWorkbook()
    wb.addWorksheet('One').addRow(['a'])
    wb.addWorksheet('Two').addRow(['b'])
    const bytes = wb.xlsx()
    const rows: unknown[] = []
    const meta = readRows(bytes, 1, (cells) => rows.push(cells[0]?.value))
    expect(meta.sheetName).toBe('Two')
    expect(rows).toEqual(['b'])
  })

  it('throws for an unknown sheet name or out-of-range index', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['x'])
    const bytes = wb.xlsx()
    expect(() => readRows(bytes, 'Nope', () => {})).toThrow(/no sheet "Nope"/)
    expect(() => readRows(bytes, 5, () => {})).toThrow(/no sheet/)
  })

  it('delivers ragged rows as sparse arrays, holes undefined', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.setCell('A1', 'first')
    s.setCell('C1', 'third')
    const rows: (unknown[])[] = []
    readRows(wb.xlsx(), 'S', (cells) => rows.push(cells))
    expect(rows[0]![1]).toBeUndefined()
    expect((rows[0] as { value?: unknown }[])[0]!.value).toBe('first')
    expect((rows[0] as { value?: unknown }[])[2]!.value).toBe('third')
  })

  it('resolves styles per cell under { styles: true }', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.setCell('A1', 5, { numFmt: '0.00' })
    const rows: { numFmt?: string }[][] = []
    readRows(wb.xlsx(), 'S', (cells) => rows.push(cells as { numFmt?: string }[]), { styles: true })
    expect(rows[0]![0]!.numFmt).toBe('0.00')
  })

  it('does NOT attach hyperlinks or comments to streamed rows (documented limitation)', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.setCell('A1', { hyperlink: 'https://example.com', text: 'link' })
    s.setComment('A1', 'a note')
    const rows: { hyperlink?: unknown; comment?: unknown }[][] = []
    readRows(wb.xlsx(), 'S', (cells) => rows.push(cells as never))
    expect(rows[0]![0]!.hyperlink).toBeUndefined()
    expect(rows[0]![0]!.comment).toBeUndefined()
    // but readWorkbook still resolves them normally
    const full = readWorkbook(wb.xlsx()).sheet('S')!.cell('A1')!
    expect(full.hyperlink).toBeDefined()
    expect(full.comment).toBeDefined()
  })

  it('still enforces maxCells even in streaming mode', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('Big')
    for (let i = 1; i <= 50; i++) s.addRow([i, i, i])
    expect(() =>
      readRows(wb.xlsx(), 'Big', () => {}, { limits: { maxCells: 10 } }),
    ).toThrow(/read limit/)
  })

  it('surfaces dataValidations/conditionalFormats/merges/pageSetup alongside the stream', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    s.merge('A1:B1')
    s.setDataValidation('C1:C10', { type: 'list', list: ['A', 'B'] })
    s.addConditionalFormat('D1:D10', { type: 'dataBar', color: 'FF0000' })
    s.setPageSetup({ orientation: 'landscape' })
    const meta = readRows(wb.xlsx(), 'S', () => {})
    expect(meta.merges).toEqual(['A1:B1'])
    expect(meta.dataValidations).toHaveLength(1)
    expect(meta.conditionalFormats).toHaveLength(1)
    expect(meta.pageSetup).toMatchObject({ orientation: 'landscape' })
  })

  it('an empty sheet reports zero rows without throwing', () => {
    const wb = createWorkbook()
    wb.addWorksheet('Empty')
    let count = 0
    const meta = readRows(wb.xlsx(), 'Empty', () => count++)
    expect(count).toBe(0)
    expect(meta.dimension).toEqual({ rows: 0, cols: 0 })
  })
})
