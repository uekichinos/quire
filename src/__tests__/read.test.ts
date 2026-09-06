import { zipSync, strToU8 } from 'fflate'
import { describe, expect, it } from 'vitest'
import { createWorkbook } from '../workbook'
import { readWorkbook } from '../read'
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

  it('accepts an ArrayBuffer', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['x'])
    const buf = wb.xlsx().buffer
    expect(readWorkbook(buf as ArrayBuffer).sheet('S')!.cell('A1')!.value).toBe('x')
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
