import { zipSync, strToU8 } from 'fflate'
import { describe, expect, it } from 'vitest'
import { readWorkbook } from '../read'

/**
 * Small hand-built workbooks reproducing quirks seen in files from Excel,
 * Google Sheets, LibreOffice and openpyxl — things quire's own writer never
 * emits but a real import will hit.
 */

const CT = `<?xml version="1.0"?><Types xmlns="x"><Default Extension="rels" ContentType="x"/><Default Extension="xml" ContentType="x"/></Types>`
const RELS = `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="xl/workbook.xml"/></Relationships>`
const WB = `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
const WB_RELS = `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`

function xlsx(sheet: string, extra: Record<string, string> = {}): Uint8Array {
  const src: Record<string, string> = {
    '[Content_Types].xml': CT,
    '_rels/.rels': RELS,
    'xl/workbook.xml': WB,
    'xl/_rels/workbook.xml.rels': WB_RELS,
    'xl/worksheets/sheet1.xml': sheet,
    ...extra,
  }
  const z: Record<string, Uint8Array> = {}
  for (const [k, v] of Object.entries(src)) z[k] = strToU8(v)
  return zipSync(z)
}

describe('real-world quirks', () => {
  it('ignores <sheetPr>, <sheetViews> and an mc:AlternateContent wrapper (Excel)', () => {
    const sheet = `<?xml version="1.0"?>
      <worksheet xmlns:mc="x" mc:Ignorable="x14ac">
        <sheetPr><outlinePr summaryBelow="0"/></sheetPr>
        <dimension ref="A1:B2"/>
        <sheetViews><sheetView workbookViewId="0">
          <mc:AlternateContent><mc:Choice Requires="x"><pane/></mc:Choice></mc:AlternateContent>
        </sheetView></sheetViews>
        <sheetFormatPr defaultRowHeight="15"/>
        <sheetData>
          <row r="1" spans="1:2"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c></row>
          <row r="2" spans="1:2"><c r="A2"><v>30</v></c></row>
        </sheetData>
      </worksheet>`
    const s = readWorkbook(xlsx(sheet)).sheet('S')!
    expect(s.values()).toEqual([
      [10, 20],
      [30, null],
    ])
  })

  it('tolerates attribute order variance and whitespace around <v> (Sheets export)', () => {
    const sheet = `<?xml version="1.0"?><worksheet><sheetData>
      <row r="1">
        <c s="0" r="A1" t="s"><v> 0 </v></c>
        <c t="n" r="B1"><v>1.0</v></c>
        <c r="C1"><v>
          42
        </v></c>
      </row>
    </sheetData></worksheet>`
    const s = readWorkbook(
      xlsx(sheet, {
        'xl/sharedStrings.xml': `<?xml version="1.0"?><sst count="1" uniqueCount="1"><si><t>hello</t></si></sst>`,
        '[Content_Types].xml': CT,
      }),
    ).sheet('S')!
    expect(s.values()[0]).toEqual(['hello', 1, 42])
  })

  it('reads openpyxl-style output: BOM, no xml:space, style ids present', () => {
    const sheet = `﻿<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="x"><dimension ref="A1:A2"/><sheetData>
        <row r="1"><c r="A1" s="1"><v>44927</v></c></row>
        <row r="2"><c r="A2" t="s"><v>0</v></c></row>
      </sheetData></worksheet>`
    const styles = `<?xml version="1.0"?><styleSheet xmlns="x">
      <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>
      <fonts count="1"><font><sz val="11"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2"><xf/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs>
    </styleSheet>`
    const sst = `<?xml version="1.0"?><sst><si><t>note</t></si></sst>`
    const s = readWorkbook(
      xlsx(sheet, { 'xl/styles.xml': styles, 'xl/sharedStrings.xml': sst }),
    ).sheet('S')!
    expect(s.cell('A1')!.type).toBe('date')
    expect((s.cell('A1')!.value as Date).getFullYear()).toBe(2023)
    expect(s.cell('A1')!.numFmt).toBe('yyyy\\-mm\\-dd')
    expect(s.cell('A2')!.value).toBe('note')
  })

  it('reads a sheet target given as an absolute /xl/... path', () => {
    const wbRels = `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="/xl/worksheets/sheet1.xml"/></Relationships>`
    const sheet = `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`
    const s = readWorkbook(xlsx(sheet, { 'xl/_rels/workbook.xml.rels': wbRels })).sheet('S')!
    expect(s.cell('A1')!.value).toBe(7)
  })
})
