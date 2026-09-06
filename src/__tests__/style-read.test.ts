import { zipSync, strToU8 } from 'fflate'
import { describe, expect, it } from 'vitest'
import { createWorkbook } from '../workbook'
import { readWorkbook } from '../read'

function styledRoundTrip(fn: (wb: ReturnType<typeof createWorkbook>) => void) {
  const wb = createWorkbook()
  fn(wb)
  return readWorkbook(wb.xlsx(), { styles: true })
}

describe('styles on read — round-trips the writer', () => {
  it('font: bold / italic / underline / size / name / colour', () => {
    const back = styledRoundTrip((w) => {
      w.addWorksheet('S').addRow([
        {
          value: 'H',
          style: {
            font: { bold: true, italic: true, underline: true, size: 14, name: 'Arial', color: 'FF0000' },
          },
        },
      ])
    })
    expect(back.sheet('S')!.cell('A1')!.style).toEqual({
      font: { bold: true, italic: true, underline: true, size: 14, name: 'Arial', color: 'FFFF0000' },
    })
  })

  it('solid fill colour', () => {
    const back = styledRoundTrip((w) => {
      w.addWorksheet('S').addRow([{ value: 1, style: { fill: '2F5597' } }])
    })
    expect(back.sheet('S')!.cell('A1')!.style).toEqual({ fill: 'FF2F5597' })
  })

  it('borders per edge, colour defaulting to black', () => {
    const back = styledRoundTrip((w) => {
      w.addWorksheet('S').addRow([
        { value: 1, style: { border: { top: { style: 'thin' }, bottom: { style: 'medium', color: '0000FF' } } } },
      ])
    })
    expect(back.sheet('S')!.cell('A1')!.style!.border).toEqual({
      top: { style: 'thin', color: 'FF000000' },
      bottom: { style: 'medium', color: 'FF0000FF' },
    })
  })

  it('alignment', () => {
    const back = styledRoundTrip((w) => {
      w.addWorksheet('S').addRow([
        { value: 1, style: { align: { horizontal: 'center', vertical: 'middle', wrapText: true, indent: 2 } } },
      ])
    })
    // writer maps vertical 'middle' → OOXML 'center'
    expect(back.sheet('S')!.cell('A1')!.style!.align).toEqual({
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
      indent: 2,
    })
  })

  it('resolved number-format code (custom and built-in)', () => {
    const back = styledRoundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow([{ value: 1.5, style: { numFmt: '0.000"kg"' } }])
      s.addRow([{ value: 0.25, style: { numFmt: '0.00%' } }])
    })
    expect(back.sheet('S')!.cell('A1')!.style!.numFmt).toBe('0.000"kg"')
    expect(back.sheet('S')!.cell('A2')!.style!.numFmt).toBe('0.00%') // built-in id 10
  })

  it('un-styled cells have no style; row-default style is resolved on each cell', () => {
    const back = styledRoundTrip((w) => {
      const s = w.addWorksheet('S')
      s.addRow(['plain'])
      s.addRow(['a', 'b'], { style: { font: { bold: true } } })
    })
    expect(back.sheet('S')!.cell('A1')!.style).toBeUndefined()
    // the writer always records sz + name on a font, so they come back too
    const a2 = back.sheet('S')!.cell('A2')!.style
    expect(a2!.font).toMatchObject({ bold: true, size: 11, name: 'Calibri' })
    expect(back.sheet('S')!.cell('B2')!.style!.font).toMatchObject({ bold: true })
  })

  it('is off by default', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow([{ value: 1, style: { font: { bold: true } } }])
    const back = readWorkbook(wb.xlsx())
    expect(back.sheet('S')!.cell('A1')!.style).toBeUndefined()
  })
})

/* --- files with indexed / theme colours (the writer never emits these) ------ */

const CT = `<?xml version="1.0"?><Types xmlns="x"><Default Extension="rels" ContentType="x"/><Default Extension="xml" ContentType="x"/></Types>`
const RELS = `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="xl/workbook.xml"/></Relationships>`
const WB = `<?xml version="1.0"?><workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
const WB_RELS = `<?xml version="1.0"?><Relationships xmlns="x"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`

function make(styles: string, theme: string, sheet: string): Uint8Array {
  const z: Record<string, Uint8Array> = {}
  for (const [k, v] of Object.entries({
    '[Content_Types].xml': CT,
    '_rels/.rels': RELS,
    'xl/workbook.xml': WB,
    'xl/_rels/workbook.xml.rels': WB_RELS,
    'xl/styles.xml': styles,
    'xl/theme/theme1.xml': theme,
    'xl/worksheets/sheet1.xml': sheet,
  })) {
    z[k] = strToU8(v)
  }
  return zipSync(z)
}

const THEME1 = `<?xml version="1.0"?><a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
</a:clrScheme></a:themeElements></a:theme>`

describe('styles on read — indexed & theme colours', () => {
  it('resolves an indexed font colour', () => {
    const styles = `<?xml version="1.0"?><styleSheet xmlns="x">
      <fonts count="2"><font><sz val="11"/></font><font><color indexed="10"/><sz val="11"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>
    </styleSheet>`
    const sheet = `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>`
    const wb = readWorkbook(make(styles, THEME1, sheet), { styles: true })
    expect(wb.sheet('S')!.cell('A1')!.style!.font!.color).toBe('FFFF0000') // indexed 10 = red
  })

  it('resolves a theme fill colour with the 0/1 swap and applies tint', () => {
    const styles = `<?xml version="1.0"?><styleSheet xmlns="x">
      <fonts count="1"><font><sz val="11"/></font></fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor theme="4"/></patternFill></fill>
      </fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2"><xf/><xf fillId="2" applyFill="1"/></cellXfs>
    </styleSheet>`
    const sheet = `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>`
    const wb = readWorkbook(make(styles, THEME1, sheet), { styles: true })
    // theme 4 = accent1 = 4472C4
    expect(wb.sheet('S')!.cell('A1')!.style!.fill).toBe('FF4472C4')
  })

  it('keeps the style on an empty styled cell', () => {
    const styles = `<?xml version="1.0"?><styleSheet xmlns="x">
      <fonts count="1"><font><sz val="11"/></font></fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
      </fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2"><xf/><xf fillId="2" applyFill="1"/></cellXfs>
    </styleSheet>`
    const sheet = `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>`
    const wb = readWorkbook(make(styles, THEME1, sheet), { styles: true })
    const c = wb.sheet('S')!.cell('A1')!
    expect(c.type).toBe('empty')
    expect(c.value).toBeNull()
    expect(c.style).toEqual({ fill: 'FFFFFF00' })
  })

  it('theme index 1 maps to dk1 (text), index 0 to lt1 (background)', () => {
    const styles = `<?xml version="1.0"?><styleSheet xmlns="x">
      <fonts count="2"><font><sz val="11"/></font><font><color theme="1"/><sz val="11"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>
    </styleSheet>`
    const sheet = `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>`
    const wb = readWorkbook(make(styles, THEME1, sheet), { styles: true })
    expect(wb.sheet('S')!.cell('A1')!.style!.font!.color).toBe('FF000000') // dk1 windowText
  })
})
