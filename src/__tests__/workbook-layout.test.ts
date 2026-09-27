import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { readWorkbook } from '../read'
import { createWorkbook } from '../workbook'

function build(fn: (wb: ReturnType<typeof createWorkbook>) => void): Record<string, string> {
  const wb = createWorkbook()
  fn(wb)
  const entries = unzipSync(wb.xlsx())
  const out: Record<string, string> = {}
  for (const [path, data] of Object.entries(entries)) out[path] = strFromU8(data)
  return out
}

describe('merged cells', () => {
  it('emits <mergeCells> after <sheetData>', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setCell('A1', 'Title')
      s.merge('A1:C1')
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>')
    expect(sheet.indexOf('</sheetData>')).toBeLessThan(sheet.indexOf('<mergeCells'))
    expect(sheet).toContain('<dimension ref="A1:C1"/>') // merge extends the dimension
  })

  it('normalises and de-duplicates via count', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setCell('A1', 'x')
      s.merge('A1:B2')
      s.merge('D1:D4')
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<mergeCells count="2">')
  })

  it('rejects a single-cell merge, a backwards range, and overlaps', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    expect(() => s.merge('A1:A1')).toThrow(/single cell/)
    expect(() => s.merge('C3:A1')).toThrow(/top-left to bottom-right/)
    s.merge('A1:C3')
    expect(() => s.merge('B2:D4')).toThrow(/overlaps/)
  })
})

describe('column specs', () => {
  it('emits <cols> before <sheetData> with width + customWidth', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').setColumn(1, { width: 24 }).addRow(['x'])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<cols><col min="1" max="1" width="24" customWidth="1"/></cols>')
    expect(sheet.indexOf('<cols>')).toBeLessThan(sheet.indexOf('<sheetData>'))
  })

  it('supports hidden columns and a column default style', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setColumn(2, { hidden: true })
      s.setColumn(3, { style: { numFmt: '0.00%' } })
      s.addRow([1, 2, 0.5])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<col min="2" max="2" hidden="1"/>')
    expect(sheet).toMatch(/<col min="3" max="3" style="\d+"\/>/)
    // the cell in column 3 picks up the column's style
    expect(sheet).toMatch(/<c r="C1" s="\d+">/)
    // 0.00% is built-in numFmtId 10 (no custom <numFmt> emitted)
    expect(parts['xl/styles.xml']).toMatch(/<xf numFmtId="10"[^>]*applyNumberFormat="1"/)
  })

  it('resolves style precedence column < row < cell', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setColumn(1, { style: { font: { name: 'Arial' } } })
      s.addRow([{ value: 'x', style: { font: { bold: true } } }], {
        style: { font: { italic: true } },
      })
    })
    // effective font on A1 = Arial (col) + italic (row) + bold (cell)
    expect(parts['xl/styles.xml']).toMatch(/<font><b\/><i\/><sz val="11"\/><name val="Arial"\/>/)
  })

  it('accepts a columns array on addWorksheet', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S', { columns: [{ width: 10 }, { width: 20 }] }).addRow(['a', 'b'])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<col min="1" max="1" width="10" customWidth="1"/>')
    expect(sheet).toContain('<col min="2" max="2" width="20" customWidth="1"/>')
  })

  it('rejects an out-of-range column index', () => {
    const s = createWorkbook().addWorksheet('S')
    expect(() => s.setColumn(0, {})).toThrow()
    expect(() => s.setColumn(16385, {})).toThrow()
  })
})

describe('row height', () => {
  it('sets ht + customHeight on the row', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').addRow(['tall'], { height: 30 })
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<row r="1" ht="30" customHeight="1">')
  })

  it('setRow can style/size a row that has no cells', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['r1'])
      s.setRow(3, { height: 40 })
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<row r="3" ht="40" customHeight="1"></row>')
  })

  it('rejects a non-positive height', () => {
    const s = createWorkbook().addWorksheet('S')
    expect(() => s.addRow(['x'], { height: 0 })).toThrow(/height/)
  })
})

describe('freeze panes', () => {
  it('freezes the first row', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S', { freeze: { ySplit: 1 } }).addRow(['h']).addRow(['v'])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain(
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
    )
    expect(sheet.indexOf('<sheetViews>')).toBeLessThan(sheet.indexOf('<sheetData>'))
  })

  it('freezes the first column', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').freeze({ xSplit: 1 }).addRow(['a', 'b'])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>',
    )
  })

  it('freezes rows and columns together', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').freeze({ xSplit: 2, ySplit: 1 }).addRow(['a'])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain(
      '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>',
    )
  })

  it('freeze({}) / freeze zero clears the pane', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S').freeze({ xSplit: 1 }).freeze({}).addRow(['a'])
    })
    expect(parts['xl/worksheets/sheet1.xml']).not.toContain('<sheetViews>')
  })

  it('rejects negative / non-integer splits', () => {
    const s = createWorkbook().addWorksheet('S')
    expect(() => s.freeze({ ySplit: -1 })).toThrow()
    expect(() => s.freeze({ xSplit: 1.5 })).toThrow()
  })
})

describe('auto-filter', () => {
  it('adds <autoFilter> and the workbook _FilterDatabase defined name', () => {
    const parts = build((wb) => {
      wb.addWorksheet('Sales', { autoFilter: 'A1:C1' }).addRow(['a', 'b', 'c'])
    })
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<autoFilter ref="A1:C1"/>')
    expect(parts['xl/workbook.xml']).toContain(
      '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'Sales\'!$A$1:$C$1</definedName>',
    )
  })

  it('places <autoFilter> before <mergeCells>', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['a', 'b'])
      s.autoFilter('A1:B1')
      s.merge('A3:B3')
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet.indexOf('<autoFilter')).toBeLessThan(sheet.indexOf('<mergeCells'))
  })

  it('has no defined name when no sheet filters', () => {
    const parts = build((wb) => wb.addWorksheet('S').addRow([1]))
    expect(parts['xl/workbook.xml']).not.toContain('definedNames')
  })
})

describe('outline (grouping)', () => {
  it('emits hidden + outlineLevel on rows and columns', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.setColumn(2, { hidden: true, outlineLevel: 1 })
      s.addRow(['a'])
      s.addRow(['b'], { hidden: true, outlineLevel: 2 })
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<col min="2" max="2" hidden="1" outlineLevel="1"/>')
    expect(sheet).toContain('<row r="2" hidden="1" outlineLevel="2">')
  })

  it('emits <sheetPr><outlinePr> when summary direction is non-default', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S', { outline: { summaryBelow: false, summaryRight: false } }).addRow(['a'])
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet).toContain('<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>')
    expect(sheet.indexOf('<sheetPr>')).toBeLessThan(sheet.indexOf('<dimension'))
  })

  it('omits <sheetPr> entirely when outline direction is left at Excel defaults', () => {
    const parts = build((wb) => {
      wb.addWorksheet('S', { outline: {} }).addRow(['a'])
    })
    expect(parts['xl/worksheets/sheet1.xml']).not.toContain('<sheetPr>')
  })

  it('rejects an out-of-range outline level', () => {
    const s = createWorkbook().addWorksheet('S')
    expect(() => s.setColumn(1, { outlineLevel: 8 })).toThrow(/outlineLevel/)
    expect(() => s.addRow(['a'], { outlineLevel: -1 })).toThrow(/outlineLevel/)
  })

  it('round-trips row/column outline info through readWorkbook', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.setColumn(2, { hidden: true, outlineLevel: 1 })
    s.addRow(['a'])
    s.addRow(['b'], { outlineLevel: 2 })
    const sheet = readWorkbook(wb.xlsx()).sheet('S')!
    expect(sheet.columnInfo.get(2)).toMatchObject({ hidden: true, outlineLevel: 1 })
    expect(sheet.rowInfo.get(2)).toMatchObject({ outlineLevel: 2 })
  })
})

describe('sheet/workbook protection', () => {
  it('protect() with no options emits sheet="1" only (Excel UI defaults)', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['a'])
    s.protect()
    const sheet = strFromU8(unzipSync(wb.xlsx())['xl/worksheets/sheet1.xml']!)
    expect(sheet).toContain('<sheetProtection sheet="1"/>')
  })

  it('translates positive "allow" options into inverted-polarity XML attrs', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['a'])
    s.protect({
      allowSelectLockedCells: false,
      allowSelectUnlockedCells: false,
      allowFormatCells: true,
      allowSort: true,
    })
    const sheet = strFromU8(unzipSync(wb.xlsx())['xl/worksheets/sheet1.xml']!)
    expect(sheet).toContain('selectLockedCells="1"')
    expect(sheet).toContain('selectUnlockedCells="1"')
    expect(sheet).toContain('formatCells="0"')
    expect(sheet).toContain('sort="0"')
    expect(sheet).not.toContain('formatColumns')
  })

  it('places <sheetProtection> after </sheetData> and before <autoFilter>', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S')
      s.addRow(['a'])
      s.autoFilter('A1:A1')
      s.protect()
    })
    const sheet = parts['xl/worksheets/sheet1.xml']!
    expect(sheet.indexOf('</sheetData>')).toBeLessThan(sheet.indexOf('<sheetProtection'))
    expect(sheet.indexOf('<sheetProtection')).toBeLessThan(sheet.indexOf('<autoFilter'))
  })

  it('wb.protect() defaults lockStructure to true, and emits lockWindows when set', () => {
    const wb = createWorkbook()
    wb.addWorksheet('S').addRow(['a'])
    wb.protect()
    expect(strFromU8(unzipSync(wb.xlsx())['xl/workbook.xml']!)).toContain(
      '<workbookProtection lockStructure="1"/>',
    )

    const wb2 = createWorkbook()
    wb2.addWorksheet('S').addRow(['a'])
    wb2.protect({ lockStructure: false, lockWindows: true })
    expect(strFromU8(unzipSync(wb2.xlsx())['xl/workbook.xml']!)).toContain(
      '<workbookProtection lockWindows="1"/>',
    )
  })

  it('round-trips sheet + workbook protection through readWorkbook', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['a'])
    s.protect({ allowSort: true })
    wb.protect({ lockWindows: true })
    const read = readWorkbook(wb.xlsx())
    expect(read.protection).toEqual({ lockStructure: true, lockWindows: true })
    expect(read.sheet('S')!.protection).toMatchObject({ allowSort: true, allowFormatCells: false })
  })

  it('cell-level protection styling is only carried when non-default, and round-trips', () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('S')
    s.setCell('A1', 'unlocked', { protection: { locked: false } })
    s.setCell('A2', 'hidden-formula', { protection: { hidden: true } })
    s.setCell('A3', 'default', { protection: { locked: true, hidden: false } })
    const bytes = wb.xlsx()
    const styles = strFromU8(unzipSync(bytes)['xl/styles.xml']!)
    expect(styles).toContain('<protection locked="0"/>')
    expect(styles).toContain('<protection hidden="1"/>')

    const read = readWorkbook(bytes, { styles: true }).sheet('S')!
    expect(read.cell('A1')!.style?.protection).toEqual({ locked: false })
    expect(read.cell('A2')!.style?.protection).toEqual({ hidden: true })
    expect(read.cell('A3')!.style?.protection).toBeUndefined()
  })
})

describe('child-element order (schema compliance)', () => {
  it('dimension → sheetViews → cols → sheetData → autoFilter → mergeCells', () => {
    const parts = build((wb) => {
      const s = wb.addWorksheet('S', {
        columns: [{ width: 10 }],
        freeze: { ySplit: 1 },
        autoFilter: 'A1:B1',
      })
      s.addRow(['a', 'b'])
      s.merge('A3:B3')
    })
    const xml = parts['xl/worksheets/sheet1.xml']!
    const order = ['<dimension', '<sheetViews>', '<cols>', '<sheetData>', '<autoFilter', '<mergeCells']
    const positions = order.map((tag) => xml.indexOf(tag))
    expect(positions).toEqual([...positions].sort((a, b) => a - b))
    expect(positions.every((p) => p >= 0)).toBe(true)
  })
})
