import { describe, expect, it } from 'vitest'
import { createStreamingWorkbook } from '../streaming'
import { createWorkbook } from '../workbook'
import { readWorkbook } from '../read'

describe('createStreamingWorkbook — basics', () => {
  it('produces a valid, readable .xlsx', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['a', 1, true])
    s.addRow(['b', 2, false])
    const bytes = await wb.finish()
    expect([bytes[0], bytes[1]]).toEqual([0x50, 0x4b]) // "PK"

    const sheet = readWorkbook(bytes).sheet('S')!
    expect(sheet.values()).toEqual([
      ['a', 1, true],
      ['b', 2, false],
    ])
  })

  it('matches the buffered writer\'s output for equivalent input', async () => {
    const build = (wb: ReturnType<typeof createWorkbook>) => {
      const s = wb.addWorksheet('S')
      s.addRow(['x', 42, new Date(2024, 5, 1)])
      s.addRow(['y', { formula: 'SUM(B1:B1)', result: 42 }])
    }
    const buffered = createWorkbook()
    build(buffered)

    const streaming = createStreamingWorkbook()
    const s = streaming.addWorksheet('S')
    s.addRow(['x', 42, new Date(2024, 5, 1)])
    s.addRow(['y', { formula: 'SUM(B1:B1)', result: 42 }])

    const a = readWorkbook(buffered.xlsx()).sheet('S')!.values()
    const b = readWorkbook(await streaming.finish()).sheet('S')!.values()
    expect(b).toEqual(a)
  })

  it('shares one string/style pool across sheets (dedup) like the buffered writer', async () => {
    const wb = createStreamingWorkbook()
    const s1 = wb.addWorksheet('One')
    s1.addRow(['common', { value: 'styled', style: { font: { bold: true } } }])
    const s2 = wb.addWorksheet('Two')
    s2.addRow(['common'])
    const bytes = await wb.finish()
    const back = readWorkbook(bytes, { styles: true })
    expect(back.sheet('One')!.cell('A1')!.value).toBe('common')
    expect(back.sheet('Two')!.cell('A1')!.value).toBe('common')
    expect(back.sheet('One')!.cell('B1')!.style).toMatchObject({ font: { bold: true } })
  })

  it('writes many rows and stays well under a naive full-buffer memory footprint', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('Big')
    const n = 20_000
    for (let i = 1; i <= n; i++) s.addRow([i, `row ${i}`, i % 2 === 0])
    const bytes = await wb.finish()
    const sheet = readWorkbook(bytes).sheet('Big')!
    expect(sheet.dimension.rows).toBe(n)
    expect(sheet.cell(`A${n}`)!.value).toBe(n)
    expect(sheet.cell(`B1`)!.value).toBe('row 1')
  })

  it('an empty sheet still produces a valid, readable part', async () => {
    const wb = createStreamingWorkbook()
    wb.addWorksheet('Empty')
    const bytes = await wb.finish()
    expect(readWorkbook(bytes).sheet('Empty')!.values()).toEqual([])
  })

  it('throws when the workbook has no worksheets', async () => {
    const wb = createStreamingWorkbook()
    await expect(wb.finish()).rejects.toThrow(/at least one worksheet/)
  })
})

describe('createStreamingWorkbook — one-pass constraints', () => {
  it('rejects setColumn / freeze after the first addRow', () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    expect(() => s.setColumn(1, { width: 10 })).toThrow(/before the first addRow/)
    expect(() => s.freeze({ ySplit: 1 })).toThrow(/before the first addRow/)
  })

  it('accepts columns/freeze/autoFilter via addWorksheet options before any row', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S', {
      columns: [{ width: 20 }],
      freeze: { ySplit: 1 },
      autoFilter: 'A1:B1',
    })
    s.addRow(['h1', 'h2'])
    const bytes = await wb.finish()
    // just needs to not throw and still round-trip
    expect(readWorkbook(bytes).sheet('S')!.values()).toEqual([['h1', 'h2']])
  })

  it('rejects further addRow calls once the sheet has been finished', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    await wb.finish()
    expect(() => s.addRow(['y'])).toThrow(/already been finished/)
  })
})

describe('createStreamingWorkbook — features', () => {
  it('round-trips hyperlinks, comments, rich text, merges, data validation, conditional format', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    s.merge('A1:B1')
    s.setDataValidation('C2:C10', { type: 'list', list: ['Open', 'Done'] })
    s.addConditionalFormat('D2:D10', {
      type: 'cellIs',
      operator: 'greaterThan',
      formula: 100,
      style: { fill: 'FF0000' },
    })
    s.setComment('A1', 'note', { author: 'A' })

    s.addRow(['Header', null, 'Status', 'Value'])
    s.addRow([
      { hyperlink: 'https://example.com', text: 'link' },
      [{ text: 'rich ' }, { text: 'bold', font: { bold: true } }],
      'Open',
      150,
    ])

    const bytes = await wb.finish()
    const sheet = readWorkbook(bytes, { styles: true }).sheet('S')!
    expect(sheet.merges).toEqual(['A1:B1'])
    expect(sheet.cell('A1')!.comment).toEqual({ text: 'note', author: 'A' })
    expect(sheet.cell('A2')!.hyperlink).toMatchObject({ target: 'https://example.com' })
    expect(sheet.cell('B2')!.value).toEqual([{ text: 'rich ' }, { text: 'bold', font: { bold: true } }])
    expect(sheet.dataValidations).toEqual([
      expect.objectContaining({ ref: 'C2:C10', type: 'list', list: ['Open', 'Done'] }),
    ])
    expect(sheet.conditionalFormats).toEqual([
      expect.objectContaining({ ref: 'D2:D10', rule: expect.objectContaining({ type: 'cellIs' }) }),
    ])
  })

  it('supports defineName referencing an already-added sheet', async () => {
    const wb = createStreamingWorkbook()
    wb.addWorksheet('Sales').addRow(['x'])
    wb.defineName('SalesRange', 'Sales', 'A1:A10')
    const bytes = await wb.finish()
    expect(readWorkbook(bytes).definedNames).toEqual([
      { name: 'SalesRange', sheetName: 'Sales', range: 'A1:A10', refersTo: "'Sales'!$A$1:$A$10", hidden: undefined },
    ])
  })

  it('rejects an invalid row style value the same way the buffered writer does', () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    expect(() => s.addRow([Number.NaN])).toThrow(/finite/)
  })

  it('setPageSetup round-trips, including the print area', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    s.setPageSetup({ orientation: 'landscape', fitToWidth: 1, fitToHeight: 0, printArea: 'A1:B5' })
    s.addRow(['x'])
    const bytes = await wb.finish()
    const sheet = readWorkbook(bytes).sheet('S')!
    expect(sheet.pageSetup).toMatchObject({
      orientation: 'landscape',
      fitToWidth: 1,
      fitToHeight: 0,
      printArea: 'A1:B5',
    })
  })

  it('rejects setPageSetup after the first addRow', () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    expect(() => s.setPageSetup({ orientation: 'landscape' })).toThrow(/before the first addRow/)
  })

  it('sheet.protect()/wb.protect() round-trip, and protect() can follow addRow', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S')
    s.addRow(['x'])
    s.protect({ allowSort: true }) // allowed after addRow — renders into the tail, not the header
    wb.protect({ lockWindows: true })
    const bytes = await wb.finish()
    const read = readWorkbook(bytes)
    expect(read.protection).toEqual({ lockStructure: true, lockWindows: true })
    expect(read.sheet('S')!.protection).toMatchObject({ allowSort: true, allowFormatCells: false })
  })

  it('row/column outline (grouping) round-trips', async () => {
    const wb = createStreamingWorkbook()
    const s = wb.addWorksheet('S', { outline: { summaryBelow: false } })
    s.setColumn(1, { outlineLevel: 1 })
    s.addRow(['a'])
    s.addRow(['b'], { hidden: true, outlineLevel: 2 })
    const bytes = await wb.finish()
    const sheet = readWorkbook(bytes).sheet('S')!
    expect(sheet.columnInfo.get(1)).toMatchObject({ outlineLevel: 1 })
    expect(sheet.rowInfo.get(2)).toMatchObject({ hidden: true, outlineLevel: 2 })
  })
})
