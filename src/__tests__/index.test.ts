import { describe, expect, it } from 'vitest'
import * as quire from '../index'

describe('@uekichinos/quire — public surface', () => {
  it('exports the writer entry point and helpers', () => {
    expect(typeof quire.createWorkbook).toBe('function')
    expect(typeof quire.colLetter).toBe('function')
    expect(typeof quire.parseRef).toBe('function')
    expect(typeof quire.escapeXml).toBe('function')
  })

  it('exports the reader entry point and errors', () => {
    expect(typeof quire.readWorkbook).toBe('function')
    expect(typeof quire.XlsxReadError).toBe('function')
    expect(typeof quire.XmlError).toBe('function')
  })

  it('write then read round-trips through the public API', () => {
    const wb = quire.createWorkbook()
    wb.addWorksheet('Report')
      .addRow(['Product', 'Revenue'])
      .addRow(['Widget', 15003.4])
      .addRow(['Gadget', 2450])

    const back = quire.readWorkbook(wb.xlsx())
    expect(back.sheetNames).toEqual(['Report'])
    expect(back.sheet('Report')!.values()).toEqual([
      ['Product', 'Revenue'],
      ['Widget', 15003.4],
      ['Gadget', 2450],
    ])
  })

  it.todo('R2 — dates, formulas, merges')
  it.todo('R3 — styles on read')
})
