import { describe, expect, it } from 'vitest'
import * as quire from '../index'

describe('@uekichinos/quire — public surface', () => {
  it('exports the entry point and helpers', () => {
    expect(typeof quire.createWorkbook).toBe('function')
    expect(typeof quire.colLetter).toBe('function')
    expect(typeof quire.parseRef).toBe('function')
    expect(typeof quire.escapeXml).toBe('function')
    expect(quire.version).toBe('0.0.0')
  })

  it('runs the README example end to end', () => {
    const wb = quire.createWorkbook()
    const sheet = wb.addWorksheet('Report')
    sheet
      .addRow(['Product', 'Revenue'])
      .addRow(['Widget', 15003.4])
      .addRow(['Gadget', 2450])
    expect(wb.xlsx()).toBeInstanceOf(Uint8Array)
  })

  it.todo('Phase 2 — styles, dates and formula cells')
  it.todo('Phase 3 — merges, column widths, freeze panes')
  it.todo('Phase 4 — LibreOffice round-trip gate in CI')
})
