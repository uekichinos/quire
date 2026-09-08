// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest'
import { createWorkbook, readWorkbook } from '../index'

/** Guards the "identical in Node and the browser" claim — no `node:` globals leak in. */
describe('browser environment', () => {
  it('write → blob → read works with DOM Blob / URL', async () => {
    const wb = createWorkbook()
    const s = wb.addWorksheet('Sales')
    s.addRow(['Product', 'Revenue'])
    s.addRow(['Widget', 15003.4])
    s.addRow([{ value: 0.25, style: { numFmt: '0.00%' } }, new Date(2024, 2, 1)])

    const blob = wb.blob()
    expect(blob).toBeInstanceOf(Blob)
    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    expect(blob.size).toBeGreaterThan(400)

    const url = URL.createObjectURL(blob)
    expect(typeof url).toBe('string')
    URL.revokeObjectURL(url)

    const bytes = new Uint8Array(await blob.arrayBuffer())
    const back = readWorkbook(bytes, { styles: true })
    expect(back.sheet('Sales')!.values()).toEqual([
      ['Product', 'Revenue'],
      ['Widget', 15003.4],
      [0.25, new Date(2024, 2, 1)],
    ])
    expect(back.sheet('Sales')!.cell('A3')!.numFmt).toBe('0.00%')
  })
})
