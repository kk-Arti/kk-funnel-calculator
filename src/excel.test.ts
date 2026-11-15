import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { writeFileSync } from 'node:fs'
import { buildWorkbook } from './excel'
import { EXAMPLE } from './calc'

describe('excel export', () => {
  it('writes inputs, formulas and cached results that match the page', async () => {
    const blob = await buildWorkbook({ ...EXAMPLE })
    const buf = Buffer.from(await blob.arrayBuffer())
    if (process.env.XLSX_OUT) writeFileSync(process.env.XLSX_OUT, buf)
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(buf)
    const ws = wb.getWorksheet('Funnel')!
    expect(ws.getCell('C6').value).toBe(10000)
    const res = (a: string) => (ws.getCell(a).value as { result: number }).result
    expect(res('D22')).toBe(204)
    expect(+res('F22').toFixed(2)).toBe(73.53)
    expect(+res('F20').toFixed(2)).toBe(22.06)
  })
})
