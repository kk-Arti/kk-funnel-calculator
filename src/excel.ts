// Branded Excel export: a live copy of the funnel calculator with the visitor's numbers and working formulas.
// Formulas mirror calc.ts: people are rounded at every step, so Excel and the page show the same numbers.
import type { Workbook, Cell, PaperSize } from 'exceljs'
import { LOGO_PNG_BASE64 } from './assets/logo64'
import { EXAMPLE, HINTS, STAGES, compute, levers, parseNum, validate, type Inputs } from './calc'

const MESSAGE_URL =
  'https://www.linkedin.com/messaging/compose/?profileUrn=urn%3Ali%3Afsd_profile%3AACoAADxgieABwE7747g9IZ4BfTrBChanNiZJwr4'
const TOOL_URL = 'https://kk-arti.github.io/kk-funnel-calculator/'

const C = {
  dark: 'FF1C1C1C', dark2: 'FF242424', green: 'FF179158', greenDeep: 'FF072B1A', greenText: 'FF0F7A49',
  inputFill: 'FFEAF6EF', outFill: 'FFF5F5F5', grey: 'FF6B6B6B', line: 'FFD9D9D9', red: 'FFC0392B', blue: 'FF0000FF', white: 'FFFFFFFF',
}
const FONT = 'Arial'
const fill = (argb: string) => ({ type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb } })
const thin = { style: 'thin' as const, color: { argb: C.line } }
const box = { top: thin, left: thin, bottom: thin, right: thin }

function f(cell: Cell, formula: string, result: number | string | undefined) {
  cell.value = result === undefined ? { formula } : { formula, result }
}

export async function buildWorkbook(inp: Inputs): Promise<Blob> {
  const ExcelJS = (await import('exceljs')).default
  const wb: Workbook = new ExcelJS.Workbook()
  wb.creator = 'K&K Media'
  wb.title = 'Funnel Cost Calculator'
  const ws = wb.addWorksheet('Funnel', {
    views: [{ showGridLines: false }],
    pageSetup: { paperSize: 1 as unknown as PaperSize /* US Letter */, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  })
  ;[3, 34, 16, 16, 16, 18, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w))

  // Header band
  for (const r of [1, 2, 3, 4]) for (let c = 1; c <= 7; c++) ws.getCell(r, c).fill = fill(C.dark)
  ws.getRow(1).height = 12; ws.getRow(2).height = 30; ws.getRow(3).height = 20; ws.getRow(4).height = 12
  const logo = wb.addImage({ base64: LOGO_PNG_BASE64, extension: 'png' })
  ws.addImage(logo, { tl: { col: 1.05, row: 1.05 }, ext: { width: 42, height: 42 } })
  ws.getCell('C2').value = 'Funnel Cost Calculator'
  ws.getCell('C2').font = { name: FONT, size: 18, bold: true, color: { argb: C.white } }
  ws.getCell('C3').value = 'K&K Media, built by Artem Kalachev'
  ws.getCell('C3').font = { name: FONT, size: 10, color: { argb: 'FFCDCDCD' } }
  const msg = ws.getCell('E5')
  msg.value = { text: 'Message me on LinkedIn', hyperlink: MESSAGE_URL }
  msg.font = { name: FONT, size: 11, bold: true, color: { argb: C.greenText }, underline: true }
  const web = ws.getCell('G5')
  web.value = { text: 'Open the online calculator', hyperlink: TOOL_URL }
  web.font = { name: FONT, size: 10, color: { argb: C.grey }, underline: true }

  // Inputs (C6:C13). Blank inputs fall back to the example so the file always opens with a working funnel.
  const val = (k: keyof Inputs) => {
    const n = parseNum(inp[k])
    const use = n === null || Number.isNaN(n) ? parseNum(EXAMPLE[k])! : n
    return ['load', 'cart', 'checkout', 'order', 'margin'].includes(k) ? use / 100 : use
  }
  const rows: [string, keyof Inputs, string][] = [
    ['Clicks', 'clicks', '#,##0'],
    ['Cost per click ($)', 'cpc', '$#,##0.00'],
    ['Click to page loaded', 'load', '0.0%'],
    ['Page view to add to cart', 'cart', '0.0%'],
    ['Cart to checkout', 'checkout', '0.0%'],
    ['Checkout to order', 'order', '0.0%'],
    ['Average order value ($)', 'aov', '$#,##0.00'],
    ['Margin after COGS, shipping and fees', 'margin', '0.0%'],
  ]
  ws.getCell('B5').value = 'Your inputs'
  ws.getCell('B5').font = { name: FONT, size: 12, bold: true }
  rows.forEach(([label, k, fmt], i) => {
    const r = 6 + i
    ws.getCell(`B${r}`).value = label
    ws.getCell(`B${r}`).font = { name: FONT, size: 10 }
    const c = ws.getCell(`C${r}`)
    c.value = val(k)
    c.numFmt = fmt
    c.font = { name: FONT, size: 11, bold: true, color: { argb: C.blue } }
    c.fill = fill(C.inputFill)
    c.border = box
    const h = ws.getCell(`D${r}`)
    h.value = HINTS[k]
    h.font = { name: FONT, size: 9, italic: true, color: { argb: C.grey } }
    if (fmt === '0.0%') {
      c.dataValidation = { type: 'decimal', operator: 'between', formulae: [0.0001, 1], allowBlank: false, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Enter a percentage', error: 'Type 8% (or 0.08), not 8.' }
    }
  })
  ws.getCell('B14').value = 'Green cells with blue text are inputs. Everything else calculates.'
  ws.getCell('B14').font = { name: FONT, size: 9, italic: true, color: { argb: C.grey } }

  // Calculated values for cached results (so previews show numbers before Excel recalculates)
  const parsed = validate(Object.fromEntries(rows.map(([, k]) => [k, String(['load', 'cart', 'checkout', 'order', 'margin'].includes(k) ? (val(k) as number) * 100 : val(k))])) as Inputs).parsed!
  const res = compute(parsed)
  const lev = levers(parsed)

  // Funnel table
  const H = 16
  ws.getCell(`B${H}`).value = 'What each step costs'
  ws.getCell(`B${H}`).font = { name: FONT, size: 12, bold: true }
  const heads = ['Step', 'Rate', 'Reached', 'Dropped', 'Ad cost each ($)']
  heads.forEach((h, i) => {
    const c = ws.getCell(H + 1, 2 + i)
    c.value = h
    c.font = { name: FONT, size: 10, bold: true, color: { argb: C.white } }
    c.fill = fill(i === 0 ? C.green : C.dark2)
    c.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' }
    c.border = box
  })
  const first = H + 2 // clicks row
  STAGES.forEach((name, i) => {
    const r = first + i
    ws.getCell(`B${r}`).value = name
    const rate = ws.getCell(`C${r}`)
    if (i > 0) f(rate, `C${7 + i}`, [parsed.rates.load, parsed.rates.cart, parsed.rates.checkout, parsed.rates.order][i - 1])
    rate.numFmt = '0.0%'
    const reached = ws.getCell(`D${r}`)
    f(reached, i === 0 ? 'ROUND(C6,0)' : `ROUND(D${r - 1}*C${r},0)`, res.stages[i].reached)
    reached.numFmt = '#,##0'
    const dropped = ws.getCell(`E${r}`)
    if (i > 0) f(dropped, `D${r - 1}-D${r}`, res.stages[i].dropped!)
    dropped.numFmt = '#,##0'
    const cost = ws.getCell(`F${r}`)
    f(cost, i === 0 ? 'C7' : `IF(D${r}=0,"",C6*C7/D${r})`, i === 0 ? parsed.cpc : (res.stages[i].costPer ?? ''))
    cost.numFmt = '$#,##0.00'
    for (const col of 'BCDEF') {
      const c = ws.getCell(`${col}${r}`)
      c.border = box
      c.fill = fill(C.outFill)
      c.font = { name: FONT, size: 10, bold: col === 'F' || i === 4 }
    }
  })
  const last = first + 4 // orders row

  // Summary
  const S = last + 2
  const sum: [string, string, number | string | null, string][] = [
    ['Ad spend', 'C6*C7', res.spend, '$#,##0'],
    ['Orders', `D${last}`, res.orders, '#,##0'],
    ['Cost per order (CPA)', `IF(D${last}=0,"",C6*C7/D${last})`, res.cpa, '$#,##0.00'],
    ['Revenue', `D${last}*C12`, res.revenue, '$#,##0'],
    ['First-order ROAS', `IF(C6*C7=0,"",D${last}*C12/(C6*C7))`, res.roas, '0.00"x"'],
    ['Left per order before ads', 'C12*C13', res.contributionPerOrder, '$#,##0.00'],
    ['Left after ads, before opex', `D${last}*C12*C13-C6*C7`, res.afterAds, '+$#,##0;-$#,##0'],
  ]
  sum.forEach(([label, formula, result, fmt], i) => {
    const r = S + i
    ws.getCell(`B${r}`).value = label
    ws.getCell(`B${r}`).font = { name: FONT, size: 10, bold: i === 2 || i === 6 }
    const c = ws.getCell(`C${r}`)
    f(c, formula, result ?? undefined)
    c.numFmt = fmt
    c.font = { name: FONT, size: 11, bold: true }
    c.fill = fill(C.outFill)
    c.border = box
  })
  ws.addConditionalFormatting({
    ref: `C${S + 6}`,
    rules: [{ type: 'cellIs', operator: 'lessThan', formulae: [0], priority: 1, style: { font: { color: { argb: C.red }, bold: true } } }],
  })

  // Levers: +1 point on each step
  const L = S + sum.length + 1
  ws.getCell(`B${L}`).value = 'Which step to lift first: +1 point on each step'
  ws.getCell(`B${L}`).font = { name: FONT, size: 12, bold: true }
  ;['Lever', 'New rate', 'Orders', 'New CPA ($)', 'More after ads ($)'].forEach((h, i) => {
    const c = ws.getCell(L + 1, 2 + i)
    c.value = h
    c.font = { name: FONT, size: 10, bold: true, color: { argb: C.white } }
    c.fill = fill(i === 0 ? C.green : C.dark2)
    c.alignment = { horizontal: i === 0 ? 'left' : 'center' }
    c.border = box
  })
  // Orders with one rate raised by a point: recompute the whole chain with ROUND at each step.
  const chain = (bump: number) => {
    const rate = (j: number) => (j === bump ? `MIN(1,C${8 + j}+0.01)` : `C${8 + j}`)
    let expr = 'ROUND(C6,0)'
    for (let j = 0; j < 4; j++) expr = `ROUND(${expr}*${rate(j)},0)`
    return expr
  }
  const names = ['Click to page loaded', 'View to cart', 'Cart to checkout', 'Checkout to order']
  names.forEach((n, j) => {
    const r = L + 2 + j
    ws.getCell(`B${r}`).value = n
    f(ws.getCell(`C${r}`), `MIN(1,C${8 + j}+0.01)`, Math.min(1, [parsed.rates.load, parsed.rates.cart, parsed.rates.checkout, parsed.rates.order][j] + 0.01))
    ws.getCell(`C${r}`).numFmt = '0.0%'
    f(ws.getCell(`D${r}`), chain(j), lev[j].orders)
    ws.getCell(`D${r}`).numFmt = '#,##0'
    f(ws.getCell(`E${r}`), `IF(D${r}=0,"",C6*C7/D${r})`, lev[j].cpa ?? '')
    ws.getCell(`E${r}`).numFmt = '$#,##0.00'
    f(ws.getCell(`F${r}`), `(D${r}-D${last})*C12*C13`, lev[j].extraAfterAds ?? '')
    ws.getCell(`F${r}`).numFmt = '+$#,##0;-$#,##0'
  })
  const rc = L + 6
  ws.getCell(`B${rc}`).value = 'Clicks 10% cheaper'
  f(ws.getCell(`C${rc}`), 'C7*0.9', parsed.cpc * 0.9)
  ws.getCell(`C${rc}`).numFmt = '$#,##0.00'
  f(ws.getCell(`D${rc}`), `D${last}`, res.orders)
  ws.getCell(`D${rc}`).numFmt = '#,##0'
  f(ws.getCell(`E${rc}`), `IF(D${last}=0,"",C6*C7*0.9/D${last})`, lev[4].cpa ?? '')
  ws.getCell(`E${rc}`).numFmt = '$#,##0.00'
  f(ws.getCell(`F${rc}`), 'C6*C7*0.1', lev[4].extraAfterAds ?? '')
  ws.getCell(`F${rc}`).numFmt = '+$#,##0;-$#,##0'
  for (let r = L + 2; r <= rc; r++) for (const col of 'BCDEF') {
    const c = ws.getCell(`${col}${r}`)
    c.border = box
    c.fill = fill(C.outFill)
    if (!c.font?.bold) c.font = { name: FONT, size: 10, bold: col === 'F' }
  }

  // Notes and CTA
  let r = rc + 2
  ws.getCell(`B${r}`).value = 'How the numbers work'
  ws.getCell(`B${r}`).font = { name: FONT, size: 12, bold: true }
  for (const n of [
    'People at each step = people at the previous step × the step rate, rounded to whole people.',
    'Ad cost each = ad spend ÷ people who reached the step.',
    '+10% on any step gives about the same +10% orders. +1 point compares steps you can realistically move.',
    'Left after ads = orders × (order value × margin) minus ad spend. Before opex, so it is not net profit.',
    'First orders only. Repeat purchases and subscriptions change the picture: judge them by cohort payback days.',
  ]) {
    r++
    ws.getCell(`B${r}`).value = n
    ws.getCell(`B${r}`).font = { name: FONT, size: 10, color: { argb: 'FF3A3A3A' } }
  }
  r += 2
  for (const rr of [r, r + 1, r + 2]) for (let c = 1; c <= 7; c++) ws.getCell(rr, c).fill = fill(C.inputFill)
  ws.getRow(r).height = 10; ws.getRow(r + 2).height = 10; ws.getRow(r + 1).height = 22
  ws.getCell(`B${r + 1}`).value = 'Not sure which step is cheapest to lift in your store?'
  ws.getCell(`B${r + 1}`).font = { name: FONT, size: 11, bold: true, color: { argb: C.greenDeep } }
  const cta = ws.getCell(`F${r + 1}`)
  cta.value = { text: 'Send me your numbers', hyperlink: MESSAGE_URL }
  cta.font = { name: FONT, size: 11, bold: true, color: { argb: C.greenText }, underline: true }
  ws.headerFooter.oddFooter = '&L&8K&K Media, Funnel Cost Calculator&R&8linkedin.com/in/artem-kalachev'

  const buf = await wb.xlsx.writeBuffer()
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}
