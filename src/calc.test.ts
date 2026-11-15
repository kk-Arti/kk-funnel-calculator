import { describe, expect, it } from 'vitest'
import { EXAMPLE, compute, levers, reach, validate } from './calc'

const parsed = () => validate(EXAMPLE).parsed!

describe('funnel math matches post #6', () => {
  it('reaches 8,500 / 680 / 340 / 204 from 10,000 clicks', () => {
    expect(reach(10000, parsed().rates)).toEqual([10000, 8500, 680, 340, 204])
  })
  it('costs $1.76 / $22.06 / $44.12 / $73.53 per step on $15,000', () => {
    const r = compute(parsed())
    expect(r.spend).toBe(15000)
    expect(r.stages.slice(1).map((s) => +s.costPer!.toFixed(2))).toEqual([1.76, 22.06, 44.12, 73.53])
    expect(r.stages.slice(1).map((s) => s.dropped)).toEqual([1500, 7820, 340, 136])
  })
  it('8.8% view to cart gives 224 orders and a $66.96 CPA', () => {
    const p = parsed()
    const r = compute({ ...p, rates: { ...p.rates, cart: 0.088 } })
    expect(r.orders).toBe(224)
    expect(r.cpa!.toFixed(2)).toBe('66.96')
  })
  it('example economics: $24,480 revenue, 1.63x, $66 per order, -$1,536 after ads', () => {
    const r = compute(parsed())
    expect(r.revenue).toBe(24480)
    expect(r.roas!.toFixed(2)).toBe('1.63')
    expect(r.contributionPerOrder).toBeCloseTo(66)
    expect(r.afterAds!).toBeCloseTo(-1536)
  })
  it('+1 point levers: view to cart is the biggest step lever in the example', () => {
    const l = levers(parsed())
    const steps = l.slice(0, 4)
    const best = steps.reduce((a, b) => ((b.extraAfterAds ?? 0) > (a.extraAfterAds ?? 0) ? b : a))
    expect(best.label).toBe('View to cart')
    expect(l[1].orders).toBe(230)
    expect(l[4].extraAfterAds).toBeCloseTo(1500)
  })
  it('rejects bad input', () => {
    expect(validate({ ...EXAMPLE, cart: '0' }).errors.cart).toBeTruthy()
    expect(validate({ ...EXAMPLE, clicks: '' }).errors.clicks).toBeTruthy()
    expect(validate({ ...EXAMPLE, aov: '' }).parsed!.aov).toBeNull()
  })
})
