// Pure funnel math. People at each step are whole numbers (rounded), so the calculator
// matches the post: 8.8% view to cart gives 224 orders and a $66.96 CPA.

export type Key = 'clicks' | 'cpc' | 'load' | 'cart' | 'checkout' | 'order' | 'aov' | 'margin'

export type Inputs = Record<Key, string>

export const PERCENT_KEYS: Key[] = ['load', 'cart', 'checkout', 'order', 'margin']
export const MONEY_KEYS: Key[] = ['cpc', 'aov']

export const LABELS: Record<Key, string> = {
  clicks: 'Clicks',
  cpc: 'Cost per click',
  load: 'Click to page loaded',
  cart: 'Page view to add to cart',
  checkout: 'Cart to checkout',
  order: 'Checkout to order',
  aov: 'Average order value',
  margin: 'Margin after COGS, shipping and fees',
}

export const EXAMPLE: Inputs = {
  clicks: '10000',
  cpc: '1.50',
  load: '85',
  cart: '8',
  checkout: '50',
  order: '60',
  aov: '120',
  margin: '55',
}

export const STEP_KEYS = ['load', 'cart', 'checkout', 'order'] as const
export type StepKey = (typeof STEP_KEYS)[number]

export const STAGES = ['Clicks', 'Page views', 'Carts', 'Checkouts', 'Orders'] as const

export interface Parsed {
  clicks: number
  cpc: number
  rates: Record<StepKey, number> // fractions
  aov: number | null
  margin: number | null // fraction
}

export interface Stage {
  name: (typeof STAGES)[number]
  reached: number
  costPer: number | null // ad spend ÷ reached; null for clicks (that is CPC)
  dropped: number | null // left between the previous stage and this one
}

export interface Result {
  spend: number
  stages: Stage[]
  orders: number
  cpa: number | null
  revenue: number | null
  roas: number | null
  contributionPerOrder: number | null
  contribution: number | null
  afterAds: number | null
}

export interface Lever {
  label: string
  detail: string
  orders: number
  cpa: number | null
  extraOrders: number
  extraAfterAds: number | null
}

export function parseNum(raw: string): number | null {
  const s = raw.trim().replace(/[$\s%,]/g, '')
  if (s === '') return null
  if (!/^\d*\.?\d+$|^\d+\.$/.test(s)) return NaN
  return Number(s)
}

export function validate(inp: Inputs): { parsed?: Parsed; errors: Partial<Record<Key, string>> } {
  const errors: Partial<Record<Key, string>> = {}
  const v: Partial<Record<Key, number | null>> = {}
  for (const k of Object.keys(LABELS) as Key[]) {
    const n = parseNum(inp[k])
    const optional = k === 'aov' || k === 'margin'
    if (n === null) {
      if (!optional) errors[k] = 'Required'
      v[k] = null
    } else if (Number.isNaN(n)) errors[k] = PERCENT_KEYS.includes(k) ? 'Use a number, for example 8 for 8%' : 'Use a number'
    else if (PERCENT_KEYS.includes(k) && (n <= 0 || n > 100)) errors[k] = 'Enter a percent above 0 and up to 100'
    else if (k === 'clicks' && (n < 1 || !Number.isInteger(n))) errors[k] = 'Enter whole clicks, at least 1'
    else if (MONEY_KEYS.includes(k) && n <= 0) errors[k] = 'Must be above 0'
    else v[k] = PERCENT_KEYS.includes(k) ? n / 100 : n
  }
  if (Object.keys(errors).length) return { errors }
  return {
    errors,
    parsed: {
      clicks: v.clicks as number,
      cpc: v.cpc as number,
      rates: { load: v.load as number, cart: v.cart as number, checkout: v.checkout as number, order: v.order as number },
      aov: (v.aov as number | null) ?? null,
      margin: (v.margin as number | null) ?? null,
    },
  }
}

/** People reached at each stage, rounded to whole people at every step. */
export function reach(clicks: number, rates: Record<StepKey, number>): number[] {
  const out = [Math.round(clicks)]
  for (const k of STEP_KEYS) out.push(Math.round(out[out.length - 1] * rates[k]))
  return out
}

export function compute(p: Parsed): Result {
  const spend = p.clicks * p.cpc
  const r = reach(p.clicks, p.rates)
  const stages: Stage[] = r.map((reached, i) => ({
    name: STAGES[i],
    reached,
    costPer: i === 0 ? null : reached > 0 ? spend / reached : null,
    dropped: i === 0 ? null : r[i - 1] - reached,
  }))
  const orders = r[4]
  const cpa = orders > 0 ? spend / orders : null
  const revenue = p.aov !== null ? orders * p.aov : null
  const roas = revenue !== null ? revenue / spend : null
  const contributionPerOrder = p.aov !== null && p.margin !== null ? p.aov * p.margin : null
  const contribution = contributionPerOrder !== null ? orders * contributionPerOrder : null
  const afterAds = contribution !== null ? contribution - spend : null
  return { spend, stages, orders, cpa, revenue, roas, contributionPerOrder, contribution, afterAds }
}

const STEP_LABEL: Record<StepKey, string> = {
  load: 'Click to page loaded',
  cart: 'View to cart',
  checkout: 'Cart to checkout',
  order: 'Checkout to order',
}

/** What +1 percentage point on each step is worth, plus 10% cheaper clicks for comparison. */
export function levers(p: Parsed): Lever[] {
  const base = compute(p)
  const per = base.contributionPerOrder
  const out: Lever[] = STEP_KEYS.map((k) => {
    const before = p.rates[k]
    const after = Math.min(1, before + 0.01)
    const res = compute({ ...p, rates: { ...p.rates, [k]: after } })
    const extra = res.orders - base.orders
    return {
      label: STEP_LABEL[k],
      detail: `${fmtPctPoint(before)} to ${fmtPctPoint(after)}`,
      orders: res.orders,
      cpa: res.cpa,
      extraOrders: extra,
      extraAfterAds: per !== null ? extra * per : null,
    }
  })
  const cheaper = compute({ ...p, cpc: p.cpc * 0.9 })
  out.push({
    label: 'Clicks 10% cheaper',
    detail: `${fmtMoney(p.cpc)} to ${fmtMoney(p.cpc * 0.9)} per click`,
    orders: cheaper.orders,
    cpa: cheaper.cpa,
    extraOrders: 0,
    extraAfterAds: base.spend - cheaper.spend,
  })
  return out
}

export const fmtMoney = (v: number, sign = false) => {
  const abs = Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  if (v < 0) return `-$${abs}`
  return sign ? `+$${abs}` : `$${abs}`
}
export const fmtMoney0 = (v: number, sign = false) => {
  const abs = Math.round(Math.abs(v)).toLocaleString('en-US')
  if (Math.round(v) < 0) return `-$${abs}`
  return sign && Math.round(v) > 0 ? `+$${abs}` : `$${abs}`
}
export const fmtInt = (v: number) => Math.round(v).toLocaleString('en-US')
export const fmtPctPoint = (f: number) => `${+(f * 100).toFixed(2)}%`
export const fmtRoas = (v: number) => `${v.toFixed(2)}x`
