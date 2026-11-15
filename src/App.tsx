import { useEffect, useMemo, useState } from 'react'
import {
  EXAMPLE,
  HINTS,
  LABELS,
  MONEY_KEYS,
  PERCENT_KEYS,
  STEP_KEYS,
  compute,
  fmtInt,
  fmtMoney,
  fmtMoney0,
  fmtRoas,
  levers,
  validate,
  type Inputs,
  type Key,
} from './calc'

const LINKEDIN = 'https://www.linkedin.com/in/artem-kalachev/'
// Opens a new LinkedIn message to Artem (same link as the Message button on his profile)
const LINKEDIN_MESSAGE =
  'https://www.linkedin.com/messaging/compose/?profileUrn=urn%3Ali%3Afsd_profile%3AACoAADxgieABwE7747g9IZ4BfTrBChanNiZJwr4'
const LOGO = `${import.meta.env.BASE_URL}kk-logo-96.png`
const STORE = 'kk-funnel-calculator:v1'
const ROAS_TOOL = 'https://kk-arti.github.io/kk-roas-calculator/'

function load(): Inputs {
  try {
    const raw = window.localStorage.getItem(STORE)
    if (!raw) return { ...EXAMPLE }
    const s = JSON.parse(raw) as Inputs
    return { ...EXAMPLE, ...s }
  } catch {
    return { ...EXAMPLE }
  }
}

const GROUPS: { title: string; hint: string; keys: Key[] }[] = [
  { title: 'Ads', hint: 'One channel, one period, for example last month.', keys: ['clicks', 'cpc'] },
  {
    title: 'Step rates',
    hint: "Share of people who make it to the next step, for this channel's traffic only. In GA4 the same steps are in Reports, Monetization, Purchase journey, filtered by session source.",
    keys: ['load', 'cart', 'checkout', 'order'],
  },
  { title: 'Order economics', hint: 'Optional. Adds revenue, ROAS and money left after ads.', keys: ['aov', 'margin'] },
]

export default function App() {
  const [inp, setInp] = useState<Inputs>(load)
  useEffect(() => {
    try {
      window.localStorage.setItem(STORE, JSON.stringify(inp))
    } catch {
      // storage blocked: the calculator still works for this visit
    }
  }, [inp])

  const v = useMemo(() => validate(inp), [inp])
  const res = v.parsed ? compute(v.parsed) : null
  const lev = v.parsed ? levers(v.parsed) : null
  const isExample = (Object.keys(EXAMPLE) as Key[]).every((k) => inp[k] === EXAMPLE[k])
  const set = (k: Key, val: string) => setInp((s) => ({ ...s, [k]: val }))

  const [busy, setBusy] = useState(false)
  const downloadExcel = async () => {
    setBusy(true)
    try {
      const { buildWorkbook } = await import('./excel')
      const blob = await buildWorkbook(inp)
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'KK Funnel Cost Calculator.xlsx'
      document.body.appendChild(a)
      a.click()
      a.remove()
      setTimeout(() => URL.revokeObjectURL(url), 1000)
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <div className="screen">
        <header className="site-header">
          <div className="wrap header-row">
            <div className="brand">
              <img src={LOGO} alt="" width="36" height="36" />
              <span>K&amp;K Media</span>
            </div>
            <a className="header-link" href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
              Message me
            </a>
          </div>
        </header>

        <main className="wrap">
          <section className="hero">
            <p className="eyebrow">Free tool for DTC brands</p>
            <h1>Funnel cost calculator</h1>
            <p className="lead">
              See what each step from click to order costs you in ad dollars, and which step is worth lifting first.
            </p>
          </section>

          <section className="card" aria-labelledby="inputs-h">
            <div className="section-head">
              <h2 id="inputs-h">Your funnel</h2>
              <div className="toolbar">
                {isExample ? (
                  <button type="button" className="btn ghost" onClick={() => setInp({ clicks: '', cpc: '', load: '', cart: '', checkout: '', order: '', aov: '', margin: '' })}>
                    Clear example
                  </button>
                ) : (
                  <button type="button" className="btn ghost" onClick={() => setInp({ ...EXAMPLE })}>
                    Load example
                  </button>
                )}
                <button type="button" className="btn ghost" onClick={downloadExcel} disabled={busy}>
                  {busy ? 'Preparing file…' : 'Download Excel'}
                </button>
                <button type="button" className="btn ghost" onClick={() => window.print()}>
                  Print / Save as PDF
                </button>
              </div>
            </div>
            <p className="one-channel">Count one ad channel at a time: clicks and spend from that channel's ad account, step rates for that channel's traffic only.</p>
            {isExample && <p className="ex-note">Example numbers. Replace them with yours.</p>}
            <p className="help-note">
              Stuck on a number or not sure where to find it?{' '}
              <a href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
                Message me
              </a>{' '}
              and I'll help.
            </p>
            <div className="groups">
              {GROUPS.map((g) => (
                <fieldset key={g.title} className="group">
                  <legend>{g.title}</legend>
                  <p className="group-hint">{g.hint}</p>
                  <div className="fields">
                    {g.keys.map((k) => (
                      <Field key={k} k={k} value={inp[k]} error={v.errors[k]} onChange={set} />
                    ))}
                  </div>
                </fieldset>
              ))}
            </div>
          </section>

          {res && lev ? (
            <>
              <section className="card" aria-labelledby="sum-h">
                <h2 id="sum-h">The result</h2>
                <div className="tiles">
                  <Tile label="Ad spend" value={fmtMoney0(res.spend)} />
                  <Tile label="Orders" value={fmtInt(res.orders)} />
                  <Tile label="Cost per order (CPA)" value={res.cpa !== null ? fmtMoney(res.cpa) : 'no orders'} strong />
                  {res.roas !== null && <Tile label="First-order ROAS" value={fmtRoas(res.roas)} />}
                  {res.afterAds !== null && (
                    <Tile
                      label="Left after ads, before opex"
                      value={fmtMoney0(res.afterAds, true)}
                      tone={res.afterAds >= 0 ? 'pos' : 'neg'}
                    />
                  )}
                </div>
                {res.contributionPerOrder !== null && res.cpa !== null && (
                  <p className="verdict">
                    Each order leaves {fmtMoney(res.contributionPerOrder)} before ads and costs {fmtMoney(res.cpa)} to get.{' '}
                    {res.cpa <= res.contributionPerOrder ? (
                      <strong className="pos">The first order covers its ad cost.</strong>
                    ) : (
                      <strong className="neg">
                        The first order loses {fmtMoney(res.cpa - res.contributionPerOrder)} after ads.
                      </strong>
                    )}
                  </p>
                )}
              </section>

              <section className="card" aria-labelledby="steps-h">
                <h2 id="steps-h">What each step costs</h2>
                <p className="muted">Ad spend divided by the people who reached the step.</p>
                <div className="steps">
                  {res.stages.map((s, i) => {
                    const max = res.stages[res.stages.length - 1].costPer ?? 1
                    const w = s.costPer !== null ? Math.max(2, (s.costPer / max) * 100) : 0
                    return (
                      <div className="step-row" key={s.name}>
                        <div className="step-name">
                          {s.name}
                          <span className="unit">
                            {fmtInt(s.reached)} {i === 0 ? 'people' : 'reached'}
                            {s.dropped !== null && s.dropped > 0 ? `, ${fmtInt(s.dropped)} dropped` : ''}
                          </span>
                        </div>
                        <div className="step-track" aria-hidden="true">
                          {s.costPer !== null && <i style={{ width: `${w}%` }} />}
                        </div>
                        <div className="step-cost">
                          {s.costPer !== null ? fmtMoney(s.costPer) : fmtMoney(v.parsed!.cpc)}
                          <span className="unit">{i === 0 ? 'per click' : 'each'}</span>
                        </div>
                      </div>
                    )
                  })}
                </div>
              </section>

              <section className="card" aria-labelledby="lev-h">
                <h2 id="lev-h">Which step to lift first</h2>
                <p className="muted">
                  +10% on any step gives about the same +10% orders, because the steps multiply. So compare what you can
                  realistically move: one percentage point on each step.
                </p>
                <Levers levers={lev} hasMoney={res.contributionPerOrder !== null} />
              </section>

              <div className="mix-cta">
                <div className="mix-body">
                  <div>
                    <p className="mix-title">Several channels, subscriptions or repeat orders?</p>
                    <p className="muted">
                      That's where the real answer is, and it's different for every store. Send me your numbers and I'll
                      tell you which step I would lift first.
                    </p>
                  </div>
                  <a className="btn primary" href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
                    Message me
                  </a>
                </div>
              </div>
            </>
          ) : (
            <section className="card">
              <p className="muted">Fill in clicks, cost per click and the four step rates to see the result.</p>
            </section>
          )}

          <Notes />
        </main>

        <footer className="site-footer">
          <div className="wrap footer-inner">
            <div className="footer-cta">
              <p className="footer-title">Want a second pair of eyes on your funnel?</p>
              <a className="btn primary big" href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
                Message me on LinkedIn
              </a>
            </div>
            <div className="footer-meta">
              <img src={LOGO} alt="" width="28" height="28" />
              <span>
                Built by{' '}
                <a href={LINKEDIN} target="_blank" rel="noopener noreferrer">
                  Artem Kalachev
                </a>
                , K&amp;K Media
              </span>
            </div>
            <p className="privacy small muted">No tracking, no cookies. Your numbers stay in this browser.</p>
          </div>
        </footer>
      </div>

      {res && lev && <PrintView inp={inp} res={res} lev={lev} />}
    </>
  )
}

function Field({ k, value, error, onChange }: { k: Key; value: string; error?: string; onChange: (k: Key, v: string) => void }) {
  const isPct = PERCENT_KEYS.includes(k)
  const isMoney = MONEY_KEYS.includes(k)
  const id = `f-${k}`
  return (
    <div className="field">
      <label htmlFor={id}>{LABELS[k]}</label>
      <div className={`num${error ? ' invalid' : ''}`}>
        {isMoney && <span className="adorn pre">$</span>}
        <input
          id={id}
          inputMode="decimal"
          autoComplete="off"
          value={value}
          placeholder="0"
          onChange={(e) => onChange(k, e.target.value)}
          aria-invalid={!!error}
          aria-describedby={error ? `${id}-err` : `${id}-hint`}
          className={isPct ? 'has-post' : isMoney ? 'has-pre' : ''}
        />
        {isPct && <span className="adorn post">%</span>}
      </div>
      {error ? (
        <span id={`${id}-err`} className="field-error" role="alert">
          {error}
        </span>
      ) : (
        <span id={`${id}-hint`} className="field-hint">
          {HINTS[k]}
          {k === 'margin' && (
            <>
              {' '}
              <a href={ROAS_TOOL} target="_blank" rel="noopener noreferrer">
                Work it out per product
              </a>
            </>
          )}
        </span>
      )}
    </div>
  )
}

function Tile({ label, value, strong, tone }: { label: string; value: string; strong?: boolean; tone?: 'pos' | 'neg' }) {
  return (
    <div className={`tile${strong ? ' strong' : ''}`}>
      <span className="metric-label">{label}</span>
      <span className={`tile-value ${tone ?? ''}`}>{value}</span>
    </div>
  )
}

function Levers({ levers: l, hasMoney }: { levers: ReturnType<typeof levers>; hasMoney: boolean }) {
  const steps = l.slice(0, STEP_KEYS.length)
  const best = steps.reduce((a, b) => (b.extraOrders > a.extraOrders ? b : a))
  return (
    <div className="lev">
      {l.map((x) => (
        <div className={`lev-row${hasMoney ? ' has-money' : ''}${x === best ? ' best' : ''}`} key={x.label}>
          <div className="lev-name">
            {x.label}
            <span className="unit">{x.detail}</span>
          </div>
          <div className="lev-num">
            {x.extraOrders > 0 ? `+${fmtInt(x.extraOrders)}` : 'same'}
            <span className="unit">orders</span>
          </div>
          <div className="lev-num">
            {x.cpa !== null ? fmtMoney(x.cpa) : ''}
            <span className="unit">new CPA</span>
          </div>
          {hasMoney && (
            <div className="lev-num money">
              {x.extraAfterAds !== null ? fmtMoney0(x.extraAfterAds, true) : ''}
              <span className="unit">more after ads</span>
            </div>
          )}
          {x === best && <span className="pill up">Biggest step lever</span>}
        </div>
      ))}
    </div>
  )
}

const NOTES: [string, string][] = [
  ['Ad spend', 'Clicks × cost per click.'],
  ['People at each step', 'People at the previous step × the step rate, rounded to whole people.'],
  ['Cost per step', 'Ad spend ÷ people who reached that step. A lost checkout costs more than a lost cart because fewer people get that far on the same spend.'],
  ['CPA', 'Ad spend ÷ orders.'],
  ['Left after ads', 'Orders × (average order value × margin) minus ad spend. Before opex, so it is not net profit.'],
  ['+1 point', 'Adds one percentage point to one step and keeps the rest. Clicks 10% cheaper keeps every rate and cuts ad spend instead.'],
]

function Notes() {
  return (
    <section className="card notes" aria-labelledby="notes-h">
      <h2 id="notes-h">How the numbers work</h2>
      <dl>
        {NOTES.map(([t, d]) => (
          <div key={t}>
            <dt>{t}</dt>
            <dd>{d}</dd>
          </div>
        ))}
      </dl>
      <p className="callout">
        First orders only. Repeat purchases and subscriptions change the picture: judge them by cohort payback days.
      </p>
    </section>
  )
}

function PrintView({ inp, res, lev }: { inp: Inputs; res: ReturnType<typeof compute>; lev: ReturnType<typeof levers> }) {
  const date = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })
  return (
    <div className="print-only" aria-hidden="true">
      <div className="pv-head">
        <img src={LOGO} alt="" width="40" height="40" />
        <div>
          <h1>Funnel cost calculator</h1>
          <p>
            K&amp;K Media. {date}. {fmtInt(Number(inp.clicks))} clicks at ${inp.cpc}, ad spend {fmtMoney0(res.spend)}
          </p>
        </div>
      </div>
      <table className="pv-table">
        <thead>
          <tr>
            <th>Step</th>
            <th>Rate</th>
            <th>Reached</th>
            <th>Dropped</th>
            <th>Ad cost each</th>
          </tr>
        </thead>
        <tbody>
          {res.stages.map((s, i) => (
            <tr key={s.name}>
              <td>{s.name}</td>
              <td>{i === 0 ? '' : `${inp[STEP_KEYS[i - 1]]}%`}</td>
              <td>{fmtInt(s.reached)}</td>
              <td>{s.dropped !== null ? fmtInt(s.dropped) : ''}</td>
              <td>{s.costPer !== null ? fmtMoney(s.costPer) : `${fmtMoney(Number(inp.cpc))} per click`}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <p>
        <strong>CPA:</strong> {res.cpa !== null ? fmtMoney(res.cpa) : 'no orders'}
        {res.roas !== null && (
          <>
            {'. '}
            <strong>First-order ROAS:</strong> {fmtRoas(res.roas)}
          </>
        )}
        {res.afterAds !== null && (
          <>
            {'. '}
            <strong>Left after ads, before opex:</strong> {fmtMoney0(res.afterAds, true)}
          </>
        )}
      </p>
      <h2>+1 point on each step</h2>
      <table className="pv-table">
        <thead>
          <tr>
            <th>Lever</th>
            <th>Change</th>
            <th>Extra orders</th>
            <th>New CPA</th>
            <th>More after ads</th>
          </tr>
        </thead>
        <tbody>
          {lev.map((x) => (
            <tr key={x.label}>
              <td>{x.label}</td>
              <td>{x.detail}</td>
              <td>{x.extraOrders > 0 ? `+${fmtInt(x.extraOrders)}` : 'same'}</td>
              <td>{x.cpa !== null ? fmtMoney(x.cpa) : ''}</td>
              <td>{x.extraAfterAds !== null ? fmtMoney0(x.extraAfterAds, true) : ''}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <h2>How the numbers work</h2>
      <ul>
        {NOTES.map(([t, d]) => (
          <li key={t}>
            <strong>{t}:</strong> {d}
          </li>
        ))}
      </ul>
      <p className="pv-foot">Built by Artem Kalachev, K&amp;K Media. linkedin.com/in/artem-kalachev</p>
    </div>
  )
}
