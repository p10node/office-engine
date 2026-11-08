import { useEffect, useRef, useState } from 'react'
import type { PageSet, PageSetup } from '../../core/docx/types'

/** paper sizes in pt (portrait) */
export const PAPERS: { k: string; label: string; w: number; h: number }[] = [
  { k: 'a4', label: 'A4 (21 × 29.7 cm)', w: 595.3, h: 841.9 },
  { k: 'letter', label: 'Letter (8.5 × 11 in)', w: 612, h: 792 },
  { k: 'legal', label: 'Legal (8.5 × 14 in)', w: 612, h: 1008 },
  { k: 'a3', label: 'A3 (29.7 × 42 cm)', w: 841.9, h: 1190.6 },
  { k: 'a5', label: 'A5 (14.8 × 21 cm)', w: 419.5, h: 595.3 },
  { k: 'b5', label: 'B5 (17.6 × 25 cm)', w: 498.9, h: 708.7 },
]

const CM = 72 / 2.54
const toCm = (pt: number) => +(pt / CM).toFixed(2)
const toPt = (cm: number) => +(cm * CM).toFixed(2)
const near = (a: number, b: number) => Math.abs(a - b) < 2

/** the paper of a page (either way round), or 'custom' */
export function paperOf(page: { w: number; h: number }): string {
  const lo = Math.min(page.w, page.h)
  const hi = Math.max(page.w, page.h)
  return PAPERS.find((p) => near(p.w, lo) && near(p.h, hi))?.k ?? 'custom'
}

const MARGIN_PRESETS: { label: string; cm: [number, number, number, number] }[] = [
  { label: 'Normal', cm: [2.54, 2.54, 2.54, 2.54] },
  { label: 'Narrow', cm: [1.27, 1.27, 1.27, 1.27] },
  { label: 'Moderate', cm: [2.54, 1.91, 2.54, 1.91] },
  { label: 'Wide', cm: [2.54, 5.08, 2.54, 5.08] },
  { label: 'Vietnam (TT 01/2011)', cm: [2, 1.5, 2, 3] },
]

/**
 * Page setup for the whole document: paper, orientation, margins (cm), text columns.
 * `onOk` gets only what changed.
 */
export function PageSetupDialog({ page, onOk, onClose }: { page: PageSetup; onOk: (set: PageSet) => void; onClose: () => void }) {
  const [paper, setPaper] = useState(paperOf(page))
  const [land, setLand] = useState(page.w > page.h)
  // margins as typed (cm): top, right, bottom, left
  const [mar, setMar] = useState([page.mt, page.mr, page.mb, page.ml].map((v) => String(toCm(v))))
  const [cols, setCols] = useState(page.cols ?? 1)
  const [gap, setGap] = useState(String(toCm(page.colGap ?? 36)))
  const first = useRef<HTMLSelectElement>(null)
  useEffect(() => first.current?.focus(), [])

  const nums = mar.map(Number)
  const gapN = Number(gap)
  const size = (() => {
    const p = PAPERS.find((x) => x.k === paper)
    const lo = p ? p.w : Math.min(page.w, page.h)
    const hi = p ? p.h : Math.max(page.w, page.h)
    return land ? { w: hi, h: lo } : { w: lo, h: hi }
  })()
  const bad =
    nums.some((v) => !Number.isFinite(v) || v < 0 || v > 15) ||
    toPt(nums[1] + nums[3]) > size.w - 72 ||
    toPt(nums[0] + nums[2]) > size.h - 72 ||
    !Number.isFinite(gapN) ||
    gapN < 0 ||
    gapN > 10

  const submit = () => {
    if (bad) return
    const set: PageSet = {}
    if (!near(size.w, page.w) || !near(size.h, page.h)) Object.assign(set, { w: size.w, h: size.h })
    const keys = ['mt', 'mr', 'mb', 'ml'] as const
    keys.forEach((k, i) => {
      const v = toPt(nums[i])
      if (Math.abs(v - page[k]) > 0.3) set[k] = v
    })
    if (cols !== (page.cols ?? 1)) set.cols = cols
    if (cols > 1 && Math.abs(toPt(gapN) - (page.colGap ?? 36)) > 0.3) set.colGap = toPt(gapN)
    onClose()
    if (Object.keys(set).length) onOk(set)
  }

  const field = (i: number, label: string) => (
    <label className="de-ps-field">
      <span>{label}</span>
      <input
        type="number"
        step="any"
        min="0"
        aria-label={`${label} margin (cm)`}
        value={mar[i]}
        onChange={(e) => setMar((m) => m.map((x, k) => (k === i ? e.target.value : x)))}
      />
      <em>cm</em>
    </label>
  )

  return (
    <div
      className="de-dialog-back"
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          onClose()
        }
      }}
    >
      <form
        className="de-dialog de-ps"
        role="dialog"
        noValidate
        aria-label="Page setup"
        onSubmit={(e) => {
          e.preventDefault()
          submit()
        }}
      >
        <div className="de-dialog-title">Page setup</div>
        <div className="de-ps-note">Applies to the whole document.</div>

        <div className="de-ps-row">
          <label className="de-ps-field de-ps-wide">
            <span>Paper size</span>
            <select ref={first} aria-label="Paper size" value={paper} onChange={(e) => setPaper(e.target.value)}>
              {PAPERS.map((p) => (
                <option key={p.k} value={p.k}>
                  {p.label}
                </option>
              ))}
              {paper === 'custom' && <option value="custom">Custom ({toCm(page.w)} × {toCm(page.h)} cm)</option>}
            </select>
          </label>
        </div>

        <div className="de-ps-row" role="radiogroup" aria-label="Orientation">
          <span className="de-ps-head">Orientation</span>
          <label className="de-ps-radio">
            <input type="radio" name="de-ps-orient" checked={!land} onChange={() => setLand(false)} /> Portrait
          </label>
          <label className="de-ps-radio">
            <input type="radio" name="de-ps-orient" checked={land} onChange={() => setLand(true)} /> Landscape
          </label>
        </div>

        <div className="de-ps-head">
          Margins
          <select
            className="de-ps-preset"
            aria-label="Margin preset"
            value=""
            onChange={(e) => {
              const p = MARGIN_PRESETS[Number(e.target.value)]
              if (p) setMar(p.cm.map(String))
            }}
          >
            <option value="">Presets…</option>
            {MARGIN_PRESETS.map((p, i) => (
              <option key={p.label} value={i}>
                {p.label}
              </option>
            ))}
          </select>
        </div>
        <div className="de-ps-grid">
          {field(0, 'Top')}
          {field(2, 'Bottom')}
          {field(3, 'Left')}
          {field(1, 'Right')}
        </div>

        <div className="de-ps-row">
          <span className="de-ps-head">Columns</span>
          {[1, 2, 3].map((n) => (
            <label key={n} className="de-ps-radio">
              <input type="radio" name="de-ps-cols" checked={cols === n} onChange={() => setCols(n)} /> {n}
            </label>
          ))}
          {cols > 1 && (
            <label className="de-ps-field">
              <span>Spacing</span>
              <input type="number" step="any" min="0" aria-label="Column spacing (cm)" value={gap} onChange={(e) => setGap(e.target.value)} />
              <em>cm</em>
            </label>
          )}
        </div>

        {bad && <div className="de-dialog-note de-ps-bad">The margins leave no room for text.</div>}
        <div className="de-dialog-btns">
          <button type="button" className="de-btn" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className="de-btn de-btn-primary" disabled={bad}>
            Apply
          </button>
        </div>
      </form>
    </div>
  )
}
