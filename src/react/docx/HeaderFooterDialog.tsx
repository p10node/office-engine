import { useEffect, useRef, useState } from 'react'
import type { Block, DocxMeta, HFKind, HFLine, HFSet } from '../../core/docx/types'

const KINDS: { k: HFKind; label: string }[] = [
  { k: 'header', label: 'Header' },
  { k: 'footer', label: 'Footer' },
  { k: 'firstHeader', label: 'First page header' },
  { k: 'firstFooter', label: 'First page footer' },
]

/** quick page numbers: where, and how */
const NUMBERS: { label: string; k: HFKind; line: HFLine }[] = [
  { label: 'Bottom center: 1', k: 'footer', line: { text: '{PAGE}', align: 'center' } },
  { label: 'Bottom right: 1', k: 'footer', line: { text: '{PAGE}', align: 'right' } },
  { label: 'Bottom center: Page 1 of 5', k: 'footer', line: { text: 'Page {PAGE} of {NUMPAGES}', align: 'center' } },
  { label: 'Bottom center: Trang 1/5', k: 'footer', line: { text: 'Trang {PAGE}/{NUMPAGES}', align: 'center' } },
  { label: 'Top right: 1', k: 'header', line: { text: '{PAGE}', align: 'right' } },
]

interface Part {
  text: string
  align: 'left' | 'center' | 'right'
  b: boolean
  sz: string
  /** holds pictures, tables or shapes the dialog can't show: saving replaces them */
  rich: boolean
}

/** the editable form of a header / footer as converted for display */
export function partOf(blocks: Block[] | undefined): Part {
  const paras = (blocks ?? []).filter((b) => b.t === 'p')
  const rich = (blocks ?? []).some((b) => b.t !== 'p' || b.runs.some((r) => r.t !== 'text' && r.t !== 'field'))
  const lines = paras.map((p) => (p.t === 'p' ? p.runs.map((r) => (r.t === 'text' ? r.text.replace(/\n/g, ' ') : r.t === 'field' ? `{${r.kind}}` : '')).join('') : ''))
  const p0 = paras.find((p) => p.t === 'p' && p.runs.length) ?? paras[0]
  const runs = p0?.t === 'p' ? p0.runs.filter((r) => r.t === 'text' || r.t === 'field') : []
  const align = p0?.t === 'p' && (p0.ps.align === 'center' || p0.ps.align === 'right') ? p0.ps.align : 'left'
  const first = runs[0] as { s?: { b?: boolean; sz?: number } } | undefined
  // trailing empty lines carry nothing
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop()
  return { text: lines.join('\n'), align, b: runs.length > 0 && runs.every((r) => (r as { s?: { b?: boolean } }).s?.b), sz: first?.s?.sz ? String(first.s.sz) : '', rich }
}

const linesOf = (p: Part): HFLine[] =>
  p.text.split('\n').length === 1 && !p.text.trim()
    ? []
    : p.text.split('\n').map((text) => ({ text, align: p.align, ...(p.b ? { b: true } : {}), ...(Number(p.sz) >= 4 ? { sz: Number(p.sz) } : {}) }))

/**
 * Headers and footers of the document: text (a line per paragraph), alignment, bold, size, page number / page count,
 * a different first page. `onOk` gets only the parts that changed.
 */
export function HeaderFooterDialog({ meta, focus, onOk, onClose }: { meta: DocxMeta; focus?: HFKind; onOk: (set: HFSet) => void; onClose: () => void }) {
  const init = useRef(Object.fromEntries(KINDS.map(({ k }) => [k, partOf(meta[k])])) as Record<HFKind, Part>)
  const [parts, setParts] = useState(init.current)
  const [first, setFirst] = useState(!!meta.titlePg)
  const [active, setActive] = useState<HFKind>(focus && (first || !focus.startsWith('first')) ? focus : 'header')
  const area = useRef<HTMLTextAreaElement>(null)
  useEffect(() => area.current?.focus(), [active])

  const cur = parts[active]
  const put = (patch: Partial<Part>, k: HFKind = active) => setParts((p) => ({ ...p, [k]: { ...p[k], ...patch } }))
  const insert = (token: string) => {
    const el = area.current
    const at = el ? el.selectionStart : cur.text.length
    const end = el ? el.selectionEnd : at
    put({ text: cur.text.slice(0, at) + token + cur.text.slice(end) })
    requestAnimationFrame(() => {
      el?.focus()
      el?.setSelectionRange(at + token.length, at + token.length)
    })
  }

  const submit = () => {
    const set: HFSet = {}
    for (const { k } of KINDS) {
      const a = parts[k]
      const b = init.current[k]
      if (a.text !== b.text || a.align !== b.align || a.b !== b.b || a.sz !== b.sz) (set.parts ??= {})[k] = linesOf(a)
    }
    if (first !== !!meta.titlePg) set.titlePg = first
    onClose()
    if (set.parts || set.titlePg !== undefined) onOk(set)
  }

  const shown = KINDS.filter(({ k }) => first || !k.startsWith('first'))
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
        className="de-dialog de-hfd"
        role="dialog"
        aria-label="Header and footer"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          submit()
        }}
      >
        <div className="de-dialog-title">Header &amp; footer</div>
        <div className="de-hfd-tabs" role="tablist">
          {shown.map(({ k, label }) => (
            <button key={k} type="button" role="tab" aria-selected={active === k} className={active === k ? 'on' : undefined} onClick={() => setActive(k)}>
              {label}
            </button>
          ))}
        </div>
        <textarea
          ref={area}
          rows={3}
          aria-label={`${KINDS.find((x) => x.k === active)!.label} text`}
          placeholder="Type the text; one line per paragraph"
          value={cur.text}
          style={{ textAlign: cur.align, fontWeight: cur.b ? 700 : undefined }}
          onChange={(e) => put({ text: e.target.value })}
        />
        {cur.rich && <div className="de-dialog-note de-ps-bad">This part also holds pictures, tables or shapes: changing it here removes them.</div>}
        <div className="de-ps-row">
          {(['left', 'center', 'right'] as const).map((a) => (
            <label key={a} className="de-ps-radio">
              <input type="radio" name="de-hfd-align" checked={cur.align === a} onChange={() => put({ align: a })} /> {a[0].toUpperCase() + a.slice(1)}
            </label>
          ))}
          <label className="de-ps-radio">
            <input type="checkbox" checked={cur.b} onChange={(e) => put({ b: e.target.checked })} /> Bold
          </label>
          <label className="de-ps-field">
            <span>Size</span>
            <input type="number" step="any" min="4" aria-label="Font size (pt)" placeholder="auto" value={cur.sz} onChange={(e) => put({ sz: e.target.value })} />
          </label>
        </div>
        <div className="de-ps-row">
          <button type="button" className="de-btn" onClick={() => insert('{PAGE}')}>
            + Page number
          </button>
          <button type="button" className="de-btn" onClick={() => insert('{NUMPAGES}')}>
            + Page count
          </button>
          <select
            className="de-ps-preset"
            aria-label="Page number preset"
            value=""
            onChange={(e) => {
              const n = NUMBERS[Number(e.target.value)]
              if (!n) return
              put({ text: n.line.text, align: n.line.align ?? 'left' }, n.k)
              setActive(n.k)
            }}
          >
            <option value="">Page numbers…</option>
            {NUMBERS.map((n, i) => (
              <option key={n.label} value={i}>
                {n.label}
              </option>
            ))}
          </select>
        </div>
        <label className="de-ps-radio">
          <input
            type="checkbox"
            checked={first}
            onChange={(e) => {
              setFirst(e.target.checked)
              if (!e.target.checked && active.startsWith('first')) setActive(active === 'firstFooter' ? 'footer' : 'header')
            }}
          />{' '}
          Different first page
        </label>
        <div className="de-dialog-btns">
          <button type="button" className="de-btn" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className="de-btn de-btn-primary">
            Apply
          </button>
        </div>
      </form>
    </div>
  )
}
