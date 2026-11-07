import { memo, type CSSProperties, type ReactNode } from 'react'
import type { Block, Inline, Paragraph, ParaStyle, RunStyle, Table } from '../../core/docx/types'
import type { DocumentClient } from '../../core/worker/client'
import { useMediaUrl } from '../common/hooks'

export interface RenderCtx {
  client: DocumentClient
  /** 0-based page index, -1 while measuring */
  page: number
  pages: number
  measuring: boolean
  editable: boolean
}

const pt = (v: number | undefined) => (v ? `${+v.toFixed(2)}pt` : undefined)

const UNDERLINE_STYLE: Record<string, CSSProperties['textDecorationStyle']> = {
  double: 'double', dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashedHeavy: 'dashed', dashLong: 'dashed',
  wave: 'wavy', wavyHeavy: 'wavy', wavyDouble: 'wavy',
}

export function runCss(s: RunStyle | undefined): CSSProperties {
  const css: CSSProperties = {}
  if (!s) return css
  if (s.b) css.fontWeight = 700
  else if (s.b === false) css.fontWeight = 400
  if (s.i) css.fontStyle = 'italic'
  else if (s.i === false) css.fontStyle = 'normal'
  const u = s.u && s.u !== 'none'
  const deco = [u ? 'underline' : '', s.strike || s.dstrike ? 'line-through' : ''].filter(Boolean).join(' ')
  if (deco) css.textDecoration = deco
  else if (s.u === 'none' || s.strike === false) css.textDecoration = 'none'
  if (u && UNDERLINE_STYLE[s.u!]) css.textDecorationStyle = UNDERLINE_STYLE[s.u!]
  if (s.dstrike) css.textDecorationStyle = 'double'
  if (s.color) css.color = s.color
  if (s.sz) css.fontSize = `${s.va ? s.sz * 0.65 : s.sz}pt`
  if (s.font) css.fontFamily = s.font
  if (s.hl || s.shd) css.background = s.hl ?? s.shd
  if (s.va) css.verticalAlign = s.va === 'sup' ? 'super' : 'sub'
  if (s.caps) css.textTransform = 'uppercase'
  if (s.smallCaps) css.fontVariant = 'small-caps'
  if (s.spacing) css.letterSpacing = `${s.spacing}pt`
  return css
}

function lineHeight(ps: ParaStyle): string | number | undefined {
  if (!ps.line) return 1.17
  if (ps.lineRule === 'exact') return `${ps.line}pt`
  if (ps.lineRule === 'atLeast') return `max(${ps.line}pt, 1.17em)`
  return +(ps.line * 1.17).toFixed(3)
}

/** contextual spacing flags: 'b' drop space before, 'a' drop space after */
export type Spacing = '' | 'b' | 'a' | 'ba'

export function contextual(p: Block | undefined, prev: Block | undefined, next: Block | undefined): Spacing {
  if (!p || p.t !== 'p' || !p.ps.contextual) return ''
  const b = prev?.t === 'p' && prev.sid === p.sid ? 'b' : ''
  const a = next?.t === 'p' && next.sid === p.sid ? 'a' : ''
  return (b + a) as Spacing
}

function paraCss(ps: ParaStyle, sp: Spacing): CSSProperties {
  const css: CSSProperties = {
    paddingTop: sp.includes('b') ? undefined : pt(ps.before),
    paddingBottom: sp.includes('a') ? undefined : pt(ps.after),
    paddingLeft: pt(ps.indL),
    paddingRight: pt(ps.indR),
    textIndent: pt(ps.indFirst),
    textAlign: ps.align,
    lineHeight: lineHeight(ps),
  }
  if (ps.shd) css.background = ps.shd
  const b = (x: ParaStyle['bTop']) => (x ? `${x.w}pt ${x.style} ${x.color}` : undefined)
  if (ps.bTop) css.borderTop = b(ps.bTop)
  if (ps.bBottom) css.borderBottom = b(ps.bBottom)
  if (ps.bLeft) css.borderLeft = b(ps.bLeft)
  if (ps.bRight) css.borderRight = b(ps.bRight)
  return css
}

function Img({ ctx, target, w, h, alt }: { ctx: RenderCtx; target: string; w: number; h: number; alt?: string }) {
  const url = useMediaUrl(ctx.client, target, !ctx.measuring)
  const size = { width: `${w}pt`, height: `${h}pt` }
  if (ctx.measuring || !url) return <span className="de-img-missing" style={{ ...size, borderStyle: 'none' }} />
  if (url === 'unsupported') {
    return (
      <span className="de-img-missing" style={size} title={target}>
        {target.split('.').pop()?.toUpperCase()} image
      </span>
    )
  }
  return <img src={url} alt={alt ?? ''} style={{ ...size, verticalAlign: 'bottom', display: 'inline-block' }} draggable={false} />
}

function InlineView({ x, ctx, o }: { x: Inline; ctx: RenderCtx; o: number }): ReactNode {
  const obj = { 'data-o': o, 'data-obj': '', contentEditable: ctx.editable ? false : undefined } as const
  switch (x.t) {
    case 'text': {
      const span = (
        <span data-o={o} style={runCss(x.s)} data-cm={x.cm?.join(' ')}>
          {x.text}
        </span>
      )
      return x.href && !ctx.measuring && !ctx.editable ? (
        <a href={x.href} target={x.href.startsWith('#') ? undefined : '_blank'} rel="noreferrer noopener">
          {span}
        </a>
      ) : (
        span
      )
    }
    case 'field': {
      const v = x.kind === 'PAGE' ? (ctx.page >= 0 ? ctx.page + 1 : 1) : ctx.pages || 1
      return (
        <span {...obj} style={runCss(x.s)}>
          {v}
        </span>
      )
    }
    case 'ref':
      return (
        <span {...obj} style={runCss(x.s)}>
          {x.text}
        </span>
      )
    case 'img': {
      const img = <Img ctx={ctx} target={x.target} w={x.w} h={x.h} alt={x.alt} />
      return (
        <span {...obj} data-cm={x.cm?.join(' ')}>
          {x.href && !ctx.measuring && !ctx.editable ? (
            <a href={x.href} target="_blank" rel="noreferrer noopener">
              {img}
            </a>
          ) : (
            img
          )}
        </span>
      )
    }
    case 'box':
      return (
        <span
          {...obj}
          className="de-tbox"
          style={{ width: x.w ? `${x.w}pt` : undefined, minHeight: x.h ? `${x.h}pt` : undefined, background: x.fill, border: x.border }}
        >
          {x.blocks.map((b, i) => (
            <BlockView key={i} block={b} ctx={ctx} />
          ))}
        </span>
      )
    case 'shape':
      return x.label ? (
        <span {...obj} className="de-ph" style={{ width: `${x.w}pt`, height: `${x.h}pt` }}>
          {x.label}
        </span>
      ) : (
        <span {...obj} />
      )
  }
}

function ParaView({ p, ctx, pa, sp }: { p: Paragraph; ctx: RenderCtx; pa?: string; sp: Spacing }) {
  const hang = p.ps.indFirst && p.ps.indFirst < 0 ? -p.ps.indFirst : 0
  const visible = p.runs.length > 0
  const style = paraCss(p.ps, sp)
  if (!visible) {
    style.fontSize = pt(p.markSz)
    if (p.markFont) style.fontFamily = p.markFont
  }
  const editable = ctx.editable && pa !== undefined
  let o = 0
  const runs = p.runs.map((x, i) => {
    const node = <InlineView key={i} x={x} ctx={ctx} o={o} />
    o += x.t === 'text' ? x.text.length : 1
    return node
  })
  return (
    <div className="de-p" style={style} data-pa={pa} contentEditable={ctx.editable && !editable ? false : undefined}>
      {p.marker && (
        <span className="de-marker" contentEditable={ctx.editable ? false : undefined} style={{ ...runCss(p.marker.s), minWidth: p.marker.tab && hang ? `${hang}pt` : undefined }}>
          {p.marker.text}
          {p.marker.tab && !hang ? '\t' : ''}
        </span>
      )}
      {runs}
      {!visible &&
        (editable ? (
          <span data-o={0} data-empty="">
            {'\u200B'}
          </span>
        ) : (
          '\u200B'
        ))}
    </div>
  )
}

function TableView({ t, ctx, pa }: { t: Table; ctx: RenderCtx; pa?: string }) {
  const total = t.cols.reduce((a, b) => a + b, 0)
  const style: CSSProperties = { width: total ? `${total}pt` : '100%' }
  if (t.align === 'center') style.margin = '0 auto'
  else if (t.align === 'right') style.marginLeft = 'auto'
  else if (t.indL) style.marginLeft = `${t.indL}pt`
  return (
    <table className="de-tbl" style={style}>
      {total > 0 && (
        <colgroup>
          {t.cols.map((w, i) => (
            <col key={i} style={{ width: `${w}pt` }} />
          ))}
        </colgroup>
      )}
      <tbody>
        {t.rows.map((row, ri) => (
          <tr key={ri} style={{ height: row.h ? `${row.h}pt` : undefined }}>
            {row.cells.map((c, ci) => (
              <td
                key={ci}
                colSpan={c.colSpan > 1 ? c.colSpan : undefined}
                rowSpan={c.rowSpan > 1 ? c.rowSpan : undefined}
                style={{
                  borderTop: c.borders.t,
                  borderRight: c.borders.r,
                  borderBottom: c.borders.b,
                  borderLeft: c.borders.l,
                  background: c.shd,
                  verticalAlign: c.vAlign,
                  padding: `${c.pad[0]}pt ${c.pad[1]}pt ${c.pad[2]}pt ${c.pad[3]}pt`,
                }}
              >
                {c.blocks.map((b, i) => (
                  <BlockView
                    key={b.t === 'p' ? `${i}:${b.ver ?? 0}` : i}
                    block={b}
                    ctx={ctx}
                    pa={pa !== undefined && b.t === 'p' && b.pid !== undefined && !b.ro ? `${pa}.${ri}.${ci}.${i}` : undefined}
                    sp={contextual(b, c.blocks[i - 1], c.blocks[i + 1])}
                    nested
                  />
                ))}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

export interface BlockViewProps {
  block: Block
  ctx: RenderCtx
  /** editable address (top-level: block index; table: block index used as prefix) */
  pa?: string
  sp?: Spacing
  nested?: boolean
}

export const BlockView = memo(function BlockView({ block, ctx, pa, sp = '', nested }: BlockViewProps) {
  switch (block.t) {
    case 'p':
      return <ParaView p={block} ctx={ctx} pa={pa} sp={sp} />
    case 'tbl':
      return (
        <div style={{ overflow: 'visible' }} contentEditable={ctx.editable && (nested || pa === undefined) ? false : undefined}>
          <TableView t={block} ctx={ctx} pa={nested ? undefined : pa} />
        </div>
      )
    default:
      return <div />
  }
})
