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
}

const pt = (v: number | undefined) => (v ? `${+v.toFixed(2)}pt` : undefined)

const UNDERLINE_STYLE: Record<string, CSSProperties['textDecorationStyle']> = {
  double: 'double', dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashedHeavy: 'dashed', dashLong: 'dashed',
  wave: 'wavy', wavyHeavy: 'wavy', wavyDouble: 'wavy',
}

export function runCss(s: RunStyle): CSSProperties {
  const css: CSSProperties = {}
  if (s.b) css.fontWeight = 700
  else if (s.b === false) css.fontWeight = 400
  if (s.i) css.fontStyle = 'italic'
  const u = s.u && s.u !== 'none'
  const deco = [u ? 'underline' : '', s.strike || s.dstrike ? 'line-through' : ''].filter(Boolean).join(' ')
  if (deco) css.textDecoration = deco
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

function paraCss(ps: ParaStyle): CSSProperties {
  const css: CSSProperties = {
    paddingTop: pt(ps.before),
    paddingBottom: pt(ps.after),
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

function InlineView({ x, ctx }: { x: Inline; ctx: RenderCtx }): ReactNode {
  switch (x.t) {
    case 'text': {
      const span = <span style={runCss(x.s)}>{x.text}</span>
      return x.href && !ctx.measuring ? (
        <a href={x.href} target={x.href.startsWith('#') ? undefined : '_blank'} rel="noreferrer noopener">
          {span}
        </a>
      ) : (
        span
      )
    }
    case 'tab':
      return <span style={runCss(x.s)}>{'\t'}</span>
    case 'br':
      return <br />
    case 'field': {
      const v = x.kind === 'PAGE' ? (ctx.page >= 0 ? ctx.page + 1 : 1) : ctx.pages || 1
      return <span style={runCss(x.s)}>{v}</span>
    }
    case 'img': {
      const img = <Img ctx={ctx} target={x.target} w={x.w} h={x.h} alt={x.alt} />
      return x.href && !ctx.measuring ? (
        <a href={x.href} target="_blank" rel="noreferrer noopener">
          {img}
        </a>
      ) : (
        img
      )
    }
    case 'box':
      return (
        <span
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
        <span className="de-ph" style={{ width: `${x.w}pt`, height: `${x.h}pt` }}>
          {x.label}
        </span>
      ) : null
  }
}

function ParaView({ p, ctx }: { p: Paragraph; ctx: RenderCtx }) {
  const hang = p.ps.indFirst && p.ps.indFirst < 0 ? -p.ps.indFirst : 0
  const visible = p.runs.length > 0
  const style = paraCss(p.ps)
  if (!visible) {
    style.fontSize = pt(p.markSz)
    if (p.markFont) style.fontFamily = p.markFont
  }
  return (
    <div className="de-p" style={style}>
      {p.marker && (
        <span className="de-marker" style={{ ...runCss(p.marker.s), minWidth: p.marker.tab && hang ? `${hang}pt` : undefined }}>
          {p.marker.text}
          {p.marker.tab && !hang ? '\t' : ''}
        </span>
      )}
      {p.runs.map((x, i) => (
        <InlineView key={i} x={x} ctx={ctx} />
      ))}
      {!visible && '\u200B'}
    </div>
  )
}

function TableView({ t, ctx }: { t: Table; ctx: RenderCtx }) {
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
                  <BlockView key={i} block={b} ctx={ctx} />
                ))}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

export const BlockView = memo(function BlockView({ block, ctx }: { block: Block; ctx: RenderCtx }) {
  switch (block.t) {
    case 'p':
      return <ParaView p={block} ctx={ctx} />
    case 'tbl':
      return (
        <div style={{ overflow: 'visible' }}>
          <TableView t={block} ctx={ctx} />
        </div>
      )
    default:
      return <div />
  }
})
