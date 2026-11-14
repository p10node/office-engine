import { memo, useId, type CSSProperties, type ReactNode } from 'react'
import { cssColor, cssPaint } from '../../core/color'
import type { Box, Fill, SlideEl, SlideModel, Stroke, TCell, TextBody, TRun } from '../../core/pptx/types'
import { safeHref } from '../../core/url'
import type { DocumentClient } from '../../core/worker/client'
import { useMediaUrl } from '../common/hooks'
import { geomPath, LINE_GEOMS } from './geometry'

function boxStyle(b: Box): CSSProperties {
  const t: string[] = []
  if (b.rot) t.push(`rotate(${b.rot}deg)`)
  return { left: b.x, top: b.y, width: b.w, height: b.h, transform: t.length ? t.join(' ') : undefined }
}

/**
 * A fill as CSS. `background` loads what a url() in it names, so a color or gradient of the model (read from the
 * file, or sent by a peer) is only used when it is exactly that: see `cssColor`.
 */
function useFillCss(client: DocumentClient, fill: Fill | undefined, enabled: boolean): CSSProperties | undefined {
  const url = useMediaUrl(client, fill?.t === 'img' ? fill.target : undefined, enabled)
  if (!fill) return undefined
  switch (fill.t) {
    case 'solid': {
      const c = cssColor(fill.c)
      return c ? { background: c } : undefined
    }
    case 'grad': {
      const css = cssPaint(fill.css)
      return css ? { background: css } : undefined
    }
    case 'img':
      return url && url !== 'unsupported'
        ? { backgroundImage: `url("${url}")`, backgroundSize: fill.tile ? 'auto' : '100% 100%', backgroundRepeat: fill.tile ? 'repeat' : 'no-repeat' }
        : undefined
    default:
      return undefined
  }
}

function runStyle(r: TRun): CSSProperties {
  const s: CSSProperties = { fontSize: `${r.baseline ? r.sz * 0.65 : r.sz}pt` }
  if (r.b) s.fontWeight = 700
  if (r.i) s.fontStyle = 'italic'
  const deco = [r.u ? 'underline' : '', r.strike ? 'line-through' : ''].filter(Boolean).join(' ')
  if (deco) s.textDecoration = deco
  if (r.color) s.color = cssColor(r.color)
  if (r.font) s.fontFamily = r.font
  if (r.baseline) s.verticalAlign = r.baseline > 0 ? 'super' : 'sub'
  if (r.caps) s.textTransform = 'uppercase'
  if (r.hl) s.background = cssColor(r.hl)
  return s
}

export interface TextEdit {
  /** contentEditable root callback */
  root: (el: HTMLDivElement | null) => void
}

export function TextView({ text, flow, edit }: { text: TextBody; flow?: boolean; edit?: TextEdit }) {
  const [l, t, r, b] = text.ins
  const style: CSSProperties = {
    ...(flow ? { position: 'relative', inset: 'auto', minHeight: '100%' } : {}),
    padding: `${t}px ${r}px ${b}px ${l}px`,
    justifyContent: text.anchor === 'ctr' ? 'center' : text.anchor === 'b' ? 'flex-end' : 'flex-start',
  }
  if (text.vert) style.writingMode = text.vert === 'vert270' ? 'sideways-lr' : 'vertical-rl'
  if (text.empty && !edit) return <div className="de-txt" style={style} />
  return (
    <div
      ref={edit?.root}
      className={(text.wrap ? 'de-txt' : 'de-txt de-nowrap') + (edit ? ' de-txt-edit' : '')}
      style={style}
      contentEditable={edit ? true : undefined}
      suppressContentEditableWarning
      spellCheck={edit ? false : undefined}
    >
      {text.paras.map((p, i) => {
        const ps: CSSProperties = {
          textAlign: p.algn,
          paddingLeft: p.marL || undefined,
          textIndent: p.indent || undefined,
          paddingTop: p.spcBef || undefined,
          paddingBottom: p.spcAft || undefined,
          lineHeight: p.lnPx ? `${p.lnPx}px` : p.lnPct ? +(p.lnPct * 1.2).toFixed(3) : 1.2,
        }
        if (!p.runs.length || p.runs.every((x) => !x.text && !x.br)) ps.fontSize = `${p.endSz}pt`
        const hang = p.indent < 0 ? -p.indent : 0
        let o = 0
        const nodes: ReactNode[] = p.runs.map((r, k) => {
          const at = o
          o += r.fld ? 1 : r.br ? 1 : r.text.length
          if (edit) {
            if (r.fld)
              return (
                <span key={k} data-o={at} data-obj="" contentEditable={false} style={runStyle(r)}>
                  {r.text}
                </span>
              )
            return (
              <span key={k} data-o={at} style={runStyle(r)}>
                {r.br ? '\n' : r.text}
              </span>
            )
          }
          const href = safeHref(r.href)
          return r.br ? (
            <br key={k} />
          ) : href ? (
            <a key={k} href={href} target="_blank" rel="noreferrer noopener" style={{ color: 'inherit' }}>
              <span style={runStyle(r)}>{r.text}</span>
            </a>
          ) : (
            <span key={k} style={runStyle(r)}>
              {r.text}
            </span>
          )
        })
        const empty = o === 0
        return (
          <div key={i} style={ps} data-pa={edit ? String(i) : undefined}>
            {p.bullet && !empty && (
              <span
                className="de-bul"
                contentEditable={edit ? false : undefined}
                style={{ minWidth: hang || undefined, paddingRight: hang ? undefined : '0.4em', fontSize: `${p.bullet.sz}pt`, color: cssColor(p.bullet.color), fontFamily: p.bullet.font }}
              >
                {p.bullet.text}
              </span>
            )}
            {nodes}
            {empty &&
              (edit ? (
                <span data-o={0} data-empty="">
                  {'\u200B'}
                </span>
              ) : (
                '\u200B'
              ))}
          </div>
        )
      })}
    </div>
  )
}

function Marker({ id, s }: { id: string; s: Stroke }) {
  return (
    <marker id={id} viewBox="0 0 10 10" refX="5" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse">
      <path d="M0 0 L10 5 L0 10 Z" fill={cssColor(s.c)} />
    </marker>
  )
}

export interface SlideEditing {
  sid: number
  text: TextBody
  edit: TextEdit
}

function ShapeView({ el, client, live, editing }: { el: Extract<SlideEl, { t: 'shape' }>; client: DocumentClient; live: boolean; editing?: SlideEditing }) {
  const id = useId().replace(/:/g, '')
  const { box } = el
  const isLine = LINE_GEOMS.has(el.geom)
  const fillCss = useFillCss(client, el.fill, live)
  const d = el.path ?? geomPath(el.geom, box.w, box.h, el.adj)
  const flipped = box.flipH || box.flipV
  const flip = flipped ? `translate(${box.flipH ? box.w : 0} ${box.flipV ? box.h : 0}) scale(${box.flipH ? -1 : 1} ${box.flipV ? -1 : 1})` : undefined
  const isRect = !isLine && !el.path && (el.geom === 'rect' || el.geom === 'flowChartProcess')
  // an SVG fill / stroke can name a url() too
  const solid = !isLine && el.fill?.t === 'solid' ? cssColor(el.fill.c) : undefined
  const css = !isLine && fillCss && el.fill?.t !== 'solid' ? fillCss : undefined
  const s = el.stroke
  const w = Math.max(box.w, 1)
  const h = Math.max(box.h, 1)
  return (
    <div className="de-el" style={boxStyle(box)} data-sid={el.sid !== undefined && el.text ? el.sid : undefined} data-top={el.top}>
      {isRect && !s && (solid || css) ? (
        <div style={{ position: 'absolute', inset: 0, ...(css ?? { background: solid }) }} />
      ) : (
        <>
          {css && (
            <div
              style={{
                position: 'absolute',
                inset: 0,
                ...css,
                // quoted into the value: path data and nothing that could end the quote
                clipPath: isRect || !/^[\sA-Za-z0-9.,+-]*$/.test(d) ? undefined : `path("${d}")`,
                transform: flipped ? `scale(${box.flipH ? -1 : 1}, ${box.flipV ? -1 : 1})` : undefined,
              }}
            />
          )}
          {(solid || s) && (
            <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ position: 'absolute', left: 0, top: 0, overflow: 'visible' }}>
              {s && (s.head || s.tail) && (
                <defs>
                  {s.head && <Marker id={`h${id}`} s={s} />}
                  {s.tail && <Marker id={`t${id}`} s={s} />}
                </defs>
              )}
              <path
                d={d}
                transform={flip}
                fill={solid ?? 'none'}
                fillRule="evenodd"
                stroke={cssColor(s?.c) ?? 'none'}
                strokeWidth={s?.w}
                strokeDasharray={s?.dash}
                strokeLinejoin="round"
                markerStart={s?.head ? `url(#h${id})` : undefined}
                markerEnd={s?.tail ? `url(#t${id})` : undefined}
              />
            </svg>
          )}
        </>
      )}
      {editing && editing.sid === el.sid ? <TextView text={editing.text} edit={editing.edit} /> : el.text && <TextView text={el.text} />}
    </div>
  )
}

function PicView({ el, client, live }: { el: Extract<SlideEl, { t: 'pic' }>; client: DocumentClient; live: boolean }) {
  const url = useMediaUrl(client, el.target, live)
  const { box } = el
  const style: CSSProperties = { ...boxStyle(box), overflow: 'hidden' }
  const top = el.top
  if (el.geom === 'ellipse') style.borderRadius = '50%'
  else if (el.geom === 'roundRect') style.borderRadius = Math.min(box.w, box.h) * 0.1667
  if (el.stroke) style.outline = `${el.stroke.w}px solid ${cssColor(el.stroke.c) ?? 'transparent'}`
  let img: CSSProperties = { position: 'absolute', left: 0, top: 0, width: '100%', height: '100%' }
  if (el.crop) {
    const [l, t, r, b] = el.crop
    const sw = box.w / Math.max(0.01, 1 - l - r)
    const sh = box.h / Math.max(0.01, 1 - t - b)
    img = { position: 'absolute', left: -l * sw, top: -t * sh, width: sw, height: sh }
  }
  const flip = box.flipH || box.flipV ? `scale(${box.flipH ? -1 : 1}, ${box.flipV ? -1 : 1})` : undefined
  if (flip) img.transform = flip
  return (
    <div className="de-el" style={style} data-top={top}>
      {url && url !== 'unsupported' ? (
        <img src={url} alt="" style={img} draggable={false} />
      ) : url === 'unsupported' ? (
        <div className="de-ph" style={{ width: '100%', height: '100%' }}>
          {el.target?.split('.').pop()?.toUpperCase()} image
        </div>
      ) : null}
      {el.video && <div className="de-ph" style={{ position: 'absolute', right: 8, bottom: 8, padding: '2px 6px', border: 0, background: 'rgba(0,0,0,.6)', color: '#fff' }}>▶ media</div>}
    </div>
  )
}

const strokeCss = (s: Stroke | undefined) => (s ? `${Math.max(1, s.w)}px solid ${cssColor(s.c) ?? 'transparent'}` : undefined)

function CellView({ cell, client, live, h }: { cell: TCell; client: DocumentClient; live: boolean; h: number }) {
  const fill = useFillCss(client, cell.fill, live)
  return (
    <td
      colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
      rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
      style={{ ...(fill ?? {}), borderLeft: strokeCss(cell.bl), borderRight: strokeCss(cell.br), borderTop: strokeCss(cell.bt), borderBottom: strokeCss(cell.bb), height: h }}
    >
      {cell.text && <TextView text={cell.text} flow />}
    </td>
  )
}

function TableView({ el, client, live }: { el: Extract<SlideEl, { t: 'table' }>; client: DocumentClient; live: boolean }) {
  return (
    <div className="de-el" style={{ ...boxStyle(el.box), height: 'auto' }} data-top={el.top}>
      <table className="de-stbl">
        <colgroup>
          {el.cols.map((w, i) => (
            <col key={i} style={{ width: w }} />
          ))}
        </colgroup>
        <tbody>
          {el.rows.map((r, ri) => (
            <tr key={ri}>
              {r.cells.map((c, ci) => (c.hidden ? null : <CellView key={ci} cell={c} client={client} live={live} h={r.h} />))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function ElView({ el, client, live, editing }: { el: SlideEl; client: DocumentClient; live: boolean; editing?: SlideEditing }) {
  switch (el.t) {
    case 'shape':
      return <ShapeView el={el} client={client} live={live} editing={editing} />
    case 'pic':
      return <PicView el={el} client={client} live={live} />
    case 'table':
      return <TableView el={el} client={client} live={live} />
    case 'placeholder':
      return (
        <div className="de-el de-ph" style={{ ...boxStyle(el.box), display: 'flex' }} data-top={el.top}>
          {el.label}
        </div>
      )
  }
}

export interface SlideViewProps {
  model: SlideModel
  client: DocumentClient
  scale: number
  /** load media (false for far-away thumbnails) */
  live?: boolean
  editing?: SlideEditing
  /** extra layer drawn in slide coordinates (selection handles) */
  overlay?: ReactNode
}

export const SlideView = memo(function SlideView({ model, client, scale, live = true, editing, overlay }: SlideViewProps) {
  const bg = useFillCss(client, model.bg, live)
  return (
    <div className={model.hidden ? 'de-slide de-hidden-slide' : 'de-slide'} style={{ width: model.w * scale, height: model.h * scale }}>
      <div className="de-slide-inner" style={{ width: model.w, height: model.h, transform: `scale(${scale})`, background: '#fff', ...(bg ?? {}) }}>
        {model.els.map((el, i) => (
          <ElView key={i} el={el} client={client} live={live} editing={editing} />
        ))}
        {overlay}
      </div>
    </div>
  )
})
