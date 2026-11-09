/**
 * @p10node/graphviz: Graphviz DOT text → SVG, with no DOM (browser, worker, Node, Bun), no server and no dependency.
 */
import { renderDotAll } from '../../../src/core/diagram/dot'
import { show, svgToPng } from '../../diagram/src/dom'

/** A drawing: a self-contained SVG document and its size in CSS pixels. */
export interface Diagram {
  svg: string
  width: number
  height: number
}

/** The source cannot be drawn (unsupported diagram kind, syntax the reader does not follow). */
export class DiagramError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'DiagramError'
  }
}

/** Every drawing of a source (one per graph). */
export function renderEach(source: string): Diagram[] {
  try {
    return renderDotAll(source).map((d) => ({ svg: d.svg, width: d.w, height: d.h }))
  } catch (e) {
    throw new DiagramError((e as Error).message)
  }
}

/** The (first) drawing of a source. */
export function render(source: string): Diagram {
  return renderEach(source)[0]
}

/** Draws a source into an element (its content is replaced). */
export function renderTo(target: Element, source: string): Diagram[] {
  const ds = renderEach(source)
  show(
    target,
    ds.map((d) => d.svg),
  )
  return ds
}

/** A PNG of a drawing's SVG, `scale` times its size. Browser only. */
export function toPng(svg: string, scale = 2): Promise<Blob> {
  return svgToPng(svg, scale)
}
