/**
 * @p10node/mermaid: Mermaid text → SVG, with no DOM (browser, worker, Node, Bun), no server and no dependency.
 */
import { renderMermaid } from '../../../src/core/diagram/mermaid'
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

/** The drawing of a source. */
export function render(source: string): Diagram {
  try {
    const d = renderMermaid(source)
    return { svg: d.svg, width: d.w, height: d.h }
  } catch (e) {
    throw new DiagramError((e as Error).message)
  }
}

/** Draws a source into an element (its content is replaced). */
export function renderTo(target: Element, source: string): Diagram {
  const d = render(source)
  show(target, [d.svg])
  return d
}

/** A PNG of a drawing's SVG, `scale` times its size. Browser only. */
export function toPng(svg: string, scale = 2): Promise<Blob> {
  return svgToPng(svg, scale)
}
