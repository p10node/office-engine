/**
 * @p10node/diagram: Mermaid, PlantUML and Graphviz DOT text → SVG, with no DOM (browser, worker, Node, Bun),
 * no server and no dependency.
 */
import { renderDiagrams, sniffDiagram } from '../../../src/core/diagram'
import { show, svgToPng } from './dom'

export type DiagramLang = 'mermaid' | 'plantuml' | 'graphviz'

/** A drawing: a self-contained SVG document and its size in CSS pixels. */
export interface Diagram {
  svg: string
  width: number
  height: number
}

/** The source cannot be drawn (unknown or unsupported diagram kind, syntax the reader does not follow). */
export class DiagramError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'DiagramError'
  }
}

/** The language of a diagram source: `@startuml`, `digraph {`, a Mermaid header (`graph TD`, `sequenceDiagram`…), or null. */
export function detect(source: string): DiagramLang | null {
  return sniffDiagram(source)
}

/** Every drawing of a source (a PlantUML file or a DOT file may hold several). `lang`: detected when left out. */
export function renderEach(source: string, lang?: DiagramLang): Diagram[] {
  const l = lang ?? sniffDiagram(source)
  if (!l) throw new DiagramError('Not a Mermaid, PlantUML or Graphviz diagram')
  try {
    return renderDiagrams(l, source).map((d) => ({ svg: d.svg, width: d.w, height: d.h }))
  } catch (e) {
    throw new DiagramError((e as Error).message)
  }
}

/** The (first) drawing of a source. */
export function render(source: string, lang?: DiagramLang): Diagram {
  return renderEach(source, lang)[0]
}

/** Draws a source into an element (its content is replaced). */
export function renderTo(target: Element, source: string, lang?: DiagramLang): Diagram[] {
  const ds = renderEach(source, lang)
  show(
    target,
    ds.map((d) => d.svg),
  )
  return ds
}

const BLOCKS: [string, DiagramLang | null][] = [
  ['.mermaid', 'mermaid'],
  ['.language-mermaid', 'mermaid'],
  ['.plantuml', 'plantuml'],
  ['.language-plantuml', 'plantuml'],
  ['.language-puml', 'plantuml'],
  ['.graphviz', 'graphviz'],
  ['.language-graphviz', 'graphviz'],
  ['.language-dot', 'graphviz'],
  ['.diagram', null],
]

/**
 * Draws the diagram blocks of a page: `<pre class="mermaid">`, `<div class="plantuml">`, `<pre><code
 * class="language-dot">`, `<pre class="diagram">` (language detected)… Each becomes a
 * `<figure class="diagram diagram-<lang>">`; a block that cannot be drawn stays and gets a `data-diagram-error`
 * attribute with the reason. Returns how many were drawn. Browser only.
 */
export function renderPage(root: ParentNode = document): number {
  let drawn = 0
  for (const [sel, lang] of BLOCKS) {
    for (const el of Array.from(root.querySelectorAll<HTMLElement>(sel))) {
      if (!el.isConnected || el.hasAttribute('data-diagram-error') || el.closest('figure.diagram')) continue
      const host = el.tagName === 'CODE' && el.parentElement?.tagName === 'PRE' ? el.parentElement : el
      const source = el.textContent ?? ''
      try {
        const ds = renderEach(source, lang ?? undefined)
        const fig = el.ownerDocument.createElement('figure')
        fig.className = `diagram diagram-${lang ?? sniffDiagram(source)}`
        show(
          fig,
          ds.map((d) => d.svg),
        )
        host.replaceWith(fig)
        drawn++
      } catch (e) {
        el.setAttribute('data-diagram-error', (e as Error).message)
      }
    }
  }
  return drawn
}

/** A PNG of a drawing's SVG, `scale` times its size. Browser only. */
export function toPng(svg: string, scale = 2): Promise<Blob> {
  return svgToPng(svg, scale)
}
