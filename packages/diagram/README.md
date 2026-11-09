# @p10node/diagram

Mermaid, PlantUML and Graphviz DOT text → SVG. No DOM needed (browser, Web Worker, Node, Bun), no server, no WebAssembly, no dependency.

Part of [office-engine](https://github.com/p10node/office-engine); single-language packages: [`@p10node/mermaid`](../mermaid), [`@p10node/plantuml`](../plantuml), [`@p10node/graphviz`](../graphviz). Markdown with diagrams: [`@p10node/markdown`](../markdown).

## Install

```sh
npm i @p10node/diagram
```

## Use

```ts
import { detect, render, renderEach, renderPage, renderTo, toPng } from '@p10node/diagram'

const { svg, width, height } = render('graph LR\n  A --> B') // language detected
render('A -> B : hi', 'plantuml')                              // or given
renderEach(pumlFileText)                                       // every @startuml block, every DOT graph
detect('digraph { a -> b }')                                   // 'graphviz'

renderTo(document.getElementById('out')!, source)               // into an element
renderPage()                                                   // every diagram block of the page
const png = await toPng(svg, 2)                                // PNG Blob, 2× size (browser)
```

`renderPage(root = document)` draws `<pre class="mermaid">`, `<div class="plantuml">`, `<pre class="graphviz">`, `<pre><code class="language-mermaid | plantuml | puml | dot | graphviz">` and `<pre class="diagram">` (language detected). Each becomes `<figure class="diagram diagram-<lang>">`; a block that cannot be drawn stays, with a `data-diagram-error` attribute. It returns how many it drew.

From a CDN, no build step:

```html
<pre class="mermaid">sequenceDiagram
  Alice->>Bob: Hello</pre>
<script type="module">
  import { renderPage } from 'https://cdn.jsdelivr.net/npm/@p10node/diagram/dist/index.js'
  renderPage()
</script>
```

## API

| Function                           | Returns               | Notes                                                              |
|------------------------------------|-----------------------|--------------------------------------------------------------------|
| `render(source, lang?)`            | `Diagram`             | The first drawing; `lang`: `'mermaid' \| 'plantuml' \| 'graphviz'` |
| `renderEach(source, lang?)`        | `Diagram[]`           | Every drawing of the source                                        |
| `detect(source)`                   | `DiagramLang \| null` | From `@startuml`, `digraph {`, a Mermaid header                    |
| `renderTo(element, source, lang?)` | `Diagram[]`           | Replaces the element's content (browser)                           |
| `renderPage(root?)`                | `number`              | Draws the page's diagram blocks (browser)                          |
| `toPng(svg, scale = 2)`            | `Promise<Blob>`       | PNG of a drawing (browser)                                         |

`Diagram` is `{ svg: string; width: number; height: number }` (CSS pixels). A source that cannot be drawn throws a `DiagramError` with the reason.

## What is drawn

| Language | Diagrams                                                                                                |
|----------|---------------------------------------------------------------------------------------------------------|
| Mermaid  | flowchart / graph (subgraphs, `classDef`, `style`), sequence, class, state, ER, pie, Gantt, mind map    |
| PlantUML | sequence, class / object, use case, component / deployment, state, activity, mind map, WBS, `@startdot` |
| Graphviz | graph / digraph, clusters, default attributes, edge chains, `rank=same`, record and HTML labels (text)  |

Laid out by a built-in layered layout (like Mermaid's and Graphviz's `dot`), so drawings look close to, not exactly like, the originals. Not drawn: other Mermaid kinds (git graph, journey, timeline, C4…), other PlantUML kinds (Gantt, JSON / YAML, salt, timing), Graphviz engines other than `dot`.

## Showing the SVG

The SVG has a fixed `width` / `height`; to fit a column:

```css
figure.diagram svg { max-width: 100%; height: auto; }
```

## Safety

Text from the source is escaped; only plain colors are kept from styles; the SVG has no scripts, links, `<foreignObject>` or external references. `renderTo` / `renderPage` only ever insert SVG built this way.

## License

MIT
