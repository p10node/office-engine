# @p10node/plantuml

PlantUML text → SVG. No DOM needed (browser, Web Worker, Node, Bun), no server, no dependency.

Part of [office-engine](https://github.com/p10node/office-engine). All languages in one package, and drawing the diagram blocks of a page: [`@p10node/diagram`](../diagram).

## Install

```sh
npm i @p10node/plantuml
```

## Use

```ts
import { render, renderEach, renderTo, toPng } from '@p10node/plantuml'

const { svg, width, height } = render('@startuml\nAlice -> Bob : hello\n@enduml')
const all = renderEach(text)                       // one per @startuml … @enduml block
renderTo(document.getElementById('out')!, text)    // into an element (browser)
const png = await toPng(svg, 2)                    // PNG Blob, 2× size (browser)
```

`render` returns `{ svg, width, height }` (CSS pixels); a source that cannot be drawn throws a `DiagramError` with the reason.

From a CDN, no build step:

```html
<div id="out"></div>
<script type="module">
  import { renderTo } from 'https://cdn.jsdelivr.net/npm/@p10node/plantuml/dist/index.js'
  renderTo(document.getElementById('out'), '@startuml\nAlice -> Bob : hello\n@enduml')
</script>
```

## What is drawn

sequence (activations, `return`, notes, alt / loop / group…, dividers, delays), class / object, use case, component / deployment (packages, nodes, clouds…), state, activity (new syntax: if / elseif / switch / while / repeat / fork / partitions), mind map, WBS, `@startdot` (Graphviz).

Laid out by a built-in layered layout, so drawings look close to, not exactly like, the original tool's. Not drawn yet: Gantt, JSON / YAML, salt, timing, network, ditaa, old-style activity `if`.

The SVG has a fixed `width` / `height`; to fit a column: `svg { max-width: 100%; height: auto; }`.

## Safety

Text from the source is escaped; only plain colors are kept; the SVG has no scripts, links, `<foreignObject>` or external references.

## License

MIT
