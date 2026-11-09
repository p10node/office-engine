# @p10node/mermaid

Mermaid text → SVG. No DOM needed (browser, Web Worker, Node, Bun), no server, no dependency.

Part of [office-engine](https://github.com/p10node/office-engine). All languages in one package, and drawing the diagram blocks of a page: [`@p10node/diagram`](../diagram).

## Install

```sh
npm i @p10node/mermaid
```

## Use

```ts
import { render, renderTo, toPng } from '@p10node/mermaid'

const { svg, width, height } = render('graph LR\n  A[Write] --> B[Draw]')
renderTo(document.getElementById('out')!, text)    // into an element (browser)
const png = await toPng(svg, 2)                    // PNG Blob, 2× size (browser)
```

`render` returns `{ svg, width, height }` (CSS pixels); a source that cannot be drawn throws a `DiagramError` with the reason.

From a CDN, no build step:

```html
<div id="out"></div>
<script type="module">
  import { renderTo } from 'https://cdn.jsdelivr.net/npm/@p10node/mermaid/dist/index.js'
  renderTo(document.getElementById('out'), 'graph LR\n  A[Write] --> B[Draw]')
</script>
```

## What is drawn

flowchart / graph (subgraphs, `classDef`, `style`, `linkStyle`), sequence (activations, notes, loop / alt / opt / par / critical / break / rect), class, state (composite states, fork / join / choice), ER (crow's foot), pie, Gantt, mind map.

Laid out by a built-in layered layout, so drawings look close to, not exactly like, the original tool's. Not drawn yet: git graph, journey, timeline, quadrant, requirement, C4, Sankey, XY chart, block, packet, kanban, architecture.

The SVG has a fixed `width` / `height`; to fit a column: `svg { max-width: 100%; height: auto; }`.

## Safety

Text from the source is escaped; only plain colors are kept; the SVG has no scripts, links, `<foreignObject>` or external references.

## License

MIT
