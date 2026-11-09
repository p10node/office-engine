# @p10node/graphviz

Graphviz DOT text → SVG. No DOM needed (browser, Web Worker, Node, Bun), no server, no dependency.

Part of [office-engine](https://github.com/p10node/office-engine). All languages in one package, and drawing the diagram blocks of a page: [`@p10node/diagram`](../diagram).

## Install

```sh
npm i @p10node/graphviz
```

## Use

```ts
import { render, renderEach, renderTo, toPng } from '@p10node/graphviz'

const { svg, width, height } = render('digraph { rankdir=LR; a -> b -> c }')
const all = renderEach(text)                       // one per graph in the file
renderTo(document.getElementById('out')!, text)    // into an element (browser)
const png = await toPng(svg, 2)                    // PNG Blob, 2× size (browser)
```

`render` returns `{ svg, width, height }` (CSS pixels); a source that cannot be drawn throws a `DiagramError` with the reason.

From a CDN, no build step:

```html
<div id="out"></div>
<script type="module">
  import { renderTo } from 'https://cdn.jsdelivr.net/npm/@p10node/graphviz/dist/index.js'
  renderTo(document.getElementById('out'), 'digraph { rankdir=LR; a -> b -> c }')
</script>
```

## What is drawn

graph / digraph / strict, subgraphs and `cluster_*` boxes, `node` / `edge` / `graph` defaults, edge chains and `{a b} -> c`, `rank=same`, shapes (box, ellipse, circle, diamond, record, cylinder, note…), styles (filled, dashed, dotted, bold, rounded, invis), colors (names, `#rgb`, HSV, X11 grays), arrowheads, `dir`, labels (`\n` `\l`, record fields, HTML labels as text).

Laid out by a built-in layered layout, so drawings look close to, not exactly like, the original tool's. Not drawn yet: layout engines other than `dot` (`neato`, `circo`, `fdp`… are laid out as `dot`), ports (edges meet the node), HTML table layouts (their text is shown).

The SVG has a fixed `width` / `height`; to fit a column: `svg { max-width: 100%; height: auto; }`.

## Safety

Text from the source is escaped; only plain colors are kept; the SVG has no scripts, links, `<foreignObject>` or external references.

## License

MIT
