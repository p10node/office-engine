# @p10node/markdown

Markdown → HTML that is safe to show, with ` ```mermaid `, ` ```plantuml ` and ` ```dot ` code blocks drawn as SVG. No DOM needed, no server, no dependency.

Part of [office-engine](https://github.com/p10node/office-engine). Diagrams alone: [`@p10node/diagram`](../diagram).

## Install

```sh
npm i @p10node/markdown
```

## Use

```ts
import { render, renderTo } from '@p10node/markdown'

const html = render(markdownText)                       // string
renderTo(document.getElementById('doc')!, markdownText) // into an element
render(text, { diagrams: false, headingIds: false })
```

| Option       | Default | Notes                                                                            |
|--------------|---------|----------------------------------------------------------------------------------|
| `diagrams`   | `true`  | Draw ` ```mermaid `, ` ```plantuml ` / ` ```puml `, ` ```dot ` / ` ```graphviz ` |
| `headingIds` | `true`  | `id` on headings from their text (`# Getting started` → `#getting-started`)      |

## What is read

- CommonMark blocks: ATX and setext headings, paragraphs, block quotes (nested, lazy lines), bullet and numbered lists (nested, tight / loose, `start`), fenced and indented code, thematic breaks
- Inline: emphasis, strong, code spans, links and pictures (inline and reference), autolinks, backslash escapes, hard line breaks
- GitHub: tables (column alignment), task lists, strikethrough, bare `https://` / `www.` links
- Code blocks get `class="language-…"`; diagrams `<figure class="md-diagram md-<lang>">`; a diagram that cannot be drawn shows why and its code (`<div class="md-diagram-error">`)

## Safety

The HTML can go straight into `innerHTML`: all text is escaped and **HTML written in the Markdown is shown as text**, links keep web, mail, phone and relative targets (`javascript:` and the like become plain text), and pictures load web, relative or `data:image/…` sources only.

## License

MIT
