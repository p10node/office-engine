import { describe, expect, test } from 'bun:test'
import * as diagram from '../packages/diagram/src/index'
import * as graphviz from '../packages/graphviz/src/index'
import * as markdown from '../packages/markdown/src/index'
import * as mermaid from '../packages/mermaid/src/index'
import * as plantuml from '../packages/plantuml/src/index'

const noSvg = (html: string) => html.replace(/<svg[\s\S]*?<\/svg>/g, '<svg/>')

describe('diagram packages', () => {
  test('render: an SVG and its size', () => {
    for (const d of [mermaid.render('graph LR\n  A --> B'), plantuml.render('@startuml\nA -> B\n@enduml'), graphviz.render('digraph { a -> b }'), diagram.render('graph TD\n  X --> Y')]) {
      expect(d.svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true)
      expect(d.width).toBeGreaterThan(0)
      expect(d.height).toBeGreaterThan(0)
      expect(d.svg).toContain(`width="${d.width}" height="${d.height}"`)
    }
  })

  test('several drawings per source; detection', () => {
    expect(plantuml.renderEach('@startuml\nA -> B\n@enduml\n@startmindmap\n* Root\n** Leaf\n@endmindmap')).toHaveLength(2)
    expect(graphviz.renderEach('digraph a { x -> y }\ngraph b { p -- q }')).toHaveLength(2)
    expect(diagram.detect('@startuml\nA -> B\n@enduml')).toBe('plantuml')
    expect(diagram.detect('strict digraph { a -> b }')).toBe('graphviz')
    expect(diagram.detect('sequenceDiagram\n  A->>B: hi')).toBe('mermaid')
    expect(diagram.detect('just text')).toBeNull()
    expect(diagram.renderEach('digraph { a -> b }\ndigraph { c -> d }')).toHaveLength(2)
    expect(diagram.render('A -> B : hi', 'plantuml').svg).toContain('>hi</text>')
  })

  test("errors are the package's DiagramError", () => {
    for (const [fn, src] of [
      [() => mermaid.render('gitGraph\n  commit'), 'gitGraph'],
      [() => plantuml.render('@startgantt\n[T] lasts 2 days\n@endgantt'), 'gantt'],
      [() => graphviz.render('nothing here'), 'graph'],
      [() => diagram.render('nothing here'), 'Not a Mermaid'],
    ] as [() => unknown, string][]) {
      let err: unknown
      try {
        fn()
      } catch (e) {
        err = e
      }
      expect(err).toBeInstanceOf(Error)
      expect((err as Error).name).toBe('DiagramError')
      expect((err as Error).message).toContain(src)
    }
    expect(() => mermaid.render('gitGraph')).toThrow(mermaid.DiagramError)
    expect(() => diagram.render('x')).toThrow(diagram.DiagramError)
  })
})

describe('@p10node/markdown', () => {
  const md = markdown.render

  test('blocks: headings (ids), paragraphs, rules, quotes, code', () => {
    expect(md('# One\n\nText *a* **b** ~~c~~ `d`\n\nTwo\n---\n\n***\n\n> q\nlazy\n\n    code\n\n```js\nx < 1\n```')).toBe(
      [
        '<h1 id="one">One</h1>',
        '<p>Text <em>a</em> <strong>b</strong> <del>c</del> <code>d</code></p>',
        '<h2 id="two">Two</h2>',
        '<hr>',
        '<blockquote>\n<p>q\nlazy</p>\n</blockquote>',
        '<pre><code>code\n</code></pre>',
        '<pre><code class="language-js">x &lt; 1\n</code></pre>',
      ].join('\n'),
    )
    expect(md('# A\n# A\n## Hello, *World*!', { headingIds: true })).toBe('<h1 id="a">A</h1>\n<h1 id="a-1">A</h1>\n<h2 id="hello-world">Hello, <em>World</em>!</h2>')
    expect(md('# A', { headingIds: false })).toBe('<h1>A</h1>')
  })

  test('lists: tight, loose, nested, numbered from n, tasks', () => {
    expect(md('- a\n- b\n  - c\n- [ ] d\n- [x] e')).toBe(
      '<ul>\n<li>a</li>\n<li>b\n<ul>\n<li>c</li>\n</ul></li>\n<li class="task-list-item"><input type="checkbox" disabled> d</li>\n<li class="task-list-item"><input type="checkbox" disabled checked> e</li>\n</ul>',
    )
    expect(md('3. x\n\n4. y')).toBe('<ol start="3">\n<li><p>x</p></li>\n<li><p>y</p></li>\n</ol>')
    // a list changes type: two lists
    expect(md('- a\n* b')).toBe('<ul>\n<li>a</li>\n</ul>\n<ul>\n<li>b</li>\n</ul>')
  })

  test('tables with alignment and escaped pipes', () => {
    expect(md('| a | b |\n|:-:|--:|\n| 1 | x \\| y |')).toBe(
      '<table>\n<thead>\n<tr><th style="text-align:center">a</th><th style="text-align:right">b</th></tr>\n</thead>\n<tbody>\n<tr><td style="text-align:center">1</td><td style="text-align:right">x | y</td></tr>\n</tbody>\n</table>',
    )
  })

  test('links and pictures: inline, reference, auto, bare', () => {
    expect(md('[a](https://a.com "T") [b][r] [r] <https://c.com> www.d.com ![i](/p.png)\n\n[r]: https://r.com')).toBe(
      '<p><a href="https://a.com" title="T">a</a> <a href="https://r.com">b</a> <a href="https://r.com">r</a> <a href="https://c.com">https://c.com</a> <a href="http://www.d.com">www.d.com</a> <img src="/p.png" alt="i"></p>',
    )
  })

  test('safe: HTML is text, script links and pictures are dropped', () => {
    const html = md('<script>alert(1)</script>\n\n[x](javascript:alert(1)) [y](JaVaScRiPt:alert(1)) ![z](javascript:alert(1)) ![w](data:text/html;base64,AAAA) ![ok](data:image/png;base64,AAAA)\n\n<img src=x onerror=alert(1)> "quoted" & `<b>`')
    expect(html).not.toMatch(/<script|<img src="x|javascript:|data:text/i)
    expect(html).toContain('&lt;script&gt;alert(1)&lt;/script&gt;')
    expect(html).toContain('<img src="data:image/png;base64,AAAA" alt="ok">')
    expect(html).toContain('&lt;img src=x onerror=alert(1)&gt; &quot;quoted&quot; &amp; <code>&lt;b&gt;</code>')
    // link text stays when the target is dropped
    expect(html).toContain('<p>x y z w')
  })

  test('diagram code blocks are drawn, unless asked not to; a broken one shows why', () => {
    const src = '```mermaid\ngraph LR\n  A --> B\n```\n\n```plantuml\nA -> B\n```\n\n~~~dot\ndigraph { a -> b }\n~~~'
    expect(noSvg(md(src))).toBe('<figure class="md-diagram md-mermaid"><svg/></figure>\n<figure class="md-diagram md-plantuml"><svg/></figure>\n<figure class="md-diagram md-graphviz"><svg/></figure>')
    expect(md('```mermaid\ngraph LR\n  A --> B\n```', { diagrams: false })).toBe('<pre><code class="language-mermaid">graph LR\n  A --&gt; B\n</code></pre>')
    expect(md('```mermaid\ngitGraph\n```')).toBe(
      '<div class="md-diagram-error"><p>Mermaid diagram not drawn: Mermaid “gitGraph” diagrams are not supported yet</p><pre><code class="language-mermaid">gitGraph\n</code></pre></div>',
    )
  })
})
