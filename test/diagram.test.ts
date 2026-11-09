import { describe, expect, test } from 'bun:test'
import type { Block, Paragraph } from '../src/core/docx/types'
import { parseDot } from '../src/core/diagram/dot'
import type { Graph } from '../src/core/diagram/graph'
import { DiagramError, renderDiagrams, sniffDiagram, type DiagramLang } from '../src/core/diagram'
import { layered, type LCluster, type LEdge, type LNode } from '../src/core/diagram/layout'
import { parseDate, parseMermaid } from '../src/core/diagram/mermaid'
import type { Model } from '../src/core/diagram/model'
import { parsePlantUml } from '../src/core/diagram/plantuml'
import { color, textWidth } from '../src/core/diagram/svg'
import { plainFormat } from '../src/core/plain'
import { Engine } from '../src/core/worker/engine'

const enc = (s: string) => new TextEncoder().encode(s)

const graphOf = (m: Model): Graph => {
  if (m.t !== 'graph') throw new Error(`expected a graph, got ${m.t}`)
  return m.graph
}
const edgeList = (g: Graph) => g.edges.map((e) => `${e.from}>${e.to}`)

/** what every picture must be: one self-contained SVG with finite numbers and nothing active */
function checkSvg(svg: string, w: number, h: number) {
  expect(svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true)
  expect(svg.endsWith('</svg>')).toBe(true)
  expect(svg).not.toMatch(/NaN|Infinity|undefined/)
  expect(svg).not.toMatch(/<script|<foreignObject|javascript:|url\(|href=|<image|on\w+=/i)
  expect(Number.isFinite(w) && w > 0 && w < 20000).toBe(true)
  expect(Number.isFinite(h) && h > 0 && h < 50000).toBe(true)
  // balanced tags
  const open = (svg.match(/<(svg|text)\b/g) ?? []).length
  const close = (svg.match(/<\/(svg|text)>/g) ?? []).length
  expect(open).toBe(close)
}

const texts = (svg: string) => [...svg.matchAll(/<text[^>]*>([^<]*)<\/text>/g)].map((m) => m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"'))

// ---------------------------------------------------------------- samples: every kind draws

const SAMPLES: [string, DiagramLang, string, string[]][] = [
  ['mermaid flowchart', 'mermaid', 'flowchart TD\n  A[Start] -->|go| B{Ok?}\n  B -- yes --> C((Done))\n  B -. no .-> A', ['Start', 'Ok?', 'Done', 'go', 'yes', 'no']],
  ['mermaid sequence', 'mermaid', 'sequenceDiagram\n  autonumber\n  Alice->>+Bob: Hello\n  loop Every minute\n    Bob-->>-Alice: Hi\n  end\n  Note over Alice,Bob: done', ['Alice', 'Bob', 'Hello', 'Hi', 'loop', 'done']],
  ['mermaid class', 'mermaid', 'classDiagram\n  Animal <|-- Duck\n  Animal : +int age\n  class Duck {\n    +swim()\n  }', ['Animal', 'Duck', '+int age', '+swim()']],
  ['mermaid state', 'mermaid', 'stateDiagram-v2\n  [*] --> Still\n  Still --> Moving : push\n  state Moving {\n    [*] --> Slow\n  }\n  Moving --> [*]', ['Still', 'Moving', 'push', 'Slow']],
  ['mermaid er', 'mermaid', 'erDiagram\n  CUSTOMER ||--o{ ORDER : places\n  ORDER {\n    int id PK\n  }', ['CUSTOMER', 'ORDER', 'places', 'id', 'PK']],
  ['mermaid pie', 'mermaid', 'pie title Pets\n  "Dogs" : 386\n  "Cats" : 85', ['Pets', 'Dogs', 'Cats']],
  ['mermaid gantt', 'mermaid', 'gantt\n  title Plan\n  dateFormat YYYY-MM-DD\n  section A\n  Design :a1, 2024-01-01, 10d\n  Build :after a1, 2w', ['Plan', 'Design', 'Build']],
  ['mermaid mindmap', 'mermaid', 'mindmap\n  root((Topic))\n    One\n      Leaf\n    Two', ['Topic', 'One', 'Leaf', 'Two']],
  ['plantuml sequence', 'plantuml', '@startuml\nactor User\nUser -> App ++ : login\nApp --> User -- : ok\n@enduml', ['User', 'App', 'login', 'ok']],
  ['plantuml class', 'plantuml', '@startuml\nclass Car {\n  - speed : int\n  + drive()\n}\ninterface Vehicle\nVehicle <|.. Car\n@enduml', ['Car', 'Vehicle', '- speed : int', '+ drive()']],
  ['plantuml use case', 'plantuml', '@startuml\nactor Customer\nrectangle Shop {\n  (Checkout)\n}\nCustomer --> (Checkout)\n@enduml', ['Customer', 'Shop', 'Checkout']],
  ['plantuml component', 'plantuml', '@startuml\n[Web] --> [API] : HTTPS\ndatabase DB\n[API] --> DB\n@enduml', ['Web', 'API', 'DB', 'HTTPS']],
  ['plantuml state', 'plantuml', '@startuml\n[*] --> Idle\nIdle --> Busy : job\nBusy --> [*]\n@enduml', ['Idle', 'Busy', 'job']],
  ['plantuml activity', 'plantuml', '@startuml\nstart\n:Read;\nif (ok?) then (yes)\n  :Use;\nelse (no)\n  :Fix;\nendif\nstop\n@enduml', ['Read', 'ok?', 'Use', 'Fix', 'yes', 'no']],
  ['plantuml mindmap', 'plantuml', '@startmindmap\n* Root\n** Right\n-- Left\n@endmindmap', ['Root', 'Right', 'Left']],
  ['plantuml wbs', 'plantuml', '@startwbs\n* Project\n** Phase 1\n*** Task\n@endwbs', ['Project', 'Phase 1', 'Task']],
  ['plantuml dot', 'plantuml', '@startdot\ndigraph { a -> b }\n@enddot', ['a', 'b']],
  ['graphviz', 'graphviz', 'digraph G {\n  rankdir=LR\n  subgraph cluster_x { label="Group"; a -> b }\n  b -> c [label="next"]\n  c [shape=box, style=filled, fillcolor=lightblue]\n}', ['Group', 'a', 'b', 'c', 'next']],
]

describe('diagrams draw', () => {
  for (const [name, lang, src, labels] of SAMPLES) {
    test(name, () => {
      const [r] = renderDiagrams(lang, src)
      checkSvg(r.svg, r.w, r.h)
      const shown = texts(r.svg).join('\n')
      for (const l of labels) expect(shown).toContain(l)
    })
  }

  test('several PlantUML diagrams in one file, several DOT graphs', () => {
    expect(renderDiagrams('plantuml', '@startuml\nA -> B\n@enduml\n\n@startuml\nclass X\n@enduml')).toHaveLength(2)
    expect(renderDiagrams('graphviz', 'digraph a { x -> y }\ngraph b { p -- q }')).toHaveLength(2)
  })

  test('unsupported kinds and broken text are DiagramErrors', () => {
    expect(() => renderDiagrams('mermaid', 'gitGraph\n  commit')).toThrow(DiagramError)
    expect(() => renderDiagrams('mermaid', '')).toThrow(DiagramError)
    expect(() => renderDiagrams('plantuml', '@startgantt\n[Task] lasts 5 days\n@endgantt')).toThrow(/not supported/)
    expect(() => renderDiagrams('graphviz', 'not a graph')).toThrow(DiagramError)
    expect(() => renderDiagrams('mermaid', 'sequenceDiagram\n')).toThrow(DiagramError)
  })

  test('text and colors from the source cannot inject markup', () => {
    const [r] = renderDiagrams('mermaid', 'flowchart LR\n  A["<script>alert(1)</script> a < b & c > d"] --> B\n  style A fill:url(javascript:alert(1)),stroke:red" onload="x\n  classDef c fill:#f00,stroke:expression(1)\n  class B c')
    checkSvg(r.svg, r.w, r.h)
    expect(r.svg).not.toContain('<script')
    expect(r.svg).toContain('alert(1) a &lt; b &amp; c &gt; d')
    expect(r.svg).not.toContain('onload')
    const [d] = renderDiagrams('graphviz', 'digraph { a [label="</text><script>x</script>", color="red\\" onclick=\\"x", fillcolor="url(#x)", style=filled] }')
    checkSvg(d.svg, d.w, d.h)
    expect(d.svg).not.toContain('onclick')
  })

  test('colors: hex, rgb(), names; anything else dropped', () => {
    expect(color('#ABC')).toBe('#abc')
    expect(color('#ff000080')).toBe('#ff000080')
    expect(color('rgb(1, 2, 3)')).toBe('rgb(1, 2, 3)')
    expect(color('LightBlue')).toBe('lightblue')
    expect(color('red;x')).toBeUndefined()
    expect(color('url(#a)')).toBeUndefined()
    expect(color('expression(alert(1))')).toBeUndefined()
  })

  test('text widths: wider text, wider boxes; marks take no room', () => {
    expect(textWidth('WWW')).toBeGreaterThan(textWidth('iii'))
    expect(textWidth('Tiếng Việt')).toBeLessThan(textWidth('Tiếng Việt') + 1)
    expect(textWidth('é')).toBe(textWidth('e'))
    expect(textWidth('漢字')).toBe(28)
  })
})

// ---------------------------------------------------------------- what the readers understand

describe('Mermaid', () => {
  test('flowchart edges: labels, chains, & lists, lengths, kinds', () => {
    const g = graphOf(parseMermaid('graph LR\n  A -- text --> B\n  B -->|pipe| C --> D\n  A & B --> E & F\n  E -.-> F\n  F ==> G\n  G ---> H\n  H --o I\n  I --x J\n  J <--> K\n  K ~~~ L'))
    expect(g.dir).toBe('LR')
    expect(edgeList(g)).toEqual(['A>B', 'B>C', 'C>D', 'A>E', 'A>F', 'B>E', 'B>F', 'E>F', 'F>G', 'G>H', 'H>I', 'I>J', 'J>K', 'K>L'])
    const by = (k: string) => g.edges[edgeList(g).indexOf(k)]
    expect(by('A>B').label).toBe('text')
    expect(by('B>C').label).toBe('pipe')
    expect(by('E>F').line).toBe('dashed')
    expect(by('F>G').line).toBe('thick')
    expect(by('G>H').minlen).toBe(2)
    expect(by('C>D').minlen).toBe(1)
    expect(by('H>I').head).toBe('circle')
    expect(by('I>J').head).toBe('cross')
    expect(by('J>K')).toMatchObject({ head: 'arrow', tail: 'arrow' })
    expect(by('K>L').line).toBe('hidden')
  })

  test('flowchart nodes: shapes, quoted text, entities, <br>, markdown strings', () => {
    const g = graphOf(parseMermaid('flowchart TD\n  a([stadium]) --> b[(db)]\n  b --> c{{hex}}\n  c --> d[/lean/]\n  d --> e[\\trap/]\n  e --> f>flag]\n  f --> g["quoted (with) [brackets]"]\n  g --> h["A #quot;B#quot;<br/>C"]\n  h --> i["`**bold** text`"]\n  i --> node-1(((double)))'))
    const n = Object.fromEntries(g.nodes.map((x) => [x.id, x]))
    expect(n.a.shape).toBe('stadium')
    expect(n.b.shape).toBe('cylinder')
    expect(n.c.shape).toBe('hexagon')
    expect(n.d.shape).toBe('lean-r')
    expect(n.e.shape).toBe('trap-t')
    expect(n.f.shape).toBe('flag')
    expect(n.g.label).toBe('quoted (with) [brackets]')
    expect(n.h.label).toBe('A "B"\nC')
    expect(n.i.label).toBe('bold text')
    expect(n['node-1'].shape).toBe('doublecircle')
  })

  test('subgraphs, styles and classes', () => {
    const g = graphOf(parseMermaid('flowchart TB\n  x --> a\n  subgraph one [First]\n    a --> b\n    subgraph inner\n      c\n    end\n  end\n  classDef hot fill:#f96,stroke:#333,stroke-width:3px\n  class a hot\n  b:::hot\n  style c fill:red\n  linkStyle 0 stroke:#00f'))
    expect(g.clusters).toEqual([
      { id: 'one', label: 'First', parent: undefined },
      { id: 'inner', label: 'inner', parent: 'one' },
    ])
    const n = Object.fromEntries(g.nodes.map((x) => [x.id, x]))
    expect(n.x.cluster).toBeUndefined()
    expect(n.a.cluster).toBe('one')
    expect(n.c.cluster).toBe('inner')
    expect(n.a.style).toMatchObject({ fill: '#f96', stroke: '#333', sw: 3 })
    expect(n.b.style?.fill).toBe('#f96')
    expect(n.c.style?.fill).toBe('red')
    expect(g.edges[0].color).toBe('#00f')
  })

  test('front matter title, comments and directives', () => {
    const g = graphOf(parseMermaid('---\ntitle: My chart\n---\n%%{init: {"theme": "dark"}}%%\n%% a comment\ngraph TD; A-->B; B-->C'))
    expect(g.title).toBe('My chart')
    expect(edgeList(g)).toEqual(['A>B', 'B>C'])
  })

  test('sequence: messages, activations, notes, blocks', () => {
    const m = parseMermaid('sequenceDiagram\n  participant A as Alice\n  actor B\n  A->>+B: call\n  B-->>-A: back\n  A-)B: async\n  A-xB: lost\n  A<<->>B: both\n  alt yes\n    A->B: a\n  else no\n    A-->B: b\n  end\n  box Aqua Group\n  participant C\n  end\n  Note left of A: n')
    if (m.t !== 'sequence') throw new Error()
    const s = m.seq
    expect(s.participants.map((p) => [p.id, p.label, p.kind])).toEqual([
      ['A', 'Alice', 'participant'],
      ['B', 'B', 'actor'],
      ['C', 'C', 'participant'],
    ])
    const msgs = s.events.filter((e) => e.t === 'msg')
    expect(msgs.map((e) => e.t === 'msg' && [e.text, e.dashed, e.head])).toEqual([
      ['call', false, 'filled'],
      ['back', true, 'filled'],
      ['async', false, 'open'],
      ['lost', false, 'cross'],
      ['both', false, 'filled'],
      ['a', false, 'none'],
      ['b', true, 'none'],
    ])
    expect(msgs[0]).toMatchObject({ activate: true })
    expect(msgs[1]).toMatchObject({ deactivate: true })
    expect(s.events.map((e) => e.t).filter((t) => t !== 'msg')).toEqual(['block', 'else', 'end', 'note'])
  })

  test('class diagram: members, generics, relations with cardinality', () => {
    const g = graphOf(parseMermaid('classDiagram\n  class Shape~T~ {\n    <<interface>>\n    +area() double\n    +name String\n  }\n  Shape <|-- Circle\n  Circle "1" *-- "many" Point : has\n  Circle ..> Util'))
    const shape = g.nodes.find((n) => n.id === 'Shape')!
    expect(shape.sections!.map((s) => s.rows.map((r) => r.join(' ')))).toEqual([['«interface»', 'Shape<T>'], ['+name String'], ['+area() double']])
    expect(g.edges.map((e) => [e.from, e.to, e.tail, e.head, e.line, e.label, e.tailLabel, e.headLabel])).toEqual([
      ['Shape', 'Circle', 'triangle', 'none', 'solid', undefined, undefined, undefined],
      ['Circle', 'Point', 'diamond', 'none', 'solid', 'has', '1', 'many'],
      ['Circle', 'Util', 'none', 'open', 'dashed', undefined, undefined, undefined],
    ])
  })

  test('ER cardinalities', () => {
    const g = graphOf(parseMermaid('erDiagram\n  A ||--o{ B : x\n  B }|..|{ C : "y z"\n  C |o--o| D : w\n  D one or more to zero or many E : v'))
    expect(g.edges.map((e) => [e.tail, e.head, e.line, e.label])).toEqual([
      ['one', 'zero-many', 'solid', 'x'],
      ['one-many', 'one-many', 'dashed', 'y z'],
      ['zero-one', 'zero-one', 'solid', 'w'],
      ['one-many', 'zero-many', 'solid', 'v'],
    ])
  })

  test('gantt dates: formats, durations, dependencies', () => {
    expect(parseDate('2024-03-05', 'YYYY-MM-DD')).toBe(Date.UTC(2024, 2, 5))
    expect(parseDate('05/03/24 14:30', 'DD/MM/YY HH:mm')).toBe(Date.UTC(2024, 2, 5, 14, 30))
    expect(parseDate('1700000000', 'X')).toBe(1700000000000)
    const m = parseMermaid('gantt\n  dateFormat YYYY-MM-DD\n  section S\n  A :a, 2024-01-01, 3d\n  B :b, after a, 1w\n  C :crit, 2024-01-20, 2024-01-25\n  M :milestone, m, after b, 0d')
    if (m.t !== 'gantt') throw new Error()
    const day = 864e5
    expect(m.gantt.sections[0].tasks.map((t) => [(t.start - Date.UTC(2024, 0, 1)) / day, (t.end - t.start) / day, !!t.crit, !!t.milestone])).toEqual([
      [0, 3, false, false],
      [3, 7, false, false],
      [19, 5, true, false],
      [10, 0, false, true],
    ])
  })
})

describe('PlantUML', () => {
  test('blocks, comments, skinparams and titles', () => {
    const [m] = parsePlantUml("@startuml\n' comment\n/' block\ncomment '/\nskinparam monochrome true\nskinparam sequence {\n  ArrowColor red\n}\ntitle Hello\nA -> B : hi\n@enduml")
    if (m.t !== 'sequence') throw new Error(m.t)
    expect(m.seq.title).toBe('Hello')
    expect(m.seq.participants.map((p) => p.id)).toEqual(['A', 'B'])
  })

  test('sequence arrows and return', () => {
    const [m] = parsePlantUml('@startuml\nparticipant "Long Name" as L\nA -> L ++ : call\nL --> A : dashed\nA ->> L : thin\nA <- L : left\nA -> L : x\nreturn done\nA ->x L : lost\n== Part ==\n...\n|||\n@enduml')
    if (m.t !== 'sequence') throw new Error(m.t)
    expect(m.seq.participants.map((p) => [p.id, p.label])).toEqual([
      ['L', 'Long Name'],
      ['A', 'A'],
    ])
    const msgs = m.seq.events.flatMap((e) => (e.t === 'msg' ? [[e.from, e.to, e.text, e.dashed, e.head]] : []))
    expect(msgs).toEqual([
      ['A', 'L', 'call', false, 'filled'],
      ['L', 'A', 'dashed', true, 'filled'],
      ['A', 'L', 'thin', false, 'open'],
      ['L', 'A', 'left', false, 'filled'],
      ['A', 'L', 'x', false, 'filled'],
      ['L', 'A', 'done', true, 'filled'],
      ['A', 'L', 'lost', false, 'cross'],
    ])
    expect(m.seq.events.map((e) => e.t).filter((t) => t !== 'msg')).toEqual(['divider', 'delay', 'space'])
  })

  test('class diagram: bodies, relations, packages', () => {
    const g = graphOf(parsePlantUml('@startuml\npackage shop {\n  abstract class Item {\n    - price : int\n    + total() : int\n  }\n  enum Kind {\n    A\n    B\n  }\n}\nItem <|-- Book\nItem "1" o-- "*" Kind : has >\nBook ..> Kind\n@enduml')[0])
    expect(g.clusters.map((c) => c.id)).toEqual(['shop'])
    const item = g.nodes.find((n) => n.id === 'Item')!
    expect(item.cluster).toBe('shop')
    expect(item.sections!.map((s) => s.rows.map((r) => r[0]))).toEqual([['«abstract»', 'Item'], ['- price : int', '+ total() : int']])
    expect(g.edges.map((e) => [e.from, e.to, e.tail, e.head, e.line, e.label, e.tailLabel, e.headLabel])).toEqual([
      ['Item', 'Book', 'triangle', 'none', 'solid', undefined, undefined, undefined],
      ['Item', 'Kind', 'odiamond', 'none', 'solid', 'has', '1', '*'],
      ['Book', 'Kind', 'none', 'arrow', 'dashed', undefined, undefined, undefined],
    ])
  })

  test('activity: branches, loops and forks are joined up', () => {
    const g = graphOf(parsePlantUml('@startuml\nstart\n:a;\nif (c?) then (y)\n  :b;\nelseif (d?) then (z)\n  :e;\nelse (n)\n  :f;\nendif\nwhile (more?)\n  :g;\nendwhile\nrepeat\n  :h;\nrepeat while (again?) is (yes)\nfork\n  :i;\nfork again\n  :j;\nend fork\nstop\n@enduml')[0])
    const label = (id: string) => g.nodes.find((n) => n.id === id)!.label || g.nodes.find((n) => n.id === id)!.shape
    const pairs = g.edges.map((e) => `${label(e.from)}>${label(e.to)}${e.label ? `(${e.label})` : ''}`)
    for (const p of ['start>a', 'a>c?', 'c?>b(y)', 'c?>d?', 'd?>e(z)', 'd?>f(n)', 'b>more?', 'e>more?', 'f>more?', 'more?>g', 'g>more?', 'h>again?', 'i>bar', 'j>bar', 'bar>end']) expect(pairs).toContain(p)
  })

  test('use case and component elements', () => {
    const g = graphOf(parsePlantUml('@startuml\nactor :Shopper: as S\nusecase "Pay now" as UC\nS --> UC\n[Front end] -up-> [API]\nnode "Box" as N\ncloud Net\nN --> Net\n@enduml')[0])
    const n = Object.fromEntries(g.nodes.map((x) => [x.id, x]))
    expect(n.S).toMatchObject({ shape: 'actor', label: 'Shopper' })
    expect(n.UC).toMatchObject({ shape: 'ellipse', label: 'Pay now' })
    expect(n['Front end'].shape).toBe('component')
    expect(n.N.shape).toBe('box3d')
    expect(n.Net.shape).toBe('cloud')
    // "up": laid out from API to Front end, the arrow still at API
    expect(g.edges[1]).toMatchObject({ from: 'API', to: 'Front end', tail: 'arrow', head: 'none' })
  })
})

describe('Graphviz', () => {
  test('defaults, chains, subgraph ends, clusters, rank=same, attributes', () => {
    const [g] = parseDot(`
      # preprocessor line
      strict digraph "G" {
        graph [rankdir=LR, label="Title"]
        node [shape=box color=blue]
        edge [style=dashed]
        a -> b -> c [label="x"];
        {d e} -> f
        subgraph cluster_one { label = "One"; style=filled; fillcolor="#eeeeee"; g; h }
        { rank = same; b; f }
        a:port1 -> g:n [dir=back, arrowtail=odiamond, color="0.000 1.000 1.000"]
        i [label="{A|<p>B|C}", shape=record]
        j [label=<<b>bold</b><br/>two>, shape=plaintext]
        k [label="\\N line\\lnext"]
        /* comment */ l [style=invis] // trailing
      }`)
    expect(g.dir).toBe('LR')
    expect(g.title).toBe('Title')
    expect(edgeList(g)).toEqual(['a>b', 'b>c', 'd>f', 'e>f', 'a>g'])
    expect(g.edges[0]).toMatchObject({ label: 'x', line: 'dashed', head: 'arrow', tail: 'none' })
    expect(g.edges[4]).toMatchObject({ head: 'none', tail: 'odiamond', color: '#ff0000' })
    const n = Object.fromEntries(g.nodes.map((x) => [x.id, x]))
    expect(n.a).toMatchObject({ shape: 'rect', style: { stroke: 'blue' } })
    expect(n.g.cluster).toBe('cluster_one')
    expect(g.clusters[0]).toMatchObject({ label: 'One', style: { fill: '#eeeeee' } })
    expect(g.sameRank).toEqual([['b', 'f']])
    expect(n.i.sections!.map((s) => s.rows[0][0])).toEqual(['A', 'B', 'C'])
    expect(n.j.label).toBe('bold\ntwo')
    expect(n.k.label).toBe('k line\nnext')
    expect(n.l.style?.hidden).toBe(true)
  })

  test('undirected graphs have no arrowheads; gray levels and hex colors', () => {
    const [g] = parseDot('graph { a -- b [color=gray50]; b -- c [color="#00ff00"] }')
    expect(g.edges.map((e) => [e.head, e.tail, e.color])).toEqual([
      ['none', 'none', '#808080'],
      ['none', 'none', '#00ff00'],
    ])
  })
})

// ---------------------------------------------------------------- the layout keeps things apart

function rng(seed: number) {
  return () => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff
    return seed / 0x7fffffff
  }
}

describe('layered layout', () => {
  test('random graphs with clusters: nodes never overlap, cluster boxes hold their nodes and no others', () => {
    for (let run = 0; run < 40; run++) {
      const r = rng(run + 1)
      const nc = Math.floor(r() * 4)
      const clusters: LCluster[] = []
      for (let c = 0; c < nc; c++) clusters.push({ parent: c > 0 && r() < 0.3 ? Math.floor(r() * c) : -1, pad: [10, 10, 24, 10], minW: 40 })
      const N = 3 + Math.floor(r() * 18)
      const nodes: LNode[] = Array.from({ length: N }, () => ({ w: 30 + r() * 90, h: 20 + r() * 40, cluster: nc && r() < 0.6 ? Math.floor(r() * nc) : -1 }))
      const edges: LEdge[] = []
      const E = Math.floor(N * (0.8 + r()))
      for (let k = 0; k < E; k++) edges.push({ v: Math.floor(r() * N), w: Math.floor(r() * N), minlen: 1 + (r() < 0.2 ? 1 : 0), weight: 1, label: r() < 0.2 ? { w: 40, h: 16 } : undefined })
      const out = layered(nodes, edges, clusters, { nodesep: 30, ranksep: 40, edgesep: 10 })
      const box = (k: number) => ({ x0: out.nodes[k].x - nodes[k].w / 2, x1: out.nodes[k].x + nodes[k].w / 2, y0: out.nodes[k].y - nodes[k].h / 2, y1: out.nodes[k].y + nodes[k].h / 2 })
      const overlap = (a: ReturnType<typeof box>, b: ReturnType<typeof box>) => a.x0 < b.x1 - 0.5 && b.x0 < a.x1 - 0.5 && a.y0 < b.y1 - 0.5 && b.y0 < a.y1 - 0.5
      for (let a = 0; a < N; a++) {
        expect(Number.isFinite(out.nodes[a].x) && Number.isFinite(out.nodes[a].y)).toBe(true)
        for (let b = a + 1; b < N; b++) expect(overlap(box(a), box(b))).toBe(false)
      }
      const inside = (k: number, c: number) => {
        for (let x = nodes[k].cluster; x >= 0; x = clusters[x].parent) if (x === c) return true
        return false
      }
      clusters.forEach((_, c) => {
        const cb = out.clusters[c]
        if (!cb) return
        const b = { x0: cb.x, x1: cb.x + cb.w, y0: cb.y, y1: cb.y + cb.h }
        for (let k = 0; k < N; k++) {
          const nb = box(k)
          if (inside(k, c)) expect(nb.x0 >= b.x0 - 0.5 && nb.x1 <= b.x1 + 0.5 && nb.y0 >= b.y0 - 0.5 && nb.y1 <= b.y1 + 0.5).toBe(true)
          else expect(overlap(nb, b)).toBe(false)
        }
      })
      for (const e of out.edges) {
        if (!e) continue
        for (const s of e.segs) for (const p of s.p) expect(Number.isFinite(p.x) && Number.isFinite(p.y)).toBe(true)
      }
    }
  })

  test('edges point down the ranks; a cycle is kept, one edge drawn backwards', () => {
    const nodes: LNode[] = [0, 1, 2].map(() => ({ w: 40, h: 20, cluster: -1 }))
    const out = layered(nodes, [{ v: 0, w: 1, minlen: 1, weight: 1 }, { v: 1, w: 2, minlen: 1, weight: 1 }, { v: 2, w: 0, minlen: 1, weight: 1 }], [], { nodesep: 20, ranksep: 30, edgesep: 10 })
    expect(out.nodes[0].y).toBeLessThan(out.nodes[1].y)
    expect(out.nodes[1].y).toBeLessThan(out.nodes[2].y)
    // the back edge starts at node 2 (bottom) and ends at node 0
    const back = out.edges[2]!
    expect(back.start.y).toBeGreaterThan(out.nodes[0].y)
    const end = back.segs[back.segs.length - 1].p.at(-1)!
    expect(end.y).toBeLessThan(out.nodes[1].y)
  })
})

// ---------------------------------------------------------------- opening files

async function open(src: Blob | Uint8Array, name?: string) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, name ? { name } : {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const paras = blocks.filter((b): b is Paragraph => b.t === 'p')
  return { engine, info, paras }
}
const imgs = (ps: Paragraph[]) => ps.flatMap((p) => p.runs.filter((r) => r.t === 'img'))
const textOf = (ps: Paragraph[]) => ps.map((p) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join(''))

describe('diagram files open as documents', () => {
  test('format from the extension, the MIME type or the content', () => {
    const head = (s: string) => enc(s)
    expect(plainFormat('a.mmd', undefined, head('graph TD\nA-->B'), true)).toBe('mermaid')
    expect(plainFormat('a.mermaid', undefined, head('x'), true)).toBe('mermaid')
    expect(plainFormat('a.puml', undefined, head('@startuml'), true)).toBe('plantuml')
    expect(plainFormat('a.gv', undefined, head('digraph {}'), true)).toBe('graphviz')
    expect(plainFormat('a.dot', undefined, head('digraph {}'), true)).toBe('graphviz')
    // a Word 97 template (.dot) is an OLE file, not a graph
    expect(plainFormat('a.dot', undefined, new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), false)).toBeNull()
    expect(plainFormat(undefined, 'text/plain', head('sequenceDiagram\n A->>B: x'), true)).toBe('mermaid')
    expect(plainFormat(undefined, undefined, head('@startuml\nA -> B\n@enduml'), true)).toBe('plantuml')
    expect(plainFormat(undefined, undefined, head('digraph G {\n a -> b\n}'), true)).toBe('graphviz')
    expect(plainFormat(undefined, 'text/vnd.graphviz', head('graph { a -- b }'), true)).toBe('graphviz')
    // still text / CSV / Markdown
    expect(plainFormat(undefined, 'text/plain', head('graph theory notes'), true)).toBe('text')
    expect(plainFormat('notes.txt', undefined, head('graph TD\nA-->B'), true)).toBe('text')
    expect(sniffDiagram('# Heading\n\ngraph TD')).toBeNull()
    expect(sniffDiagram('pie is my favourite dessert')).toBeNull()
    expect(sniffDiagram('timeline of the project:\n- start')).toBeNull()
    expect(sniffDiagram('graph theory notes {\n}')).toBeNull()
    expect(sniffDiagram('pie showData title Pets\n "a": 1')).toBe('mermaid')
    expect(sniffDiagram('graph LR; A-->B')).toBe('mermaid')
    expect(sniffDiagram('strict digraph "my graph" {\n}')).toBe('graphviz')
    expect(sniffDiagram('graph G\n{\n a -- b\n}')).toBe('graphviz')
    expect(sniffDiagram('graph TD\n  A --> B')).toBe('mermaid')
    expect(sniffDiagram('%% comment\nflowchart LR\n A-->B')).toBe('mermaid')
    expect(sniffDiagram('/* c */\ndigraph {\n}')).toBe('graphviz')
  })

  test('a .mmd file: one picture, an SVG part, saved and reopened as Word', async () => {
    const { info, paras, engine } = await open(enc('flowchart LR\n  A[Hello] --> B[World]'), 'hello.mmd')
    expect(info).toMatchObject({ kind: 'docx', imported: 'mermaid', name: 'hello.mmd' })
    const [img] = imgs(paras)
    expect(img).toBeDefined()
    const part = (await engine.call('part', [(img as { target: string }).target])) as Blob
    expect(part.type).toBe('image/svg+xml')
    const svg = await part.text()
    expect(svg).toContain('>Hello</text>')
    const saved = (await engine.call('save', [])) as Blob
    const again = await open(saved)
    expect(again.info.imported).toBeUndefined()
    expect(imgs(again.paras)).toHaveLength(1)
  })

  test('wide drawings get a landscape page; big ones are scaled to fit', async () => {
    const chain = Array.from({ length: 14 }, (_, k) => `n${k} --> n${k + 1}`).join('\n')
    const { paras, engine } = await open(enc(`graph LR\n${chain}`), 'wide.mmd')
    const [img] = imgs(paras) as { w: number; h: number }[]
    // landscape text area: 9" = 648 pt
    expect(img.w).toBeLessThanOrEqual(648.5)
    const saved = (await engine.call('save', [])) as Blob
    const { ZipArchive } = await import('../src/core/zip')
    const { BlobSource } = await import('../src/core/source')
    const zip = await ZipArchive.open(new BlobSource(saved))
    expect(await zip.text('word/document.xml')).toContain('w:orient="landscape"')
  })

  test('a diagram that cannot be drawn: why, then its source', async () => {
    const { paras } = await open(enc('journey\n  title My day\n  section Go\n    Walk: 5: Me'), 'day.mmd')
    expect(imgs(paras)).toHaveLength(0)
    const t = textOf(paras)
    expect(t[0]).toBe('Mermaid diagram not drawn: Mermaid “journey” diagrams are not supported yet')
    expect(t[1]).toContain('title My day')
  })

  test('Markdown: diagram code blocks drawn, other code blocks kept as code', async () => {
    const md = '# Doc\n\nBefore\n\n```mermaid\ngraph TD\n  A --> B\n```\n\n```plantuml\n@startuml\nA -> B : hi\n@enduml\n```\n\n~~~dot\ndigraph { x -> y }\n~~~\n\n```js\nconst graph = 1\n```\n\nAfter'
    const { paras, info } = await open(enc(md), 'doc.md')
    expect(info.imported).toBe('markdown')
    expect(imgs(paras)).toHaveLength(3)
    const t = textOf(paras).filter(Boolean)
    expect(t).toEqual(['Doc', 'Before', 'const graph = 1', 'After'])
  })

  test('HTML: mermaid / plantuml / dot blocks of a web page drawn', async () => {
    const html =
      '<h1>Page</h1><p>Intro</p><pre class="mermaid">graph TD\n  A --&gt; B</pre>' +
      '<div class="diagram mermaid">sequenceDiagram\n  A-&gt;&gt;B: hi</div>' +
      '<pre><code class="language-dot">digraph { x -&gt; y }</code></pre>' +
      '<pre><code class="language-js">let a = 1</code></pre><p>End</p>'
    const { paras, info } = await open(enc(html), 'page.html')
    expect(info.imported).toBe('html')
    expect(imgs(paras)).toHaveLength(3)
    expect(textOf(paras).filter(Boolean)).toEqual(['Page', 'Intro', 'let a = 1', 'End'])
  })
})
