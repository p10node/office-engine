/**
 * Syntax colors for the source editor: Markdown (diagram code blocks in their own language), HTML, SVG, Mermaid,
 * PlantUML and Graphviz DOT. Line based, with the few states that span lines (code blocks, block comments).
 */
import type { PlainFormat } from '../../core/plain'

/** [class, text]; class '' is plain text. Classes: c comment, k keyword, s string, o arrow / operator, n number,
 * h heading, a attribute / link, t tag, m code, b bold, i italic */
export type Token = [string, string]

type Rule = [RegExp, string]
type Lang = 'markdown' | 'html' | 'mermaid' | 'plantuml' | 'graphviz'

const sticky = (rules: [string | RegExp, string][]): Rule[] => rules.map(([re, c]) => [new RegExp(typeof re === 'string' ? re : re.source, 'y' + (typeof re === 'string' ? '' : re.flags.replace(/[gy]/g, ''))), c])

const MERMAID_KW =
  'graph|flowchart|subgraph|end|direction|classDef|class|style|linkStyle|click|sequenceDiagram|participant|actor|loop|alt|else|opt|par|and|critical|option|break|rect|[Nn]ote|over|left of|right of|activate|deactivate|autonumber|title|classDiagram(?:-v2)?|stateDiagram(?:-v2)?|state|erDiagram|pie|showData|gantt|dateFormat|axisFormat|excludes|section|mindmap|namespace|TB|TD|BT|RL|LR|as|box|create|destroy|done|active|crit|milestone|after'

const RULES: Record<Lang, Rule[]> = {
  mermaid: sticky([
    [/%%.*/, 'c'],
    [/"[^"]*"/, 's'],
    [/\|[^|\n]*\|(?=\s*[\w"([{])/, 's'],
    [`\\b(?:${MERMAID_KW})\\b`, 'k'],
    [/<<-->>|<<->>|-->>|->>|--[x)]|-[x)]|<?-\.+->?|<?={2,}>?|<?-{2,}[>ox]?|->|~~~|<\|--|--\|>|\*--|--\*|o--|--o|\.\.>|\.\.\|>|<\.\.|\.\.|[|}][|o]--[|o][|{]|[|}][|o]\.\.[|o][|{]|:::|&/, 'o'],
    [/\b\d+(?:\.\d+)?\b/, 'n'],
    [/#[0-9a-fA-F]{3,8}\b/, 'n'],
  ]),
  plantuml: sticky([
    [/@(?:start|end)\w*/, 'k'],
    [/![a-z]+\b.*/, 'k'],
    [/"[^"]*"/, 's'],
    [/^\s*:[^;\n]*[;|<>/\]}]?/, 's'],
    [
      /\b(?:participant|actor|boundary|control|entity|database|collections|queue|class|interface|abstract|enum|annotation|package|namespace|node|component|usecase|rectangle|cloud|folder|frame|artifact|storage|card|file|object|state|start|stop|end|if|then|else|elseif|endif|while|endwhile|repeat|fork|again|merge|note|of|left|right|top|bottom|over|alt|opt|loop|par|break|critical|group|return|activate|deactivate|destroy|create|autonumber|title|skinparam|as|is|not|detach|kill|partition|split|switch|case|endswitch|hide|show|footbox|direction|to)\b/,
      'k',
    ],
    [/<<[^>\n]*>>/, 'a'],
    [/[<*o#x}|^+]{0,2}(?:-+|\.{2,}|=+)(?:\[[^\]\n]*\])?(?:up|down|left|right)?(?:-+|\.+|=+)?(?:\|>|>>?|[*o#x{|^+\\/]{1,2})|<\|?(?:-+|\.+)|--+|\.\.+|\+\+|\*\*|!!/, 'o'],
    [/#[0-9a-zA-Z]{3,8}\b/, 'n'],
    [/\b\d+\b/, 'n'],
  ]),
  graphviz: sticky([
    [/\/\/.*/, 'c'],
    [/"(?:\\.|[^"\\])*"/, 's'],
    [/<(?:[^<>]|<[^<>]*>)*>/, 's'],
    [/\b(?:strict|graph|digraph|subgraph|node|edge)\b/i, 'k'],
    [/[\w.]+(?=\s*=)/, 'a'],
    [/->|--/, 'o'],
    [/-?\b\d+(?:\.\d+)?\b/, 'n'],
  ]),
  html: sticky([
    [/<!\[CDATA\[.*?\]\]>/, 'c'],
    [/<!doctype[^>]*>/i, 'k'],
    [/<\/?[A-Za-z][\w:-]*/, 't'],
    [/\/?>/, 't'],
    [/[\w:-]+(?==)/, 'a'],
    [/"[^"\n]*"|'[^'\n]*'/, 's'],
    [/&(?:#\d+|#x[0-9a-fA-F]+|\w+);/, 'n'],
  ]),
  markdown: sticky([
    [/`[^`\n]+`/, 'm'],
    [/!?\[[^\]\n]*\]\([^)\n]*\)/, 'a'],
    [/<https?:\/\/[^>\s]+>/, 'a'],
    [/(\*\*|__)(?=\S)[^\n]*?\S\1/, 'b'],
    [/(?<![\w*])\*(?=\S)[^*\n]*?\S\*(?!\*)|(?<!\w)_(?=\S)[^_\n]*?\S_(?!\w)/, 'i'],
    [/~~(?=\S)[^\n]*?\S~~/, 'c'],
    [/<\/?[A-Za-z][^>\n]*>/, 't'],
  ]),
}

const FENCE: Record<string, Lang> = { mermaid: 'mermaid', mmd: 'mermaid', plantuml: 'plantuml', puml: 'plantuml', uml: 'plantuml', dot: 'graphviz', graphviz: 'graphviz', gv: 'graphviz', html: 'html' }

function scan(line: string, rules: Rule[], out: Token[]) {
  let plain = ''
  let i = 0
  while (i < line.length) {
    let hit = false
    for (const [re, c] of rules) {
      re.lastIndex = i
      const m = re.exec(line)
      if (m && m[0].length) {
        if (plain) out.push(['', plain])
        plain = ''
        out.push([c, m[0]])
        i += m[0].length
        hit = true
        break
      }
    }
    if (!hit) {
      // a word goes whole (keywords match at word starts only)
      const w = /^[\p{L}\p{N}_]+/u.exec(line.slice(i, i + 64))
      const step = w ? w[0].length : 1
      plain += line.slice(i, i + step)
      i += step
    }
  }
  if (plain) out.push(['', plain])
}

/** a part of `line` from `from` up to the end of a block `close` (or the end of the line); returns where it ended, -1 for still open */
function until(line: string, from: number, close: string, c: string, out: Token[], skip = 0): number {
  const e = line.indexOf(close, from + skip)
  if (e < 0) {
    out.push([c, line.slice(from)])
    return -1
  }
  out.push([c, line.slice(from, e + close.length)])
  return e + close.length
}

/** Tokens of the whole text, lines separated by "\n" tokens. Over `max` characters: no colors. */
export function highlight(format: PlainFormat, text: string, max = 300_000): Token[] {
  // SVG is XML: colored as HTML
  const lang = (format === 'svg' ? 'html' : format) as Lang
  if (text.length > max || !(lang in RULES)) return [['', text]]
  const out: Token[] = []
  let fence: { mark: string; lang?: Lang } | null = null
  let block: string | null = null
  const lines = text.split('\n')
  lines.forEach((line, k) => {
    if (k) out.push(['', '\n'])
    // markdown code blocks: the fence line, then the body in its own language
    if (lang === 'markdown') {
      const f = /^(\s*)(`{3,}|~{3,})\s*([\w-]*).*$/.exec(line)
      if (fence) {
        if (f && f[2][0] === fence.mark[0] && f[2].length >= fence.mark.length && !f[3]) {
          out.push(['k', line])
          fence = null
        } else if (fence.lang) scanLine(fence.lang, line)
        else out.push(['m', line])
        return
      }
      if (f) {
        fence = { mark: f[2], lang: FENCE[f[3].toLowerCase()] }
        out.push(['k', line])
        return
      }
      if (/^\s{0,3}#{1,6}(\s|$)/.test(line)) return void out.push(['h', line])
      if (/^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(line)) return void out.push(['o', line])
      const lead = /^(\s*(?:>\s*)*)((?:[-*+]|\d+[.)])\s+(?:\[[ xX]\]\s+)?)?/.exec(line)!
      if (lead[0]) {
        if (lead[1].includes('>')) out.push(['c', lead[1]])
        else if (lead[1]) out.push(['', lead[1]])
        if (lead[2]) out.push(['o', lead[2]])
      }
      scan(line.slice(lead[0].length), RULES.markdown, out)
      return
    }
    scanLine(lang, line)
  })
  return out

  function scanLine(l: Lang, line: string) {
    let i = 0
    // comments spanning lines
    if (block) {
      const e = until(line, 0, block, 'c', out)
      if (e < 0) return
      block = null
      i = e
    }
    const opens: [RegExp, string][] =
      l === 'html' ? [[/<!--/, '-->']] : l === 'graphviz' ? [[/\/\*/, '*/']] : l === 'plantuml' ? [[/\/'/, "'/"]] : []
    while (i <= line.length) {
      let at = -1
      let close = ''
      for (const [re, c] of opens) {
        const m = re.exec(line.slice(i))
        if (m && (at < 0 || m.index + i < at)) {
          at = m.index + i
          close = c
        }
      }
      const head = at < 0 ? line.slice(i) : line.slice(i, at)
      if (l === 'plantuml' && /^\s*'/.test(head) && i === 0) out.push(['c', head])
      else if (l === 'graphviz' && i === 0 && /^\s*#/.test(head)) out.push(['c', head])
      else scan(head, RULES[l], out)
      if (at < 0) return
      // the search for the end starts after the opening mark
      const e = until(line, at, close, 'c', out, close === '-->' ? 4 : 2)
      if (e < 0) {
        block = close
        return
      }
      i = e
      if (i === line.length) return
    }
  }
}
