/**
 * Checks the font reader (src/core/font) on real font files against HarfBuzz, which browsers shape text with:
 *   bun scripts/font-check.ts /usr/share/fonts/truetype/liberation/*.ttf "/System/Library/Fonts/Helvetica.ttc#1"
 * (`file#n`: font n of a collection). Needs `hb-shape` and `hb-view` (HarfBuzz's tools: `brew install harfbuzz`,
 * `apt install libharfbuzz-bin`). No font ships with the repository; this is run by hand on the fonts at hand.
 *
 * - Runs: every run `layoutText` calls exact has HarfBuzz's glyphs at HarfBuzz's places (kerning and ligatures on
 *   and off). A run it calls not exact may differ: counted, with how many did.
 * - Subsets: a subset has the widths of its font, and draws the same picture (`hb-view`, with the font's own rules
 *   off, which a subset does not carry).
 */
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Font, layoutText, subsetFont } from '../src/core/font'
import { segments } from '../src/core/font/text'

const TEXTS = [
  'AVATAR To Ty. WAVE Yo, LT "quoted" 1/2 – (parentheses) [brackets]',
  'office affluent fjord ffi ffl fi fl ff shuffle The quick brown fox jumps over the lazy dog.',
  'Hamburgefonstiv 0123456789 $1,234.56 + 7% = x; a_b c@d.e #tag T. V. P. F,',
  'Tiếng Việt: Trường Đại học Quốc gia Hà Nội, ngày 23 tháng 11',
  'Größe, façade, naïve, Ærøskøbing, Łódź, Dvořák, İstanbul, Ål',
  'Съешь ещё этих мягких французских булок, да выпей чаю. ГДЕ ТУТ АУ',
  'Ξεσκεπάζω την ψυχοφθόρα βδελυγμία ΑΥΤΟ',
  '日本語のテキスト、漢字かな交じり文。東京都',
  '한국어 텍스트 서울특별시',
]
const OPTS = [{}, { kerning: false }, { ligatures: false }, { kerning: false, ligatures: false }]
const RULES_OFF = '--features=-kern,-liga,-clig,-calt,-ccmp,-locl,-rlig,-rclt,-mark,-mkmk,-dlig'

const files = process.argv.slice(2)
if (!files.length) {
  console.error('usage: bun scripts/font-check.ts <font file>[#index] …')
  process.exit(2)
}
for (const tool of ['hb-shape', 'hb-view']) {
  if (Bun.spawnSync(['which', tool]).exitCode !== 0) {
    console.error(`${tool} is not installed (HarfBuzz's tools)`)
    process.exit(2)
  }
}
const dir = mkdtempSync(join(tmpdir(), 'font-check-'))
const stats = { fonts: 0, unread: 0, runs: 0, runsDiffer: 0, notExact: 0, notExactDiffer: 0, noGlyphs: 0, subsets: 0, subsetsDiffer: 0, pictures: 0 }
const complain = (what: string) => {
  if (stats.runsDiffer + stats.subsetsDiffer <= 20) console.log(what)
}

for (const arg of files) {
  const [file, n] = arg.split('#')
  const index = Number(n ?? 0)
  const label = `${file.split('/').pop()}${index ? '#' + index : ''}`
  let font: Font
  try {
    font = await Font.open(new Uint8Array(readFileSync(file)), index)
  } catch (e) {
    stats.unread++
    console.log(`not read: ${label}: ${(e as Error).message}`)
    continue
  }
  stats.fonts++
  for (const whole of TEXTS) {
    // one script at a time, as hb-shape takes its text for one script
    for (const seg of segments(whole)) {
      const text = whole.slice(seg.from, seg.to)
      for (const o of OPTS) {
        const run = layoutText(font, text, o)
        if (run.missing) {
          stats.noGlyphs++
          continue
        }
        const off = [...(o.kerning === false ? ['-kern'] : []), ...(o.ligatures === false ? ['-liga', '-clig'] : [])]
        const p = Bun.spawnSync(['hb-shape', `--font-file=${file}`, `--face-index=${index}`, '--no-glyph-names', '--output-format=json', `--script=${seg.iso === 'Zyyy' ? 'Latn' : seg.iso}`, ...(off.length ? [`--features=${off.join(',')}`] : []), text])
        const hb = JSON.parse(p.stdout.toString()) as { g: number; dx: number; ax: number }[]
        // where each glyph is drawn: a kern table's value is split between two glyphs by HarfBuzz, the places are the same
        let pen = 0
        const theirs = hb.map((x) => {
          const at = pen + x.dx
          pen += x.ax
          return at
        })
        const width = pen
        pen = 0
        const mine = run.glyphs.map((_, i) => {
          const at = pen + run.dx[i]
          pen += run.advances[i]
          return at
        })
        const same = hb.length === run.glyphs.length && hb.every((x, i) => x.g === run.glyphs[i] && theirs[i] === mine[i]) && width === run.width
        if (run.exact) {
          stats.runs++
          if (!same) {
            stats.runsDiffer++
            const i = hb.findIndex((x, k) => x.g !== run.glyphs[k] || theirs[k] !== mine[k])
            complain(`RUN DIFFERS ${label} ${JSON.stringify(o)} "${text.slice(0, 30)}" at glyph ${i}: HarfBuzz ${hb[i]?.g} at ${theirs[i]}, here ${run.glyphs[i]} at ${mine[i]}; width ${width} / ${run.width}`)
          }
        } else {
          stats.notExact++
          if (!same) stats.notExactDiffer++
        }
      }
    }
  }
  for (const text of [TEXTS[2], TEXTS[5], TEXTS[7]]) {
    const run = layoutText(font, text, { kerning: false, ligatures: false })
    if (run.missing) continue
    stats.subsets++
    const chars: [number, number][] = [...text].map((c) => [c.codePointAt(0)!, font.glyphOf(c.codePointAt(0)!)])
    const errors: string[] = []
    try {
      const sub = subsetFont(font, [...run.glyphs, ...chars.map((c) => c[1])], chars)
      const small = await Font.open(sub.bytes)
      const width = run.glyphs.reduce((a, x) => a + small.advanceOf(sub.glyphs.get(x)!), 0)
      if (width !== run.width) errors.push(`width ${width}, the font's ${run.width}`)
      if (small.numGlyphs !== sub.glyphs.size) errors.push('glyph count')
      // the picture, when the text's glyphs are the characters' own and the font has no rules hb-view cannot be told to leave out
      const plain = run.glyphs.every((x, i) => x === font.glyphOf(text.codePointAt(run.clusters[i])!))
      if (plain && !font.aat) {
        stats.pictures++
        const out = join(dir, `subset.${sub.outlines === 'cff' ? 'otf' : 'ttf'}`)
        writeFileSync(out, sub.bytes)
        const view = (f: string, i: number) => Bun.spawnSync(['hb-view', `--font-file=${f}`, `--face-index=${i}`, '--font-size=48', RULES_OFF, '-O', 'png', text])
        const a = view(file, index)
        const b = view(out, 0)
        if (a.exitCode !== 0 || b.exitCode !== 0) errors.push(`hb-view failed: ${b.stderr.toString().slice(0, 100)}`)
        else if (Buffer.compare(a.stdout, b.stdout) !== 0) errors.push('the picture differs')
      }
    } catch (e) {
      errors.push((e as Error).message)
    }
    if (errors.length) {
      stats.subsetsDiffer++
      complain(`SUBSET DIFFERS ${label} "${text.slice(0, 20)}": ${errors.join('; ')}`)
    }
  }
}
rmSync(dir, { recursive: true, force: true })
console.log(stats)
process.exit(stats.runsDiffer || stats.subsetsDiffer ? 1 : 0)
