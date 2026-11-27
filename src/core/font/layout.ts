import { DAMAGED, FontError } from './container'

/**
 * What a font says about runs of its glyphs: the GSUB table (one glyph for another, ligatures), the GPOS table
 * (kerning) and the older kern table.
 *
 * Applied here: what plain left-to-right text needs. Single, multiple and ligature substitutions; single and pair
 * adjustments. The other kinds of lookup are read only as far as it takes to say whether they would do something to
 * a run: contextual ones are matched (and not applied), mark attachment and cursive joining are known by the glyphs
 * they start at. A run they would change is laid out without them and said not to be exact, so nothing passes for
 * exact that a shaping engine would have set differently.
 */

const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)

/** how many records a font may still make the reader hold (kerning pairs, ligatures, subtables) */
export interface Budget {
  left: number
}

function spend(budget: Budget, n: number) {
  budget.left -= n
  if (budget.left < 0) throw new FontError('The font file is too large: its layout tables are too long')
}

/** glyph → its index in the coverage, -1 when it is not covered */
export type Coverage = (glyph: number) => number

function readCoverage(v: DataView, at: number, budget: Budget): Coverage {
  const format = v.getUint16(at)
  const count = v.getUint16(at + 2)
  spend(budget, count)
  if (format === 1) {
    const glyphs = new Uint16Array(count)
    for (let i = 0; i < count; i++) glyphs[i] = v.getUint16(at + 4 + i * 2)
    return (g) => {
      let lo = 0
      let hi = count - 1
      while (lo <= hi) {
        const mid = (lo + hi) >> 1
        if (glyphs[mid] < g) lo = mid + 1
        else if (glyphs[mid] > g) hi = mid - 1
        else return mid
      }
      return -1
    }
  }
  if (format !== 2) throw new FontError(`${DAMAGED}: a coverage table of an unknown kind`)
  const start = new Uint16Array(count)
  const end = new Uint16Array(count)
  const first = new Uint16Array(count)
  for (let i = 0; i < count; i++) {
    start[i] = v.getUint16(at + 4 + i * 6)
    end[i] = v.getUint16(at + 6 + i * 6)
    first[i] = v.getUint16(at + 8 + i * 6)
  }
  return (g) => {
    let lo = 0
    let hi = count - 1
    while (lo <= hi) {
      const mid = (lo + hi) >> 1
      if (end[mid] < g) lo = mid + 1
      else if (start[mid] > g) hi = mid - 1
      else return first[mid] + g - start[mid]
    }
    return -1
  }
}

/** glyph → its class (0: in no class) */
function readClassDef(v: DataView, at: number, budget: Budget): (glyph: number) => number {
  const format = v.getUint16(at)
  if (format === 1) {
    const first = v.getUint16(at + 2)
    const count = v.getUint16(at + 4)
    spend(budget, count)
    const classes = new Uint16Array(count)
    for (let i = 0; i < count; i++) classes[i] = v.getUint16(at + 6 + i * 2)
    return (g) => classes[g - first] ?? 0
  }
  if (format !== 2) throw new FontError(`${DAMAGED}: a class table of an unknown kind`)
  const count = v.getUint16(at + 2)
  spend(budget, count)
  const start = new Uint16Array(count)
  const end = new Uint16Array(count)
  const cls = new Uint16Array(count)
  for (let i = 0; i < count; i++) {
    start[i] = v.getUint16(at + 4 + i * 6)
    end[i] = v.getUint16(at + 6 + i * 6)
    cls[i] = v.getUint16(at + 8 + i * 6)
  }
  return (g) => {
    let lo = 0
    let hi = count - 1
    while (lo <= hi) {
      const mid = (lo + hi) >> 1
      if (end[mid] < g) lo = mid + 1
      else if (start[mid] > g) hi = mid - 1
      else return cls[mid]
    }
    return 0
  }
}

/**
 * A lookup of a kind that is not applied here. `hits` says what it would do at glyph `i` of a run: 0 nothing, 1 a
 * rule matches that changes nothing (fonts write such rules as exceptions: the glyph is done with for this lookup),
 * 2 it would change the run.
 */
interface Other {
  kind: 'other'
  hits(ids: number[], i: number): 0 | 1 | 2
}

/** a lookup that would change every run that has a glyph of `cov` */
const startsAt = (cov: Coverage): Other => ({ kind: 'other', hits: (ids, i) => (cov(ids[i]) >= 0 ? 2 : 0) })
const NEVER: Other = { kind: 'other', hits: () => 0 }

type SubLookup =
  | { kind: 'single'; cov: Coverage; delta: number; subs: Uint16Array | null }
  | { kind: 'multiple'; cov: Coverage; seqs: number[][] }
  | { kind: 'ligature'; cov: Coverage; sets: { glyph: number; rest: number[] }[][] }
  | Other

type PosLookup =
  | { kind: 'single'; cov: Coverage; place: Int16Array; advance: Int16Array }
  | { kind: 'pairs'; cov: Coverage; second: boolean; sets: Map<number, [number, number, number, number]>[] }
  | { kind: 'classes'; cov: Coverage; second: boolean; class1: (g: number) => number; class2: (g: number) => number; count1: number; count2: number; v: DataView; at: number; size: number; f1: number; f2: number }
  | Other

interface Lookup<T> {
  /** lookup flag: which glyphs it passes over */
  flag: number
  subs: T[]
}

interface LayoutTable<T> {
  /** script tag → the features (indices) of its default language system */
  scripts: Map<string, number[]>
  features: { tag: string; lookups: number[] }[]
  lookups: Lookup<T>[]
}

export type Substitution = LayoutTable<SubLookup>
export type Positioning = LayoutTable<PosLookup>

/** the script list, the feature list and the lookups of a GSUB or GPOS table; `sub` reads one subtable of a lookup */
function readLayout<T>(b: Uint8Array, extension: number, budget: Budget, sub: (v: DataView, at: number, type: number) => T): LayoutTable<T> {
  const v = view(b)
  const scriptsAt = v.getUint16(4)
  const featuresAt = v.getUint16(6)
  const lookupsAt = v.getUint16(8)
  const scripts = new Map<string, number[]>()
  const tag = (at: number) => String.fromCharCode(v.getUint8(at), v.getUint8(at + 1), v.getUint8(at + 2), v.getUint8(at + 3))
  if (scriptsAt) {
    const n = v.getUint16(scriptsAt)
    spend(budget, n)
    for (let i = 0; i < n; i++) {
      const rec = scriptsAt + 2 + i * 6
      const script = scriptsAt + v.getUint16(rec + 4)
      const def = v.getUint16(script)
      const features: number[] = []
      if (def) {
        const lang = script + def
        const required = v.getUint16(lang + 2)
        if (required !== 0xffff) features.push(required)
        const count = v.getUint16(lang + 4)
        spend(budget, count)
        for (let k = 0; k < count; k++) features.push(v.getUint16(lang + 6 + k * 2))
      }
      scripts.set(tag(rec), features)
    }
  }
  const features: { tag: string; lookups: number[] }[] = []
  if (featuresAt) {
    const n = v.getUint16(featuresAt)
    spend(budget, n)
    for (let i = 0; i < n; i++) {
      const rec = featuresAt + 2 + i * 6
      const feature = featuresAt + v.getUint16(rec + 4)
      const count = v.getUint16(feature + 2)
      spend(budget, count)
      const lookups: number[] = []
      for (let k = 0; k < count; k++) lookups.push(v.getUint16(feature + 4 + k * 2))
      features.push({ tag: tag(rec), lookups })
    }
  }
  const lookups: Lookup<T>[] = []
  if (lookupsAt) {
    const n = v.getUint16(lookupsAt)
    spend(budget, n)
    for (let i = 0; i < n; i++) {
      const lookup = lookupsAt + v.getUint16(lookupsAt + 2 + i * 2)
      const type = v.getUint16(lookup)
      const flag = v.getUint16(lookup + 2)
      const count = v.getUint16(lookup + 4)
      spend(budget, count)
      const subs: T[] = []
      for (let k = 0; k < count; k++) {
        let at = lookup + v.getUint16(lookup + 6 + k * 2)
        let kind = type
        // an extension subtable: the real one, farther away than 16 bits reach
        if (kind === extension) {
          kind = v.getUint16(at + 2)
          at += v.getUint32(at + 4)
        }
        subs.push(sub(v, at, kind))
      }
      lookups.push({ flag, subs })
    }
  }
  return { scripts, features, lookups }
}

/** one rule of a contextual lookup: what has to be before, at and after a glyph (glyphs or classes), and whether it then does something */
interface Rule {
  /** before the first glyph, nearest first */
  back: number[]
  /** after the first glyph */
  input: number[]
  ahead: number[]
  acts: boolean
}

/**
 * A contextual lookup (GSUB 5 and 6, GPOS 7 and 8; `chained`: the second of each pair), read to be matched. The
 * lookups its rules call are not applied. Glyphs are compared one after the other: a lookup that passes over marks
 * is matched as if the run had none, which is the case of the runs laid out here.
 */
function readContext(v: DataView, at: number, chained: boolean, budget: Budget): Other {
  const format = v.getUint16(at)
  const list = (p: number, n: number) => {
    spend(budget, n)
    const out: number[] = []
    for (let i = 0; i < n; i++) out.push(v.getUint16(p + i * 2))
    return out
  }
  /** the rules of a rule set at `set` */
  const rules = (set: number): Rule[] => {
    const n = v.getUint16(set)
    spend(budget, n)
    const out: Rule[] = []
    for (let i = 0; i < n; i++) {
      let p = set + v.getUint16(set + 2 + i * 2)
      if (!chained) {
        const count = v.getUint16(p)
        out.push({ back: [], input: list(p + 4, Math.max(0, count - 1)), ahead: [], acts: v.getUint16(p + 2) > 0 })
        continue
      }
      const back = list(p + 2, v.getUint16(p))
      p += 2 + back.length * 2
      const input = list(p + 2, Math.max(0, v.getUint16(p) - 1))
      p += 2 + input.length * 2
      const ahead = list(p + 2, v.getUint16(p))
      p += 2 + ahead.length * 2
      out.push({ back, input, ahead, acts: v.getUint16(p) > 0 })
    }
    return out
  }
  /** the rule sets at `p`: one for each first glyph (by coverage index) or class; an offset of 0 is no set */
  const sets = (p: number): (Rule[] | null)[] => {
    const n = v.getUint16(p)
    spend(budget, n)
    const out: (Rule[] | null)[] = []
    for (let i = 0; i < n; i++) {
      const o = v.getUint16(p + 2 + i * 2)
      out.push(o ? rules(at + o) : null)
    }
    return out
  }
  const same = (g: number) => g
  const matcher = (first: (g: number) => number, all: (Rule[] | null)[], back: (g: number) => number, input: (g: number) => number, ahead: (g: number) => number): Other['hits'] => {
    return (ids, i) => {
      const k = first(ids[i])
      if (k < 0) return 0
      for (const r of all[k] ?? []) {
        const end = i + r.input.length
        if (i - r.back.length < 0 || end + r.ahead.length >= ids.length) continue
        if (r.input.every((x, n) => input(ids[i + 1 + n]) === x) && r.back.every((x, n) => back(ids[i - 1 - n]) === x) && r.ahead.every((x, n) => ahead(ids[end + 1 + n]) === x)) return r.acts ? 2 : 1
      }
      return 0
    }
  }
  if (format === 1) {
    const cov = readCoverage(v, at + v.getUint16(at + 2), budget)
    return { kind: 'other', hits: matcher(cov, sets(at + 4), same, same, same) }
  }
  if (format === 2) {
    const cov = readCoverage(v, at + v.getUint16(at + 2), budget)
    const cls = (p: number) => (v.getUint16(p) ? readClassDef(v, at + v.getUint16(p), budget) : () => 0)
    const back = chained ? cls(at + 4) : same
    const input = cls(chained ? at + 6 : at + 4)
    const ahead = chained ? cls(at + 8) : same
    const all = sets(chained ? at + 10 : at + 6)
    // the set is that of the first glyph's class; a glyph out of the coverage is in no set
    return { kind: 'other', hits: matcher((g) => (cov(g) < 0 ? -1 : input(g)), all, back, input, ahead) }
  }
  if (format !== 3) throw new FontError(`${DAMAGED}: a contextual rule of an unknown kind`)
  // one coverage for each glyph of the sequence
  let p = at + 2
  const covs = (n: number, from: number) => {
    spend(budget, n)
    const out: Coverage[] = []
    for (let i = 0; i < n; i++) out.push(readCoverage(v, at + v.getUint16(from + i * 2), budget))
    return out
  }
  let back: Coverage[] = []
  let ahead: Coverage[] = []
  let input: Coverage[]
  let acts: boolean
  if (chained) {
    back = covs(v.getUint16(p), p + 2)
    p += 2 + back.length * 2
    input = covs(v.getUint16(p), p + 2)
    p += 2 + input.length * 2
    ahead = covs(v.getUint16(p), p + 2)
    p += 2 + ahead.length * 2
    acts = v.getUint16(p) > 0
  } else {
    input = covs(v.getUint16(p), p + 4)
    acts = v.getUint16(p + 2) > 0
  }
  if (!input.length) return NEVER
  return {
    kind: 'other',
    hits: (ids, i) => {
      const end = i + input.length - 1
      if (i - back.length < 0 || end + ahead.length >= ids.length) return 0
      return input.every((c, n) => c(ids[i + n]) >= 0) && back.every((c, n) => c(ids[i - 1 - n]) >= 0) && ahead.every((c, n) => c(ids[end + 1 + n]) >= 0) ? (acts ? 2 : 1) : 0
    },
  }
}

export function readGsub(b: Uint8Array, budget: Budget): Substitution {
  return readLayout<SubLookup>(b, 7, budget, (v, at, type) => {
    const cov = () => readCoverage(v, at + v.getUint16(at + 2), budget)
    if (type === 1) {
      if (v.getUint16(at) === 1) return { kind: 'single', cov: cov(), delta: v.getInt16(at + 4), subs: null }
      const n = v.getUint16(at + 4)
      spend(budget, n)
      const subs = new Uint16Array(n)
      for (let i = 0; i < n; i++) subs[i] = v.getUint16(at + 6 + i * 2)
      return { kind: 'single', cov: cov(), delta: 0, subs }
    }
    if (type === 2) {
      const n = v.getUint16(at + 4)
      const seqs: number[][] = []
      for (let i = 0; i < n; i++) {
        const s = at + v.getUint16(at + 6 + i * 2)
        const count = v.getUint16(s)
        spend(budget, count + 1)
        const seq: number[] = []
        for (let k = 0; k < count; k++) seq.push(v.getUint16(s + 2 + k * 2))
        seqs.push(seq)
      }
      return { kind: 'multiple', cov: cov(), seqs }
    }
    if (type === 4) {
      const n = v.getUint16(at + 4)
      const sets: { glyph: number; rest: number[] }[][] = []
      for (let i = 0; i < n; i++) {
        const set = at + v.getUint16(at + 6 + i * 2)
        const count = v.getUint16(set)
        const ligs: { glyph: number; rest: number[] }[] = []
        for (let k = 0; k < count; k++) {
          const lig = set + v.getUint16(set + 2 + k * 2)
          const comps = v.getUint16(lig + 2)
          spend(budget, comps + 1)
          const rest: number[] = []
          for (let c = 1; c < comps; c++) rest.push(v.getUint16(lig + 4 + (c - 1) * 2))
          ligs.push({ glyph: v.getUint16(lig), rest })
        }
        sets.push(ligs)
      }
      return { kind: 'ligature', cov: cov(), sets }
    }
    if (type === 5 || type === 6) return readContext(v, at, type === 6, budget)
    // alternates (3) are not applied to plain text by anyone; reverse chaining (8) is known by the glyphs it starts at
    if (type === 3) return NEVER
    if (type === 8) return startsAt(cov())
    throw new FontError(`${DAMAGED}: a substitution of an unknown kind`)
  })
}

const bits = (n: number) => {
  let c = 0
  for (; n; n &= n - 1) c++
  return c
}

/** the horizontal parts of a value record at `at`: [xPlacement, xAdvance] */
function value(v: DataView, at: number, format: number): [number, number] {
  let p = at
  let place = 0
  let advance = 0
  if (format & 1) {
    place = v.getInt16(p)
    p += 2
  }
  if (format & 2) p += 2
  if (format & 4) advance = v.getInt16(p)
  return [place, advance]
}

export function readGpos(b: Uint8Array, budget: Budget): Positioning {
  return readLayout<PosLookup>(b, 9, budget, (v, at, type) => {
    const cov = () => readCoverage(v, at + v.getUint16(at + 2), budget)
    const format = v.getUint16(at)
    if (type === 1) {
      const f = v.getUint16(at + 4)
      const n = format === 1 ? 1 : v.getUint16(at + 6)
      spend(budget, n)
      const place = new Int16Array(n)
      const advance = new Int16Array(n)
      const size = bits(f & 0xff) * 2
      for (let i = 0; i < n; i++) [place[i], advance[i]] = value(v, at + (format === 1 ? 6 : 8) + i * size, f)
      return { kind: 'single', cov: cov(), place, advance }
    }
    if (type === 2) {
      const f1 = v.getUint16(at + 4)
      const f2 = v.getUint16(at + 6)
      const s1 = bits(f1 & 0xff) * 2
      const s2 = bits(f2 & 0xff) * 2
      if (format === 1) {
        const n = v.getUint16(at + 8)
        const sets: Map<number, [number, number, number, number]>[] = []
        for (let i = 0; i < n; i++) {
          const set = at + v.getUint16(at + 10 + i * 2)
          const count = v.getUint16(set)
          spend(budget, count + 1)
          const pairs = new Map<number, [number, number, number, number]>()
          for (let k = 0; k < count; k++) {
            const rec = set + 2 + k * (2 + s1 + s2)
            const second = v.getUint16(rec)
            // (the first record of a glyph is the one that counts)
            if (!pairs.has(second)) pairs.set(second, [...value(v, rec + 2, f1), ...value(v, rec + 2 + s1, f2)])
          }
          sets.push(pairs)
        }
        return { kind: 'pairs', cov: cov(), second: f2 !== 0, sets }
      }
      if (format !== 2) throw new FontError(`${DAMAGED}: a pair adjustment of an unknown kind`)
      const count1 = v.getUint16(at + 12)
      const count2 = v.getUint16(at + 14)
      const size = s1 + s2
      // every record is read where it is, when its pair is asked for: all of them have to be in the table
      if (at + 16 + count1 * count2 * size > v.byteLength) throw new FontError(`${DAMAGED}: a pair adjustment ends early`)
      return {
        kind: 'classes',
        cov: cov(),
        second: f2 !== 0,
        class1: readClassDef(v, at + v.getUint16(at + 8), budget),
        class2: readClassDef(v, at + v.getUint16(at + 10), budget),
        count1,
        count2,
        v,
        at: at + 16,
        size,
        f1,
        f2,
      }
    }
    if (type === 7 || type === 8) return readContext(v, at, type === 8, budget)
    // cursive joining (3), marks on a base, a ligature or another mark (4, 5, 6): known by the glyphs they start at
    if (type >= 3 && type <= 6) return startsAt(cov())
    throw new FontError(`${DAMAGED}: an adjustment of an unknown kind`)
  })
}

/**
 * The kern table: pairs (left glyph × 65536 + right glyph) → how much closer or farther the two are set.
 * `unread`: it has subtables of a kind that is not read (so their kerning is missing).
 */
export function readKern(b: Uint8Array, budget: Budget): { pairs: Map<number, number>; unread: boolean } {
  const v = view(b)
  const pairs = new Map<number, number>()
  let unread = false
  // version 0 (Microsoft, OpenType) has 16-bit counts and lengths, version 1.0 (Apple) 32-bit ones
  const apple = v.getUint16(0) === 1
  const tables = apple ? v.getUint32(4) : v.getUint16(2)
  let at = apple ? 8 : 4
  for (let t = 0; t < tables && at + (apple ? 8 : 6) <= b.length; t++) {
    const length = apple ? v.getUint32(at) : v.getUint16(at + 2)
    const coverage = apple ? v.getUint16(at + 4) : v.getUint16(at + 4)
    const format = apple ? coverage & 0xff : coverage >> 8
    // horizontal, not a minimum, not across the line
    const plain = apple ? !(coverage & 0xe000) : (coverage & 0x07) === 1
    const override = !apple && !!(coverage & 0x08)
    const head = apple ? 8 : 6
    if (format === 0 && plain) {
      const n = v.getUint16(at + head)
      // (a table of more pairs than its 16-bit length can say gives a wrong length: the pair count is what is true)
      const count = Math.min(n, Math.floor((b.length - at - head - 8) / 6))
      spend(budget, count)
      for (let i = 0; i < count; i++) {
        const rec = at + head + 8 + i * 6
        const key = v.getUint16(rec) * 65536 + v.getUint16(rec + 2)
        const val = v.getInt16(rec + 4)
        pairs.set(key, override ? val : (pairs.get(key) ?? 0) + val)
      }
    } else unread = true
    if (!length) break
    at += length
  }
  return { pairs, unread }
}

/** the lookups (their indices, in the order they are applied) of the features `tags` for `script` */
function lookupsOf<T>(table: LayoutTable<T>, script: string, tags: ReadonlySet<string>): number[] {
  // the script of the text, else the font's default one, else Latin: as shaping engines choose
  const features = table.scripts.get(script) ?? table.scripts.get('DFLT') ?? table.scripts.get('dflt') ?? table.scripts.get('latn')
  if (!features) return []
  const out = new Set<number>()
  for (const f of features) {
    const feature = table.features[f]
    if (feature && tags.has(feature.tag)) for (const l of feature.lookups) if (l < table.lookups.length) out.add(l)
  }
  return [...out].sort((a, b) => a - b)
}

/** the substitutions a browser makes in plain text without being asked */
const GSUB_ON = new Set(['ccmp', 'locl', 'rlig', 'calt', 'clig', 'liga', 'rclt'])
const GSUB_NO_LIGA = new Set(['ccmp', 'locl', 'rlig', 'calt', 'rclt'])
/** …and the adjustments */
const GPOS_ON = new Set(['kern', 'mark', 'mkmk', 'curs', 'dist', 'abvm', 'blwm'])
const GPOS_NO_KERN = new Set(['mark', 'mkmk', 'curs', 'dist', 'abvm', 'blwm'])
const KERN = new Set(['kern'])

/** lookup flags that pass over base glyphs or ligatures: matching then needs the glyph classes, which are not read */
const SKIPS = 0x0006

/**
 * Applies the font's own substitutions to a run of glyphs, in place: `ids` and, for each glyph, the index of the
 * character it starts at (`clusters`). Returns false when the run has glyphs a lookup that is not applied could
 * change: the result is then not what a shaping engine gives.
 */
export function substitute(table: Substitution, script: string, ids: number[], clusters: number[], ligatures: boolean): boolean {
  let exact = true
  for (const l of lookupsOf(table, script, ligatures ? GSUB_ON : GSUB_NO_LIGA)) {
    const lookup = table.lookups[l]
    for (let i = 0; i < ids.length; i++) {
      for (const s of lookup.subs) {
        if (s.kind === 'other' || lookup.flag & SKIPS) {
          const hit = s.kind === 'other' ? s.hits(ids, i) : s.cov(ids[i]) >= 0 ? 2 : 0
          if (hit === 2) exact = false
          if (hit) break
          continue
        }
        const c = s.cov(ids[i])
        if (c < 0) continue
        if (s.kind === 'single') {
          ids[i] = s.subs ? (s.subs[c] ?? ids[i]) : (ids[i] + s.delta) & 0xffff
          break
        }
        if (s.kind === 'multiple') {
          const seq = s.seqs[c]
          if (!seq?.length) continue
          ids.splice(i, 1, ...seq)
          clusters.splice(i, 1, ...seq.map(() => clusters[i]))
          i += seq.length - 1
          break
        }
        const lig = s.sets[c]?.find((x) => i + x.rest.length < ids.length && x.rest.every((g, k) => ids[i + 1 + k] === g))
        if (lig) {
          ids.splice(i, lig.rest.length + 1, lig.glyph)
          clusters.splice(i + 1, lig.rest.length)
          break
        }
      }
    }
  }
  return exact
}

/**
 * The font's own adjustments to a run of glyphs: added to `advances` (how far the pen moves after each glyph) and
 * `offsets` (how far each glyph is drawn from the pen). `kerning` false leaves pair kerning out, as text set with
 * `font-kerning: none` or in Word without "Kerning for fonts". Returns false when the run has glyphs a lookup that is
 * not applied could move.
 */
export function position(table: Positioning | null, kern: Map<number, number> | null, script: string, ids: number[], advances: number[], offsets: number[], kerning: boolean): boolean {
  let exact = true
  // the kern table is the font's kerning only when its GPOS table has none for the script
  const gposKerns = !!table && lookupsOf(table, script, KERN).length > 0
  if (table) {
    for (const l of lookupsOf(table, script, kerning ? GPOS_ON : GPOS_NO_KERN)) {
      const lookup = table.lookups[l]
      for (let i = 0; i < ids.length; ) {
        let next = i + 1
        for (const s of lookup.subs) {
          if (s.kind === 'other' || lookup.flag & SKIPS) {
            const hit = s.kind === 'other' ? s.hits(ids, i) : s.cov(ids[i]) >= 0 ? 2 : 0
            if (hit === 2) exact = false
            if (hit) break
            continue
          }
          const c = s.cov(ids[i])
          if (c < 0) continue
          if (s.kind === 'single') {
            const k = s.place.length === 1 ? 0 : c
            offsets[i] += s.place[k] ?? 0
            advances[i] += s.advance[k] ?? 0
            break
          }
          if (i + 1 >= ids.length) continue
          let rec: [number, number, number, number] | undefined
          if (s.kind === 'pairs') rec = s.sets[c]?.get(ids[i + 1])
          else {
            const c1 = s.class1(ids[i])
            const c2 = s.class2(ids[i + 1])
            if (c1 < s.count1 && c2 < s.count2) {
              const at = s.at + (c1 * s.count2 + c2) * s.size
              rec = [...value(s.v, at, s.f1), ...value(s.v, at + bits(s.f1 & 0xff) * 2, s.f2)]
            }
          }
          if (!rec) continue
          offsets[i] += rec[0]
          advances[i] += rec[1]
          offsets[i + 1] += rec[2]
          advances[i + 1] += rec[3]
          // a pair that also sets its second glyph is done with it: the next pair starts after it
          if (s.second) next = i + 2
          break
        }
        i = next
      }
    }
  }
  if (kerning && !gposKerns && kern) for (let i = 0; i + 1 < ids.length; i++) advances[i] += kern.get(ids[i] * 65536 + ids[i + 1]) ?? 0
  return exact
}

/** true when the font kerns text of `script` at all (so that a caller can tell "no kerning" from "kerning of 0") */
export function hasKerning(table: Positioning | null, kern: Map<number, number> | null, script: string): boolean {
  return (!!table && lookupsOf(table, script, KERN).length > 0) || (!!kern && kern.size > 0)
}
