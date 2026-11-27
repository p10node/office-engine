/**
 * Helvetica as every PDF reader has it: text no declared font file serves is set in it, in Windows-1252 codes (a
 * character outside them is a question mark). Widths are Adobe's, in 1/1000 em.
 */
import type { StdFont } from './types'

/** U+0020 to U+007E */
const REGULAR = [
  278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
  1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
  333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
]
const BOLD = [
  278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611,
  975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556,
  333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
]

/** Windows-1252 0x80-0x9F (0 where it has no character) */
const HIGH = [
  0x20ac, 0, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0, 0x17d, 0,
  0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0, 0x17e, 0x178,
]
const CODE = new Map<number, number>(HIGH.map((cp, i) => [cp, 0x80 + i] as [number, number]).filter(([cp]) => cp))

/** widths of the codes past ASCII that are not a letter with a mark (whose width is its letter's) */
const OTHER: Record<number, [number, number]> = {
  0x80: [556, 556], 0x82: [222, 278], 0x83: [556, 556], 0x84: [333, 500], 0x85: [1000, 1000], 0x86: [556, 556], 0x87: [556, 556], 0x88: [333, 333],
  0x89: [1000, 1000], 0x8b: [333, 333], 0x8c: [1000, 1000], 0x91: [222, 278], 0x92: [222, 278], 0x93: [333, 500], 0x94: [333, 500], 0x95: [350, 350],
  0x96: [556, 556], 0x97: [1000, 1000], 0x98: [333, 333], 0x99: [1000, 1000], 0x9b: [333, 333], 0x9c: [944, 944], 0xa0: [278, 278], 0xa1: [333, 333],
  0xa2: [556, 556], 0xa3: [556, 556], 0xa4: [556, 556], 0xa5: [556, 556], 0xa6: [260, 280], 0xa7: [556, 556], 0xa8: [333, 333], 0xa9: [737, 737],
  0xaa: [370, 370], 0xab: [556, 556], 0xac: [584, 584], 0xad: [333, 333], 0xae: [737, 737], 0xaf: [333, 333], 0xb0: [400, 400], 0xb1: [584, 584],
  0xb2: [333, 333], 0xb3: [333, 333], 0xb4: [333, 333], 0xb5: [556, 611], 0xb6: [537, 556], 0xb7: [278, 278], 0xb8: [333, 333], 0xb9: [333, 333],
  0xba: [365, 365], 0xbb: [556, 556], 0xbc: [834, 834], 0xbd: [834, 834], 0xbe: [834, 834], 0xbf: [611, 611], 0xc6: [1000, 1000], 0xd0: [722, 722],
  0xd7: [584, 584], 0xd8: [778, 778], 0xde: [667, 667], 0xdf: [611, 611], 0xe6: [889, 889], 0xf0: [556, 611], 0xf7: [584, 584], 0xf8: [611, 611],
  0xfe: [556, 611],
}

export const stdFont = (bold?: boolean, italic?: boolean): StdFont => (bold ? (italic ? 'Helvetica-BoldOblique' : 'Helvetica-Bold') : italic ? 'Helvetica-Oblique' : 'Helvetica')

/** the Windows-1252 code of a character; 0x3f (?) for one it has none for */
export function stdCode(cp: number): number {
  if ((cp >= 0x20 && cp <= 0x7e) || (cp >= 0xa0 && cp <= 0xff)) return cp
  return CODE.get(cp) ?? (cp === 0x2011 || cp === 0x2010 ? 0x2d : cp === 0x202f ? 0xa0 : 0x3f)
}

/** the character of a Windows-1252 code */
export const stdChar = (code: number): string => String.fromCharCode(code >= 0x80 && code < 0xa0 ? HIGH[code - 0x80] || 0x3f : code)

const widths = new Map<string, number>()

/** a code's width in 1/1000 em */
export function stdWidth(code: number, bold: boolean): number {
  if (code >= 0x20 && code <= 0x7e) return (bold ? BOLD : REGULAR)[code - 0x20]
  const other = OTHER[code]
  if (other) return other[bold ? 1 : 0]
  const key = `${bold ? 1 : 0}${code}`
  let w = widths.get(key)
  if (w === undefined) {
    // a letter with a mark is as wide as its letter
    const base = stdChar(code).normalize('NFD').charCodeAt(0)
    w = base >= 0x20 && base <= 0x7e ? (bold ? BOLD : REGULAR)[base - 0x20] : 556
    widths.set(key, w)
  }
  return w
}

export const isBold = (f: StdFont) => f.startsWith('Helvetica-Bold')
