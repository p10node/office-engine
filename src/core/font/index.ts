/**
 * Font files read without a browser: glyph widths, line metrics, kerning and ligatures, the file to use for a piece
 * of text among those a host declares, and subsets to embed. See docs/api.md, "Reading font files".
 */
export { FontError, FONT_LIMITS, fontFormat } from './container'
export { Font, type FontMetrics, type FontNames, type LineMetrics } from './font'
export { layoutText, needsShaping, type GlyphRun, type Inexact, type RunOptions, type ShapedGlyph, type Shaper } from './text'
export { FontBook, type FontBookOptions, type TextMeasure, type TextPiece, type TextSpec, type Unsure } from './book'
export { subsetFont, toUnicodeCMap, type FontSubset } from './subset'
