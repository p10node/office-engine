import { useEffect, useState, type CSSProperties } from 'react'
import type { CanonStyle, StyleInfo } from '../../core/docx/types'
import { MOD } from '../common/ContextMenu'
import { EditorBar } from '../common/Slots'
import { Btn, ColorBtn, Dropdown, fontName, fontOptions, GridPicker, Icon, Menu, pickFiles, Pop, Select, Sep, SizeBox, type DropOption } from '../edit/Toolbar'
import type { BorderPreset } from './tables'
import type { DocxEditor, FormatState } from './DocxEditor'

const HIGHLIGHTS = ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#0000ff', '#ff0000', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0', '#000000']

const CELL_BORDERS: { k: BorderPreset; icon: string; title: string }[] = [
  { k: 'all', icon: 'bAll', title: 'All borders' },
  { k: 'inner', icon: 'bInner', title: 'Inner borders' },
  { k: 'outer', icon: 'bOuter', title: 'Outer borders' },
  { k: 'none', icon: 'bNone', title: 'No borders' },
  { k: 'top', icon: 'bTop', title: 'Top border' },
  { k: 'bottom', icon: 'bBottom', title: 'Bottom border' },
  { k: 'left', icon: 'bLeft', title: 'Left border' },
  { k: 'right', icon: 'bRight', title: 'Right border' },
]
/** cell border lines: value is the CSS width + style */
export const CELL_LINES = [
  { value: '0.5pt solid', label: 'Thin', hint: '───' },
  { value: '1.5pt solid', label: 'Medium', hint: '━━━' },
  { value: '3pt solid', label: 'Thick', hint: '▬▬▬' },
  { value: '0.5pt dashed', label: 'Dashed', hint: '- - -' },
  { value: '0.5pt dotted', label: 'Dotted', hint: '·····' },
  { value: '2.25pt double', label: 'Double', hint: '═══' },
]
const VALIGN = [
  { value: 'top', label: 'Top', icon: 'vTop' },
  { value: 'middle', label: 'Middle', icon: 'vMiddle' },
  { value: 'bottom', label: 'Bottom', icon: 'vBottom' },
]

const LINE_SPACING = [
  { value: 1, label: 'Single' },
  { value: 1.15, label: '1.15' },
  { value: 1.5, label: '1.5' },
  { value: 2, label: 'Double' },
  { value: 2.5, label: '2.5' },
  { value: 3, label: 'Triple' },
]

/** built-in styles always offered (created in the document on first use) */
const CANON: { k: CanonStyle; label: string; look: CSSProperties }[] = [
  { k: 'normal', label: 'Normal text', look: {} },
  { k: 'title', label: 'Title', look: { fontSize: 22 } },
  { k: 'subtitle', label: 'Subtitle', look: { fontSize: 15, color: '#5a5a5a' } },
  { k: 'h1', label: 'Heading 1', look: { fontSize: 18, color: '#2f5496' } },
  { k: 'h2', label: 'Heading 2', look: { fontSize: 16, color: '#2f5496' } },
  { k: 'h3', label: 'Heading 3', look: { fontSize: 14, color: '#1f3763' } },
  { k: 'h4', label: 'Heading 4', look: { fontSize: 13, color: '#2f5496', fontStyle: 'italic' } },
  { k: 'h5', label: 'Heading 5', look: { fontSize: 13, color: '#2f5496' } },
  { k: 'h6', label: 'Heading 6', look: { fontSize: 13, color: '#1f3763', fontStyle: 'italic' } },
]

const SYMBOLS = [
  '©', '®', '™', '§', '¶', '†', '‡', '•', '…', '–', '—', '‰',
  '€', '£', '¥', '¢', '₫', '$', '°', '±', '×', '÷', '≈', '≠',
  '≤', '≥', '∞', '√', '∑', '∏', '∫', 'π', 'µ', 'Ω', 'α', 'β',
  '←', '→', '↑', '↓', '↔', '⇒', '✓', '✗', '★', '☆', '♥', '♦',
  '«', '»', '“', '”', '‘', '’', '¼', '½', '¾', '¹', '²', '³',
]

/** the paragraph styles offered: the built-in ones first (in the document or not: `@kind` then), then the document's other quick styles */
export function styleChoices(styles: StyleInfo[]) {
  const out: { value: string; label: string; k?: CanonStyle; s?: StyleInfo; look?: CSSProperties; sep: boolean }[] = CANON.map(({ k, label, look }, i) => {
    const s = styles.find((x) => x.canon === k)
    return { value: s?.id ?? `@${k}`, label, k, s, look, sep: i === 1 }
  })
  styles.filter((s) => !s.canon).forEach((s, i) => out.push({ value: s.id, label: s.name, s, sep: i === 0 }))
  return out
}

/** how a style looks in the picker (size clamped to fit the menu) */
function stylePreview(s: StyleInfo | undefined): CSSProperties | undefined {
  const p = s?.preview
  if (!p) return undefined
  return {
    fontSize: p.sz ? Math.max(12, Math.min(22, p.sz)) : undefined,
    fontWeight: p.b ? 700 : undefined,
    fontStyle: p.i ? 'italic' : undefined,
    color: p.color && !/^#?f{6}$/i.test(p.color.replace('#', '')) ? p.color : undefined,
    fontFamily: p.font,
    lineHeight: 1.25,
  }
}

export interface ImageActions {
  original: () => void
  replace: () => void
  alt: () => void
  /** undefined when downloads are disabled */
  download?: () => void
}

export function DocxToolbar(props: {
  editor: DocxEditor
  styles: StyleInfo[]
  onFind: () => void
  outline: boolean
  onOutline: () => void
  painter: boolean
  onLink: () => void
  image: ImageActions
  /** opens the page setup dialog */
  onPageSetup?: () => void
}) {
  const { editor, styles, image } = props
  const [st, setSt] = useState<FormatState | null>(null)
  const [bColor, setBColor] = useState('#000000')
  const [bLine, setBLine] = useState('0.5pt solid')
  useEffect(() => {
    editor.onState = setSt
    return () => {
      editor.onState = undefined
    }
  }, [editor])
  const font = fontName(st?.font)
  const dis = !st
  const sz = st?.sz ? Math.round(st.sz * 2) / 2 : undefined
  const line = st?.line ? Math.round(st.line * 100) / 100 : ''

  // style picker: built-ins first (with shortcuts), then the document's other quick styles
  const alt = /Mac|iPhone|iPad/.test(typeof navigator !== 'undefined' ? navigator.platform : '') ? '⌥' : 'Alt+'
  const styleOpts: DropOption<string>[] = styleChoices(styles).map(({ value, label, k, s, look, sep }) => {
    const hint = k === 'normal' ? `${MOD}${alt}0` : k?.startsWith('h') ? `${MOD}${alt}${k.slice(1)}` : undefined
    return { value, label, hint, style: k ? { lineHeight: 1.25, ...(k === 'normal' ? {} : (stylePreview(s) ?? look)) } : stylePreview(s), sep }
  })
  const styleValue = st?.style && styleOpts.some((o) => o.value === st.style) ? st.style : st && !st.style ? (styles.find((x) => x.canon === 'normal')?.id ?? '') : ''

  return (
    <EditorBar>
      <Btn icon="undo" title={`Undo (${MOD}Z)`} disabled={!editor.canUndo} onClick={() => editor.undo()} />
      <Btn icon="redo" title={`Redo (${MOD}Y)`} disabled={!editor.canRedo} onClick={() => editor.redo()} />
      <Btn icon="paint" title="Paint format (click, then select text)" active={props.painter} disabled={dis} onClick={() => editor.copyFormat()} />
      {st?.inTable && (
        <>
          <Sep />
          <Menu
            title="Table"
            label="Table"
            icon="table"
            items={[
              { value: 'rowAbove', label: 'Insert row above', icon: 'rowAbove' },
              { value: 'rowBelow', label: 'Insert row below', icon: 'rowBelow' },
              { value: 'colLeft', label: 'Insert column left', icon: 'colLeft' },
              { value: 'colRight', label: 'Insert column right', icon: 'colRight' },
              { value: 'delRow', label: 'Delete row', icon: 'delRow', sep: true },
              { value: 'delCol', label: 'Delete column', icon: 'delCol' },
              { value: 'delTable', label: 'Delete table', icon: 'trash' },
              { value: 'distribute', label: 'Distribute columns evenly', icon: 'colAdd', sep: true },
            ]}
            onPick={(v) => (v === 'distribute' ? editor.distributeColumns() : editor.tableOp(v as Parameters<DocxEditor['tableOp']>[0]))}
          />
          <Btn
            icon="merge"
            title={st.cells?.merged && (st.cells?.n ?? 0) < 2 ? 'Unmerge cells' : 'Merge cells'}
            active={st.cells?.merged && (st.cells?.n ?? 0) < 2}
            disabled={!st.cells || (st.cells.n < 2 && !st.cells.merged)}
            onClick={() => ((st.cells?.n ?? 0) > 1 ? editor.mergeCells() : editor.unmergeCell())}
          />
          <ColorBtn icon="fill" title="Cell color" value={st.cells?.shd} onPick={(c) => editor.setCellShade(c)} noneLabel="No fill" />
          <Pop icon="borders" title="Cell borders" chevron>
            {(close) => (
              <div className="de-borders-pop">
                <div className="de-borders" role="group" aria-label="Cell borders">
                  {CELL_BORDERS.map((x) => (
                    <button key={x.k} type="button" className="de-tb" title={x.title} aria-label={x.title} onClick={() => (close(), editor.setCellBorders(x.k, `${bLine} ${bColor}`))}>
                      <Icon name={x.icon} />
                    </button>
                  ))}
                </div>
                <div className="de-borders-side">
                  <ColorBtn icon="color" title="Border color" value={bColor} onPick={(c) => setBColor(c ?? '#000000')} noneLabel="Black" />
                  <Dropdown title="Border style" icon="lineStyle" label={null} menuWidth={150} value={bLine} options={CELL_LINES} onChange={(v) => setBLine(v)} />
                </div>
              </div>
            )}
          </Pop>
          <Dropdown
            title="Vertical align"
            icon={VALIGN.find((x) => x.value === st.cells?.vAlign)?.icon ?? 'vTop'}
            label={null}
            menuWidth={130}
            value={st.cells?.vAlign ?? 'top'}
            options={VALIGN}
            onChange={(v) => editor.setCellVAlign(v as 'top' | 'middle' | 'bottom')}
          />
        </>
      )}
      {st?.img && (
        <>
          <Sep />
          <Btn icon="minus" title="Smaller" onClick={() => editor.resizeImage(0.8)} />
          <span className="de-tb-hint" style={{ margin: 0 }}>
            {Math.round(st.img.w)}×{Math.round(st.img.h)}pt
          </span>
          <Btn icon="plus" title="Larger" onClick={() => editor.resizeImage(1.25)} />
          <Btn icon="fit" title="Fit page width" onClick={() => editor.resizeImage('fit')} />
          <Btn icon="imgOrig" title="Original size" onClick={image.original} />
          <Btn icon="imgReplace" title="Replace image" onClick={image.replace} />
          <Btn icon="alt" title="Alt text" active={!!st.img.alt} onClick={image.alt} />
          {image.download && <Btn icon="save" title="Download image" onClick={image.download} />}
          <Btn icon="trash" title="Delete image" onClick={() => editor.deleteSelection()} />
        </>
      )}
      <Sep />
      <Select
        title="Paragraph style"
        width={118}
        menuWidth={250}
        disabled={dis}
        value={styleValue}
        placeholder="Style"
        options={styleOpts}
        onChange={(v) => (v.startsWith('@') ? void editor.applyCanon(v.slice(1) as CanonStyle) : editor.setPara({ style: v }))}
      />
      <Sep />
      <Select title="Font" width={118} menuWidth={200} disabled={dis} value={font} placeholder="Font" options={fontOptions(font)} onChange={(v) => editor.setFont(v)} />
      <Sep />
      <SizeBox value={sz} disabled={dis} onChange={(v) => editor.setSize(v)} />
      <Sep />
      <Btn icon="bold" title={`Bold (${MOD}B)`} active={st?.b} disabled={dis} onClick={() => editor.toggle('b')} />
      <Btn icon="italic" title={`Italic (${MOD}I)`} active={st?.i} disabled={dis} onClick={() => editor.toggle('i')} />
      <Btn icon="underline" title={`Underline (${MOD}U)`} active={st?.u} disabled={dis} onClick={() => editor.toggle('u')} />
      <Btn icon="strike" title="Strikethrough" active={st?.strike} disabled={dis} onClick={() => editor.toggle('strike')} />
      <ColorBtn icon="color" title="Text color" value={st?.color} onPick={(c) => editor.setColor(c)} noneLabel="Automatic" />
      <ColorBtn icon="highlight" title="Highlight color" value={st?.hl} palette={HIGHLIGHTS} onPick={(c) => editor.setHighlight(c)} />
      <Sep />
      <Btn icon="link" title={`Insert link (${MOD}K)`} active={!!st?.link} disabled={dis} onClick={props.onLink} />
      <Btn icon="image" title="Insert image" disabled={dis} onClick={() => void pickFiles().then((f) => void (f.length && editor.insertImages(f)))} />
      <Pop icon="table" title="Insert table" disabled={dis || st?.inTable}>
        {(close) => <GridPicker onPick={(r, c) => (close(), editor.insertTable(r, c))} />}
      </Pop>
      <Btn icon="pageSetup" title="Page setup (size, orientation, margins, columns)" disabled={!props.onPageSetup} onClick={() => props.onPageSetup?.()} />
      <Pop icon="omega" title="Special characters" disabled={dis}>
        {(close) => (
          <div className="de-symbols">
            {SYMBOLS.map((c) => (
              <button key={c} type="button" title={`U+${c.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`} onClick={() => (close(), editor.insertText(c))}>
                {c}
              </button>
            ))}
          </div>
        )}
      </Pop>
      <Sep />
      <Btn icon="alignLeft" title="Align left" active={st?.align === 'left'} disabled={dis} onClick={() => editor.setPara({ align: 'left' }, { align: 'left' })} />
      <Btn icon="alignCenter" title="Center" active={st?.align === 'center'} disabled={dis} onClick={() => editor.setPara({ align: 'center' }, { align: 'center' })} />
      <Btn icon="alignRight" title="Align right" active={st?.align === 'right'} disabled={dis} onClick={() => editor.setPara({ align: 'right' }, { align: 'right' })} />
      <Btn icon="justify" title="Justify" active={st?.align === 'justify'} disabled={dis} onClick={() => editor.setPara({ align: 'justify' }, { align: 'justify' })} />
      <Dropdown
        title="Line & paragraph spacing"
        icon="lineSpacing"
        label={null}
        menuWidth={150}
        disabled={dis}
        value={line}
        options={LINE_SPACING}
        onChange={(v) => editor.setPara({ line: v }, { line: v, lineRule: 'auto' })}
      />
      <Sep />
      <Btn icon="bullets" title="Bulleted list" active={st?.lst === 'bullet'} disabled={dis} onClick={() => editor.toggleList('bullet')} />
      <Btn icon="numbers" title="Numbered list" active={st?.lst === 'number'} disabled={dis} onClick={() => editor.toggleList('number')} />
      <Btn icon="outdent" title="Decrease indent (Shift+Tab)" disabled={dis} onClick={() => editor.indent(-1)} />
      <Btn icon="indent" title="Increase indent (Tab)" disabled={dis} onClick={() => editor.indent(1)} />
      <Sep />
      <Btn icon="sup" title="Superscript" active={st?.va === 'sup'} disabled={dis} onClick={() => editor.format({ va: st?.va === 'sup' ? undefined : 'sup' }, { va: st?.va === 'sup' ? null : 'sup' })} />
      <Btn icon="sub" title="Subscript" active={st?.va === 'sub'} disabled={dis} onClick={() => editor.format({ va: st?.va === 'sub' ? undefined : 'sub' }, { va: st?.va === 'sub' ? null : 'sub' })} />
      <Btn icon="pagebreak" title={`Page break (${MOD}Enter)`} disabled={dis || st?.inTable} onClick={() => editor.insertPageBreak()} />
      <Btn icon="clear" title={`Clear formatting (${MOD}\\)`} disabled={dis} onClick={() => editor.clearFormat()} />
      <Sep />
      <Btn icon="outline" title="Document outline" active={props.outline} onClick={props.onOutline} />
      <Btn icon="search" title={`Find & replace (${MOD}F)`} onClick={props.onFind} />
    </EditorBar>
  )
}
