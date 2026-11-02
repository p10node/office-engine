import { useEffect, useState } from 'react'
import type { ShapeKind } from '../../core/pptx/types'
import { MOD } from '../common/ContextMenu'
import { EditorBar } from '../common/Slots'
import { Btn, ColorBtn, Dropdown, fontName, fontOptions, Menu, Select, Sep, SizeBox, type DropOption } from '../edit/Toolbar'
import type { PptxEditor, PptxFormat } from './PptxEditor'

const HIGHLIGHTS = ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#ff0000', '#0000ff', '#ffc000', '#92d050', '#c0c0c0', '#000000']

export const SHAPES: DropOption<string>[] = [
  { value: 'text', label: 'Text box', icon: 'textbox' },
  { value: 'image', label: 'Image…', icon: 'image' },
  { value: 'line', label: 'Line', icon: 'line' },
  { value: 'rect', label: 'Rectangle', icon: 'rect', sep: true },
  { value: 'roundRect', label: 'Rounded rectangle', icon: 'roundRect' },
  { value: 'ellipse', label: 'Ellipse', icon: 'ellipse' },
  { value: 'triangle', label: 'Triangle', icon: 'triangle' },
  { value: 'rtTriangle', label: 'Right triangle' },
  { value: 'diamond', label: 'Diamond' },
  { value: 'parallelogram', label: 'Parallelogram' },
  { value: 'trapezoid', label: 'Trapezoid' },
  { value: 'pentagon', label: 'Pentagon' },
  { value: 'hexagon', label: 'Hexagon' },
  { value: 'octagon', label: 'Octagon' },
  { value: 'donut', label: 'Donut' },
  { value: 'can', label: 'Cylinder' },
  { value: 'heart', label: 'Heart' },
  { value: 'plus', label: 'Plus' },
  { value: 'rightArrow', label: 'Arrow right', icon: 'arrow', sep: true },
  { value: 'leftArrow', label: 'Arrow left' },
  { value: 'upArrow', label: 'Arrow up' },
  { value: 'downArrow', label: 'Arrow down' },
  { value: 'leftRightArrow', label: 'Double arrow' },
  { value: 'chevron', label: 'Chevron' },
  { value: 'homePlate', label: 'Pentagon arrow' },
  { value: 'star4', label: '4-point star', sep: true },
  { value: 'star5', label: '5-point star', icon: 'star' },
  { value: 'star6', label: '6-point star' },
  { value: 'wedgeRectCallout', label: 'Callout', sep: true },
  { value: 'wedgeRoundRectCallout', label: 'Rounded callout' },
]

const LINE_W = [0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => ({ value: w, label: `${w} pt` }))
const DASHES = [
  { value: 'solid', label: 'Solid', hint: '────' },
  { value: 'sysDot', label: 'Dotted', hint: '········' },
  { value: 'dash', label: 'Dashed', hint: '- - - -' },
  { value: 'lgDash', label: 'Long dash', hint: '— — —' },
  { value: 'dashDot', label: 'Dash dot', hint: '-·-·-·' },
]
const SPACING = [1, 1.15, 1.5, 2].map((v) => ({ value: v, label: v === 1 ? 'Single' : v === 2 ? 'Double' : String(v) }))
const VALIGN = [
  { value: 't', label: 'Top', icon: 'vTop' },
  { value: 'ctr', label: 'Middle', icon: 'vMiddle' },
  { value: 'b', label: 'Bottom', icon: 'vBottom' },
] as const

export type Arrange = 'front' | 'forward' | 'backward' | 'back'
export type Rotate = 'cw' | 'ccw' | 'flipH' | 'flipV'
export type AlignTo = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'
export type SlideCmd = 'add' | 'duplicate' | 'delete' | 'up' | 'down' | 'hide'

export function PptxToolbar(props: {
  editor: PptxEditor
  /** selected shape (not editing its text) */
  shape: { canText: boolean; anchor?: 't' | 'ctr' | 'b' } | null
  /** anchor of the text box being edited */
  textAnchor?: 't' | 'ctr' | 'b'
  canUndo: boolean
  onUndo: (redo: boolean) => void
  onInsert: (k: ShapeKind) => void
  onFill: (c: string | null) => void
  onLine: (p: { color?: string | null; w?: number; dash?: string }) => void
  onArrange: (k: Arrange) => void
  onRotate: (k: Rotate) => void
  onAlign: (k: AlignTo) => void
  onAnchor: (a: 't' | 'ctr' | 'b') => void
  onDuplicate: () => void
  onDelete: () => void
  onBackground: (c: string | null) => void
  onSlide: (k: SlideCmd) => void
  onPresent: (fromStart: boolean) => void
}) {
  const { editor } = props
  const [st, setSt] = useState<PptxFormat | null>(null)
  useEffect(() => {
    editor.onState = setSt
    return () => {
      editor.onState = undefined
    }
  }, [editor])
  const text = !!st && editor.active
  const font = fontName(st?.font)
  const sz = st?.sz ? Math.round(st.sz * 2) / 2 : undefined
  const anchor = text ? props.textAnchor : props.shape?.anchor
  return (
    <EditorBar>
      <Btn icon="undo" title={`Undo (${MOD}Z)`} disabled={!text && !props.canUndo} onClick={() => (text ? editor.undo() : props.onUndo(false))} />
      <Btn icon="redo" title={`Redo (${MOD}Y)`} onClick={() => (text ? editor.redo() : props.onUndo(true))} />
      <Sep />
      <Menu title="Insert" label="Insert" icon="plus" items={SHAPES as { value: string; label: string; icon?: string; sep?: boolean }[]} onPick={(v) => props.onInsert(v)} />
      <Menu
        title="Slide"
        label="Slide"
        icon="slideAdd"
        items={[
          { value: 'add', label: 'New slide', icon: 'slideAdd', hint: `${MOD}M` },
          { value: 'duplicate', label: 'Duplicate slide', icon: 'copy' },
          { value: 'delete', label: 'Delete slide', icon: 'trash' },
          { value: 'up', label: 'Move slide up', icon: 'up', sep: true },
          { value: 'down', label: 'Move slide down', icon: 'down' },
          { value: 'hide', label: 'Hide / show slide', icon: 'hide', sep: true },
        ]}
        onPick={(v) => props.onSlide(v as SlideCmd)}
      />
      <Sep />
      {text ? (
        <>
          <Select title="Font" width={112} menuWidth={200} value={font} placeholder="Font" options={fontOptions(font)} onChange={(v) => editor.setFont(v)} />
          <Sep />
          <SizeBox value={sz} onChange={(v) => editor.setSize(v)} />
          <Sep />
          <Btn icon="bold" title={`Bold (${MOD}B)`} active={st?.b} onClick={() => editor.toggle('b')} />
          <Btn icon="italic" title={`Italic (${MOD}I)`} active={st?.i} onClick={() => editor.toggle('i')} />
          <Btn icon="underline" title={`Underline (${MOD}U)`} active={st?.u} onClick={() => editor.toggle('u')} />
          <Btn icon="strike" title="Strikethrough" active={st?.strike} onClick={() => editor.toggle('strike')} />
          <ColorBtn icon="color" title="Text color" value={st?.color} onPick={(c) => editor.setColor(c)} noneLabel="Automatic" />
          <ColorBtn icon="highlight" title="Highlight color" value={st?.hl} palette={HIGHLIGHTS} onPick={(c) => editor.setHighlight(c)} />
          <Btn icon="sup" title="Superscript" active={(st?.baseline ?? 0) > 0} onClick={() => editor.toggleBaseline('sup')} />
          <Btn icon="sub" title="Subscript" active={(st?.baseline ?? 0) < 0} onClick={() => editor.toggleBaseline('sub')} />
          <Sep />
          <Btn icon="alignLeft" title="Align left" active={st?.algn === 'left'} onClick={() => editor.setAlign('left')} />
          <Btn icon="alignCenter" title="Center" active={st?.algn === 'center'} onClick={() => editor.setAlign('center')} />
          <Btn icon="alignRight" title="Align right" active={st?.algn === 'right'} onClick={() => editor.setAlign('right')} />
          <Btn icon="justify" title="Justify" active={st?.algn === 'justify'} onClick={() => editor.setAlign('justify')} />
          <Dropdown title="Vertical align" icon={VALIGN.find((x) => x.value === anchor)?.icon ?? 'vTop'} label={null} menuWidth={140} value={anchor ?? 't'} options={VALIGN.map((x) => ({ ...x }))} onChange={props.onAnchor} />
          <Dropdown
            title="Line spacing"
            icon="lineSpacing"
            label={null}
            menuWidth={140}
            value={st?.lnPct ? Math.round(st.lnPct * 100) / 100 : ''}
            options={SPACING}
            onChange={(v) => editor.setLineSpacing(v)}
          />
          <Sep />
          <Btn icon="bullets" title="Bulleted list" active={st?.bullet === 'bullet'} onClick={() => editor.toggleBullet('bullet')} />
          <Btn icon="numbers" title="Numbered list" active={st?.bullet === 'number'} onClick={() => editor.toggleBullet('number')} />
          <Btn icon="outdent" title="Decrease indent (Shift+Tab)" disabled={!st?.lvl} onClick={() => editor.indentLevel(-1)} />
          <Btn icon="indent" title="Increase indent" onClick={() => editor.indentLevel(1)} />
          <Btn icon="clear" title="Clear formatting" onClick={() => editor.clearFormat()} />
        </>
      ) : props.shape ? (
        <>
          <ColorBtn icon="fill" title="Fill color" onPick={props.onFill} noneLabel="Transparent" />
          <ColorBtn icon="borders" title="Border color" onPick={(c) => props.onLine({ color: c })} noneLabel="Transparent" />
          <Dropdown title="Border weight" icon="lineWeight" label={null} menuWidth={120} options={LINE_W} onChange={(w) => props.onLine({ w })} />
          <Dropdown title="Border dash" icon="lineStyle" label={null} menuWidth={170} options={DASHES} onChange={(dash) => props.onLine({ dash })} />
          <Sep />
          <Menu
            title="Arrange"
            label="Arrange"
            icon="front"
            items={[
              { value: 'front', label: 'Bring to front', icon: 'front', hint: `${MOD}⇧↑` },
              { value: 'forward', label: 'Bring forward', hint: `${MOD}↑` },
              { value: 'backward', label: 'Send backward', hint: `${MOD}↓` },
              { value: 'back', label: 'Send to back', icon: 'back', hint: `${MOD}⇧↓` },
            ]}
            onPick={(v) => props.onArrange(v as Arrange)}
          />
          <Dropdown
            title="Rotate"
            icon="rotate"
            label={null}
            menuWidth={220}
            options={[
              { value: 'cw', label: 'Rotate 90° clockwise', icon: 'rotate' },
              { value: 'ccw', label: 'Rotate 90° counter-clockwise', icon: 'rotateCcw' },
              { value: 'flipH', label: 'Flip horizontally', icon: 'flipH', sep: true },
              { value: 'flipV', label: 'Flip vertically', icon: 'flipV' },
            ]}
            onChange={(v) => props.onRotate(v as Rotate)}
          />
          <Dropdown
            title="Align on slide"
            icon="alignObj"
            label={null}
            menuWidth={150}
            options={[
              { value: 'left', label: 'Left', icon: 'alignLeft' },
              { value: 'center', label: 'Center', icon: 'alignCenter' },
              { value: 'right', label: 'Right', icon: 'alignRight' },
              { value: 'top', label: 'Top', icon: 'vTop', sep: true },
              { value: 'middle', label: 'Middle', icon: 'vMiddle' },
              { value: 'bottom', label: 'Bottom', icon: 'vBottom' },
            ]}
            onChange={(v) => props.onAlign(v as AlignTo)}
          />
          {props.shape.canText && (
            <Dropdown title="Vertical text align" icon={VALIGN.find((x) => x.value === anchor)?.icon ?? 'vTop'} label={null} menuWidth={140} value={anchor ?? 't'} options={VALIGN.map((x) => ({ ...x }))} onChange={props.onAnchor} />
          )}
          <Sep />
          <Btn icon="copy" title={`Duplicate (${MOD}D)`} onClick={props.onDuplicate} />
          <Btn icon="trash" title="Delete (Del)" onClick={props.onDelete} />
        </>
      ) : (
        <>
          <ColorBtn icon="fill" title="Slide background" onPick={props.onBackground} noneLabel="Layout default" />
          <span className="de-tb-hint">Click a shape or text to edit it</span>
        </>
      )}
      <Sep />
      <Dropdown
        title="Present"
        icon="play"
        label="Present"
        menuWidth={200}
        options={[
          { value: 'cur', label: 'From current slide', hint: 'F5' },
          { value: 'start', label: 'From beginning', hint: '⇧F5' },
        ]}
        onChange={(v) => props.onPresent(v === 'start')}
      />
    </EditorBar>
  )
}
