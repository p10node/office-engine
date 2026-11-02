import { useState } from 'react'
import type { BorderKind, FmtPatch } from '../../core/xlsx/types'
import { MOD } from '../common/ContextMenu'
import { EditorBar } from '../common/Slots'
import { Btn, ColorBtn, Dropdown, fontName, fontOptions, Icon, Menu, Pop, Select, Sep, SizeBox } from '../edit/Toolbar'

const FORMATS = [
  { value: 'General', label: 'Automatic', hint: '1000.12' },
  { value: '@', label: 'Plain text', hint: 'abc' },
  { value: '0', label: 'Number', hint: '1000', sep: true },
  { value: '0.00', label: 'Number', hint: '1000.12' },
  { value: '#,##0', label: 'Number', hint: '1,000' },
  { value: '#,##0.00', label: 'Number', hint: '1,000.12' },
  { value: '0.00E+00', label: 'Scientific', hint: '1.01E+03' },
  { value: '0%', label: 'Percent', hint: '10%', sep: true },
  { value: '0.00%', label: 'Percent', hint: '10.12%' },
  { value: '"$"#,##0.00', label: 'Currency', hint: '$1,000.12', sep: true },
  { value: '"$"#,##0', label: 'Currency rounded', hint: '$1,000' },
  { value: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)', label: 'Accounting', hint: '$ (1,000.12)' },
  { value: '#,##0 "₫"', label: 'Currency', hint: '1,000 ₫' },
  { value: '# ?/?', label: 'Fraction', hint: '1000 1/8' },
  { value: 'yyyy-mm-dd', label: 'Date', hint: '2026-09-26', sep: true },
  { value: 'dd/mm/yyyy', label: 'Date', hint: '26/09/2026' },
  { value: 'd mmmm yyyy', label: 'Long date', hint: '26 September 2026' },
  { value: 'h:mm', label: 'Time', hint: '15:59' },
  { value: 'h:mm:ss AM/PM', label: 'Time', hint: '3:59:00 PM' },
  { value: 'yyyy-mm-dd h:mm', label: 'Date time', hint: '2026-09-26 15:59' },
  { value: '[h]:mm:ss', label: 'Duration', hint: '24:01:00' },
]
const SIZES = [6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48]
const BORDERS: { k: BorderKind; icon: string; title: string }[] = [
  { k: 'all', icon: 'bAll', title: 'All borders' },
  { k: 'inner', icon: 'bInner', title: 'Inner borders' },
  { k: 'outer', icon: 'bOuter', title: 'Outer borders' },
  { k: 'none', icon: 'bNone', title: 'Clear borders' },
  { k: 'top', icon: 'bTop', title: 'Top border' },
  { k: 'bottom', icon: 'bBottom', title: 'Bottom border' },
  { k: 'left', icon: 'bLeft', title: 'Left border' },
  { k: 'right', icon: 'bRight', title: 'Right border' },
]
const BORDER_STYLES = [
  { value: 'thin', label: 'Thin', hint: '───' },
  { value: 'medium', label: 'Medium', hint: '━━━' },
  { value: 'thick', label: 'Thick', hint: '▬▬▬' },
  { value: 'dashed', label: 'Dashed', hint: '- - -' },
  { value: 'dotted', label: 'Dotted', hint: '·····' },
  { value: 'double', label: 'Double', hint: '═══' },
] as const
const HALIGN = [
  { value: 'left', label: 'Left', icon: 'alignLeft' },
  { value: 'center', label: 'Center', icon: 'alignCenter' },
  { value: 'right', label: 'Right', icon: 'alignRight' },
] as const
const VALIGN = [
  { value: 'top', label: 'Top', icon: 'vTop' },
  { value: 'center', label: 'Middle', icon: 'vMiddle' },
  { value: 'bottom', label: 'Bottom', icon: 'vBottom' },
] as const
const ROTATE = [
  { value: 0, label: 'None', icon: 'rotNone' },
  { value: 45, label: 'Tilt up', icon: 'rotUp' },
  { value: -45, label: 'Tilt down', icon: 'rotDown' },
  { value: 255, label: 'Stack vertically', icon: 'rotStack' },
  { value: 90, label: 'Rotate up', icon: 'rotUp90' },
  { value: -90, label: 'Rotate down', icon: 'rotDown90' },
]
const FUNCS = ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN']

export type MergeMode = 'all' | 'across' | 'down' | 'unmerge'

/** number format with one decimal place more / less (Google Sheets .0 / .00 buttons) */
function decimals(fmt: string, d: number): string {
  const base = fmt === 'General' || fmt === '@' ? '0' : fmt
  const m = /0(\.0*)?/.exec(base)
  if (!m) return d > 0 ? '0.0' : '0'
  const n = Math.max(0, (m[1] ? m[1].length - 1 : 0) + d)
  return base.slice(0, m.index) + '0' + (n ? '.' + '0'.repeat(n) : '') + base.slice(m.index + m[0].length)
}

export function XlsxToolbar(props: {
  css: Record<string, string | number>
  /** workbook default font (cells without an explicit font) */
  base: { fontFamily?: string; fontSize?: string | number }
  format: string
  merged: boolean
  frozen: boolean
  gridlines: boolean
  painter: boolean
  onToggle: (k: 'b' | 'i' | 'u' | 'strike') => void
  onFormat: (p: FmtPatch) => void
  onUndo: () => void
  onRedo: () => void
  onPaint: () => void
  onMerge: (mode: MergeMode) => void
  onFreeze: (k: 'none' | 'row' | 'col' | 'sel') => void
  onGridlines: () => void
  onFunc: (fn: string | null) => void
  onSort: (desc: boolean) => void
  onRows: (k: 'rowAbove' | 'rowBelow' | 'colLeft' | 'colRight' | 'delRow' | 'delCol' | 'hideRows' | 'hideCols' | 'unhideRows' | 'unhideCols') => void
  onFind: () => void
}) {
  const { css, onFormat: f } = props
  const [bColor, setBColor] = useState('#000000')
  const [bStyle, setBStyle] = useState<(typeof BORDER_STYLES)[number]['value']>('thin')
  const deco = String(css.textDecoration ?? '')
  const align = (css.textAlign as string | undefined) ?? ''
  const valign = css.alignItems === 'flex-start' ? 'top' : css.alignItems === 'center' ? 'center' : 'bottom'
  const rot = Number(css.textRotate ?? 0)
  const pt = (v: unknown) => (typeof v === 'string' && v.endsWith('pt') ? parseFloat(v) : typeof v === 'number' ? Math.round(v * 0.75 * 2) / 2 : undefined)
  const sz = pt(css.fontSize) ?? pt(props.base.fontSize)
  const font = fontName((css.fontFamily ?? props.base.fontFamily) as string | undefined)
  const known = FORMATS.find((x) => x.value === props.format)
  const formats = known ? FORMATS : [{ value: props.format, label: 'Custom', hint: props.format }, ...FORMATS]
  return (
    <EditorBar>
      <Btn icon="undo" title={`Undo (${MOD}Z)`} onClick={props.onUndo} />
      <Btn icon="redo" title={`Redo (${MOD}Y)`} onClick={props.onRedo} />
      <Btn icon="paint" title="Paint format (click, then select cells)" active={props.painter} onClick={props.onPaint} />
      <Sep />
      <Btn icon="currency" title="Format as currency" onClick={() => f({ numFmt: '"$"#,##0.00' })} />
      <Btn icon="percent" title="Format as percent" onClick={() => f({ numFmt: '0%' })} />
      <Btn icon="decLess" title="Decrease decimal places" onClick={() => f({ numFmt: decimals(props.format, -1) })} />
      <Btn icon="decMore" title="Increase decimal places" onClick={() => f({ numFmt: decimals(props.format, 1) })} />
      <Select title={`Number format: ${known?.label ?? 'Custom'}`} width={96} menuWidth={270} value={props.format} options={formats} onChange={(v) => f({ numFmt: v })} />
      <Sep />
      <Select title="Font" width={112} menuWidth={200} value={font} placeholder="Font" options={fontOptions(font)} onChange={(v) => f({ font: v })} />
      <Sep />
      <SizeBox value={sz} sizes={SIZES} onChange={(v) => f({ sz: v })} />
      <Sep />
      <Btn icon="bold" title={`Bold (${MOD}B)`} active={css.fontWeight === 700} onClick={() => props.onToggle('b')} />
      <Btn icon="italic" title={`Italic (${MOD}I)`} active={css.fontStyle === 'italic'} onClick={() => props.onToggle('i')} />
      <Btn icon="underline" title={`Underline (${MOD}U)`} active={deco.includes('underline')} onClick={() => props.onToggle('u')} />
      <Btn icon="strike" title={`Strikethrough (${MOD}⇧X)`} active={deco.includes('line-through')} onClick={() => props.onToggle('strike')} />
      <ColorBtn icon="color" title="Text color" value={css.color as string | undefined} onPick={(c) => f({ color: c })} noneLabel="Automatic" />
      <Sep />
      <ColorBtn icon="fill" title="Fill color" value={typeof css.background === 'string' && !css.background.includes('gradient') ? css.background : undefined} onPick={(c) => f({ fill: c })} noneLabel="No fill" />
      <Pop icon="borders" title="Borders" chevron>
        {(close) => (
          <div className="de-borders-pop">
            <div className="de-borders" role="group" aria-label="Borders">
              {BORDERS.map((x) => (
                <button
                  key={x.k}
                  type="button"
                  className="de-tb"
                  title={x.title}
                  aria-label={x.title}
                  onClick={() => {
                    close()
                    f({ border: x.k, borderColor: bColor, borderStyle: bStyle })
                  }}
                >
                  <Icon name={x.icon} />
                </button>
              ))}
            </div>
            <div className="de-borders-side">
              <ColorBtn icon="color" title="Border color" value={bColor} onPick={(c) => setBColor(c ?? '#000000')} noneLabel="Black" />
              <Dropdown title="Border style" icon="lineStyle" label={null} menuWidth={150} value={bStyle} options={[...BORDER_STYLES]} onChange={(v) => setBStyle(v)} />
            </div>
          </div>
        )}
      </Pop>
      <Dropdown
        title="Merge cells"
        icon="merge"
        label={null}
        menuWidth={190}
        options={[
          { value: 'all', label: 'Merge all' },
          { value: 'across', label: 'Merge horizontally' },
          { value: 'down', label: 'Merge vertically' },
          { value: 'unmerge', label: 'Unmerge', disabled: !props.merged, sep: true },
        ]}
        onChange={(v) => props.onMerge(v as MergeMode)}
      />
      <Sep />
      <Dropdown
        title="Horizontal align"
        icon={HALIGN.find((x) => x.value === align)?.icon ?? 'alignLeft'}
        label={null}
        menuWidth={140}
        value={align}
        options={HALIGN.map((x) => ({ ...x }))}
        onChange={(v) => f({ align: v === align ? null : (v as 'left') })}
      />
      <Dropdown title="Vertical align" icon={VALIGN.find((x) => x.value === valign)!.icon} label={null} menuWidth={140} value={valign} options={VALIGN.map((x) => ({ ...x }))} onChange={(v) => f({ valign: v })} />
      <Btn icon="wrap" title="Wrap text" active={css.whiteSpace === 'pre-wrap'} onClick={() => f({ wrap: css.whiteSpace !== 'pre-wrap' })} />
      <Dropdown title="Text rotation" icon={ROTATE.find((x) => x.value === rot)?.icon ?? 'rotNone'} label={null} menuWidth={180} value={rot} options={ROTATE} onChange={(v) => f({ rotate: v || null })} />
      <Sep />
      <Menu
        title="Insert, delete or hide rows & columns"
        label="Rows"
        icon="rowAdd"
        items={[
          { value: 'rowAbove', label: 'Insert row above', icon: 'rowAbove' },
          { value: 'rowBelow', label: 'Insert row below', icon: 'rowBelow' },
          { value: 'colLeft', label: 'Insert column left', icon: 'colLeft' },
          { value: 'colRight', label: 'Insert column right', icon: 'colRight' },
          { value: 'delRow', label: 'Delete row(s)', icon: 'delRow', sep: true },
          { value: 'delCol', label: 'Delete column(s)', icon: 'delCol' },
          { value: 'hideRows', label: 'Hide row(s)', icon: 'hide', sep: true },
          { value: 'hideCols', label: 'Hide column(s)', icon: 'hide' },
          { value: 'unhideRows', label: 'Unhide rows', icon: 'view' },
          { value: 'unhideCols', label: 'Unhide columns', icon: 'view' },
        ]}
        onPick={(v) => props.onRows(v as 'rowAbove')}
      />
      <Dropdown
        title="Functions"
        icon="sum"
        label={null}
        menuWidth={170}
        options={[...FUNCS.map((x) => ({ value: x, label: x })), { value: '', label: 'More functions…', sep: true }]}
        onChange={(v) => props.onFunc(v || null)}
      />
      <Dropdown
        title="Sort range"
        icon="sort"
        label={null}
        menuWidth={200}
        options={[
          { value: 'asc', label: 'Sort range A → Z', icon: 'sortAZ' },
          { value: 'desc', label: 'Sort range Z → A', icon: 'sortZA' },
        ]}
        onChange={(v) => props.onSort(v === 'desc')}
      />
      <Menu
        title="View: freeze, gridlines"
        label="View"
        icon="freeze"
        items={[
          { value: 'row', label: 'Freeze 1 row' },
          { value: 'col', label: 'Freeze 1 column' },
          { value: 'sel', label: 'Freeze up to current cell' },
          { value: 'none', label: 'Unfreeze', disabled: !props.frozen },
          { value: 'grid', label: props.gridlines ? 'Hide gridlines' : 'Show gridlines', icon: 'gridlines', sep: true },
        ]}
        onPick={(v) => (v === 'grid' ? props.onGridlines() : props.onFreeze(v as 'row'))}
      />
      <Sep />
      <Btn icon="clear" title={`Clear formatting (${MOD}\\)`} onClick={() => f({ clear: true })} />
      <Btn icon="search" title={`Find & replace (${MOD}F)`} onClick={props.onFind} />
    </EditorBar>
  )
}
