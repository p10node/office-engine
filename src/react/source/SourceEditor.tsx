import { useMemo, useRef, type KeyboardEvent } from 'react'
import type { PlainFormat } from '../../core/plain'
import { Icon } from '../edit/Toolbar'
import { highlight } from './highlight'

const LABEL: Partial<Record<PlainFormat, string>> = { markdown: 'Markdown', html: 'HTML', mermaid: 'Mermaid', plantuml: 'PlantUML', graphviz: 'Graphviz' }

export interface SourceEditorProps {
  text: string
  format: PlainFormat
  onChange: (text: string) => void
  readOnly?: boolean
  /** a line under the header (e.g. what editing the source does) */
  notice?: string
  onCopy?: () => void
  onDownload?: () => void
}

/**
 * The text of a Markdown / HTML / diagram file, editable, with line numbers and syntax colors: a transparent
 * textarea over the colored text, both sized by their content in one grid cell (they scroll together).
 */
export function SourceEditor({ text, format, onChange, readOnly, notice, onCopy, onDownload }: SourceEditorProps) {
  const tokens = useMemo(() => highlight(format, text), [format, text])
  const lines = useMemo(() => {
    let n = 1
    for (let i = text.indexOf('\n'); i >= 0; i = text.indexOf('\n', i + 1)) n++
    return n
  }, [text])
  const gutter = useMemo(() => Array.from({ length: lines }, (_, k) => k + 1).join('\n'), [lines])
  // Escape, then Tab: leave the editor (Tab alone indents)
  const leaving = useRef(false)
  const insert = (el: HTMLTextAreaElement, s: string) => {
    // execCommand keeps the native undo history; setRangeText where it is gone
    if (!el.ownerDocument.execCommand?.('insertText', false, s)) {
      el.setRangeText(s, el.selectionStart, el.selectionEnd, 'end')
      onChange(el.value)
    }
  }
  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    const el = e.currentTarget
    if (e.key === 'Escape') {
      leaving.current = true
      return
    }
    const wasLeaving = leaving.current
    leaving.current = false
    if (readOnly || e.altKey || e.metaKey || e.ctrlKey || e.nativeEvent.isComposing) return
    if (e.key === 'Tab' && !e.shiftKey && !wasLeaving) {
      e.preventDefault()
      insert(el, '  ')
    } else if (e.key === 'Enter' && !e.shiftKey) {
      // keep the indentation of the line
      const start = el.value.lastIndexOf('\n', el.selectionStart - 1) + 1
      const indent = /^[ \t]*/.exec(el.value.slice(start, el.selectionStart))![0]
      if (!indent) return
      e.preventDefault()
      insert(el, '\n' + indent)
    }
  }
  return (
    <div className="de-src" data-testid="de-source">
      <div className="de-src-head">
        <span className="de-src-title">
          <Icon name="code" size={15} />
          {LABEL[format] ?? 'Source'}
        </span>
        <span className="de-src-lines">
          {lines} {lines === 1 ? 'line' : 'lines'}
        </span>
        <span style={{ flex: 1 }} />
        {onCopy && (
          <button type="button" className="de-src-btn" onClick={onCopy} title="Copy the source">
            <Icon name="copy" size={15} />
          </button>
        )}
        {onDownload && (
          <button type="button" className="de-src-btn" onClick={onDownload} title="Download the source">
            <Icon name="save" size={15} />
          </button>
        )}
      </div>
      {notice && <div className="de-src-notice">{notice}</div>}
      <div className="de-src-scroll">
        <pre className="de-src-gutter" aria-hidden="true">
          {gutter}
        </pre>
        <div className="de-src-code">
          <pre className="de-src-hl" aria-hidden="true">
            {tokens.map(([c, t], k) => (c ? <span key={k} className={`de-tk-${c}`}>{t}</span> : t))}
            {'\n'}
          </pre>
          <textarea
            className="de-src-input"
            value={text}
            readOnly={readOnly}
            onChange={(e) => onChange(e.target.value)}
            onKeyDown={onKeyDown}
            spellCheck={false}
            autoCapitalize="off"
            autoComplete="off"
            autoCorrect="off"
            wrap="off"
            rows={1}
            cols={1}
            aria-label={`${LABEL[format] ?? 'Source'} source`}
            data-testid="de-source-input"
          />
        </div>
      </div>
    </div>
  )
}
