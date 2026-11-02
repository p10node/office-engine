import { Fragment, useEffect, useRef, useState } from 'react'
import type { DocumentClient } from '../../core/worker/client'

const CORE: [string, string][] = [
  ['title', 'Title'],
  ['subject', 'Subject'],
  ['creator', 'Author'],
  ['keywords', 'Keywords'],
  ['description', 'Comments'],
  ['lastModifiedBy', 'Last modified by'],
  ['created', 'Created'],
  ['modified', 'Modified'],
]

/** Core properties (docProps/core.xml) that are set. */
async function coreProps(client: DocumentClient): Promise<[string, string][]> {
  const xml = await (await client.call('part', 'docProps/core.xml')).text()
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  const out: [string, string][] = []
  for (const [tag, label] of CORE) {
    const v = doc.getElementsByTagNameNS('*', tag)[0]?.textContent?.trim()
    if (!v) continue
    const date = tag === 'created' || tag === 'modified' ? new Date(v) : null
    out.push([label, date && !isNaN(date.getTime()) ? date.toLocaleString() : v])
  }
  return out
}

export function DetailsDialog(props: { client: DocumentClient; file: [string, string][]; stats: [string, string][]; onClose: () => void }) {
  const [core, setCore] = useState<[string, string][]>([])
  const closeRef = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    closeRef.current?.focus()
    let live = true
    coreProps(props.client).then(
      (c) => live && setCore(c),
      () => {},
    )
    return () => {
      live = false
    }
  }, [props.client])
  const groups = [props.file, core, props.stats].filter((g) => g.length)
  return (
    <div
      className="de-dialog-back"
      onMouseDown={(e) => e.target === e.currentTarget && props.onClose()}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          props.onClose()
        }
      }}
    >
      <div className="de-dialog" role="dialog" aria-label="Document details">
        <div className="de-dialog-title">Document details</div>
        <dl className="de-details">
          {groups.map((g, k) => (
            <Fragment key={k}>
              {k > 0 && <div className="de-details-sep" />}
              {g.map(([label, value]) => (
                <Fragment key={label}>
                  <dt>{label}</dt>
                  <dd>{value}</dd>
                </Fragment>
              ))}
            </Fragment>
          ))}
        </dl>
        <div className="de-dialog-btns">
          <button ref={closeRef} type="button" className="de-btn de-btn-primary" onClick={props.onClose}>
            Close
          </button>
        </div>
      </div>
    </div>
  )
}
