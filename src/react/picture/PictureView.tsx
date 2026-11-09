import { useEffect, useState } from 'react'
import { svgImage } from '../../core/svgfile'
import type { DocumentClient } from '../../core/worker/client'
import { useFileOps } from '../common/Slots'

type Pic = { url: string; w: number; h: number; alt: string } | { error: string }

/**
 * An SVG file: the picture alone, no page around it, scaled down to fit the view and at its own size otherwise.
 * Drawn from the file's text (the Code pane edits it); the document built from it is what Download as Word gives.
 */
export function PictureView({ client }: { client: DocumentClient }) {
  const [pic, setPic] = useState<Pic | null>(null)
  useEffect(() => {
    let live = true
    let url = ''
    // the document holding the picture is loaded too: Download as Word saves it
    client.call('docx.start').catch(() => {})
    client.call('plain.source').then(
      (text) => {
        if (!live) return
        try {
          const img = svgImage((text as string | null) ?? '')
          url = URL.createObjectURL(new Blob([img.svg], { type: 'image/svg+xml' }))
          setPic({ url, w: img.w, h: img.h, alt: img.alt ?? '' })
        } catch (e) {
          setPic({ error: (e as Error).message })
        }
      },
      (e: Error) => live && setPic({ error: e.message }),
    )
    return () => {
      live = false
      if (url) URL.revokeObjectURL(url)
    }
  }, [client])
  const shown = pic && 'url' in pic ? pic : null
  useFileOps({
    printable: async () => {
      if (!shown) throw new Error('The picture is not shown')
      return { size: 'auto', margin: '1cm', node: <img src={shown.url} alt={shown.alt} style={{ width: `${shown.w}px`, maxWidth: '100%' }} /> }
    },
    text: async () => ((await client.call('plain.source')) as string | null) ?? '',
    stats: () => (shown ? [['Picture size', `${Math.round(shown.w)} × ${Math.round(shown.h)} px`]] : []),
  })
  if (!pic)
    return (
      <div className="de-center">
        <div className="de-spinner" />
      </div>
    )
  if ('error' in pic)
    return (
      <div className="de-center">
        <div className="de-error" role="alert">
          <strong>SVG picture not drawn.</strong>
          <div>{pic.error}</div>
        </div>
      </div>
    )
  return (
    <div className="de-scroll de-picture" data-testid="de-picture">
      <img src={pic.url} alt={pic.alt} draggable={false} style={{ width: `${pic.w}px`, height: `${pic.h}px` }} />
    </div>
  )
}
