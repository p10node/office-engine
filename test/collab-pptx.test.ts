import { describe, expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import type { PParaSpec, SlideModel } from '../src/core/pptx/types'
import { CollabHub } from '../src/collab/hub'
import { startCollab, type Collab } from '../src/collab'
import { MemoryTransport } from '../src/collab/transport'
import { ZipArchive } from '../src/core/zip'
import { openSource } from '../src/core/source'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

async function open(hub: CollabHub, name: string, latency: () => number) {
  const client = await DocumentClient.open(Bun.file(S('sample.pptx')) as unknown as Blob, { worker: false })
  const collab = startCollab(client, { transport: new MemoryTransport(hub, latency), room: 'p', user: { name } })
  return { client, collab }
}

async function settle(cs: Collab[]) {
  for (let i = 0; i < 300; i++) {
    await sleep(10)
    await Promise.all(cs.map((c) => c.session.idle()))
    if (cs.every((c) => c.pending === 0 && c.status === 'live') && new Set(cs.map((c) => c.session.rev)).size === 1) return
  }
  throw new Error('did not settle')
}

/** every slide part of the saved file */
async function dump(client: DocumentClient) {
  const zip = await ZipArchive.open(await openSource(await client.call('save')))
  const meta = client.info.pptx!
  const out: string[] = [JSON.stringify(meta.slides.map((s) => [s.path, s.hidden]))]
  for (const s of meta.slides) out.push(new TextDecoder().decode(await zip.bytes(s.path)))
  return out
}

const textShapes = (m: SlideModel) => m.els.filter((e) => e.t === 'shape' && e.sid !== undefined && e.text) as (SlideModel['els'][number] & { t: 'shape' })[]

describe('collab pptx', () => {
  for (const start of [5, 17, 29, 61]) {
    test(`concurrent text, shape and slide edits converge (seed ${start})`, async () => {
      let seed = start
      const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
      const hub = new CollabHub()
      const a = await open(hub, 'Ann', () => Math.floor(rand() * 6))
      const b = await open(hub, 'Bob', () => Math.floor(rand() * 6))
      await settle([a.collab, b.collab])
      const ops: Promise<unknown>[] = []
      const run = (p: Promise<unknown>) => ops.push(p.catch((e) => console.error(e)))
      // text edits that took effect where they were made (a later shape / slide delete or undo may still drop them all)
      let texts = 0
      for (let i = 0; i < 50; i++) {
        const who = rand() < 0.5 ? a : b
        const meta = who.client.info.pptx!
        const si = Math.floor(rand() * Math.min(meta.slides.length, 3))
        // a collaborator may be deleting that slide right now
        const m = await who.client.call('pptx.slide', si).catch(() => null)
        if (!m) continue
        const x = rand()
        const shapes = m.els.filter((e) => e.top !== undefined)
        const el = shapes[Math.floor(rand() * shapes.length)]
        if (x < 0.35) {
          const ts = textShapes(m)
          const t = ts[Math.floor(rand() * ts.length)]
          if (t?.text) {
            const paras: PParaSpec[] = t.text.paras.map((p, k) => ({ k: p.k ?? k, runs: p.runs.map((r) => ({ text: r.text ?? '', src: r.src })) }))
            const mark = (who === a ? ' a' : ' b') + i
            if (paras[0]) paras[0].runs.push({ text: mark })
            run(
              who.client.call('pptx.setText', si, t.sid!, paras).then((r) => {
                if (JSON.stringify(r ?? null).includes(mark)) texts++
              }),
            )
          }
        } else if (x < 0.55 && el) run(who.client.call('pptx.shapes', si, [{ op: 'xfrm', top: el.top!, box: { ...el.box, x: el.box.x + 5 * i } }]))
        else if (x < 0.65 && el) run(who.client.call('pptx.shapes', si, [{ op: 'fill', top: el.top!, color: who === a ? '#ff0000' : '#00ff00' }]))
        else if (x < 0.72) run(who.client.call('pptx.shapes', si, [{ op: 'insert', kind: 'rect', box: { x: 10 * i, y: 10, w: 50, h: 50 } }]))
        else if (x < 0.76 && el) run(who.client.call('pptx.shapes', si, [{ op: 'delete', top: el.top! }]))
        else if (x < 0.8) run(who.client.call('pptx.slides', { op: 'add', after: si }))
        else if (x < 0.83) run(who.client.call('pptx.slides', { op: 'move', from: si, to: 0 }))
        else if (x < 0.85 && meta.slides.length > 2) run(who.client.call('pptx.slides', { op: 'delete', index: si }))
        else if (x < 0.95) run(who.client.call('pptx.undo', false))
        else run(who.client.call('pptx.undo', true))
        if (rand() < 0.5) await sleep(Math.floor(rand() * 4))
      }
      await Promise.all(ops)
      await settle([a.collab, b.collab])
      const da = await dump(a.client)
      const db = await dump(b.client)
      expect(texts).toBeGreaterThan(0)
      expect(da).toEqual(db)
      const c = await open(hub, 'Cid', () => 0)
      await settle([a.collab, b.collab, c.collab])
      expect(await dump(c.client)).toEqual(da)
      for (const x of [a, b, c]) {
        x.collab.close()
        x.client.destroy()
      }
    }, 30_000)
  }
})
