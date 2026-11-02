/** The reference server end to end: directory storage, snapshots, restart, version history. */
import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { startCollab, type Collab } from '../src/collab'
import { DocumentClient } from '../src/core/worker/client'

const SAMPLE = `${import.meta.dir}/../playground/public/samples/sample.xlsx`
const ROOT = mkdtempSync(join(tmpdir(), 'oe-collab-'))
const DATA = join(ROOT, 'data')
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))
let server: ReturnType<typeof Bun.spawn> | null = null

afterAll(() => {
  server?.kill()
  rmSync(ROOT, { recursive: true, force: true })
})

async function start(port: number) {
  server?.kill()
  await server?.exited
  server = Bun.spawn(['bun', 'scripts/collab-server.ts'], {
    cwd: `${import.meta.dir}/..`,
    env: { ...process.env, PORT: String(port), DATA, SNAPSHOT_EVERY: '3' },
    stdout: 'ignore',
    stderr: 'inherit',
  })
  for (let i = 0; i < 100; i++) {
    try {
      if ((await fetch(`http://localhost:${port}`)).ok) return `http://localhost:${port}`
    } catch {
      // not listening yet
    }
    await sleep(50)
  }
  throw new Error('collab server did not start')
}

async function enter(http: string, name: string) {
  const { snapshot } = (await (await fetch(`${http}/rooms/r1`)).json()) as { snapshot: number }
  const client = await DocumentClient.open(`${http}/rooms/r1/file?seq=${snapshot}`, { worker: false })
  await client.call('xlsx.load', 0)
  await sleep(30)
  const collab = startCollab(client, { transport: http.replace('http', 'ws') + '/collab', room: 'r1', user: { name }, base: snapshot })
  for (let i = 0; i < 300 && collab.status !== 'live'; i++) await sleep(10)
  expect(collab.status).toBe('live')
  return { client, collab, snapshot }
}

async function settle(cs: Collab[]) {
  for (let i = 0; i < 300; i++) {
    if (cs.every((c) => c.pending === 0) && new Set(cs.map((c) => c.session.rev)).size === 1) return
    await sleep(10)
  }
  throw new Error('did not settle')
}

/** column J of rows 0..n-1 */
async function colJ(client: DocumentClient, n: number) {
  const b = await client.call('xlsx.block', 0, 0, n)
  return Array.from({ length: n }, (_, r) => b.rows.find((x) => x.r === r)?.cells.find((c) => c[0] === 9)?.[1])
}

describe('reference collab server', () => {
  test('edits are stored as an op log, snapshotted, survive a restart and form a version history', async () => {
    const port = 19_000 + Math.floor(Math.random() * 1000)
    let http = await start(port)
    const put = await fetch(`${http}/rooms/r1/file`, { method: 'PUT', body: Bun.file(SAMPLE), headers: { 'X-File-Name': 'sample.xlsx' } })
    expect(await put.json()).toMatchObject({ room: 'r1', seq: 0 })

    const a = await enter(http, 'Ann')
    const b = await enter(http, 'Bob')
    for (let i = 0; i < 8; i++) {
      await (i % 2 ? b : a).client.call('xlsx.set', 0, [{ r: i, c: 9, input: `v${i}` }])
      await settle([a.collab, b.collab])
    }
    const want = await colJ(a.client, 9)
    expect(want).toEqual(['v0', 'v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', undefined])
    expect(await colJ(b.client, 9)).toEqual(want)
    for (const p of [a, b]) {
      p.collab.close()
      p.client.destroy()
    }
    // the last editor left: an idle snapshot catches up with every edit
    let meta: { seq: number; snapshot: number } = { seq: -1, snapshot: -2 }
    for (let i = 0; i < 200 && meta.snapshot !== meta.seq; i++) {
      meta = await (await fetch(`${http}/rooms/r1`)).json()
      await sleep(50)
    }
    expect(meta).toMatchObject({ seq: 8, snapshot: 8 })

    // restart: rooms come back from the directory; a newcomer opens snapshot 8 and replays nothing
    http = await start(port + 1)
    const c = await enter(http, 'Cid')
    expect(c.snapshot).toBe(8)
    expect(c.collab.session.rev).toBe(8)
    expect(await colJ(c.client, 9)).toEqual(want)
    await c.client.call('xlsx.set', 0, [{ r: 8, c: 9, input: 'after restart' }])
    for (let i = 0; i < 100 && c.collab.pending; i++) await sleep(10)
    expect(c.collab.session.rev).toBe(9)

    const history = (await (await fetch(`${http}/rooms/r1/history`)).json()) as { seq: number; snapshots: { seq: number }[]; edits: { seq: number; user: string; time: number }[] }
    expect(history.seq).toBe(9)
    expect(history.snapshots.map((s) => s.seq)).toContain(0)
    expect(history.snapshots.map((s) => s.seq)).toContain(8)
    expect(history.edits.map((e) => e.seq)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9])
    expect(new Set(history.edits.map((e) => e.user))).toEqual(new Set(['Ann', 'Bob', 'Cid']))
    expect(history.edits.every((e) => e.time > 0)).toBe(true)

    // any version: a snapshot, or one built on demand between snapshots
    for (const seq of [0, 5, 8]) {
      const res = await fetch(`${http}/rooms/r1/file?seq=${seq}`, { method: 'HEAD' })
      expect(res.headers.get('X-Collab-Seq')).toBe(String(seq))
      const old = await DocumentClient.open(`${http}/rooms/r1/file?seq=${seq}`, { worker: false })
      await old.call('xlsx.load', 0)
      await sleep(30)
      expect(await colJ(old, 9)).toEqual(want.map((v, i) => (i < seq ? v : undefined)))
      old.destroy()
    }
    expect((await fetch(`${http}/rooms/r1/file?seq=99`)).status).toBe(404)
    expect((await fetch(`${http}/rooms/missing`)).status).toBe(404)

    // room ids stay inside the data directory
    await fetch(`${http}/rooms/..%2F..%2Fescape/file`, { method: 'PUT', body: 'x' })
    expect(existsSync(join(ROOT, 'escape'))).toBe(false)
    expect(readdirSync(ROOT)).toEqual(['data'])
    c.collab.close()
    c.client.destroy()
  }, 60_000)
})
