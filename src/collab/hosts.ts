/**
 * Views register their live editing objects here so a collaboration session
 * (created independently, e.g. by DocViewer) can reach them. Tiny on purpose:
 * views import only this file.
 */

type Fn = (host: unknown) => void

const hosts = new WeakMap<object, Map<string, unknown>>()
const waiters = new WeakMap<object, Map<string, Set<Fn>>>()

/** registers `host` for (client, kind); returns the unregister function */
export function registerHost<T>(client: object, kind: string, host: T): () => void {
  let m = hosts.get(client)
  if (!m) hosts.set(client, (m = new Map()))
  m.set(kind, host)
  for (const fn of waiters.get(client)?.get(kind) ?? []) fn(host)
  return () => {
    if (m!.get(kind) === host) m!.delete(kind)
  }
}

export function getHost<T>(client: object, kind: string): T | undefined {
  return hosts.get(client)?.get(kind) as T | undefined
}

/** calls `fn` with the current host and every host registered later */
export function watchHost<T>(client: object, kind: string, fn: (host: T) => void): () => void {
  let w = waiters.get(client)
  if (!w) waiters.set(client, (w = new Map()))
  let set = w.get(kind)
  if (!set) w.set(kind, (set = new Set()))
  set.add(fn as Fn)
  const cur = getHost<T>(client, kind)
  if (cur) fn(cur)
  return () => set!.delete(fn as Fn)
}

// ---------------------------------------------------------------- presence channel (views ↔ session)

export interface PeerView {
  site: number
  name: string
  color: string
  presence?: unknown
}

interface Channel {
  peers: PeerView[]
  peerFns: Set<() => void>
  localFns: Set<(data: unknown) => void>
}

const channels = new WeakMap<object, Channel>()

function channel(client: object): Channel {
  let c = channels.get(client)
  if (!c) channels.set(client, (c = { peers: NO_PEERS, peerFns: new Set(), localFns: new Set() }))
  return c
}

/** session → views: collaborators and where they are */
export function setPeers(client: object, peers: PeerView[]) {
  const c = channel(client)
  c.peers = peers
  for (const fn of c.peerFns) fn()
}

const NO_PEERS: PeerView[] = []

/** stable between changes (usable as a useSyncExternalStore snapshot) */
export function getPeers(client: object): PeerView[] {
  return channels.get(client)?.peers ?? NO_PEERS
}

export function subscribePeers(client: object, fn: () => void): () => void {
  const c = channel(client)
  c.peerFns.add(fn)
  return () => c.peerFns.delete(fn)
}

/** views → session: the local user's position (selection, active slide…) */
export function publishLocal(client: object, data: unknown) {
  for (const fn of channels.get(client)?.localFns ?? []) fn(data)
}

export function onLocal(client: object, fn: (data: unknown) => void): () => void {
  const c = channel(client)
  c.localFns.add(fn)
  return () => c.localFns.delete(fn)
}
