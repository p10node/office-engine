/**
 * Comments in the viewer: the threads the worker keeps, the panel state (open, active thread, a new thread being
 * written) and the actions. Views publish where comments go (`CommentHost`): the selection (DOCX), the active cell
 * (XLSX), the selected shape (PPTX), and where each thread is.
 */
import { createContext, useContext, useSyncExternalStore } from 'react'
import { initialsOf, nowIso, newCommentId, threadsOf, type CommentAnchor, type CommentThread, type CommentUser, type DocComment } from '../../core/comments'
import type { DocKind } from '../../core/opc'
import type { DocumentClient } from '../../core/worker/client'

export interface CommentOptions {
  /** who writes the comments made here (default: the collaboration user, else "Me") */
  user?: CommentUser
  /** who can be @mentioned: a list, or a function of what was typed after "@" */
  mentions?: CommentUser[] | ((query: string) => CommentUser[] | Promise<CommentUser[]>)
  /** whether comments can be added, answered and resolved here (default: while editing) */
  canComment?: boolean
  /** each comment change made here (not those of collaborators): notify the people mentioned, log activity… */
  onEvent?: (e: CommentEvent) => void
  /** open the comments panel when the document opens */
  open?: boolean
  /** a thread (or a comment of it) to show when the document opens: the panel opens on it and the document scrolls to it */
  focus?: string
}

export type CommentAction = 'add' | 'reply' | 'edit' | 'delete' | 'resolve' | 'reopen'

export interface CommentEvent {
  type: CommentAction
  /** the comment added / changed / deleted (the thread's root for resolve / reopen) */
  comment: DocComment
  /** the root of its thread */
  thread: DocComment
  /** people mentioned by this change: all of a new comment's, the ones an edit added */
  mentioned: CommentUser[]
  /** the people who wrote the thread's other comments (once each, not the one who made this change) */
  participants: CommentUser[]
  /** what the thread points at: the commented text (DOCX), "Sheet1!B3" (XLSX), "Slide 3" (PPTX) */
  quote: string
}

/** where a thread being written would go */
export interface PendingComment {
  quote: string
  /** XLSX / PPTX: the thread's anchor */
  at?: CommentAnchor
  /** DOCX: the range to mark (view-specific) */
  range?: unknown
  /** the place already has a thread (XLSX cell): commenting answers it */
  existing?: string
}

export interface PlacedThread extends CommentThread {
  quote: string
}

/** what a view tells the comments panel */
export interface CommentHost {
  kind: DocKind
  /** where a new thread would go now; null when there is nothing to comment on */
  pending(): PendingComment | null
  /** puts a new thread in place (DOCX marks its text); false when it can't */
  attach?(id: string, p: PendingComment): boolean | Promise<boolean>
  /** the threads that have a place in the document, in document order */
  placed(threads: CommentThread[]): PlacedThread[]
  /** shows thread `id`'s place */
  reveal(id: string): void
  /** what to do to comment, when `pending()` finds nothing */
  hint: string
}

export interface CommentsState {
  rev: number
  open: boolean
  /** thread shown as the current one (its place highlighted) */
  active: string | null
  /** a new thread being written */
  draft: PendingComment | null
  showResolved: boolean
  /** message about the last action that failed, or why a thread can't be started */
  notice: string | null
}

export class CommentsCtl {
  readonly list = new Map<string, DocComment>()
  state: CommentsState = { rev: 0, open: false, active: null, draft: null, showResolved: false, notice: null }
  host: CommentHost | null = null
  opts: CommentOptions = {}
  me: CommentUser = { name: 'Me' }
  canComment = false
  loaded = false
  private subs = new Set<() => void>()
  private offs: (() => void)[] = []
  private touchTimer?: ReturnType<typeof setTimeout>
  private placedCache: { rev: number; host: CommentHost | null; out: PlacedThread[] } | null = null

  constructor(readonly client: DocumentClient) {
    this.offs.push(
      client.on('comment', ({ id, c }) => {
        if (c) this.list.set(id, c)
        else this.list.delete(id)
        if (!c && this.state.active === id) this.set({ active: null })
        else this.set({})
      }),
    )
    client
      .call('comments.list')
      .then((all) => {
        for (const c of all) if (!this.list.has(c.id)) this.list.set(c.id, c)
        this.loaded = true
        this.set({})
      })
      .catch((e: Error) => console.warn('[office-engine] comments not loaded:', e.message))
  }

  dispose() {
    this.offs.forEach((f) => f())
    clearTimeout(this.touchTimer)
    this.subs.clear()
  }

  // ---------------------------------------------------------------- store

  subscribe = (fn: () => void) => {
    this.subs.add(fn)
    return () => {
      this.subs.delete(fn)
    }
  }

  snapshot = () => this.state

  private set(p: Partial<CommentsState>) {
    this.state = { ...this.state, ...p, rev: this.state.rev + 1 }
    for (const fn of this.subs) fn()
  }

  /** the document changed: places and quotes are looked up again (soon, not on every keystroke) */
  touch() {
    if (this.touchTimer) return
    this.touchTimer = setTimeout(() => {
      this.touchTimer = undefined
      this.set({})
    }, 250)
  }

  setHost(h: CommentHost | null) {
    this.host = h
    this.set({})
  }

  // ---------------------------------------------------------------- reading

  /** the threads with a place in the document, in document order (resolved ones too) */
  threads(): PlacedThread[] {
    const c = this.placedCache
    if (c && c.rev === this.state.rev && c.host === this.host) return c.out
    const out = this.host ? this.host.placed(threadsOf(this.list.values())) : []
    this.placedCache = { rev: this.state.rev, host: this.host, out }
    return out
  }

  /** threads shown in the document (not resolved) */
  visible(): PlacedThread[] {
    return this.threads().filter((t) => !t.root.done)
  }

  isMine(c: DocComment): boolean {
    return this.me.id ? c.author.id === this.me.id : !c.author.id && c.author.name === this.me.name
  }

  // ---------------------------------------------------------------- panel

  toggle(open = !this.state.open) {
    this.set(open ? { open } : { open, draft: null, notice: null })
  }

  /** makes thread `id` the current one; `reveal`: scroll the document to it */
  select(id: string | null, reveal = true) {
    if (id && !this.state.open) this.set({ open: true, active: id, notice: null })
    else if (id !== this.state.active) this.set({ active: id, notice: null })
    if (id && reveal) this.host?.reveal(id)
  }

  /** starts a new thread on the selection (or at `at`); false (with a notice) when there is nothing to comment on */
  start(at?: PendingComment): boolean {
    if (!this.canComment || !this.host) return false
    const draft = at ?? this.host.pending()
    if (!draft) {
      this.set({ open: true, draft: null, notice: this.host.hint })
      return false
    }
    if (draft.existing) {
      this.select(draft.existing)
      return true
    }
    this.set({ open: true, draft, active: null, notice: null })
    return true
  }

  cancelDraft() {
    if (this.state.draft) this.set({ draft: null })
  }

  dismiss() {
    if (this.state.notice) this.set({ notice: null })
  }

  setShowResolved(on: boolean) {
    this.set({ showResolved: on })
  }

  // ---------------------------------------------------------------- actions

  private author(): CommentUser {
    const u = this.me
    return { name: u.name, ...(u.id ? { id: u.id } : {}), initials: u.initials ?? initialsOf(u.name) }
  }

  private async write(id: string, c: DocComment | null): Promise<boolean> {
    try {
      await this.client.call('comments.set', id, c)
      return true
    } catch (e) {
      this.set({ notice: `Couldn't save the comment: ${(e as Error).message}` })
      return false
    }
  }

  private quoteOf(rootId: string): string {
    return this.threads().find((t) => t.root.id === rootId)?.quote ?? ''
  }

  /** authors of the comments of thread `root` other than `except` and this user, once each */
  private participants(root: string, except: string): CommentUser[] {
    const out = new Map<string, CommentUser>()
    for (const c of this.list.values()) {
      if (c.id === except || (c.id !== root && c.parent !== root) || this.isMine(c)) continue
      const k = c.author.id ?? c.author.name
      if (!out.has(k)) out.set(k, c.author)
    }
    return [...out.values()]
  }

  private emit(type: CommentAction, comment: DocComment, mentioned: CommentUser[] = []) {
    const thread = comment.parent ? (this.list.get(comment.parent) ?? comment) : comment
    try {
      this.opts.onEvent?.({ type, comment, thread, mentioned, participants: this.participants(thread.id, comment.id), quote: this.quoteOf(thread.id) })
    } catch (e) {
      console.error('[office-engine] comment event handler failed', e)
    }
  }

  /** posts the thread being written */
  async post(text: string, mentions: CommentUser[]): Promise<boolean> {
    const draft = this.state.draft
    const host = this.host
    if (!draft || !host || !text.trim()) return false
    const id = newCommentId(host.kind)
    const c: DocComment = { id, author: this.author(), time: nowIso(), text: text.trim(), ...(mentions.length ? { mentions } : {}), ...(draft.at ? { at: draft.at } : {}) }
    if (!(await this.write(id, c))) return false
    // another thread may have been started meanwhile: it stays
    const mine = () => (this.state.draft === draft ? { draft: null } : {})
    if (host.attach && !(await host.attach(id, draft))) {
      await this.write(id, null)
      this.set({ ...mine(), notice: 'The commented text is no longer there' })
      return false
    }
    this.set(this.state.draft === draft ? { draft: null, active: id } : {})
    this.emit('add', c, mentions)
    return true
  }

  async reply(rootId: string, text: string, mentions: CommentUser[]): Promise<boolean> {
    const root = this.list.get(rootId)
    if (!root || !text.trim()) return false
    const id = newCommentId(this.host?.kind ?? 'docx')
    const c: DocComment = { id, parent: rootId, author: this.author(), time: nowIso(), text: text.trim(), ...(mentions.length ? { mentions } : {}) }
    // answering a resolved thread opens it again, as in Word
    if (root.done && !(await this.write(rootId, { ...root, done: undefined }))) return false
    if (!(await this.write(id, c))) return false
    this.emit('reply', c, mentions)
    return true
  }

  async edit(id: string, text: string, mentions: CommentUser[]): Promise<boolean> {
    const c = this.list.get(id)
    if (!c || !text.trim()) return false
    const next: DocComment = { ...c, text: text.trim() }
    if (mentions.length) next.mentions = mentions
    else delete next.mentions
    if (!(await this.write(id, next))) return false
    const had = new Set((c.mentions ?? []).map((m) => m.id ?? m.name))
    this.emit('edit', next, mentions.filter((m) => !had.has(m.id ?? m.name)))
    return true
  }

  /** deletes a comment; deleting a thread's root deletes the thread */
  async remove(id: string): Promise<boolean> {
    const c = this.list.get(id)
    if (!c) return false
    // the event needs the thread as it was
    const quote = this.quoteOf(c.parent ?? id)
    if (!(await this.write(id, null))) return false
    try {
      this.opts.onEvent?.({ type: 'delete', comment: c, thread: c.parent ? (this.list.get(c.parent) ?? c) : c, mentioned: [], participants: this.participants(c.parent ?? id, id), quote })
    } catch (e) {
      console.error('[office-engine] comment event handler failed', e)
    }
    return true
  }

  async resolve(id: string, done: boolean): Promise<boolean> {
    const c = this.list.get(id)
    if (!c || c.parent || !!c.done === done) return false
    const next: DocComment = { ...c }
    if (done) next.done = true
    else delete next.done
    if (!(await this.write(id, next))) return false
    if (done && this.state.active === id && !this.state.showResolved) this.set({ active: null })
    this.emit(done ? 'resolve' : 'reopen', next)
    return true
  }

  /** people matching what was typed after "@" */
  async people(query: string): Promise<CommentUser[]> {
    const m = this.opts.mentions
    if (!m) return []
    const q = query.trim().toLowerCase()
    const all = typeof m === 'function' ? await m(query) : m.filter((u) => !q || u.name.toLowerCase().includes(q) || (u.id ?? '').toLowerCase().includes(q))
    return all.slice(0, 8)
  }
}

export const CommentsContext = createContext<CommentsCtl | null>(null)

export const useCommentsCtl = () => useContext(CommentsContext)

const NONE: CommentsState = { rev: 0, open: false, active: null, draft: null, showResolved: false, notice: null }
const noop = () => () => {}

/** the comments state, re-rendering on every change */
export function useCommentsState(ctl: CommentsCtl | null): CommentsState {
  return useSyncExternalStore(ctl?.subscribe ?? noop, ctl ? ctl.snapshot : () => NONE, ctl ? ctl.snapshot : () => NONE)
}

/** CSS that highlights the threads' places: `[data-cm~="id"]` inside `scope` */
export function markCss(scope: string, ids: string[], active: string | null): string {
  const sel = (id: string) => `${scope} [data-cm~="${id.replace(/["\\]/g, '')}"]`
  let css = ''
  if (ids.length) css += `${ids.map(sel).join(',')}{background-image:linear-gradient(var(--de-cm-bg),var(--de-cm-bg))!important;cursor:pointer}`
  if (active && ids.includes(active)) css += `${sel(active)}{background-image:linear-gradient(var(--de-cm-on),var(--de-cm-on))!important;box-shadow:0 2px 0 var(--de-cm-line)}`
  return css
}
