/**
 * Conversions on worker threads: a file that takes too long, or a caller that gives up, ends its thread (and with it
 * whatever the engine was doing), and the server's own thread is never blocked by a large file.
 */
import { availableParallelism } from 'node:os'
import { Worker } from 'node:worker_threads'
import { convert, ConvertError, type ConvertOptions, type Converted } from './convert'
import { systemFonts, type FontFile } from './fonts'
import type { FontsMessage, Job, JobResult } from './worker'

export interface ConverterOptions {
  /** threads converting at once (default: the number of CPUs); 0 converts on this thread, without time limits */
  workers?: number
  /** conversions waiting for a thread before more are refused with a `BusyError` (default 100) */
  queue?: number
  /** bytes of input running and waiting before more are refused with a `BusyError` (default 1 GB; one file is always taken when none is) */
  bytes?: number
  /** milliseconds a conversion may take before its thread is ended (default 120000; 0: no limit) */
  timeout?: number
  /** megabytes of JavaScript heap a thread may use (Node only: Bun has no such limit) */
  memory?: number
  /**
   * the font files PDFs set their text in, for a conversion that names none (default the system's, found once); a
   * function is called when a PDF is first asked for. A list is taken as it is when the converter is made
   */
  fonts?: FontFile[] | Promise<FontFile[]> | (() => Promise<FontFile[]>)
}

/** the conversion took longer than `timeout` */
export class TimeoutError extends Error {
  constructor(ms: number) {
    super(`The conversion took longer than ${ms / 1000} s`)
    this.name = 'TimeoutError'
  }
}

/** the thread converting the file stopped (out of memory, a crash) */
export class ThreadError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'ThreadError'
  }
}

/** every thread is busy and the queue is full */
export class BusyError extends Error {
  constructor() {
    super('The converter is busy: try again later')
    this.name = 'BusyError'
  }
}

interface Task {
  data: Uint8Array
  opts: Omit<ConvertOptions, 'signal'>
  signal?: AbortSignal
  resolve: (r: Converted) => void
  reject: (e: Error) => void
  /** while it waits: what takes it out of the queue when its signal aborts */
  onAbort?: () => void
}

interface Thread {
  worker: Worker
  /** the thread has the pool's font list (sent once, not with each conversion) */
  fonts?: boolean
  task?: Task
  timer?: ReturnType<typeof setTimeout>
  abort?: () => void
}

// the source (`bun packages/convert/src/cli.ts`) or the built package (dist/index.js next to dist/worker.js)
const WORKER_URL = new URL(import.meta.url.endsWith('.ts') ? './worker.ts' : './worker.js', import.meta.url)

/** Converts files on a pool of worker threads, one file per thread at a time. */
export class Converter {
  readonly workers: number
  readonly timeout: number
  private readonly queueMax: number
  private readonly bytesMax: number
  private readonly memory?: number
  /** input bytes of the conversions running and waiting */
  private bytes = 0
  private tasks = 0
  private readonly idle: Thread[] = []
  private readonly busy = new Set<Thread>()
  private readonly waiting: Task[] = []
  private seq = 0
  private closed = false

  constructor(opts: ConverterOptions = {}) {
    this.workers = Math.max(0, Math.floor(opts.workers ?? availableParallelism()))
    this.timeout = opts.timeout ?? 120_000
    this.queueMax = opts.queue ?? 100
    this.bytesMax = opts.bytes ?? 1024 * 1024 * 1024
    this.memory = opts.memory
    this.fontSpec = opts.fonts
    // (a copy: the threads that have it and the ones sent it later have the same list)
    if (Array.isArray(opts.fonts)) this.catalog = [...opts.fonts]
  }

  private fontSpec: ConverterOptions['fonts']
  /** found once, here, not in each thread */
  private fontList?: Promise<FontFile[]>
  /** the list once found: a thread is sent it once */
  private catalog?: FontFile[]
  /** PDF conversions waiting for the font list: queued work, given up with their caller or the converter */
  private fontWaits = new Set<(e: unknown) => void>()

  /** conversions running and waiting */
  get load(): { running: number; queued: number } {
    return { running: this.busy.size, queued: this.waiting.length + this.fontWaits.size }
  }

  /** the font list, found when first asked for; one that could not be found is looked for again next time */
  private fonts(): Promise<FontFile[]> {
    if (this.catalog) return Promise.resolve(this.catalog)
    if (!this.fontList) {
      const spec = this.fontSpec
      const p = Promise.resolve().then(() => (typeof spec === 'function' ? spec() : (spec ?? systemFonts())))
      this.fontList = p.then((list) => (this.catalog = list))
      this.fontList.catch(() => (this.fontList = undefined))
    }
    return this.fontList
  }

  convert(data: Uint8Array, opts: ConvertOptions = {}): Promise<Converted> {
    if (this.closed) return Promise.reject(new Error('The converter is closed'))
    if (opts.to === 'pdf' && !opts.fonts) {
      if (this.catalog) return this.convert(data, { ...opts, fonts: this.catalog })
      // the font list is being looked for: the conversion waits for it as queued work does, its input counted
      if (opts.signal?.aborted) return Promise.reject(opts.signal.reason)
      const size = data.byteLength
      const free = this.idle.length + (this.workers - this.busy.size) - this.fontWaits.size
      if (free <= 0 && this.waiting.length + this.fontWaits.size >= this.queueMax) return Promise.reject(new BusyError())
      if (this.tasks && this.bytes + size > this.bytesMax) return Promise.reject(new BusyError())
      this.bytes += size
      this.tasks++
      const signal = opts.signal
      // (let go of when the conversion is given up: an input is not held until the font folders are read)
      let input: Uint8Array | null = data
      return new Promise<Converted>((resolve, reject) => {
        const settle = () => {
          this.fontWaits.delete(fail)
          signal?.removeEventListener('abort', onAbort)
          this.bytes -= size
          this.tasks--
          const d = input!
          input = null
          return d
        }
        const fail = (e: unknown) => {
          settle()
          reject(e)
        }
        const onAbort = () => fail(signal!.reason)
        this.fontWaits.add(fail)
        signal?.addEventListener('abort', onAbort, { once: true })
        this.fonts().then(
          (fonts) => {
            // (given up while the fonts were looked for)
            if (!this.fontWaits.has(fail)) return
            this.convert(settle(), { ...opts, fonts }).then(resolve, reject)
          },
          (e) => this.fontWaits.has(fail) && fail(e),
        )
      })
    }
    if (!this.workers) return convert(data, opts)
    const { signal, ...rest } = opts
    if (signal?.aborted) return Promise.reject(signal.reason)
    const size = data.byteLength
    if (this.tasks && this.bytes + size > this.bytesMax) return Promise.reject(new BusyError())
    const idle = this.idle.length > 0 || this.busy.size < this.workers
    if (!idle && this.waiting.length >= this.queueMax) return Promise.reject(new BusyError())
    this.bytes += size
    this.tasks++
    const done = () => {
      this.bytes -= size
      this.tasks--
    }
    return new Promise<Converted>((resolve, reject) => {
      const task: Task = {
        data,
        opts: rest,
        signal,
        resolve: (r) => (done(), resolve(r)),
        reject: (e) => (done(), reject(e)),
      }
      let thread: Thread | undefined
      try {
        thread = this.idle.pop() ?? (this.busy.size < this.workers ? this.spawn() : undefined)
      } catch (e) {
        // (a thread that cannot start: the conversion fails, its bytes are given back)
        return task.reject(e as Error)
      }
      if (thread) return this.run(thread, task)
      this.waiting.push(task)
      if (signal) {
        task.onAbort = () => {
          const at = this.waiting.indexOf(task)
          if (at < 0) return
          this.waiting.splice(at, 1)
          task.reject(signal.reason)
        }
        signal.addEventListener('abort', task.onAbort, { once: true })
      }
    })
  }

  /** the next conversion waiting, its abort listener gone */
  private dequeue(): Task | undefined {
    const task = this.waiting.shift()
    if (task?.onAbort) task.signal?.removeEventListener('abort', task.onAbort)
    return task
  }

  /** ends every thread; conversions under way and waiting are rejected */
  async close(): Promise<void> {
    this.closed = true
    for (const fail of [...this.fontWaits]) fail(new Error('The converter is closed'))
    for (let t = this.dequeue(); t; t = this.dequeue()) t.reject(new Error('The converter is closed'))
    const all = [...this.idle.splice(0), ...this.busy]
    this.busy.clear()
    for (const t of all) this.end(t, new Error('The converter is closed'))
    await Promise.all(all.map((t) => t.worker.terminate()))
  }

  private spawn(): Thread {
    const worker = new Worker(WORKER_URL, this.memory ? { resourceLimits: { maxOldGenerationSizeMb: this.memory } } : {})
    const thread: Thread = { worker }
    worker.on('message', (msg: JobResult) => {
      const task = thread.task
      if (!task) return
      this.release(thread)
      if (msg.ok) task.resolve({ ...msg.result, data: new Uint8Array(msg.result.data) })
      else task.reject(msg.code ? new ConvertError(msg.error, msg.code) : new Error(msg.error))
    })
    // a thread that dies (out of memory, a crash) takes its conversion with it; the next one gets a new thread
    worker.on('error', (e: Error) => this.discard(thread, new ThreadError(`The conversion thread failed: ${e.message}`)))
    worker.on('exit', () => this.discard(thread, new ThreadError('The conversion thread stopped')))
    return thread
  }

  private run(thread: Thread, task: Task): void {
    if (task.signal?.aborted) {
      task.reject(task.signal.reason)
      return this.next(thread)
    }
    thread.task = task
    this.busy.add(thread)
    // (a thread at work keeps the process alive, an idle one does not)
    thread.worker.ref()
    if (this.timeout) thread.timer = setTimeout(() => this.kill(thread, new TimeoutError(this.timeout)), this.timeout)
    if (task.signal) {
      const signal = task.signal
      thread.abort = () => this.kill(thread, signal.reason instanceof Error ? signal.reason : new Error('The conversion was aborted'))
      signal.addEventListener('abort', thread.abort, { once: true })
    }
    const id = ++this.seq
    try {
      // the pool's font list goes to a thread once; a conversion then says it uses it
      let opts = task.opts
      if (opts.fonts && opts.fonts === this.catalog) {
        if (!thread.fonts) {
          thread.worker.postMessage({ fonts: opts.fonts } satisfies FontsMessage)
          thread.fonts = true
        }
        opts = { ...opts, fonts: undefined, sharedFonts: true } as typeof opts
      }
      // a copy of its own moves to the thread: the caller's bytes stay as they are (a Node Buffer may share its memory with others)
      const data = new Uint8Array(task.data)
      thread.worker.postMessage({ id, data, opts } satisfies Job, [data.buffer as ArrayBuffer])
    } catch (e) {
      // (options that cannot be sent: the thread takes the next conversion)
      this.release(thread)
      task.reject(e as Error)
    }
  }

  /** the thread is done with its task: it takes the next one, or waits */
  private release(thread: Thread) {
    clearTimeout(thread.timer)
    if (thread.abort) thread.task?.signal?.removeEventListener('abort', thread.abort)
    thread.task = thread.timer = thread.abort = undefined
    this.busy.delete(thread)
    this.next(thread)
  }

  private next(thread: Thread): void {
    if (this.closed) return
    const task = this.dequeue()
    if (task) return this.run(thread, task)
    thread.worker.unref()
    this.idle.push(thread)
  }

  /** ends a thread in the middle of its conversion; a new one takes the queue on */
  private kill(thread: Thread, err: Error) {
    this.discard(thread, err)
    void thread.worker.terminate()
  }

  private discard(thread: Thread, err: Error) {
    const known = this.busy.delete(thread) || this.idle.includes(thread)
    const at = this.idle.indexOf(thread)
    if (at >= 0) this.idle.splice(at, 1)
    this.end(thread, err)
    if (!known || this.closed) return
    const task = this.dequeue()
    if (!task) return
    try {
      this.run(this.spawn(), task)
    } catch (e) {
      task.reject(e as Error)
    }
  }

  private end(thread: Thread, err: Error) {
    const task = thread.task
    clearTimeout(thread.timer)
    if (thread.abort) task?.signal?.removeEventListener('abort', thread.abort)
    thread.task = thread.timer = thread.abort = undefined
    task?.reject(err)
  }
}
