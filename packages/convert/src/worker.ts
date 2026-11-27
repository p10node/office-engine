/** A conversion thread of the pool (pool.ts): one file at a time, the bytes moved in and out, never copied. */
import { parentPort } from 'node:worker_threads'
import { convert, ConvertError, type ConvertOptions } from './convert'

export interface Job {
  id: number
  data: Uint8Array
  /** `sharedFonts`: the fonts of the last `FontsMessage` */
  opts: Omit<ConvertOptions, 'signal'> & { sharedFonts?: boolean }
}

/** the pool's font list, sent to a thread once */
export interface FontsMessage {
  fonts: NonNullable<ConvertOptions['fonts']>
}

export type JobResult =
  | { id: number; ok: true; result: Awaited<ReturnType<typeof convert>> }
  | { id: number; ok: false; error: string; code?: ConvertError['code'] }

let shared: ConvertOptions['fonts']

parentPort!.on('message', async (msg: Job | FontsMessage) => {
  if (!('id' in msg)) {
    shared = msg.fonts
    return
  }
  const { id, data } = msg
  const { sharedFonts, ...opts } = msg.opts
  try {
    const result = await convert(data, sharedFonts ? { ...opts, fonts: shared } : opts)
    parentPort!.postMessage({ id, ok: true, result } satisfies JobResult, [result.data.buffer as ArrayBuffer])
  } catch (e) {
    const err = e as Error
    parentPort!.postMessage({ id, ok: false, error: err?.message ?? String(e), ...(e instanceof ConvertError ? { code: e.code } : {}) } satisfies JobResult)
  }
})
