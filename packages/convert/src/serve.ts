/** Listens for HTTP requests with the handler of server.ts: Bun.serve under Bun, the http module under Node. */
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import { Readable } from 'node:stream'

export interface Listening {
  url: string
  close(): Promise<void>
}

type Handler = (req: Request) => Promise<Response>

export async function listen(handler: Handler, opts: { port?: number; host?: string; maxBytes?: number } = {}): Promise<Listening> {
  const port = opts.port ?? 3000
  const host = opts.host ?? '0.0.0.0'
  const bun = (globalThis as { Bun?: { serve(o: object): { url: URL; stop(force?: boolean): Promise<void> } } }).Bun
  if (bun) {
    // (a little over the limit: the handler answers 413 itself, with its message)
    const server = bun.serve({ port, hostname: host, fetch: handler, maxRequestBodySize: (opts.maxBytes ?? 100 * 1024 * 1024) + 1024 * 1024 })
    return { url: server.url.href, close: () => server.stop(true) }
  }
  const server = createServer((req, res) => void handle(handler, req, res))
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(port, host, resolve)
  })
  const addr = server.address()
  const at = typeof addr === 'object' && addr ? `${addr.family === 'IPv6' ? `[${addr.address}]` : addr.address}:${addr.port}` : `${host}:${port}`
  return {
    url: `http://${at}/`,
    close: () =>
      new Promise((resolve) => {
        server.closeAllConnections()
        server.close(() => resolve())
      }),
  }
}

async function handle(handler: Handler, req: IncomingMessage, res: ServerResponse) {
  // the client going away gives up its conversions
  const ac = new AbortController()
  res.on('close', () => {
    if (!res.writableFinished) ac.abort(new DOMException('The client closed the connection', 'AbortError'))
  })
  try {
    const headers = new Headers()
    for (const [k, v] of Object.entries(req.headers)) {
      if (Array.isArray(v)) for (const x of v) headers.append(k, x)
      else if (v !== undefined) headers.set(k, v)
    }
    const method = req.method ?? 'GET'
    const hasBody = method !== 'GET' && method !== 'HEAD'
    const request = new Request(`http://${req.headers.host ?? 'localhost'}${req.url ?? '/'}`, {
      method,
      headers,
      signal: ac.signal,
      ...(hasBody ? { body: Readable.toWeb(req), duplex: 'half' } : {}),
    } as RequestInit)
    const response = await handler(request)
    res.statusCode = response.status
    // (a header given more than once comes joined with ", ": the warnings are percent-encoded, so they stay apart)
    response.headers.forEach((v, k) => res.setHeader(k, v))
    if (!response.body || method === 'HEAD') return void res.end()
    for await (const chunk of response.body as unknown as AsyncIterable<Uint8Array>) res.write(chunk)
    res.end()
  } catch (e) {
    if (!res.headersSent) res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' })
    res.end(`${(e as Error)?.message ?? e}\n`)
  }
}
