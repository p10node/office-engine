import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import type { Plugin, Rolldown } from 'vite'

// Keeps MIT/BSD-style notice requirements for code bundled from node_modules:
// - prefixes each output chunk with a /*! ... */ list of bundled packages (postBanner runs after minify)
// - writes their full license texts to one file shared by every build (lib, embed, worker)

interface Pkg {
  name: string
  version: string
  license: string
  text: string
}

const SEPARATOR = '\n\n' + '-'.repeat(80) + '\n\n'
const pkgCache = new Map<string, Pkg | null>()

function pkgDirOf(id: string): string | null {
  const clean = id.split('?')[0]
  const marker = '/node_modules/'
  const at = clean.lastIndexOf(marker)
  if (clean.startsWith('\0') || at < 0) return null
  const rest = clean.slice(at + marker.length).split('/')
  const name = rest[0].startsWith('@') ? `${rest[0]}/${rest[1]}` : rest[0]
  return clean.slice(0, at + marker.length) + name
}

function readPkg(dir: string): Pkg | null {
  if (pkgCache.has(dir)) return pkgCache.get(dir)!
  let pkg: Pkg | null = null
  const manifest = join(dir, 'package.json')
  if (existsSync(manifest)) {
    const json = JSON.parse(readFileSync(manifest, 'utf8'))
    const licenseFile = readdirSync(dir).find((f) => /^licen[cs]e(\.(md|txt))?$/i.test(f))
    const license = typeof json.license === 'string' ? json.license : (json.license?.type ?? 'UNKNOWN')
    pkg = {
      name: json.name,
      version: json.version,
      license,
      text: licenseFile ? readFileSync(join(dir, licenseFile), 'utf8').trim() : `License: ${license} (no license file shipped)`,
    }
  }
  pkgCache.set(dir, pkg)
  return pkg
}

function bundledPkgs(chunk: Rolldown.RenderedChunk): Pkg[] {
  const byName = new Map<string, Pkg>()
  for (const [id, mod] of Object.entries(chunk.modules)) {
    if (!mod.renderedLength) continue
    const dir = pkgDirOf(id)
    const pkg = dir && readPkg(dir)
    if (pkg) byName.set(`${pkg.name}@${pkg.version}`, pkg)
  }
  return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name))
}

function mergeInto(file: string, pkgs: Pkg[]): void {
  const sections = new Map<string, string>()
  if (existsSync(file)) {
    const [, ...prev] = readFileSync(file, 'utf8').split(SEPARATOR)
    for (const s of prev) sections.set(s.split('\n', 1)[0], s.trim())
  }
  for (const p of pkgs) {
    const head = `${p.name}@${p.version} (${p.license})`
    sections.set(head, `${head}\n\n${p.text}`)
  }
  if (!sections.size) return
  const header = 'Third-party software bundled in this package. Each notice below applies to the named package.'
  const body = [...sections.keys()].sort().map((k) => sections.get(k)!)
  mkdirSync(dirname(file), { recursive: true })
  writeFileSync(file, [header, ...body].join(SEPARATOR) + '\n')
}

export function thirdPartyLicenses(outFile = 'dist/THIRD_PARTY_LICENSES.txt'): Plugin {
  const file = resolve(outFile)
  const name = outFile.split('/').pop()
  return {
    name: 'third-party-licenses',
    outputOptions(opts) {
      const prev = opts.postBanner
      opts.postBanner = async (chunk) => {
        const before = typeof prev === 'function' ? await prev(chunk) : (prev ?? '')
        const pkgs = bundledPkgs(chunk)
        if (!pkgs.length) return before
        const list = pkgs.map((p) => `${p.name}@${p.version} (${p.license})`).join(', ')
        return [before, `/*! Bundles ${list}. License texts: ${name} */`].filter(Boolean).join('\n')
      }
      return opts
    },
    generateBundle(_, bundle) {
      const pkgs = Object.values(bundle).flatMap((out) => (out.type === 'chunk' ? bundledPkgs(out) : []))
      mergeInto(file, pkgs)
    },
  }
}
