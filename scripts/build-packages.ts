/**
 * Builds the standalone SDK packages (packages/*): one ESM bundle (shared code compiled in, no dependency) and one
 * self-contained index.d.ts each. Usage: bun scripts/build-packages.ts [name…]
 */
import { spawnSync } from 'node:child_process'
import { copyFileSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { gzipSync } from 'node:zlib'
import { build } from 'vite'

const ALL = ['diagram', 'mermaid', 'plantuml', 'graphviz', 'markdown']
const names = process.argv.slice(2).length ? process.argv.slice(2) : ALL
const root = resolve(import.meta.dir, '..')
const abs = (p: string) => resolve(root, p)

for (const name of names) {
  if (!ALL.includes(name)) throw new Error(`Unknown package “${name}”: ${ALL.join(', ')}`)
  const dir = abs(`packages/${name}`)
  rmSync(`${dir}/dist`, { recursive: true, force: true })
  await build({
    configFile: false,
    root,
    logLevel: 'warn',
    build: {
      outDir: `${dir}/dist`,
      emptyOutDir: true,
      sourcemap: true,
      target: 'es2022',
      minify: true,
      lib: { entry: `${dir}/src/index.ts`, formats: ['es'], fileName: () => 'index.js' },
    },
  })
  const js = readFileSync(`${dir}/dist/index.js`, 'utf8')
  if (/^\s*import\s|\bimport\s*\(/m.test(js)) throw new Error(`${name}: the bundle imports another module`)

  // types: each entry declares its whole API, so its .d.ts stands alone
  const tmp = abs(`node_modules/.cache/pkg-types/${name}`)
  rmSync(tmp, { recursive: true, force: true })
  mkdirSync(tmp, { recursive: true })
  const cfg = `${tmp}/tsconfig.json`
  writeFileSync(
    cfg,
    JSON.stringify({
      extends: abs('tsconfig.json'),
      compilerOptions: { noEmit: false, declaration: true, emitDeclarationOnly: true, outDir: `${tmp}/out`, rootDir: root, types: [] },
      include: [],
      files: [`${dir}/src/index.ts`],
    }),
  )
  const tsc = spawnSync(abs('node_modules/.bin/tsc'), ['-p', cfg], { stdio: 'inherit' })
  if (tsc.status) throw new Error(`${name}: tsc failed`)
  const dts = readFileSync(`${tmp}/out/packages/${name}/src/index.d.ts`, 'utf8')
  if (/^\s*import\s|\bfrom\s+['"]/m.test(dts)) throw new Error(`${name}: index.d.ts refers to other files`)
  writeFileSync(`${dir}/dist/index.d.ts`, dts)
  copyFileSync(abs('LICENSE'), `${dir}/LICENSE`)
  const size = statSync(`${dir}/dist/index.js`).size
  console.log(`@p10node/${name}: ${(size / 1024).toFixed(1)} KB, ${(gzipSync(js).length / 1024).toFixed(1)} KB gzip`)
}
