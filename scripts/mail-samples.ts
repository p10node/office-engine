/** Writes the mail demo files: playground/public/samples/sample.eml and sample.msg (built by test/mail-builder.ts). Usage: bun scripts/mail-samples.ts */
import { sampleEml, sampleMsg } from '../test/mail-builder'

const dir = `${import.meta.dir}/../playground/public/samples`
await Bun.write(`${dir}/sample.eml`, sampleEml())
await Bun.write(`${dir}/sample.msg`, sampleMsg())
console.log('wrote sample.eml and sample.msg in', dir)
