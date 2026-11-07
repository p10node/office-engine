/** Writes the iWork demo files: playground/public/samples/sample.pages, .key and .numbers. Usage: bun scripts/iwork-samples.ts */
import { sampleKeynote, sampleNumbers, samplePages } from '../test/iwork-builder'

const dir = `${import.meta.dir}/../playground/public/samples`
await Bun.write(`${dir}/sample.pages`, await samplePages())
await Bun.write(`${dir}/sample.key`, await sampleKeynote())
await Bun.write(`${dir}/sample.numbers`, await sampleNumbers())
console.log('wrote sample.pages, sample.key, sample.numbers in', dir)
