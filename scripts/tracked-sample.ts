/**
 * playground/public/samples/tracked.docx: sample.docx with tracked changes (two authors), for the review features.
 *   bun scripts/tracked-sample.ts
 */
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'

const dir = `${import.meta.dir}/../playground/public/samples`
const zip = await ZipArchive.open(new BlobSource(Bun.file(`${dir}/sample.docx`) as unknown as Blob))
let doc = await zip.text('word/document.xml')
const swap = (from: string, to: string) => {
  if (!doc.includes(from)) throw new Error(`not found: ${from}`)
  doc = doc.replace(from, to)
}
const LAN = 'w:author="Lan Nguyen" w:date="2026-11-01T09:00:00Z"'
const MINH = 'w:author="Minh Tran" w:date="2026-11-02T14:30:00Z"'
swap(
  '<w:r><w:t>Revenue grew 24% year over year</w:t></w:r>',
  `<w:r><w:t xml:space="preserve">Revenue grew </w:t></w:r><w:del w:id="101" ${LAN}><w:r><w:delText>24%</w:delText></w:r></w:del><w:ins w:id="102" ${LAN}><w:r><w:t>31%</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> year over year</w:t></w:r>`,
)
swap(
  '<w:r><w:t>Two new regions launched</w:t></w:r>',
  `<w:r><w:t xml:space="preserve">Two new regions launched</w:t></w:r><w:ins w:id="103" ${MINH}><w:r><w:t xml:space="preserve"> in Q3</w:t></w:r></w:ins>`,
)
swap(
  '<w:r><w:t>Expand to APAC</w:t></w:r>',
  `<w:r><w:t xml:space="preserve">Expand to </w:t></w:r><w:del w:id="104" ${MINH}><w:r><w:delText>APAC</w:delText></w:r></w:del><w:ins w:id="105" ${MINH}><w:r><w:t>Vietnam</w:t></w:r></w:ins>`,
)
const out = await rebuildZip(zip, new Map([['word/document.xml', doc]]), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
await Bun.write(`${dir}/tracked.docx`, out)
console.log('tracked.docx', out.size)
