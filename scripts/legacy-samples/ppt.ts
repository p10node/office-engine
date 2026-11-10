/**
 * Writes test/fixtures/legacy/sample.ppt, the PowerPoint 97-2003 sample built by test/ppt-builder.ts (the content of
 * the playground's PPTX sample). Usage: bun scripts/legacy-samples/ppt.ts
 */
import { samplePpt } from '../../test/ppt-builder'

await Bun.write(`${import.meta.dir}/../../test/fixtures/legacy/sample.ppt`, samplePpt())
