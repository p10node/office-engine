/**
 * A large CSV for load tests: `bun scripts/large-csv.ts 25000000 playground/public/samples/large-rows-25m.csv`
 * (the same columns as fixtures.py's large-rows.csv; large-* samples are not committed). 25M rows ≈ 2.7 GB in ~2 min.
 */
const rows = Number(process.argv[2] ?? 1_000_000)
const out = process.argv[3] ?? 'playground/public/samples/large-rows.csv'
const w = Bun.file(out).writer()
w.write('id,name,email,city,country,amount,qty,date,flag,note\n')
for (let i = 1; i <= rows; i++) {
  w.write(`${i},User ${i},user${i}@example.com,"City, ${i % 500}",VN,${(i * 1.37).toFixed(2)},${i % 97},2026-0${1 + (i % 9)}-1${i % 9},${i % 2 ? 'true' : 'false'},"note ""${i}"""\n`)
}
await w.end()
console.log(`${out}: ${rows.toLocaleString()} rows, ${(Bun.file(out).size / 1048576 / 1024).toFixed(2)} GB`)

export {}
