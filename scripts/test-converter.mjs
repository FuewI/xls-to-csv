// Run: npm test   — verifies the converter against the sample and synthetic edge cases.
import fs from 'node:fs'
import assert from 'node:assert/strict'
import * as XLSX from 'xlsx'
import {
  parseOrderWorkbook, buildCsv, buildCsvBytes, buildCsvRows, buildXlsxBytes, buildFileName, ConversionError, EXPORT_FORMATS,
} from '../src/lib/converter.js'

let failed = 0
// The real sample order is not committed to the repository (it contains company data).
// Tests that need it are skipped when the file is absent (e.g. in GitHub Actions).
const SAMPLE = 'samples/Order0511202600000000082792000039011.xls'
const hasSample = fs.existsSync(SAMPLE)
const sampleTest = (name, fn) => (hasSample ? test(name, fn) : console.log('-', name, '(skipped: sample file not present)'))
const test = (name, fn) => {
  try { fn(); console.log('✓', name) } catch (e) { failed++; console.error('✗', name, '\n ', e.message) }
}
const toXls = (aoa, merges = [], bookType = 'biff8') => {
  const ws = XLSX.utils.aoa_to_sheet(aoa)
  ws['!merges'] = merges.map((m) => XLSX.utils.decode_range(m))
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  return XLSX.write(wb, { type: 'array', bookType })
}

sampleTest('real sample file', () => {
  const r = parseOrderWorkbook(fs.readFileSync(SAMPLE))
  assert.equal(r.header.supplierColumn, 'F')
  assert.equal(r.header.qtyColumn, 'H')
  assert.equal(r.shop.name, 'КОМЕ КНЯЖЕВО')
  assert.equal(r.category.category, 'Млечни')
  assert.equal(r.products.length, 9)
  assert.equal(r.warnings.length, 0)
  assert.equal(buildFileName(r.shop.name, r.category.category), 'КОМЕ КНЯЖЕВО_Млечни.csv')
  assert.equal(
    buildCsv(r.products, 'standard'),
    ['107106,,3', '116125,,8', '138003,,3', '107129,,5', '107111,,2', '191101,,20', '197201,,5', '191104,,5', '191114,,3']
      .join('\r\n') + '\r\n',
  )
})

const base = (extra = []) => [
  ['ЗАЯВКА N 1/01.01.2026'],
  ['От фирма: X ООД', null, null, null, null, 'Към фирма: ДОСТАВЧИК ЕООД'],
  ['Магазин за доставка: Тест Център'],
  ...extra,
  [],
  ['№', 'Код', 'Баркод', 'Име', 'МЕ', 'Код на дост.', 'Пакетаж', 'Кол.'],
  [1, '1', '', 'Продукт млечни', 'БР', '00123', 1, 2],
  [],
  [2, '2', '', 'Продукт 2', 'БР', 4567, 1, 1.5],
  ['Общо: 10.00'],
]

test('no category -> plain filename, leading zeros kept', () => {
  const r = parseOrderWorkbook(toXls(base()))
  assert.equal(r.category.category, '')
  assert.equal(buildFileName(r.shop.name, r.category.category), 'Тест Център.csv')
  assert.equal(buildCsv(r.products, 'standard'), '00123,,2\r\n4567,,1.5\r\n')
})

test('"Местни" in separate box -> Месни', () => {
  const r = parseOrderWorkbook(toXls(base([[null, null, null, null, null, null, 'МЕСТНИ']])))
  assert.equal(r.category.category, 'Месни')
})

test('full header "Код на доставчика" + merged header cells', () => {
  const aoa = [
    ['Магазин: Шоп 1'],
    [],
    ['Код на доставчика', null, 'Кол.'],
    ['A-1', null, 3],
  ]
  const r = parseOrderWorkbook(toXls(aoa, ['A3:B3']))
  assert.equal(r.shop.name, 'Шоп 1')
  assert.equal(buildCsv(r.products, 'standard'), 'A-1,,3\r\n')
})

test('code without quantity is reported, not silently written', () => {
  const aoa = [['Магазин: Шоп'], ['Код на дост.', 'Кол.'], ['111', 1], ['222', null]]
  const r = parseOrderWorkbook(toXls(aoa))
  assert.equal(r.products.length, 1)
  assert.equal(r.warnings.length, 1)
})

test('missing columns -> error', () => {
  assert.throws(() => parseOrderWorkbook(toXls([['Код', 'Име'], ['1', 'x']])), ConversionError)
})

test('garbage file -> error', () => {
  assert.throws(() => parseOrderWorkbook(new TextEncoder().encode('not an excel file at all \u0000\u0001')), ConversionError)
})

test('xlsx format works too', () => {
  const r = parseOrderWorkbook(toXls(base(), [], 'xlsx'))
  assert.equal(r.products.length, 2)
})

test('CSV escaping only when needed', () => {
  assert.equal(buildCsv([{ code: 'A,1', qty: '2' }], 'standard'), '"A,1",,2\r\n')
})

const sampleProducts = () => parseOrderWorkbook(fs.readFileSync(SAMPLE)).products

sampleTest('Excel CSV: BOM + sep=; + semicolons, exactly 3 fields', () => {
  const products = sampleProducts()
  const bytes = buildCsvBytes(products, EXPORT_FORMATS.EXCEL)
  assert.deepEqual([...bytes.slice(0, 3)], [0xef, 0xbb, 0xbf])
  const text = new TextDecoder().decode(bytes.slice(3))
  const lines = text.split('\r\n')
  assert.equal(lines[0], 'sep=;')
  assert.equal(lines.at(-1), '') // trailing newline only
  const data = lines.slice(1, -1)
  assert.equal(data.length, products.length)
  for (const l of data) {
    const f = l.split(';')
    assert.equal(f.length, 3, l)
    assert.equal(f[1], '')
    assert.ok(f[0] && f[2])
    assert.ok(!l.includes(','), 'no commas in Excel mode')
  }
  assert.equal(data[0], '107106;;3')
  if (process.env.OUT_DIR) fs.writeFileSync(`${process.env.OUT_DIR}/excel.csv`, bytes)
})

sampleTest('Standard CSV: no BOM, no sep line, commas', () => {
  const products = sampleProducts()
  const bytes = buildCsvBytes(products, 'standard')
  assert.notEqual(bytes[0], 0xef)
  const text = new TextDecoder().decode(bytes)
  assert.ok(!text.includes('sep='))
  assert.equal(text.split('\r\n')[0], '107106,,3')
  for (const l of text.trimEnd().split('\r\n')) assert.equal(l.split(',').length, 3)
})

sampleTest('Preview rows never contain the sep= directive', () => {
  const rows = buildCsvRows(sampleProducts(), 'excel')
  assert.equal(rows[0], '107106;;3')
  assert.ok(rows.every((r) => !r.startsWith('sep=')))
})

test('Excel CSV quotes values containing ";"', () => {
  assert.equal(buildCsvRows([{ code: 'A;1', qty: '2' }], 'excel')[0], '"A;1";;2')
  assert.equal(buildCsvRows([{ code: 'A,1', qty: '2' }], 'excel')[0], 'A,1;;2')
})

sampleTest('XLSX export: text codes with leading zeros, empty B, numeric C, no header', () => {
  const products = [{ code: '00123', qty: '2' }, { code: '4567', qty: '1.5' }, ...sampleProducts()]
  const bytes = buildXlsxBytes(products)
  if (process.env.OUT_DIR) fs.writeFileSync(`${process.env.OUT_DIR}/export.xlsx`, bytes)
  const wb = XLSX.read(bytes, { type: 'array', cellNF: true })
  const ws = wb.Sheets[wb.SheetNames[0]]
  assert.equal(ws['!ref'], `A1:C${products.length}`)
  assert.equal(ws.A1.t, 's')
  assert.equal(ws.A1.v, '00123')
  assert.equal(ws.A1.z, '@')
  assert.equal(ws.C1.t, 'n')
  assert.equal(ws.C1.v, 2)
  assert.equal(ws.C2.v, 1.5)
  for (let r = 1; r <= products.length; r++) assert.equal(ws[`B${r}`], undefined)
  assert.equal(ws.A3.v, '107106')
})

test('file name extensions', () => {
  assert.equal(buildFileName('КОМЕ КНЯЖЕВО', 'Млечни', 'xlsx'), 'КОМЕ КНЯЖЕВО_Млечни.xlsx')
  assert.equal(buildFileName('Шоп', ''), 'Шоп.csv')
})

// Same layout as the real orders (№, Код, Баркод, Име, МЕ, Код на дост., Пакетаж, Кол.)
const orderWithGaps = () => toXls([
  ['ЗАЯВКА N 1/01.01.2026'],
  ['Магазин за доставка: Тест'],
  [],
  ['№', 'Код', 'Баркод', 'Име', 'МЕ', 'Код на дост.', 'Пакетаж', 'Кол.'],
  [1, '207573', '3800106204053', 'Сирене 500гр', 'БР', '3800106204053', 18, 5],
  [2, '212234', '3800106201000', 'Продукт без код', 'БР', '', 24, 10],
  [],
  [3, '212241', '3800106201069', 'Кашкавал', 'БР', '3800106201069', 0, 3],
  [4, '212250', '', 'Нулево количество', 'БР', '0042', 1, 0],
  [5, '212260', '', 'Без код и нула', 'БР', null, 1, 0],
  [null, null, null, null, null, null, null, 18], // unlabeled total of quantities
  ['Общо: 135.46'],
])

test('rows with quantity but no supplier code are kept, in order, with empty A', () => {
  const r = parseOrderWorkbook(orderWithGaps())
  assert.deepEqual(r.products.map((p) => [p.code, p.qty, p.row]), [
    ['3800106204053', '5', 5],
    ['', '10', 6],
    ['3800106201069', '3', 8],
    ['0042', '0', 9],
    ['', '0', 10],
  ])
  assert.equal(r.warnings.length, 0)
  assert.deepEqual(buildCsvRows(r.products, 'excel'), [
    '3800106204053;;5', ';;10', '3800106201069;;3', '0042;;0', ';;0',
  ])
  assert.equal(
    buildCsv(r.products, 'standard'),
    '3800106204053,,5\r\n,,10\r\n3800106201069,,3\r\n0042,,0\r\n,,0\r\n',
  )
  // Total rows are still skipped
  assert.ok(r.skipped.some((x) => x.row === 11 && /обобщен/i.test(x.reason)))
  assert.ok(r.skipped.some((x) => x.row === 12))
})

test('XLSX: missing supplier code leaves A truly empty, quantity stays on its row', () => {
  const r = parseOrderWorkbook(orderWithGaps())
  const ws = XLSX.read(buildXlsxBytes(r.products), { type: 'array' }).Sheets['Поръчка']
  assert.equal(ws.A1.v, '3800106204053')
  assert.equal(ws.A2, undefined)
  assert.equal(ws.C2.v, 10)
  assert.equal(ws.A3.v, '3800106201069')
  assert.equal(ws.C3.v, 3)
  assert.equal(ws.A4.v, '0042')
  assert.equal(ws.C5.v, 0)
})

if (failed) { console.error(`\n${failed} test(s) failed`); process.exit(1) }
console.log('\nAll tests passed')
