/**
 * Conversion logic: Excel order file -> "supplier_code,,quantity" CSV.
 *
 * Pure module (no React / DOM) so it can be tested from Node as well.
 * The original file is only read, never modified.
 */
import * as XLSX from 'xlsx'
import * as cptable from 'xlsx/dist/cpexcel.full.mjs'

// Legacy BIFF5 .xls files store text in a code page (e.g. cp1251 for Cyrillic).
XLSX.set_cptable(cptable)

export const CATEGORIES = {
  NONE: '',
  MILK: 'Млечни',
  MEAT: 'Месни',
}

export class ConversionError extends Error {
  constructor(message) {
    super(message)
    this.name = 'ConversionError'
  }
}

/* ------------------------------------------------------------------ */
/* Text helpers                                                        */
/* ------------------------------------------------------------------ */

const clean = (s) =>
  String(s ?? '')
    .replace(/[   ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

/** Lower-case, collapsed whitespace, trailing ":" / "." removed (except "кол."). */
const norm = (s) => clean(s).toLowerCase()
const normLabel = (s) => norm(s).replace(/[\s:：\-–—]+$/, '')

const CYR = 'а-яёѐ-ѝa-z0-9'
const wordRe = (w) => new RegExp(`(^|[^${CYR}])(${w})(?=[^${CYR}]|$)`, 'i')

/* ------------------------------------------------------------------ */
/* Header patterns                                                     */
/* ------------------------------------------------------------------ */

// Matches "Код на доставчика" and its abbreviations ("Код на дост.", "Код дост.").
// Plain "Код" (internal article code) is intentionally NOT matched.
const isSupplierHeader = (t) => /^код\s*(на\s*)?дост(авчика?|\.)?\s*\.?$/.test(normLabel(t))

// Exact "Кол." is preferred; a few common variants are accepted as fallback.
const qtyHeaderScore = (t) => {
  const n = norm(t).replace(/[:\s]+$/, '')
  if (n === 'кол.') return 3
  if (n === 'кол' || n === 'количество') return 2
  if (n === 'к-во' || n === 'к-во.' || n === 'кол-во' || n === 'кол-во.') return 1
  return 0
}

const SUMMARY_RE = new RegExp(
  `^(общо|всичко|сума|обща сума|общо сума|тотал|total|итого|subtotal|междинна сума)(${'[^' + CYR + ']'}|$)`,
  'i',
)

/* ------------------------------------------------------------------ */
/* Value extraction                                                    */
/* ------------------------------------------------------------------ */

/** Supplier code exactly as shown in Excel, keeping leading zeros. */
export function codeValue(cell) {
  if (!cell || cell.v == null || cell.v === '') return ''
  if (cell.t === 's' || cell.t === 'str') return clean(cell.v)
  if (cell.t === 'n') {
    const w = clean(cell.w)
    // Custom number formats like "00000" render leading zeros in .w
    if (/^\d+$/.test(w)) return w
    if (Number.isInteger(cell.v)) return String(cell.v) // avoid "1.23E+12" from General format
    return w || String(cell.v)
  }
  return clean(cell.w ?? cell.v)
}

const formatNumber = (n) =>
  Number.isInteger(n) ? String(n) : String(parseFloat(n.toPrecision(12)))

/** Quantity as a plain number string, or null if the cell is not a number. */
export function qtyValue(cell) {
  if (!cell || cell.v == null || cell.v === '') return null
  if (cell.t === 'n') return Number.isFinite(cell.v) ? formatNumber(cell.v) : null
  if (cell.t === 's' || cell.t === 'str') {
    const s = clean(cell.v).replace(/\s/g, '')
    if (/^-?\d+([.,]\d+)?$/.test(s)) return formatNumber(Number(s.replace(',', '.')))
  }
  return null
}

/* ------------------------------------------------------------------ */
/* Sheet model                                                         */
/* ------------------------------------------------------------------ */

function buildSheetModel(ws) {
  if (!ws || !ws['!ref']) return null
  const range = XLSX.utils.decode_range(ws['!ref'])
  const merges = ws['!merges'] || []

  // address -> merge range, for every cell covered by a merge
  const mergeOf = new Map()
  for (const m of merges) {
    for (let r = m.s.r; r <= m.e.r; r++)
      for (let c = m.s.c; c <= m.e.c; c++) mergeOf.set(`${r}:${c}`, m)
  }

  const raw = (r, c) => ws[XLSX.utils.encode_cell({ r, c })]
  /** Cell value, resolving merged areas to their top-left cell. */
  const cell = (r, c) => {
    const m = mergeOf.get(`${r}:${c}`)
    return m ? raw(m.s.r, m.s.c) : raw(r, c)
  }
  const text = (r, c) => {
    const x = cell(r, c)
    return x && x.v != null ? clean(x.w ?? x.v) : ''
  }
  const isMergeTopLeft = (r, c) => {
    const m = mergeOf.get(`${r}:${c}`)
    return !m || (m.s.r === r && m.s.c === c)
  }
  /** Non-empty cells of a row, merged areas counted once. */
  const rowCells = (r) => {
    const out = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      const m = mergeOf.get(`${r}:${c}`)
      if (m && (m.s.c !== c)) continue
      const t = text(r, c)
      if (t) out.push({ c, t, cell: cell(r, c), merge: m })
    }
    return out
  }

  return { ws, range, mergeOf, raw, cell, text, isMergeTopLeft, rowCells }
}

/* ------------------------------------------------------------------ */
/* Header detection                                                    */
/* ------------------------------------------------------------------ */

function findHeader(model) {
  const { range, text, mergeOf } = model
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      if (!model.isMergeTopLeft(r, c)) continue
      if (!isSupplierHeader(text(r, c))) continue

      // Quantity header: same row preferred, then nearby rows (multi-row headers).
      let best = null
      for (const dr of [0, 1, -1, 2, -2]) {
        const rr = r + dr
        if (rr < range.s.r || rr > range.e.r) continue
        for (let cc = range.s.c; cc <= range.e.c; cc++) {
          if (cc === c || !model.isMergeTopLeft(rr, cc)) continue
          const score = qtyHeaderScore(text(rr, cc))
          if (score && (!best || score > best.score)) best = { r: rr, c: cc, score }
        }
        if (best && best.score === 3) break
      }
      if (!best) continue

      const supMerge = mergeOf.get(`${r}:${c}`)
      const qtyMerge = mergeOf.get(`${best.r}:${best.c}`)
      const bottom = Math.max(supMerge ? supMerge.e.r : r, qtyMerge ? qtyMerge.e.r : best.r)
      return {
        supplierCol: c,
        supplierSpan: supMerge ? [supMerge.s.c, supMerge.e.c] : [c, c],
        qtyCol: best.c,
        qtySpan: qtyMerge ? [qtyMerge.s.c, qtyMerge.e.c] : [best.c, best.c],
        headerRow: Math.min(r, best.r),
        headerBottom: bottom,
        supplierHeaderText: text(r, c),
        qtyHeaderText: text(best.r, best.c),
      }
    }
  }
  return null
}

/* ------------------------------------------------------------------ */
/* Product rows                                                        */
/* ------------------------------------------------------------------ */

/** First non-empty value inside the header's column span (handles merged data cells). */
function cellInSpan(model, r, span) {
  for (let c = span[0]; c <= span[1]; c++) {
    const x = model.raw(r, c)
    if (x && x.v != null && x.v !== '') return { cell: x, c }
  }
  return { cell: null, c: span[0] }
}

/** True if a data cell belongs to a merged block wider than the header column (a note/footer line). */
function isWideMerge(model, r, c, span) {
  const m = model.mergeOf.get(`${r}:${c}`)
  if (!m) return false
  return m.s.c < span[0] || m.e.c > span[1]
}

function extractProducts(model, header) {
  const { range } = model
  const products = []
  const warnings = []
  const skipped = []
  const productRowSet = new Set()

  for (let r = header.headerBottom + 1; r <= range.e.r; r++) {
    const cells = model.rowCells(r)
    if (cells.length === 0) continue // blank row

    const excelRow = r + 1
    const rowText = cells.map((x) => x.t).join(' | ')

    // Repeated header (multi-page exports)
    if (cells.some((x) => isSupplierHeader(x.t))) {
      skipped.push({ row: excelRow, reason: 'Повторен заглавен ред', text: rowText })
      continue
    }

    // Summary / totals rows
    if (cells.some((x) => SUMMARY_RE.test(norm(x.t)))) {
      skipped.push({ row: excelRow, reason: 'Обобщен ред (общо/сума)', text: rowText })
      continue
    }

    const sup = cellInSpan(model, r, header.supplierSpan)
    const qty = cellInSpan(model, r, header.qtySpan)

    const supIsNote = sup.cell && isWideMerge(model, r, sup.c, header.supplierSpan)
    const qtyIsNote = qty.cell && isWideMerge(model, r, qty.c, header.qtySpan)
    if (supIsNote || qtyIsNote) {
      skipped.push({ row: excelRow, reason: 'Допълнителна информация', text: rowText })
      continue
    }

    const code = codeValue(sup.cell)
    const q = qtyValue(qty.cell)

    if (code && q !== null) {
      products.push({ code, qty: q, row: excelRow })
      productRowSet.add(r)
      continue
    }

    if (code && q === null) {
      const qtyText = qty.cell ? clean(qty.cell.w ?? qty.cell.v) : ''
      warnings.push({
        row: excelRow,
        message: qtyText
          ? `Ред ${excelRow}: код „${code}“ има невалидно количество „${qtyText}“ — редът е пропуснат.`
          : `Ред ${excelRow}: код „${code}“ е без количество — редът е пропуснат.`,
      })
      continue
    }

    if (!code && q !== null) {
      // Only numbers in the row -> most likely an unlabeled total line.
      const hasText = cells.some((x) => x.cell && x.cell.t === 's' && qtyValue(x.cell) === null)
      if (!hasText) {
        skipped.push({ row: excelRow, reason: 'Вероятно обобщен ред (само числа)', text: rowText })
      } else {
        warnings.push({
          row: excelRow,
          message: `Ред ${excelRow}: има количество ${q}, но липсва код на доставчика — редът е пропуснат.`,
        })
      }
      continue
    }

    skipped.push({ row: excelRow, reason: 'Допълнителна информация', text: rowText })
  }

  return { products, warnings, skipped, productRowSet }
}

/* ------------------------------------------------------------------ */
/* Shop name detection                                                 */
/* ------------------------------------------------------------------ */

// Ordered by priority. Labels that refer to the supplier are deliberately excluded.
const SHOP_LABELS = [
  ['магазин за доставка', 'обект за доставка', 'име на магазин', 'име на обект', 'търговски обект', 'магазин', 'обект'],
  ['получател', 'доставка до', 'адрес за доставка', 'място на доставка', 'място за доставка'],
  ['клиент', 'купувач', 'поръчител', 'заявител', 'контрагент'],
]
const ALL_SHOP_LABELS = SHOP_LABELS.flat().sort((a, b) => b.length - a.length)

const NOT_A_NAME_RE =
  /^(поръчка|заявка|дата|№|номер|доставчик|код|тел|телефон|адрес|еик|булстат|ддс|мол|страница|стр\.|забележка|подпис|категория|млечни|месни|местни|общо|сума)\b/i

const looksLikeValue = (t) => {
  const n = normLabel(t)
  if (!n || n.length < 2) return false
  if (!/[a-zа-яё]/i.test(n)) return false // numbers / dates only
  if (ALL_SHOP_LABELS.includes(n)) return false
  if (isSupplierHeader(n) || qtyHeaderScore(n)) return false
  return true
}

function findShopName(model, excludedRows, header) {
  const { range } = model
  const candidates = []

  for (let r = range.s.r; r <= range.e.r; r++) {
    if (excludedRows.has(r)) continue
    const cells = model.rowCells(r)
    for (let i = 0; i < cells.length; i++) {
      const { c, t } = cells[i]
      const n = norm(t)
      for (let p = 0; p < SHOP_LABELS.length; p++) {
        for (const label of SHOP_LABELS[p]) {
          // "Магазин: Х" inside the same cell
          const inline = n.match(
            new RegExp(`(^|[^${CYR}])${label}\\s*[:：]\\s*(.+)$`, 'i'),
          )
          if (inline && /доставчик/.test(n.slice(0, inline.index + inline[0].length - inline[2].length))) continue
          if (inline) {
            const value = clean(t.slice(t.length - inline[2].length))
            if (looksLikeValue(value)) {
              candidates.push({ value, priority: p, r, c, label })
              continue
            }
          }
          // Label alone in a cell -> value in next cell to the right, or below.
          if (normLabel(t) === label) {
            const right = cells.slice(i + 1).find((x) => looksLikeValue(x.t))
            if (right) {
              candidates.push({ value: right.t, priority: p, r, c: right.c, label })
              continue
            }
            const below = model.text(r + 1, c)
            if (below && looksLikeValue(below) && !excludedRows.has(r + 1)) {
              candidates.push({ value: below, priority: p, r: r + 1, c, label })
            }
          }
        }
      }
    }
  }

  if (candidates.length) {
    candidates.sort((a, b) => a.priority - b.priority || a.r - b.r || a.c - b.c)
    const best = candidates[0]
    return {
      name: best.value,
      source: `клетка ${XLSX.utils.encode_cell({ r: best.r, c: best.c })} (етикет „${best.label}“)`,
      confident: true,
    }
  }

  // Fallback: first descriptive text above the table header (low confidence).
  if (header) {
    for (let r = range.s.r; r < header.headerRow; r++) {
      for (const { c, t } of model.rowCells(r)) {
        if (NOT_A_NAME_RE.test(norm(t))) continue
        if (!looksLikeValue(t) || t.length > 80) continue
        return {
          name: t,
          source: `клетка ${XLSX.utils.encode_cell({ r, c })} (предположение)`,
          confident: false,
        }
      }
    }
  }

  return { name: '', source: '', confident: false }
}

/* ------------------------------------------------------------------ */
/* Category detection                                                  */
/* ------------------------------------------------------------------ */

const MILK_RE = wordRe('млечни')
const MEAT_RE = wordRe('месни')
const LOCAL_RE = wordRe('местни')
const stripCategoryPrefix = (n) => n.replace(/^(категория|група|отдел|вид)\s*[:：-]?\s*/, '')

function classifyCategory(t) {
  const n = normLabel(t)
  const bare = stripCategoryPrefix(n)
  const exact = bare === 'млечни' ? CATEGORIES.MILK : bare === 'месни' || bare === 'местни' ? CATEGORIES.MEAT : null
  if (exact) return { category: exact, strength: 2 }
  // Contained in longer text, e.g. supplier name "…-МАДЖАРОВ-2-МЛЕЧНИ".
  if (MILK_RE.test(n)) return { category: CATEGORIES.MILK, strength: 1 }
  if (MEAT_RE.test(n)) return { category: CATEGORIES.MEAT, strength: 1 }
  if (LOCAL_RE.test(n)) return { category: CATEGORIES.MEAT, strength: 1 }
  return null
}

function findCategory(model, productRowSet, sheetName) {
  const { range } = model
  const hits = []

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (const { c, t } of model.rowCells(r)) {
      const hit = classifyCategory(t)
      if (!hit) continue
      // Inside product rows only an exact standalone label counts (product names may contain the word).
      if (productRowSet.has(r) && hit.strength < 2) continue
      hits.push({ ...hit, where: `клетка ${XLSX.utils.encode_cell({ r, c })}`, text: t })
    }
  }
  if (!hits.length && sheetName) {
    const hit = classifyCategory(sheetName)
    if (hit) hits.push({ ...hit, where: `име на листа „${sheetName}“`, text: sheetName })
  }

  if (!hits.length) return { category: CATEGORIES.NONE, source: '', ambiguous: false }

  const maxStrength = Math.max(...hits.map((h) => h.strength))
  const top = hits.filter((h) => h.strength === maxStrength)
  const kinds = new Set(top.map((h) => h.category))
  if (kinds.size > 1) {
    return {
      category: CATEGORIES.NONE,
      source: top.map((h) => `${h.where}: „${h.text}“`).join('; '),
      ambiguous: true,
    }
  }
  return { category: top[0].category, source: top[0].where, ambiguous: false }
}

/* ------------------------------------------------------------------ */
/* Public API                                                          */
/* ------------------------------------------------------------------ */

/**
 * Parse an Excel file (ArrayBuffer / Uint8Array) and extract order data.
 * Throws ConversionError with a Bulgarian, user-facing message on failure.
 */
export function parseOrderWorkbook(data) {
  let wb
  try {
    wb = XLSX.read(data, { type: 'array', cellText: true, cellDates: false, dense: false })
  } catch {
    throw new ConversionError('Файлът не може да бъде прочетен. Уверете се, че е валиден Excel файл (.xls или .xlsx).')
  }
  if (!wb || !wb.SheetNames || wb.SheetNames.length === 0) {
    throw new ConversionError('Файлът не съдържа работни листове.')
  }

  let model = null
  let header = null
  let sheetName = null
  for (const name of wb.SheetNames) {
    const m = buildSheetModel(wb.Sheets[name])
    if (!m) continue
    const h = findHeader(m)
    if (h) {
      model = m
      header = h
      sheetName = name
      break
    }
  }

  if (!header) {
    throw new ConversionError(
      'Не са открити задължителните колони „Код на доставчика“ и „Кол.“. Проверете дали файлът е поръчка в очаквания формат.',
    )
  }

  const { products, warnings, skipped, productRowSet } = extractProducts(model, header)

  if (products.length === 0) {
    throw new ConversionError(
      'Колоните „Код на доставчика“ и „Кол.“ са открити, но под тях няма редове с продукти.',
    )
  }

  // Rows that belong to the table header or products are not searched for the shop name.
  const excluded = new Set(productRowSet)
  for (let r = header.headerRow; r <= header.headerBottom; r++) excluded.add(r)

  const shop = findShopName(model, excluded, header)
  const category = findCategory(model, productRowSet, sheetName)

  const XL = (c) => XLSX.utils.encode_col(c)
  return {
    sheetName,
    sheetCount: wb.SheetNames.length,
    header: {
      supplierColumn: XL(header.supplierCol),
      qtyColumn: XL(header.qtyCol),
      headerRow: header.headerRow + 1,
      supplierHeaderText: header.supplierHeaderText,
      qtyHeaderText: header.qtyHeaderText,
    },
    products,
    warnings,
    skipped,
    shop,
    category,
  }
}

/* ------------------------------------------------------------------ */
/* CSV + filename                                                      */
/* ------------------------------------------------------------------ */

export const EXPORT_FORMATS = {
  /** Semicolon-delimited, UTF-8 with BOM, "sep=;" directive — opens in columns when double-clicked in Excel. */
  EXCEL: { id: 'excel', label: 'Excel CSV', delimiter: ';', bom: true, sepLine: true },
  /** Plain RFC 4180-style CSV: comma-delimited, UTF-8, no BOM, no directive. */
  STANDARD: { id: 'standard', label: 'Стандартен CSV', delimiter: ',', bom: false, sepLine: false },
}

const resolveFormat = (format) =>
  typeof format === 'string'
    ? Object.values(EXPORT_FORMATS).find((f) => f.id === format) || EXPORT_FORMATS.EXCEL
    : format || EXPORT_FORMATS.EXCEL

const csvField = (v, delimiter) => {
  const s = String(v ?? '')
  const needsQuotes = s.includes(delimiter) || /["\r\n]/.test(s) || /^\s|\s$/.test(s)
  return needsQuotes ? `"${s.replace(/"/g, '""')}"` : s
}

/** Data rows only (no header, no "sep=" directive): supplier_code<d><d>quantity. */
export function buildCsvRows(products, format) {
  const { delimiter } = resolveFormat(format)
  return products.map((p) => [csvField(p.code, delimiter), '', csvField(p.qty, delimiter)].join(delimiter))
}

/** Full CSV text (without BOM). Excel mode starts with the "sep=;" line. */
export function buildCsv(products, format) {
  const f = resolveFormat(format)
  const lines = buildCsvRows(products, f)
  if (f.sepLine) lines.unshift(`sep=${f.delimiter}`)
  return lines.join('\r\n') + '\r\n'
}

/** Encoded file bytes, with UTF-8 BOM when the format requires it. */
export function buildCsvBytes(products, format) {
  const f = resolveFormat(format)
  const body = new TextEncoder().encode(buildCsv(products, f))
  if (!f.bom) return body
  const out = new Uint8Array(body.length + 3)
  out.set([0xef, 0xbb, 0xbf], 0)
  out.set(body, 3)
  return out
}

/**
 * Real .xlsx workbook: A = supplier code (text cell, "@" format), B = empty, C = quantity (number).
 * No header row. Returns a Uint8Array.
 */
export function buildXlsxBytes(products) {
  const ws = {}
  products.forEach((p, r) => {
    ws[XLSX.utils.encode_cell({ r, c: 0 })] = { t: 's', v: p.code, z: '@' }
    const n = Number(p.qty)
    ws[XLSX.utils.encode_cell({ r, c: 2 })] = Number.isFinite(n) ? { t: 'n', v: n } : { t: 's', v: p.qty }
  })
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(products.length - 1, 0), c: 2 } })
  ws['!cols'] = [{ wch: 18 }, { wch: 4 }, { wch: 10 }]
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Поръчка')
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true }))
}

/** Remove characters that are not allowed in file names. */
export function sanitizeFileName(s) {
  return clean(s)
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^[.\s]+|[.\s]+$/g, '')
    .trim()
}

export function buildFileName(shopName, category, extension = 'csv') {
  const shop = sanitizeFileName(shopName)
  if (!shop) return ''
  return category ? `${shop}_${category}.${extension}` : `${shop}.${extension}`
}
