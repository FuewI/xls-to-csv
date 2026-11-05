import { useMemo, useState } from 'react'
import DropZone from './components/DropZone.jsx'
import {
  CATEGORIES,
  ConversionError,
  EXPORT_FORMATS,
  buildCsvBytes,
  buildCsvRows,
  buildFileName,
  buildXlsxBytes,
  parseOrderWorkbook,
} from './lib/converter.js'

const CATEGORY_OPTIONS = [
  { value: CATEGORIES.NONE, label: 'Няма категория' },
  { value: CATEGORIES.MILK, label: 'Млечни' },
  { value: CATEGORIES.MEAT, label: 'Месни' },
]

const PREVIEW_LIMIT = 200

const FORMAT_OPTIONS = [
  {
    format: EXPORT_FORMATS.EXCEL,
    title: 'Excel CSV',
    hint: 'точка и запетая (;), UTF-8 с BOM, sep=; — отваря се директно в Excel',
  },
  {
    format: EXPORT_FORMATS.STANDARD,
    title: 'Стандартен CSV',
    hint: 'запетая (,), UTF-8 — за други програми и системи',
  },
]

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

function saveFile(bytes, fileName, mime) {
  const blob = new Blob([bytes], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

export default function App() {
  const [fileName, setFileName] = useState('')
  const [result, setResult] = useState(null)
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [shopName, setShopName] = useState('')
  const [category, setCategory] = useState(CATEGORIES.NONE)
  const [showSkipped, setShowSkipped] = useState(false)
  const [exportFormat, setExportFormat] = useState(EXPORT_FORMATS.EXCEL)

  const handleFile = async (file) => {
    setError('')
    setResult(null)
    setFileName(file.name)
    setShowSkipped(false)

    if (!/\.(xls|xlsx)$/i.test(file.name)) {
      setError('Невалиден файл. Моля, качете Excel файл с разширение .xls или .xlsx.')
      return
    }

    setBusy(true)
    try {
      // The file is only read into memory — the original is never modified.
      const buffer = await file.arrayBuffer()
      const parsed = parseOrderWorkbook(new Uint8Array(buffer))
      setResult(parsed)
      setShopName(parsed.shop.name)
      setCategory(parsed.category.category)
    } catch (e) {
      setError(
        e instanceof ConversionError
          ? e.message
          : 'Възникна неочаквана грешка при обработката на файла. Проверете дали файлът е валиден Excel документ.',
      )
      console.error(e)
    } finally {
      setBusy(false)
    }
  }

  const csvFileName = useMemo(() => buildFileName(shopName, category), [shopName, category])
  const missingCodeCount = useMemo(
    () => (result ? result.products.filter((p) => !p.code).length : 0),
    [result],
  )
  const xlsxFileName = useMemo(() => buildFileName(shopName, category, 'xlsx'), [shopName, category])
  // Preview shows data rows only — the "sep=;" directive is not a data row.
  const csvPreviewRows = useMemo(
    () => (result ? buildCsvRows(result.products.slice(0, 5), exportFormat) : []),
    [result, exportFormat],
  )

  const downloadCsv = () => {
    if (!result || !csvFileName) return
    saveFile(buildCsvBytes(result.products, exportFormat), csvFileName, 'text/csv;charset=utf-8')
  }

  const downloadXlsx = () => {
    if (!result || !xlsxFileName) return
    saveFile(buildXlsxBytes(result.products), xlsxFileName, XLSX_MIME)
  }

  const reset = () => {
    setResult(null)
    setError('')
    setFileName('')
  }

  return (
    <div className="min-h-screen">
      <header className="border-b border-slate-200 bg-white/80 backdrop-blur">
        <div className="mx-auto flex max-w-4xl items-center gap-3 px-6 py-4">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-blue-600 text-white shadow-sm">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="h-5 w-5">
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 5h16v14H4zM4 10h16M10 5v14" />
            </svg>
          </div>
          <div>
            <h1 className="text-lg font-semibold leading-tight text-slate-900">Конвертор на поръчки</h1>
            <p className="text-xs text-slate-500">Excel → CSV · обработката става изцяло във вашия браузър</p>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-4xl space-y-6 px-6 py-8">
        {!result && (
          <section className="text-center">
            <h2 className="text-2xl font-semibold text-slate-900">Качете Excel файл с поръчка</h2>
            <p className="mt-1 text-slate-500">
              Данните се извличат автоматично: код на доставчика и количество.
            </p>
          </section>
        )}

        <DropZone onFile={handleFile} compact={!!result} busy={busy} />

        {error && (
          <div className="flex items-start gap-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-red-800">
            <span className="mt-0.5 text-lg leading-none">⚠</span>
            <div>
              <p className="font-medium">Грешка{fileName ? ` във файла „${fileName}“` : ''}</p>
              <p className="text-sm">{error}</p>
            </div>
          </div>
        )}

        {result && (
          <>
            {/* File information */}
            <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
              <div className="mb-5 flex items-center justify-between">
                <h2 className="text-base font-semibold text-slate-900">Информация за файла</h2>
                <button onClick={reset} className="text-sm text-slate-500 hover:text-slate-800">
                  Изчисти
                </button>
              </div>

              <dl className="grid gap-x-8 gap-y-4 sm:grid-cols-2">
                <Info label="Оригинален файл" value={fileName} />
                <Info
                  label="Брой продукти"
                  value={
                    <>
                      <span className="text-xl font-semibold text-blue-700">{result.products.length}</span>
                      {missingCodeCount > 0 && (
                        <span className="ml-2 text-sm text-slate-500">
                          (от тях {missingCodeCount} без код на доставчика — колона A остава празна)
                        </span>
                      )}
                    </>
                  }
                />
                <Info
                  label="Колони в Excel"
                  value={`„${result.header.supplierHeaderText}“ → ${result.header.supplierColumn}, „${result.header.qtyHeaderText}“ → ${result.header.qtyColumn} (заглавен ред ${result.header.headerRow})`}
                />
                <Info
                  label="Открита категория"
                  value={
                    result.category.category ? (
                      <>
                        {result.category.category}{' '}
                        <span className="text-xs text-slate-400">({result.category.source})</span>
                      </>
                    ) : (
                      <span className="text-slate-500">Няма</span>
                    )
                  }
                />
              </dl>

              <div className="mt-6 grid gap-4 sm:grid-cols-2">
                <label className="block">
                  <span className="mb-1 block text-sm font-medium text-slate-700">Име на магазина</span>
                  <input
                    value={shopName}
                    onChange={(e) => setShopName(e.target.value)}
                    placeholder="Въведете име на магазина"
                    className={[
                      'w-full rounded-lg border bg-white px-3 py-2 text-slate-900 outline-none transition focus:ring-4',
                      shopName.trim()
                        ? 'border-slate-300 focus:border-blue-500 focus:ring-blue-100'
                        : 'border-amber-400 focus:border-amber-500 focus:ring-amber-100',
                    ].join(' ')}
                  />
                  <span className="mt-1 block text-xs text-slate-500">
                    {result.shop.name
                      ? `Открито в ${result.shop.source}`
                      : 'Името на магазина не беше открито автоматично — моля, въведете го.'}
                  </span>
                  {result.shop.name && !result.shop.confident && (
                    <span className="mt-1 block text-xs text-amber-700">
                      Името е предположение — моля, проверете го.
                    </span>
                  )}
                </label>

                <label className="block">
                  <span className="mb-1 block text-sm font-medium text-slate-700">Категория</span>
                  <select
                    value={category}
                    onChange={(e) => setCategory(e.target.value)}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-slate-900 outline-none transition focus:border-blue-500 focus:ring-4 focus:ring-blue-100"
                  >
                    {CATEGORY_OPTIONS.map((o) => (
                      <option key={o.label} value={o.value}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                  {result.category.ambiguous && (
                    <span className="mt-1 block text-xs text-amber-700">
                      Открити са и двете категории ({result.category.source}) — изберете ръчно.
                    </span>
                  )}
                </label>
              </div>

              <fieldset className="mt-6">
                <legend className="mb-1 text-sm font-medium text-slate-700">Формат на CSV</legend>
                <div className="grid gap-2 sm:grid-cols-2">
                  {FORMAT_OPTIONS.map(({ format, title, hint }) => {
                    const active = exportFormat.id === format.id
                    return (
                      <label
                        key={format.id}
                        className={[
                          'flex cursor-pointer items-start gap-3 rounded-lg border px-3 py-2.5 transition',
                          active
                            ? 'border-blue-500 bg-blue-50/60 ring-2 ring-blue-100'
                            : 'border-slate-300 bg-white hover:border-blue-300',
                        ].join(' ')}
                      >
                        <input
                          type="radio"
                          name="export-format"
                          checked={active}
                          onChange={() => setExportFormat(format)}
                          className="mt-1 accent-blue-600"
                        />
                        <span>
                          <span className="block text-sm font-medium text-slate-900">
                            {title}
                            {format.id === EXPORT_FORMATS.EXCEL.id && (
                              <span className="ml-2 rounded bg-blue-100 px-1.5 py-0.5 text-[10px] font-semibold uppercase text-blue-700">
                                препоръчан
                              </span>
                            )}
                          </span>
                          <span className="block text-xs text-slate-500">{hint}</span>
                        </span>
                      </label>
                    )
                  })}
                </div>
                <p className="mt-2 text-xs text-slate-500">
                  Файлът съдържа кодовете точно както са в поръчката. При отваряне на CSV Excel може да покаже
                  дългите кодове като 3,8E+12 или да скрие водещите нули — за точен вид в Excel използвайте
                  „Изтегли Excel (.xlsx)“.
                </p>
              </fieldset>

              <div className="mt-6 rounded-xl bg-blue-50 px-4 py-3">
                <p className="text-xs font-medium uppercase tracking-wide text-blue-700">Име на CSV файла</p>
                <p className="mt-0.5 break-all font-mono text-lg text-blue-950">
                  {csvFileName || <span className="text-slate-400">— въведете име на магазина —</span>}
                </p>
              </div>
            </section>

            {/* Warnings */}
            {result.warnings.length > 0 && (
              <section className="rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-amber-900">
                <p className="font-medium">
                  Внимание: {result.warnings.length} ред(а) не бяха включени, защото данните са непълни
                </p>
                <ul className="mt-1 list-disc space-y-0.5 pl-5 text-sm">
                  {result.warnings.map((w) => (
                    <li key={w.row}>{w.message}</li>
                  ))}
                </ul>
              </section>
            )}

            {/* Preview */}
            <section className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
              <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4">
                <h2 className="text-base font-semibold text-slate-900">Преглед на данните</h2>
                <span className="text-sm text-slate-500">{result.products.length} реда</span>
              </div>
              <div className="max-h-[28rem] overflow-auto">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
                    <tr>
                      <th className="w-14 px-4 py-2 font-medium">#</th>
                      <th className="px-4 py-2 font-medium">
                        A <span className="normal-case text-slate-400">· код на доставчика</span>
                      </th>
                      <th className="px-4 py-2 font-medium">
                        B <span className="normal-case text-slate-400">· празна</span>
                      </th>
                      <th className="px-4 py-2 text-right font-medium">
                        C <span className="normal-case text-slate-400">· кол.</span>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.products.slice(0, PREVIEW_LIMIT).map((p, i) => (
                      <tr key={i} className="hover:bg-blue-50/40">
                        <td className="px-4 py-2 text-slate-400">{i + 1}</td>
                        <td className="px-4 py-2 font-mono text-slate-900">
                          {p.code || (
                            <span className="rounded bg-slate-100 px-1.5 py-0.5 font-sans text-xs text-slate-500">
                              без код
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-2 text-slate-300">—</td>
                        <td className="px-4 py-2 text-right font-mono text-slate-900">{p.qty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {result.products.length > PREVIEW_LIMIT && (
                <p className="border-t border-slate-200 px-6 py-2 text-xs text-slate-500">
                  Показани са първите {PREVIEW_LIMIT} реда. CSV файлът съдържа всички {result.products.length}.
                </p>
              )}
              <div className="border-t border-slate-200 bg-slate-50 px-6 py-3">
                <p className="mb-1 text-xs font-medium text-slate-500">Съдържание на CSV (първи редове)</p>
                <pre className="overflow-x-auto font-mono text-xs text-slate-700">
                  {csvPreviewRows.join('\n')}
                  {result.products.length > 5 ? '\n…' : ''}
                </pre>
              </div>
            </section>

            {/* Download */}
            <div className="space-y-3">
              <button
                onClick={downloadCsv}
                disabled={!csvFileName}
                className="flex w-full items-center justify-center gap-3 rounded-2xl bg-blue-600 px-6 py-5 text-lg font-semibold text-white shadow-lg shadow-blue-600/20 transition hover:bg-blue-700 active:scale-[0.99] disabled:cursor-not-allowed disabled:bg-slate-300 disabled:shadow-none"
              >
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="h-6 w-6">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v12m0 0-4-4m4 4 4-4M4 20h16" />
                </svg>
                Изтегли CSV
              </button>
              <button
                onClick={downloadXlsx}
                disabled={!xlsxFileName}
                className="flex w-full items-center justify-center gap-2 rounded-xl border border-blue-200 bg-white px-6 py-3 text-base font-medium text-blue-700 transition hover:border-blue-400 hover:bg-blue-50 disabled:cursor-not-allowed disabled:border-slate-200 disabled:text-slate-400"
              >
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="h-5 w-5">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 5h16v14H4zM4 10h16M10 5v14" />
                </svg>
                Изтегли Excel (.xlsx)
                {xlsxFileName && <span className="font-mono text-xs text-slate-400">{xlsxFileName}</span>}
              </button>
            </div>

            {/* Skipped rows (transparency) */}
            {result.skipped.length > 0 && (
              <section className="text-sm text-slate-500">
                <button
                  onClick={() => setShowSkipped((v) => !v)}
                  className="font-medium text-slate-600 hover:text-slate-900"
                >
                  {showSkipped ? '▾' : '▸'} Пропуснати редове без продукти ({result.skipped.length})
                </button>
                {showSkipped && (
                  <ul className="mt-2 space-y-1 rounded-xl border border-slate-200 bg-white p-4">
                    {result.skipped.map((s) => (
                      <li key={s.row}>
                        <span className="font-mono text-slate-400">Ред {s.row}</span> · {s.reason}:{' '}
                        <span className="text-slate-700">{s.text}</span>
                      </li>
                    ))}
                  </ul>
                )}
              </section>
            )}
          </>
        )}
      </main>

      <footer className="pb-8 text-center text-xs text-slate-400">
        Файловете не се качват на сървър и оригиналът не се променя.
      </footer>
    </div>
  )
}

function Info({ label, value }) {
  return (
    <div>
      <dt className="text-xs font-medium uppercase tracking-wide text-slate-500">{label}</dt>
      <dd className="mt-0.5 break-words text-slate-900">{value}</dd>
    </div>
  )
}
