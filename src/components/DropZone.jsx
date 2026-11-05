import { useRef, useState } from 'react'

const ACCEPT = '.xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

export default function DropZone({ onFile, compact = false, busy = false }) {
  const inputRef = useRef(null)
  const [dragging, setDragging] = useState(false)

  const pick = (files) => {
    const file = files && files[0]
    if (file) onFile(file)
  }

  return (
    <div
      role="button"
      tabIndex={0}
      onClick={() => inputRef.current?.click()}
      onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && inputRef.current?.click()}
      onDragOver={(e) => {
        e.preventDefault()
        setDragging(true)
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault()
        setDragging(false)
        pick(e.dataTransfer.files)
      }}
      className={[
        'group relative flex cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed text-center transition',
        'focus:outline-none focus-visible:ring-4 focus-visible:ring-blue-200',
        compact ? 'px-6 py-6' : 'px-6 py-14',
        dragging
          ? 'border-blue-500 bg-blue-50'
          : 'border-slate-300 bg-white hover:border-blue-400 hover:bg-blue-50/40',
      ].join(' ')}
    >
      <input
        ref={inputRef}
        type="file"
        accept={ACCEPT}
        className="hidden"
        onChange={(e) => {
          pick(e.target.files)
          e.target.value = '' // allow re-selecting the same file
        }}
      />
      <div
        className={[
          'mb-3 flex items-center justify-center rounded-full bg-blue-100 text-blue-600 transition group-hover:scale-105',
          compact ? 'h-10 w-10' : 'h-14 w-14',
        ].join(' ')}
      >
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" className={compact ? 'h-5 w-5' : 'h-7 w-7'}>
          <path strokeLinecap="round" strokeLinejoin="round" d="M12 16V4m0 0-4 4m4-4 4 4M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" />
        </svg>
      </div>
      <p className={compact ? 'text-sm font-medium text-slate-700' : 'text-lg font-semibold text-slate-800'}>
        {busy ? 'Обработване…' : compact ? 'Качете друг файл' : 'Пуснете Excel файла тук'}
      </p>
      <p className="mt-1 text-sm text-slate-500">
        или <span className="font-medium text-blue-600 underline-offset-2 group-hover:underline">изберете файл</span> от компютъра · .xls, .xlsx
      </p>
    </div>
  )
}
