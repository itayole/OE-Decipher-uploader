import { Fragment, useMemo, useState } from 'react'
import './App.css'
import { APP_VERSION, RELEASE_NOTES } from './releaseNotes'

const TYPE_LABELS = {
  regular: 'רגילה',
  ab: 'AB — מודעות ספונטנית (TOM/אחרים)',
  closed_others: 'סגורה + אחר',
}

const TYPE_BADGE = {
  regular: 'רגילה',
  ab: 'AB',
  closed_others: 'סגורה + אחר',
  log: 'לוג ניקוי',
}

function parseAbSuffix(name) {
  const m = /^(.*?)([aAbB])$/.exec(name)
  if (!m) return null
  return { base: m[1], letter: m[2].toLowerCase() }
}

const STEPS = [
  { key: 'upload', label: 'העלאה' },
  { key: 'mapping', label: 'מיפוי שאלות' },
  { key: 'results', label: 'תוצאות' },
]

function StepIndicator({ step, onStepClick }) {
  const activeIndex = STEPS.findIndex((s) => s.key === step)
  return (
    <div className="steps">
      {STEPS.map((s, i) => {
        const clickable = i < activeIndex
        return (
          <Fragment key={s.key}>
            <div
              className={`step ${i === activeIndex ? 'active' : ''} ${i < activeIndex ? 'done' : ''} ${clickable ? 'clickable' : ''}`}
              onClick={clickable ? () => onStepClick(s.key) : undefined}
              role={clickable ? 'button' : undefined}
              tabIndex={clickable ? 0 : undefined}
            >
              <span className="step-dot">{i < activeIndex ? '✓' : i + 1}</span>
              <span>{s.label}</span>
            </div>
            {i < STEPS.length - 1 && (
              <div className={`step-line ${i < activeIndex ? 'done' : ''}`} />
            )}
          </Fragment>
        )
      })}
    </div>
  )
}

function DropzoneInput({ label, hint, file, disabled, onChange }) {
  return (
    <div className="field">
      <label>
        {label} <span className="hint">{hint}</span>
      </label>
      <label className={`dropzone ${disabled ? 'disabled' : ''}`}>
        <span className="dropzone-icon">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M12 16V4M12 4l-4 4M12 4l4 4" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </span>
        <span className="dropzone-text">
          {file ? <strong>{file.name}</strong> : disabled ? 'לא זמין בשלב זה' : 'לחצו לבחירת קובץ xlsx'}
        </span>
        <input type="file" accept=".xlsx,.xlsm" disabled={disabled} onChange={onChange} />
      </label>
    </div>
  )
}

function App() {
  const [step, setStep] = useState('upload') // 'upload' | 'mapping' | 'results'
  const [otcFile, setOtcFile] = useState(null)
  const [rawFile, setRawFile] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  const [jobId, setJobId] = useState(null)
  const [blocks, setBlocks] = useState([])
  const [typeByName, setTypeByName] = useState({})
  const [cleanCodeByBase, setCleanCodeByBase] = useState({})

  const [result, setResult] = useState(null)

  const blocksByName = useMemo(() => {
    const map = {}
    blocks.forEach((b) => (map[b.name] = b))
    return map
  }, [blocks])

  const abPairs = useMemo(() => {
    const pairs = {}
    blocks.forEach((b) => {
      if (typeByName[b.name] !== 'ab') return
      const parsed = parseAbSuffix(b.name)
      if (!parsed) return
      pairs[parsed.base] = pairs[parsed.base] || {}
      pairs[parsed.base][parsed.letter] = b.name
    })
    return pairs
  }, [blocks, typeByName])

  const incompletePairs = useMemo(() => {
    return Object.entries(abPairs)
      .filter(([, p]) => p.a && p.b)
      .filter(([base]) => !cleanCodeByBase[base])
      .map(([base]) => base)
  }, [abPairs, cleanCodeByBase])

  const handleUpload = async (e) => {
    e.preventDefault()
    if (!otcFile) return
    setLoading(true)
    setError(null)

    const formData = new FormData()
    formData.append('otc_file', otcFile)

    try {
      const res = await fetch('/api/upload', { method: 'POST', body: formData })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || 'שגיאה בעיבוד הקובץ')
      setJobId(data.job_id)
      setBlocks(data.blocks)
      const initialTypes = {}
      data.blocks.forEach((b) => (initialTypes[b.name] = b.suggested_type))
      setTypeByName(initialTypes)
      setCleanCodeByBase({})
      setStep('mapping')
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  const handleGenerate = async () => {
    setLoading(true)
    setError(null)

    const mapping = blocks.map((b) => {
      const type = typeByName[b.name] || 'regular'
      const entry = { name: b.name, type }
      if (type === 'ab') {
        const parsed = parseAbSuffix(b.name)
        if (parsed && cleanCodeByBase[parsed.base]) {
          entry.cleaned_code = cleanCodeByBase[parsed.base]
        }
      }
      return entry
    })

    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ job_id: jobId, mapping }),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || 'שגיאה בהפקת הקבצים')
      setResult(data)
      setStep('results')
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  const handleDownload = () => {
    if (jobId) window.location.href = `/api/download/${jobId}`
  }

  const handleReset = () => {
    setStep('upload')
    setOtcFile(null)
    setJobId(null)
    setBlocks([])
    setTypeByName({})
    setCleanCodeByBase({})
    setResult(null)
    setError(null)
  }

  const handleStepClick = (key) => {
    if (key === 'upload') {
      handleReset()
    } else {
      setStep(key)
    }
  }

  return (
    <>
      <header className="app-header">
        <div className="app-header-row">
          <span className="brand-mark">OE</span>
          <h1>העלאת שאלות פתוחות מקודדות ל-Decipher</h1>
        </div>
        <p className="subtitle">
          העלו את קובץ ה-OTC המקודד, מפו את סוגי השאלות, וקבלו קבצי .dat מוכנים להעלאה ל-Decipher.
        </p>
      </header>

      <StepIndicator step={step} onStepClick={handleStepClick} />

      {step === 'upload' && (
        <form className="upload-card" onSubmit={handleUpload}>
          <DropzoneInput
            label="קובץ OTC מקודד (xlsx)"
            hint="חובה"
            file={otcFile}
            onChange={(e) => setOtcFile(e.target.files?.[0] ?? null)}
          />

          <DropzoneInput
            label="קובץ נתונים גולמי מ-Decipher (xlsx)"
            hint="אופציונלי — עדיין לא בשימוש בשלב זה"
            file={rawFile}
            disabled
            onChange={(e) => setRawFile(e.target.files?.[0] ?? null)}
          />

          <button type="submit" className="primary" disabled={!otcFile || loading}>
            {loading ? 'מעבד...' : 'המשך למיפוי שאלות'}
          </button>
        </form>
      )}

      {error && <div className="error-box">{error}</div>}

      {step === 'mapping' && (
        <section>
          <div className="results-header">
            <h2>מיפוי {blocks.length} שאלות פתוחות</h2>
            <button
              type="button"
              className="primary"
              disabled={loading || incompletePairs.length > 0}
              onClick={handleGenerate}
            >
              {loading ? 'מפיק קבצים...' : 'הפקת קבצים'}
            </button>
          </div>
          <p className="subtitle">
            בדקו את סוג השאלה שזוהה אוטומטית ושנו במידת הצורך. עבור שאלות AB, בחרו את
            "התשובה לניקוי" (למשל "לא יודע") מתוך רשימת הקטגוריות.
          </p>

          <div className="mapping-table">
            {blocks.map((block) => {
              const type = typeByName[block.name] || 'regular'
              const parsed = parseAbSuffix(block.name)
              const pair = parsed ? abPairs[parsed.base] : null
              const isPaired = type === 'ab' && pair && pair.a && pair.b
              const isRoleA = parsed?.letter === 'a'
              const partnerName = parsed && pair ? (isRoleA ? pair.b : pair.a) : null

              const unionCategories = (() => {
                if (!isPaired || !isRoleA) return []
                const a = blocksByName[pair.a]?.categories || []
                const b = blocksByName[pair.b]?.categories || []
                const seen = new Map()
                ;[...a, ...b].forEach((c) => {
                  if (!seen.has(c.code)) seen.set(c.code, c)
                })
                return [...seen.values()]
              })()

              return (
                <div className="mapping-row" key={block.name}>
                  <div className="mapping-row-main">
                    <span className="mapping-name">{block.name}</span>
                    <span className="block-meta">{block.code_count} עמודות קוד</span>
                    <select
                      value={type}
                      onChange={(e) =>
                        setTypeByName((prev) => ({ ...prev, [block.name]: e.target.value }))
                      }
                    >
                      {Object.entries(TYPE_LABELS).map(([value, label]) => (
                        <option key={value} value={value}>
                          {label}
                        </option>
                      ))}
                    </select>
                  </div>

                  {type === 'ab' && !isPaired && (
                    <div className="mapping-warning">
                      לא נמצא זוג מתאים (שם עם סיומת a/b תואמת) — השאלה תטופל כרגילה
                    </div>
                  )}

                  {isPaired && isRoleA && (
                    <div className="mapping-ab-config">
                      <span>מזווג עם: {partnerName} (TOM = {block.name}, אחרים = {partnerName})</span>
                      <label>
                        תשובה לניקוי:
                        <select
                          value={cleanCodeByBase[parsed.base] || ''}
                          onChange={(e) =>
                            setCleanCodeByBase((prev) => ({
                              ...prev,
                              [parsed.base]: e.target.value,
                            }))
                          }
                        >
                          <option value="">בחר תשובה...</option>
                          {unionCategories.map((c) => (
                            <option key={c.code} value={c.code}>
                              {c.code} — {c.label}
                            </option>
                          ))}
                        </select>
                      </label>
                    </div>
                  )}

                  {isPaired && !isRoleA && (
                    <div className="mapping-ab-config">
                      <span>
                        מזווג עם: {partnerName} — התשובה לניקוי נבחרת בשורת {partnerName}
                      </span>
                    </div>
                  )}
                </div>
              )
            })}
          </div>

          {incompletePairs.length > 0 && (
            <div className="error-box">
              יש לבחור "תשובה לניקוי" עבור הזוגות: {incompletePairs.join(', ')}
            </div>
          )}
        </section>
      )}

      {step === 'results' && result && (
        <section>
          <div className="results-header">
            <h2>הופקו {result.blocks.length} קבצים</h2>
            <div className="results-actions">
              <button type="button" className="secondary" onClick={handleDownload}>
                הורדת כל הקבצים (zip)
              </button>
              <button type="button" className="secondary" onClick={handleReset}>
                קובץ חדש
              </button>
            </div>
          </div>

          {result.warnings?.length > 0 && (
            <div className="error-box">
              {result.warnings.map((w) => (
                <div key={w}>{w}</div>
              ))}
            </div>
          )}

          {result.blocks.map((block) => (
            <details className="block-card" key={block.filename}>
              <summary>
                <span className="block-title">
                  {block.question_name}
                  <span className={`type-badge ${block.type}`}>{TYPE_BADGE[block.type] || block.type}</span>
                </span>
                <span className="block-meta">
                  {block.row_count} רשומות · {block.filename}
                </span>
              </summary>
              <div className="table-wrap">
                <table className="preview">
                  <thead>
                    <tr>
                      {block.columns.map((col) => (
                        <th key={col}>{col}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {block.preview_rows.map((row, i) => (
                      <tr key={i}>
                        {row.map((cell, j) => (
                          <td key={j}>{cell ?? ''}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </details>
          ))}
        </section>
      )}

      <footer className="app-footer">
        <details>
          <summary>גרסה {APP_VERSION}</summary>
          <ul>
            {RELEASE_NOTES[0].notes.map((note) => (
              <li key={note}>{note}</li>
            ))}
          </ul>
        </details>
      </footer>
    </>
  )
}

export default App
