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
  xml: 'XML',
}

function parseAbSuffix(name) {
  const m = /^(.*?)([aAbB])$/.exec(name)
  if (!m) return null
  return { base: m[1], letter: m[2].toLowerCase() }
}

const DONT_KNOW_KEYWORDS = [
  'לא ראיתי',
  'לא ראתה',
  'לא מכיר',
  'לא מכירה',
  'לא מוכר',
  'לא יודע',
  'לא יודעת',
  'אינני מכיר',
  'איני מכיר',
  'לא זוכר',
  'לא זוכרת',
  "don't know",
  'dont know',
  'not familiar',
  'unfamiliar',
]

// Words that mark a conditional/hedged answer (e.g. "ראיתי אבל לא זוכר של מי")
// rather than a plain "don't know" — these should be deprioritized.
const CONDITIONAL_MARKERS = ['אבל', 'אך ', 'אולם', 'רק ']

function guessDontKnowCode(categories) {
  const candidates = categories.filter((c) =>
    DONT_KNOW_KEYWORDS.some((kw) => (c.label || '').toLowerCase().includes(kw))
  )
  if (!candidates.length) return null

  const simple = candidates.filter(
    (c) => !CONDITIONAL_MARKERS.some((kw) => (c.label || '').includes(kw))
  )
  const pool = simple.length ? simple : candidates
  const best = [...pool].sort((a, b) => (a.label || '').length - (b.label || '').length)[0]
  return best.code
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

function DropzoneInput({ label, hint, file, disabled, onFile }) {
  const [isDragging, setIsDragging] = useState(false)

  const handleDragOver = (e) => {
    if (disabled) return
    e.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = () => setIsDragging(false)

  const handleDrop = (e) => {
    e.preventDefault()
    setIsDragging(false)
    if (disabled) return
    const dropped = e.dataTransfer.files?.[0]
    if (dropped) onFile(dropped)
  }

  return (
    <div className="field">
      <label>
        {label} <span className="hint">{hint}</span>
      </label>
      <label
        className={`dropzone ${disabled ? 'disabled' : ''} ${isDragging ? 'dragging' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <span className="dropzone-icon">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M12 16V4M12 4l-4 4M12 4l4 4" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </span>
        <span className="dropzone-text">
          {file ? (
            <strong>{file.name}</strong>
          ) : disabled ? (
            'לא זמין בשלב זה'
          ) : (
            'לחצו לבחירת קובץ xlsx או גררו אותו לכאן'
          )}
        </span>
        <input
          type="file"
          accept=".xlsx,.xlsm"
          disabled={disabled}
          onChange={(e) => onFile(e.target.files?.[0] ?? null)}
        />
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

  const [adHocTemplateFile, setAdHocTemplateFile] = useState(null)
  const [templateUpdateFile, setTemplateUpdateFile] = useState(null)
  const [templateStatus, setTemplateStatus] = useState(null) // {type: 'success'|'error', message}
  const [templateUpdating, setTemplateUpdating] = useState(false)

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

      const blocksByNameTemp = {}
      data.blocks.forEach((b) => (blocksByNameTemp[b.name] = b))
      const initialCleanCodes = {}
      data.blocks.forEach((b) => {
        if (b.suggested_type !== 'ab' || b.role !== 'A' || !b.paired_with) return
        const parsed = parseAbSuffix(b.name)
        if (!parsed) return
        const partner = blocksByNameTemp[b.paired_with]
        const seen = new Map()
        ;[...(b.categories || []), ...(partner?.categories || [])].forEach((c) => {
          if (!seen.has(c.code)) seen.set(c.code, c)
        })
        const guess = guessDontKnowCode([...seen.values()])
        if (guess) initialCleanCodes[parsed.base] = guess
      })
      setCleanCodeByBase(initialCleanCodes)
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

    const formData = new FormData()
    formData.append('job_id', jobId)
    formData.append('mapping', JSON.stringify(mapping))
    if (adHocTemplateFile) {
      formData.append('xml_template_file', adHocTemplateFile)
    }

    try {
      const res = await fetch('/api/generate', { method: 'POST', body: formData })
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

  const handleDownloadTemplate = () => {
    window.location.href = '/api/xml-template'
  }

  const handleUpdateTemplate = async () => {
    if (!templateUpdateFile) return
    setTemplateUpdating(true)
    setTemplateStatus(null)
    const formData = new FormData()
    formData.append('template_file', templateUpdateFile)
    try {
      const res = await fetch('/api/xml-template', { method: 'POST', body: formData })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || 'שגיאה בעדכון התבנית')
      setTemplateStatus({ type: 'success', message: 'תבנית ברירת המחדל עודכנה בהצלחה' })
      setTemplateUpdateFile(null)
    } catch (err) {
      setTemplateStatus({ type: 'error', message: err.message })
    } finally {
      setTemplateUpdating(false)
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
    setAdHocTemplateFile(null)
    setTemplateStatus(null)
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
          <h1>הכנת שאלות פתוחות להעלאה ל-Decipher</h1>
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
            onFile={(f) => setOtcFile(f ?? null)}
          />

          <DropzoneInput
            label="קובץ נתונים גולמי מ-Decipher (xlsx)"
            hint="אופציונלי — עדיין לא בשימוש בשלב זה"
            file={rawFile}
            disabled
            onFile={(f) => setRawFile(f ?? null)}
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
                    <span className="block-meta">
                      {block.code_count} עמודות קוד · {block.answered_count} תשובות לא ריקות
                    </span>
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

          <details className="template-panel">
            <summary>תבנית XML לייצוא</summary>
            <div className="template-panel-body">
              <div className="template-row">
                <div>
                  <strong>תבנית ברירת המחדל של השרת</strong>
                  <p className="hint-block">משמשת לכל הייצואים, אלא אם הועלתה תבנית חד-פעמית למטה.</p>
                </div>
                <button type="button" className="secondary" onClick={handleDownloadTemplate}>
                  הורדת התבנית הנוכחית
                </button>
              </div>

              <div className="template-row">
                <div>
                  <strong>עדכון תבנית ברירת המחדל (קבוע)</strong>
                  <p className="hint-block">יחליף את התבנית עבור כל הייצואים העתידיים, עד לעדכון הבא.</p>
                </div>
                <div className="template-actions">
                  <input
                    type="file"
                    accept=".txt,.xml"
                    onChange={(e) => setTemplateUpdateFile(e.target.files?.[0] ?? null)}
                  />
                  <button
                    type="button"
                    className="secondary"
                    disabled={!templateUpdateFile || templateUpdating}
                    onClick={handleUpdateTemplate}
                  >
                    {templateUpdating ? 'מעדכן...' : 'עדכון התבנית'}
                  </button>
                </div>
              </div>

              <div className="template-row">
                <div>
                  <strong>תבנית חד-פעמית לייצוא זה בלבד</strong>
                  <p className="hint-block">
                    לא נשמרת בשרת — תשמש רק להפקת הקבצים הבאה
                    {adHocTemplateFile ? ` (${adHocTemplateFile.name})` : ''}.
                  </p>
                </div>
                <div className="template-actions">
                  <input
                    type="file"
                    accept=".txt,.xml"
                    onChange={(e) => setAdHocTemplateFile(e.target.files?.[0] ?? null)}
                  />
                  {adHocTemplateFile && (
                    <button type="button" className="secondary" onClick={() => setAdHocTemplateFile(null)}>
                      ביטול
                    </button>
                  )}
                </div>
              </div>

              {templateStatus && (
                <div className={templateStatus.type === 'error' ? 'error-box' : 'success-box'}>
                  {templateStatus.message}
                </div>
              )}
            </div>
          </details>
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
                  {block.row_count != null && `${block.row_count} רשומות`}
                  {block.answered_count != null && ` · ${block.answered_count} תשובות לא ריקות`}
                  {block.row_count != null || block.answered_count != null ? ' · ' : ''}
                  {block.filename}
                </span>
              </summary>
              {block.type === 'xml' ? (
                <div className="xml-preview-wrap">
                  <pre className="xml-preview">{block.text_preview}</pre>
                </div>
              ) : (
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
              )}
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
