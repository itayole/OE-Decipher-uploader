import { useState } from 'react'
import './App.css'
import { APP_VERSION, RELEASE_NOTES } from './releaseNotes'

function App() {
  const [otcFile, setOtcFile] = useState(null)
  const [rawFile, setRawFile] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)
  const [result, setResult] = useState(null)

  const handleSubmit = async (e) => {
    e.preventDefault()
    if (!otcFile) return
    setLoading(true)
    setError(null)
    setResult(null)

    const formData = new FormData()
    formData.append('otc_file', otcFile)

    try {
      const res = await fetch('/api/process', { method: 'POST', body: formData })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || 'שגיאה בעיבוד הקובץ')
      setResult(data)
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  const handleDownload = () => {
    if (result?.job_id) {
      window.location.href = `/api/download/${result.job_id}`
    }
  }

  return (
    <>
      <h1>העלאת שאלות פתוחות מקודדות ל-Decipher</h1>
      <p className="subtitle">
        העלו את קובץ ה-OTC המקודד לפיצול אוטומטי לקבצי .dat, מוכנים להעלאה ל-Decipher.
      </p>

      <form className="upload-card" onSubmit={handleSubmit}>
        <div className="field">
          <label>
            קובץ OTC מקודד (xlsx) <span className="hint">חובה</span>
          </label>
          <input
            type="file"
            accept=".xlsx,.xlsm"
            onChange={(e) => setOtcFile(e.target.files?.[0] ?? null)}
          />
        </div>

        <div className="field">
          <label>
            קובץ נתונים גולמי מ-Decipher (xlsx){' '}
            <span className="hint">אופציונלי — עדיין לא בשימוש בשלב זה</span>
          </label>
          <input
            type="file"
            accept=".xlsx,.xlsm"
            disabled
            onChange={(e) => setRawFile(e.target.files?.[0] ?? null)}
          />
        </div>

        <button type="submit" className="primary" disabled={!otcFile || loading}>
          {loading ? 'מעבד...' : 'עבד קובץ'}
        </button>
      </form>

      {error && <div className="error-box">{error}</div>}

      {result && (
        <section>
          <div className="results-header">
            <h2>נמצאו {result.blocks.length} שאלות פתוחות</h2>
            <button type="button" className="secondary" onClick={handleDownload}>
              הורדת כל הקבצים (zip)
            </button>
          </div>

          {result.blocks.map((block) => (
            <details className="block-card" key={block.question_name} open={false}>
              <summary>
                <span className="block-title">{block.question_name}</span>
                <span className="block-meta">
                  {block.code_count} עמודות קוד · {block.row_count} רשומות · {block.filename}
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
