import { useEffect, useRef, useState } from 'react'
import { downloadUrl, getDefaultSpec, getJob, jsonUrl, startJob } from './api.js'

const STAGE_LABEL = {
  waiting: 'Waiting',
  done: 'Done',
  failed: 'Failed',
}

function ReportRow({ name, r }) {
  const pct = r.sheets_total ? Math.round((r.sheets_done / r.sheets_total) * 100) : 0
  const busy = !['done', 'failed', 'waiting'].includes(r.stage)
  return (
    <li className={`report ${r.stage}`}>
      <div className="report-top">
        <span className="file">{name}</span>
        <span className="year">{r.year ? `FY${r.year}` : ''}</span>
      </div>
      <div className="report-stage">
        {STAGE_LABEL[r.stage] || r.stage}
        {busy && r.sheets_total ? ` · ${r.sheets_done}/${r.sheets_total} sheets` : ''}
      </div>
      {busy && <div className="bar"><div style={{ width: `${pct}%` }} /></div>}
      {r.error && <div className="error small">{r.error}</div>}
    </li>
  )
}

export default function App() {
  const [company, setCompany] = useState('')
  const [companyId, setCompanyId] = useState('')
  const [files, setFiles] = useState([])
  const [spec, setSpec] = useState('')
  const [job, setJob] = useState(null)
  const [error, setError] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const timer = useRef(null)

  useEffect(() => {
    getDefaultSpec().then(setSpec).catch(() => {})
    const saved = new URLSearchParams(window.location.search).get('job')
    if (saved) poll(saved)
    return () => clearTimeout(timer.current)
  }, [])

  async function poll(id) {
    try {
      const data = await getJob(id)
      setJob(data)
      if (data.status === 'queued' || data.status === 'running') {
        timer.current = setTimeout(() => poll(id), 2000)
      }
    } catch (e) {
      setError(e.message)
    }
  }

  async function submit(e) {
    e.preventDefault()
    setError('')
    if (!company.trim()) return setError('Enter the company name.')
    if (!files.length) return setError('Choose at least one PDF.')
    const form = new FormData()
    form.append('company_name', company.trim())
    form.append('company_id', companyId.trim())
    form.append('spec_text', spec)
    files.forEach((f) => form.append('files', f))
    setSubmitting(true)
    try {
      const { job_id } = await startJob(form)
      window.history.replaceState(null, '', `?job=${job_id}`)
      poll(job_id)
    } catch (err) {
      setError(err.message)
    } finally {
      setSubmitting(false)
    }
  }

  function reset() {
    clearTimeout(timer.current)
    setJob(null)
    setFiles([])
    window.history.replaceState(null, '', window.location.pathname)
  }

  const running = job && (job.status === 'queued' || job.status === 'running')

  return (
    <main className="page">
      <h1>Annual Report Extractor</h1>
      <p className="sub">Upload a company's annual reports (one PDF per year). The agent extracts each year,
        verifies it, and joins all years into one Excel file.</p>

      {!job && (
        <form className="card" onSubmit={submit}>
          <label>Company name
            <input value={company} onChange={(e) => setCompany(e.target.value)} placeholder="e.g. MTN Nigeria Communications Plc" />
          </label>

          <label>Company ID <span className="muted">(optional)</span>
            <input value={companyId} onChange={(e) => setCompanyId(e.target.value)} placeholder="e.g. MTNN" />
          </label>

          <label>Annual report PDFs
            <input type="file" accept="application/pdf,.pdf" multiple
              onChange={(e) => setFiles(Array.from(e.target.files || []))} />
          </label>
          {files.length > 0 && (
            <ul className="files">{files.map((f) => <li key={f.name}>{f.name}</li>)}</ul>
          )}

          <details>
            <summary>Sheets &amp; columns to extract</summary>
            <textarea rows={10} value={spec} onChange={(e) => setSpec(e.target.value)} spellCheck={false} />
            <p className="muted small">One line per sheet: <code>1. SHEET_NAME: column1, column2, ...</code></p>
          </details>

          {error && <div className="error">{error}</div>}
          <button type="submit" disabled={submitting}>{submitting ? 'Uploading…' : 'Run extraction'}</button>
        </form>
      )}

      {job && (
        <section className="card">
          <div className="status-line">
            <strong>{job.inputs?.company_name}</strong>
            <span className={`pill ${job.status}`}>{running ? job.stage : job.status}</span>
          </div>

          <ul className="reports">
            {Object.entries(job.reports || {}).map(([name, r]) => <ReportRow key={name} name={name} r={r} />)}
          </ul>

          {job.warnings?.length > 0 && (
            <div className="warn">
              {job.warnings.map((w, i) => <div key={i}><b>{w.check}:</b> {w.detail}</div>)}
            </div>
          )}
          {job.error && <div className="error">{job.error}</div>}

          {job.download_ready && (
            <div className="actions">
              <a className="button" href={downloadUrl(job.id)}>Download Excel</a>
              <a className="link" href={jsonUrl(job.id)}>JSON</a>
            </div>
          )}
          {!running && <button className="secondary" onClick={reset}>New extraction</button>}
          {error && <div className="error">{error}</div>}
        </section>
      )}
    </main>
  )
}
