'use client'
import { useState, useCallback, useRef } from 'react'
import * as XLSX from 'xlsx'
import { buildFinancialModel } from './finance'
import { generateExcel } from './excel'

// ── Formatters ────────────────────────────────────────────────────────────────
const fmt$ = n => n == null || isNaN(n) ? '—' : '$' + Math.round(n).toLocaleString()
const fmtPct = n => n == null || isNaN(n) ? '—' : (n * 100).toFixed(2) + '%'
const fmtX = n => n == null || isNaN(n) ? '—' : n.toFixed(2) + 'x'

// ── Default assumptions ───────────────────────────────────────────────────────
const DEFAULT = {
  propertyName: '', address: '', borough: '', zip: '',
  grossSF: 0, taxes: 0, taxClass: '2',
  listingPrice: 0, investmentValue: 0,
  debt: { loanAmount: 0, rate: 0.065, term: 5, amort: 30 },
  expenses: {
    insurance: 1200, waterSewer: 700, superSalary: 500,
    repairsMaint: 750, commonElectric: 0.25, fuelOil: 900,
    generalAdmin: 0, managementPct: 0.04, vacancyPct: 0.03,
  },
  marketRents: { studio: 0, oneBed: 0, twoBed: 0, threeBed: 0, fourBed: 0, fiveBed: 0 },
  rsGrowthRate: 0.03, fmGrowthRate: 0.04,
}

// ── Status pill ───────────────────────────────────────────────────────────────
function StatusPill({ status }) {
  const colors = {
    FM: { bg: '#dcfce7', color: '#15803d' },
    RS: { bg: '#dbeafe', color: '#1d4ed8' },
    RC: { bg: '#fef9c3', color: '#854d0e' },
  }
  const c = colors[status] || { bg: '#f3f4f6', color: '#374151' }
  return (
    <span style={{ fontSize: 11, fontWeight: 700, padding: '2px 7px', borderRadius: 4, ...c }}>
      {status}
    </span>
  )
}

// ── Labelled input ────────────────────────────────────────────────────────────
function Field({ label, value, onChange, type = 'number', prefix }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
      <label style={{ fontSize: 11, fontWeight: 600, color: '#64748b', letterSpacing: '0.05em', textTransform: 'uppercase' }}>
        {label}
      </label>
      <div style={{ display: 'flex', alignItems: 'center', gap: 0 }}>
        {prefix && (
          <span style={{ padding: '7px 10px', background: '#f1f5f9', border: '1px solid #e2e8f0', borderRight: 'none', borderRadius: '6px 0 0 6px', fontSize: 13, color: '#64748b' }}>
            {prefix}
          </span>
        )}
        <input
          type={type === 'number' ? 'number' : 'text'}
          value={value ?? ''}
          step={type === 'number' ? 'any' : undefined}
          onChange={e => onChange(type === 'number' ? (parseFloat(e.target.value) || 0) : e.target.value)}
          style={{
            flex: 1,
            border: '1px solid #e2e8f0',
            borderRadius: prefix ? '0 6px 6px 0' : 6,
            padding: '7px 10px',
            fontSize: 13,
            fontFamily: "'DM Mono', monospace",
            background: '#fff',
            color: '#0f172a',
            outline: 'none',
            width: '100%',
            minWidth: 0,
          }}
        />
      </div>
    </div>
  )
}

// ── Section card ──────────────────────────────────────────────────────────────
function Section({ title, children }) {
  return (
    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 12, padding: '1.25rem 1.5rem', marginBottom: '1rem' }}>
      {title && (
        <p style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: '#94a3b8', marginBottom: '1rem' }}>
          {title}
        </p>
      )}
      {children}
    </div>
  )
}

function Grid({ cols = 2, children }) {
  return (
    <div style={{ display: 'grid', gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: '0.875rem' }}>
      {children}
    </div>
  )
}

// ── Metric card ───────────────────────────────────────────────────────────────
function Metric({ label, value, accent }) {
  return (
    <div style={{ background: accent ? '#0f172a' : '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '0.875rem 1rem' }}>
      <div style={{ fontSize: 11, color: accent ? '#94a3b8' : '#64748b', marginBottom: 4, fontWeight: 500 }}>{label}</div>
      <div style={{ fontSize: 20, fontWeight: 700, color: accent ? '#fff' : '#0f172a', fontFamily: "'DM Mono', monospace" }}>{value}</div>
    </div>
  )
}

// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep]               = useState('upload')
  const [file, setFile]               = useState(null)
  const [fileContent, setFileContent] = useState(null)
  const [parsing, setParsing]         = useState(false)
  const [parseError, setParseError]   = useState(null)
  const [rentRoll, setRentRoll]       = useState(null)
  const [assumptions, setAssumptions] = useState(DEFAULT)
  const [finData, setFinData]         = useState(null)
  const [building, setBuilding]       = useState(false)
  const [downloadUrl, setDownloadUrl] = useState(null)
  const [fileName, setFileName]       = useState('')
  const fileRef = useRef()

  // ── File handling ───────────────────────────────────────────────────────────
  const handleFile = useCallback(async (f) => {
    if (!f) return
    setFile(f)
    setParseError(null)
    const name = f.name.toLowerCase()
    if (name.endsWith('.csv') || name.endsWith('.txt')) {
      const text = await f.text()
      setFileContent(text)
    } else {
      const buf = await f.arrayBuffer()
      const wb  = XLSX.read(buf, { type: 'array' })
      const ws  = wb.Sheets[wb.SheetNames[0]]
      setFileContent(XLSX.utils.sheet_to_csv(ws))
    }
  }, [])

  const onDrop = useCallback(e => {
    e.preventDefault()
    const f = e.dataTransfer.files[0]
    if (f) handleFile(f)
  }, [handleFile])

  // ── AI parse ────────────────────────────────────────────────────────────────
  const parseWithAI = async () => {
    if (!fileContent) return
    setParsing(true); setParseError(null)
    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 4000,
          messages: [{
            role: 'user',
            content: `Parse this rent roll and return ONLY a JSON object. No markdown, no explanation.

RENT ROLL DATA:
${fileContent.slice(0, 8000)}

Return this exact structure:
{
  "propertyName": "",
  "address": "",
  "borough": "",
  "zip": "",
  "grossSF": 0,
  "taxes": 0,
  "listingPrice": 0,
  "investmentValue": 0,
  "units": [
    {
      "unit": "1",
      "status": "FM",
      "type": "residential",
      "bedrooms": "2 Bedroom",
      "sf": 0,
      "actualRent": 0,
      "leaseExp": "",
      "notes": ""
    }
  ]
}
Rules: FM=free market RS=rent stabilized RC=rent controlled. actualRent is monthly. bedrooms must be exactly: Studio, 1 Bedroom, 2 Bedroom, 3 Bedroom, 4 Bedroom, or 5 Bedroom. Default status to FM if unclear.`
          }],
        }),
      })
      const json = await res.json()
      const raw  = (json.content || []).map(b => b.text || '').join('')
      const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim())

      setRentRoll(parsed.units || [])
      setAssumptions(prev => ({
        ...prev,
        propertyName:   parsed.propertyName   || prev.propertyName,
        address:        parsed.address        || prev.address,
        borough:        parsed.borough        || prev.borough,
        zip:            parsed.zip            || prev.zip,
        grossSF:        parsed.grossSF        || prev.grossSF,
        taxes:          parsed.taxes          || prev.taxes,
        listingPrice:   parsed.listingPrice   || prev.listingPrice,
        investmentValue:parsed.investmentValue|| prev.investmentValue,
      }))
      setStep('assumptions')
    } catch (err) {
      setParseError('Parse failed: ' + err.message)
    } finally {
      setParsing(false)
    }
  }

  // ── Assumption setter helpers ───────────────────────────────────────────────
  const setA = (key, val) => setAssumptions(p => ({ ...p, [key]: val }))
  const setExp = (key, val) => setAssumptions(p => ({ ...p, expenses: { ...p.expenses, [key]: val } }))
  const setDebt = (key, val) => setAssumptions(p => ({ ...p, debt: { ...p.debt, [key]: val } }))
  const setMkt = (key, val) => setAssumptions(p => ({ ...p, marketRents: { ...p.marketRents, [key]: val } }))

  // ── Preview ─────────────────────────────────────────────────────────────────
  const goPreview = () => {
    const m = buildFinancialModel(rentRoll, assumptions)
    setFinData(m)
    setStep('preview')
  }

  // ── Build & download ────────────────────────────────────────────────────────
  const buildAndDownload = async () => {
    setBuilding(true)
    try {
      const m    = buildFinancialModel(rentRoll, assumptions)
      setFinData(m)
      const data = generateExcel(m)
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url  = URL.createObjectURL(blob)
      const name = `${assumptions.propertyName || 'FinancialModel'}_${new Date().toISOString().slice(0,10)}.xlsx`
      setDownloadUrl(url)
      setFileName(name)
      // Auto-trigger download
      const a = document.createElement('a')
      a.href = url; a.download = name; a.click()
      setStep('done')
    } catch (err) {
      setParseError('Build failed: ' + err.message)
    } finally {
      setBuilding(false)
    }
  }

  const reset = () => {
    setStep('upload'); setFile(null); setFileContent(null)
    setRentRoll(null); setFinData(null); setDownloadUrl(null)
    setParseError(null); setAssumptions(DEFAULT)
  }

  // ── Styles ──────────────────────────────────────────────────────────────────
  const cf = finData?.cashFlow
  const pr = finData?.pricing
  const dcf= finData?.dcf

  const STEPS = ['Upload', 'Assumptions', 'Preview', 'Download']
  const stepIdx = ['upload', 'assumptions', 'preview', 'done'].indexOf(step)

  return (
    <div style={{ minHeight: '100vh', background: '#f8fafc', padding: '2rem 1rem' }}>
      <div style={{ maxWidth: 920, margin: '0 auto' }}>

        {/* Header */}
        <div style={{ marginBottom: '2rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 8 }}>
            <div style={{ width: 36, height: 36, background: '#0f172a', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18 }}>
              🏢
            </div>
            <h1 style={{ fontSize: 24, fontWeight: 700, letterSpacing: '-0.025em', color: '#0f172a' }}>
              Rent Roll → Financial Model
            </h1>
          </div>
          <p style={{ fontSize: 14, color: '#64748b', marginLeft: 48 }}>
            Upload any rent roll · AI extracts the data · Download a complete 6-sheet Excel model
          </p>
        </div>

        {/* Step tracker */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 0, marginBottom: '1.5rem', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '0.75rem 1.25rem' }}>
          {STEPS.map((s, i) => (
            <div key={s} style={{ display: 'flex', alignItems: 'center', flex: i < 3 ? 1 : 0 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{
                  width: 26, height: 26, borderRadius: '50%',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontSize: 12, fontWeight: 700,
                  background: i < stepIdx ? '#0f172a' : i === stepIdx ? '#3b82f6' : '#e2e8f0',
                  color: i <= stepIdx ? '#fff' : '#94a3b8',
                }}>
                  {i < stepIdx ? '✓' : i + 1}
                </div>
                <span style={{ fontSize: 13, fontWeight: i === stepIdx ? 600 : 400, color: i === stepIdx ? '#0f172a' : i < stepIdx ? '#64748b' : '#94a3b8' }}>
                  {s}
                </span>
              </div>
              {i < 3 && <div style={{ flex: 1, height: 1, background: '#e2e8f0', margin: '0 12px' }} />}
            </div>
          ))}
        </div>

        {/* ── STEP 1: Upload ── */}
        {step === 'upload' && (
          <Section>
            <div
              onDrop={onDrop}
              onDragOver={e => e.preventDefault()}
              onClick={() => fileRef.current?.click()}
              style={{
                border: `2px dashed ${file ? '#86efac' : '#cbd5e1'}`,
                borderRadius: 10,
                padding: '3rem 2rem',
                textAlign: 'center',
                cursor: 'pointer',
                background: file ? '#f0fdf4' : '#f8fafc',
                transition: 'all 0.15s',
              }}
            >
              <div style={{ fontSize: 40, marginBottom: 12 }}>{file ? '✅' : '📄'}</div>
              <p style={{ fontSize: 16, fontWeight: 600, color: '#1e293b', margin: '0 0 6px' }}>
                {file ? file.name : 'Drop your rent roll here'}
              </p>
              <p style={{ fontSize: 13, color: '#94a3b8', margin: 0 }}>
                Supports CSV, Excel (.xlsx / .xls) · Click to browse
              </p>
              <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,.txt" style={{ display: 'none' }} onChange={e => handleFile(e.target.files[0])} />
            </div>

            {file && !parsing && (
              <div style={{ marginTop: '1.25rem', display: 'flex', gap: 10 }}>
                <button
                  onClick={parseWithAI}
                  style={{ background: '#0f172a', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 22px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}
                >
                  Parse with AI →
                </button>
                <button
                  onClick={() => { setFile(null); setFileContent(null) }}
                  style={{ background: '#fff', color: '#374151', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 18px', fontSize: 14, cursor: 'pointer' }}
                >
                  Clear
                </button>
              </div>
            )}

            {parsing && (
              <div style={{ marginTop: '1.25rem', display: 'flex', alignItems: 'center', gap: 10 }}>
                <div style={{ width: 18, height: 18, border: '2.5px solid #3b82f6', borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} />
                <span style={{ fontSize: 14, color: '#64748b' }}>AI is reading your rent roll…</span>
              </div>
            )}

            {parseError && <p style={{ marginTop: 12, color: '#dc2626', fontSize: 13 }}>{parseError}</p>}

            <div style={{ marginTop: '1.5rem', padding: '1rem 1.25rem', background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0' }}>
              <p style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase', color: '#94a3b8', marginBottom: 8 }}>
                Expected columns
              </p>
              <code style={{ fontSize: 12, color: '#475569', display: 'block', background: '#e2e8f0', padding: '8px 12px', borderRadius: 6, lineHeight: 1.6 }}>
                Unit # · Status (FM/RS/RC) · Bedrooms · SF · Monthly Rent · Lease Expiration
              </code>
              <p style={{ fontSize: 12, color: '#94a3b8', marginTop: 8 }}>Column names don't need to match exactly — AI figures it out.</p>
            </div>

            <style>{`@keyframes spin { to { transform: rotate(360deg) } }`}</style>
          </Section>
        )}

        {/* ── STEP 2: Assumptions ── */}
        {step === 'assumptions' && rentRoll && (
          <>
            {/* Parsed units table */}
            <Section title={`Parsed Units · ${rentRoll.length} total`}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12, marginTop: -4 }}>
                <span style={{ fontSize: 13, color: '#64748b' }}>
                  {rentRoll.filter(u=>u.type!=='commercial').length} residential · {rentRoll.filter(u=>u.type==='commercial').length} commercial
                </span>
              </div>
              <div style={{ maxHeight: 220, overflowY: 'auto', borderRadius: 6, border: '1px solid #e2e8f0' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: '#f8fafc', position: 'sticky', top: 0 }}>
                      {['Unit', 'Status', 'Bedrooms', 'SF', 'Actual Rent/Mo', 'Lease Exp', 'Notes'].map(h => (
                        <th key={h} style={{ padding: '7px 10px', textAlign: 'left', fontSize: 11, fontWeight: 600, color: '#64748b', borderBottom: '1px solid #e2e8f0', whiteSpace: 'nowrap' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rentRoll.map((u, i) => (
                      <tr key={i} style={{ borderBottom: '1px solid #f1f5f9' }}>
                        <td style={{ padding: '6px 10px', fontWeight: 600 }}>{u.unit}</td>
                        <td style={{ padding: '6px 10px' }}><StatusPill status={u.status} /></td>
                        <td style={{ padding: '6px 10px', color: '#475569' }}>{u.bedrooms}</td>
                        <td style={{ padding: '6px 10px', fontFamily: 'DM Mono', color: '#475569' }}>{u.sf?.toLocaleString()}</td>
                        <td style={{ padding: '6px 10px', fontFamily: 'DM Mono', fontWeight: 600 }}>{fmt$(u.actualRent)}</td>
                        <td style={{ padding: '6px 10px', color: '#94a3b8', fontSize: 12 }}>{u.leaseExp}</td>
                        <td style={{ padding: '6px 10px', color: '#94a3b8', fontSize: 12 }}>{u.notes}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Section>

            {/* Property info */}
            <Section title="Property Information">
              <Grid cols={2}>
                <Field label="Property Name" value={assumptions.propertyName} onChange={v=>setA('propertyName',v)} type="text" />
                <Field label="Address" value={assumptions.address} onChange={v=>setA('address',v)} type="text" />
                <Field label="Borough / City" value={assumptions.borough} onChange={v=>setA('borough',v)} type="text" />
                <Field label="Zip Code" value={assumptions.zip} onChange={v=>setA('zip',v)} type="text" />
                <Field label="Gross Square Footage" value={assumptions.grossSF} onChange={v=>setA('grossSF',v)} />
                <Field label="Annual Property Taxes" value={assumptions.taxes} onChange={v=>setA('taxes',v)} prefix="$" />
                <Field label="Tax Class" value={assumptions.taxClass} onChange={v=>setA('taxClass',v)} type="text" />
              </Grid>
            </Section>

            {/* Market rents */}
            <Section title="Market Rate Rents (Monthly)">
              <Grid cols={3}>
                <Field label="Studio" value={assumptions.marketRents.studio} onChange={v=>setMkt('studio',v)} prefix="$" />
                <Field label="1 Bedroom" value={assumptions.marketRents.oneBed} onChange={v=>setMkt('oneBed',v)} prefix="$" />
                <Field label="2 Bedroom" value={assumptions.marketRents.twoBed} onChange={v=>setMkt('twoBed',v)} prefix="$" />
                <Field label="3 Bedroom" value={assumptions.marketRents.threeBed} onChange={v=>setMkt('threeBed',v)} prefix="$" />
                <Field label="4 Bedroom" value={assumptions.marketRents.fourBed} onChange={v=>setMkt('fourBed',v)} prefix="$" />
                <Field label="5 Bedroom" value={assumptions.marketRents.fiveBed} onChange={v=>setMkt('fiveBed',v)} prefix="$" />
              </Grid>
            </Section>

            {/* Expenses */}
            <Section title="Expense Assumptions">
              <Grid cols={3}>
                <Field label="Insurance ($/unit)" value={assumptions.expenses.insurance} onChange={v=>setExp('insurance',v)} prefix="$" />
                <Field label="Water & Sewer ($/unit)" value={assumptions.expenses.waterSewer} onChange={v=>setExp('waterSewer',v)} prefix="$" />
                <Field label="Super Salary ($/unit)" value={assumptions.expenses.superSalary} onChange={v=>setExp('superSalary',v)} prefix="$" />
                <Field label="Repairs & Maint. ($/unit)" value={assumptions.expenses.repairsMaint} onChange={v=>setExp('repairsMaint',v)} prefix="$" />
                <Field label="Fuel – Oil ($/unit)" value={assumptions.expenses.fuelOil} onChange={v=>setExp('fuelOil',v)} prefix="$" />
                <Field label="Common Electric ($/sf)" value={assumptions.expenses.commonElectric} onChange={v=>setExp('commonElectric',v)} prefix="$" />
                <Field label="General Admin ($/unit)" value={assumptions.expenses.generalAdmin} onChange={v=>setExp('generalAdmin',v)} prefix="$" />
                <Field label="Management Fee" value={assumptions.expenses.managementPct} onChange={v=>setExp('managementPct',v)} />
                <Field label="Vacancy Rate" value={assumptions.expenses.vacancyPct} onChange={v=>setExp('vacancyPct',v)} />
              </Grid>
            </Section>

            {/* Debt */}
            <Section title="Debt / Financing">
              <Grid cols={4}>
                <Field label="Loan Amount" value={assumptions.debt.loanAmount} onChange={v=>setDebt('loanAmount',v)} prefix="$" />
                <Field label="Interest Rate" value={assumptions.debt.rate} onChange={v=>setDebt('rate',v)} />
                <Field label="Term (years)" value={assumptions.debt.term} onChange={v=>setDebt('term',v)} />
                <Field label="Amortization (years)" value={assumptions.debt.amort} onChange={v=>setDebt('amort',v)} />
              </Grid>
            </Section>

            {/* Pricing */}
            <Section title="Pricing & Growth">
              <Grid cols={4}>
                <Field label="Listing Price" value={assumptions.listingPrice} onChange={v=>setA('listingPrice',v)} prefix="$" />
                <Field label="Investment Value" value={assumptions.investmentValue} onChange={v=>setA('investmentValue',v)} prefix="$" />
                <Field label="RS Growth Rate" value={assumptions.rsGrowthRate} onChange={v=>setA('rsGrowthRate',v)} />
                <Field label="FM Growth Rate" value={assumptions.fmGrowthRate} onChange={v=>setA('fmGrowthRate',v)} />
              </Grid>
            </Section>

            <div style={{ display: 'flex', gap: 10 }}>
              <button onClick={goPreview} style={{ background: '#0f172a', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 22px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}>
                Preview Model →
              </button>
              <button onClick={() => setStep('upload')} style={{ background: '#fff', color: '#374151', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 18px', fontSize: 14, cursor: 'pointer' }}>
                ← Back
              </button>
            </div>
          </>
        )}

        {/* ── STEP 3: Preview ── */}
        {step === 'preview' && finData && (
          <>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '0.75rem', marginBottom: '1rem' }}>
              <Metric label="NOI (Current)" value={fmt$(cf?.noi)} />
              <Metric label="NOI (Pro Forma)" value={fmt$(cf?.pfNoi)} accent />
              <Metric label="Current Cap Rate" value={fmtPct(pr?.listCapCurrent)} />
              <Metric label="Pro Forma Cap Rate" value={fmtPct(pr?.listCapPF)} accent />
              <Metric label="GRM (Current)" value={fmtX(pr?.listGRMCurrent)} />
              <Metric label="GRM (Pro Forma)" value={fmtX(pr?.listGRMPF)} />
              <Metric label="Cash-on-Cash (PF)" value={fmtPct(pr?.listCoCPF)} />
              <Metric label="DSCR (Pro Forma)" value={fmtX(cf && cf.annualDebtService > 0 ? cf.pfNoi / cf.annualDebtService : null)} />
              <Metric label="Unlevered IRR" value={fmtPct(dcf?.unlevIRR)} accent />
              <Metric label="Levered IRR" value={fmtPct(dcf?.levIRR)} accent />
              <Metric label="Equity Multiple" value={dcf?.equityMultiple ? fmtX(dcf.equityMultiple) : '—'} />
              <Metric label="Exit Value (Yr 10)" value={fmt$(dcf?.reversionProceeds)} />
            </div>

            {/* I&E table */}
            <Section title="Income & Expense Summary">
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <thead>
                  <tr>
                    {['', 'Current Annual', 'Pro Forma Annual'].map((h,i) => (
                      <th key={h} style={{ padding: '6px 10px', textAlign: i > 0 ? 'right' : 'left', fontSize: 11, fontWeight: 600, color: '#64748b', borderBottom: '1px solid #e2e8f0' }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[
                    ['Gross Potential Residential Rent', cf?.gprResi,   cf?.pfGprResi,   false],
                    ['Gross Potential Commercial Rent',  cf?.gprComm,   cf?.pfGprComm,   false],
                    ['Gross Income',                     cf?.grossIncome, cf?.pfGrossIncome, false],
                    ['Vacancy / Collection Loss',        -(cf?.vacancy||0), -(cf?.pfVacancy||0), false],
                    ['Effective Gross Income',           cf?.egi,   cf?.pfEgi,   true],
                    ['Total Expenses',                   -(cf?.totalExp||0), -(cf?.pfTotalExp||0), false],
                    ['Net Operating Income',             cf?.noi,   cf?.pfNoi,   true],
                    ['Annual Debt Service',              -(cf?.annualDebtService||0), -(cf?.annualDebtService||0), false],
                    ['Cash Flow After Debt',             cf?.cashFlowAfterDebt, cf?.pfCashFlowAfterDebt, true],
                  ].map(([label, cur, pf, bold]) => (
                    <tr key={label} style={{ borderBottom: '1px solid #f8fafc', background: bold ? '#f8fafc' : 'transparent' }}>
                      <td style={{ padding: '7px 10px', fontWeight: bold ? 700 : 400 }}>{label}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right', fontFamily: 'DM Mono', fontWeight: bold ? 700 : 400, color: (cur||0) < 0 ? '#dc2626' : '#0f172a' }}>{fmt$(cur)}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right', fontFamily: 'DM Mono', fontWeight: bold ? 700 : 400, color: (pf||0) < 0 ? '#dc2626' : '#0f172a' }}>{fmt$(pf)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Section>

            {/* 10yr bar */}
            <Section title="10-Year NOI Projection">
              <div style={{ display: 'flex', alignItems: 'flex-end', gap: 4, height: 100 }}>
                {dcf?.noi?.map((v, i) => {
                  const max = Math.max(...(dcf.noi.filter(Boolean)))
                  const h = max > 0 ? Math.max(4, Math.round((v / max) * 90)) : 4
                  return (
                    <div key={i} style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
                      <div style={{ width: '100%', height: h, background: i === 0 ? '#cbd5e1' : `hsl(${210 + i * 5}, 60%, ${40 - i}%)`, borderRadius: '3px 3px 0 0', transition: 'height 0.4s ease' }} title={fmt$(v)} />
                      <span style={{ fontSize: 9, color: '#94a3b8' }}>Y{i}</span>
                    </div>
                  )
                })}
              </div>
            </Section>

            <div style={{ display: 'flex', gap: 10 }}>
              <button
                onClick={buildAndDownload}
                disabled={building}
                style={{ background: '#16a34a', color: '#fff', border: 'none', borderRadius: 8, padding: '11px 24px', fontSize: 14, fontWeight: 600, cursor: building ? 'wait' : 'pointer', opacity: building ? 0.7 : 1 }}
              >
                {building ? 'Building Excel…' : '⬇ Build & Download Excel'}
              </button>
              <button onClick={() => setStep('assumptions')} style={{ background: '#fff', color: '#374151', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 18px', fontSize: 14, cursor: 'pointer' }}>
                ← Edit Assumptions
              </button>
            </div>
            {parseError && <p style={{ marginTop: 10, color: '#dc2626', fontSize: 13 }}>{parseError}</p>}
          </>
        )}

        {/* ── STEP 4: Done ── */}
        {step === 'done' && finData && (
          <>
            <Section>
              <div style={{ textAlign: 'center', padding: '2rem 1rem' }}>
                <div style={{ fontSize: 52, marginBottom: 12 }}>📥</div>
                <h2 style={{ fontSize: 22, fontWeight: 700, letterSpacing: '-0.02em', margin: '0 0 8px' }}>
                  Download started automatically
                </h2>
                <p style={{ fontSize: 14, color: '#64748b', margin: '0 0 6px' }}>
                  Your file <strong>{fileName}</strong> should be in your Downloads folder.
                </p>
                <p style={{ fontSize: 13, color: '#94a3b8', margin: '0 0 1.5rem' }}>
                  6 sheets: INPUTS · RENT ROLL · CASH FLOW · PRICING · 10 YEAR DCF · VALUE SUMMARY
                </p>
                <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
                  {downloadUrl && (
                    <a href={downloadUrl} download={fileName}
                      style={{ background: '#0f172a', color: '#fff', borderRadius: 8, padding: '10px 22px', fontSize: 14, fontWeight: 600, textDecoration: 'none', display: 'inline-block' }}>
                      ⬇ Download Again
                    </a>
                  )}
                  <button onClick={reset} style={{ background: '#fff', color: '#374151', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 18px', fontSize: 14, cursor: 'pointer' }}>
                    ← New Deal
                  </button>
                </div>
              </div>
            </Section>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '0.75rem' }}>
              {[
                ['NOI (Current)', fmt$(cf?.noi)],
                ['NOI (Pro Forma)', fmt$(cf?.pfNoi)],
                ['Cap Rate (PF)', fmtPct(pr?.listCapPF)],
                ['Levered IRR', fmtPct(dcf?.levIRR)],
              ].map(([l, v]) => <Metric key={l} label={l} value={v} />)}
            </div>
          </>
        )}

      </div>
    </div>
  )
}
