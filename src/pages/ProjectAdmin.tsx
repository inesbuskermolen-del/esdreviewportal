import { useEffect, useState, useRef, useCallback, useMemo } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import axios from 'axios'
import NavBar from '@/components/NavBar'
import { CreditInfoButton } from '@/components/CreditInfoButton'
import type { Project, Credit, DrawingRequirement, Reviewer, ESDExcellenceOpportunity, RevisionSummary } from '@/types'

type Tab = 'matrix' | 'drawings'
type GenStatus = 'idle' | 'running' | 'complete' | 'error'

const GDFT_CREDIT_NAMES = [
  'Ventilation - Natural - Apartments',
  'Thermal Performance Rating - Residential',
]

// Maps innovation credit name → GDFT principle (null = N/A, no badge shown)
const INNOVATION_GDFT: Record<string, string> = {
  'ESD As-built verification':                       'High Performance Thermal Envelope',
  'Design for Disassembly Plan':                     'Material Circularity',
  'Battery Storage':                                 'Net-Positive Energy',
  'Triple Glazing':                                  'High Performance Thermal Envelope',
  'Grey water recycling':                            'Water Resilience',
  'Low GWP / No Refrigerants HHW':                  'Life Cycle Decarbonisation',
  'Low GWP / No Refrigerants HVAC':                 'Life Cycle Decarbonisation',
  'Carbon Neutral / Low Carbon Concrete':            'Life Cycle Decarbonisation',
  'Carbon Neutral Power Agreement – Base Build':     'Net-Positive Energy',
  'Carbon Neutral Power Agreement – Apartments':     'Net-Positive Energy',
  'Construction Waste Reduction':                    'Zero Waste',
  'Building Integrated Solar PV':                    'Net-Positive Energy',
  'Airtightness testing (sample apartments)':        'High Performance Thermal Envelope',
  'Airtightness testing (whole building)':           'High Performance Thermal Envelope',
  'LCA':                                             'Life Cycle Decarbonisation',
  'Embodied Carbon Assessment':                      'Life Cycle Decarbonisation',
  'Material Passport':                               'Material Circularity',
  'Micro grid':                                      'Net-Positive Energy',
  'Centralised HRV/ERV system':                      'High Performance Thermal Envelope',
  'Composter':                                       'Zero Waste',
  'Electric Bike Fleet and E-Bike Charging Stations':'Life Cycle Decarbonisation',
  'Water leak detection system':                     'Water Resilience',
  'Zero waste strategy':                             'Zero Waste',
  'Solshare system':                                 'Net-Positive Energy',
  'Recycled/reused/repurposed materials':            'Material Circularity',
  'Share Economy':                                   'Material Circularity',
  'Formal Pre-plaster inspection':                   'High Performance Thermal Envelope',
  'Bicycle Repair Station':                          'Material Circularity',
  'Smart Grid Integration':                          'Net-Positive Energy',
  'Natural Insulation Materials':                    'Life Cycle Decarbonisation',
  'Geothermal Systems':                              'Life Cycle Decarbonisation',
  'Demand Response Systems':                         'Net-Positive Energy',
}

const VALID_STATUSES = new Set<GenStatus>(['idle', 'running', 'complete', 'error'])

function toGenStatus(s: string): GenStatus {
  return VALID_STATUSES.has(s as GenStatus) ? (s as GenStatus) : 'idle'
}

export default function ProjectAdmin() {
  const { id } = useParams<{ id: string }>()
  const navigate = useNavigate()

  const [project, setProject] = useState<Project | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [activeTab, setActiveTab] = useState<Tab>('matrix')
  const [genStatus, setGenStatus] = useState<GenStatus>('idle')
  const [exportLoading, setExportLoading] = useState(false)
  const [excellenceItems, setExcellenceItems] = useState<ESDExcellenceOpportunity[]>([])
  const [computedBESS, setComputedBESS] = useState<number | null>(null)
  const [localFlags, setLocalFlags] = useState<Record<string, string>>({})
  const [flagging, setFlagging] = useState<Record<string, boolean>>({})
  const [inviteEmail, setInviteEmail] = useState('')
  const [inviteName, setInviteName] = useState('')
  const [inviteDiscipline, setInviteDiscipline] = useState('')
  const [inviting, setInviting] = useState(false)
  const [inviteError, setInviteError] = useState('')
  const [inviteSuccess, setInviteSuccess] = useState(false)
  const [inviteWarning, setInviteWarning] = useState('')
  const [notifyEmail, setNotifyEmail] = useState('')
  const [notifyEmailSaved, setNotifyEmailSaved] = useState(false)
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null)

  const interactiveBESS = useMemo(() => {
    if (project?.bessScore == null) return null
    let score = project.bessScore
    for (const item of excellenceItems) {
      const flag = localFlags[item.id] ?? item.flag
      if (flag !== 'Yes') continue
      if (item.creditReference === 'Innovation' && item.bessPoints) {
        const raw = Number(item.bessPoints)
        if (!isNaN(raw)) score += Math.round(raw * 10) / 10
      } else if (item.additionalBessPoints != null) {
        score += item.additionalBessPoints
      }
    }
    return Math.round(score)
  }, [project?.bessScore, excellenceItems, localFlags])

  const stopPolling = useCallback(() => {
    if (pollRef.current) {
      clearInterval(pollRef.current)
      pollRef.current = null
    }
  }, [])

  // Load excellence items from dedicated endpoint
  const loadExcellence = useCallback(async () => {
    if (!id) return
    try {
      const res = await axios.get<{ items: ESDExcellenceOpportunity[], computedBESS: number }>(
        `/api/projects/${id}/excellence`,
        { withCredentials: true },
      )
      setExcellenceItems(res.data.items)
      setComputedBESS(res.data.computedBESS)
      const flags: Record<string, string> = {}
      for (const item of res.data.items) flags[item.id] = item.flag
      setLocalFlags(flags)
    } catch (err) { console.error('[loadExcellence]', err) }
  }, [id])

  useEffect(() => { loadExcellence() }, [loadExcellence])

  const startPolling = useCallback(() => {
    if (!id) return
    stopPolling()
    pollRef.current = setInterval(async () => {
      try {
        const res = await axios.get<{ status: string }>(
          `/api/projects/${id}/generation-status`,
          { withCredentials: true },
        )
        const status = toGenStatus(res.data.status)
        setGenStatus(status)
        if (status === 'complete' || status === 'error') {
          stopPolling()
          if (status === 'complete') {
            // Refresh credits so commentsGIW is up to date in the matrix
            axios
              .get<Project>(`/api/projects/${id}`, { withCredentials: true })
              .then((r) => setProject(r.data))
              .catch(console.error)
            loadExcellence()
          }
        }
      } catch {
        // ignore transient poll errors
      }
    }, 3000)
  }, [id, stopPolling, loadExcellence])

  useEffect(() => {
    if (!id) return
    axios
      .get<Project>(`/api/projects/${id}`, { withCredentials: true })
      .then((res) => {
        setProject(res.data)
        setNotifyEmail(res.data.notifyEmail ?? '')
        const status = toGenStatus(res.data.generationStatus)
        setGenStatus(status)
        if (status === 'running') startPolling()
      })
      .catch(() => setError('Failed to load project.'))
      .finally(() => setLoading(false))
  }, [id, startPolling])

  useEffect(() => {
    if (project?.name) {
      document.title = `ESD Review Portal — ${project.name} | GIW Environmental Solutions`
    }
    return () => { document.title = 'ESD Review Portal | GIW Environmental Solutions' }
  }, [project?.name])

  // Clean up poll on unmount
  useEffect(() => () => stopPolling(), [stopPolling])

  async function handleFlag(itemId: string, flag: string) {
    setFlagging((f) => ({ ...f, [itemId]: true }))
    try {
      await axios.patch(`/api/excellence/${itemId}/flag`, { flag, flaggedBy: 'GIW' }, { withCredentials: true })
      setLocalFlags((f) => ({ ...f, [itemId]: flag }))
    } catch { /* silent */ } finally {
      setFlagging((f) => { const n = { ...f }; delete n[itemId]; return n })
    }
  }

  async function handleDeleteExcellence(itemId: string) {
    if (!window.confirm('Remove this opportunity from the review?')) return
    try {
      await axios.delete(`/api/excellence/${itemId}`, { withCredentials: true })
      setExcellenceItems((items) => items.filter((i) => i.id !== itemId))
    } catch { /* silent */ }
  }

  async function handleDeleteAllExcellence() {
    if (!id) return
    if (!window.confirm('Remove all ESD excellence opportunities and innovation credits from this review?')) return
    try {
      await axios.delete(`/api/projects/${id}/excellence`, { withCredentials: true })
      setExcellenceItems([])
    } catch { /* silent */ }
  }

  async function handleExport() {
    if (!id) return
    setExportLoading(true)
    try {
      const res = await axios.post(
        `/api/projects/${id}/export`,
        {},
        { withCredentials: true, responseType: 'blob' },
      )
      const blob = new Blob([res.data as BlobPart])
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `${project?.name ?? 'project'}-review-matrix.xlsx`
      a.click()
      URL.revokeObjectURL(url)
    } catch {
      // stub — server returns 200 JSON for now
    } finally {
      setExportLoading(false)
    }
  }

  const [changelogLoading, setChangelogLoading] = useState(false)
  const [reportLoading, setReportLoading] = useState(false)
  const [reportModalOpen, setReportModalOpen] = useState(false)
  const [reportGiwref, setReportGiwref] = useState('')
  const [reportClient, setReportClient] = useState('')
  const [reportArchitect, setReportArchitect] = useState('')

  async function handleChangelog() {
    if (!id) return
    setChangelogLoading(true)
    try {
      const res = await axios.get(
        `/api/projects/${id}/changelog`,
        { withCredentials: true, responseType: 'blob' },
      )
      const blob = new Blob([res.data as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `Changelog-${project?.name ?? 'project'}-Rev${project?.revision ?? ''}.docx`
      a.click()
      URL.revokeObjectURL(url)
    } catch {
      alert('Changelog not available — this may be the first revision.')
    } finally {
      setChangelogLoading(false)
    }
  }

  function handleGenerateReport() {
    setReportGiwref('GIW')
    setReportClient(project?.client ?? '')
    setReportArchitect(project?.architect ?? '')
    setReportModalOpen(true)
  }

  async function submitReport() {
    if (!id) return
    setReportModalOpen(false)
    setReportLoading(true)
    try {
      const res = await axios.post(
        `/api/projects/${id}/report`,
        { giwref: reportGiwref.trim(), client: reportClient.trim(), architect: reportArchitect.trim() },
        { withCredentials: true, responseType: 'blob' },
      )
      const blob = new Blob([res.data as BlobPart], { type: 'application/zip' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `SMP-Report-${project?.name ?? 'project'}-Rev${project?.revision ?? 'A'}.zip`
      a.click()
      URL.revokeObjectURL(url)
    } catch (err) {
      if (axios.isAxiosError(err) && err.response) {
        const text = await (err.response.data as Blob).text()
        try {
          const parsed = JSON.parse(text) as { error?: string }
          alert(parsed.error ?? 'Failed to generate report.')
        } catch {
          alert('Failed to generate report.')
        }
      } else {
        alert('Failed to generate report. Please try again.')
      }
    } finally {
      setReportLoading(false)
    }
  }

  async function handleRegenerate() {
    if (!id) return
    try {
      await axios.post(`/api/projects/${id}/generate`, {}, { withCredentials: true })
      setGenStatus('running')
      startPolling()
    } catch {
      setGenStatus('error')
    }
  }

  async function handleInvite() {
    if (!id || !inviteEmail.trim() || !inviteDiscipline.trim()) return
    setInviting(true)
    setInviteError('')
    setInviteSuccess(false)
    setInviteWarning('')
    try {
      const { data } = await axios.post<{ ok: boolean; reviewerId: string; emailWarning?: string }>(`/api/projects/${id}/invite`, {
        email: inviteEmail.trim(),
        discipline: inviteDiscipline.trim(),
        name: inviteName.trim() || undefined,
      }, { withCredentials: true })
      if (data.emailWarning) {
        setInviteWarning(data.emailWarning)
      } else {
        setInviteSuccess(true)
        setTimeout(() => setInviteSuccess(false), 5000)
      }
      setInviteEmail('')
      setInviteName('')
      setInviteDiscipline('')
      // Refresh reviewers list
      const res = await axios.get<Project>(`/api/projects/${id}`, { withCredentials: true })
      setProject(res.data)
    } catch {
      setInviteError('Failed to send invite. Check the email address and try again.')
    } finally {
      setInviting(false)
    }
  }

  async function handleSaveNotifyEmail() {
    if (!id) return
    try {
      await axios.patch(`/api/projects/${id}`, { notifyEmail: notifyEmail.trim() || null }, { withCredentials: true })
      setNotifyEmailSaved(true)
      setTimeout(() => setNotifyEmailSaved(false), 2000)
    } catch { /* silent */ }
  }

  if (loading) {
    return (
      <div className="min-h-screen" style={{ backgroundColor: '#FFFFFF' }}>
        <NavBar />
        <div style={{ backgroundColor: '#00602B', padding: '24px 32px' }}>
          <div className="max-w-6xl mx-auto flex items-start justify-between gap-6 flex-wrap">
            <div className="flex flex-col gap-2">
              <div className="skeleton-shimmer" style={{ height: 24, width: 280 }} />
              <div className="skeleton-shimmer" style={{ height: 14, width: 180 }} />
            </div>
            <div className="skeleton-shimmer" style={{ height: 60, width: 80, borderRadius: '4px' }} />
          </div>
        </div>
        <main className="max-w-6xl mx-auto px-6 py-8 space-y-5">
          <div className="giw-card flex flex-col gap-3">
            <div className="skeleton-shimmer" style={{ height: 12, width: 100 }} />
            <div className="skeleton-shimmer" style={{ height: 36, borderRadius: '2px' }} />
          </div>
          <div className="giw-card flex flex-col gap-3">
            <div className="skeleton-shimmer" style={{ height: 12, width: 80 }} />
            {Array.from({ length: 3 }).map((_, i) => (
              <div key={i} className="skeleton-shimmer" style={{ height: 16 }} />
            ))}
          </div>
        </main>
      </div>
    )
  }

  if (error || !project) {
    return (
      <div className="min-h-screen" style={{ backgroundColor: '#FFFFFF' }}>
        <NavBar />
        <div className="flex flex-col items-center justify-center py-32 gap-4">
          <p style={{ fontFamily: 'Open Sans, sans-serif', color: '#B94040' }}>
            {error || 'Project not found.'}
          </p>
          <button className="btn-secondary" onClick={() => navigate('/admin')}>
            Back to Projects
          </button>
        </div>
      </div>
    )
  }

  const credits = project.credits ?? []
  const drawingItems = project.drawingItems ?? []
  const reviewers = (project as Project & { reviewers?: Reviewer[] }).reviewers ?? []

  const DISCIPLINES = ['Architect', 'Services Engineer', 'Civil Engineer', 'Landscape Architect', 'Developer', 'Waste Consultant', 'ESD Consultant', 'Other']

  return (
    <div className="min-h-screen" style={{ backgroundColor: '#FFFFFF' }}>
      <NavBar />

      {/* Project header band */}
      <div style={{ backgroundColor: '#00602B', padding: '24px 32px' }}>
        <div className="max-w-6xl mx-auto flex items-start justify-between gap-6 flex-wrap">
          <div>
            <h1
              className="font-semibold"
              style={{
                fontFamily: 'Montserrat, sans-serif',
                fontSize: '24px',
                color: '#FFFFFF',
                marginBottom: '4px',
              }}
            >
              {project.name}
            </h1>
            {project.address && (
              <p
                style={{
                  fontFamily: 'Open Sans, sans-serif',
                  fontSize: '14px',
                  color: 'rgba(255,255,255,0.7)',
                  marginBottom: '6px',
                }}
              >
                {project.address}
              </p>
            )}
            <p
              style={{
                fontFamily: 'Open Sans, sans-serif',
                fontSize: '12px',
                color: 'rgba(255,255,255,0.6)',
              }}
            >
              {project.date
                ? new Date(project.date).toLocaleDateString('en-AU', {
                    day: 'numeric',
                    month: 'long',
                    year: 'numeric',
                  })
                : 'No date'}
              {project.revision ? ` · Rev ${project.revision}` : ''}
              {project.projectId ? ` · ${project.projectId}` : ''}
            </p>

            {/* Revision switcher */}
            {(project.revisionFamily ?? []).length > 1 && (
              <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', marginTop: '10px' }}>
                {(project.revisionFamily as RevisionSummary[]).map((r) => {
                  const isCurrent = r.id === project.id
                  return (
                    <button
                      key={r.id}
                      onClick={() => !isCurrent && navigate(`/admin/projects/${r.id}`)}
                      style={{
                        fontFamily: 'Montserrat, sans-serif', fontSize: '11px', fontWeight: 600,
                        padding: '3px 10px', borderRadius: '3px', cursor: isCurrent ? 'default' : 'pointer',
                        border: '1px solid rgba(255,255,255,0.6)',
                        backgroundColor: isCurrent ? 'rgba(255,255,255,0.25)' : 'transparent',
                        color: '#fff',
                      }}
                      title={`Rev ${r.revision ?? '?'} — ${new Date(r.createdAt).toLocaleDateString('en-AU')}${r.bessScore != null ? ` · ${r.bessScore}%` : ''}`}
                    >
                      Rev {r.revision ?? '?'}
                    </button>
                  )
                })}
              </div>
            )}
          </div>

          {project.bessScore != null && (
            <div style={{ display: 'flex', gap: '10px', flexShrink: 0 }}>
              <div style={{ backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '4px', padding: '8px 20px', textAlign: 'center' }}>
                <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '28px', fontWeight: 600, color: '#FFFFFF', lineHeight: 1 }}>
                  {project.bessScore}%
                </p>
                <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '11px', color: 'rgba(255,255,255,0.8)', marginTop: '4px' }}>
                  Baseline BESS Score
                </p>
              </div>
              {interactiveBESS != null && (
                <div style={{ backgroundColor: '#C8E6D4', borderRadius: '4px', padding: '8px 20px', textAlign: 'center' }}>
                  <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '28px', fontWeight: 600, color: '#004D22', lineHeight: 1 }}>
                    {interactiveBESS}%
                  </p>
                  <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '11px', color: '#004D22', marginTop: '4px' }}>
                    Improved BESS Score
                  </p>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <main className="max-w-6xl mx-auto px-6 py-8">
        {/* Two-column layout */}
        <div className="flex gap-6 flex-col lg:flex-row mb-8">
          {/* Main column (2/3) */}
          <div className="flex-1 min-w-0 space-y-6">
            {/* Invite Reviewer card */}
            <div className="giw-card">
              <p
                className="text-xs font-semibold mb-3 uppercase tracking-wide"
                style={{ fontFamily: 'Montserrat, sans-serif', color: '#C0C0C0' }}
              >
                Invite Reviewer
              </p>
              <div className="flex gap-2 flex-wrap mb-2">
                <input
                  type="text"
                  value={inviteName}
                  onChange={(e) => setInviteName(e.target.value)}
                  placeholder="Full name"
                  className="giw-input flex-1 text-sm"
                  style={{ fontFamily: 'Open Sans, sans-serif', minWidth: '160px' }}
                />
                <input
                  type="email"
                  value={inviteEmail}
                  onChange={(e) => setInviteEmail(e.target.value)}
                  placeholder="Email address"
                  className="giw-input flex-1 text-sm"
                  style={{ fontFamily: 'Open Sans, sans-serif', minWidth: '180px' }}
                />
              </div>
              <div className="flex gap-2 flex-wrap">
                <select
                  value={inviteDiscipline}
                  onChange={(e) => setInviteDiscipline(e.target.value)}
                  className="giw-input text-sm"
                  style={{ fontFamily: 'Open Sans, sans-serif' }}
                >
                  <option value="">Select discipline…</option>
                  {DISCIPLINES.map((d) => (
                    <option key={d} value={d}>{d}</option>
                  ))}
                </select>
                <button
                  className="btn-primary whitespace-nowrap"
                  onClick={handleInvite}
                  disabled={inviting || !inviteEmail.trim() || !inviteDiscipline.trim()}
                >
                  {inviting ? 'Sending…' : 'Send Invite'}
                </button>
              </div>
              {inviteSuccess && (
                <p className="mt-2 text-xs" style={{ fontFamily: 'Open Sans, sans-serif', color: '#004D22' }}>
                  Invite sent!
                </p>
              )}
              {inviteWarning && (
                <p className="mt-2 text-xs" style={{ fontFamily: 'Open Sans, sans-serif', color: '#B35900' }}>
                  {inviteWarning}
                </p>
              )}
              {inviteError && (
                <p className="mt-2 text-xs" style={{ fontFamily: 'Open Sans, sans-serif', color: '#B94040' }}>
                  {inviteError}
                </p>
              )}
            </div>

            {/* Reviewer table */}
            <div className="giw-card">
              <p
                className="text-xs font-semibold mb-3 uppercase tracking-wide"
                style={{ fontFamily: 'Montserrat, sans-serif', color: '#C0C0C0' }}
              >
                Reviewers
              </p>
              {reviewers.length === 0 ? (
                <div className="text-center py-8">
                  <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '10px' }}>
                    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" opacity="0.3">
                      <circle cx="16" cy="10" r="5" stroke="#00602B" strokeWidth="2" />
                      <path d="M6 28c0-5.523 4.477-10 10-10s10 4.477 10 10" stroke="#00602B" strokeWidth="2" strokeLinecap="round" />
                    </svg>
                  </div>
                  <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#C0C0C0' }}>
                    No reviewers have accessed this project yet
                  </p>
                </div>
              ) : (
                <table className="w-full text-sm">
                  <thead>
                    <tr
                      style={{
                        borderBottom: '1px solid #C0C0C0',
                        fontFamily: 'Montserrat, sans-serif',
                        fontSize: '12px',
                        color: '#C0C0C0',
                      }}
                    >
                      <th className="text-left py-2 pr-4 font-medium">Name</th>
                      <th className="text-left py-2 pr-4 font-medium">Email</th>
                      <th className="text-left py-2 pr-4 font-medium">Discipline</th>
                      <th className="text-left py-2 pr-4 font-medium">Submitted</th>
                      <th className="text-left py-2 pr-4 font-medium">Timestamp</th>
                      <th className="text-left py-2 font-medium">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {reviewers.map((r) => (
                      <tr
                        key={r.id}
                        style={{
                          borderBottom: '1px solid #C0C0C0',
                          fontFamily: 'Open Sans, sans-serif',
                          color: '#2C2C2C',
                          backgroundColor: '#FFFFFF',
                        }}
                      >
                        <td className="py-2 pr-4">{r.name ?? '—'}</td>
                        <td className="py-2 pr-4">{r.email}</td>
                        <td className="py-2 pr-4">{r.discipline}</td>
                        <td className="py-2 pr-4">
                          {r.hasSubmitted ? (
                            <span style={{ color: '#004D22', fontWeight: 600 }}>Yes</span>
                          ) : (
                            <span style={{ color: '#C0C0C0' }}>No</span>
                          )}
                        </td>
                        <td className="py-2 pr-4">
                          {r.submittedAt
                            ? new Date(r.submittedAt).toLocaleString('en-AU')
                            : '—'}
                        </td>
                        <td className="py-2">
                          <button
                            className="btn-secondary"
                            style={{ fontSize: '11px', padding: '2px 10px' }}
                            disabled={inviting}
                            onClick={async () => {
                              if (!id || !r.email || !r.discipline) return
                              setInviting(true)
                              setInviteError('')
                              setInviteSuccess(false)
                              setInviteWarning('')
                              try {
                                const { data } = await axios.post<{ ok: boolean; emailWarning?: string }>(`/api/projects/${id}/invite`, {
                                  email: r.email,
                                  discipline: r.discipline,
                                  ...(r.name ? { name: r.name } : {}),
                                }, { withCredentials: true })
                                if (data.emailWarning) {
                                  setInviteWarning(data.emailWarning)
                                } else {
                                  setInviteSuccess(true)
                                  setTimeout(() => setInviteSuccess(false), 5000)
                                }
                              } catch {
                                setInviteError('Failed to resend invite.')
                              } finally {
                                setInviting(false)
                              }
                            }}
                          >
                            Resend Invite
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>

          {/* Sidebar (1/3) */}
          <div className="lg:w-72 flex-shrink-0 space-y-4">
            <div className="giw-card space-y-3">
              <button
                className="btn-primary w-full"
                onClick={handleExport}
                disabled={exportLoading}
              >
                {exportLoading ? 'Exporting…' : 'Export Excel'}
              </button>

              {project?.parentProjectId && (
                <button
                  className="btn-secondary w-full"
                  onClick={handleChangelog}
                  disabled={changelogLoading}
                >
                  {changelogLoading ? 'Generating…' : 'Download Changelog'}
                </button>
              )}

              <button
                className="btn-secondary w-full"
                onClick={handleRegenerate}
                disabled={genStatus === 'running'}
              >
                {genStatus === 'running' ? 'Regenerating…' : 'Regenerate Comments'}
              </button>

              <button
                className="btn-primary w-full"
                onClick={handleGenerateReport}
                disabled={reportLoading}
              >
                {reportLoading ? 'Generating…' : 'Generate Report'}
              </button>

              {/* Generation status indicator */}
              <div className="flex items-center gap-2 pt-1">
                {genStatus === 'idle' && (
                  <span
                    className="text-xs"
                    style={{ fontFamily: 'Open Sans, sans-serif', color: '#C0C0C0' }}
                  >
                    Comments not yet generated
                  </span>
                )}
                {genStatus === 'running' && (
                  <>
                    <span
                      className="w-2 h-2 rounded-full animate-pulse flex-shrink-0"
                      style={{ backgroundColor: '#00602B' }}
                    />
                    <span
                      className="text-xs"
                      style={{ fontFamily: 'Open Sans, sans-serif', color: '#00602B' }}
                    >
                      Generating…
                    </span>
                  </>
                )}
                {genStatus === 'complete' && (
                  <>
                    <svg
                      width="14"
                      height="14"
                      viewBox="0 0 14 14"
                      fill="none"
                      className="flex-shrink-0"
                    >
                      <circle cx="7" cy="7" r="7" fill="#004D22" />
                      <path
                        d="M4 7l2 2 4-4"
                        stroke="white"
                        strokeWidth="1.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                    <span
                      className="text-xs"
                      style={{ fontFamily: 'Open Sans, sans-serif', color: '#004D22' }}
                    >
                      Comments generated
                    </span>
                  </>
                )}
                {genStatus === 'error' && (
                  <>
                    <svg
                      width="14"
                      height="14"
                      viewBox="0 0 14 14"
                      fill="none"
                      className="flex-shrink-0"
                    >
                      <circle cx="7" cy="7" r="7" fill="#B94040" />
                      <path
                        d="M5 5l4 4M9 5l-4 4"
                        stroke="white"
                        strokeWidth="1.5"
                        strokeLinecap="round"
                      />
                    </svg>
                    <span
                      className="text-xs"
                      style={{ fontFamily: 'Open Sans, sans-serif', color: '#B94040' }}
                    >
                      Generation failed — try again
                    </span>
                  </>
                )}
              </div>

              {/* Notify email */}
              <div>
                <label style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#8C8C8C', display: 'block', marginBottom: '4px' }}>
                  Notify email (on submission)
                </label>
                <div style={{ display: 'flex', gap: '6px' }}>
                  <input
                    type="email"
                    value={notifyEmail}
                    onChange={(e) => setNotifyEmail(e.target.value)}
                    onBlur={handleSaveNotifyEmail}
                    placeholder="email@giw.com.au"
                    className="giw-input flex-1 text-sm"
                  />
                  {notifyEmailSaved && (
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0, marginTop: '8px' }}>
                      <circle cx="8" cy="8" r="8" fill="#004D22" />
                      <path d="M5 8l2 2 4-4" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Tabs */}
        <div style={{ borderBottom: '1px solid #C0C0C0', marginBottom: '24px' }}>
          {(['matrix', 'drawings'] as Tab[]).map((tab) => (
            <button
              key={tab}
              onClick={() => setActiveTab(tab)}
              style={{
                fontFamily: 'Montserrat, sans-serif',
                fontSize: '14px',
                fontWeight: activeTab === tab ? 600 : 400,
                color: activeTab === tab ? '#2C2C2C' : '#C0C0C0',
                padding: '10px 20px',
                marginBottom: '-1px',
                background: 'none',
                border: 'none',
                borderBottom: activeTab === tab ? '2px solid #00602B' : '2px solid transparent',
                cursor: 'pointer',
              }}
            >
              {tab === 'matrix' ? 'Review Matrix' : 'Drawing Requirements'}
            </button>
          ))}
        </div>

        {/* Tab content */}
        {activeTab === 'matrix' && (
          <>
            <ReviewMatrix credits={credits} gdft={project?.gdft ?? false} genStatus={genStatus} />
            <AdminExcellenceSection
              items={excellenceItems}
              localFlags={localFlags}
              flagging={flagging}
              onFlag={handleFlag}
              onDelete={handleDeleteExcellence}
              gdft={project?.gdft ?? false}
            />
          </>
        )}
        {activeTab === 'drawings' && (
          <DrawingRequirementsTab items={drawingItems} />
        )}
      </main>

      {/* Report details modal */}
      {reportModalOpen && (
        <div
          style={{
            position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.45)',
            display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 50,
          }}
          onClick={(e) => { if (e.target === e.currentTarget) setReportModalOpen(false) }}
        >
          <div className="giw-card" style={{ width: '400px', maxWidth: '90vw' }}>
            <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '16px', fontWeight: 600, color: '#2C2C2C', marginBottom: '20px' }}>
              Generate SMP Report
            </p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
              <div>
                <label style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#8C8C8C', display: 'block', marginBottom: '4px' }}>
                  GIW Reference
                </label>
                <input
                  type="text"
                  value={reportGiwref}
                  onChange={(e) => setReportGiwref(e.target.value)}
                  placeholder="e.g. GIW-2024-001"
                  className="giw-input w-full"
                  style={{ fontFamily: 'Open Sans, sans-serif' }}
                  autoFocus
                />
              </div>
              <div>
                <label style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#8C8C8C', display: 'block', marginBottom: '4px' }}>
                  Client
                </label>
                <input
                  type="text"
                  value={reportClient}
                  onChange={(e) => setReportClient(e.target.value)}
                  placeholder="Client name"
                  className="giw-input w-full"
                  style={{ fontFamily: 'Open Sans, sans-serif' }}
                />
              </div>
              <div>
                <label style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#8C8C8C', display: 'block', marginBottom: '4px' }}>
                  Architect
                </label>
                <input
                  type="text"
                  value={reportArchitect}
                  onChange={(e) => setReportArchitect(e.target.value)}
                  placeholder="Architect name"
                  className="giw-input w-full"
                  style={{ fontFamily: 'Open Sans, sans-serif' }}
                  onKeyDown={(e) => { if (e.key === 'Enter') submitReport() }}
                />
              </div>
            </div>
            <div style={{ display: 'flex', gap: '8px', marginTop: '20px' }}>
              <button className="btn-primary" style={{ flex: 1 }} onClick={submitReport}>
                Generate
              </button>
              <button className="btn-secondary" style={{ flex: 1 }} onClick={() => setReportModalOpen(false)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

/* ── Auto-resizing textarea ── */

function AutoTextarea({
  value,
  onChange,
  onBlur,
  hasError,
}: {
  value: string
  onChange: (v: string) => void
  onBlur: () => void
  hasError: boolean
}) {
  const ref = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    const el = ref.current
    if (!el) return
    el.style.height = 'auto'
    el.style.height = `${el.scrollHeight}px`
  }, [value])

  return (
    <textarea
      ref={ref}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      onBlur={onBlur}
      rows={1}
      className="no-scrollbar"
      style={{
        width: '100%',
        resize: 'none',
        overflow: 'hidden',
        border: `1px solid ${hasError ? '#B94040' : '#C0C0C0'}`,
        borderRadius: '2px',
        padding: '4px 6px',
        fontSize: '12px',
        fontFamily: 'Open Sans, sans-serif',
        color: '#2C2C2C',
        backgroundColor: '#fff',
        lineHeight: '1.5',
        minHeight: '28px',
        boxSizing: 'border-box',
      }}
    />
  )
}

/* ── Review Matrix ── */

const STATUS_LABELS: Record<string, string> = {
  Y: 'Achieved',
  N: 'Not Achieved',
  ScopedOut: 'Scoped Out',
}

const STATUS_COLOURS: Record<string, { bg: string; text: string }> = {
  Y: { bg: '#C6EFCE', text: '#2D6A2D' },
  N: { bg: '#FCE4D6', text: '#9C3D1E' },
  ScopedOut: { bg: '#C0C0C0', text: '#555555' },
}

function creditStatusStyle(status: string) {
  return STATUS_COLOURS[status] ?? { bg: '#FFFFFF', text: '#2C2C2C' }
}

function getCategoryOrder(creditId: string, category = ''): number {
  const id = creditId.toLowerCase()
  const cat = category.toLowerCase()
  if (id.startsWith('management')  || cat.startsWith('management'))              return 1
  if (id.startsWith('iwm')         || cat.includes('water'))                     return 2
  if (id.startsWith('oe')          || cat.includes('operational energy'))        return 3
  if (id.startsWith('ieq')         || cat.includes('indoor environmental'))      return 4
  if (id.startsWith('transport')   || cat.includes('transport'))                 return 5
  if (id.startsWith('waste')       || cat.includes('waste'))                     return 6
  if (id.startsWith('urban')       || cat.includes('urban'))                     return 7
  if (id.startsWith('innovation')  || cat.includes('innovation'))                return 8
  return 99
}

function groupByCategory(credits: Credit[]): { category: string; order: number; items: Credit[] }[] {
  const map = new Map<string, { category: string; order: number; items: Credit[] }>()
  for (const c of credits) {
    const order = getCategoryOrder(c.creditId, c.category)
    if (!map.has(c.category)) {
      map.set(c.category, { category: c.category, order, items: [] })
    }
    map.get(c.category)!.items.push(c)
  }
  return [...map.values()].sort((a, b) => a.order - b.order)
}

function ReviewMatrix({ credits, gdft, genStatus }: { credits: Credit[]; gdft: boolean; genStatus?: GenStatus }) {
  const [giwComments, setGiwComments] = useState<Record<string, string>>(() => {
    const init: Record<string, string> = {}
    for (const c of credits) init[c.id] = c.commentsGIW ?? ''
    return init
  })
  const [saved, setSaved] = useState<Set<string>>(new Set())
  const [errors, setErrors] = useState<Set<string>>(new Set())

  // Sync textarea values when credits prop updates (e.g. after generation completes)
  useEffect(() => {
    setGiwComments((prev) => {
      const updated = { ...prev }
      for (const c of credits) updated[c.id] = c.commentsGIW ?? ''
      return updated
    })
  }, [credits])

  async function saveGIWComment(creditId: string) {
    const text = giwComments[creditId] ?? ''
    try {
      await axios.patch(
        `/api/credits/${creditId}/giw-comment`,
        { commentText: text },
        { withCredentials: true },
      )
      setSaved((prev) => new Set([...prev, creditId]))
      setErrors((prev) => { const s = new Set(prev); s.delete(creditId); return s })
      setTimeout(() => setSaved((prev) => { const s = new Set(prev); s.delete(creditId); return s }), 2000)
    } catch {
      setErrors((prev) => new Set([...prev, creditId]))
    }
  }

  if (credits.length === 0) {
    if (genStatus === 'running') {
      return (
        <div className="space-y-4">
          {[1, 2, 3].map((i) => (
            <div key={i} className="giw-card p-0 overflow-hidden">
              <div style={{ backgroundColor: '#C8E6D4', padding: '10px 20px', borderBottom: '1px solid #C0C0C0' }}>
                <div className="skeleton-shimmer" style={{ height: 14, width: 140, borderRadius: 2 }} />
              </div>
              <div style={{ padding: '12px 20px', display: 'flex', flexDirection: 'column', gap: 10 }}>
                {[1, 2, 3].map((j) => (
                  <div key={j} className="skeleton-shimmer" style={{ height: 12, width: `${70 + j * 8}%`, borderRadius: 2 }} />
                ))}
              </div>
            </div>
          ))}
          <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#8C8C8C', textAlign: 'center', paddingTop: '8px' }}>
            Generating review matrix…
          </p>
        </div>
      )
    }
    return (
      <div className="text-center py-16">
        <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '12px' }}>
          <svg width="40" height="40" viewBox="0 0 40 40" fill="none" opacity="0.3">
            <rect x="4" y="8" width="32" height="26" rx="2" stroke="#00602B" strokeWidth="2" />
            <path d="M4 14h32" stroke="#00602B" strokeWidth="2" />
            <path d="M12 20h16M12 26h10" stroke="#00602B" strokeWidth="2" strokeLinecap="round" />
          </svg>
        </div>
        <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#C0C0C0' }}>
          Upload a BESS PDF to generate the review matrix
        </p>
      </div>
    )
  }

  const groups = groupByCategory(credits)

  return (
    <div className="giw-card p-0 overflow-hidden">
      <table className="text-sm" style={{ tableLayout: 'fixed', width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr
            style={{
              borderBottom: '1px solid #C0C0C0',
              backgroundColor: '#FFFFFF',
              fontFamily: 'Montserrat, sans-serif',
              fontSize: '11px',
              color: '#C0C0C0',
            }}
          >
            <th className="text-left py-2 px-5 font-medium" style={{ width: '70px' }}>ID</th>
            <th className="text-left py-2 pr-4 font-medium" style={{ width: '120px' }}>Credit</th>
            <th className="text-left py-2 pr-4 font-medium" style={{ width: '160px' }}>Requirement</th>
            <th className="text-left py-2 pr-4 font-medium" style={{ width: '112px' }}>Status</th>
            <th className="text-left py-2 pr-4 font-medium" style={{ width: '50px' }}>Score</th>
            <th className="text-left py-2 pr-4 font-medium" style={{ width: '55px' }}>Weight</th>
            <th className="text-left py-2 pr-4 font-medium" style={{ width: '65px' }}>Mandatory</th>
            <th className="text-left py-2 pr-4 font-medium" style={{ width: '110px' }}>Responsible Party</th>
            <th className="text-left py-2 pl-2 pr-5 font-medium">Comments GIW</th>
          </tr>
        </thead>
        <tbody>
          {groups.map((group) => (
            <>
              <tr key={`cat-${group.order}`}>
                <td
                  colSpan={9}
                  style={{
                    backgroundColor: '#C8E6D4',
                    padding: '10px 20px',
                    borderBottom: '1px solid #C0C0C0',
                    borderTop: '1px solid #C0C0C0',
                    fontFamily: 'Montserrat, sans-serif',
                    fontSize: '13px',
                    fontWeight: 600,
                    color: '#004D22',
                  }}
                >
                  {group.category}
                </td>
              </tr>
              {group.items.map((credit) => {
                const style = creditStatusStyle(credit.creditStatus)
                return (
                  <tr
                    key={credit.id}
                    style={{
                      borderBottom: '1px solid #C0C0C0',
                      fontFamily: 'Open Sans, sans-serif',
                      color: '#2C2C2C',
                      backgroundColor: '#FFFFFF',
                    }}
                  >
                    <td
                      className="py-2 px-5 font-medium"
                      style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '12px' }}
                    >
                      {credit.creditId}
                    </td>
                    <td className="py-2 pr-4 text-sm" style={{ overflow: 'hidden', wordBreak: 'break-word' }}>{credit.creditName}</td>
                    <td className="py-2 pr-4 text-sm" style={{ color: '#555', overflow: 'hidden', wordBreak: 'break-word' }}>
                      {credit.creditRequirement ?? '—'}
                      <CreditInfoButton creditId={credit.creditId} />
                    </td>
                    <td className="py-2 pr-4">
                      <span
                        style={{
                          backgroundColor: style.bg,
                          color: style.text,
                          padding: '2px 8px',
                          borderRadius: '3px',
                          fontSize: '11px',
                          fontFamily: 'Montserrat, sans-serif',
                          fontWeight: 600,
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {STATUS_LABELS[credit.creditStatus] ?? credit.creditStatus}
                      </span>
                    </td>
                    <td className="py-2 pr-4 text-sm">
                      {credit.creditScore != null ? credit.creditScore : '—'}
                    </td>
                    <td className="py-2 pr-4 text-sm">
                      {credit.creditWeight != null ? credit.creditWeight : '—'}
                    </td>
                    <td className="py-2 pr-4 text-sm" style={{ textAlign: 'center' }}>
                      {(credit.mandatory || (gdft && GDFT_CREDIT_NAMES.includes(credit.creditName))) ? (
                        <span style={{ color: '#004D22', fontWeight: 600 }}>Yes</span>
                      ) : (
                        <span style={{ color: '#C0C0C0' }}>No</span>
                      )}
                      {gdft && GDFT_CREDIT_NAMES.includes(credit.creditName) && (
                        <span style={{ display: 'block', marginTop: '4px', padding: '3px 8px', borderRadius: '3px', backgroundColor: '#FFF0E0', fontFamily: 'Montserrat, sans-serif', fontWeight: 700, fontSize: '13px', color: '#B84D00', textAlign: 'center' }}>GDFT</span>
                      )}
                    </td>
                    <td className="py-2 pr-4 text-sm" style={{ color: '#555', overflow: 'hidden', wordBreak: 'break-word' }}>
                      {credit.responsibleParty ?? '—'}
                    </td>
                    <td className="py-2 pl-2 pr-5">
                      {credit.creditStatus === 'ScopedOut' && !giwComments[credit.id]?.trim() ? (
                        <span style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#8C8C8C', fontStyle: 'italic' }}>
                          Not targeted
                        </span>
                      ) : (
                        <div style={{ position: 'relative' }}>
                          <AutoTextarea
                            value={giwComments[credit.id] ?? ''}
                            onChange={(v) =>
                              setGiwComments((prev) => ({ ...prev, [credit.id]: v }))
                            }
                            onBlur={() => saveGIWComment(credit.id)}
                            hasError={errors.has(credit.id)}
                          />
                          {saved.has(credit.id) && (
                            <svg width="14" height="14" viewBox="0 0 14 14" fill="none"
                              style={{ position: 'absolute', bottom: 6, right: 6 }}>
                              <circle cx="7" cy="7" r="7" fill="#004D22" />
                              <path d="M4 7l2 2 4-4" stroke="white" strokeWidth="1.4"
                                strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          )}
                        </div>
                      )}
                    </td>
                  </tr>
                )
              })}
            </>
          ))}
        </tbody>
      </table>
    </div>
  )
}

/* ── Drawing Requirements ── */

const DRAWING_STATUS_COLOURS: Record<string, { bg: string; text: string }> = {
  NotStarted: { bg: '#C0C0C0', text: '#555555' },
  InProgress: { bg: '#FCE4D6', text: '#9C3D1E' },
  Complete: { bg: '#C6EFCE', text: '#2D6A2D' },
}

const DRAWING_STATUS_OPTIONS = [
  { value: 'NotStarted', label: 'Not Started' },
  { value: 'InProgress', label: 'In Progress' },
  { value: 'Complete',   label: 'Complete' },
]

function DrawingRequirementsTab({ items }: { items: DrawingRequirement[] }) {
  const [statuses, setStatuses] = useState<Record<string, string>>({})

  useEffect(() => {
    setStatuses(Object.fromEntries(items.map((i) => [i.id, i.status])))
  }, [items])

  async function handleStatusChange(id: string, status: string) {
    setStatuses((s) => ({ ...s, [id]: status }))
    try {
      await axios.patch(`/api/drawing-requirements/${id}`, { status })
    } catch {
      // revert on failure
      setStatuses((s) => ({ ...s, [id]: items.find((i) => i.id === id)?.status ?? s[id] }))
    }
  }

  if (items.length === 0) {
    return (
      <div className="text-center py-16">
        <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '12px' }}>
          <svg width="40" height="40" viewBox="0 0 40 40" fill="none" opacity="0.3">
            <rect x="4" y="4" width="32" height="32" rx="2" stroke="#00602B" strokeWidth="2" />
            <path d="M12 14h16M12 20h10M12 26h8" stroke="#00602B" strokeWidth="2" strokeLinecap="round" />
          </svg>
        </div>
        <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#C0C0C0' }}>
          No drawing annotations required
        </p>
      </div>
    )
  }

  return (
    <div className="giw-card p-0 overflow-hidden">
      <table className="w-full text-sm" style={{ tableLayout: 'fixed' }}>
        <thead>
          <tr
            style={{
              borderBottom: '1px solid #C0C0C0',
              backgroundColor: '#FFFFFF',
              fontFamily: 'Montserrat, sans-serif',
              fontSize: '11px',
              color: '#C0C0C0',
            }}
          >
            <th className="text-left py-2 px-5 font-medium w-28">Credit Ref</th>
            <th className="text-left py-2 pr-4 font-medium w-32">Type</th>
            <th className="text-left py-2 pr-4 font-medium">Requirement</th>
            <th className="text-left py-2 pr-5 font-medium w-36">Status</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item) => {
            const currentStatus = statuses[item.id] ?? item.status
            const colours = DRAWING_STATUS_COLOURS[currentStatus] ?? { bg: '#FFFFFF', text: '#2C2C2C' }
            return (
              <tr
                key={item.id}
                style={{
                  borderBottom: '1px solid #C0C0C0',
                  fontFamily: 'Open Sans, sans-serif',
                  color: '#2C2C2C',
                }}
              >
                <td
                  className="py-2 px-5 font-medium"
                  style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '12px' }}
                >
                  {item.creditReference}
                </td>
                <td className="py-2 pr-4">{item.drawingType}</td>
                <td className="py-2 pr-4">{item.requirement}</td>
                <td className="py-2 pr-5">
                  <select
                    value={currentStatus}
                    onChange={(e) => handleStatusChange(item.id, e.target.value)}
                    style={{
                      backgroundColor: colours.bg,
                      color: colours.text,
                      padding: '2px 6px',
                      borderRadius: '3px',
                      fontSize: '11px',
                      fontFamily: 'Montserrat, sans-serif',
                      fontWeight: 600,
                      border: 'none',
                      cursor: 'pointer',
                      outline: 'none',
                    }}
                  >
                    {DRAWING_STATUS_OPTIONS.map((o) => (
                      <option key={o.value} value={o.value}>{o.label}</option>
                    ))}
                  </select>
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

/* ── ESD Excellence Opportunities (admin view) ── */

const FLAG_OPTIONS = ['Yes', 'No', 'Maybe'] as const

interface AdminExcellenceSectionProps {
  items: ESDExcellenceOpportunity[]
  localFlags: Record<string, string>
  flagging: Record<string, boolean>
  onFlag: (id: string, flag: string) => void
  onDelete: (id: string) => void
  gdft: boolean
}

function AdminExcellenceSection({ items, localFlags, flagging, onFlag, onDelete, gdft }: AdminExcellenceSectionProps) {
  const [notesMap, setNotesMap] = useState<Record<string, string>>({})
  const [notesSaved, setNotesSaved] = useState<Set<string>>(new Set())
  const [notesErrors, setNotesErrors] = useState<Set<string>>(new Set())

  const [descMap, setDescMap] = useState<Record<string, string>>({})
  const [descSaved, setDescSaved] = useState<Set<string>>(new Set())
  const [descErrors, setDescErrors] = useState<Set<string>>(new Set())

  // Sync maps when items load or change (items starts as [] and arrives asynchronously)
  useEffect(() => {
    setNotesMap((prev) => {
      const m = { ...prev }
      for (const item of items) if (!(item.id in m)) m[item.id] = item.reviewerNotes ?? ''
      return m
    })
    setDescMap((prev) => {
      const m = { ...prev }
      for (const item of items) if (!(item.id in m)) m[item.id] = item.improvementDescription ?? ''
      return m
    })
  }, [items])

  async function saveNotes(id: string) {
    try {
      await axios.patch(`/api/excellence/${id}/notes`, { reviewerNotes: notesMap[id] ?? '' }, { withCredentials: true })
      setNotesSaved((s) => new Set([...s, id]))
      setNotesErrors((s) => { const n = new Set(s); n.delete(id); return n })
      setTimeout(() => setNotesSaved((s) => { const n = new Set(s); n.delete(id); return n }), 2000)
    } catch {
      setNotesErrors((s) => new Set([...s, id]))
    }
  }

  async function saveDesc(id: string) {
    try {
      await axios.patch(`/api/excellence/${id}/description`, { improvementDescription: descMap[id] ?? '' }, { withCredentials: true })
      setDescSaved((s) => new Set([...s, id]))
      setDescErrors((s) => { const n = new Set(s); n.delete(id); return n })
      setTimeout(() => setDescSaved((s) => { const n = new Set(s); n.delete(id); return n }), 2000)
    } catch {
      setDescErrors((s) => new Set([...s, id]))
    }
  }

  if (items.length === 0) return null

  const regularItems = [...items.filter(i => i.creditReference !== 'Innovation')]
    .sort((a, b) => {
      const oa = getCategoryOrder(a.creditReference)
      const ob = getCategoryOrder(b.creditReference)
      return oa !== ob ? oa - ob : a.creditReference.localeCompare(b.creditReference)
    })
    .filter((item, idx, arr) => arr.findIndex(x => x.creditReference === item.creditReference) === idx)

  const innovationItems = [...items.filter(i => i.creditReference === 'Innovation')]
    .sort((a, b) => a.creditName.localeCompare(b.creditName))
    .filter((item, idx, arr) => arr.findIndex(x => x.creditName === item.creditName) === idx)

  function cardProps(item: ESDExcellenceOpportunity) {
    return {
      item,
      flag: localFlags[item.id] ?? item.flag,
      isFlagging: !!flagging[item.id],
      desc: descMap[item.id] ?? '',
      onDescChange: (v: string) => setDescMap((m) => ({ ...m, [item.id]: v })),
      onDescBlur: () => saveDesc(item.id),
      descSaved: descSaved.has(item.id),
      descError: descErrors.has(item.id),
      notes: notesMap[item.id] ?? '',
      onNotesChange: (v: string) => setNotesMap((m) => ({ ...m, [item.id]: v })),
      onNotesBlur: () => saveNotes(item.id),
      notesSaved: notesSaved.has(item.id),
      notesError: notesErrors.has(item.id),
      onFlag: (f: string) => onFlag(item.id, f),
      onDelete: () => onDelete(item.id),
    }
  }

  const EXCELLENCE_COLS: [string, number][] = [
    ['Credit Ref', 90], ['Credit Name', 180], ['Current Score', 100], ['BESS Points', 100],
    ['Improvement Description', 260], ['Flag', 100], ['Comments', 220], ['', 60],
  ]

  function renderTable(rows: ESDExcellenceOpportunity[], mb?: string, hideCurrentScore?: boolean, showGdft?: boolean) {
    const cols = hideCurrentScore ? EXCELLENCE_COLS.filter(([label]) => label !== 'Current Score') : EXCELLENCE_COLS
    return (
      <div style={{ borderRadius: '4px', border: '1px solid #C0C0C0', marginBottom: mb }}>
        <table style={{ borderCollapse: 'collapse', tableLayout: 'fixed', width: '100%', backgroundColor: '#fff', fontSize: '13px', fontFamily: 'Open Sans, sans-serif' }}>
          <thead>
            <tr style={{ backgroundColor: '#00602B' }}>
              {cols.map(([label, w]) => (
                <th key={label} style={{
                  fontFamily: 'Montserrat, sans-serif', fontSize: '12px', fontWeight: 500,
                  color: '#fff', textTransform: 'uppercase', padding: '8px 12px',
                  textAlign: 'left', whiteSpace: 'nowrap', width: w,
                }}>
                  {label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((item) => (
              <AdminExcellenceRow
                key={item.id}
                {...cardProps(item)}
                hideCurrentScore={hideCurrentScore}
                gdftPrinciple={showGdft ? (INNOVATION_GDFT[item.creditName] ?? null) : null}
              />
            ))}
          </tbody>
        </table>
      </div>
    )
  }

  return (
    <div style={{ marginTop: '48px' }}>

      {regularItems.length > 0 && (
        <>
          <div style={{ marginBottom: '16px' }}>
            <h2 style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '20px', fontWeight: 600, color: '#2C2C2C', marginBottom: '6px' }}>
              ESD Excellence Opportunities
            </h2>
            <div style={{ height: '2px', width: '48px', backgroundColor: '#00602B', borderRadius: '1px' }} />
          </div>
          {renderTable(regularItems, '48px')}
        </>
      )}

      {innovationItems.length > 0 && (
        <>
          <div style={{ marginBottom: '16px' }}>
            <h2 style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '20px', fontWeight: 600, color: '#2C2C2C', marginBottom: '6px' }}>
              Innovation Credits
            </h2>
            <div style={{ height: '2px', width: '48px', backgroundColor: '#00602B', borderRadius: '1px' }} />
          </div>
          {renderTable(innovationItems, undefined, true, gdft)}
        </>
      )}
    </div>
  )
}

const excellenceTdBase: React.CSSProperties = {
  padding: '8px 12px',
  borderBottom: '1px solid #C0C0C0',
  verticalAlign: 'top',
  overflowWrap: 'break-word',
  wordBreak: 'break-word',
}

interface AdminExcellenceRowProps {
  item: ESDExcellenceOpportunity
  flag: string
  isFlagging: boolean
  desc: string
  onDescChange: (v: string) => void
  onDescBlur: () => void
  descSaved: boolean
  descError: boolean
  notes: string
  onNotesChange: (v: string) => void
  onNotesBlur: () => void
  notesSaved: boolean
  notesError: boolean
  onFlag: (flag: string) => void
  onDelete: () => void
  hideCurrentScore?: boolean
  gdftPrinciple?: string | null
}

function AdminExcellenceRow({ item, flag, isFlagging, desc, onDescChange, onDescBlur, descSaved, descError, notes, onNotesChange, onNotesBlur, notesSaved, notesError, onFlag, onDelete, hideCurrentScore, gdftPrinciple }: AdminExcellenceRowProps) {
  const [hovered, setHovered] = useState(false)
  return (
    <tr
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{ backgroundColor: hovered ? '#C8E6D4' : '#fff', transition: 'background-color 0.1s' }}
    >
      <td style={{ ...excellenceTdBase, minWidth: 90 }}>
        <span style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '13px', fontWeight: 500, color: '#2C2C2C' }}>
          {item.creditReference}
        </span>
      </td>

      <td style={{ ...excellenceTdBase, minWidth: 180, color: '#2C2C2C' }}>
        {item.creditName}
        {gdftPrinciple && (
          <span style={{ display: 'inline-block', marginLeft: '6px', padding: '2px 7px', borderRadius: '3px', backgroundColor: '#FFF0E0', fontFamily: 'Montserrat, sans-serif', fontWeight: 700, fontSize: '11px', color: '#B84D00', whiteSpace: 'nowrap' }}>
            GDFT: {gdftPrinciple}
          </span>
        )}
      </td>

      {!hideCurrentScore && (
        <td style={{ ...excellenceTdBase, minWidth: 100 }}>
          {item.currentScore != null ? (
            <span style={{ display: 'inline-block', backgroundColor: '#C8E6D4', color: '#004D22', fontFamily: 'Open Sans, sans-serif', fontSize: '12px', padding: '2px 10px', borderRadius: '12px' }}>
              {item.currentScore}%
            </span>
          ) : (
            <span style={{ color: '#C0C0C0', fontSize: '12px' }}>—</span>
          )}
        </td>
      )}

      <td style={{ ...excellenceTdBase, minWidth: 100 }}>
        {item.bessPoints ? (
          <span style={{ display: 'inline-block', backgroundColor: '#E8EDD8', color: '#4E5A2A', fontFamily: 'Open Sans, sans-serif', fontSize: '12px', padding: '2px 10px', borderRadius: '12px' }}>
            {Number(item.bessPoints)} BESS points
          </span>
        ) : item.additionalBessPoints != null ? (
          <span style={{ display: 'inline-block', backgroundColor: '#E8EDD8', color: '#4E5A2A', fontFamily: 'Open Sans, sans-serif', fontSize: '12px', padding: '2px 10px', borderRadius: '12px' }}>
            {item.additionalBessPoints} BESS {item.additionalBessPoints === 1 ? 'point' : 'points'}
          </span>
        ) : (
          <span style={{ color: '#C0C0C0', fontSize: '12px' }}>—</span>
        )}
      </td>

      <td style={{ ...excellenceTdBase, minWidth: 260 }}>
        <div style={{ position: 'relative' }}>
          <AutoTextarea value={desc} onChange={onDescChange} onBlur={onDescBlur} hasError={descError} />
          {descSaved && (
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none" style={{ position: 'absolute', bottom: 6, right: 6 }}>
              <circle cx="7" cy="7" r="7" fill="#004D22" />
              <path d="M4 7l2 2 4-4" stroke="white" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          )}
        </div>
      </td>

      <td style={{ ...excellenceTdBase, minWidth: 100 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
          {FLAG_OPTIONS.map((opt) => {
            const active = flag === opt
            return (
              <button
                key={opt}
                onClick={() => onFlag(opt)}
                disabled={isFlagging}
                style={{
                  fontFamily: 'Montserrat, sans-serif', fontSize: '11px', fontWeight: 500,
                  padding: '3px 10px', borderRadius: '2px', cursor: 'pointer',
                  border: '1px solid #00602B',
                  backgroundColor: active ? '#00602B' : 'transparent',
                  color: active ? '#fff' : '#00602B',
                  opacity: isFlagging ? 0.5 : 1,
                  transition: 'background-color 0.15s, color 0.15s',
                }}
              >
                {opt}
              </button>
            )
          })}
        </div>
      </td>

      <td style={{ ...excellenceTdBase, minWidth: 220 }}>
        <div style={{ position: 'relative' }}>
          <textarea
            value={notes}
            onChange={(e) => onNotesChange(e.target.value)}
            onBlur={onNotesBlur}
            rows={3}
            placeholder="Add a comment…"
            style={{
              width: '100%', resize: 'vertical',
              border: `1px solid ${notesError ? '#B94040' : '#C0C0C0'}`,
              borderRadius: '2px', padding: '6px 8px',
              fontSize: '12px', fontFamily: 'Open Sans, sans-serif',
              color: '#2C2C2C', backgroundColor: 'transparent', boxSizing: 'border-box',
            }}
          />
          {notesSaved && (
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none" style={{ position: 'absolute', bottom: 8, right: 8 }}>
              <circle cx="7" cy="7" r="7" fill="#004D22" />
              <path d="M4 7l2 2 4-4" stroke="white" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          )}
        </div>
      </td>

      <td style={{ ...excellenceTdBase, minWidth: 60, textAlign: 'center' }}>
        <button
          onClick={onDelete}
          title="Delete"
          style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '4px', color: '#B94040', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', borderRadius: '2px' }}
        >
          <svg width="15" height="15" viewBox="0 0 15 15" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M5.5 1.5h4a.5.5 0 0 1 0 1h-4a.5.5 0 0 1 0-1zM2 3.5h11a.5.5 0 0 1 0 1H12.5l-.9 8.1a1 1 0 0 1-1 .9H4.4a1 1 0 0 1-1-.9L2.5 4.5H1.5a.5.5 0 0 1 0-1H2zm1.51 1 .86 7.7a.1.1 0 0 0 .03.06V12h6.2l.86-7.5H3.51z" fill="currentColor"/>
          </svg>
        </button>
      </td>
    </tr>
  )
}
