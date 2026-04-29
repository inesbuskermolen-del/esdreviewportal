import {
  useEffect, useState, useRef, useCallback, useMemo, Fragment,
} from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import axios from 'axios'
import NavBar from '@/components/NavBar'
import type { Credit, ESDExcellenceOpportunity, ReviewerSession } from '@/types'

/* ── localStorage session ── */

const SESSION_KEY = 'giw_reviewer_session'
const SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000

function getSession(token: string): ReviewerSession | null {
  try {
    const raw = localStorage.getItem(SESSION_KEY)
    if (!raw) return null
    const s = JSON.parse(raw) as ReviewerSession
    if (s.reviewLinkToken !== token) return null
    if (Date.now() - new Date(s.createdAt).getTime() > SESSION_TTL_MS) {
      localStorage.removeItem(SESSION_KEY)
      return null
    }
    return s
  } catch { return null }
}

/* ── Helpers ── */

function getCategoryOrder(creditId: string, category = ''): number {
  const id = creditId.toLowerCase()
  const cat = category.toLowerCase()
  if (id.startsWith('management')  || cat.startsWith('management'))         return 1
  if (id.startsWith('iwm')         || cat.includes('water'))                return 2
  if (id.startsWith('oe')          || cat.includes('operational energy'))   return 3
  if (id.startsWith('ieq')         || cat.includes('indoor environmental')) return 4
  if (id.startsWith('transport')   || cat.includes('transport'))            return 5
  if (id.startsWith('waste')       || cat.includes('waste'))                return 6
  if (id.startsWith('urban')       || cat.includes('urban'))                return 7
  if (id.startsWith('innovation')  || cat.includes('innovation'))           return 8
  return 99
}

/* ── Save indicator ── */

function SaveIndicator({ saved, error }: { saved: boolean; error: boolean }) {
  if (saved) return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0 }}>
      <circle cx="8" cy="8" r="8" fill="#004D22" />
      <path d="M5 8l2 2 4-4" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
  if (error) return (
    <span title="Save failed — try again">
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0 }}>
        <circle cx="8" cy="8" r="8" fill="#B94040" />
        <path d="M5 5l6 6M11 5l-6 6" stroke="white" strokeWidth="1.5" strokeLinecap="round" />
      </svg>
    </span>
  )
  return null
}

/* ── Types ── */

type Tab = 'matrix'

interface ProjectInfo {
  id: string
  name: string
  address?: string | null
  bessScore?: number | null
  gdft?: boolean
}

const GDFT_CREDIT_NAMES = [
  'Ventilation - Natural - Apartments',
  'Thermal Performance Rating - Residential',
]

/* ── Main component ── */

export default function ReviewMatrix() {
  const { reviewLinkToken } = useParams<{ reviewLinkToken: string }>()
  const navigate = useNavigate()

  // Reviewer mode when a localStorage session exists for this token;
  // otherwise treat as GIW/admin direct access.
  const hasReviewerSession = !!(reviewLinkToken && getSession(reviewLinkToken))
  const isGIW = !hasReviewerSession

  const [session, setSession] = useState<ReviewerSession | null>(null)
  const [project, setProject] = useState<ProjectInfo | null>(null)
  const [credits, setCredits] = useState<Credit[]>([])
  const [excellenceItems, setExcellenceItems] = useState<ESDExcellenceOpportunity[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [activeTab] = useState<Tab>('matrix')

  // Per-credit comment state
  const [reviewerComments, setReviewerComments] = useState<Record<string, string>>({})
  const [savingCredits, setSavingCredits] = useState<Set<string>>(new Set())
  const [savedTicks, setSavedTicks] = useState<Set<string>>(new Set())
  const [saveErrors, setSaveErrors] = useState<Set<string>>(new Set())
  const tickTimers = useRef<Record<string, ReturnType<typeof setTimeout>>>({})

  // Per-credit GIW comment state
  const [giwComments, setGiwComments] = useState<Record<string, string>>({})
  const [giwSaveTicks, setGiwSaveTicks] = useState<Set<string>>(new Set())
  const [giwSaveErrors, setGiwSaveErrors] = useState<Set<string>>(new Set())

  // Submit state
  const [submitting, setSubmitting] = useState(false)
  const [submitSuccess, setSubmitSuccess] = useState(false)
  const [submitError, setSubmitError] = useState('')

  // Sticky badge bar visibility
  const headerRef = useRef<HTMLDivElement>(null)
  const [badgesSticky, setBadgesSticky] = useState(false)
  useEffect(() => {
    const el = headerRef.current
    if (!el) return
    const observer = new IntersectionObserver(
      ([entry]) => setBadgesSticky(!entry.isIntersecting),
      { threshold: 0 },
    )
    observer.observe(el)
    return () => observer.disconnect()
  }, [])

  // Excellence flag state
  const [flagging, setFlagging] = useState<Record<string, boolean>>({})
  const [localFlags, setLocalFlags] = useState<Record<string, string>>({})

  const interactiveBESS = useMemo(() => {
    if (project?.bessScore == null) return null
    let score = project.bessScore
    for (const item of excellenceItems) {
      const flag = localFlags[item.id] ?? item.flag
      if (flag !== 'Yes') continue
      if (item.creditReference === 'Innovation' && item.bessPoints) {
        const raw = Number(item.bessPoints)
        if (!isNaN(raw)) score += Math.round(raw * 0.9 * 10) / 10
      } else if (item.additionalBessPoints != null) {
        score += item.additionalBessPoints
      }
    }
    return Math.round(score)
  }, [project?.bessScore, excellenceItems, localFlags])

  useEffect(() => {
    if (project?.name) {
      document.title = `ESD Review Portal — ${project.name} | GIW Environmental Solutions`
    }
    return () => { document.title = 'ESD Review Portal | GIW Environmental Solutions' }
  }, [project?.name])

  useEffect(() => {
    if (!reviewLinkToken) return

    let resolvedSession: ReviewerSession | null = null
    if (!isGIW) {
      resolvedSession = getSession(reviewLinkToken)
      if (!resolvedSession) {
        navigate(`/review/${reviewLinkToken}`, { replace: true })
        return
      }
      setSession(resolvedSession)
    }

    const reviewerEmail = resolvedSession?.reviewerEmail
    const reviewerDiscipline = resolvedSession?.reviewerDiscipline

    ;(async () => {
      try {
        const projRes = await axios.get<ProjectInfo>(`/api/review/${reviewLinkToken}/project`)
        setProject(projRes.data)
        const pid = projRes.data.id

        const [creditsRes, exRes] = await Promise.all([
          axios.get<Credit[]>(
            isGIW
              ? `/api/projects/${pid}/credits`
              : `/api/projects/${pid}/credits?reviewerEmail=${encodeURIComponent(reviewerEmail!)}&reviewerDiscipline=${encodeURIComponent(reviewerDiscipline!)}`,
            { withCredentials: true },
          ),
          axios.get<{ items: ESDExcellenceOpportunity[], computedBESS: number } | ESDExcellenceOpportunity[]>(
            isGIW
              ? `/api/projects/${pid}/excellence`
              : `/api/projects/${pid}/excellence?reviewerEmail=${encodeURIComponent(reviewerEmail!)}&reviewerDiscipline=${encodeURIComponent(reviewerDiscipline!)}`,
            { withCredentials: true },
          ),
        ])

        setCredits(creditsRes.data)
        const exData = exRes.data
        const exItems = Array.isArray(exData) ? exData : exData.items
        setExcellenceItems(exItems)

        const initGIW: Record<string, string> = {}
        const initReviewer: Record<string, string> = {}
        for (const c of creditsRes.data) {
          initGIW[c.id] = c.commentsGIW ?? ''
          if (!isGIW) initReviewer[c.id] = c.comments?.[0]?.commentText ?? ''
        }
        setGiwComments(initGIW)
        setReviewerComments(initReviewer)

        const flags: Record<string, string> = {}
        for (const item of exItems) flags[item.id] = item.flag
        setLocalFlags(flags)
      } catch {
        setError('Failed to load review data.')
      } finally {
        setLoading(false)
      }
    })()
  }, [reviewLinkToken, isGIW]) // eslint-disable-line react-hooks/exhaustive-deps

  /* Comment save helpers */

  const showTick = useCallback((id: string, setTicks: React.Dispatch<React.SetStateAction<Set<string>>>) => {
    setTicks((s) => new Set(s).add(id))
    if (tickTimers.current[id]) clearTimeout(tickTimers.current[id])
    tickTimers.current[id] = setTimeout(() => {
      setTicks((s) => { const n = new Set(s); n.delete(id); return n })
    }, 2000)
  }, [])

  const saveReviewerComment = useCallback(
    async (creditId: string, commentText: string) => {
      if (!session) return
      setSavingCredits((s) => new Set(s).add(creditId))
      try {
        await axios.post(`/api/credits/${creditId}/comment`, {
          reviewerEmail: session.reviewerEmail,
          reviewerDiscipline: session.reviewerDiscipline,
          commentText,
        })
        showTick(creditId, setSavedTicks)
        setSaveErrors((s) => { const n = new Set(s); n.delete(creditId); return n })
      } catch {
        setSaveErrors((s) => new Set(s).add(creditId))
        setTimeout(() => {
          setSaveErrors((s) => { const n = new Set(s); n.delete(creditId); return n })
        }, 3000)
      } finally {
        setSavingCredits((s) => { const n = new Set(s); n.delete(creditId); return n })
      }
    },
    [session, showTick],
  )

  const saveGIWComment = useCallback(async (creditId: string, commentText: string) => {
    try {
      await axios.patch(
        `/api/credits/${creditId}/giw-comment`,
        { commentText },
        { withCredentials: true },
      )
      showTick(creditId, setGiwSaveTicks)
      setGiwSaveErrors((s) => { const n = new Set(s); n.delete(creditId); return n })
    } catch {
      setGiwSaveErrors((s) => new Set(s).add(creditId))
      setTimeout(() => {
        setGiwSaveErrors((s) => { const n = new Set(s); n.delete(creditId); return n })
      }, 3000)
    }
  }, [showTick])

  async function handleFlag(itemId: string, flag: string) {
    setFlagging((f) => ({ ...f, [itemId]: true }))
    try {
      const flaggedBy = isGIW ? 'admin@giw.com.au' : (session?.reviewerEmail ?? '')
      await axios.patch(`/api/excellence/${itemId}/flag`, { flag, flaggedBy })
      setLocalFlags((f) => ({ ...f, [itemId]: flag }))
    } catch { /* silent */ } finally {
      setFlagging((f) => { const n = { ...f }; delete n[itemId]; return n })
    }
  }

  async function handleDelete(itemId: string) {
    if (!window.confirm('Remove this opportunity from the review?')) return
    try {
      await axios.delete(`/api/excellence/${itemId}`, { withCredentials: true })
      setExcellenceItems((items) => items.filter((i) => i.id !== itemId))
    } catch { /* silent */ }
  }

  async function handleSubmit() {
    if (!session || !project) return
    const hasComment = Object.values(reviewerComments).some((v) => v.trim() !== '')
    if (!hasComment) {
      setSubmitError('Please add at least one project team comment before submitting.')
      return
    }
    setSubmitting(true)
    setSubmitError('')
    try {
      await axios.post(`/api/review/${project.id}/submit`, {
        reviewerEmail: session.reviewerEmail,
        reviewerDiscipline: session.reviewerDiscipline,
      })
      setSubmitSuccess(true)
    } catch {
      setSubmitError('Submission failed. Please try again.')
    } finally {
      setSubmitting(false)
    }
  }

  /* Derived */
  const commentedCount = Object.values(reviewerComments).filter((v) => v.trim() !== '').length
  const totalCount = credits.length
  const creditGroups = groupByCategory(credits)
  const colCount = isGIW ? 8 : 6

  /* ── Render ── */

  if (error) {
    return (
      <div className="min-h-screen" style={{ backgroundColor: '#FFFFFF' }}>
        <NavBar />
        <div className="flex flex-col items-center justify-center py-32 gap-4">
          <p style={{ fontFamily: 'Open Sans, sans-serif', color: '#B94040' }}>{error}</p>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen" style={{ backgroundColor: '#FFFFFF' }}>
      <NavBar />

      {/* Project header */}
      <div ref={headerRef} style={{ backgroundColor: '#00602B', padding: '20px 32px' }}>
        <div className="max-w-7xl mx-auto flex items-center justify-between gap-6 flex-wrap">
          <div style={{ minWidth: 0 }}>
            {loading
              ? <div className="skeleton-shimmer" style={{ height: 22, width: 260, marginBottom: 8 }} />
              : (
                <h1 style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '22px', fontWeight: 600, color: '#fff', marginBottom: '4px' }}>
                  {project?.name ?? ''}
                </h1>
              )
            }
            {loading
              ? <div className="skeleton-shimmer" style={{ height: 13, width: 180 }} />
              : project?.address && (
                <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: 'rgba(255,255,255,0.7)' }}>
                  {project.address}
                </p>
              )
            }
          </div>
          {loading
            ? <div className="skeleton-shimmer" style={{ height: 58, width: 80, borderRadius: '4px' }} />
            : project?.bessScore != null && (
              <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
                <div style={{ backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '4px', padding: '8px 18px', textAlign: 'center' }}>
                  <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '24px', fontWeight: 600, color: '#fff', lineHeight: 1 }}>
                    {project.bessScore}%
                  </p>
                  <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '11px', color: 'rgba(255,255,255,0.8)', marginTop: '2px' }}>
                    Baseline BESS Score
                  </p>
                </div>
                {interactiveBESS != null && (
                  <div style={{ backgroundColor: '#C8E6D4', borderRadius: '4px', padding: '8px 18px', textAlign: 'center' }}>
                    <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '24px', fontWeight: 600, color: '#004D22', lineHeight: 1 }}>
                      {interactiveBESS}%
                    </p>
                    <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '11px', color: '#004D22', marginTop: '2px' }}>
                      Improved BESS Score
                    </p>
                  </div>
                )}
              </div>
            )
          }
        </div>
      </div>

      {/* Sticky BESS badge bar */}
      {project?.bessScore != null && interactiveBESS != null && (
        <div style={{
          position: 'fixed', top: 60, right: 0, left: 0, zIndex: 40,
          backgroundColor: '#00602B',
          padding: '8px 32px',
          display: 'flex', justifyContent: 'flex-end', alignItems: 'center', gap: '8px',
          boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
          transform: badgesSticky ? 'translateY(0)' : 'translateY(-110%)',
          transition: 'transform 0.25s ease',
          pointerEvents: badgesSticky ? 'auto' : 'none',
        }}>
          <div style={{ backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '4px', padding: '5px 14px', textAlign: 'center' }}>
            <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '16px', fontWeight: 600, color: '#fff', lineHeight: 1 }}>
              {project.bessScore}%
            </p>
            <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '10px', color: 'rgba(255,255,255,0.8)', marginTop: '2px' }}>
              Baseline BESS Score
            </p>
          </div>
          <div style={{ backgroundColor: '#C8E6D4', borderRadius: '4px', padding: '5px 14px', textAlign: 'center' }}>
            <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '16px', fontWeight: 600, color: '#004D22', lineHeight: 1 }}>
              {interactiveBESS}%
            </p>
            <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '10px', color: '#004D22', marginTop: '2px' }}>
              Improved BESS Score
            </p>
          </div>
        </div>
      )}

      {/* Reviewer instructions */}
      {!isGIW && !loading && credits.length > 0 && (
        <div className="max-w-7xl mx-auto px-6 pt-6">
          <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#555', lineHeight: '1.6' }}>
            Please see below for the baseline BESS matrix and additional ESD Excellence and Innovation credits. Please provide comments where required and flag any ESD excellence / Innovation credits as yes, no or maybe. The interactive improved BESS score at the top of page will display the new BESS score based on selected additional credits. Once done, please submit review at the bottom of the page.
          </p>
        </div>
      )}

      {/* Progress bar (reviewer only) */}
      {!isGIW && !loading && (
        <div className="max-w-7xl mx-auto px-6 pt-5">
          <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#C0C0C0', marginBottom: '6px' }}>
            {commentedCount} of {totalCount} credit{totalCount !== 1 ? 's' : ''} with project team comments
          </p>
          <div style={{ height: '4px', backgroundColor: '#C0C0C0', borderRadius: '2px', overflow: 'hidden' }}>
            <div style={{
              height: '100%', backgroundColor: '#00602B', borderRadius: '2px',
              width: totalCount > 0 ? `${(commentedCount / totalCount) * 100}%` : '0%',
              transition: 'width 0.4s ease',
            }} />
          </div>
        </div>
      )}

      {/* Tab content */}
      <div className="max-w-7xl mx-auto px-6 py-6">
        {activeTab === 'matrix' && (
          <>
            {loading ? (
              <CreditTableSkeleton />
            ) : credits.length === 0 ? (
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
            ) : (
              <div style={{ borderRadius: '4px', border: '1px solid #C0C0C0' }}>
                <table style={{ borderCollapse: 'collapse', width: '100%', backgroundColor: '#fff', fontSize: '13px', fontFamily: 'Open Sans, sans-serif' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#00602B' }}>
                      {[
                        ['Credit Name', 200],
                        ['Credit Requirement', 220],
                        ['Mandatory', 70],
                        ['Responsible Party', 150],
                        ...(isGIW ? [['Credit Score', 80], ['Credit Weight', 80]] as [string, number][] : []),
                        ['Comments GIW', 220],
                        ['Comments Project Team', 260],
                      ].map(([label, w], i) => (
                        <th
                          key={i}
                          style={{
                            fontFamily: 'Montserrat, sans-serif', fontSize: '12px', fontWeight: 500,
                            color: '#fff', textTransform: 'uppercase', padding: '8px 12px',
                            textAlign: 'left', whiteSpace: 'nowrap', minWidth: w as number,
                            ...(i === 0 ? { position: 'sticky', left: 0, zIndex: 2, backgroundColor: '#00602B' } : {}),
                          }}
                        >
                          {label as string}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {creditGroups.map((group) => (
                      <Fragment key={group.order}>
                        <tr>
                          <td
                            colSpan={colCount}
                            style={{
                              backgroundColor: '#C8E6D4', padding: '8px 12px',
                              fontFamily: 'Montserrat, sans-serif', fontSize: '13px', fontWeight: 600, color: '#004D22',
                              position: 'sticky', left: 0,
                            }}
                          >
                            {group.category}
                          </td>
                        </tr>
                        {group.items.map((credit) => (
                          <CreditRow
                            key={credit.id}
                            credit={credit}
                            isGIW={isGIW}
                            gdft={project?.gdft ?? false}
                            currentReviewerEmail={session?.reviewerEmail}
                            giwComment={giwComments[credit.id] ?? ''}
                            onGIWCommentChange={(v) => setGiwComments((p) => ({ ...p, [credit.id]: v }))}
                            onGIWCommentBlur={() => saveGIWComment(credit.id, giwComments[credit.id] ?? '')}
                            isGIWSaved={giwSaveTicks.has(credit.id)}
                            isGIWError={giwSaveErrors.has(credit.id)}
                            reviewerComment={reviewerComments[credit.id] ?? ''}
                            onReviewerCommentChange={(v) => setReviewerComments((p) => ({ ...p, [credit.id]: v }))}
                            onReviewerCommentBlur={() => saveReviewerComment(credit.id, reviewerComments[credit.id] ?? '')}
                            isSaving={savingCredits.has(credit.id)}
                            isSaved={savedTicks.has(credit.id)}
                            isSaveError={saveErrors.has(credit.id)}
                          />
                        ))}
                      </Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* ESD Excellence Opportunities */}
            {(isGIW || ['architect', 'developer', 'esd consultant', 'services engineer'].includes((session?.reviewerDiscipline ?? '').toLowerCase())) && !loading && (
              excellenceItems.length > 0 ? (
                <ExcellenceSection
                  items={excellenceItems}
                  localFlags={localFlags}
                  flagging={flagging}
                  isGIW={isGIW}
                  onFlag={handleFlag}
                />
              ) : credits.length > 0 && (
                <div style={{ marginTop: '48px' }}>
                  <div style={{ marginBottom: '16px' }}>
                    <h2 style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '20px', fontWeight: 600, color: '#2C2C2C', marginBottom: '6px' }}>
                      ESD Excellence Opportunities
                    </h2>
                    <div style={{ height: '2px', width: '48px', backgroundColor: '#00602B', borderRadius: '1px' }} />
                  </div>
                  <div className="text-center py-12">
                    <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '12px' }}>
                      <svg width="36" height="36" viewBox="0 0 36 36" fill="none" opacity="0.3">
                        <path d="M18 4l3.5 7 7.5 1.1-5.5 5.3 1.3 7.6L18 21.5l-6.8 3.5 1.3-7.6L7 12.1l7.5-1.1z" stroke="#00602B" strokeWidth="2" strokeLinejoin="round" />
                      </svg>
                    </div>
                    <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#C0C0C0' }}>
                      No improvement opportunities identified
                    </p>
                  </div>
                </div>
              )
            )}

            {/* Submit button (reviewer only) */}
            {!isGIW && !loading && !submitSuccess && credits.length > 0 && (
              <div className="mt-8">
                {submitError && (
                  <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#B94040', marginBottom: '8px' }}>
                    {submitError}
                  </p>
                )}
                <button className="btn-primary" onClick={handleSubmit} disabled={submitting}>
                  {submitting ? 'Submitting…' : 'Submit My Review'}
                </button>
              </div>
            )}

            {submitSuccess && (
              <div style={{
                marginTop: '24px', padding: '14px 20px', borderRadius: '4px',
                backgroundColor: '#C8E6D4', border: '1px solid #00602B', color: '#004D22',
                fontFamily: 'Open Sans, sans-serif', fontSize: '14px',
              }}>
                Your review has been submitted. Thank you.
              </div>
            )}

            {loading && (
              <div style={{ marginTop: '48px' }}>
                <div style={{ marginBottom: '16px' }}>
                  <div className="skeleton-shimmer" style={{ height: 20, width: 240 }} />
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))', gap: '12px' }}>
                  {Array.from({ length: 3 }).map((_, i) => (
                    <div key={i} className="giw-card flex flex-col gap-3">
                      <div className="skeleton-shimmer" style={{ height: 14, width: '80%' }} />
                      <div className="skeleton-shimmer" style={{ height: 24, width: 80, borderRadius: '12px' }} />
                      <div className="skeleton-shimmer" style={{ height: 12 }} />
                      <div className="skeleton-shimmer" style={{ height: 12, width: '70%' }} />
                    </div>
                  ))}
                </div>
              </div>
            )}
          </>
        )}

      </div>
    </div>
  )
}

/* ── Credit row grouping ── */

function groupByCategory(credits: Credit[]) {
  const map = new Map<string, { category: string; order: number; items: Credit[] }>()
  for (const c of credits) {
    const order = getCategoryOrder(c.creditId, c.category)
    if (!map.has(c.category)) map.set(c.category, { category: c.category, order, items: [] })
    map.get(c.category)!.items.push(c)
  }
  return [...map.values()].sort((a, b) => a.order - b.order)
}

/* ── CreditRow ── */

interface CreditRowProps {
  credit: Credit
  isGIW: boolean
  gdft: boolean
  currentReviewerEmail?: string
  giwComment: string
  onGIWCommentChange: (v: string) => void
  onGIWCommentBlur: () => void
  isGIWSaved: boolean
  isGIWError: boolean
  reviewerComment: string
  onReviewerCommentChange: (v: string) => void
  onReviewerCommentBlur: () => void
  isSaving: boolean
  isSaved: boolean
  isSaveError: boolean
}

const tdBase: React.CSSProperties = {
  padding: '8px 12px', borderBottom: '1px solid #C0C0C0', verticalAlign: 'top',
  overflowWrap: 'break-word', wordBreak: 'break-word',
}

function CreditRow({
  credit, isGIW, gdft, currentReviewerEmail,
  giwComment, onGIWCommentChange, onGIWCommentBlur, isGIWSaved, isGIWError,
  reviewerComment, onReviewerCommentChange, onReviewerCommentBlur,
  isSaving, isSaved, isSaveError,
}: CreditRowProps) {
  const [hovered, setHovered] = useState(false)
  const rowBg = hovered ? '#C8E6D4' : '#fff'

  const teamCommentForGIW = (credit.comments ?? [])
    .filter(c => c.commentText.trim())
    .map(c => `${c.commentText.trim()}\n— ${c.reviewerDiscipline}`)
    .join('\n\n')

  return (
    <tr
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{ backgroundColor: rowBg, transition: 'background-color 0.1s' }}
    >
      {/* Credit Name — sticky */}
      <td style={{ ...tdBase, position: 'sticky', left: 0, backgroundColor: rowBg, zIndex: 1, minWidth: 200 }}>
        <span style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '13px', fontWeight: 500, color: '#2C2C2C' }}>
          {credit.creditId} {credit.creditName}
        </span>
      </td>

      {/* Requirement — collapsible */}
      <RequirementCell requirement={credit.creditRequirement} />

      {/* Mandatory */}
      <td style={{ ...tdBase, textAlign: 'center' }}>
        {(credit.mandatory || (gdft && GDFT_CREDIT_NAMES.includes(credit.creditName))) && (
          <span style={{ fontFamily: 'Montserrat, sans-serif', fontWeight: 700, fontSize: '12px', color: '#2C2C2C' }}>M</span>
        )}
        {gdft && GDFT_CREDIT_NAMES.includes(credit.creditName) && (
          <span style={{ display: 'block', marginTop: '4px', padding: '3px 8px', borderRadius: '3px', backgroundColor: '#FFF0E0', fontFamily: 'Montserrat, sans-serif', fontWeight: 700, fontSize: '13px', color: '#B84D00' }}>GDFT</span>
        )}
      </td>

      {/* Responsible Party */}
      <td style={{ ...tdBase, color: '#C0C0C0', fontSize: '12px' }}>{credit.responsibleParty ?? ''}</td>

      {/* Score + Weight (GIW only) */}
      {isGIW && (
        <>
          <td style={{ ...tdBase, fontSize: '12px' }}>{credit.creditScore != null ? `${credit.creditScore}%` : ''}</td>
          <td style={{ ...tdBase, fontSize: '12px' }}>{credit.creditWeight != null ? `${credit.creditWeight}%` : ''}</td>
        </>
      )}

      {/* Comments GIW */}
      <td style={{ ...tdBase, minWidth: 220 }}>
        {isGIW ? (
          <div style={{ position: 'relative' }}>
            <textarea
              value={giwComment}
              onChange={(e) => onGIWCommentChange(e.target.value)}
              onBlur={onGIWCommentBlur}
              rows={3}
              style={{
                width: '100%', resize: 'vertical', border: '1px solid #C0C0C0',
                borderRadius: '2px', padding: '4px 6px', fontSize: '12px',
                fontFamily: 'Open Sans, sans-serif', color: '#2C2C2C', backgroundColor: '#fff',
              }}
            />
            {(isGIWSaved || isGIWError) && (
              <span style={{ position: 'absolute', bottom: '6px', right: '8px' }}>
                <SaveIndicator saved={isGIWSaved} error={isGIWError} />
              </span>
            )}
          </div>
        ) : (
          <p style={{ fontSize: '12px', color: '#2C2C2C', margin: 0, whiteSpace: 'pre-wrap' }}>{giwComment}</p>
        )}
      </td>

      {/* Comments Project Team */}
      <td style={{ ...tdBase, minWidth: 260 }}>
        {isGIW ? (
          <p style={{ fontSize: '12px', color: '#2C2C2C', margin: 0, whiteSpace: 'pre-wrap' }}>{teamCommentForGIW}</p>
        ) : (
          <div>
            {/* Own comment input */}
            <div style={{ position: 'relative' }}>
              <textarea
                value={reviewerComment}
                onChange={(e) => onReviewerCommentChange(e.target.value)}
                onBlur={() => { if (!isSaving) onReviewerCommentBlur() }}
                rows={3}
                placeholder="Add your comments…"
                style={{
                  width: '100%', resize: 'vertical', border: '1px solid #C0C0C0',
                  borderRadius: '2px', padding: '4px 6px', fontSize: '12px',
                  fontFamily: 'Open Sans, sans-serif', backgroundColor: '#fff',
                }}
              />
              {(isSaved || isSaveError) && (
                <span style={{ position: 'absolute', bottom: '6px', right: '8px' }}>
                  <SaveIndicator saved={isSaved} error={isSaveError} />
                </span>
              )}
            </div>
            {/* Other reviewers' comments */}
            {(credit.comments ?? [])
              .filter(c => c.reviewerEmail !== currentReviewerEmail && c.commentText.trim())
              .map(c => (
                <div
                  key={c.id}
                  style={{
                    marginTop: '8px', padding: '6px 8px',
                    backgroundColor: '#F7F5F0', borderLeft: '3px solid #6B7A3B',
                    borderRadius: '2px',
                  }}
                >
                  <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '10px', fontWeight: 600, color: '#6B7A3B', margin: '0 0 3px 0' }}>
                    {c.reviewerDiscipline}
                  </p>
                  <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#2C2C2C', margin: 0, whiteSpace: 'pre-wrap' }}>
                    {c.commentText}
                  </p>
                </div>
              ))
            }
          </div>
        )}
      </td>
    </tr>
  )
}

/* ── RequirementCell ── */

function RequirementCell({ requirement }: { requirement?: string | null }) {
  const text = requirement ?? ''
  return (
    <td style={{ ...tdBase, color: '#C0C0C0', fontSize: '12px', minWidth: 220 }}>
      {text}
    </td>
  )
}

/* ── Credit table skeleton ── */

function CreditTableSkeleton() {
  const cols = [200, 220, 70, 150, 50, 50, 220, 260]
  return (
    <div style={{ borderRadius: '4px', border: '1px solid #C0C0C0' }}>
      <table style={{ borderCollapse: 'collapse', width: '100%', backgroundColor: '#fff' }}>
        <thead>
          <tr style={{ backgroundColor: '#00602B' }}>
            {cols.map((w, i) => (
              <th key={i} style={{ minWidth: w, padding: '10px 12px' }}>
                <div className="skeleton-shimmer" style={{ height: 11, borderRadius: 2 }} />
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: 8 }).map((_, i) => (
            <tr key={i} style={{ borderBottom: '1px solid #C0C0C0' }}>
              {cols.map((w, j) => (
                <td key={j} style={{ padding: '12px', minWidth: w }}>
                  <div className="skeleton-shimmer" style={{ height: 12, borderRadius: 2 }} />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}


/* ── ESD Excellence Section ── */

const FLAG_OPTIONS = ['Yes', 'No', 'Maybe'] as const

function AutoTextarea({
  value, onChange, onBlur, hasError,
}: { value: string; onChange: (v: string) => void; onBlur: () => void; hasError: boolean }) {
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
      style={{
        width: '100%', resize: 'none', overflow: 'hidden',
        border: `1px solid ${hasError ? '#B94040' : '#C0C0C0'}`,
        borderRadius: '2px', padding: '4px 6px',
        fontSize: '12px', fontFamily: 'Open Sans, sans-serif',
        color: '#2C2C2C', backgroundColor: '#fff',
        lineHeight: '1.5', minHeight: '28px', boxSizing: 'border-box',
      }}
    />
  )
}

interface ExcellenceSectionProps {
  items: ESDExcellenceOpportunity[]
  localFlags: Record<string, string>
  flagging: Record<string, boolean>
  isGIW: boolean
  onFlag: (id: string, flag: string) => void
}

function ExcellenceSection({ items, localFlags, flagging, isGIW, onFlag }: ExcellenceSectionProps) {
  const [notesMap, setNotesMap] = useState<Record<string, string>>({})
  const [notesSaved, setNotesSaved] = useState<Set<string>>(new Set())
  const [notesErrors, setNotesErrors] = useState<Set<string>>(new Set())

  const [descMap, setDescMap] = useState<Record<string, string>>({})
  const [descSaved, setDescSaved] = useState<Set<string>>(new Set())
  const [descErrors, setDescErrors] = useState<Set<string>>(new Set())

  // Sync maps when items load or change
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
      await axios.patch(`/api/excellence/${id}/notes`, { reviewerNotes: notesMap[id] ?? '' })
      setNotesSaved((s) => new Set([...s, id]))
      setNotesErrors((s) => { const n = new Set(s); n.delete(id); return n })
      setTimeout(() => setNotesSaved((s) => { const n = new Set(s); n.delete(id); return n }), 2000)
    } catch {
      setNotesErrors((s) => new Set([...s, id]))
    }
  }

  async function saveDesc(id: string) {
    try {
      await axios.patch(`/api/excellence/${id}/description`, { improvementDescription: descMap[id] ?? '' })
      setDescSaved((s) => new Set([...s, id]))
      setDescErrors((s) => { const n = new Set(s); n.delete(id); return n })
      setTimeout(() => setDescSaved((s) => { const n = new Set(s); n.delete(id); return n }), 2000)
    } catch {
      setDescErrors((s) => new Set([...s, id]))
    }
  }

  const regularItems = [...items.filter(i => i.creditReference !== 'Innovation')]
    .sort((a, b) => {
      const oa = getCategoryOrder(a.creditReference)
      const ob = getCategoryOrder(b.creditReference)
      if (oa !== ob) return oa - ob
      return (a.currentScore ?? 0) - (b.currentScore ?? 0)
    })
    .filter((item, idx, arr) => arr.findIndex(x => x.creditReference === item.creditReference) === idx)

  const innovationItems = [...items.filter(i => i.creditReference === 'Innovation')]
    .sort((a, b) => a.creditName.localeCompare(b.creditName))
    .filter((item, idx, arr) => arr.findIndex(x => x.creditName === item.creditName) === idx)

  function cardProps(item: ESDExcellenceOpportunity) {
    return {
      key: item.id,
      item,
      flag: localFlags[item.id] ?? item.flag,
      isFlagging: !!flagging[item.id],
      isGIW,
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
    }
  }

  const baseCols: [string, number][] = [
    ['Credit Ref', 90], ['Credit Name', 180], ['Current Score', 100], ['BESS Points', 100],
    ['Improvement Description', 260], ['Flag', 100], ['Comments', 220],
  ]
  const cols = baseCols

  function ExcellenceTable({ rows, mb, hideCurrentScore }: { rows: ESDExcellenceOpportunity[], mb?: string, hideCurrentScore?: boolean }) {
    const tableCols = hideCurrentScore ? cols.filter(([label]) => label !== 'Current Score') : cols
    return (
      <div style={{ borderRadius: '4px', border: '1px solid #C0C0C0', marginBottom: mb }}>
        <table style={{ borderCollapse: 'collapse', width: '100%', backgroundColor: '#fff', fontSize: '13px', fontFamily: 'Open Sans, sans-serif' }}>
          <thead>
            <tr style={{ backgroundColor: '#00602B' }}>
              {tableCols.map(([label, w]) => (
                <th key={label} style={{
                  fontFamily: 'Montserrat, sans-serif', fontSize: '12px', fontWeight: 500,
                  color: '#fff', textTransform: 'uppercase', padding: '8px 12px',
                  textAlign: 'left', whiteSpace: 'nowrap', minWidth: w,
                }}>
                  {label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((item) => <ExcellenceRow {...cardProps(item)} hideCurrentScore={hideCurrentScore} />)}
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
            <div style={{ height: '2px', width: '48px', backgroundColor: '#00602B', borderRadius: '1px', marginBottom: '8px' }} />
            <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#8C8C8C' }}>
              The following additional BESS points have been identified to improve the BESS score. Note that these items are not mandatory.
            </p>
          </div>
          <ExcellenceTable rows={regularItems} mb="48px" />
        </>
      )}

      {innovationItems.length > 0 && (
        <>
          <div style={{ marginBottom: '16px' }}>
            <h2 style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '20px', fontWeight: 600, color: '#2C2C2C', marginBottom: '6px' }}>
              Innovation Credits
            </h2>
            <div style={{ height: '2px', width: '48px', backgroundColor: '#00602B', borderRadius: '1px', marginBottom: '8px' }} />
            <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#8C8C8C' }}>
              Maximum 10 innovation points can be claimed per project
            </p>
          </div>
          <ExcellenceTable rows={innovationItems} hideCurrentScore />
        </>
      )}
    </div>
  )
}

interface ExcellenceRowProps {
  item: ESDExcellenceOpportunity
  flag: string
  isFlagging: boolean
  isGIW: boolean
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
  hideCurrentScore?: boolean
}

function ExcellenceRow({ item, flag, isFlagging, isGIW, desc, onDescChange, onDescBlur, descSaved, descError, notes, onNotesChange, onNotesBlur, notesSaved, notesError, onFlag, hideCurrentScore }: ExcellenceRowProps) {
  const [hovered, setHovered] = useState(false)
  return (
    <tr
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{ backgroundColor: hovered ? '#C8E6D4' : '#fff', transition: 'background-color 0.1s' }}
    >
      <td style={{ ...tdBase, minWidth: 90 }}>
        <span style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '13px', fontWeight: 500, color: '#2C2C2C' }}>
          {item.creditReference}
        </span>
      </td>

      <td style={{ ...tdBase, minWidth: 180, color: '#2C2C2C' }}>
        {item.creditName}
      </td>

      {!hideCurrentScore && (
        <td style={{ ...tdBase, minWidth: 100 }}>
          {item.currentScore != null ? (
            <span style={{ display: 'inline-block', backgroundColor: '#C8E6D4', color: '#004D22', fontFamily: 'Open Sans, sans-serif', fontSize: '12px', padding: '2px 10px', borderRadius: '12px' }}>
              {item.currentScore}%
            </span>
          ) : (
            <span style={{ color: '#C0C0C0', fontSize: '12px' }}>—</span>
          )}
        </td>
      )}

      <td style={{ ...tdBase, minWidth: 100 }}>
        {item.bessPoints ? (
          <span style={{ display: 'inline-block', backgroundColor: '#E8EDD8', color: '#4E5A2A', fontFamily: 'Open Sans, sans-serif', fontSize: '12px', padding: '2px 10px', borderRadius: '12px' }}>
            {Math.round(Number(item.bessPoints) * 0.9 * 10) / 10} BESS {Number(item.bessPoints) * 0.9 === 1 ? 'point' : 'points'}
          </span>
        ) : item.additionalBessPoints != null ? (
          <span style={{ display: 'inline-block', backgroundColor: '#E8EDD8', color: '#4E5A2A', fontFamily: 'Open Sans, sans-serif', fontSize: '12px', padding: '2px 10px', borderRadius: '12px' }}>
            {item.additionalBessPoints} BESS {item.additionalBessPoints === 1 ? 'point' : 'points'}
          </span>
        ) : (
          <span style={{ color: '#C0C0C0', fontSize: '12px' }}>—</span>
        )}
      </td>

      <td style={{ ...tdBase, minWidth: 260 }}>
        {isGIW ? (
          <div style={{ position: 'relative' }}>
            <AutoTextarea value={desc} onChange={onDescChange} onBlur={onDescBlur} hasError={descError} />
            <SaveIndicator saved={descSaved} error={descError} />
          </div>
        ) : (
          desc
            ? <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#2C2C2C', margin: 0, lineHeight: 1.6, whiteSpace: 'pre-line' }}>{desc}</p>
            : <span style={{ color: '#C0C0C0', fontSize: '12px' }}>—</span>
        )}
      </td>

      <td style={{ ...tdBase, minWidth: 100 }}>
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

      <td style={{ ...tdBase, minWidth: 220 }}>
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
          <SaveIndicator saved={notesSaved} error={notesError} />
        </div>
      </td>

    </tr>
  )
}
