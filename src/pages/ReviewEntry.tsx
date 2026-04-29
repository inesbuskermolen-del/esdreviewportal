import { useState, useEffect, FormEvent } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import axios from 'axios'
import NavBar from '@/components/NavBar'
import type { ReviewerSession } from '@/types'

const SESSION_KEY = 'giw_reviewer_session'
const SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000 // 30 days

function getSession(reviewLinkToken: string): ReviewerSession | null {
  try {
    const raw = localStorage.getItem(SESSION_KEY)
    if (!raw) return null
    const s = JSON.parse(raw) as ReviewerSession
    if (s.reviewLinkToken !== reviewLinkToken) return null
    if (Date.now() - new Date(s.createdAt).getTime() > SESSION_TTL_MS) {
      localStorage.removeItem(SESSION_KEY)
      return null
    }
    return s
  } catch {
    return null
  }
}

function saveSession(session: ReviewerSession): void {
  localStorage.setItem(SESSION_KEY, JSON.stringify(session))
}

const DISCIPLINES = [
  'Architect',
  'Services Engineer',
  'Civil Engineer',
  'Landscape Architect',
  'Developer',
  'Waste Consultant',
  'ESD Consultant',
  'Other',
]

interface ProjectInfo {
  name: string
  address?: string | null
  bessScore?: number | null
}

export default function ReviewEntry() {
  const { reviewLinkToken } = useParams<{ reviewLinkToken: string }>()
  const navigate = useNavigate()

  const [project, setProject] = useState<ProjectInfo | null>(null)
  const [loadError, setLoadError] = useState('')

  const [email, setEmail] = useState('')
  const [discipline, setDiscipline] = useState('')
  const [otherDiscipline, setOtherDiscipline] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [formError, setFormError] = useState('')

  // Check for existing valid session before rendering form
  useEffect(() => {
    if (!reviewLinkToken) return

    // Fast path: valid existing session
    const existing = getSession(reviewLinkToken)
    if (existing) {
      navigate(`/review/${reviewLinkToken}/matrix`, { replace: true })
      return
    }

    // Fetch project info for display
    axios
      .get<ProjectInfo>(`/api/review/${reviewLinkToken}/project`)
      .then((res) => setProject(res.data))
      .catch(() => setLoadError('This review link is invalid or has expired.'))
  }, [reviewLinkToken, navigate])

  function isValidEmail(v: string) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault()
    setFormError('')

    if (!isValidEmail(email.trim())) {
      setFormError('Please enter a valid email address.')
      return
    }

    const finalDiscipline =
      discipline === 'Other' ? otherDiscipline.trim() || 'Other' : discipline

    if (!finalDiscipline) {
      setFormError('Please select a discipline.')
      return
    }

    setSubmitting(true)
    try {
      const res = await axios.post<{
        reviewerId: string
        projectId: string
        projectName: string
        projectAddress: string | null
      }>('/api/review/identify', {
        reviewLinkToken,
        email: email.trim().toLowerCase(),
        discipline: finalDiscipline,
      })

      const session: ReviewerSession = {
        reviewerEmail: email.trim().toLowerCase(),
        reviewerDiscipline: finalDiscipline,
        reviewerId: res.data.reviewerId,
        projectId: res.data.projectId,
        reviewLinkToken: reviewLinkToken!,
        createdAt: new Date().toISOString(),
      }
      saveSession(session)
      navigate(`/review/${reviewLinkToken}/matrix`, { replace: true })
    } catch {
      setFormError('Something went wrong. Please try again.')
    } finally {
      setSubmitting(false)
    }
  }

  if (loadError) {
    return (
      <div className="min-h-screen" style={{ backgroundColor: '#E1E1E1' }}>
        <NavBar />
        <div className="flex items-start justify-center pt-24 px-4">
          <div
            style={{
              maxWidth: '480px',
              width: '100%',
              backgroundColor: '#FFFFFF',
              border: '1px solid #C0C0C0',
              borderRadius: '4px',
              padding: '32px',
            }}
          >
            <p
              style={{
                fontFamily: 'Open Sans, sans-serif',
                fontSize: '14px',
                color: '#B94040',
              }}
            >
              {loadError}
            </p>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen" style={{ backgroundColor: '#E1E1E1' }}>
      <NavBar />

      <div className="flex items-start justify-center pt-16 px-4 pb-16">
        <div
          style={{
            maxWidth: '480px',
            width: '100%',
            backgroundColor: '#FFFFFF',
            border: '1px solid #C0C0C0',
            borderRadius: '4px',
            padding: '32px',
          }}
        >
          {/* GIW branding */}
          <p
            style={{
              fontFamily: 'Montserrat, sans-serif',
              fontSize: '14px',
              fontWeight: 600,
              color: '#2C2C2C',
              marginBottom: '20px',
            }}
          >
            GIW Environmental Solutions
          </p>

          {/* Project info */}
          {project ? (
            <>
              <h1
                style={{
                  fontFamily: 'Montserrat, sans-serif',
                  fontSize: '22px',
                  fontWeight: 600,
                  color: '#2C2C2C',
                  marginBottom: project.address ? '6px' : '0',
                }}
              >
                {project.name}
              </h1>
              {project.address && (
                <p
                  style={{
                    fontFamily: 'Open Sans, sans-serif',
                    fontSize: '13px',
                    color: '#C0C0C0',
                    marginBottom: '0',
                  }}
                >
                  {project.address}
                </p>
              )}
            </>
          ) : (
            <div style={{ height: '52px' }} /> /* placeholder while loading */
          )}

          {/* Divider */}
          <div
            style={{
              borderTop: '1px solid #C0C0C0',
              margin: '20px 0',
            }}
          />

          <p
            style={{
              fontFamily: 'Montserrat, sans-serif',
              fontSize: '16px',
              fontWeight: 500,
              color: '#2C2C2C',
              marginBottom: '20px',
            }}
          >
            Enter your details to access this review
          </p>

          {formError && (
            <p
              className="mb-4 text-sm"
              style={{ fontFamily: 'Open Sans, sans-serif', color: '#B94040' }}
            >
              {formError}
            </p>
          )}

          <form onSubmit={handleSubmit} className="space-y-4">
            {/* Email */}
            <div>
              <label
                htmlFor="reviewer-email"
                className="block text-sm mb-1"
                style={{ fontFamily: 'Open Sans, sans-serif', color: '#2C2C2C' }}
              >
                Email address
              </label>
              <input
                id="reviewer-email"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
                placeholder="your@email.com"
                autoComplete="email"
                className="giw-input"
              />
            </div>

            {/* Discipline */}
            <div>
              <label
                htmlFor="reviewer-discipline"
                className="block text-sm mb-1"
                style={{ fontFamily: 'Open Sans, sans-serif', color: '#2C2C2C' }}
              >
                Discipline
              </label>
              <select
                id="reviewer-discipline"
                value={discipline}
                onChange={(e) => setDiscipline(e.target.value)}
                required
                className="giw-input"
                style={{ cursor: 'pointer' }}
              >
                <option value="" disabled>
                  Select your discipline…
                </option>
                {DISCIPLINES.map((d) => (
                  <option key={d} value={d}>
                    {d}
                  </option>
                ))}
              </select>
            </div>

            {/* Other free text */}
            {discipline === 'Other' && (
              <div>
                <label
                  htmlFor="reviewer-other"
                  className="block text-sm mb-1"
                  style={{ fontFamily: 'Open Sans, sans-serif', color: '#2C2C2C' }}
                >
                  Please specify
                </label>
                <input
                  id="reviewer-other"
                  type="text"
                  value={otherDiscipline}
                  onChange={(e) => setOtherDiscipline(e.target.value)}
                  placeholder="Your discipline"
                  className="giw-input"
                />
              </div>
            )}

            <button
              type="submit"
              disabled={submitting}
              className="btn-primary w-full"
            >
              {submitting ? 'Accessing…' : 'Access Review'}
            </button>
          </form>

          <p
            style={{
              fontFamily: 'Open Sans, sans-serif',
              fontSize: '11px',
              color: '#C0C0C0',
              marginTop: '12px',
            }}
          >
            Your details are used only to route comments to the correct team
            members.
          </p>
        </div>
      </div>
    </div>
  )
}
