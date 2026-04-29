import { useEffect, useState } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import axios from 'axios'
import NavBar from '@/components/NavBar'
import type { ReviewerSession } from '@/types'

const SESSION_KEY = 'giw_reviewer_session'

function saveSession(session: ReviewerSession): void {
  localStorage.setItem(SESSION_KEY, JSON.stringify(session))
}

export default function ReviewInvite() {
  const { inviteToken } = useParams<{ inviteToken: string }>()
  const navigate = useNavigate()
  const [error, setError] = useState('')

  useEffect(() => {
    if (!inviteToken) return
    axios.get<{
      reviewerId: string
      projectId: string
      reviewLinkToken: string
      projectName: string
      reviewerEmail: string
      reviewerDiscipline: string
    }>(`/api/review/invite/${inviteToken}`)
      .then((res) => {
        const session: ReviewerSession = {
          reviewerEmail: res.data.reviewerEmail,
          reviewerDiscipline: res.data.reviewerDiscipline,
          reviewerId: res.data.reviewerId,
          projectId: res.data.projectId,
          reviewLinkToken: res.data.reviewLinkToken,
          createdAt: new Date().toISOString(),
        }
        saveSession(session)
        navigate(`/review/${res.data.reviewLinkToken}/matrix`, { replace: true })
      })
      .catch(() => setError('This invite link is invalid or has expired. Please contact GIW Environmental Solutions.'))
  }, [inviteToken, navigate])

  return (
    <div className="min-h-screen" style={{ backgroundColor: '#E1E1E1' }}>
      <NavBar />
      <div className="flex items-start justify-center pt-24 px-4">
        <div style={{
          maxWidth: '480px', width: '100%',
          backgroundColor: '#FFFFFF', border: '1px solid #C0C0C0',
          borderRadius: '4px', padding: '32px', textAlign: 'center',
        }}>
          {error ? (
            <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#B94040' }}>
              {error}
            </p>
          ) : (
            <>
              <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '16px' }}>
                {/* Simple CSS spinner */}
                <div style={{
                  width: 32, height: 32, border: '3px solid #E1E1E1',
                  borderTop: '3px solid #00602B', borderRadius: '50%',
                  animation: 'spin 0.8s linear infinite',
                }} />
              </div>
              <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#2C2C2C' }}>
                Opening your review…
              </p>
              <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
            </>
          )}
        </div>
      </div>
    </div>
  )
}
