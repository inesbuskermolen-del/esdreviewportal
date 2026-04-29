import { useEffect, useState } from 'react'
import { useSearchParams, useNavigate, Link } from 'react-router-dom'
import axios from 'axios'

export default function AuthVerify() {
  const [params] = useSearchParams()
  const navigate = useNavigate()
  const [error, setError] = useState('')

  useEffect(() => {
    const token = params.get('token')
    if (!token) {
      setError('Invalid login link.')
      return
    }

    axios
      .get(`/api/auth/verify?token=${encodeURIComponent(token)}`, {
        withCredentials: true,
      })
      .then(() => navigate('/admin', { replace: true }))
      .catch(() =>
        setError(
          'This login link has expired or is invalid. Please request a new one.',
        ),
      )
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div
      className="min-h-screen flex items-start justify-center pt-24 px-4"
      style={{ backgroundColor: '#E1E1E1' }}
    >
      <div
        style={{
          width: '100%',
          maxWidth: '480px',
          backgroundColor: '#FFFFFF',
          border: '1px solid #C0C0C0',
          borderRadius: '4px',
          padding: '32px',
          textAlign: 'center',
        }}
      >
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

        {error ? (
          <>
            <p
              style={{
                fontFamily: 'Open Sans, sans-serif',
                fontSize: '14px',
                color: '#B94040',
                marginBottom: '16px',
              }}
            >
              {error}
            </p>
            <Link
              to="/admin/login"
              style={{
                fontFamily: 'Open Sans, sans-serif',
                fontSize: '13px',
                color: '#00602B',
              }}
            >
              Back to sign in
            </Link>
          </>
        ) : (
          <p
            style={{
              fontFamily: 'Open Sans, sans-serif',
              fontSize: '14px',
              color: '#C0C0C0',
            }}
          >
            Verifying…
          </p>
        )}
      </div>
    </div>
  )
}
