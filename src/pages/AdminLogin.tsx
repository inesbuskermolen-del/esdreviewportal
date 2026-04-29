import { useState, FormEvent } from 'react'
import axios from 'axios'

export default function AdminLogin() {
  const [email, setEmail] = useState('')
  const [loading, setLoading] = useState(false)
  const [sent, setSent] = useState(false)

  async function handleSubmit(e: FormEvent) {
    e.preventDefault()
    setLoading(true)
    try {
      await axios.post('/api/auth/request-link', { email }, { withCredentials: true })
    } finally {
      setLoading(false)
      setSent(true)
    }
  }

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
        }}
      >
        {/* Logo text */}
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

        <h1
          style={{
            fontFamily: 'Montserrat, sans-serif',
            fontSize: '22px',
            fontWeight: 600,
            color: '#2C2C2C',
            marginBottom: '24px',
          }}
        >
          Staff Sign In
        </h1>

        {sent ? (
          <p
            style={{
              fontFamily: 'Open Sans, sans-serif',
              fontSize: '14px',
              color: '#004D22',
            }}
          >
            Check your GIW email for a login link.
          </p>
        ) : (
          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label
                htmlFor="email"
                className="block text-sm mb-1"
                style={{ fontFamily: 'Open Sans, sans-serif', color: '#2C2C2C' }}
              >
                Email address
              </label>
              <input
                id="email"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
                autoComplete="email"
                placeholder="you@giw.com.au"
                className="giw-input"
              />
            </div>

            <button
              type="submit"
              disabled={loading}
              className="btn-primary w-full"
            >
              {loading ? 'Sending…' : 'Send Login Link'}
            </button>
          </form>
        )}
      </div>
    </div>
  )
}
