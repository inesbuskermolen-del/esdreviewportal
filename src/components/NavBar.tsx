import { useState, useEffect } from 'react'
import { Link, useLocation } from 'react-router-dom'
import { useAuth } from '@/lib/auth'

export default function NavBar() {
  const { user } = useAuth()
  const location = useLocation()
  const [menuOpen, setMenuOpen] = useState(false)

  // Close menu on route change
  useEffect(() => { setMenuOpen(false) }, [location.pathname])

  // Prevent body scroll when menu open
  useEffect(() => {
    document.body.style.overflow = menuOpen ? 'hidden' : ''
    return () => { document.body.style.overflow = '' }
  }, [menuOpen])

  const isActive = (path: string) => location.pathname === path

  return (
    <>
      <nav
        className="w-full sticky top-0 z-50 flex items-center justify-between px-6"
        style={{ backgroundColor: '#00602B', height: '60px' }}
      >
        {/* Logo */}
        <Link to={user ? '/admin' : '/'} className="flex items-center no-underline">
          <img src={`${import.meta.env.BASE_URL}GIW logo.png`} alt="GIW Environmental Solutions" style={{ height: '50px' }} />
        </Link>

        {/* Desktop nav */}
        <div className="hidden md:flex items-center gap-6">
          {user && (
            <Link
              to="/admin"
              className="relative text-white hover:text-gray-200 transition-colors pb-1"
              style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '14px' }}
            >
              Admin
              {isActive('/admin') && (
                <span
                  className="absolute bottom-0 left-0 right-0 h-0.5"
                  style={{ backgroundColor: '#ffffff' }}
                />
              )}
            </Link>
          )}
        </div>

        {/* Mobile hamburger */}
        <button
          className="md:hidden flex flex-col gap-1 p-1"
          onClick={() => setMenuOpen(true)}
          aria-label="Open menu"
          style={{ background: 'none', border: 'none', cursor: 'pointer' }}
        >
          <span style={{ display: 'block', width: 22, height: 2, backgroundColor: '#fff', borderRadius: 1 }} />
          <span style={{ display: 'block', width: 22, height: 2, backgroundColor: '#fff', borderRadius: 1 }} />
          <span style={{ display: 'block', width: 22, height: 2, backgroundColor: '#fff', borderRadius: 1 }} />
        </button>
      </nav>

      {/* Mobile slide-in menu */}
      {menuOpen && (
        <div style={{ position: 'fixed', inset: 0, zIndex: 200 }}>
          {/* Overlay */}
          <div
            style={{ position: 'absolute', inset: 0, backgroundColor: 'rgba(0,0,0,0.5)' }}
            onClick={() => setMenuOpen(false)}
          />
          {/* Panel */}
          <div
            style={{
              position: 'absolute', top: 0, right: 0,
              width: 280, height: '100%',
              backgroundColor: '#00602B',
              padding: '24px',
              display: 'flex', flexDirection: 'column', gap: '8px',
            }}
          >
            {/* Close button */}
            <button
              onClick={() => setMenuOpen(false)}
              style={{
                alignSelf: 'flex-end', background: 'none', border: 'none',
                color: 'rgba(255,255,255,0.6)', fontSize: '24px',
                cursor: 'pointer', lineHeight: 1, marginBottom: '16px',
              }}
              aria-label="Close menu"
            >
              ×
            </button>

            {user && (
              <Link
                to="/admin"
                style={{
                  fontFamily: 'Montserrat, sans-serif', fontSize: '15px',
                  color: '#fff', textDecoration: 'none', padding: '10px 0',
                  borderBottom: '1px solid rgba(255,255,255,0.1)',
                }}
              >
                Admin
              </Link>
            )}
          </div>
        </div>
      )}
    </>
  )
}
