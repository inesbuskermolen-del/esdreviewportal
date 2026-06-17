import { useState, useEffect, useRef } from 'react'
import { CREDIT_SUMMARIES } from '@/lib/creditSummaries'

export function CreditInfoButton({ creditId }: { creditId: string }) {
  const summary = CREDIT_SUMMARIES[creditId]
  const [open, setOpen] = useState(false)
  const [pos, setPos] = useState({ top: 0, left: 0 })
  const btnRef = useRef<HTMLButtonElement>(null)
  const popupRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    function handleDown(e: MouseEvent) {
      if (
        popupRef.current && !popupRef.current.contains(e.target as Node) &&
        btnRef.current && !btnRef.current.contains(e.target as Node)
      ) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', handleDown)
    return () => document.removeEventListener('mousedown', handleDown)
  }, [open])

  if (!summary) return null

  function handleClick(e: React.MouseEvent) {
    e.stopPropagation()
    if (!open && btnRef.current) {
      const rect = btnRef.current.getBoundingClientRect()
      // Clamp so the popup doesn't overflow the right edge of the viewport
      const popupWidth = 320
      const left = Math.min(rect.left, window.innerWidth - popupWidth - 12)
      setPos({ top: rect.bottom + 6, left })
    }
    setOpen(v => !v)
  }

  return (
    <>
      <button
        ref={btnRef}
        onClick={handleClick}
        style={{
          display: 'inline',
          background: 'none',
          border: 'none',
          padding: 0,
          marginLeft: '6px',
          cursor: 'pointer',
          fontFamily: 'Open Sans, sans-serif',
          fontSize: '11px',
          color: '#004D22',
          textDecoration: 'underline',
          lineHeight: 'inherit',
          verticalAlign: 'baseline',
        }}
      >
        More info
      </button>
      {open && (
        <div
          ref={popupRef}
          style={{
            position: 'fixed',
            top: pos.top,
            left: pos.left,
            zIndex: 9999,
            backgroundColor: '#004D22',
            color: '#FFFFFF',
            padding: '12px 16px',
            borderRadius: '4px',
            width: '320px',
            fontFamily: 'Open Sans, sans-serif',
            fontSize: '12px',
            lineHeight: '1.6',
            boxShadow: '0 4px 16px rgba(0,0,0,0.25)',
          }}
        >
          {summary}
        </div>
      )}
    </>
  )
}
