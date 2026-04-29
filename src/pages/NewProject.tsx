import { useState, useRef, DragEvent, ChangeEvent } from 'react'
import { useNavigate } from 'react-router-dom'
import axios, { AxiosProgressEvent } from 'axios'
import NavBar from '@/components/NavBar'

type Phase = 'idle' | 'uploading' | 'parsing' | 'done' | 'error'

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

export default function NewProject() {
  const navigate = useNavigate()
  const fileInputRef = useRef<HTMLInputElement>(null)

  const [file, setFile] = useState<File | null>(null)
  const [dragging, setDragging] = useState(false)
  const [phase, setPhase] = useState<Phase>('idle')
  const [uploadPct, setUploadPct] = useState(0)
  const [error, setError] = useState('')

  function acceptFile(f: File) {
    if (f.type !== 'application/pdf') {
      setError('Only PDF files are accepted.')
      return
    }
    if (f.size > 20 * 1024 * 1024) {
      setError('File must be 20 MB or smaller.')
      return
    }
    setError('')
    setFile(f)
  }

  function handleInputChange(e: ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0]
    if (f) acceptFile(f)
    e.target.value = ''
  }

  function handleDrop(e: DragEvent<HTMLDivElement>) {
    e.preventDefault()
    setDragging(false)
    const f = e.dataTransfer.files[0]
    if (f) acceptFile(f)
  }

  async function handleUpload() {
    if (!file) return
    setError('')
    setPhase('uploading')
    setUploadPct(0)

    const data = new FormData()
    data.append('pdf', file)

    try {
      const res = await axios.post<{ projectId: string }>(
        '/api/projects/create-from-pdf',
        data,
        {
          withCredentials: true,
          headers: { 'Content-Type': 'multipart/form-data' },
          onUploadProgress: (event: AxiosProgressEvent) => {
            if (event.total) {
              const pct = Math.round((event.loaded / event.total) * 90)
              setUploadPct(pct)
              if (pct >= 90) setPhase('parsing')
            }
          },
        },
      )
      setPhase('done')
      navigate(`/admin/projects/${res.data.projectId}`)
    } catch (err) {
      setPhase('error')
      const msg =
        axios.isAxiosError(err) && err.response?.data?.error
          ? String(err.response.data.error)
          : 'Upload failed. Please try again.'
      setError(msg)
    }
  }

  const isUploading = phase === 'uploading' || phase === 'parsing'
  const showBar = isUploading || phase === 'done'

  return (
    <div className="min-h-screen" style={{ backgroundColor: '#FFFFFF' }}>
      <NavBar />

      <main className="max-w-2xl mx-auto px-6 py-10">
        <div className="mb-8">
          <button
            onClick={() => navigate('/admin')}
            className="text-sm mb-4 flex items-center gap-1 hover:opacity-70 transition-opacity"
            style={{ fontFamily: 'Open Sans, sans-serif', color: '#C0C0C0' }}
          >
            ← Back to Projects
          </button>
          <h1
            className="font-semibold"
            style={{
              fontFamily: 'Montserrat, sans-serif',
              fontSize: '28px',
              fontWeight: 600,
              color: '#2C2C2C',
            }}
          >
            New Project
          </h1>
          <p
            className="mt-1 text-sm"
            style={{ fontFamily: 'Open Sans, sans-serif', color: '#C0C0C0' }}
          >
            Upload a BESS assessment PDF to create a new project. Project details will be
            extracted automatically.
          </p>
        </div>

        {/* Upload zone */}
        <div
          role="button"
          tabIndex={0}
          onClick={() => !isUploading && fileInputRef.current?.click()}
          onKeyDown={(e) => {
            if ((e.key === 'Enter' || e.key === ' ') && !isUploading)
              fileInputRef.current?.click()
          }}
          onDragOver={(e) => {
            e.preventDefault()
            if (!isUploading) setDragging(true)
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          style={{
            border: `2px dashed ${dragging ? '#00602B' : '#C0C0C0'}`,
            backgroundColor: dragging ? '#C8E6D4' : '#FFFFFF',
            borderRadius: '4px',
            minHeight: '200px',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            gap: '8px',
            cursor: isUploading ? 'default' : 'pointer',
            transition: 'border-color 0.15s, background-color 0.15s',
            padding: '32px 24px',
          }}
        >
          <svg
            width="36"
            height="36"
            viewBox="0 0 36 36"
            fill="none"
            aria-hidden="true"
          >
            <path
              d="M18 24V12M18 12l-6 6M18 12l6 6"
              stroke="#00602B"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
            <rect
              x="6"
              y="28"
              width="24"
              height="2"
              rx="1"
              fill="#00602B"
              opacity="0.4"
            />
          </svg>
          <p
            style={{
              fontFamily: 'Montserrat, sans-serif',
              fontSize: '15px',
              fontWeight: 500,
              color: '#2C2C2C',
              textAlign: 'center',
            }}
          >
            Drop BESS PDF here or click to browse
          </p>
          <p
            style={{
              fontFamily: 'Open Sans, sans-serif',
              fontSize: '13px',
              color: '#C0C0C0',
            }}
          >
            PDF files only · Max 20 MB
          </p>
        </div>

        <input
          ref={fileInputRef}
          type="file"
          accept="application/pdf"
          className="hidden"
          onChange={handleInputChange}
        />

        {/* File selected indicator */}
        {file && !isUploading && (
          <div
            className="mt-3 flex items-center gap-2"
            style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#2C2C2C' }}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <circle cx="8" cy="8" r="8" fill="#004D22" />
              <path
                d="M5 8l2 2 4-4"
                stroke="white"
                strokeWidth="1.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
            <span className="font-medium">{file.name}</span>
            <span style={{ color: '#C0C0C0' }}>{formatBytes(file.size)}</span>
          </div>
        )}

        {/* Progress bar */}
        {showBar && (
          <div
            className="mt-4"
            style={{
              height: '6px',
              backgroundColor: '#C0C0C0',
              borderRadius: '3px',
              overflow: 'hidden',
            }}
          >
            <div
              style={{
                height: '100%',
                backgroundColor: '#00602B',
                borderRadius: '3px',
                width: phase === 'done' ? '100%' : phase === 'parsing' ? '90%' : `${uploadPct}%`,
                transition: 'width 0.3s ease',
                animation: phase === 'parsing' ? 'pulse 1.5s ease-in-out infinite' : undefined,
              }}
            />
          </div>
        )}

        {/* Status text while processing */}
        {isUploading && (
          <p
            className="mt-2 text-sm"
            style={{ fontFamily: 'Open Sans, sans-serif', color: '#00602B' }}
          >
            {phase === 'uploading' ? 'Uploading…' : 'Parsing PDF with AI…'}
          </p>
        )}

        {/* Error message */}
        {error && (
          <p
            className="mt-3 text-sm"
            style={{ fontFamily: 'Open Sans, sans-serif', color: '#B94040' }}
          >
            {error}
          </p>
        )}

        {/* Upload button */}
        <div className="mt-6">
          <button
            className="btn-primary"
            disabled={!file || isUploading}
            onClick={handleUpload}
          >
            {isUploading ? 'Processing…' : 'Upload & Create Project'}
          </button>
        </div>
      </main>
    </div>
  )
}
