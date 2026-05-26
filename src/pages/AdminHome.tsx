import { useState, useEffect, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import axios from 'axios'
import NavBar from '@/components/NavBar'
import type { Project, ReviewerSummary, RevisionSummary } from '@/types'

export default function AdminHome() {
  const [projects, setProjects] = useState<Project[]>([])
  const [loading, setLoading] = useState(true)
  const [slowApi, setSlowApi] = useState(false)
  const navigate = useNavigate()
  const slowTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  useEffect(() => {
    document.title = 'ESD Review Portal | GIW Environmental Solutions'
  }, [])

  useEffect(() => {
    slowTimer.current = setTimeout(() => setSlowApi(true), 4000)
    axios
      .get<Project[]>('/api/projects', { withCredentials: true })
      .then((res) => setProjects(res.data))
      .catch(console.error)
      .finally(() => {
        if (slowTimer.current) clearTimeout(slowTimer.current)
        setSlowApi(false)
        setLoading(false)
      })
    return () => { if (slowTimer.current) clearTimeout(slowTimer.current) }
  }, [])

  const copyReviewLink = (token: string) => {
    const url = `${window.location.origin}${import.meta.env.BASE_URL}review/${token}`
    navigator.clipboard
      .writeText(url)
      .then(() => alert('Review link copied to clipboard.'))
      .catch(() => alert('Could not copy link — please copy manually:\n' + url))
  }

  const deleteProject = async (id: string, name: string) => {
    if (!window.confirm(`Delete "${name}"? This cannot be undone.`)) return
    try {
      await axios.delete(`/api/projects/${id}`, { withCredentials: true })
      setProjects((prev) => prev.filter((p) => p.id !== id))
    } catch {
      alert('Failed to delete project. Please try again.')
    }
  }

  return (
    <div className="min-h-screen" style={{ backgroundColor: '#FFFFFF' }}>
      <NavBar />

      {slowApi && (
        <div style={{
          backgroundColor: '#FFF8E1', borderBottom: '1px solid #FFE082',
          padding: '10px 24px', textAlign: 'center',
          fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#6B5900',
        }}>
          The server is waking up — this may take up to 60 seconds on first load.
        </div>
      )}

      <main className="max-w-7xl mx-auto px-6 py-10">
        {/* Page header */}
        <div className="flex items-start justify-between mb-8">
          <div>
            <h1 style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '28px', fontWeight: 600, color: '#2C2C2C', lineHeight: 1.2 }}>
              Projects
            </h1>
            <p className="mt-1" style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#C0C0C0' }}>
              Manage ESD review projects and reviewer submissions.
            </p>
          </div>
          <button onClick={() => navigate('/admin/projects/new')} className="btn-primary">
            + New Project
          </button>
        </div>

        {loading ? (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
            {Array.from({ length: 6 }).map((_, i) => (
              <ProjectCardSkeleton key={i} />
            ))}
          </div>
        ) : projects.length === 0 ? (
          <div className="giw-card text-center py-20">
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '16px' }}>
              <svg width="40" height="40" viewBox="0 0 40 40" fill="none" opacity="0.3">
                <rect x="4" y="8" width="32" height="26" rx="2" stroke="#00602B" strokeWidth="2" />
                <path d="M4 14h32" stroke="#00602B" strokeWidth="2" />
                <path d="M12 20h16M12 26h10" stroke="#00602B" strokeWidth="2" strokeLinecap="round" />
              </svg>
            </div>
            <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '14px', color: '#C0C0C0' }}>
              No projects yet. Create your first project to get started.
            </p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
            {projects.map((project) => {
              // Latest revision = last child, or the root itself if no children
              const children = project.revisions ?? []
              const latest = children.length > 0 ? children[children.length - 1] : null
              const latestId = latest?.id ?? project.id
              const latestToken = latest?.reviewLinkToken ?? project.reviewLinkToken
              return (
                <ProjectCard
                  key={project.id}
                  project={project}
                  onOpenProject={() => navigate(`/admin/projects/${latestId}`)}
                  onOpenRevision={(id) => navigate(`/admin/projects/${id}`)}
                  onCopyLink={() => copyReviewLink(latestToken)}
                  onDelete={() => deleteProject(project.id, project.name)}
                />
              )
            })}
          </div>
        )}
      </main>
    </div>
  )
}

/* ── Skeleton card ── */

function ProjectCardSkeleton() {
  return (
    <div className="giw-card flex flex-col gap-3">
      <div className="flex items-start justify-between gap-3">
        <div className="skeleton-shimmer flex-1" style={{ height: 18, maxWidth: '70%' }} />
        <div className="skeleton-shimmer" style={{ height: 32, width: 56, borderRadius: '2px' }} />
      </div>
      <div className="skeleton-shimmer" style={{ height: 13, width: '55%' }} />
      <div className="skeleton-shimmer" style={{ height: 12, width: '40%' }} />
      <div className="skeleton-shimmer" style={{ height: 12, width: '30%' }} />
      <div className="flex gap-2 mt-auto pt-2">
        <div className="skeleton-shimmer flex-1" style={{ height: 34, borderRadius: '2px' }} />
        <div className="skeleton-shimmer flex-1" style={{ height: 34, borderRadius: '2px' }} />
      </div>
    </div>
  )
}

/* ── Project card ── */

interface ProjectCardProps {
  project: Project
  onOpenProject: () => void
  onOpenRevision: (id: string) => void
  onCopyLink: () => void
  onDelete: () => void
}

function ProjectCard({ project, onOpenProject, onOpenRevision, onCopyLink, onDelete }: ProjectCardProps) {
  const [hovered, setHovered] = useState(false)
  const [gdft, setGdft] = useState(project.gdft ?? false)
  const reviewers = (project.reviewers ?? []) as ReviewerSummary[]
  const submitted = reviewers.filter((r) => r.hasSubmitted).length
  const total = reviewers.length

  async function handleGdftToggle(e: React.ChangeEvent<HTMLInputElement>) {
    e.stopPropagation()
    const val = e.target.checked
    setGdft(val)
    try {
      await axios.patch(`/api/projects/${project.id}`, { gdft: val }, { withCredentials: true })
    } catch {
      setGdft(!val)
    }
  }

  const formattedDate = project.date
    ? new Date(project.date).toLocaleDateString('en-AU', {
        day: 'numeric',
        month: 'short',
        year: 'numeric',
      })
    : null

  return (
    <div
      className="giw-card flex flex-col gap-3"
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        border: `1px solid ${hovered ? '#00602B' : '#C0C0C0'}`,
        transition: 'border-color 0.15s',
        cursor: 'default',
      }}
    >
      {/* Name + BESS badge */}
      <div className="flex items-start justify-between gap-3">
        <h2
          style={{
            fontFamily: 'Montserrat, sans-serif', fontSize: '18px', fontWeight: 600,
            color: '#2C2C2C', lineHeight: 1.3,
          }}
        >
          {project.name}
        </h2>

        {project.bessScore != null && (
          <div
            style={{
              backgroundColor: '#00602B', borderRadius: '2px',
              padding: '4px 12px', textAlign: 'center', flexShrink: 0,
            }}
          >
            <p style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '20px', fontWeight: 600, color: '#fff', lineHeight: 1.1 }}>
              {project.bessScore}%
            </p>
            <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '10px', color: 'rgba(255,255,255,0.8)' }}>
              BESS
            </p>
          </div>
        )}
      </div>

      {/* Address */}
      {project.address && (
        <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '13px', color: '#C0C0C0', lineHeight: 1.4 }}>
          {project.address}
        </p>
      )}

      {/* Date */}
      {formattedDate && (
        <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#C0C0C0' }}>
          {formattedDate}
        </p>
      )}

      {/* Revision pills — root (Rev A) + all children */}
      {(() => {
        const children: RevisionSummary[] = project.revisions ?? []
        if (children.length === 0 && !project.revision) return null
        const allRevs: Array<{ id: string; label: string }> = [
          { id: project.id, label: project.revision ?? 'A' },
          ...children.map((r) => ({ id: r.id, label: r.revision ?? '?' })),
        ]
        if (allRevs.length <= 1) return null
        return (
          <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
            {allRevs.map(({ id, label }) => (
              <button
                key={id}
                onClick={(e) => { e.stopPropagation(); onOpenRevision(id) }}
                style={{
                  fontFamily: 'Montserrat, sans-serif', fontSize: '11px', fontWeight: 600,
                  padding: '2px 8px', borderRadius: '3px', border: '1px solid #00602B',
                  backgroundColor: '#fff', color: '#00602B', cursor: 'pointer',
                }}
                title={`Open revision ${label}`}
              >
                Rev {label}
              </button>
            ))}
          </div>
        )
      })()}

      {/* Reviewer count + GDFT */}
      <div className="flex items-center justify-between">
        <p style={{ fontFamily: 'Open Sans, sans-serif', fontSize: '12px', color: '#C0C0C0' }}>
          {total === 0
            ? 'No reviewers yet'
            : `${submitted} of ${total} reviewer${total !== 1 ? 's' : ''} submitted`}
        </p>
        <label
          onClick={(e) => e.stopPropagation()}
          style={{ display: 'flex', alignItems: 'center', gap: '5px', cursor: 'pointer', userSelect: 'none' }}
        >
          <input
            type="checkbox"
            checked={gdft}
            onChange={handleGdftToggle}
            style={{ accentColor: '#00602B', width: '14px', height: '14px', cursor: 'pointer' }}
          />
          <span style={{ fontFamily: 'Montserrat, sans-serif', fontSize: '11px', fontWeight: 600, color: gdft ? '#00602B' : '#8C8C8C' }}>
            GDFT
          </span>
        </label>
      </div>

      {/* Actions */}
      <div className="flex gap-2 mt-auto pt-2">
        <button onClick={onOpenProject} className="btn-primary flex-1">
          Open Project
        </button>
        <button onClick={onCopyLink} className="btn-secondary flex-1">
          Copy Link
        </button>
        <button
          onClick={onDelete}
          className="btn-danger"
          style={{ padding: '0 12px' }}
          title="Delete project"
        >
          <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
            <path d="M1 3h12M5 3V2h4v1M3 3l.9 9h6.2L11 3" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
          </svg>
        </button>
      </div>
    </div>
  )
}
