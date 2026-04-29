export interface RevisionSummary {
  id: string
  revision: string | null
  createdAt: string
  bessScore: number | null
  reviewLinkToken: string
}

export interface Project {
  id: string
  name: string
  address?: string | null
  bessScore?: number | null
  date?: string | null
  revision?: string | null
  projectId?: string | null
  reviewLinkToken: string
  generationStatus: string
  notifyEmail?: string | null
  gdft?: boolean
  createdAt: string
  parentProjectId?: string | null
  reviewers?: ReviewerSummary[]
  revisions?: RevisionSummary[]       // on list: child revisions of this root
  revisionFamily?: RevisionSummary[]  // on detail: all revisions incl. self
  credits?: Credit[]
  drawingItems?: DrawingRequirement[]
  excellenceItems?: ESDExcellenceOpportunity[]
}

export interface ReviewerSummary {
  id: string
  hasSubmitted: boolean
  email?: string
  discipline?: string
  submittedAt?: string | null
}

export interface Reviewer {
  id: string
  projectId: string
  email: string
  discipline: string
  hasSubmitted: boolean
  submittedAt?: string | null
}

export interface Credit {
  id: string
  projectId: string
  category: string
  categoryOrder: number
  creditId: string
  creditName: string
  creditRequirement?: string | null
  mandatory: boolean
  responsibleParty?: string | null
  creditStatus: string
  creditScore?: number | null
  creditWeight?: number | null
  commentsGIW?: string | null
  scopedOutReason?: string | null
  rawDataPoints?: string | null
  lastEditedBy?: string | null
  lastEditedAt?: string | null
  deletedByGIW?: boolean
  hiddenFromPortal?: boolean
  /** GIW: all reviewer comments. Reviewer: their own comment only. */
  comments?: CreditComment[]
}

export interface CreditComment {
  id: string
  creditId: string
  projectId: string
  reviewerEmail: string
  reviewerDiscipline: string
  commentText: string
  submittedAt: string
}

export interface DrawingRequirement {
  id: string
  projectId: string
  creditReference: string
  drawingType: string
  requirement: string
  discipline?: string | null
  status: string
  notes?: string | null
}

export interface ESDExcellenceOpportunity {
  id: string
  projectId: string
  creditId?: string | null
  creditReference: string
  creditName: string
  currentScore?: number | null
  improvementDescription?: string | null
  flag: string
  flaggedBy?: string | null
  flaggedAt?: string | null
  reviewerNotes?: string | null
  bessPoints?: string | null
  additionalBessPoints?: number | null
  deletedByGIW: boolean
}

export interface AuthUser {
  email: string
  isGIW: boolean
}

export interface ReviewerSession {
  reviewerEmail: string
  reviewerDiscipline: string
  reviewerId: string
  projectId: string
  reviewLinkToken: string
  createdAt: string // ISO date
}
