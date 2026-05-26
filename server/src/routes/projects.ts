import { Router, Request, Response } from 'express'
import Anthropic from '@anthropic-ai/sdk'
import { prisma } from '../lib/prisma'
import { requireGIW, GIWRequest } from '../middleware/auth'
import { uploadPDF } from '../middleware/upload'
import { sendReviewInviteByEmail } from '../lib/email'
import { computeItemsBessPoints, computeCurrentBESS } from '../lib/bess'

function computeInteractiveBESS(
  bessScore: number | null,
  items: Array<{ flag: string; creditReference: string; bessPoints?: string | null; additionalBessPoints?: number | null }>,
): number {
  let score = bessScore ?? 0
  for (const item of items) {
    if (item.flag !== 'Yes') continue
    if (item.creditReference === 'Innovation' && item.bessPoints) {
      const raw = Number(item.bessPoints)
      if (!isNaN(raw)) score += Math.round(raw * 0.9 * 10) / 10
    } else if (item.additionalBessPoints != null) {
      score += item.additionalBessPoints
    }
  }
  return Math.round(score)
}
import {
  triggerCommentGeneration,
  triggerDrawingGeneration,
  generateGIWComments,
  generateExcellenceOpportunities,
  applyAutoVisibilityRules,
  INNOVATION_INITIATIVES,
  BLOCKED_INNOVATION_NAMES,
} from '../lib/generate'
import { exportProjectToExcel } from '../lib/export'

// eslint-disable-next-line @typescript-eslint/no-require-imports
const pdfParse = require('pdf-parse') as (
  buffer: Buffer,
) => Promise<{ text: string; numpages: number }>

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

/* ── Hardcoded credit metadata ── */

interface CreditMeta {
  mandatory: boolean
  responsibleParty: string
}

const CREDIT_META: Record<string, CreditMeta> = {
  'Management 1.1': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Management 2.1': { mandatory: true,  responsibleParty: 'ESD Consultant' },
  'Management 2.2': { mandatory: true,  responsibleParty: 'ESD Consultant' },
  'Management 2.3': { mandatory: false, responsibleParty: 'ESD Consultant' },
  'Management 3.1': { mandatory: false, responsibleParty: 'Services / Developer' },
  'Management 3.2': { mandatory: false, responsibleParty: 'Services / Developer' },
  'Management 3.3': { mandatory: false, responsibleParty: 'Services / Developer' },
  'Management 4.1': { mandatory: false, responsibleParty: 'Developer' },
  'IWM 1.1':        { mandatory: true,  responsibleParty: 'Architect / Developer' },
  'IWM 2.1':        { mandatory: true,  responsibleParty: 'Civil / Services / Architect / Developer' },
  'IWM 3.1':        { mandatory: false, responsibleParty: 'Landscape / Architect / Developer' },
  'IWM 4.1':        { mandatory: false, responsibleParty: 'Services / Developer / Architect' },
  'OE 1.1':         { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'OE 1.2':         { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'OE 2.1':         { mandatory: false, responsibleParty: 'ESD Consultant' },
  'OE 2.2':         { mandatory: false, responsibleParty: 'Developer / Services' },
  'OE 2.6':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 2.7':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 3.1':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 3.2':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 3.3':         { mandatory: false, responsibleParty: 'Developer / Services' },
  'OE 3.4':         { mandatory: false, responsibleParty: 'Developer / Architect' },
  'OE 3.5':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 3.6':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 3.7':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 4.1':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 4.2':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 4.4':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'OE 4.5':         { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'IEQ 1.1':        { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'IEQ 1.2':        { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'IEQ 1.3':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'IEQ 1.4':        { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'IEQ 1.5':        { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'IEQ 1.6':        { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'IEQ 2.1':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'IEQ 2.2':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'IEQ 2.3':        { mandatory: true,  responsibleParty: 'Developer / Architect / Services' },
  'IEQ 3.1':        { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'IEQ 3.2':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'IEQ 3.3':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'IEQ 3.4':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'IEQ 3.5':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'IEQ 4.1':        { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.1':  { mandatory: false, responsibleParty: 'Developer / Architect / Traffic' },
  'Transport 1.2':  { mandatory: false, responsibleParty: 'Developer / Architect / Traffic' },
  'Transport 1.3':  { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.4':  { mandatory: false, responsibleParty: 'Developer / Architect / Traffic' },
  'Transport 1.5':  { mandatory: false, responsibleParty: 'Developer / Architect / Traffic' },
  'Transport 1.6':  { mandatory: false, responsibleParty: 'Developer / Architect / Traffic' },
  'Transport 2.1':  { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Transport 2.2':  { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 2.3':  { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Waste 1.1':      { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Waste 2.1':      { mandatory: false, responsibleParty: 'Developer / Architect / Waste Consultant' },
  'Waste 2.2':      { mandatory: false, responsibleParty: 'Developer / Architect / Waste Consultant' },
  'Urban Ecology 1.1': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Urban Ecology 2.1': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 2.2': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 2.3': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 2.4': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Urban Ecology 3.1': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 3.2': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Innovation 1.1': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.2': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.3': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.4': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.5': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.6': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.7': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.8': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Innovation 1.9': { mandatory: false, responsibleParty: 'Developer / Architect' },
}

/**
 * Normalise a creditId extracted from a PDF so it matches the CREDIT_META keys.
 * Handles abbreviations and verbose category names the AI may produce.
 */
function normaliseCreditId(raw: string): string {
  // Extract the X.X number suffix first
  const numMatch = raw.match(/(\d+\.\d+)/)
  if (!numMatch) return raw

  const num = numMatch[1]
  const lower = raw.toLowerCase()

  if (/^(management|mgmt|mgt)\b/.test(lower))                           return `Management ${num}`
  if (/^(iwm\b|integrated water)/.test(lower))                          return `IWM ${num}`
  if (/^(oe\b|operational energy)/.test(lower))                         return `OE ${num}`
  if (/^(ieq\b|indoor environment)/.test(lower))                        return `IEQ ${num}`
  if (/^transport/.test(lower))                                          return `Transport ${num}`
  if (/^(w\b|wrr\b|waste)/.test(lower))                                 return `Waste ${num}`
  if (/^(ue\b|urban ecology|urban)/.test(lower))                        return `Urban Ecology ${num}`
  if (/^innovation/.test(lower))                                         return `Innovation ${num}`

  return raw
}

function getCategoryOrder(creditId: string, category = ''): number {
  const id = creditId.toLowerCase()
  const cat = category.toLowerCase()
  if (id.startsWith('management')  || cat.startsWith('management'))              return 1
  if (id.startsWith('iwm')         || cat.includes('water'))                     return 2
  if (id.startsWith('oe')          || cat.includes('operational energy'))        return 3
  if (id.startsWith('ieq')         || cat.includes('indoor environmental'))      return 4
  if (id.startsWith('transport')   || cat.includes('transport'))                 return 5
  if (id.startsWith('waste')       || cat.includes('waste'))                     return 6
  if (id.startsWith('urban')       || cat.includes('urban'))                     return 7
  if (id.startsWith('innovation')  || cat.includes('innovation'))                return 8
  return 99
}

function extractJSON(raw: string): string {
  // Strip markdown code fences if present
  const block = raw.match(/```(?:json)?\s*([\s\S]*?)```/)
  if (block) return block[1].trim()
  // Find outermost {...} — handles any leading/trailing text
  const start = raw.indexOf('{')
  const end = raw.lastIndexOf('}')
  if (start !== -1 && end !== -1 && end > start) return raw.slice(start, end + 1)
  return raw.trim()
}

/* ── Router ── */

const router = Router()

/* Disciplines that map to a shorter search term for responsibleParty matching */
const DISCIPLINE_SEARCH: Record<string, string> = {
  'Services Engineer': 'Services',
  'Civil Engineer': 'Civil',
  'Landscape Architect': 'Landscape',
}

/** Returns the substring to search in responsibleParty, or null for no filter (show all credits). */
function disciplineSearchTerm(discipline: string): string | null {
  if (discipline === 'Other' || discipline === 'ESD Consultant') return null
  return DISCIPLINE_SEARCH[discipline] ?? discipline
}

// All routes below require GIW auth EXCEPT the credits and excellence endpoints,
// which support both GIW and reviewer access.

/* POST /api/projects/create-from-pdf */
router.post(
  '/create-from-pdf',
  requireGIW,
  uploadPDF.single('pdf'),
  async (req: GIWRequest, res: Response): Promise<void> => {
    const file = (req as GIWRequest & { file?: Express.Multer.File }).file

    if (!file) {
      res.status(400).json({ error: 'No PDF file uploaded' })
      return
    }

    try {
      /* 1. Extract text from PDF */
      const { text: pdfText } = await pdfParse(file.buffer)

      // Trim excessively long PDFs to stay within a safe input token budget
      const trimmedText = pdfText.length > 60000 ? pdfText.slice(0, 60000) + '\n[text truncated]' : pdfText

      /* 2. Two-pass parse:
            Pass A — lightweight: project info + credit statuses/scores (small output).
            Pass B — requirements + rawDataPoints fetched per-credit by generate.ts.
         This keeps Pass A well under the 8192-token output limit. */
      const userPrompt = `Extract from this BESS assessment PDF and return a single JSON object:
{
  "project": {
    "name": "string",
    "address": "string or null",
    "projectId": "string or null",
    "bessScore": number or null,
    "date": "ISO date string or null"
  },
  "credits": [
    {
      "creditId": "use exactly: Management, IWM, OE, IEQ, Transport, Waste, Urban Ecology, or Innovation — followed by the credit number (e.g. Management 1.1, IWM 2.1, OE 3.2, IEQ 1.1, Transport 1.1, Waste 2.1, Urban Ecology 2.1, Innovation 1.1)",
      "creditName": "string",
      "category": "string",
      "creditRequirement": "string or null",
      "creditScore": number or null,
      "creditWeight": number or null,
      "creditStatus": "Y | N | ScopedOut",
      "scopedOutReason": "string or null",
      "rawDataPoints": "string or null"
    }
  ],
  "supportingEvidence": [
    {
      "creditReference": "string",
      "type": "Floor Plan Annotation",
      "requirement": "string"
    }
  ]
}

Rules:
- creditStatus: Achieved → "Y", Not Achieved or 0% → "N", Scoped Out / N/A / Disabled → "ScopedOut"
- EXCLUDE credits with status Disabled from the credits array entirely
- creditRequirement: the specific compliance requirement for this credit as stated in the BESS document (e.g. "Provide a minimum of 24 long-stay bicycle spaces"). 1–2 sentences. Null if not stated.
- rawDataPoints: key numbers and specs only for this credit (e.g. "Required 24 long-stay bicycle spaces + 4 short-stay. Provided 20 long-stay + 2 short-stay." or "WELS 6-star taps, 3-star showers, no rainwater tank. 28% potable water reduction achieved."). 1–2 sentences max. Null if no specific data in PDF for this credit. Exception: for IWM 1.1 (potable water / fixtures), list every fixture type and its WELS star rating on a separate line — do not truncate.
- Return ONLY the JSON object — no markdown, no explanation

BESS text:
${trimmedText}`

      const message = await anthropic.messages.stream({
        model: 'claude-sonnet-4-6',
        max_tokens: 32000,
        system: 'You are a JSON extraction tool. Output only a valid JSON object. No markdown fences, no commentary.',
        messages: [{ role: 'user', content: userPrompt }],
      }).finalMessage()

      const rawContent = message.content[0]
      if (rawContent.type !== 'text') {
        res.status(422).json({ error: 'Unexpected response from AI parser' })
        return
      }

      const jsonText = rawContent.text
      console.log('[create-from-pdf] stop_reason:', message.stop_reason, '— response length:', jsonText.length)

      let parsed: {
        project: {
          name: string
          address?: string
          projectId?: string
          bessScore?: number
          date?: string
        }
        credits: Array<{
          creditId: string
          creditName: string
          category: string
          creditRequirement?: string
          creditScore?: number
          creditWeight?: number
          creditStatus: string
          scopedOutReason?: string
          rawDataPoints?: string
        }>
        supportingEvidence: Array<{
          creditReference: string
          type: string
          requirement: string
        }>
      }

      try {
        parsed = JSON.parse(extractJSON(jsonText))
      } catch (parseErr) {
        console.error('[create-from-pdf] JSON parse failed. stop_reason:', message.stop_reason)
        console.error('[create-from-pdf] Raw response:\n', jsonText.slice(0, 1200))
        res.status(422).json({ error: 'AI parser returned invalid JSON. Try uploading again.' })
        return
      }

      /* 3. Detect existing project by address → auto-assign revision */
      const extractedAddress = parsed.project.address?.trim() || null
      const normalizeAddr = (a: string) => a.toLowerCase().replace(/\s+/g, ' ').trim()

      let parentProjectId: string | null = null
      let revisionLabel = 'A'

      if (extractedAddress) {
        const rootProjects = await prisma.project.findMany({
          where: { parentProjectId: null, address: { not: null } },
          select: { id: true, address: true, revisions: { select: { id: true } } },
        })
        const match = rootProjects.find(
          (p) => p.address && normalizeAddr(p.address) === normalizeAddr(extractedAddress),
        )
        if (match) {
          parentProjectId = match.id
          const nextIndex = 1 + match.revisions.length
          revisionLabel = String.fromCharCode(65 + nextIndex)
        }
      }

      /* 4. Persist to database */
      const project = await prisma.project.create({
        data: {
          name: parsed.project.name?.trim() || file.originalname,
          address: extractedAddress,
          bessScore:
            parsed.project.bessScore != null
              ? parseFloat(String(parsed.project.bessScore))
              : null,
          date: parsed.project.date ? new Date(parsed.project.date) : null,
          projectId: parsed.project.projectId?.trim() || null,
          revision: revisionLabel,
          parentProjectId,
          generationStatus: 'running',
        },
      })

      /* 4. Create credits */
      const creditData = (parsed.credits ?? []).map((c) => {
        const meta = CREDIT_META[normaliseCreditId(c.creditId)] ?? CREDIT_META[c.creditId] ?? {
          mandatory: false,
          responsibleParty: null,
        }
        const normalisedCreditId = normaliseCreditId(c.creditId)
        return {
          projectId: project.id,
          creditId: normalisedCreditId,
          creditName: c.creditName,
          category: c.category,
          categoryOrder: getCategoryOrder(c.creditId, c.category),
          creditRequirement: c.creditRequirement?.trim() ?? null,
          creditScore: c.creditScore != null ? parseFloat(String(c.creditScore)) : null,
          creditWeight: c.creditWeight != null ? parseFloat(String(c.creditWeight)) : null,
          creditStatus: c.creditStatus,
          scopedOutReason: c.scopedOutReason ?? null,
          rawDataPoints: c.rawDataPoints?.trim() ?? null,
          mandatory: meta.mandatory,
          responsibleParty: meta.responsibleParty ?? null,
        }
      })

      await prisma.credit.createMany({ data: creditData })


      /* 5. Create drawing requirements — floor plan annotations for achieved credits only */
      const DRAWING_EXCLUDED_CREDITS = new Set(['ieq 1.1', 'ieq 1.2', 'ieq 2.1', 'ieq 3.1', 'management 2.3', 'iwm 2.1'])
      const achievedCreditIds = new Set(
        (parsed.credits ?? [])
          .filter((c) => c.creditStatus === 'Y')
          .map((c) => c.creditId.trim().toLowerCase()),
      )
      const drawingData = (parsed.supportingEvidence ?? [])
        .filter((e) => {
          const ref = e.creditReference.trim().toLowerCase()
          return (
            e.type !== 'Supporting Document' &&
            achievedCreditIds.has(ref) &&
            !DRAWING_EXCLUDED_CREDITS.has(ref)
          )
        })
        .map((e) => ({
          projectId: project.id,
          creditReference: e.creditReference,
          drawingType: e.type,
          requirement: e.requirement,
        }))

      if (drawingData.length > 0) {
        await prisma.drawingRequirement.createMany({ data: drawingData })
      }

      /* 6. Trigger stub generation jobs */
      triggerCommentGeneration(project.id).catch(console.error)
      triggerDrawingGeneration(project.id).catch(console.error)

      res.status(201).json({ projectId: project.id })
    } catch (err: unknown) {
      const e = err as Error
      console.error('[create-from-pdf] FAILED:', e?.message ?? err)
      console.error('[create-from-pdf] Stack:', e?.stack?.split('\n')[1])
      res.status(500).json({ error: `Failed to process PDF: ${e?.message ?? 'unknown error'}` })
    }
  },
)

/* GET /api/projects */
router.get('/', requireGIW, async (_req: GIWRequest, res: Response): Promise<void> => {
  try {
    const projects = await prisma.project.findMany({
      where: { parentProjectId: null },
      orderBy: { createdAt: 'desc' },
      include: {
        reviewers: { select: { id: true, hasSubmitted: true } },
        revisions: {
          orderBy: { createdAt: 'asc' },
          select: { id: true, revision: true, createdAt: true, bessScore: true, reviewLinkToken: true },
        },
      },
    })
    res.json(projects)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to fetch projects' })
  }
})

/* POST /api/projects */
router.post('/', requireGIW, async (req: GIWRequest, res: Response): Promise<void> => {
  try {
    const { name, address, bessScore, date, revision, projectId } =
      req.body as {
        name: string
        address?: string
        bessScore?: number | string
        date?: string
        revision?: string
        projectId?: string
      }

    if (!name?.trim()) {
      res.status(400).json({ error: 'Project name is required' })
      return
    }

    const project = await prisma.project.create({
      data: {
        name: name.trim(),
        address: address?.trim() || null,
        bessScore:
          bessScore !== undefined && bessScore !== ''
            ? parseFloat(String(bessScore))
            : null,
        date: date ? new Date(date) : null,
        revision: revision?.trim() || null,
        projectId: projectId?.trim() || null,
      },
    })

    res.status(201).json(project)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to create project' })
  }
})

/* GET /api/projects/:id */
router.get('/:id', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    // Apply auto-visibility rules so existing projects stay up to date
    await applyAutoVisibilityRules(req.params.id).catch(console.error)

    const project = await prisma.project.findUnique({
      where: { id: req.params.id },
      include: {
        credits: {
          where: { deletedByGIW: false },
          orderBy: [{ categoryOrder: 'asc' }, { creditId: 'asc' }],
        },
        reviewers: true,
        drawingItems: true,
        excellenceItems: { where: { deletedByGIW: false } },
      },
    })

    if (!project) {
      res.status(404).json({ error: 'Project not found' })
      return
    }

    // Build revision family: root + all its children, ordered by creation date
    const rootId = project.parentProjectId ?? project.id
    const revisionFamily = await prisma.project.findMany({
      where: { OR: [{ id: rootId }, { parentProjectId: rootId }] },
      orderBy: { createdAt: 'asc' },
      select: { id: true, revision: true, createdAt: true, bessScore: true },
    })

    res.json({ ...project, revisionFamily })
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to fetch project' })
  }
})

/* PATCH /api/projects/:id */
router.patch('/:id', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    const { name, address, bessScore, date, revision, projectId, generationStatus, notifyEmail, gdft } =
      req.body as Partial<{
        name: string
        address: string
        bessScore: number
        date: string
        revision: string
        projectId: string
        generationStatus: string
        notifyEmail: string | null
        gdft: boolean
      }>

    const project = await prisma.project.update({
      where: { id: req.params.id },
      data: {
        ...(name !== undefined && { name }),
        ...(address !== undefined && { address }),
        ...(bessScore !== undefined && { bessScore: parseFloat(String(bessScore)) }),
        ...(date !== undefined && { date: date ? new Date(date) : null }),
        ...(revision !== undefined && { revision }),
        ...(projectId !== undefined && { projectId }),
        ...(generationStatus !== undefined && { generationStatus }),
        ...(notifyEmail !== undefined && { notifyEmail: notifyEmail || null }),
        ...(gdft !== undefined && { gdft }),
      },
    })

    res.json(project)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to update project' })
  }
})

/* POST /api/projects/:id/invite */
router.post('/:id/invite', requireGIW, async (req: Request, res: Response): Promise<void> => {
  const { email, discipline, name } = req.body as { email?: string; discipline?: string; name?: string }
  if (!email?.trim() || !discipline?.trim()) {
    res.status(400).json({ error: 'email and discipline are required' })
    return
  }
  const normalEmail = email.trim().toLowerCase()
  try {
    const project = await prisma.project.findUniqueOrThrow({
      where: { id: req.params.id },
      select: { id: true, name: true, address: true },
    })
    const inviteToken = crypto.randomUUID()
    // Upsert: update existing reviewer (by email+project) or create new one
    const existing = await prisma.reviewer.findFirst({
      where: { projectId: project.id, email: normalEmail },
    })
    let reviewer
    if (existing) {
      reviewer = await prisma.reviewer.update({
        where: { id: existing.id },
        data: { discipline: discipline.trim(), inviteToken },
      })
    } else {
      reviewer = await prisma.reviewer.create({
        data: {
          projectId: project.id,
          email: normalEmail,
          discipline: discipline.trim(),
          inviteToken,
        },
      })
    }
    const base = process.env.BASE_URL || 'http://localhost:5173'
    const inviteLink = `${base}/review/invite/${inviteToken}`
    // Respond immediately — email is sent in the background so the UI doesn't wait on SMTP
    res.json({ ok: true, reviewerId: reviewer.id })
    sendReviewInviteByEmail(normalEmail, inviteLink, project.name, discipline.trim(), name?.trim() || null, project.address || null)
      .catch(err => console.error('[invite] email send failed:', err))
  } catch (err) {
    console.error('[invite] error:', err)
    res.status(500).json({ error: 'Failed to send invite' })
  }
})

/* DELETE /api/projects/:id */
router.delete('/:id', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    await prisma.project.delete({ where: { id: req.params.id } })
    res.status(204).end()
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to delete project' })
  }
})

/* POST /api/projects/:id/export */
router.post('/:id/export', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    const project = await prisma.project.findUnique({
      where: { id: req.params.id },
      select: { id: true, name: true, date: true },
    })
    if (!project) {
      res.status(404).json({ error: 'Project not found' })
      return
    }

    const buffer = await exportProjectToExcel(project.id)
    const safeName = project.name.replace(/[^a-zA-Z0-9\s-]/g, '').trim().replace(/\s+/g, '-')
    const dateStr = (project.date ?? new Date()).toISOString().slice(0, 10)
    const filename = `GIW-ESD-Review-${safeName}-${dateStr}.xlsx`

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': buffer.length,
    })
    res.send(buffer)
  } catch (err) {
    console.error('[export]', err)
    res.status(500).json({ error: 'Failed to generate export' })
  }
})

/* POST /api/projects/:id/recalculate-bess */
router.post('/:id/recalculate-bess', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    const exists = await prisma.project.findUnique({ where: { id: req.params.id }, select: { id: true } })
    if (!exists) { res.status(404).json({ error: 'Project not found' }); return }
    const { recalculateBessScore } = await import('../lib/bess')
    const score = await recalculateBessScore(req.params.id)
    res.json({ bessScore: score })
  } catch (err) {
    console.error('[recalculate-bess]', err)
    res.status(500).json({ error: 'Failed to recalculate BESS score' })
  }
})

/* GET /api/projects/:id/changelog */
router.get('/:id/changelog', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    const project = await prisma.project.findUnique({
      where: { id: req.params.id },
      select: { id: true, name: true, address: true, revision: true, bessScore: true, parentProjectId: true },
    })
    if (!project) { res.status(404).json({ error: 'Project not found' }); return }
    if (!project.parentProjectId) { res.status(400).json({ error: 'This is the first revision — no previous revision to compare against' }); return }

    // Find all revisions of this family ordered by revision label
    const family = await prisma.project.findMany({
      where: {
        OR: [{ id: project.parentProjectId }, { parentProjectId: project.parentProjectId }],
      },
      select: { id: true, revision: true, bessScore: true, createdAt: true },
      orderBy: { revision: 'asc' },
    })

    const currentIdx = family.findIndex(r => r.id === project.id)
    if (currentIdx <= 0) { res.status(400).json({ error: 'Could not determine previous revision' }); return }
    const prevRevisionMeta = family[currentIdx - 1]

    // Fetch credits for both revisions
    const [prevCredits, newCredits] = await Promise.all([
      prisma.credit.findMany({
        where: { projectId: prevRevisionMeta.id, deletedByGIW: false },
        select: { creditId: true, creditName: true, category: true, categoryOrder: true, creditStatus: true, creditScore: true },
      }),
      prisma.credit.findMany({
        where: { projectId: project.id, deletedByGIW: false },
        select: { creditId: true, creditName: true, category: true, categoryOrder: true, creditStatus: true, creditScore: true },
      }),
    ])

    const { buildChangelogDocx } = await import('../lib/changelog')
    const buffer = await buildChangelogDocx({
      projectName: project.name,
      address: project.address ?? null,
      prevRevision: prevRevisionMeta.revision ?? '?',
      newRevision: project.revision ?? '?',
      prevBessScore: prevRevisionMeta.bessScore ?? null,
      newBessScore: project.bessScore ?? null,
      prevCredits,
      newCredits,
    })

    const safeName = project.name.replace(/[^a-zA-Z0-9\s-]/g, '').trim().replace(/\s+/g, '-')
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="Changelog-${safeName}-Rev${prevRevisionMeta.revision}-to-Rev${project.revision}.docx"`,
      'Content-Length': buffer.length,
    })
    res.send(buffer)
  } catch (err) {
    console.error('[changelog]', err)
    res.status(500).json({ error: 'Failed to generate changelog' })
  }
})

/* POST /api/projects/:id/generate */
router.post('/:id/generate', requireGIW, async (req: Request, res: Response): Promise<void> => {
  const projectId = req.params.id
  try {
    const exists = await prisma.project.findUnique({
      where: { id: projectId },
      select: { id: true },
    })
    if (!exists) {
      res.status(404).json({ error: 'Project not found' })
      return
    }

    await prisma.project.update({
      where: { id: projectId },
      data: { generationStatus: 'running' },
    })

    // Return 202 immediately — generation runs async
    res.status(202).json({ status: 'running' })

    triggerCommentGeneration(projectId).catch(console.error)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to start generation' })
  }
})

/* GET /api/projects/:id/generation-status */
router.get('/:id/generation-status', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    const project = await prisma.project.findUnique({
      where: { id: req.params.id },
      select: { generationStatus: true },
    })
    if (!project) {
      res.status(404).json({ error: 'Project not found' })
      return
    }
    res.json({ status: project.generationStatus })
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to fetch generation status' })
  }
})

/* GET /api/projects/:id/credits
   GIW session  → all credits with scores.
   Reviewer     → discipline-filtered, creditScore and creditWeight omitted.
   Query params for reviewer: reviewerEmail, reviewerDiscipline */
router.get('/:id/credits', async (req: Request, res: Response): Promise<void> => {
  const { reviewerEmail, reviewerDiscipline } = req.query as Record<string, string | undefined>
  // Reviewer mode when query params are present; otherwise GIW/admin mode
  const giw = !reviewerEmail || !reviewerDiscipline

  try {
    // Always apply auto-visibility rules before serving credits
    await applyAutoVisibilityRules(req.params.id).catch(console.error)

    if (giw) {
      const credits = await prisma.credit.findMany({
        where: { projectId: req.params.id, deletedByGIW: false, hiddenFromPortal: false },
        orderBy: [{ categoryOrder: 'asc' }, { creditId: 'asc' }],
        include: {
          comments: { orderBy: { submittedAt: 'asc' } },
        },
      })
      res.json(credits)
      return
    }

    // Reviewer path
    const searchTerm = disciplineSearchTerm(reviewerDiscipline!)
    const credits = await prisma.credit.findMany({
      where: {
        projectId: req.params.id,
        deletedByGIW: false,
        hiddenFromPortal: false,
        // Scoped-out and not-achieved credits are never shown to reviewers
        creditStatus: { notIn: ['ScopedOut', 'N'] },
        ...(searchTerm
          ? { responsibleParty: { contains: searchTerm, mode: 'insensitive' } }
          : {}),
      },
      orderBy: [{ categoryOrder: 'asc' }, { creditId: 'asc' }],
      select: {
        id: true,
        projectId: true,
        category: true,
        categoryOrder: true,
        creditId: true,
        creditName: true,
        creditRequirement: true,
        mandatory: true,
        responsibleParty: true,
        creditStatus: true,
        commentsGIW: true,
        scopedOutReason: true,
        rawDataPoints: true,
        lastEditedBy: true,
        lastEditedAt: true,
        // creditScore and creditWeight intentionally excluded for reviewers
        comments: {
          orderBy: { submittedAt: 'asc' as const },
        },
      },
    })
    res.json(credits)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to fetch credits' })
  }
})

/* GET /api/projects/:id/drawing-requirements — reviewer or GIW */
router.get('/:id/drawing-requirements', async (req: Request, res: Response): Promise<void> => {
  try {
    const exists = await prisma.project.findUnique({
      where: { id: req.params.id },
      select: { id: true },
    })
    if (!exists) {
      res.status(404).json({ error: 'Project not found' })
      return
    }
    const drawings = await prisma.drawingRequirement.findMany({
      where: {
        projectId: req.params.id,
        creditReference: { notIn: ['IEQ 1.1', 'IEQ 1.2', 'IEQ 2.1', 'IEQ 3.1', 'Management 2.3', 'IWM 2.1'] },
      },
      orderBy: { creditReference: 'asc' },
    })
    res.json(drawings)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to fetch drawing requirements' })
  }
})

/* GET /api/projects/:id/excellence
   GIW session  → all non-deleted items.
   Reviewer     → discipline-filtered (via linked credit responsibleParty). */
router.get('/:id/excellence', async (req: Request, res: Response): Promise<void> => {
  const { reviewerEmail, reviewerDiscipline } = req.query as Record<string, string | undefined>
  const giw = !reviewerEmail || !reviewerDiscipline

  try {
    if (giw) {
      // Remove any blocked innovation items that may exist in the DB (case-insensitive match)
      const toDelete = await prisma.eSDExcellenceOpportunity.findMany({
        where: { projectId: req.params.id, creditReference: 'Innovation' },
        select: { id: true, creditName: true },
      })
      const blockedIds = toDelete
        .filter((r) => BLOCKED_INNOVATION_NAMES.has(r.creditName.toLowerCase()))
        .map((r) => r.id)
      if (blockedIds.length > 0) {
        await prisma.eSDExcellenceOpportunity.deleteMany({ where: { id: { in: blockedIds } } })
      }

      // Also remove blocked innovation line items from the credits table
      const blockedCredits = await prisma.credit.findMany({
        where: { projectId: req.params.id, creditId: 'Innovation' },
        select: { id: true, creditName: true },
      })
      const blockedCreditIds = blockedCredits
        .filter((r) => BLOCKED_INNOVATION_NAMES.has(r.creditName.toLowerCase()))
        .map((r) => r.id)
      if (blockedCreditIds.length > 0) {
        await prisma.credit.deleteMany({ where: { id: { in: blockedCreditIds } } })
      }

      // Auto-seed innovation items for projects that predate the innovation cards feature
      const existingInnovation = await prisma.eSDExcellenceOpportunity.count({
        where: { projectId: req.params.id, creditReference: 'Innovation' },
      })
      if (existingInnovation === 0) {
        await prisma.eSDExcellenceOpportunity.createMany({
          data: INNOVATION_INITIATIVES.map((i) => ({
            projectId: req.params.id,
            creditReference: 'Innovation',
            creditName: i.name,
            improvementDescription: i.desc || null,
            bessPoints: i.pts,
          })),
        })
      }

      const items = await prisma.eSDExcellenceOpportunity.findMany({
        where: { projectId: req.params.id, deletedByGIW: false },
        orderBy: { creditReference: 'asc' },
      })

      // Compute additionalBessPoints for non-innovation items.
      // Weight stats use only non-hidden credits (matching the Excel export BESS formula).
      // Credit lookup uses all non-deleted credits so items linked to hidden credits still resolve.
      const creditSelect = { id: true, creditId: true, category: true, creditStatus: true, creditWeight: true, creditScore: true } as const
      const [weightCredits, allCredits, projectRow] = await Promise.all([
        prisma.credit.findMany({ where: { projectId: req.params.id, deletedByGIW: false, creditId: { not: 'Innovation' }, OR: [{ hiddenFromPortal: false }, { category: { contains: 'innovation', mode: 'insensitive' } }] }, select: creditSelect }),
        prisma.credit.findMany({ where: { projectId: req.params.id, deletedByGIW: false }, select: creditSelect }),
        prisma.project.findUnique({ where: { id: req.params.id }, select: { bessScore: true } }),
      ])

      const itemsWithPoints = computeItemsBessPoints(items, weightCredits, allCredits)
      const currentBESS = computeCurrentBESS(weightCredits)
      const interactiveBESS = computeInteractiveBESS(projectRow?.bessScore ?? null, itemsWithPoints)

      res.json({ items: itemsWithPoints, computedBESS: currentBESS, interactiveBESS })
      return
    }

    // Reviewer path: find eligible creditIds first, then filter excellence items
    const searchTerm = disciplineSearchTerm(reviewerDiscipline!)
    const matchingCredits = await prisma.credit.findMany({
      where: {
        projectId: req.params.id,
        ...(searchTerm
          ? { responsibleParty: { contains: searchTerm, mode: 'insensitive' } }
          : {}),
      },
      select: { id: true },
    })
    const creditIds = matchingCredits.map((c) => c.id)

    const creditSelect = { id: true, creditId: true, category: true, creditStatus: true, creditWeight: true, creditScore: true } as const
    const allItemsSelect = { flag: true, creditReference: true, bessPoints: true, creditId: true, currentScore: true } as const
    const [items, weightCredits, allCredits, allItems, projectRow] = await Promise.all([
      prisma.eSDExcellenceOpportunity.findMany({
        where: {
          projectId: req.params.id,
          deletedByGIW: false,
          OR: [
            { creditId: { in: creditIds } },
            { creditReference: 'Innovation' },
          ],
        },
        orderBy: { creditReference: 'asc' },
      }),
      prisma.credit.findMany({ where: { projectId: req.params.id, deletedByGIW: false, creditId: { not: 'Innovation' }, OR: [{ hiddenFromPortal: false }, { category: { contains: 'innovation', mode: 'insensitive' } }] }, select: creditSelect }),
      prisma.credit.findMany({ where: { projectId: req.params.id, deletedByGIW: false }, select: creditSelect }),
      prisma.eSDExcellenceOpportunity.findMany({ where: { projectId: req.params.id, deletedByGIW: false }, select: allItemsSelect }),
      prisma.project.findUnique({ where: { id: req.params.id }, select: { bessScore: true } }),
    ])
    const filteredWithPoints = computeItemsBessPoints(items, weightCredits, allCredits)
    const currentBESS = computeCurrentBESS(weightCredits)
    const allWithPoints = computeItemsBessPoints(
      allItems.map(i => ({ ...i, id: '' })),
      weightCredits,
      allCredits,
    )
    const interactiveBESS = computeInteractiveBESS(projectRow?.bessScore ?? null, allWithPoints)
    res.json({ items: filteredWithPoints, computedBESS: currentBESS, interactiveBESS })
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: 'Failed to fetch excellence opportunities' })
  }
})

/* DELETE /api/projects/:id/excellence — soft-delete all excellence items for a project (GIW only) */
router.delete('/:id/excellence', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    const { count } = await prisma.eSDExcellenceOpportunity.updateMany({
      where: { projectId: req.params.id, deletedByGIW: false },
      data: { deletedByGIW: true },
    })
    res.json({ ok: true, count })
  } catch (err) {
    console.error('[excellence] delete-all error:', err)
    res.status(500).json({ error: 'Failed to remove all items' })
  }
})

export default router
