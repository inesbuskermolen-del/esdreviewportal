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
  INNOVATION_DISCIPLINE,
} from '../lib/generate'
import { exportProjectToExcel } from '../lib/export'
import { generateSMPReport } from '../lib/report'
// eslint-disable-next-line @typescript-eslint/no-require-imports
const { ZipArchive } = require('archiver') as { ZipArchive: new (opts?: object) => import('archiver').Archiver }

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

      /* 2. Create a stub project immediately so the response goes back before
            the AI call (which can exceed Render's 30-second request timeout). */
      const stub = await prisma.project.create({
        data: { name: 'Processing…', generationStatus: 'running' },
      })
      res.status(201).json({ projectId: stub.id })

      /* 3. All heavy work runs in the background — triggerCommentGeneration will
            set generationStatus 'complete' or 'error' when finished. */
      ;(async () => {
        try {
          await processUploadedPdf(stub.id, trimmedText, file.originalname)
        } catch (bgErr: unknown) {
          const e = bgErr as Error
          console.error('[create-from-pdf] background processing failed:', e?.message ?? bgErr)
          await prisma.project.update({
            where: { id: stub.id },
            data: { generationStatus: 'error' },
          }).catch(console.error)
        }
      })()
    } catch (err: unknown) {
      const e = err as Error
      console.error('[create-from-pdf] FAILED:', e?.message ?? err)
      if (!res.headersSent) {
        res.status(500).json({ error: `Failed to process PDF: ${e?.message ?? 'unknown error'}` })
      }
    }
  },
)

/* ── Background PDF processing (runs after response is sent) ── */
async function processUploadedPdf(projectId: string, trimmedText: string, originalName: string): Promise<void> {
  const userPrompt = `Extract from this BESS assessment PDF and return a single JSON object:
{
  "project": {
    "name": "string",
    "address": "string or null",
    "projectId": "string or null",
    "bessScore": number or null,
    "date": "ISO date string or null",
    "typology": "Multi-Residential | Mixed-Use | Townhouse | Non-Residential | null",
    "client": "string or null",
    "architect": "string or null",
    "totalDwellings": number or null,
    "buildingLevels": number or null,
    "siteArea": number or null,
    "rainwaterTankSize": number or null
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
- typology: classify as exactly one of the following based on what inputs the BESS assessment contains:
  * "Mixed-Use" — assessment includes BOTH residential inputs (apartments or townhouses) AND non-residential inputs (retail, office, commercial, or other non-residential uses)
  * "Multi-Residential" — assessment includes ONLY residential apartment inputs (no townhouses, no non-residential areas)
  * "Townhouse" — assessment includes ONLY residential townhouse inputs (no apartments, no non-residential areas)
  * "Non-Residential" — assessment includes NO residential inputs; only non-residential uses (retail, office, commercial, aged care, education, etc.)
  * null if the building type cannot be determined from the PDF
- client: the project client/owner name as stated in the PDF (cover page or project info section). Null if not stated.
- architect: the architect firm or person's name as stated in the PDF. Null if not stated.
- totalDwellings: total number of residential dwellings/apartments/units from the "Dwellings & Non-Residential Spaces" table (page 3 of BESS). If there is a "Total" row in the table, use that value. Otherwise sum all residential dwelling type rows (Studio, 1 Bed, 2 Bed, 3 Bed, etc.) — do NOT include non-residential rows (Retail, Office, Commercial, etc.). Null if not a residential project or not stated.
- buildingLevels: number of building levels/storeys from the "Height" column in the Buildings table on page 2 of the BESS PDF (e.g. if the Height cell says "8 Levels" return 8, if it says "8" return 8). Return only the integer count of levels. Null if not found.
- siteArea: site area in m² from the BESS PDF. In BESS the site area appears in the Buildings table (a row or column labelled "Site Area" with a value in m²), sometimes also on the cover page or project inputs section. Return only the integer number of square metres with no units (e.g. if it says "3,356 m²" return 3356, if it says "1234" return 1234). Return null — never 0 — if the value cannot be clearly identified in the document.
- rainwaterTankSize: rainwater tank storage volume in litres from the Rainwater Tank Profile table (look for "Tank Storage Volume", "Tank Volume", or "Storage Volume" fields). If the value is given in kL, convert to litres (e.g. 2 kL → 2000). Return only the integer number of litres with no units. Null if not found.
- creditStatus: Achieved → "Y", Not Achieved or 0% → "N", Scoped Out / N/A / Disabled → "ScopedOut"
- EXCLUDE credits with status Disabled from the credits array entirely, EXCEPT OE 4.2 which must always be included — if Disabled, include it with creditStatus "N" and creditScore 0
- creditRequirement: the specific compliance requirement for this credit as stated in the BESS document (e.g. "Provide a minimum of 24 long-stay bicycle spaces"). 1–2 sentences. Null if not stated.
- rawDataPoints: key numbers and specs only for this credit (e.g. "Required 24 long-stay bicycle spaces + 4 short-stay. Provided 20 long-stay + 2 short-stay." or "WELS 6-star taps, 3-star showers, no rainwater tank. 28% potable water reduction achieved."). 1 concise sentence max. Null if no specific data in PDF for this credit. Exception: for IWM 1.1 (potable water / fixtures), list every fixture type and its WELS star rating on a separate line — do not truncate. Exception: for Innovation credits, list each claimed initiative name on its own line as "• [initiative name]" — names only, no descriptions. Exception: for OE 2.x credits (OE 2.1, OE 2.2, OE 2.6, OE 2.7), extract every row from the "Dwellings & Non-Residential Spaces" table — for each typology list the name, quantity (number of units/tenancies), and total floor area in m² — do not truncate. Format each row as "Typology: N units, Xm² total" (e.g. "1 Bedroom: 20 units, 1100m² total", "Retail: 2 tenancies, 350m² total", "Office: 1 tenancy, 200m² total"). Also add a line "Total residential dwellings: N" summing all residential types only. Exception: for OE 3.1 and OE 3.2 (hot water) credits, always begin rawDataPoints with the exact text of the "Type of Hot Water System" field from the BESS assessment formatted as "Type of Hot Water System: [exact value as shown in BESS]", then any other key specs.
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
  if (rawContent.type !== 'text') throw new Error('Unexpected response type from AI parser')

  const jsonText = rawContent.text
  console.log(`[create-from-pdf] stop_reason: ${message.stop_reason} — input: ${message.usage.input_tokens}, output: ${message.usage.output_tokens}, response length: ${jsonText.length}`)

  let parsed: {
    project: {
      name: string
      address?: string
      projectId?: string
      bessScore?: number
      date?: string
      typology?: string
      client?: string
      architect?: string
      totalDwellings?: number
      buildingLevels?: number
      siteArea?: number
      rainwaterTankSize?: number
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
  } catch {
    console.error('[create-from-pdf] JSON parse failed. stop_reason:', message.stop_reason)
    console.error('[create-from-pdf] Raw response:\n', jsonText.slice(0, 1200))
    throw new Error('AI parser returned invalid JSON')
  }

  /* Detect existing project by address → auto-assign revision */
  const extractedAddress = parsed.project.address?.trim() || null
  const normalizeAddr = (a: string) => a.toLowerCase().replace(/\s+/g, ' ').trim()

  let parentProjectId: string | null = null
  let revisionLabel = 'A'

  if (extractedAddress) {
    const rootProjects = await prisma.project.findMany({
      where: { parentProjectId: null, address: { not: null }, id: { not: projectId } },
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

  /* Update stub project with real data */
  await prisma.project.update({
    where: { id: projectId },
    data: {
      name: parsed.project.name?.trim() || originalName,
      address: extractedAddress,
      bessScore: parsed.project.bessScore != null ? parseFloat(String(parsed.project.bessScore)) : null,
      date: parsed.project.date ? new Date(parsed.project.date) : null,
      projectId: parsed.project.projectId?.trim() || null,
      revision: revisionLabel,
      parentProjectId,
      typology: parsed.project.typology?.trim() || null,
      client: parsed.project.client?.trim() || null,
      architect: parsed.project.architect?.trim() || null,
      totalDwellings: parsed.project.totalDwellings != null ? parseInt(String(parsed.project.totalDwellings)) : null,
      buildingLevels: parsed.project.buildingLevels != null ? parseInt(String(parsed.project.buildingLevels)) : null,
      siteArea: parsed.project.siteArea != null ? parseInt(String(parsed.project.siteArea)) : null,
      rainwaterTankSize: parsed.project.rainwaterTankSize != null ? parseInt(String(parsed.project.rainwaterTankSize)) : null,
    },
  })

  /* Create credits */
  const creditData = (parsed.credits ?? []).map((c) => {
    const meta = CREDIT_META[normaliseCreditId(c.creditId)] ?? CREDIT_META[c.creditId] ?? {
      mandatory: false,
      responsibleParty: null,
    }
    return {
      projectId,
      creditId: normaliseCreditId(c.creditId),
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

  /* Ensure OE 4.2 always exists — may be Disabled in BESS (excluded from credits array by AI) */
  const hasOE42 = creditData.some(c => c.creditId.toLowerCase() === 'oe 4.2')
  if (!hasOE42) {
    const oeCredit = creditData.find(c => c.creditId.toLowerCase().startsWith('oe'))
    await prisma.credit.create({
      data: {
        projectId,
        creditId: 'OE 4.2',
        creditName: 'On-site Renewable Energy',
        category: oeCredit?.category ?? 'Operational Energy',
        categoryOrder: 3,
        creditStatus: 'N',
        creditScore: 0,
        mandatory: false,
        responsibleParty: 'Developer / Architect / Services',
        hiddenFromPortal: true,
      },
    })
  }

  /* Create drawing requirements */
  const DRAWING_EXCLUDED_CREDITS = new Set(['ieq 1.1', 'ieq 1.2', 'ieq 2.1', 'ieq 3.1', 'management 2.3', 'iwm 2.1'])
  const achievedCreditIds = new Set(
    (parsed.credits ?? []).filter((c) => c.creditStatus === 'Y').map((c) => c.creditId.trim().toLowerCase()),
  )
  const drawingData = (parsed.supportingEvidence ?? [])
    .filter((e) => {
      const ref = e.creditReference.trim().toLowerCase()
      return e.type !== 'Supporting Document' && achievedCreditIds.has(ref) && !DRAWING_EXCLUDED_CREDITS.has(ref)
    })
    .map((e) => ({ projectId, creditReference: e.creditReference, drawingType: e.type, requirement: e.requirement }))
  if (drawingData.length > 0) {
    await prisma.drawingRequirement.createMany({ data: drawingData })
  }

  /* Carry over manually-edited GIW comments from the most recent prior revision */
  if (parentProjectId) {
    const prevRevisions = await prisma.project.findMany({
      where: { OR: [{ id: parentProjectId }, { parentProjectId, id: { not: projectId } }] },
      select: { id: true, createdAt: true },
      orderBy: { createdAt: 'desc' },
    })
    const prevProjectId = prevRevisions[0]?.id
    if (prevProjectId) {
      const editedCredits = await prisma.credit.findMany({
        where: { projectId: prevProjectId, lastEditedBy: { not: null } },
        select: { creditId: true, commentsGIW: true, lastEditedBy: true, lastEditedAt: true },
      })
      if (editedCredits.length > 0) {
        const newCredits = await prisma.credit.findMany({
          where: { projectId },
          select: { id: true, creditId: true },
        })
        const editedMap = new Map(editedCredits.map(c => [c.creditId.toLowerCase().trim(), c]))
        await Promise.all(
          newCredits
            .filter(c => editedMap.has(c.creditId.toLowerCase().trim()))
            .map(c => {
              const prev = editedMap.get(c.creditId.toLowerCase().trim())!
              return prisma.credit.update({
                where: { id: c.id },
                data: { commentsGIW: prev.commentsGIW, lastEditedBy: prev.lastEditedBy, lastEditedAt: prev.lastEditedAt },
              })
            }),
        )
        console.log(`[create-from-pdf] Carried over ${editedCredits.length} manually-edited GIW comment(s) from prior revision`)
      }
    }
  }

  /* Trigger comment + drawing generation (these set generationStatus to complete/error) */
  triggerCommentGeneration(projectId).catch(console.error)
  triggerDrawingGeneration(projectId).catch(console.error)
}

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
    const { name, address, bessScore, date, revision, projectId, generationStatus, notifyEmail, gdft, typology, client, architect, siteArea, totalDwellings, buildingLevels, rainwaterTankSize } =
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
        typology: string | null
        client: string | null
        architect: string | null
        siteArea: number | null
        totalDwellings: number | null
        buildingLevels: number | null
        rainwaterTankSize: number | null
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
        ...(typology !== undefined && { typology: typology || null }),
        ...(client !== undefined && { client: client || null }),
        ...(architect !== undefined && { architect: architect || null }),
        ...(siteArea !== undefined && { siteArea: siteArea != null ? parseInt(String(siteArea)) : null }),
        ...(totalDwellings !== undefined && { totalDwellings: totalDwellings != null ? parseInt(String(totalDwellings)) : null }),
        ...(buildingLevels !== undefined && { buildingLevels: buildingLevels != null ? parseInt(String(buildingLevels)) : null }),
        ...(rainwaterTankSize !== undefined && { rainwaterTankSize: rainwaterTankSize != null ? parseInt(String(rainwaterTankSize)) : null }),
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
    const trimmedName = name?.trim() || null
    let reviewer
    if (existing) {
      reviewer = await prisma.reviewer.update({
        where: { id: existing.id },
        data: {
          discipline: discipline.trim(),
          inviteToken,
          // Only overwrite stored name if a new one was provided
          ...(trimmedName !== null && { name: trimmedName }),
        },
      })
    } else {
      reviewer = await prisma.reviewer.create({
        data: {
          projectId: project.id,
          email: normalEmail,
          name: trimmedName,
          discipline: discipline.trim(),
          inviteToken,
        },
      })
    }
    const base = process.env.BASE_URL || 'http://localhost:5173'
    const inviteLink = `${base}/review/invite/${inviteToken}`

    // Use stored name as fallback so resend also gets the name
    const resolvedName = trimmedName ?? reviewer.name ?? null

    // Race email send against a 35s timeout so the UI gets real feedback without hanging forever
    const timeout = new Promise<'timeout'>((resolve) => setTimeout(() => resolve('timeout'), 35000))
    const result = await Promise.race([
      sendReviewInviteByEmail(normalEmail, inviteLink, project.name, discipline.trim(), resolvedName, project.address || null)
        .then(() => 'sent' as const)
        .catch((err: unknown) => {
          console.error('[invite] email send failed:', err)
          return err instanceof Error ? err.message : String(err)
        }),
      timeout,
    ])

    if (result === 'sent') {
      res.json({ ok: true, reviewerId: reviewer.id })
    } else if (result === 'timeout') {
      res.json({ ok: true, reviewerId: reviewer.id, emailWarning: 'Email is taking longer than expected — it may still arrive, or check your SMTP settings.' })
    } else {
      res.json({ ok: true, reviewerId: reviewer.id, emailWarning: `Email failed to send: ${result}` })
    }
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

/* POST /api/projects/:id/report — generate and download SMP Word + Excel ZIP */
router.post('/:id/report', requireGIW, async (req: Request, res: Response): Promise<void> => {
  try {
    const project = await prisma.project.findUnique({
      where: { id: req.params.id },
      select: {
        name: true,
        address: true,
        projectId: true,
        date: true,
        revision: true,
        typology: true,
        client: true,
        architect: true,
        bessScore: true,
        totalDwellings: true,
        buildingLevels: true,
        siteArea: true,
        rainwaterTankSize: true,
        gdft: true,
      },
    })
    if (!project) {
      res.status(404).json({ error: 'Project not found' })
      return
    }

    const rawCredits = await prisma.credit.findMany({
      where: { projectId: req.params.id, deletedByGIW: false },
      select: {
        creditId: true, creditName: true, creditStatus: true,
        rawDataPoints: true, parsedValues: true, category: true, commentsGIW: true,
        comments: { select: { commentText: true } },
      },
    })
    const credits = rawCredits.map(c => ({
      creditId: c.creditId,
      creditName: c.creditName,
      creditStatus: c.creditStatus,
      rawDataPoints: c.rawDataPoints,
      parsedValues: c.parsedValues as Record<string, string> | null,
      category: c.category,
      commentsGIW: c.commentsGIW,
      reviewerComments: c.comments.map(cm => cm.commentText),
    }))

    const { client, architect, giwref } = req.body as { client?: string; architect?: string; giwref?: string }
    const { wordBuffer, excelBuffer, wordFilename, excelFilename } = await generateSMPReport(
      project,
      credits,
      { client, architect, giwref },
    )

    const safeName = project.name.replace(/[^a-zA-Z0-9\s-]/g, '').trim().replace(/\s+/g, '-')
    const rev = project.revision ?? 'A'
    const zipFilename = `SMP-Report-${safeName}-Rev${rev}.zip`

    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': `attachment; filename="${zipFilename}"`,
    })

    const archive = new ZipArchive({ zlib: { level: 6 } })
    archive.on('error', (err) => {
      console.error('[report] archiver error:', err)
      if (!res.headersSent) res.status(500).end()
    })
    archive.pipe(res)
    archive.append(wordBuffer, { name: wordFilename })
    archive.append(excelBuffer, { name: excelFilename })
    await archive.finalize()
  } catch (err: unknown) {
    const e = err as Error
    console.error('[report]', e)
    if (!res.headersSent) {
      const status = e.message?.includes('not yet available') || e.message?.includes('not set') || e.message?.includes('Unknown typology') ? 422 : 500
      res.status(status).json({ error: e.message ?? 'Failed to generate report' })
    }
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

    const reviewers = await prisma.reviewer.findMany({
      where: { projectId: req.params.id },
      select: { email: true, name: true },
    })
    const nameMap = new Map(reviewers.map(r => [r.email, r.name]))

    const annotateComments = <T extends { reviewerEmail: string }>(comments: T[]) =>
      comments.map(c => ({ ...c, reviewerName: nameMap.get(c.reviewerEmail) ?? null }))

    if (giw) {
      const credits = await prisma.credit.findMany({
        where: { projectId: req.params.id, deletedByGIW: false, hiddenFromPortal: false },
        orderBy: [{ categoryOrder: 'asc' }, { creditId: 'asc' }],
        include: {
          comments: { orderBy: { submittedAt: 'asc' } },
        },
      })
      res.json(credits.map(c => ({ ...c, comments: annotateComments(c.comments) })))
      return
    }

    // Reviewer path
    const searchTerm = disciplineSearchTerm(reviewerDiscipline!)
    const credits = await prisma.credit.findMany({
      where: {
        projectId: req.params.id,
        deletedByGIW: false,
        hiddenFromPortal: false,
        // Not-achieved credits are hidden; scoped-out visibility is managed via hiddenFromPortal
        creditStatus: { not: 'N' },
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
    res.json(credits.map(c => ({ ...c, comments: annotateComments(c.comments) })))
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
    const [linkedItems, innovationItems, weightCredits, allCredits, allItems, projectRow, reviewerNoteRows] = await Promise.all([
      // Excellence items tied to discipline-matching credits (non-innovation)
      prisma.eSDExcellenceOpportunity.findMany({
        where: { projectId: req.params.id, deletedByGIW: false, creditId: { in: creditIds } },
        orderBy: { creditReference: 'asc' },
      }),
      // Innovation items — filtered in memory by discipline
      prisma.eSDExcellenceOpportunity.findMany({
        where: { projectId: req.params.id, deletedByGIW: false, creditReference: 'Innovation' },
        orderBy: { creditName: 'asc' },
      }),
      prisma.credit.findMany({ where: { projectId: req.params.id, deletedByGIW: false, creditId: { not: 'Innovation' }, OR: [{ hiddenFromPortal: false }, { category: { contains: 'innovation', mode: 'insensitive' } }] }, select: creditSelect }),
      prisma.credit.findMany({ where: { projectId: req.params.id, deletedByGIW: false }, select: creditSelect }),
      prisma.eSDExcellenceOpportunity.findMany({ where: { projectId: req.params.id, deletedByGIW: false }, select: allItemsSelect }),
      prisma.project.findUnique({ where: { id: req.params.id }, select: { bessScore: true } }),
      prisma.eSDExcellenceNote.findMany({
        where: { reviewerEmail: reviewerEmail!, excellence: { projectId: req.params.id } },
        select: { excellenceId: true, notes: true },
      }),
    ])

    // Filter innovation items: show only those whose discipline matches the reviewer's search term.
    // Falls back to 'Architect / Developer' for any item not in the map.
    const filteredInnovation = innovationItems.filter(item => {
      if (!searchTerm) return true
      const disc = INNOVATION_DISCIPLINE[item.creditName.toLowerCase()] ?? 'Architect / Developer'
      return disc.toLowerCase().includes(searchTerm.toLowerCase())
    })
    const items = [...linkedItems, ...filteredInnovation]

    const reviewerNotesMap = Object.fromEntries(reviewerNoteRows.map(r => [r.excellenceId, r.notes]))
    const filteredWithPoints = computeItemsBessPoints(items, weightCredits, allCredits)
      .map(item => ({ ...item, reviewerNotes: reviewerNotesMap[item.id] ?? null }))
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
