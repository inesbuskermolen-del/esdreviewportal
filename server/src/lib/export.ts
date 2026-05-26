import ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as path from 'path'
import { prisma } from './prisma'
import { computeItemsBessPoints, getCatWeight } from './bess'

/* ── Colour helpers ── */
const C = {
  green:      'FF00602B',  // GIW green — header rows, category bands
  greenLight: 'FFC8E6D4',  // light green tint — totals row, score summary accent
  white:      'FFFFFFFF',
  lightGrey:  'FFE1E1E1',  // page background
  midGrey:    'FFC0C0C0',  // borders + secondary labels
  achieved:   'FFC6EFCE',
  notAch:     'FFFCE4D6',
  scoped:     'FFC0C0C0',
}

function fill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } }
}

function font(opts: Partial<ExcelJS.Font>): Partial<ExcelJS.Font> {
  return { name: 'Calibri', size: 10, ...opts }
}

function border(): Partial<ExcelJS.Borders> {
  return {
    top:    { style: 'thin', color: { argb: C.midGrey } },
    left:   { style: 'thin', color: { argb: C.midGrey } },
    bottom: { style: 'thin', color: { argb: C.midGrey } },
    right:  { style: 'thin', color: { argb: C.midGrey } },
  }
}

function getCategoryOrder(creditId: string, category = ''): number {
  const id = creditId.toLowerCase()
  const cat = category.toLowerCase()
  if (id.startsWith('management')  || cat.startsWith('management'))         return 1
  if (id.startsWith('iwm')         || cat.includes('water'))                return 2
  if (id.startsWith('oe')          || cat.includes('operational energy'))   return 3
  if (id.startsWith('ieq')         || cat.includes('indoor environmental')) return 4
  if (id.startsWith('transport')   || cat.includes('transport'))            return 5
  if (id.startsWith('waste')       || cat.includes('waste'))                return 6
  if (id.startsWith('urban')       || cat.includes('urban'))                return 7
  if (id.startsWith('innovation')  || cat.includes('innovation'))           return 8
  return 99
}

function addLogoRow(
  wb: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  lastCol: string,
  logoBuffer: Buffer,
) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const imageId = wb.addImage({ buffer: logoBuffer as any, extension: 'png' })
  const logoRow = ws.addRow([])
  logoRow.height = 60
  ws.mergeCells(`A${logoRow.number}:${lastCol}${logoRow.number}`)
  logoRow.getCell(1).fill = fill(C.green)
  ws.addImage(imageId, `A${logoRow.number}:B${logoRow.number}`)
}

type CreditWithComments = Awaited<ReturnType<typeof fetchProject>>['credits'][number]
type ProjectData = Awaited<ReturnType<typeof fetchProject>>

async function fetchProject(projectId: string): Promise<{
  id: string
  name: string
  address: string | null
  date: Date | null
  bessScore: number | null
  revision: string | null
  credits: Array<{
    id: string
    creditId: string
    creditName: string
    category: string
    categoryOrder: number
    creditRequirement: string | null
    mandatory: boolean
    responsibleParty: string | null
    creditStatus: string
    creditScore: number | null
    creditWeight: number | null
    commentsGIW: string | null
    comments: Array<{
      reviewerEmail: string
      reviewerDiscipline: string
      commentText: string
    }>
  }>
  drawingItems: Array<{
    creditReference: string
    drawingType: string
    requirement: string
    discipline: string | null
    status: string
    notes: string | null
  }>
  excellenceItems: Array<{
    id: string
    creditReference: string
    creditName: string
    currentScore: number | null
    improvementDescription: string | null
    flag: string
    flaggedBy: string | null
    reviewerNotes: string | null
    bessPoints: string | null
    creditId: string | null
    additionalBessPoints?: number | null
    notesList: Array<{ reviewerEmail: string; notes: string }>
  }>
}> {
  return prisma.project.findUniqueOrThrow({
    where: { id: projectId },
    select: {
      id: true,
      name: true,
      address: true,
      date: true,
      bessScore: true,
      revision: true,
      credits: {
        where: { deletedByGIW: false, hiddenFromPortal: false },
        orderBy: [{ categoryOrder: 'asc' }, { creditId: 'asc' }],
        select: {
          id: true,
          creditId: true,
          creditName: true,
          category: true,
          categoryOrder: true,
          creditRequirement: true,
          mandatory: true,
          responsibleParty: true,
          creditStatus: true,
          creditScore: true,
          creditWeight: true,
          commentsGIW: true,
          comments: {
            orderBy: { submittedAt: 'asc' },
            select: { reviewerEmail: true, reviewerDiscipline: true, commentText: true },
          },
        },
      },
      drawingItems: {
        orderBy: { creditReference: 'asc' },
        select: {
          creditReference: true,
          drawingType: true,
          requirement: true,
          discipline: true,
          status: true,
          notes: true,
        },
      },
      excellenceItems: {
        where: { deletedByGIW: false },
        orderBy: { creditReference: 'asc' },
        select: {
          id: true,
          creditReference: true,
          creditName: true,
          reviewerNotes: true,
          currentScore: true,
          improvementDescription: true,
          flag: true,
          flaggedBy: true,
          bessPoints: true,
          creditId: true,
          notesList: { select: { reviewerEmail: true, notes: true } },
        },
      },
    },
  })
}

/* ── Reviewer comment colour palette (ARGB, no alpha = FF) ── */
const REVIEWER_COLORS = [
  'FF1F5C9E',  // deep blue
  'FF7B2D8B',  // purple
  'FFB35900',  // burnt orange
  'FF1A7A4A',  // forest green
  'FF8B1A1A',  // dark red
  'FF1A6B7A',  // teal
  'FF5C4A1F',  // brown
  'FF3D5C1F',  // olive green
]

/** Estimate the number of text lines a value will occupy given a column width (in Excel char units). */
function estimateLines(text: string, colWidth: number): number {
  if (!text) return 1
  return text.split('\n').reduce((total, line) => {
    return total + Math.max(1, Math.ceil(line.length / Math.max(1, colWidth)))
  }, 0)
}

/** Calculate row height so all text in the given cells is visible. */
function autoRowHeight(cells: Array<{ text: string; colWidth: number }>, lineHeightPt = 14, min = 16, max = 400): number {
  const maxLines = Math.max(1, ...cells.map(({ text, colWidth }) => estimateLines(text, colWidth)))
  return Math.min(max, Math.max(min, maxLines * lineHeightPt))
}

/**
 * Build ExcelJS RichText segments for reviewer comments.
 * Format per comment: `"email": "comment text"` in the reviewer's colour.
 * Comments separated by a blank line.
 */
function buildCommentRichText(
  comments: Array<{ reviewerEmail: string; reviewerDiscipline: string; commentText: string }>,
  emailColorMap: Map<string, string>,
): ExcelJS.CellRichTextValue {
  const richText: ExcelJS.RichText[] = []
  const filtered = comments.filter(c => c.commentText.trim())
  for (let i = 0; i < filtered.length; i++) {
    const c = filtered[i]
    const argb = emailColorMap.get(c.reviewerEmail) ?? 'FF000000'
    const label = `"${c.reviewerEmail}": `
    const body = `"${c.commentText.trim()}"`
    richText.push({
      font: { name: 'Calibri', size: 10, bold: true, color: { argb } },
      text: label,
    })
    richText.push({
      font: { name: 'Calibri', size: 10, color: { argb } },
      text: body,
    })
    if (i < filtered.length - 1) {
      richText.push({ font: { name: 'Calibri', size: 10 }, text: '\n\n' })
    }
  }
  return { richText }
}

/* ── Sheet 1: BESS Review Matrix ── */
function buildReviewMatrixSheet(wb: ExcelJS.Workbook, p: ProjectData, logoBuffer: Buffer | null, interactiveBESS?: number) {
  const ws = wb.addWorksheet('BESS Review Matrix')

  const LAST = 'F'

  ws.columns = [
    { key: 'name',    width: 36 },
    { key: 'req',     width: 46 },
    { key: 'mand',    width: 11 },
    { key: 'party',   width: 30 },
    { key: 'giwcmt',  width: 50 },
    { key: 'teamcmt', width: 60 },
  ]

  // ── Row 1: Logo banner ──
  const logoRow = ws.addRow([])
  logoRow.height = 60
  ws.mergeCells(`A${logoRow.number}:${LAST}${logoRow.number}`)
  logoRow.getCell(1).fill = fill(C.green)
  if (logoBuffer) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const imageId = wb.addImage({ buffer: logoBuffer as any, extension: 'png' })
    ws.addImage(imageId, {
      tl: { col: 0, row: logoRow.number - 1 },
      ext: { width: 160, height: 48 },
      editAs: 'oneCell',
    } as Parameters<typeof ws.addImage>[1])
  }

  // ── Rows 2–4: Project info (label col A | value cols B:C) + BESS badges (D:E and F) ──
  const infoRows: [string, string][] = [
    ['Project', p.name],
    ['Date', p.date ? p.date.toLocaleDateString('en-AU') : ''],
    ['Revision', p.revision?.trim() || 'A'],
  ]
  const firstInfoRowNum = ws.rowCount + 1
  for (const [label, value] of infoRows) {
    const row = ws.addRow([label])
    row.height = 18
    const labelCell = row.getCell(1)
    labelCell.value = label
    labelCell.fill = fill('FF004D22')
    labelCell.font = font({ bold: true, color: { argb: C.white }, size: 10 })
    labelCell.alignment = { vertical: 'middle', horizontal: 'right' }
    ws.mergeCells(`B${row.number}:C${row.number}`)
    const valueCell = row.getCell(2)
    valueCell.value = value
    valueCell.fill = fill(C.green)
    valueCell.font = font({ color: { argb: C.white }, size: 10 })
    valueCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  }
  const lastInfoRowNum = ws.rowCount

  // BESS badges: D:E = Baseline, F = Improved — both span the 3 info rows
  ws.mergeCells(`D${firstInfoRowNum}:E${lastInfoRowNum}`)
  const baselineCell = ws.getCell(`D${firstInfoRowNum}`)
  baselineCell.value = {
    richText: [
      { text: 'Baseline BESS Score\n', font: { name: 'Calibri', size: 8, color: { argb: 'BBFFFFFF' } } },
      { text: p.bessScore != null ? `${p.bessScore}%` : '—', font: { name: 'Calibri', size: 18, bold: true, color: { argb: C.white } } },
    ],
  }
  baselineCell.fill = fill(C.green)
  baselineCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }

  ws.mergeCells(`F${firstInfoRowNum}:F${lastInfoRowNum}`)
  const improvedCell = ws.getCell(`F${firstInfoRowNum}`)
  improvedCell.value = {
    richText: [
      { text: 'Improved BESS Score\n', font: { name: 'Calibri', size: 8, color: { argb: 'FF004D22' } } },
      { text: interactiveBESS != null ? `${interactiveBESS}%` : '—', font: { name: 'Calibri', size: 18, bold: true, color: { argb: 'FF004D22' } } },
    ],
  }
  improvedCell.fill = fill('FFC8E6D4')
  improvedCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }

  // Column header row
  const headerRow = ws.addRow([
    'Credit Name', 'Credit Requirement', 'Mandatory', 'Responsible Party',
    'Comments GIW', 'Comments Project Team',
  ])
  headerRow.height = 22
  headerRow.eachCell((cell) => {
    cell.fill = fill(C.green)
    cell.font = font({ bold: true, color: { argb: C.white }, size: 10 })
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    cell.border = border()
  })

  // Build per-reviewer colour map (consistent across all credits in this export)
  const allEmails = [...new Set(p.credits.flatMap(c => c.comments.map(cm => cm.reviewerEmail)))]
  const emailColorMap = new Map<string, string>(
    allEmails.map((email, i) => [email, REVIEWER_COLORS[i % REVIEWER_COLORS.length]]),
  )

  // Sort credits by BESS category order then creditId
  const sortedCredits = [...p.credits].sort((a, b) => {
    const oa = getCategoryOrder(a.creditId, a.category)
    const ob = getCategoryOrder(b.creditId, b.category)
    return oa !== ob ? oa - ob : a.creditId.localeCompare(b.creditId)
  })

  // Credit rows grouped by category
  let currentCategory = ''
  for (const credit of sortedCredits) {
    if (credit.category !== currentCategory) {
      currentCategory = credit.category
      const catRow = ws.addRow([currentCategory])
      ws.mergeCells(`A${catRow.number}:${LAST}${catRow.number}`)
      const cell = catRow.getCell(1)
      cell.fill = fill(C.green)
      cell.font = font({ bold: true, color: { argb: C.white }, size: 11 })
      cell.alignment = { vertical: 'middle', wrapText: false }
      catRow.height = 20
    }

    const filteredComments = credit.comments.filter(c => c.commentText.trim())
    const commentLineCount = filteredComments.reduce((n, c) => n + c.commentText.trim().split('\n').length + 1, 0)

    const dataRow = ws.addRow([
      `${credit.creditId} ${credit.creditName}`,
      credit.creditRequirement ?? '',
      credit.mandatory ? 'Y' : '',
      credit.responsibleParty ?? '',
      credit.commentsGIW ?? '',
      '',  // placeholder — rich text set below
    ])

    // Set rich-text reviewer comments on column F
    if (filteredComments.length > 0) {
      dataRow.getCell(6).value = buildCommentRichText(filteredComments, emailColorMap)
    }

    const commentText = filteredComments.map(c => c.commentText.trim()).join('\n\n')
    dataRow.height = autoRowHeight([
      { text: `${credit.creditId} ${credit.creditName}`, colWidth: 36 },
      { text: credit.creditRequirement ?? '', colWidth: 46 },
      { text: credit.commentsGIW ?? '', colWidth: 50 },
      { text: commentText, colWidth: 60 },
    ])

    dataRow.eachCell((cell, colNumber) => {
      cell.border = border()
      cell.alignment = { vertical: 'top', wrapText: true }
      // Don't overwrite the rich-text font on the team comments column (F = 6)
      if (colNumber !== 6) cell.font = font({ size: 10 })
    })

    // Name cell bold
    dataRow.getCell(1).font = font({ bold: true, size: 10 })
  }

  // Freeze panes: logo(1) + info rows(4) + badge row(1) + column header(1) = 7
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 5 }]

  // ESD Excellence Opportunities — Yes and Maybe flagged items only
  const flaggedItems = p.excellenceItems.filter(i => i.flag === 'Yes' || i.flag === 'Maybe')
  if (flaggedItems.length > 0) {
    // Blank separator
    ws.addRow([])

    // Section header
    const sectionRow = ws.addRow(['ESD Excellence Opportunities'])
    ws.mergeCells(`A${sectionRow.number}:${LAST}${sectionRow.number}`)
    const sectionCell = sectionRow.getCell(1)
    sectionCell.fill = fill(C.green)
    sectionCell.font = font({ bold: true, color: { argb: C.white }, size: 11 })
    sectionCell.alignment = { vertical: 'middle', wrapText: false }
    sectionRow.height = 20

    // Column headers
    const exHeader = ws.addRow([
      'Credit Reference', 'Credit Name', 'Current Score %',
      'Improvement Description', 'Additional BESS Points', 'Reviewer Notes',
    ])
    exHeader.height = 20
    exHeader.eachCell((cell) => {
      cell.fill = fill(C.green)
      cell.font = font({ bold: true, color: { argb: C.white } })
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
      cell.border = border()
    })

    // Sort: regular items by category order, then innovation alphabetically
    const regularFlagged = flaggedItems.filter(i => i.creditReference !== 'Innovation').sort((a, b) => {
      const oa = getCategoryOrder(a.creditReference)
      const ob = getCategoryOrder(b.creditReference)
      return oa !== ob ? oa - ob : (a.currentScore ?? 0) - (b.currentScore ?? 0)
    })
    const innovationFlagged = flaggedItems.filter(i => i.creditReference === 'Innovation').sort((a, b) =>
      a.creditName.localeCompare(b.creditName),
    )

    for (const item of [...regularFlagged, ...innovationFlagged]) {
      let addPts: number | string = ''
      if (item.creditReference === 'Innovation' && item.bessPoints) {
        const raw = Number(item.bessPoints)
        if (!isNaN(raw)) addPts = Math.round(raw * 0.9 * 10) / 10
      } else if (item.additionalBessPoints != null) {
        addPts = item.additionalBessPoints
      }

      // Collect reviewer notes: per-reviewer notesList first, fall back to legacy shared field
      const perReviewerNotes = item.notesList.filter(n => n.notes.trim())
      const notesText = perReviewerNotes.length > 0
        ? perReviewerNotes.map(n => n.notes.trim()).join('\n\n')
        : (item.reviewerNotes ?? '')

      const row = ws.addRow([
        item.creditReference,
        item.creditName,
        item.currentScore != null ? item.currentScore : '',
        item.improvementDescription ?? '',
        addPts,
        '',  // placeholder — rich text set below
      ])

      // Build rich-text reviewer notes (coloured by reviewer, matching BESS team comments style)
      if (perReviewerNotes.length > 0) {
        row.getCell(6).value = buildCommentRichText(
          perReviewerNotes.map(n => ({ reviewerEmail: n.reviewerEmail, reviewerDiscipline: '', commentText: n.notes.trim() })),
          emailColorMap,
        )
      } else if (item.reviewerNotes) {
        row.getCell(6).value = item.reviewerNotes
      }

      row.height = autoRowHeight([
        { text: item.creditName, colWidth: 36 },
        { text: item.improvementDescription ?? '', colWidth: 46 },
        { text: notesText, colWidth: 60 },
      ])
      row.eachCell((cell) => {
        cell.border = border()
        cell.alignment = { vertical: 'top', wrapText: true }
        cell.font = font({ size: 10 })
      })
      row.getCell(2).font = font({ bold: true, size: 10 })
    }
  }
}

/* ── Sheet 2: BESS Score Summary ── */
function buildScoreSummarySheet(wb: ExcelJS.Workbook, p: ProjectData, logoBuffer: Buffer | null, allCredits?: Array<{ creditId: string; category: string; creditStatus: string; creditWeight: number | null; creditScore: number | null }>) {
  const ws = wb.addWorksheet('BESS Score Summary')
  const LAST = 'D'

  ws.columns = [
    { key: 'category', width: 36 },
    { key: 'weight',   width: 14 },
    { key: 'score',    width: 14 },
    { key: 'weighted', width: 18 },
  ]

  if (logoBuffer) addLogoRow(wb, ws, LAST, logoBuffer)

  const headerRow = ws.addRow(['Category', 'Category Weight %', 'Category Score %', 'Weighted Score'])
  headerRow.height = 20
  headerRow.eachCell((cell) => {
    cell.fill = fill(C.green)
    cell.font = font({ bold: true, color: { argb: C.white } })
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    cell.border = border()
  })

  // Fixed BESS categories in order — ensures all 8 always appear even if no credits match
  const BESS_CATEGORIES = [
    { name: 'Management',                  order: 1 },
    { name: 'Integrated Water Management', order: 2 },
    { name: 'Operational Energy',          order: 3 },
    { name: 'Indoor Environmental Quality',order: 4 },
    { name: 'Transport',                   order: 5 },
    { name: 'Waste & Resource Recovery',   order: 6 },
    { name: 'Urban Ecology',               order: 7 },
    { name: 'Innovation',                  order: 8 },
  ]

  let totalBESS = 0

  for (const { name, order } of BESS_CATEGORIES) {
    const catWeight = getCatWeight(name)
    const sourceCredits = allCredits ?? p.credits
    const credits = sourceCredits.filter(c => getCategoryOrder(c.creditId, c.category) === order)

    const eligible = credits.filter(c => c.creditStatus !== 'ScopedOut' && c.creditWeight != null)
    const eligW = eligible.reduce((s, c) => s + (c.creditWeight ?? 0), 0)
    const catScore = eligW > 0
      ? eligible.reduce((s, c) => s + (c.creditScore ?? 0) * (c.creditWeight ?? 0), 0) / eligW
      : 0

    const weightedScore = catScore * catWeight / 100
    totalBESS += weightedScore

    const row = ws.addRow([
      name,
      catWeight > 0 ? catWeight : '',
      Math.round(catScore * 10) / 10,
      Math.round(weightedScore * 10) / 10,
    ])
    row.height = 16
    row.eachCell((cell) => { cell.border = border(); cell.alignment = { vertical: 'middle' } })
  }

  // Total BESS row
  const totalRow = ws.addRow(['Total BESS Score', '', '', Math.round(totalBESS * 10) / 10])
  totalRow.height = 18
  totalRow.eachCell((cell) => {
    cell.fill = fill(C.greenLight)
    cell.font = font({ bold: true })
    cell.border = border()
    cell.alignment = { vertical: 'middle' }
  })
}

/* ── Sheet 3: Drawing Requirements ── */
function buildDrawingsSheet(wb: ExcelJS.Workbook, p: ProjectData, logoBuffer: Buffer | null) {
  const ws = wb.addWorksheet('Drawing Requirements')
  const LAST = 'F'

  ws.columns = [
    { key: 'ref',    width: 16 },
    { key: 'type',   width: 22 },
    { key: 'req',    width: 55 },
    { key: 'disc',   width: 22 },
    { key: 'status', width: 16 },
    { key: 'notes',  width: 40 },
  ]

  if (logoBuffer) addLogoRow(wb, ws, LAST, logoBuffer)

  const headerRow = ws.addRow(['Credit Reference', 'Type', 'Requirement', 'Discipline', 'Status', 'Notes'])
  headerRow.height = 20
  headerRow.eachCell((cell) => {
    cell.fill = fill(C.green)
    cell.font = font({ bold: true, color: { argb: C.white } })
    cell.alignment = { vertical: 'middle', horizontal: 'center' }
    cell.border = border()
  })

  for (const d of p.drawingItems) {
    const row = ws.addRow([d.creditReference, d.drawingType, d.requirement, d.discipline ?? '', d.status, d.notes ?? ''])
    row.height = autoRowHeight([
      { text: d.requirement ?? '', colWidth: 55 },
      { text: d.notes ?? '', colWidth: 40 },
    ])
    row.eachCell((cell) => { cell.border = border(); cell.alignment = { vertical: 'top', wrapText: true } })
  }
}


/* ── Public export function ── */
export async function exportProjectToExcel(projectId: string): Promise<Buffer> {
  const creditSelect = { id: true, creditId: true, category: true, creditStatus: true, creditWeight: true, creditScore: true } as const
  const [project, allCredits] = await Promise.all([
    fetchProject(projectId),
    prisma.credit.findMany({ where: { projectId, deletedByGIW: false }, select: creditSelect }),
  ])

  const enrichedExcellenceItems = computeItemsBessPoints(
    project.excellenceItems,
    project.credits,
    allCredits,
  )
  const enrichedProject = { ...project, excellenceItems: enrichedExcellenceItems }

  let interactiveBESS = project.bessScore ?? 0
  for (const item of enrichedExcellenceItems) {
    if (item.flag !== 'Yes') continue
    if (item.creditReference === 'Innovation' && item.bessPoints) {
      const raw = Number(item.bessPoints)
      if (!isNaN(raw)) interactiveBESS += Math.round(raw * 0.9 * 10) / 10
    } else if (item.additionalBessPoints != null) {
      interactiveBESS += item.additionalBessPoints
    }
  }
  interactiveBESS = Math.round(interactiveBESS)

  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'GIW Environmental Solutions'
  workbook.created = new Date()

  const logoPath = path.resolve(__dirname, '../../../public/GIW logo.png')
  const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : null

  buildReviewMatrixSheet(workbook, enrichedProject, logoBuffer, interactiveBESS)

  const buffer = await workbook.xlsx.writeBuffer()
  return Buffer.from(buffer)
}
