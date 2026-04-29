import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  TextRun,
  HeadingLevel,
  AlignmentType,
  WidthType,
  BorderStyle,
  ShadingType,
  VerticalAlign,
} from 'docx'

interface CreditSnapshot {
  creditId: string
  creditName: string
  category: string
  categoryOrder: number
  creditStatus: string
  creditScore: number | null
}

interface ChangelogInput {
  projectName: string
  address: string | null
  prevRevision: string
  newRevision: string
  prevBessScore: number | null
  newBessScore: number | null
  prevCredits: CreditSnapshot[]
  newCredits: CreditSnapshot[]
}

/* ── Colour palette (hex without #) ── */
const OLIVE = '4E5A2A'
const OLIVE_LIGHT = 'E8EDD8'
const CHARCOAL = '2C2C2C'
const WHITE = 'FFFFFF'
const RED_BG = 'FCE4D6'
const GREEN_BG = 'C6EFCE'
const GREY_BG = 'D9D9D9'
const AMBER_BG = 'FFF2CC'
const HEADER_BG = '3D4F1F'

const noBorder = {
  top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
}

const thinBorder = {
  top: { style: BorderStyle.SINGLE, size: 4, color: 'D8D5CE' },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: 'D8D5CE' },
  left: { style: BorderStyle.SINGLE, size: 4, color: 'D8D5CE' },
  right: { style: BorderStyle.SINGLE, size: 4, color: 'D8D5CE' },
}

function cell(
  text: string,
  opts: {
    bold?: boolean
    color?: string
    bg?: string
    width?: number
    align?: (typeof AlignmentType)[keyof typeof AlignmentType]
    size?: number
  } = {},
): TableCell {
  return new TableCell({
    borders: thinBorder,
    shading: opts.bg ? { type: ShadingType.CLEAR, fill: opts.bg, color: opts.bg } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    children: [
      new Paragraph({
        alignment: opts.align ?? AlignmentType.LEFT,
        children: [
          new TextRun({
            text,
            bold: opts.bold ?? false,
            color: opts.color ?? CHARCOAL,
            font: 'Open Sans',
            size: (opts.size ?? 9) * 2,
          }),
        ],
      }),
    ],
  })
}

function headerCell(text: string, width?: number): TableCell {
  return new TableCell({
    borders: thinBorder,
    shading: { type: ShadingType.CLEAR, fill: HEADER_BG, color: HEADER_BG },
    verticalAlign: VerticalAlign.CENTER,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
          new TextRun({
            text,
            bold: true,
            color: WHITE,
            font: 'Montserrat',
            size: 18,
          }),
        ],
      }),
    ],
  })
}

function statusBg(status: string): string {
  if (status === 'Y') return GREEN_BG
  if (status === 'ScopedOut') return GREY_BG
  return RED_BG
}

function statusLabel(status: string): string {
  if (status === 'Y') return 'Achieved'
  if (status === 'N') return 'Not Achieved'
  if (status === 'ScopedOut') return 'Scoped Out'
  return status
}

function scoreLabel(score: number | null): string {
  return score != null ? `${score}%` : '—'
}

function heading(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 80 },
    children: [
      new TextRun({
        text,
        bold: true,
        color: OLIVE,
        font: 'Montserrat',
        size: 24,
      }),
    ],
  })
}

function sectionRule(): Paragraph {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: OLIVE_LIGHT } },
    spacing: { after: 160 },
    children: [],
  })
}

function emptyRow(): Paragraph {
  return new Paragraph({ spacing: { after: 80 }, children: [] })
}

/* ── Build the Word document buffer ── */
export async function buildChangelogDocx(input: ChangelogInput): Promise<Buffer> {
  const {
    projectName, address, prevRevision, newRevision,
    prevBessScore, newBessScore, prevCredits, newCredits,
  } = input

  /* Index credits by creditId for fast lookup */
  const prevMap = new Map(prevCredits.map(c => [c.creditId, c]))
  const newMap = new Map(newCredits.map(c => [c.creditId, c]))

  /* Classify changes */
  type ChangeRow = {
    creditId: string; creditName: string; category: string; categoryOrder: number
    prevStatus: string; newStatus: string
    prevScore: number | null; newScore: number | null
    changeType: 'status' | 'score' | 'added' | 'removed'
  }

  const changes: ChangeRow[] = []

  for (const [id, nc] of newMap) {
    const pc = prevMap.get(id)
    if (!pc) {
      changes.push({ creditId: id, creditName: nc.creditName, category: nc.category, categoryOrder: nc.categoryOrder, prevStatus: '—', newStatus: nc.creditStatus, prevScore: null, newScore: nc.creditScore, changeType: 'added' })
    } else if (pc.creditStatus !== nc.creditStatus) {
      changes.push({ creditId: id, creditName: nc.creditName, category: nc.category, categoryOrder: nc.categoryOrder, prevStatus: pc.creditStatus, newStatus: nc.creditStatus, prevScore: pc.creditScore, newScore: nc.creditScore, changeType: 'status' })
    } else if (pc.creditScore !== nc.creditScore) {
      changes.push({ creditId: id, creditName: nc.creditName, category: nc.category, categoryOrder: nc.categoryOrder, prevStatus: pc.creditStatus, newStatus: nc.creditStatus, prevScore: pc.creditScore, newScore: nc.creditScore, changeType: 'score' })
    }
  }
  for (const [id, pc] of prevMap) {
    if (!newMap.has(id)) {
      changes.push({ creditId: id, creditName: pc.creditName, category: pc.category, categoryOrder: pc.categoryOrder, prevStatus: pc.creditStatus, newStatus: '—', prevScore: pc.creditScore, newScore: null, changeType: 'removed' })
    }
  }

  changes.sort((a, b) => a.categoryOrder - b.categoryOrder || a.creditId.localeCompare(b.creditId))

  /* BESS score delta */
  const scoreDelta =
    prevBessScore != null && newBessScore != null
      ? newBessScore - prevBessScore
      : null

  const deltaText =
    scoreDelta != null
      ? scoreDelta > 0 ? `+${scoreDelta.toFixed(1)}%` : `${scoreDelta.toFixed(1)}%`
      : '—'

  const deltaBg = scoreDelta == null ? WHITE : scoreDelta > 0 ? GREEN_BG : scoreDelta < 0 ? RED_BG : WHITE

  /* ── Document body ── */
  const children: (Paragraph | Table)[] = []

  /* Title */
  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: projectName, bold: true, color: OLIVE, font: 'Montserrat', size: 36 }),
      ],
    }),
    new Paragraph({
      spacing: { after: 20 },
      children: [
        new TextRun({ text: address ?? '', color: '666666', font: 'Open Sans', size: 18 }),
      ],
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [
        new TextRun({ text: `Revision Changelog: Rev ${prevRevision} → Rev ${newRevision}`, bold: true, color: CHARCOAL, font: 'Montserrat', size: 22 }),
      ],
    }),
  )

  /* BESS Score summary table */
  children.push(heading('BESS Score'))
  children.push(sectionRule())
  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            headerCell(`Rev ${prevRevision} Score`, 2800),
            headerCell(`Rev ${newRevision} Score`, 2800),
            headerCell('Change', 2000),
          ],
        }),
        new TableRow({
          children: [
            cell(prevBessScore != null ? `${prevBessScore}%` : '—', { align: AlignmentType.CENTER, size: 11 }),
            cell(newBessScore != null ? `${newBessScore}%` : '—', { align: AlignmentType.CENTER, size: 11 }),
            cell(deltaText, { align: AlignmentType.CENTER, bg: deltaBg, size: 11, bold: true }),
          ],
        }),
      ],
    }),
    emptyRow(),
  )

  /* Credit changes table */
  children.push(heading('Credit Changes'))
  children.push(sectionRule())

  if (changes.length === 0) {
    children.push(
      new Paragraph({
        children: [new TextRun({ text: 'No credit changes detected between revisions.', font: 'Open Sans', size: 18, color: '666666' })],
      }),
    )
  } else {
    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              headerCell('Credit', 1800),
              headerCell('Credit Name', 3600),
              headerCell(`Rev ${prevRevision} Status`, 1600),
              headerCell(`Rev ${newRevision} Status`, 1600),
              headerCell(`Rev ${prevRevision} Score`, 1200),
              headerCell(`Rev ${newRevision} Score`, 1200),
            ],
          }),
          ...changes.map(ch => {
            const rowBg =
              ch.changeType === 'added' ? GREEN_BG :
              ch.changeType === 'removed' ? GREY_BG :
              ch.changeType === 'status' && ch.newStatus === 'Y' ? GREEN_BG :
              ch.changeType === 'status' && ch.newStatus === 'N' ? RED_BG :
              ch.changeType === 'score' ? AMBER_BG : WHITE

            const prevStatusLabel = ch.prevStatus === '—' ? '—' : statusLabel(ch.prevStatus)
            const newStatusLabel = ch.newStatus === '—' ? '—' : statusLabel(ch.newStatus)

            return new TableRow({
              children: [
                cell(ch.creditId, { bg: rowBg, size: 8 }),
                cell(ch.creditName, { bg: rowBg, size: 8 }),
                cell(prevStatusLabel, { bg: ch.prevStatus !== '—' ? statusBg(ch.prevStatus) : rowBg, size: 8 }),
                cell(newStatusLabel, { bg: ch.newStatus !== '—' ? statusBg(ch.newStatus) : rowBg, size: 8 }),
                cell(scoreLabel(ch.prevScore), { bg: rowBg, align: AlignmentType.CENTER, size: 8 }),
                cell(scoreLabel(ch.newScore), { bg: rowBg, align: AlignmentType.CENTER, size: 8 }),
              ],
            })
          }),
        ],
      }),
      emptyRow(),
    )
  }

  /* Legend */
  children.push(
    new Paragraph({
      spacing: { before: 160, after: 40 },
      children: [new TextRun({ text: 'Legend', bold: true, font: 'Montserrat', size: 18, color: CHARCOAL })],
    }),
    new Table({
      width: { size: 5000, type: WidthType.DXA },
      rows: [
        new TableRow({ children: [cell('', { bg: GREEN_BG, width: 300 }), cell('Status improved / Credit added', { size: 8 })] }),
        new TableRow({ children: [cell('', { bg: RED_BG, width: 300 }), cell('Status declined', { size: 8 })] }),
        new TableRow({ children: [cell('', { bg: AMBER_BG, width: 300 }), cell('Score changed (status unchanged)', { size: 8 })] }),
        new TableRow({ children: [cell('', { bg: GREY_BG, width: 300 }), cell('Credit removed / Scoped Out', { size: 8 })] }),
      ],
    }),
  )

  /* Footer */
  children.push(
    new Paragraph({
      spacing: { before: 400 },
      children: [
        new TextRun({
          text: `Generated ${new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' })} · GIW Environmental Solutions`,
          font: 'Open Sans', size: 16, color: '999999', italics: true,
        }),
      ],
    }),
  )

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Open Sans', size: 18, color: CHARCOAL },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 720, bottom: 720, left: 900, right: 900 },
          },
        },
        children,
      },
    ],
  })

  return Buffer.from(await Packer.toBuffer(doc))
}
