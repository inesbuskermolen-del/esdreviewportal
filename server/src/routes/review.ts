import { Router, Request, Response } from 'express'
import { prisma } from '../lib/prisma'
import { sendSubmissionAlert, sendCompletionEmail } from '../lib/email'

const router = Router()

/* GET /api/review/invite/:inviteToken — public */
router.get('/invite/:inviteToken', async (req: Request, res: Response): Promise<void> => {
  try {
    const reviewer = await prisma.reviewer.findUnique({
      where: { inviteToken: req.params.inviteToken },
      include: { project: { select: { id: true, name: true, reviewLinkToken: true } } },
    })
    if (!reviewer) {
      res.status(404).json({ error: 'Invalid or expired invite link' })
      return
    }
    res.json({
      reviewerId: reviewer.id,
      projectId: reviewer.projectId,
      reviewLinkToken: reviewer.project.reviewLinkToken,
      projectName: reviewer.project.name,
      reviewerEmail: reviewer.email,
      reviewerDiscipline: reviewer.discipline,
    })
  } catch (err) {
    console.error('[review] invite lookup error:', err)
    res.status(500).json({ error: 'Failed to look up invite' })
  }
})

/* POST /api/review/identify */
router.post('/identify', async (req: Request, res: Response): Promise<void> => {
  const { reviewLinkToken, email, discipline } = req.body as {
    reviewLinkToken?: string
    email?: string
    discipline?: string
  }

  if (!reviewLinkToken || !email?.trim() || !discipline?.trim()) {
    res.status(400).json({ error: 'reviewLinkToken, email and discipline are required' })
    return
  }

  const normalEmail = email.trim().toLowerCase()

  try {
    const project = await prisma.project.findUnique({
      where: { reviewLinkToken },
      select: { id: true, name: true, address: true },
    })

    if (!project) {
      res.status(404).json({ error: 'Review not found' })
      return
    }

    const existing = await prisma.reviewer.findFirst({
      where: { projectId: project.id, email: normalEmail },
    })

    let reviewer
    if (existing) {
      reviewer = await prisma.reviewer.update({
        where: { id: existing.id },
        data: { discipline: discipline.trim() },
      })
    } else {
      reviewer = await prisma.reviewer.create({
        data: {
          projectId: project.id,
          email: normalEmail,
          discipline: discipline.trim(),
        },
      })
    }

    res.json({
      reviewerId: reviewer.id,
      projectId: project.id,
      projectName: project.name,
      projectAddress: project.address,
    })
  } catch (err) {
    console.error('[review] identify error:', err)
    res.status(500).json({ error: 'Failed to identify reviewer' })
  }
})

/* GET /api/review/:reviewLinkToken/project — public */
router.get('/:reviewLinkToken/project', async (req: Request, res: Response): Promise<void> => {
  try {
    const project = await prisma.project.findUnique({
      where: { reviewLinkToken: req.params.reviewLinkToken },
      select: { id: true, name: true, address: true, bessScore: true, gdft: true },
    })

    if (!project) {
      res.status(404).json({ error: 'Review not found' })
      return
    }

    res.json(project)
  } catch (err) {
    console.error('[review] project lookup error:', err)
    res.status(500).json({ error: 'Failed to fetch project info' })
  }
})

/* GET /api/review/:reviewLinkToken/drawings — public */
router.get('/:reviewLinkToken/drawings', async (req: Request, res: Response): Promise<void> => {
  try {
    const project = await prisma.project.findUnique({
      where: { reviewLinkToken: req.params.reviewLinkToken },
      select: { id: true },
    })

    if (!project) {
      res.status(404).json({ error: 'Review not found' })
      return
    }

    const drawings = await prisma.drawingRequirement.findMany({
      where: { projectId: project.id },
      orderBy: { creditReference: 'asc' },
    })

    res.json(drawings)
  } catch (err) {
    console.error('[review] drawings error:', err)
    res.status(500).json({ error: 'Failed to fetch drawing requirements' })
  }
})

/* POST /api/review/:projectId/submit */
router.post('/:projectId/submit', async (req: Request, res: Response): Promise<void> => {
  const { reviewerEmail, reviewerDiscipline } = req.body as {
    reviewerEmail?: string
    reviewerDiscipline?: string
  }

  if (!reviewerEmail?.trim() || !reviewerDiscipline?.trim()) {
    res.status(400).json({ error: 'reviewerEmail and reviewerDiscipline are required' })
    return
  }

  const normalEmail = reviewerEmail.trim().toLowerCase()

  try {
    const project = await prisma.project.findUnique({
      where: { id: req.params.projectId },
      select: { id: true, name: true, reviewLinkToken: true },
    })

    if (!project) {
      res.status(404).json({ error: 'Project not found' })
      return
    }

    const reviewer = await prisma.reviewer.findFirst({
      where: { projectId: project.id, email: normalEmail },
    })

    if (!reviewer) {
      res.status(404).json({ error: 'Reviewer not found' })
      return
    }

    const submittedAt = new Date()
    await prisma.reviewer.update({
      where: { id: reviewer.id },
      data: { hasSubmitted: true, submittedAt },
    })

    const fullProject = await prisma.project.findUnique({
      where: { id: project.id },
      select: {
        notifyEmail: true,
        reviewers: { select: { email: true, hasSubmitted: true } },
      },
    })

    sendSubmissionAlert({
      submitterEmail: normalEmail,
      submitterDiscipline: reviewerDiscipline.trim(),
      projectName: project.name,
      projectId: project.id,
      reviewLinkToken: project.reviewLinkToken,
      submittedAt,
      reviewerEmails: fullProject?.reviewers.map(r => r.email) ?? [],
      notifyEmail: fullProject?.notifyEmail ?? null,
    }).catch(err => console.error('[review] submission alert error:', err))

    const allSubmitted = fullProject?.reviewers.every(r => r.hasSubmitted) ?? false
    if (allSubmitted && (fullProject?.reviewers.length ?? 0) > 0) {
      // fire-and-forget
      ;(async () => {
        try {
          const { exportProjectToExcel } = await import('../lib/export')
          const excelBuffer = await exportProjectToExcel(project.id)
          const fileName = `${project.name.replace(/[^a-z0-9]/gi, '-')}-review-matrix.xlsx`
          await sendCompletionEmail({
            projectName: project.name,
            reviewerEmails: fullProject!.reviewers.map(r => r.email),
            excelBuffer,
            fileName,
          })
        } catch (err) {
          console.error('[review] completion email error:', err)
        }
      })()
    }

    res.json({ ok: true })
  } catch (err) {
    console.error('[review] submit error:', err)
    res.status(500).json({ error: 'Failed to submit review' })
  }
})

export default router
