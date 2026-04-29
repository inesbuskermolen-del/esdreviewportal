import { Router, Request, Response } from 'express'
import { prisma } from '../lib/prisma'
import { requireGIW, GIWRequest } from '../middleware/auth'

const router = Router()

/* PATCH /api/credits/:id/giw-comment — GIW only */
router.patch(
  '/:id/giw-comment',
  requireGIW,
  async (req: GIWRequest, res: Response): Promise<void> => {
    const { commentText } = req.body as { commentText?: string }
    try {
      const credit = await prisma.credit.update({
        where: { id: req.params.id },
        data: {
          commentsGIW: commentText?.trim() || null,
          lastEditedBy: req.giw!.email,
          lastEditedAt: new Date(),
        },
        select: { id: true, commentsGIW: true },
      })
      res.json(credit)
    } catch (err) {
      console.error('[credits] giw-comment error:', err)
      res.status(500).json({ error: 'Failed to save comment' })
    }
  },
)

/* POST /api/credits/:id/comment — reviewer (no auth, pass email in body) */
router.post('/:id/comment', async (req: Request, res: Response): Promise<void> => {
  const { reviewerEmail, reviewerDiscipline, commentText } = req.body as {
    reviewerEmail?: string
    reviewerDiscipline?: string
    commentText?: string
  }

  if (!reviewerEmail?.trim() || !reviewerDiscipline?.trim()) {
    res.status(400).json({ error: 'reviewerEmail and reviewerDiscipline are required' })
    return
  }

  const normalEmail = reviewerEmail.trim().toLowerCase()

  try {
    const credit = await prisma.credit.findUnique({
      where: { id: req.params.id },
      select: { projectId: true },
    })
    if (!credit) {
      res.status(404).json({ error: 'Credit not found' })
      return
    }

    const existing = await prisma.creditComment.findFirst({
      where: { creditId: req.params.id, reviewerEmail: normalEmail },
    })

    if (existing) {
      await prisma.creditComment.update({
        where: { id: existing.id },
        data: {
          commentText: commentText?.trim() ?? '',
          reviewerDiscipline: reviewerDiscipline.trim(),
        },
      })
    } else {
      await prisma.creditComment.create({
        data: {
          creditId: req.params.id,
          projectId: credit.projectId,
          reviewerEmail: normalEmail,
          reviewerDiscipline: reviewerDiscipline.trim(),
          commentText: commentText?.trim() ?? '',
        },
      })
    }

    res.json({ ok: true })
  } catch (err) {
    console.error('[credits] comment error:', err)
    res.status(500).json({ error: 'Failed to save comment' })
  }
})

/* DELETE /api/credits/:id — GIW only (soft delete for scoped-out credits) */
router.delete(
  '/:id',
  requireGIW,
  async (req: Request, res: Response): Promise<void> => {
    try {
      await prisma.credit.update({
        where: { id: req.params.id },
        data: { deletedByGIW: true },
      })
      res.json({ ok: true })
    } catch (err) {
      console.error('[credits] delete error:', err)
      res.status(500).json({ error: 'Failed to delete credit' })
    }
  },
)

/* PATCH /api/credits/:id/visibility — GIW only (hide/show from reviewer portal) */
router.patch(
  '/:id/visibility',
  requireGIW,
  async (req: Request, res: Response): Promise<void> => {
    const { hiddenFromPortal } = req.body as { hiddenFromPortal?: boolean }
    if (typeof hiddenFromPortal !== 'boolean') {
      res.status(400).json({ error: 'hiddenFromPortal must be a boolean' })
      return
    }
    try {
      const credit = await prisma.credit.update({
        where: { id: req.params.id },
        data: { hiddenFromPortal },
      })
      res.json(credit)
    } catch (err) {
      console.error('[credits] visibility error:', err)
      res.status(500).json({ error: 'Failed to update visibility' })
    }
  },
)

export default router
