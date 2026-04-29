import { Router, Request, Response } from 'express'
import { prisma } from '../lib/prisma'
import { requireGIW } from '../middleware/auth'

const router = Router()

/* PATCH /api/excellence/:id/flag — GIW or reviewer */
router.patch('/:id/flag', async (req: Request, res: Response): Promise<void> => {
  const { flag, flaggedBy } = req.body as { flag?: string; flaggedBy?: string }

  if (!flag) {
    res.status(400).json({ error: 'flag is required' })
    return
  }

  try {
    const item = await prisma.eSDExcellenceOpportunity.update({
      where: { id: req.params.id },
      data: {
        flag,
        flaggedBy: flaggedBy?.trim() || null,
        flaggedAt: new Date(),
      },
    })
    res.json(item)
  } catch (err) {
    console.error('[excellence] flag error:', err)
    res.status(500).json({ error: 'Failed to update flag' })
  }
})

/* PATCH /api/excellence/:id/description — GIW only */
router.patch('/:id/description', requireGIW, async (req: Request, res: Response): Promise<void> => {
  const { improvementDescription } = req.body as { improvementDescription?: string }
  try {
    const item = await prisma.eSDExcellenceOpportunity.update({
      where: { id: req.params.id },
      data: { improvementDescription: improvementDescription?.trim() || null },
    })
    res.json(item)
  } catch (err) {
    console.error('[excellence] description error:', err)
    res.status(500).json({ error: 'Failed to save description' })
  }
})

/* PATCH /api/excellence/:id/notes — GIW or reviewer */
router.patch('/:id/notes', async (req: Request, res: Response): Promise<void> => {
  const { reviewerNotes } = req.body as { reviewerNotes?: string }
  try {
    const item = await prisma.eSDExcellenceOpportunity.update({
      where: { id: req.params.id },
      data: { reviewerNotes: reviewerNotes?.trim() || null },
    })
    res.json(item)
  } catch (err) {
    console.error('[excellence] notes error:', err)
    res.status(500).json({ error: 'Failed to save notes' })
  }
})

/* DELETE /api/excellence/:id — GIW only (soft delete) */
router.delete('/:id', requireGIW, async (_req: Request, res: Response): Promise<void> => {
  try {
    await prisma.eSDExcellenceOpportunity.update({
      where: { id: _req.params.id },
      data: { deletedByGIW: true },
    })
    res.json({ ok: true })
  } catch (err) {
    console.error('[excellence] delete error:', err)
    res.status(500).json({ error: 'Failed to remove item' })
  }
})

export default router
