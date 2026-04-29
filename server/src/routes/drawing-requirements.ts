import { Router, Request, Response } from 'express'
import { prisma } from '../lib/prisma'

const router = Router()

/* PATCH /api/drawing-requirements/:id — reviewer or GIW */
router.patch('/:id', async (req: Request, res: Response): Promise<void> => {
  const { status, notes } = req.body as { status?: string; notes?: string }

  const data: Record<string, unknown> = {}
  if (status !== undefined) data.status = status
  if (notes !== undefined) data.notes = notes

  if (Object.keys(data).length === 0) {
    res.status(400).json({ error: 'At least one of status or notes is required' })
    return
  }

  const VALID_STATUSES = new Set(['NotStarted', 'InProgress', 'Complete'])
  if (status !== undefined && !VALID_STATUSES.has(status)) {
    res.status(400).json({ error: 'Invalid status value' })
    return
  }

  try {
    const item = await prisma.drawingRequirement.update({
      where: { id: req.params.id },
      data,
    })
    res.json(item)
  } catch (err) {
    console.error('[drawing-requirements] patch error:', err)
    res.status(500).json({ error: 'Failed to update drawing requirement' })
  }
})

export default router
