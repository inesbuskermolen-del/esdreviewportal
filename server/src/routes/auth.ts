import { Router, Request, Response } from 'express'
import jwt from 'jsonwebtoken'
import { requireGIW, GIWRequest } from '../middleware/auth'
import { sendMagicLink } from '../lib/email'

const router = Router()

/* POST /api/auth/request-link
   Accepts any email. Only sends a link for @giw.com.au addresses.
   Always returns 200 to avoid email enumeration. */
router.post('/request-link', async (req: Request, res: Response): Promise<void> => {
  const { email } = req.body as { email?: string }

  res.status(200).json({ ok: true })

  if (!email?.toLowerCase().endsWith('@giw.com.au')) return

  const token = jwt.sign(
    { email: email.toLowerCase(), isGIW: true },
    process.env.JWT_SECRET!,
    { expiresIn: '15m' },
  )

  const base = process.env.BASE_URL || 'http://localhost:5173'
  const link = `${base}/auth/verify?token=${encodeURIComponent(token)}`

  sendMagicLink(email.toLowerCase(), link).catch((err) =>
    console.error('[auth] Failed to send magic link:', err),
  )
})

/* GET /api/auth/verify?token=
   Called by the frontend /auth/verify page via the Vite proxy.
   Sets a 7-day session cookie and returns { ok: true }. */
router.get('/verify', async (req: Request, res: Response): Promise<void> => {
  const { token } = req.query as { token?: string }

  if (!token) {
    res.status(400).json({ error: 'Missing token' })
    return
  }

  try {
    const payload = jwt.verify(token, process.env.JWT_SECRET!) as {
      email: string
      isGIW?: boolean
    }

    if (!payload.isGIW) {
      res.status(403).json({ error: 'Not a GIW token' })
      return
    }

    const sessionToken = jwt.sign(
      { email: payload.email, isGIW: true },
      process.env.JWT_SECRET!,
      { expiresIn: '7d' },
    )

    // Cross-origin (GitHub Pages → Render) requires SameSite=None; Secure.
    const isProd = process.env.NODE_ENV === 'production'
    res.cookie('giw_token', sessionToken, {
      httpOnly: true,
      secure: isProd,
      sameSite: isProd ? 'none' : 'lax',
      maxAge: 7 * 24 * 60 * 60 * 1000,
    })

    res.json({ ok: true })
  } catch {
    res.status(401).json({ error: 'Invalid or expired link' })
  }
})

/* POST /api/auth/logout */
router.post('/logout', (_req: Request, res: Response): void => {
  res.clearCookie('giw_token')
  res.json({ ok: true })
})

/* GET /api/auth/me */
router.get('/me', requireGIW, (req: GIWRequest, res: Response): void => {
  res.json({ email: req.giw!.email, isGIW: true })
})

export default router
