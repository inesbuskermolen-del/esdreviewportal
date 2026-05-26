// Catch any uncaught synchronous throws (e.g. from module-level code in imports)
process.on('uncaughtException', (err) => {
  console.error('[startup] Uncaught exception:', err)
  process.exit(1)
})
process.on('unhandledRejection', (reason) => {
  console.error('[startup] Unhandled rejection:', reason)
  process.exit(1)
})

import express from 'express'
import cors from 'cors'
import cookieParser from 'cookie-parser'
import dotenv from 'dotenv'

import authRouter from './routes/auth'
import projectsRouter from './routes/projects'
import reviewRouter from './routes/review'
import creditsRouter from './routes/credits'
import excellenceRouter from './routes/excellence'
import drawingRequirementsRouter from './routes/drawing-requirements'

dotenv.config()

// Strip trailing slash from BASE_URL so URL construction never produces double slashes
if (process.env.BASE_URL) {
  process.env.BASE_URL = process.env.BASE_URL.replace(/\/+$/, '')
}

console.log('[startup] ESD Review Portal server initialising')
console.log('[startup] NODE_ENV    :', process.env.NODE_ENV ?? '(not set)')
console.log('[startup] PORT        :', process.env.PORT ?? '(not set — defaulting to 3001)')
console.log('[startup] DATABASE_URL:', process.env.DATABASE_URL ? '*** set ***' : '*** MISSING ***')
console.log('[startup] JWT_SECRET  :', process.env.JWT_SECRET  ? '*** set ***' : '*** MISSING ***')
console.log('[startup] BASE_URL    :', process.env.BASE_URL ?? '(not set — defaulting to http://localhost:5173)')

const REQUIRED_ENV = ['DATABASE_URL', 'JWT_SECRET']
const missing = REQUIRED_ENV.filter((k) => !process.env[k])
if (missing.length > 0) {
  console.error('[startup] FATAL: missing required environment variables:', missing.join(', '))
  console.error('[startup] Add these in the Render dashboard → Environment, then redeploy.')
  process.exit(1)
}

try {
  const app = express()
  const PORT = process.env.PORT || 3001

  // Always allow the production GitHub Pages origin and localhost dev server.
  // Also accept whatever BASE_URL is set to (extracts the origin, ignores any path suffix).
  const allowedOrigins = new Set(['https://inesbuskermolen-del.github.io', 'http://localhost:5173'])
  if (process.env.BASE_URL) {
    try { allowedOrigins.add(new URL(process.env.BASE_URL).origin) } catch { /* ignore bad URL */ }
  }
  console.log('[startup] CORS origins:', [...allowedOrigins].join(', '))
  app.use(
    cors({
      origin: (origin, cb) => cb(null, !origin || allowedOrigins.has(origin)),
      credentials: true,
    }),
  )
  app.use(express.json())
  app.use(cookieParser())

  app.use('/api/auth', authRouter)
  app.use('/api/projects', projectsRouter)
  app.use('/api/review', reviewRouter)
  app.use('/api/credits', creditsRouter)
  app.use('/api/excellence', excellenceRouter)
  app.use('/api/drawing-requirements', drawingRequirementsRouter)

  app.listen(PORT, () => {
    console.log(`[startup] Server listening on port ${PORT}`)
  })
} catch (err) {
  console.error('[startup] FATAL: failed to start server:', err)
  process.exit(1)
}
