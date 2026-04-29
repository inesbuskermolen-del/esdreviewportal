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

const app = express()
const PORT = process.env.PORT || 3001

app.use(
  cors({
    origin: process.env.BASE_URL || 'http://localhost:5173',
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
  console.log(`\nESD Review Portal server running on http://localhost:${PORT}\n`)
  console.log('Registered API routes:')
  console.log('  Auth:')
  console.log('    POST   /api/auth/request-link')
  console.log('    GET    /api/auth/verify')
  console.log('    POST   /api/auth/logout')
  console.log('    GET    /api/auth/me')
  console.log('  Projects:')
  console.log('    GET    /api/projects')
  console.log('    POST   /api/projects')
  console.log('    GET    /api/projects/:id')
  console.log('    PATCH  /api/projects/:id')
  console.log('    POST   /api/projects/:id/export')
  console.log('    POST   /api/projects/:id/generate')
  console.log('    GET    /api/projects/:id/generation-status')
  console.log('    GET    /api/projects/:id/credits')
  console.log('    GET    /api/projects/:id/excellence')
  console.log('    GET    /api/projects/:id/drawing-requirements')
  console.log('    POST   /api/projects/create-from-pdf')
  console.log('  Review:')
  console.log('    POST   /api/review/identify')
  console.log('    GET    /api/review/:token/project')
  console.log('    GET    /api/review/:token/drawings')
  console.log('    POST   /api/review/:projectId/submit')
  console.log('  Credits:')
  console.log('    PATCH  /api/credits/:id/giw-comment')
  console.log('    POST   /api/credits/:id/comment')
  console.log('  Excellence:')
  console.log('    PATCH  /api/excellence/:id/flag')
  console.log('    DELETE /api/excellence/:id')
  console.log('    DELETE /api/projects/:id/excellence')
  console.log('  Drawing Requirements:')
  console.log('    PATCH  /api/drawing-requirements/:id')
  console.log('')
})
