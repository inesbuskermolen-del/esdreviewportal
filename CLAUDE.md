# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

**ESD Review Portal** — a full-stack web app for GIW Environmental Solutions to manage BESS-to-ESD credit reviews.  
Frontend lives at the repo root; backend lives in `server/`.

## Development Commands

### Frontend (root)
```bash
npm install          # install frontend deps
npm run dev          # Vite dev server → http://localhost:5173
npm run build        # tsc + vite build
npm run lint         # ESLint
npm run preview      # preview production build
```

### Backend (server/)
```bash
cd server && npm install               # install backend deps
cd server && npm run dev               # ts-node-dev with auto-reload → http://localhost:3001
cd server && npm run build             # tsc → dist/
cd server && npm start                 # run compiled dist/index.js
cd server && npm run prisma:migrate    # prisma migrate dev
cd server && npm run prisma:generate   # regenerate Prisma client after schema changes
cd server && npm run prisma:studio     # open Prisma Studio
```

Both servers must run concurrently during development. Frontend proxies `/api/*` to `http://localhost:3001` via Vite config.

## Environment Setup

Copy `server/.env.example` to `server/.env` and fill in values before running the backend.  
Key variables:
- `DATABASE_URL` — PostgreSQL connection string
- `JWT_SECRET` — secret for signing admin JWT cookies
- `GIW_ADMIN_EMAIL` / `GIW_ADMIN_PASSWORD` — dev credentials (use `GIW_ADMIN_PASSWORD_HASH` for production)
- `SMTP_*` — Nodemailer config for email
- `BASE_URL` — frontend origin for CORS (default `http://localhost:5173`)
- `ANTHROPIC_API_KEY` — Claude API key

## Architecture

### Frontend (`src/`)
- React 18 + TypeScript, Vite, Tailwind CSS
- `src/App.tsx` — BrowserRouter with `AuthProvider` wrapping all routes
- `src/lib/auth.tsx` — `AuthContext`: calls `/api/auth/me` on mount; exposes `login`, `logout`, `user`
- `src/components/ProtectedRoute.tsx` — redirects to `/admin/login` if no GIW session
- `src/types/index.ts` — shared TypeScript interfaces for all domain models
- Import alias: `@/` → `./src/`

**Routes:**
| Path | Component | Auth |
|---|---|---|
| `/admin/login` | `AdminLogin` | Public |
| `/admin` | `AdminHome` | GIW only |
| `/admin/projects/new` | `NewProject` | GIW only |
| `/review/:token` | *(to be built)* | Reviewer (localStorage) |

### Backend (`server/src/`)
- Express 4 + TypeScript, Prisma ORM, PostgreSQL
- Entry: `server/src/index.ts`
- `server/src/middleware/auth.ts` — `requireGIW` middleware: validates `giw_token` httpOnly cookie
- `server/src/lib/prisma.ts` — shared Prisma client singleton
- `server/src/lib/email.ts` — Nodemailer helpers: `sendReviewInvite`, `sendSubmissionNotification`
- `server/src/middleware/upload.ts` — Multer (PDF only, 20 MB limit, memory storage)

**API routes:**
| Method | Path | Auth | Description |
|---|---|---|---|
| `POST` | `/api/auth/login` | Public | Set `giw_token` cookie |
| `POST` | `/api/auth/logout` | Public | Clear cookie |
| `GET` | `/api/auth/me` | GIW | Return current user |
| `GET` | `/api/projects` | GIW | List all projects |
| `POST` | `/api/projects` | GIW | Create project |
| `GET` | `/api/projects/:id` | GIW | Get project with credits/reviewers |
| `PATCH` | `/api/projects/:id` | GIW | Update project fields |

### Database (Prisma schema at `server/prisma/schema.prisma`)
Key models: `Project` → `Credit` → `CreditComment`, `Reviewer`, `DrawingRequirement`, `ESDExcellenceOpportunity`.  
Run `cd server && npx prisma migrate dev --name <name>` after schema changes.

## GIW Design System

All design tokens are in `tailwind.config.js` and `src/index.css`.

**Tailwind colour tokens:** `giw-charcoal` (#2C2C2C), `giw-olive` (#6B7A3B), `giw-olive-dark` (#4E5A2A), `giw-olive-light` (#E8EDD8), `giw-warm-white` (#F7F5F0), `giw-mid-grey` (#8C8C8C), `giw-border` (#D8D5CE), `giw-cream` (#EFEDE6), `giw-achieved` (#C6EFCE), `giw-scoped` (#D9D9D9), `giw-not-achieved` (#FCE4D6).

**Component classes (defined in `src/index.css`):**
- `.btn-primary` — olive fill button (Montserrat 500, 13px, 2px radius)
- `.btn-secondary` — olive outline button
- `.btn-danger` — red fill button
- `.giw-card` — cream card (giw-cream bg, giw-border border, 4px radius, 24px padding)
- `.giw-input` — standard form input

**Typography:** Montserrat (headings/nav), Open Sans (body) — loaded from Google Fonts in `index.html`.  
Nav bar: sticky, 60px, giw-charcoal background, olive dot brand mark.

## Auth Model

- **GIW admin:** JWT signed with `JWT_SECRET`, stored as `giw_token` httpOnly cookie. `requireGIW` middleware validates it server-side. Frontend reads session via `/api/auth/me`.
- **Reviewers (not yet built):** No auth — session state stored in `localStorage` keyed by review token.

## Notes

- `server/.env` contains real credentials — it is gitignored. Never commit it.
- Old `src/*.jsx` files from the previous Base44 scaffold are superseded by the new `*.tsx` files. They can be deleted.
- `vite.config.js` is superseded by `vite.config.ts` (Vite loads `.ts` first). Delete the `.js` version.
