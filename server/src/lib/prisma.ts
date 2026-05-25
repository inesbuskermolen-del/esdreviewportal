import { PrismaClient } from '@prisma/client'

if (!process.env.DATABASE_URL) {
  console.error('[prisma] FATAL: DATABASE_URL is not set. Add it in Railway → Variables.')
  process.exit(1)
}

const globalForPrisma = globalThis as unknown as { prisma: PrismaClient }

export const prisma = globalForPrisma.prisma ?? new PrismaClient()

if (process.env.NODE_ENV !== 'production') {
  globalForPrisma.prisma = prisma
}
