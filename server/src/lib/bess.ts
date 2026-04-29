const BESS_CAT_WEIGHTS: Record<string, number> = {
  'management': 4.5,
  'integrated water management': 22.5,
  'operational energy': 27.5,
  'indoor environmental quality': 16.5,
  'indoor environment quality': 16.5,
  'transport': 9,
  'waste & resource recovery': 5.5,
  'waste and resource recovery': 5.5,
  'urban ecology': 5.5,
  'innovation': 9,
  'man': 4.5,
  'iwm': 22.5,
  'oe': 27.5,
  'ieq': 16.5,
  'tra': 9,
  'trn': 9,
  'wrr': 5.5,
  'ue': 5.5,
  'inn': 9,
}

export function getCatWeight(category: string): number {
  const key = category.toLowerCase().trim()
  if (BESS_CAT_WEIGHTS[key] !== undefined) return BESS_CAT_WEIGHTS[key]
  for (const [name, w] of Object.entries(BESS_CAT_WEIGHTS)) {
    if (key.includes(name) || name.includes(key)) return w
  }
  return 0
}

// Hardcoded score thresholds (credit score %) per credit ID.
// Credits not listed default to [100] (binary achievement).
const SCORE_THRESHOLDS: Record<string, number[]> = {
  'oe 1.1':  [12.5, 37.5, 50, 100],
  'oe 1.2':  [50, 75, 100],
  'oe 4.5':  [50, 100],
  'ieq 1.1': [66, 100],
  'ieq 1.2': [66, 100],
  'ieq 2.1': [66, 100],
  'ieq 3.4': [66, 100],
  'ue 2.1':  [25, 50, 75, 100],
  // IWM 1.1, OE 2.1, OE 2.7, IEQ 1.4, IEQ 3.2 are continuous — default [100] applies
}

export function getNextThreshold(creditRef: string, currentScore: number): number {
  const key = creditRef.toLowerCase().trim()
  const thresholds = SCORE_THRESHOLDS[key] ?? [100]
  return thresholds.find(t => t > currentScore) ?? 100
}

type CreditLike = {
  id: string
  creditId: string
  category: string
  creditStatus: string
  creditWeight: number | null
  creditScore: number | null
}

type ItemLike = {
  creditId?: string | null
  creditReference: string
  currentScore?: number | null
}

export function computeItemsBessPoints<T extends ItemLike>(
  items: T[],
  weightCredits: CreditLike[],
  allCredits: CreditLike[],
): (T & { additionalBessPoints: number | null })[] {
  const eligible = weightCredits.filter(c => c.creditStatus !== 'ScopedOut' && c.creditWeight != null)

  const catGroups = new Map<string, typeof eligible>()
  for (const c of eligible) {
    if (!catGroups.has(c.category)) catGroups.set(c.category, [])
    catGroups.get(c.category)!.push(c)
  }

  const creditById = new Map(allCredits.map(c => [c.id, c]))
  const creditByRef = new Map(allCredits.map(c => [c.creditId.toLowerCase().trim(), c]))

  return items.map((item) => {
    if (item.creditReference === 'Innovation') return { ...item, additionalBessPoints: null }
    const credit = (item.creditId ? creditById.get(item.creditId) : undefined)
      ?? creditByRef.get(item.creditReference.toLowerCase().trim())
    if (!credit || credit.creditWeight == null || credit.creditStatus === 'ScopedOut') return { ...item, additionalBessPoints: null }
    const currentScore = item.currentScore ?? credit.creditScore ?? 0
    if (currentScore >= 100) return { ...item, additionalBessPoints: null }
    const catW = getCatWeight(credit.category)
    if (catW === 0) return { ...item, additionalBessPoints: null }
    const catCredits = catGroups.get(credit.category) ?? []
    if (catCredits.length === 0) return { ...item, additionalBessPoints: null }
    const nextThreshold = getNextThreshold(item.creditReference, currentScore)
    const additionalBessPoints = Math.round((nextThreshold - currentScore) * credit.creditWeight * catW / 10000 * 10) / 10
    return { ...item, additionalBessPoints }
  })
}

export function computeCurrentBESS(weightCredits: CreditLike[]): number {
  // Group all credits that have a weight (including scoped-out) by category
  const allWithWeight = weightCredits.filter(c => c.creditWeight != null)
  const catGroups = new Map<string, typeof allWithWeight>()
  for (const c of allWithWeight) {
    if (!catGroups.has(c.category)) catGroups.set(c.category, [])
    catGroups.get(c.category)!.push(c)
  }
  let bess = 0
  for (const [cat, catCredits] of catGroups) {
    const catW = getCatWeight(cat)
    if (catW === 0) continue
    // Numerator = only non-scoped credits contribute; scoped-out implicitly score 0
    const weightedScore = catCredits
      .filter(c => c.creditStatus !== 'ScopedOut')
      .reduce((s, c) => s + (c.creditScore ?? 0) * (c.creditWeight ?? 0), 0)
    // Formula: (sum(creditScore% * creditWeight%) / 100) * categoryWeight% / 100
    bess += weightedScore / 100 * catW / 100
  }
  return Math.round(bess)
}

/* ── Persist calculated BESS score to a project ── */

export async function recalculateBessScore(projectId: string): Promise<number> {
  const { prisma: db } = await import('./prisma')
  const credits = await db.credit.findMany({
    where: { projectId, deletedByGIW: false, creditId: { not: 'Innovation' } },
    select: { id: true, creditId: true, category: true, creditStatus: true, creditWeight: true, creditScore: true },
  })
  const score = computeCurrentBESS(credits)
  await db.project.update({ where: { id: projectId }, data: { bessScore: score } })
  return score
}
