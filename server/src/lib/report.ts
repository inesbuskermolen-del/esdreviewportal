import * as fs from 'fs'
import * as path from 'path'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import Anthropic from '@anthropic-ai/sdk'

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

const TEMPLATES_DIR = path.resolve(process.cwd(), 'templates')

const WORD_TEMPLATES: Record<string, string> = {
  'Mixed-Use':         'DATE-GIWREF-Address-SMP-MixUse-2022-A.docx',
  'Multi-Residential': 'Date-GIWREF-Address-SMP-Multi Apt-2022-A.docx',
  'Townhouse':         'Date-GIWREF-Address-SMP-TH-2022-A.docx',
  'Non-Residential':   'Date-GIWREF-Address-SMP-Comm-2022-A.docx',
}

const EXCEL_TEMPLATES: Record<string, string> = {
  'Mixed-Use': 'SMP-Visualisation-Residential-MixedUse.xlsx',
  'Multi-Residential': 'SMP-Visualisation-Residential-MixedUse.xlsx',
  'Townhouse': 'SMP-Visualisation-Residential-MixedUse.xlsx',
  'Non-Residential': 'SMP-Visualisation-NonResidential.xlsx',
}

export interface ReportProjectData {
  name: string
  address: string | null
  projectId: string | null
  date: Date | null
  revision: string | null
  typology: string | null
  client: string | null
  architect: string | null
  bessScore?: number | null
  totalDwellings?: number | null
}

export interface ReportCreditData {
  creditId: string
  creditStatus: string
  rawDataPoints: string | null
  category: string
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

function formatAddress(address: string): string {
  return address
    .replace(/,?\s*VIC\s*\d{4}\b/gi, '')   // strip ", VIC 3XXX"
    .replace(/,?\s*\b3\d{3}\b/g, '')         // strip any trailing Victorian postcode
    .replace(/\b(St|Rd|Ave|Ln|Dr|Ct|Pl|Cres|Tce|Hwy|Pde|Grv|Cct|Cl)\b\.?(?!,)(\s)/g, '$1,$2')
    .trim()
}

function extractNumber(text: string, patterns: RegExp[]): number | null {
  for (const p of patterns) {
    const m = text.match(p)
    if (m) return parseFloat(m[1].replace(/,/g, ''))
  }
  return null
}

function extractText(text: string, patterns: RegExp[]): string | null {
  for (const p of patterns) {
    const m = text.match(p)
    if (m) return m[1].trim()
  }
  return null
}

function findCredit(credits: ReportCreditData[], id: string): ReportCreditData | undefined {
  return credits.find(c => c.creditId.toLowerCase().replace(/\s+/g, ' ') === id.toLowerCase())
}

/** Like findCredit but also matches when the stored ID starts with the given prefix (e.g. 'transport 1.1' matches 'transport 1.1 resident bicycle spaces'). */
function findCreditLike(credits: ReportCreditData[], id: string): ReportCreditData | undefined {
  const lower = id.toLowerCase()
  return credits.find(c => {
    const cid = c.creditId.toLowerCase().replace(/\s+/g, ' ')
    return cid === lower || cid.startsWith(lower + ' ') || cid.startsWith(lower + '-') || cid.startsWith(lower + ':')
  })
}

/** Extract readable paragraph text from a Word document.xml string */
function extractDocParagraphs(xml: string): string[] {
  const paras: string[] = []
  for (const p of xml.split(/<\/w:p>/)) {
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(p)) !== null) texts.push(m[1])
    const line = texts.join('').trim()
    if (line) paras.push(line)
  }
  return paras
}

// ─── Claude placeholder filling ───────────────────────────────────────────────

type PlaceholderValues = Record<string, string[]>

interface FillData {
  placeholders: PlaceholderValues
  rowsToDelete: string[]
  council: string | null
}

/** Extract readable text from the first <w:tc> cell of a table row XML fragment. */
function getFirstCellText(rowXml: string): string {
  const tcMatch = rowXml.match(/<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/)
  if (!tcMatch) return ''
  const texts: string[] = []
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(tcMatch[1])) !== null) {
    if (m[1].trim()) texts.push(m[1])
  }
  return texts.join('').trim()
}

/**
 * Extract the first-cell text of every top-level <w:tr> in the document XML.
 * Uses depth-counting to skip rows inside nested tables.
 */
/** Find the next actual <w:tr> or <w:tr ...> open tag (not <w:trPr> etc.) from position `from`. */
function nextTrOpen(xml: string, from: number): number {
  let i = from
  while (i < xml.length) {
    const pos = xml.indexOf('<w:tr', i)
    if (pos === -1) return -1
    const ch = xml[pos + 5]
    if (ch === '>' || ch === ' ') return pos
    i = pos + 5
  }
  return -1
}

/** Find the next actual <w:tbl> or <w:tbl ...> open tag (not <w:tblPr>, <w:tblGrid>, etc.) from position `from`. */
function nextTblOpen(xml: string, from: number): number {
  let i = from
  while (i < xml.length) {
    const pos = xml.indexOf('<w:tbl', i)
    if (pos === -1) return -1
    const ch = xml[pos + 6]
    if (ch === '>' || ch === ' ') return pos
    i = pos + 6
  }
  return -1
}

/** Reverse search: find the last actual <w:tbl> or <w:tbl ...> open tag before position `before`. */
function lastTblOpen(xml: string, before: number): number {
  let i = before
  while (i >= 0) {
    const pos = xml.lastIndexOf('<w:tbl', i)
    if (pos === -1) return -1
    const ch = xml[pos + 6]
    if (ch === '>' || ch === ' ') return pos
    i = pos - 1
  }
  return -1
}

function extractTableRowTexts(xml: string): string[] {
  const seen = new Set<string>()
  let i = 0
  while (i < xml.length) {
    const trStart = nextTrOpen(xml, i)
    if (trStart === -1) break
    const trEnd = findRowEnd(xml, trStart)
    const rowXml = xml.slice(trStart, trEnd)
    const text = getFirstCellText(rowXml)
    if (text) seen.add(text)
    i = trEnd
  }
  return [...seen]
}

/**
 * Set the displayed text of the first dropdown whose list contains `identifyingOption`.
 * Also removes placeholder styling (grey text) so the selected value renders in black.
 */
function setDropdownContent(xml: string, identifyingOption: string, selectedValue: string): string {
  // Try exact match first, then substring match within any displayText value
  let pos = xml.indexOf(`displayText="${identifyingOption}"`)
  if (pos === -1) {
    const escaped = identifyingOption.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const m = xml.match(new RegExp(`displayText="[^"]*${escaped}[^"]*"`))
    if (!m) return xml
    pos = xml.indexOf(m[0])
    if (pos === -1) return xml
  }

  // Locate the enclosing <w:sdt> so we can also fix <w:sdtPr>
  const sdtPrStart = xml.lastIndexOf('<w:sdtPr', pos)
  const sdtStart = sdtPrStart !== -1 ? xml.lastIndexOf('<w:sdt', sdtPrStart) : -1

  const contentStart = xml.indexOf('<w:sdtContent', pos)
  if (contentStart === -1) return xml
  const contentEnd = xml.indexOf('</w:sdtContent>', contentStart) + '</w:sdtContent>'.length
  const sdtEnd = xml.indexOf('</w:sdt>', contentEnd) + '</w:sdt>'.length

  const escaped = selectedValue.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

  // Update <w:sdtContent>: replace text, strip placeholder run style, force black colour
  let content = xml.slice(contentStart, contentEnd)
  content = content.replace(/<w:t[^>]*>[^<]*<\/w:t>/, `<w:t>${escaped}</w:t>`)
  content = content.replace(/<w:rStyle w:val="PlaceholderText"\/>/g, '')
  content = content.replace(/<w:color [^>]*\/>/g, '<w:color w:val="000000"/>')
  // If no <w:color> exists inside <w:rPr>, inject one
  if (!content.includes('<w:color ')) {
    content = content.replace(/(<w:rPr>)/, '$1<w:color w:val="000000"/>')
  }

  // Remove <w:showingPlcHdr/> from <w:sdtPr> (marks content as placeholder)
  const before = sdtStart !== -1 ? xml.slice(0, sdtStart) : xml.slice(0, contentStart)
  const sdtPrBlock = sdtStart !== -1 ? xml.slice(sdtStart, contentStart) : ''
  const cleanedSdtPr = sdtPrBlock.replace(/<w:showingPlcHdr\/>/g, '')
  const after = sdtEnd !== -1 ? xml.slice(sdtEnd) : xml.slice(contentEnd)
  const sdtClose = sdtEnd !== -1 ? xml.slice(contentEnd, sdtEnd) : ''

  return before + cleanedSdtPr + content + sdtClose + after
}

/**
 * Remove the <w:sdt> dropdown wrapper after its value has been selected,
 * leaving the inner content as plain runs. Call this after setDropdownContent.
 */
function flattenDropdown(xml: string, identifyingOption: string): string {
  let pos = xml.indexOf(`displayText="${identifyingOption}"`)
  if (pos === -1) {
    const escaped = identifyingOption.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const m = xml.match(new RegExp(`displayText="[^"]*${escaped}[^"]*"`))
    if (!m) return xml
    pos = xml.indexOf(m[0])
    if (pos === -1) return xml
  }

  // Find the actual <w:sdt> open tag (not <w:sdtPr>, <w:sdtContent>, etc.)
  let sdtStart = -1
  let search = pos - 1
  while (search >= 0) {
    const p = xml.lastIndexOf('<w:sdt', search)
    if (p === -1) break
    const ch = xml[p + 6]
    if (ch === '>' || ch === ' ') { sdtStart = p; break }
    search = p - 1
  }
  if (sdtStart === -1) return xml

  // Extract inner content of <w:sdtContent>
  const contentTagStart = xml.indexOf('<w:sdtContent', pos)
  if (contentTagStart === -1) return xml
  const contentInnerStart = xml.indexOf('>', contentTagStart) + 1
  const contentEndIdx = xml.indexOf('</w:sdtContent>', contentInnerStart)
  if (contentEndIdx === -1) return xml
  const innerContent = xml.slice(contentInnerStart, contentEndIdx)

  // Find the closing </w:sdt>
  const sdtEndIdx = xml.indexOf('</w:sdt>', contentEndIdx)
  if (sdtEndIdx === -1) return xml

  return xml.slice(0, sdtStart) + innerContent + xml.slice(sdtEndIdx + 8)
}

async function getWordFillData(
  project: ReportProjectData,
  credits: ReportCreditData[],
  docParagraphs: string[],
  criteriaNames: string[],
): Promise<FillData> {
  // Credits with data points — used for placeholder filling
  const creditSummary = credits
    .filter(c => c.rawDataPoints)
    .map(c => `${c.creditId} (${c.creditStatus}): ${c.rawDataPoints}`)
    .join('\n')

  // All credits with statuses — used for row deletion decisions
  const allCreditStatus = credits
    .map(c => `${c.creditId}: ${c.creditStatus}`)
    .join('\n')

  // Include only paragraphs that contain bracketed placeholders for conciseness
  const relevantParas = docParagraphs
    .filter(p => /\[[^\]]+\]/.test(p))
    .join('\n')

  const prompt = `You are filling placeholder tags in a Sustainability Management Plan (SMP) Word document template for a building project.

The template uses square-bracket placeholders like [XX], [XXX], [XX%], [7.0], etc. Your task is to fill each one with the correct value extracted from the BESS assessment data provided.

DOCUMENT PARAGRAPHS CONTAINING PLACEHOLDERS (in document order):
${relevantParas}

PROJECT DATA:
- Project name: ${project.name}
- Address: ${project.address ?? 'Not provided'}
- BESS Score: ${project.bessScore != null ? project.bessScore + '%' : 'Not provided'}
- Client: ${project.client ?? 'Not provided'}
- Architect: ${project.architect ?? 'Not provided'}
- Revision: ${project.revision ?? 'A'}
- Date: ${project.date ? new Date(project.date).toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' }) : 'Not provided'}

BESS CREDIT DATA (creditId | status | rawDataPoints):
${creditSummary || 'No credit data available'}

ALL BESS CREDITS AND STATUSES (for row deletion):
${allCreditStatus || 'No credit data available'}

SMP REPORT TABLE CRITERIA ROWS (exact strings as they appear in the template):
${criteriaNames.join('\n')}

AVAILABLE MELBOURNE COUNCILS (pick the one matching the project address):
City of Boroondara, City of Yarra, City of Banyule, City of Bass Coast Shire, City of Darebin, City of Greater Dandenong, City of Hobsons Bay, City of Hume, City of Kingston, City of Knox, City of Manningham, City of Maribyrnong, City of Maroondah, City of Moonee Valley, City of Merri-bek, City of Port Phillip, City of Stonnington, City of Whitehorse, City of Whittlesea, City of Wyndham, City of Bayside, City of Brimbank, City of Glen Eira, City of Greater Bendigo, City of Greater Geelong, City of Melbourne, City of Monash, City of Yarra Ranges Shire

Return ONLY a valid JSON object with the following exact structure.

For the placeholder arrays: each array must contain the replacement values in DOCUMENT ORDER (i.e., the Nth element replaces the Nth occurrence of that tag). If a value cannot be determined, use the original placeholder text (e.g. "[XX]").

For "rowsToDelete": list the EXACT criteria names (from the list above) whose corresponding BESS credit is NOT claimed. A credit is not claimed if its status is "Scoped Out", "Not Achieved", "N/A", or if no matching credit exists in BESS at all. Do NOT include header rows (e.g. "Criteria", "Council Best Practice Standard", "Development Provision") or rows that don't correspond to a specific BESS credit. Use the exact criteria string from the list. IMPORTANT: Never include any Materials criteria (e.g. Embodied Energy, Structural and Reinforcing Steel, Sustainable Timber, PVC, Sustainable Products, Building Re-use) — these rows must always remain in the table.

For "council": return the EXACT council name from the list above that matches the project's suburb/address. Use your knowledge of Melbourne LGA boundaries. Return null if you cannot determine it.

{
  "XX": [<37 values in order>],
  "XXX": [<3 values — all the same rainwater tank volume, e.g. "2000">],
  "XX%": [<2 values — both the same non-residential energy improvement %, e.g. "15%">],
  "7.0": [<3 values — all the same NatHERS star rating, e.g. "7.5">],
  "XX% (XX out of XX)": [<2 values in document order>],
  "council": "City of Yarra",
  "rowsToDelete": ["exact criteria name 1", "exact criteria name 2", ...]
}

For the [XX] values in document order, the contexts are:
1. apartments count (in "will include [XX] apartments") — look in OE 2.x rawDataPoints (the dwelling profile / "Dwellings & Non-Residential Spaces" table). Find the total dwelling count, e.g. "80 apartments" or sum of bedroom type counts.
2. retail tenancies count — from OE 2.x dwelling profile rawDataPoints
3. commercial/office tenancies count — from OE 2.x dwelling profile rawDataPoints (office and commercial are equivalent)
4. number of levels/storeys — from the Building table at the start of the BESS PDF (the "Height" or "Storeys" field). Look in OE 1.x or OE 2.x rawDataPoints for patterns like "8 storeys", "8 levels", "8 floors", "Number of Storeys: 8". Return just the number (e.g. "8").
5. 1-bedroom apartments count — look in OE 2.x rawDataPoints for bedroom breakdown (e.g. "20 x 1BR", "20 x 1 bedroom", "20 x 1-bedroom apartments")
6. 2-bedroom apartments count — same source as #5
7. 3-bedroom apartments count — same source as #5
8. retail area m2
9. commercial/office area m2 (office = commercial)
10. site surface area m2 — from the front page / project inputs section of the BESS PDF (the "Site Area" field). Search all credits' rawDataPoints for patterns like "site area: 2,500m²", "2500 m² site", "total site area of 2500m²". Urban Ecology and IWM credits often reference it. Return just the number without units (e.g. "2500").
11. distance to Melbourne CBD km — use your knowledge of Melbourne geography based on the project suburb/address
12. landscape irrigation description (Key ESD Initiatives table cell, paragraph style "Landscape_Irrigation") — from IWM 3.1 rawDataPoints; if scoped out: 'Not targeted.'; if no data: 'Native vegetation has been specified for the development with no supplementary irrigation required after the establishment period.'
13. solar PV system size kW (Key ESD Initiatives table: "A [XX]kW Solar PV system is to be located on the roof of the development") — from OE 4.2 or OE 4.5 rawDataPoints
14. non-residential nominated area DF % (Key ESD Initiatives table: "2% DF to [XX]% of the nominated area") — from IEQ rawDataPoints
15. communal space m2 (Key ESD Initiatives table: "[XX]m2 of communal space") — from OE rawDataPoints
16. vegetation % of site area (Key ESD Initiatives table: "total area of vegetation is [XX]%") — from Urban Ecology rawDataPoints
17. communal food production area m2 (Key ESD Initiatives table: "[XX]m2 of communal food production area") — from Urban Ecology rawDataPoints
18. rainwater tank volume litres (SMP details: "directed into the [XX] litre rainwater tank") — from IWM 2.1 rawDataPoints; same value as [XXX] positions
19. landscape irrigation description (SMP details table "Landscape Irrigation" row) — same content as position 12
20. hot water system type (SMP details: "The development is to utilise [XX]") — from OE 2.x rawDataPoints; full phrase like 'a heat pump hot water system'
21. solar PV system size kW (SMP details section) — same value as position 13
22. solar PV expected generation kWh ("generate approximately [XX]kWh") — from OE 4.2/4.5 rawDataPoints
23. daylight % of Living areas ("[XX]% of Living areas") — from IEQ 1.3/1.4 rawDataPoints
24. daylight % of Bedrooms ("[XX]% of Bedrooms") — same source
25. non-residential nominated area DF % (SMP details, second occurrence — same value as position 14)
26. non-residential ventilation description (SMP details "Ventilation – Non-Residential" standalone cell) — from IEQ 2.3 rawDataPoints; describe the ventilation approach in one sentence
27. % of commercial areas with ceiling fans ("[XX]% of the regular use areas") — from IEQ rawDataPoints; return just the number (e.g. "50")
28. resident bicycle spaces count ("[XX] bicycle spaces are to be provided for residents") — from Transport rawDataPoints; return just the number
29. residential visitor bicycle spaces count ("[XX] bicycle spaces are to be provided for residential visitors") — same source
30. employee bicycle spaces count ("[XX] bicycle spaces are to be provided for employees") — same source
31. non-residential visitor bicycle spaces count ("[XX] bicycle spaces are to be provided for non-residential visitors") — same source
32. end of trip showers count (in "[XX] showers, [XX] lockers and changing facilities") — from Transport rawDataPoints; return just the number
33. end of trip lockers count (same sentence as position 32) — from Transport rawDataPoints; return just the number
34. motorbike/moped spaces count ("min. [XX] motorbike / moped spaces") — from Transport rawDataPoints; return just the number
35. communal space m2 (SMP details section, same value as position 15)
36. vegetation % of site area (SMP details section, same value as position 16)
37. communal food production area m2 (SMP details section, same value as position 17)

For the two [XX% (XX out of XX)] values in document order:
1. Winter sunlight result (in "[XX% (XX out of XX)] of apartments achieve at least 3 hours of sunlight"):
   - Look in IEQ 1.5 rawDataPoints for the winter sunlight percentage and apartment counts.
   - If rawDataPoints give "X out of Y" directly, use that. If only a percentage is given, calculate: compliant = round(percentage × totalApartments / 100), total = totalApartments (from [XX] position 1).
   - Format: "87% (52 out of 60)". If no data, use "[XX% (XX out of XX)]".
2. Cross-ventilation result (in "[XX% (XX out of XX)] of the development's apartments are naturally cross-ventilated"):
   - Look in IEQ 2.3 rawDataPoints for cross-ventilation percentage and counts. Same calculation method.
   - Format: "87% (52 out of 60)". If no data, use "[XX% (XX out of XX)]".

For the [XXX] values in document order (3 values, all the same rainwater tank volume):
1. rainwater tank volume litres (Key ESD Initiatives table: "A [XXX] litre rainwater tank…") — look in IWM 2.1 rawDataPoints for tank volume (e.g. "2,000 litre rainwater tank", "2000L tank"). Return just the number, e.g. "2000".
2. rainwater tank volume litres (SMP details: "A [XXX] litre rainwater tank…") — same value as #1
3. rainwater tank volume litres (SMP appendix: "…provision of a [XXX] litre rainwater tank…") — same value as #1

For the [XX%] values in document order (key "XX%" in the JSON — 2 values, both the same):
1. Non-residential energy performance improvement above BCA Section J (Key ESD Initiatives table) — look in OE 1.1 rawDataPoints for "15% above NCC", "% improvement above Section J", "% better than NCC". Return with % sign, e.g. "15%". Return "0%" if zero or not found.
2. Same value as #1 (second occurrence in SMP details)

For the [7.0] values in document order (key "7.0" in the JSON — 3 values, all the same NatHERS star rating):
1. NatHERS star rating (Key ESD Initiatives table) — look in OE 1.2 rawDataPoints for star rating. Return just the number, e.g. "7.5".
2. NatHERS star rating (SMP details) — same value as #1
3. NatHERS star rating (second SMP details occurrence) — same value as #1

Return ONLY the JSON, no explanation.`

  const message = await anthropic.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 4096,
    system: 'You are a JSON data-filling assistant. Output only valid JSON, no markdown, no commentary.',
    messages: [{ role: 'user', content: prompt }],
  })

  const raw = message.content[0]
  if (raw.type !== 'text') throw new Error('Unexpected response from Claude placeholder API')

  // Strip any accidental markdown fences
  const jsonText = raw.text.replace(/^```(?:json)?\n?/m, '').replace(/\n?```$/m, '').trim()
  try {
    const parsed = JSON.parse(jsonText)
    const rowsToDelete: string[] = Array.isArray(parsed.rowsToDelete) ? parsed.rowsToDelete : []
    const council: string | null = typeof parsed.council === 'string' ? parsed.council : null
    const xxVals: string[] = Array.isArray(parsed['XX']) ? parsed['XX'] : []
    const xxPctVals: string[] = Array.isArray(parsed['XX% (XX out of XX)']) ? parsed['XX% (XX out of XX)'] : []
    console.log('[report] Claude XX positions 12-17:', xxVals.slice(11, 17))
    console.log('[report] Claude XX% (XX out of XX) values:', xxPctVals)
    delete parsed.rowsToDelete
    delete parsed.council
    return { placeholders: parsed as PlaceholderValues, rowsToDelete, council }
  } catch {
    console.error('[report] Claude fill-data response was not valid JSON:', jsonText.slice(0, 500))
    return { placeholders: {}, rowsToDelete: [], council: null }
  }
}

/**
 * Replace "Appendix X – Title" placeholders with sequential letters (B, C, …).
 * Appendix A is already hardcoded as WSUD Response in the template.
 * Handles both single-run ("Appendix X" in one <w:t>) and split-run cases.
 */
function numberAppendices(xml: string): string {
  const SEP = '</w:p>'
  const parts = xml.split(SEP)

  // Pass 1: collect unique appendix titles in document order
  const seen = new Set<string>()
  const order: string[] = []
  for (const part of parts) {
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(part)) !== null) texts.push(m[1])
    const line = texts.join('')
    const match = line.match(/Appendix X\s*[–—-]\s*(.+?)(?:\s+\d+)?$/)
    if (!match) continue
    const raw = match[1].trim().replace(/\.\s*$/, '')
    const key = raw.toLowerCase().replace(/\W/g, '').slice(0, 10)
    if (!seen.has(key)) { seen.add(key); order.push(key) }
  }

  // Assign letters starting from B (A = WSUD Response, already in template)
  const keyToLetter = new Map(order.map((k, i) => [k, String.fromCharCode(66 + i)]))

  // Pass 2: replace X → letter in each relevant paragraph
  return parts.map(part => {
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(part)) !== null) texts.push(m[1])
    const line = texts.join('')
    const match = line.match(/Appendix X\s*[–—-]\s*(.+?)(?:\s+\d+)?$/)
    if (!match) return part
    const raw = match[1].trim().replace(/\.\s*$/, '')
    const key = raw.toLowerCase().replace(/\W/g, '').slice(0, 10)
    const letter = keyToLetter.get(key)
    if (!letter) return part
    // Single-run: "Appendix X" in one run
    let p = part.replace(/Appendix X/g, `Appendix ${letter}`)
    // Split-run: standalone "X" in its own <w:t> within this paragraph
    p = p.replace(/<w:t([^>]*)>X<\/w:t>/g, `<w:t$1>${letter}</w:t>`)
    return p
  }).join(SEP)
}

function formatBESSPercentage(raw: string, totalApts: number | null): string | null {
  const outOf = raw.match(/(\d+)\s*(?:apartments?)?\s*out\s*of\s*(\d+)/i)
             ?? raw.match(/(\d+)\s*\/\s*(\d+)/)
  const pct = raw.match(/(\d+(?:\.\d+)?)\s*%/)
  if (outOf) {
    const n = outOf[1], tot = outOf[2]
    const p = pct?.[1] ?? Math.round(parseInt(n) / parseInt(tot) * 100).toString()
    return `${p}% (${n} out of ${tot})`
  }
  if (pct && totalApts) {
    const compliant = Math.round(parseFloat(pct[1]) / 100 * totalApts)
    return `${pct[1]}% (${compliant} out of ${totalApts})`
  }
  if (pct) return `${pct[1]}%`
  return null
}

/** Combine rawDataPoints from all OE 2.x credits (dwelling profile section). */
function getHotWaterRaw(credits: ReportCreditData[]): string {
  return credits
    .filter(c => /^oe\s*2/i.test(c.creditId))
    .map(c => c.rawDataPoints ?? '')
    .filter(Boolean)
    .join(' ')
}

function extractHwDescription(raw: string): string | null {
  const r = raw.toLowerCase()
  if (/heat pump/.test(r)) return 'a heat pump hot water system'
  if (/(central|communal|common).*gas|gas.*(central|communal)/.test(r)) return 'a centralised gas hot water system'
  if (/solar/.test(r)) return 'a solar hot water system'
  if (/instant.*electric|electric.*instant/.test(r)) return 'individual electric instantaneous hot water systems'
  if (/gas/.test(r)) return 'a gas hot water system'
  if (/electric/.test(r)) return 'an electric hot water system'
  return null
}

/**
 * Extract the total dwelling count from rawDataPoints.
 * Searches OE 2.x credits first, then all credits as fallback.
 */
function getTotalApartments(credits: ReportCreditData[]): number | null {
  const oe2Raw = credits
    .filter(c => /^oe\s*2/i.test(c.creditId))
    .map(c => c.rawDataPoints ?? '')
    .join(' ')
  const allRaw = credits.map(c => c.rawDataPoints ?? '').join(' ')

  for (const raw of [oe2Raw, allRaw]) {
    if (!raw.trim()) continue

    // Direct total patterns (preferred over bedroom sums to avoid counting non-residential)
    for (const p of [
      /total\s+(?:of\s+)?(\d+)\s*(?:residential\s+)?(?:apartment|dwelling|unit)s?\b/i,
      /total[:\s]+(\d+)\s*(?:apartment|dwelling|unit)/i,
      /(\d+)\s*residential\s+(?:apartment|dwelling|unit)s?\b/i,
      /(\d+)\s*(?:total\s+)?(?:apartment|dwelling|unit)s?\b(?!\s*(?:per|m2|m²|\/))/i,
    ]) {
      const m = raw.match(p)
      if (m) return parseInt(m[1])
    }

    // Sum "N x 1BR / N x 2BR" style counts (require at least 2 types to avoid false positives)
    const xCounts = [...raw.matchAll(/(\d+)\s*[×x]\s*(?:\d+[\s-]?(?:bed|BR|bedroom)|studio)/gi)]
      .map(m => parseInt(m[1]))
    if (xCounts.length >= 2) return xCounts.reduce((a, b) => a + b, 0)

    // Sum "1 Bed: 20, 2 Bed: 40" style counts (require at least 2 types)
    const colonCounts = [...raw.matchAll(/(?:\d+[\s-]?(?:bed|BR|bedroom)|studio)\s*:\s*(\d+)/gi)]
      .map(m => parseInt(m[1]))
    if (colonCounts.length >= 2) return colonCounts.reduce((a, b) => a + b, 0)
  }
  return null
}

/**
 * Replace [XX] inside the second table cell of the first row whose first-cell text
 * matches headerPattern. Used for page-4 project-data tables where the label and
 * value are in adjacent cells rather than inline prose.
 */
function fillTableCellByHeader(xml: string, headerPattern: RegExp, value: string): string {
  let i = 0
  while (i < xml.length) {
    const trStart = nextTrOpen(xml, i)
    if (trStart === -1) break
    const trEnd = findRowEnd(xml, trStart)
    const rowXml = xml.slice(trStart, trEnd)
    const firstCell = getFirstCellText(rowXml)
    if (headerPattern.test(firstCell) && rowXml.includes('[XX]')) {
      const fixed = rowXml.replace('[XX]', () => value)
      return xml.slice(0, trStart) + fixed + xml.slice(trEnd)
    }
    i = trEnd
  }
  return xml
}

/**
 * Detect whether the project has retail/commercial areas from OE 2.x dwelling profile.
 * Falls back to typology when rawDataPoints don't mention them.
 */
function detectNonResidential(credits: ReportCreditData[], typology: string | null): { hasRetail: boolean; hasCommercial: boolean } {
  const oe2Raw = credits.filter(c => /^oe\s*2/i.test(c.creditId)).map(c => c.rawDataPoints ?? '').join(' ')
  const t = (typology ?? '').toLowerCase()
  // Pure residential — never has retail/commercial
  if (t === 'multi-residential' || t === 'townhouse') return { hasRetail: false, hasCommercial: false }
  if (t === 'non-residential') return { hasRetail: /retail/i.test(oe2Raw), hasCommercial: true }
  // Mixed-use / unknown: trust rawDataPoints
  const hasRetail = /retail/i.test(oe2Raw)
  const hasCommercial = /commercial|office/i.test(oe2Raw) // office = commercial
  return { hasRetail, hasCommercial }
}

/**
 * Remove retail/commercial line items from page 4 when not present in the project.
 * Handles:
 *  - "[XX]m2 retail" / "[XX]m2 commercial" standalone paragraphs
 *  - The composite sentence "will include [XX] apartments, [XX] retail tenancies and [XX] commercial tenancies …"
 */
function removeNonResidentialLines(xml: string, hasRetail: boolean, hasCommercial: boolean): string {
  // ── Delete area paragraphs ──────────────────────────────────────────────
  if (!hasRetail) {
    xml = deleteParagraphsByText(xml, [/^(?:\[XX\]|\d[\d,.]*)m2\s+retail$/i])
  }
  if (!hasCommercial) {
    xml = deleteParagraphsByText(xml, [/^(?:\[XX\]|\d[\d,.]*)m2\s+(?:commercial|office)$/i])
  }

  if (hasRetail && hasCommercial) return xml

  // ── Modify the "will include … apartments, … retail tenancies and … commercial tenancies …" sentence ──
  // Operate at combined-text level so we precisely target the retail/commercial substrings
  // without accidentally matching the apartment count (which precedes them in the same paragraph).
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g; let m
    while ((m = re.exec(para)) !== null) texts.push(m[1])
    const combined = texts.join('')
    if (!/will include/i.test(combined)) return para
    if (!combined.includes('retail tenancies') && !combined.includes('commercial tenancies')) return para

    const changes: Array<{ pos: number; len: number; replacement: string }> = []

    if (!hasRetail && !hasCommercial) {
      // Remove ", M retail tenancies and P commercial tenancies" — keep " constructed over"
      const hit = combined.match(/,\s*(?:\[XX\]|\d+)\s*retail tenancies and\s*(?:\[XX\]|\d+)\s*commercial tenancies/)
      if (hit?.index !== undefined) changes.push({ pos: hit.index, len: hit[0].length, replacement: '' })
    } else if (!hasRetail) {
      // Remove ", M retail tenancies and " — replace with " and " to keep commercial
      const hit = combined.match(/,\s*(?:\[XX\]|\d+)\s*retail tenancies and\s*/)
      if (hit?.index !== undefined) changes.push({ pos: hit.index, len: hit[0].length, replacement: ' and ' })
    } else {
      // !hasCommercial — remove " and P commercial tenancies"
      const hit = combined.match(/\s+and\s+(?:\[XX\]|\d+)\s+commercial tenancies/)
      if (hit?.index !== undefined) changes.push({ pos: hit.index, len: hit[0].length, replacement: '' })
    }

    if (!changes.length) return para
    return applyParaChanges(para, changes)
  })
}

/**
 * Post-render fallback: replace unfilled placeholders using rawDataPoints directly.
 * Handles cases where Claude's positional counter misfires for these specific values.
 */
function applyBESSFallbacks(xml: string, credits: ReportCreditData[], project: ReportProjectData): string {
  // Prefer the stored totalDwellings field; fall back to regex extraction from rawDataPoints
  const totalApts = project.totalDwellings ?? getTotalApartments(credits)
  console.log('[report] applyBESSFallbacks: totalApts=', totalApts, 'project.name=', project.name)

  // ── Apartments count ─────────────────────────────────────────────────────
  if (totalApts !== null) {
    const aptStr = String(totalApts)
    // Target only [XX] immediately preceding "apartments" in the combined text.
    // Using applyParaChanges handles [XX] split across multiple runs.
    // Searching for "[XX] apartments" prevents accidental replacement of other [XX] values
    // (e.g. levels count) that may appear later in the same paragraph.
    const aptKws = ['will include', 'comprises', 'containing', 'consist of', 'development of']
    xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
      const texts: string[] = []
      const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
      let m: RegExpExecArray | null
      while ((m = re.exec(para)) !== null) texts.push(m[1])
      const combined = texts.join('')
      if (!aptKws.some(kw => combined.includes(kw))) return para
      const hit = combined.match(/\[XX\]\s*apartments/)
      if (!hit || hit.index === undefined) return para
      return applyParaChanges(para, [{ pos: hit.index, len: 4, replacement: aptStr }])
    })
    // Table cell: row header "Total Dwellings", "Number of Dwellings", etc.
    for (const pat of [/total\s+dwellings?/i, /number\s+of\s+(?:dwellings?|apartments?|units?)/i, /dwellings?\s+\(/i]) {
      const prev = xml
      xml = fillTableCellByHeader(xml, pat, aptStr)
      if (xml !== prev) break
    }
  }

  // ── Number of levels ─────────────────────────────────────────────────────
  {
    // Include project name in search — BESS PDF often names the project "8 Storey Residential"
    const searchCorpus = (project.name ?? '') + ' ' + credits.map(c => c.rawDataPoints ?? '').join(' ')
    const levelsMatch =
      searchCorpus.match(/(\d+)\s*(?:storey|storeys|level|levels|floor|floors)\b/i)
      ?? searchCorpus.match(/(?:number of (?:storeys?|levels?|floors?)[:\s]+)(\d+)/i)
      ?? searchCorpus.match(/(?:height|building)[^:\n]*:\s*(?:[\d.]+ ?m[,\s]+)?(\d+)[\s-]?(?:storey|level|floor)/i)
      ?? searchCorpus.match(/(\d+)[\s-](?:storey|level|floor)\s+(?:building|development|residential|mixed)/i)
    const val = levelsMatch?.[1] ?? null

    if (val !== null) {
      // Paragraph-level replacement for "constructed over [XX]" / "comprising [XX]"
      const levelsKws = ['constructed over', 'comprising']
      xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
        const texts: string[] = []
        const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
        let m: RegExpExecArray | null
        while ((m = re.exec(para)) !== null) texts.push(m[1])
        const combined = texts.join('')
        if (!levelsKws.some(kw => combined.includes(kw))) return para
        if (!combined.includes('[XX]')) return para
        return para.replace(/(<w:t[^>]*>)([^<]*)\[XX\]([^<]*)(<\/w:t>)/, `$1$2${val}$3$4`)
      })
      // Table cell: row header "Levels", "Number of Levels", "Storeys", etc.
      for (const pat of [/number\s+of\s+(?:levels?|storeys?|floors?)/i, /^levels?$/i, /^storeys?$/i, /^height/i]) {
        const prev = xml
        xml = fillTableCellByHeader(xml, pat, val)
        if (xml !== prev) break
      }
    }
  }

  // ── OE 1.1: fill or remove thermal performance improvement % ─────────────
  {
    const oe11Raw = findCredit(credits, 'oe 1.1')?.rawDataPoints ?? ''
    const pctMatch = oe11Raw.match(/(\d+(?:\.\d+)?)\s*%\s*(?:improvement|reduction|above|better)/i) ??
                     oe11Raw.match(/(?:improvement|reduction)[^:]*:\s*(\d+(?:\.\d+)?)\s*%/i)
    const improvement = pctMatch ? parseFloat(pctMatch[1]) : 0

    // Case 1: [by XX%] as a continuous string (template tag named "by XX%")
    if (xml.includes('[by XX%]')) {
      xml = improvement === 0
        ? xml.replace(/\[by XX%\]/g, '')
        : xml.replace(/\[by XX%\]/g, `by ${improvement}%`)
    }

    // Case 2: paragraph-level — covers "by [XX%]" split across runs, and "by 0%" from Claude
    // When improvement is 0 remove the value; when >0 fill any remaining [XX%] placeholder.
    xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
      const texts: string[] = []
      const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
      let m: RegExpExecArray | null
      while ((m = re.exec(para)) !== null) texts.push(m[1])
      const combined = texts.join('')
      const hasByValue = /\bby\s+(?:\[XX%\]|0%)\b/.test(combined)
      if (!hasByValue) return para
      if (!/(?:ncc|section j|above|thermal|envelope|J1)/i.test(combined)) return para
      if (improvement === 0) {
        return para.replace(/(<w:t[^>]*>)([^<]*)(<\/w:t>)/g, (_, open, content, close) => {
          const cleaned = content.replace(/\[XX%\]/g, '').replace(/\b0%\b/g, '').replace(/\s{2,}/g, ' ')
          return `${open}${cleaned}${close}`
        })
      } else {
        return para.replace(/(<w:t[^>]*>)([^<]*)(<\/w:t>)/g, (_, open, content, close) =>
          `${open}${content.replace(/\[XX%\]/g, `${improvement}%`)}${close}`)
      }
    })

    // Case 3: delete entire paragraphs where "0%" or "[XX%]" sits in a thermal/NCC sentence
    // (covers templates that don't use "by" before the placeholder)
    if (improvement === 0) {
      xml = deleteParagraphsByText(xml, [
        /(?:0%|(?:\[XX%\]))\s*(?:above|improvement|better)[^.]*(?:section j|ncc|bca)/i,
        /(?:section j|ncc|bca)[^.]*(?:0%|(?:\[XX%\]))\s*(?:above|improvement|better)/i,
        /non[- ]residential[^.]*(?:0%|(?:\[XX%\]))[^.]*(?:above|section j|ncc)/i,
      ])
    }
  }

  // ── Site area ─────────────────────────────────────────────────────────────
  {
    const allRaw = credits.map(c => c.rawDataPoints ?? '').join(' ')
    const siteMatch = allRaw.match(/site\s+area[^:]*:?\s*([\d,]+)\s*m/i) ??
                      allRaw.match(/([\d,]+)\s*m2?\s+site/i) ??
                      allRaw.match(/total\s+site[^:]*:?\s*([\d,]+)/i) ??
                      allRaw.match(/site[^:]*:\s*([\d,]+)/i)
    if (siteMatch) {
      const val = siteMatch[1].replace(/,/g, '')
      // Inline prose: "surface area of [XX]"
      for (const kw of ['surface area of', 'site area of', 'area of']) {
        const idx = xml.indexOf(kw)
        if (idx === -1) continue
        const w = xml.slice(idx, idx + 300)
        if (!w.includes('[XX]')) continue
        xml = xml.slice(0, idx) + w.replace('[XX]', val) + xml.slice(idx + 300)
        break
      }
      // Table cell: row header "Site Area", "Total Site Area", etc.
      for (const pat of [/site\s+area/i, /total\s+site/i]) {
        const prev = xml
        xml = fillTableCellByHeader(xml, pat, val)
        if (xml !== prev) break
      }
    }
  }

  // ── Hot water: "to utilise [XX]" ─────────────────────────────────────────
  // Uses paragraph-level search to catch all occurrences (one may be a dropdown,
  // another may be a plain [XX] placeholder — indexOf only finds the first).
  {
    const hwRaw = getHotWaterRaw(credits)
    const desc = hwRaw ? extractHwDescription(hwRaw) : null
    if (desc) {
      const escaped = escapeXml(desc)
      xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
        const texts: string[] = []
        const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
        let m: RegExpExecArray | null
        while ((m = re.exec(para)) !== null) texts.push(m[1])
        const combined = texts.join('')
        if (!/to utilise/i.test(combined) || !combined.includes('[XX]')) return para
        return para.replace(/(<w:t[^>]*>)([^<]*)\[XX\]([^<]*)(<\/w:t>)/, `$1$2${escaped}$3$4`)
      })
    }
  }

  // ── Landscape irrigation: [XX] in SMP table row ─────────────────────────
  {
    const iwm31 = findCredit(credits, 'iwm 3.1')
    let landDesc: string
    if (iwm31?.rawDataPoints) {
      landDesc = iwm31.rawDataPoints.replace(/^project:\s*\w+\s*-\s*/i, '').trim()
    } else if (iwm31?.creditStatus === 'ScopedOut') {
      landDesc = 'Not targeted.'
    } else {
      landDesc = 'Native vegetation has been specified for the development with no supplementary irrigation required after the establishment period.'
    }
    xml = fillTableCellByHeader(xml, /^landscape irrigation$/i, landDesc)
  }

  // ── Winter sunlight: [XX% (XX out of XX)] of apartments achieve … ────────
  if (xml.includes('[XX% (XX out of XX)]')) {
    const ieq15 = findCredit(credits, 'ieq 1.5')
    if (ieq15?.rawDataPoints) {
      const fmt = formatBESSPercentage(ieq15.rawDataPoints, totalApts)
      if (fmt) {
        // Only replace the first occurrence (winter sunlight comes before ventilation)
        xml = xml.replace('[XX% (XX out of XX)]', () => fmt)
      }
    }
  }

  // ── Natural ventilation: second [XX% (XX out of XX)] ─────────────────────
  if (xml.includes('[XX% (XX out of XX)]')) {
    const ieq23 = findCredit(credits, 'ieq 2.3')
    if (ieq23?.rawDataPoints) {
      const fmt = formatBESSPercentage(ieq23.rawDataPoints, totalApts)
      if (fmt) xml = xml.replace('[XX% (XX out of XX)]', () => fmt)
    }
  }

  // ── IEQ 2.3: CO2 ppm — fill 'concentrations below [XX]ppm' ─────────────────
  {
    const co2Idx = xml.indexOf('concentrations below')
    if (co2Idx !== -1) {
      const window = xml.slice(co2Idx, co2Idx + 500)
      if (window.includes('[XX]')) {
        const ieq23 = findCredit(credits, 'ieq 2.3')
        const co2Match = ieq23?.rawDataPoints?.match(/(\d{3,4})\s*ppm/i) ??
                         ieq23?.rawDataPoints?.match(/CO2[^:]*:?\s*(\d{3,4})/i)
        if (co2Match) {
          const fixed = window.replace('[XX]', co2Match[1])
          xml = xml.slice(0, co2Idx) + fixed + xml.slice(co2Idx + 500)
        }
      }
    }
  }

  // ── Transport: bicycle spaces + end-of-trip showers/lockers (fallback) ──────
  // Claude fills these from XX array positions 28–33; this catches any that remain [XX].
  if (xml.includes('[XX]')) {
    const transportRaw = credits
      .filter(c => /^transport/i.test(c.creditId))
      .map(c => c.rawDataPoints ?? '')
      .join('\n')

    if (transportRaw.trim()) {
      const residentBike = extractNumber(transportRaw, [
        /(\d+)\s*resident\s+bicycle/i,
        /resident\s+bicycle[^.\n]*?:\s*(\d+)/i,
        /(\d+)\s*space[^.\n]*?for\s+residents/i,
      ])
      const visitorResBike = extractNumber(transportRaw, [
        /(\d+)\s*residential\s+visitor\s+bicycle/i,
        /(\d+)\s*visitor\s+bicycle[^.\n]*?resident/i,
        /residential\s+visitor\s+bicycle[^.\n]*?:\s*(\d+)/i,
      ])
      const employeeBike = extractNumber(transportRaw, [
        /(\d+)\s*(?:employee|staff)\s+bicycle/i,
        /(?:employee|staff)\s+bicycle[^.\n]*?:\s*(\d+)/i,
        /bicycle[^.\n]*?for\s+(?:employees?|staff)[^.\n]*?:\s*(\d+)/i,
      ])
      const nrVisitorBike = extractNumber(transportRaw, [
        /(\d+)\s*non[- ]residential\s+visitor\s+bicycle/i,
        /non[- ]residential\s+visitor\s+bicycle[^.\n]*?:\s*(\d+)/i,
      ])
      const showers = extractNumber(transportRaw, [
        /(\d+)\s*shower/i,
        /shower[^.\n]*?:\s*(\d+)/i,
      ])
      const lockers = extractNumber(transportRaw, [
        /(\d+)\s*locker/i,
        /locker[^.\n]*?:\s*(\d+)/i,
      ])

      // Each entry: [value, keywords that uniquely identify the paragraph]
      const fills: Array<[number | null, RegExp[]]> = [
        [residentBike,  [/\bresident\s+bicycle/i]],
        [visitorResBike,[/\bresidential\s+visitor\s+bicycle/i, /visitor\s+bicycle[^.]{0,60}resident/i]],
        [employeeBike,  [/bicycle[^.]{0,60}(?:employee|staff)/i, /(?:employee|staff)[^.]{0,60}bicycle/i]],
        [nrVisitorBike, [/non[- ]residential[^.]{0,60}visitor[^.]{0,60}bicycle/i]],
        [showers,       [/\bend[- ]of[- ]trip[^.]{0,40}shower/i, /\bshower\b/i]],
        [lockers,       [/\bend[- ]of[- ]trip[^.]{0,40}locker/i, /\blocker\b/i]],
      ]

      for (const [count, kws] of fills) {
        if (count === null) continue
        let filled = false
        xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
          if (filled) return para
          const texts: string[] = []
          const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
          let m: RegExpExecArray | null
          while ((m = re.exec(para)) !== null) texts.push(m[1])
          const combined = texts.join('')
          if (!combined.includes('[XX]')) return para
          if (!kws.some(k => k.test(combined))) return para
          filled = true
          return para.replace(/(<w:t[^>]*>)([^<]*)\[XX\]([^<]*)(<\/w:t>)/, `$1$2${count}$3$4`)
        })
      }
    }
  }

  return xml
}

/**
 * Insert <w:pageBreakBefore/> into Heading 2 paragraphs that follow the
 * "ESD Assessment" heading, so each BESS category section starts on a new page.
 * The "ESD Assessment" heading itself (and any headings before it) are left alone.
 */
function addCategoryPageBreaks(xml: string): string {
  const SEP = '</w:p>'
  let pastEsdAssessment = false

  // Headings that always start on a new page regardless of style
  const ALWAYS_NEW_PAGE = /^(?:Materials|Waste\s*(?:&|and)\s*Resources?\s*Recovery)$/i

  const injectBreak = (para: string): string => {
    if (para.includes('<w:pageBreakBefore')) return para
    if (para.includes('<w:pPr>')) {
      if (/<w:pStyle\b[^>]*\/>/.test(para)) {
        return para.replace(/<w:pStyle\b[^>]*\/>/, m => m + '<w:pageBreakBefore/>')
      }
      return para.replace('<w:pPr>', '<w:pPr><w:pageBreakBefore/>')
    }
    return para.replace(/(<w:r[ >])/, (_, run: string) => `<w:pPr><w:pageBreakBefore/></w:pPr>${run}`)
  }

  return xml.split(SEP).map(para => {
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(para)) !== null) texts.push(m[1])
    const line = texts.join('').trim()

    // These headings always start on a new page regardless of their style
    if (ALWAYS_NEW_PAGE.test(line)) return injectBreak(para)

    // For all other headings: only process Heading2 paragraphs after "ESD Assessment"
    if (!para.includes('w:val="Heading2"')) return para

    if (!pastEsdAssessment) {
      if (line === 'ESD Assessment') pastEsdAssessment = true
      return para
    }

    return injectBreak(para)
  }).join(SEP)
}

/**
 * Extract the table containing identifyingText from the XML, replacing it with
 * a sentinel. Returns the modified XML and a restore function that puts the
 * original table back. Used to shield specific tables from post-processing.
 */
function shieldTable(xml: string, identifyingText: string): { xml: string; restore: (x: string) => string } {
  const textIdx = xml.indexOf(identifyingText)
  if (textIdx === -1) return { xml, restore: x => x }

  const tblStart = lastTblOpen(xml, textIdx)
  if (tblStart === -1) return { xml, restore: x => x }

  let depth = 0
  let i = tblStart
  let tblEnd = -1
  while (i < xml.length) {
    const openIdx = nextTblOpen(xml, i)
    const closeIdx = xml.indexOf('</w:tbl>', i)
    if (closeIdx === -1) break
    if (openIdx !== -1 && openIdx < closeIdx) { depth++; i = openIdx + 6 }
    else { depth--; if (depth === 0) { tblEnd = closeIdx + 8; break } i = closeIdx + 8 }
  }
  if (tblEnd === -1) return { xml, restore: x => x }

  const SENTINEL = '___KEY_INIT_TABLE_PLACEHOLDER___'
  const tableXml = xml.slice(tblStart, tblEnd)
  return {
    xml: xml.slice(0, tblStart) + SENTINEL + xml.slice(tblEnd),
    restore: (x: string) => x.replace(SENTINEL, () => tableXml),
  }
}

/**
 * Remove paragraphs whose full text matches any of the given patterns.
 */
function deleteParagraphsByText(xml: string, patterns: RegExp[]): string {
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(para)) !== null) texts.push(m[1])
    const combined = texts.join('').trim()
    return patterns.some(p => p.test(combined)) ? '' : para
  })
}

/**
 * Delete the paragraph immediately following the first paragraph whose text
 * matches contextPattern. Used when two paragraphs share identical trailing text
 * and must be distinguished by their preceding count sentence.
 */
function deleteParagraphAfterMatch(xml: string, contextPattern: RegExp): string {
  const SEP = '</w:p>'
  const parts = xml.split(SEP)
  let skipNext = false
  const kept: string[] = []
  for (const part of parts) {
    if (skipNext) { skipNext = false; continue }
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(part)) !== null) texts.push(m[1])
    if (contextPattern.test(texts.join('').trim())) skipNext = true
    kept.push(part)
  }
  return kept.join(SEP)
}

/**
 * Find the end of a <w:tr> element using depth-counting, handling nested tables.
 * Returns the index just after the matching </w:tr>.
 */
function findRowEnd(xml: string, trStart: number): number {
  let depth = 0
  let i = trStart
  while (i < xml.length) {
    const nextOpen = nextTrOpen(xml, i)
    const nextClose = xml.indexOf('</w:tr>', i)
    if (nextClose === -1) return xml.length
    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth++
      i = nextOpen + 5
    } else {
      depth--
      if (depth === 0) return nextClose + 7
      i = nextClose + 7
    }
  }
  return xml.length
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Build a new <w:tr> element by cloning structural properties from a template row
 * (row properties, cell properties, run properties) and replacing text content.
 * Preserves column widths, borders, shading and font formatting.
 */
function buildTableRow(templateRowXml: string, col1Text: string, col2Text: string): string {
  const trPr = (templateRowXml.match(/<w:trPr\b[\s\S]*?<\/w:trPr>/) ?? [])[0] ?? ''

  const cellStructures: Array<{ tcPr: string; rPr: string; pPr: string }> = []
  const cellRe = /<w:tc\b[\s\S]*?<\/w:tc>/g
  let cm: RegExpExecArray | null
  while ((cm = cellRe.exec(templateRowXml)) !== null) {
    const c = cm[0]
    cellStructures.push({
      tcPr: (c.match(/<w:tcPr\b[\s\S]*?<\/w:tcPr>/) ?? [])[0] ?? '',
      rPr:  (c.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/)   ?? [])[0] ?? '',
      pPr:  (c.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/)   ?? [])[0] ?? '',
    })
  }

  const makeCell = (idx: number, text: string): string => {
    const { tcPr, rPr, pPr } = cellStructures[idx] ?? { tcPr: '', rPr: '', pPr: '' }
    return `<w:tc>${tcPr}<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:tc>`
  }

  const cells = [makeCell(0, col1Text), makeCell(1, col2Text)]
  for (let i = 2; i < cellStructures.length; i++) cells.push(makeCell(i, ''))
  return `<w:tr>${trPr}${cells.join('')}</w:tr>`
}

/**
 * For each BESS innovation credit:
 *  - If a matching row already exists in the SMP table, fill its description cell from rawDataPoints.
 *  - If no matching row exists, insert a new row after the last Innovation row.
 */
function insertInnovationRows(xml: string, credits: ReportCreditData[]): string {
  const innovationCredits = credits
    .filter(c => /^innovation/i.test(c.category ?? '') || /^innovation/i.test(c.creditId))
    .filter(c => c.rawDataPoints?.trim())

  if (!innovationCredits.length) return xml

  // Collect all table rows
  interface RowInfo { text: string; start: number; end: number; rowXml: string }
  const collectRows = (src: string): RowInfo[] => {
    const out: RowInfo[] = []
    let pos = 0
    while (pos < src.length) {
      const s = nextTrOpen(src, pos)
      if (s === -1) break
      const e = findRowEnd(src, s)
      out.push({ text: getFirstCellText(src.slice(s, e)).trim(), start: s, end: e, rowXml: src.slice(s, e) })
      pos = e
    }
    return out
  }

  let result = xml
  const unmatched: ReportCreditData[] = []

  for (const credit of innovationCredits) {
    const creditLower = credit.creditId.toLowerCase()
    const rows = collectRows(result)
    const match = rows.find(r => {
      const t = r.text.toLowerCase()
      return t === creditLower || t.startsWith(creditLower + ' ') || creditLower.startsWith(t + ' ')
    })
    if (match) {
      // Fill [XX] placeholder in second cell if present
      result = fillTableCellByHeader(result, new RegExp(escapeRegex(match.text), 'i'), (credit.rawDataPoints ?? '').trim())
    } else {
      unmatched.push(credit)
    }
  }

  if (!unmatched.length) return result

  // Re-scan to find the last Innovation row (insertion point for new rows)
  const innovRows = collectRows(result).filter(r => /^innovation/i.test(r.text))
  const lastInnovRow = innovRows[innovRows.length - 1]
  if (!lastInnovRow) {
    console.warn('[report] insertInnovationRows: no existing Innovation rows found; skipping')
    return result
  }

  const newRowsXml = unmatched
    .map(c => buildTableRow(lastInnovRow.rowXml, c.creditId, (c.rawDataPoints ?? '').trim()))
    .join('')

  return result.slice(0, lastInnovRow.end) + newRowsXml + result.slice(lastInnovRow.end)
}

/**
 * Remove table rows whose first-cell text matches any entry in criteriaToDelete.
 * Uses depth-counting to correctly handle rows that contain nested tables.
 */
function deleteTableRows(xml: string, criteriaToDelete: string[]): string {
  if (criteriaToDelete.length === 0) return xml
  const deleteSet = new Set(criteriaToDelete.map(s => s.trim().toLowerCase()))
  const orphanedBookmarkIds = new Set<string>()
  let result = ''
  let i = 0
  while (i < xml.length) {
    const trStart = nextTrOpen(xml, i)
    if (trStart === -1) { result += xml.slice(i); break }
    result += xml.slice(i, trStart)
    const trEnd = findRowEnd(xml, trStart)
    const rowXml = xml.slice(trStart, trEnd)
    const firstCell = getFirstCellText(rowXml).toLowerCase()
    if (deleteSet.has(firstCell)) {
      const bsRe = /<w:bookmarkStart[^>]*\bw:id="(\d+)"/g
      let bsM: RegExpExecArray | null
      while ((bsM = bsRe.exec(rowXml)) !== null) orphanedBookmarkIds.add(bsM[1])
    } else {
      result += rowXml
    }
    i = trEnd
  }
  if (orphanedBookmarkIds.size > 0) {
    result = result.replace(/<w:bookmarkEnd[^>]*\bw:id="(\d+)"[^/]*(\/?>)/g, (match, id) =>
      orphanedBookmarkIds.has(id) ? '' : match,
    )
  }
  return result
}

// ─── Innovation table helpers ─────────────────────────────────────────────────

/**
 * Parse "Initiatives: Name (N points) — description. Name2 ..." into structured items.
 */
function parseInnovationPoints(raw: string): Array<{ name: string; description: string }> {
  const text = raw
    .replace(/^Initiatives:\s*/i, '')
    .replace(/\s*Total\s+points?\s+claimed:?[^.]*\.\s*$/i, '')
    .trim()

  // Split around "(N points... ) — " markers; produces [name1, desc1+name2, desc2+name3, ..., descN]
  const parts = text.split(/\s*\(\d+\s*points?\b[^)]*\)\s*[—–\-]\s*/)
  const results: Array<{ name: string; description: string }> = []

  for (let i = 0; i < parts.length - 1; i++) {
    // For i=0, the whole parts[0] is the first name.
    // For i>0, parts[i] = "prev description. nextName" — take the last sentence.
    const name = (i === 0 ? parts[0] : parts[i].replace(/^[\s\S]*\.\s+/, '')).trim()
    const rawDesc = parts[i + 1]
    // Description ends before the next item's name (which is the last ". Word" in rawDesc, for non-final items)
    const desc = (i + 1 < parts.length - 1
      ? rawDesc.replace(/\.\s+[^.]+$/, '')
      : rawDesc
    ).trim().replace(/\.\s*$/, '').trim()
    if (name) results.push({ name, description: desc })
  }
  return results
}

/** Replace the text content of a single table cell, keeping the first run's style. */
function setCellText(cellXml: string, newText: string): string {
  const escaped = escapeXml(newText)
  let first = true
  return cellXml.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, (match) => {
    if (first) { first = false; return newText ? `<w:t xml:space="preserve">${escaped}</w:t>` : '<w:t></w:t>' }
    return ''
  })
}

/** Clone a template data row, filling in name (cell 0), clearing objective (cell 1), setting description (cell 2). */
function buildInnovationRow(templateRowXml: string, name: string, description: string): string {
  const cells = templateRowXml.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) || []
  if (cells.length < 3 || !cells[0] || !cells[1] || !cells[2]) return templateRowXml
  return templateRowXml
    .replace(cells[0], setCellText(cells[0], name))
    .replace(cells[1], setCellText(cells[1], ''))
    .replace(cells[2], setCellText(cells[2], description))
}

/**
 * Sync the Innovation "Council Best Practice Standard" table with BESS innovation points:
 * - Delete template rows whose criteria name has no matching BESS point
 * - Keep template rows that do match (preserving their Development Provision text)
 * - Append new rows for BESS points not already in the template
 */
function syncInnovationTable(xml: string, bessPoints: Array<{ name: string; description: string }>): string {
  // Locate the table by finding the "innovative technology" council objective paragraph
  const markerIdx = xml.toLowerCase().indexOf('innovative technology')
  if (markerIdx === -1) return xml
  const tblStart = xml.indexOf('<w:tbl', markerIdx)
  if (tblStart === -1) return xml
  const tblEnd = xml.indexOf('</w:tbl>', tblStart) + 8
  if (tblEnd < 8) return xml

  const tableXml = xml.slice(tblStart, tblEnd)
  if (!/Council Best Practice Standard/i.test(tableXml)) return xml

  const rows = tableXml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) || []
  if (rows.length < 2) return xml

  const headerRows = rows.slice(0, 2)     // "Council Best Practice Standard" + "Criteria/Development Provision"
  const dataRows   = rows.slice(2)
  const templateRow = dataRows[0] || ''   // First data row used as XML template for new rows

  const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim()

  const templateItems = dataRows.map(row => {
    const cells = row.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) || []
    const name = cells[0]
      ? (cells[0].match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join('')
      : ''
    return { name: name.trim(), xml: row }
  })

  const usedTemplateIndices = new Set<number>()
  const newDataRows: string[] = []

  console.log(`[report] Innovation table: ${dataRows.length} template data rows, ${bessPoints.length} BESS points`)

  for (const bess of bessPoints) {
    const normBess = normalize(bess.name)
    const tmplIdx = templateItems.findIndex((t, i) => !usedTemplateIndices.has(i) && normalize(t.name) === normBess)

    if (tmplIdx >= 0) {
      usedTemplateIndices.add(tmplIdx)
      newDataRows.push(templateItems[tmplIdx].xml)
      console.log(`[report] Innovation: matched template row "${bess.name}"`)
    } else if (templateRow) {
      newDataRows.push(buildInnovationRow(templateRow, bess.name, bess.description))
      console.log(`[report] Innovation: built new row "${bess.name}"`)
    } else {
      console.warn(`[report] Innovation: no templateRow, skipping "${bess.name}"`)
    }
  }

  // Reconstruct table: tblPr/tblGrid prefix + header rows + new data rows + </w:tbl>
  const rowsStart = tableXml.search(/<w:tr\b/)
  let lastTrEnd = 0
  const trScan = /<w:tr\b[\s\S]*?<\/w:tr>/gi
  let m: RegExpExecArray | null
  while ((m = trScan.exec(tableXml)) !== null) lastTrEnd = m.index + m[0].length

  const newTableXml =
    tableXml.slice(0, rowsStart) +
    headerRows.join('') +
    newDataRows.join('') +
    tableXml.slice(lastTrEnd)

  return xml.slice(0, tblStart) + newTableXml + xml.slice(tblEnd)
}

/**
 * Delete table rows whose XML contains the given identifying text.
 * Useful for removing merged-continuation rows with no text in the first cell.
 */
function deleteTableRowsContaining(xml: string, identifyingText: string): string {
  let result = ''
  let i = 0
  while (i < xml.length) {
    const trStart = nextTrOpen(xml, i)
    if (trStart === -1) { result += xml.slice(i); break }
    result += xml.slice(i, trStart)
    const trEnd = findRowEnd(xml, trStart)
    const rowXml = xml.slice(trStart, trEnd)
    result += rowXml.includes(identifyingText) ? '' : rowXml
    i = trEnd
  }
  return result
}

// ─── Excel cell patching (PizZip-based, preserves charts and drawings) ───────

function findSheetXmlPath(zip: PizZip, sheetName: string): string {
  const wbXml = zip.file('xl/workbook.xml')!.asText()
  const relsXml = zip.file('xl/_rels/workbook.xml.rels')!.asText()
  // XML-encode the name (& → &amp; etc.) then regex-escape for the pattern
  const xmlEncoded = sheetName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const escaped = xmlEncoded.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const sheetM = wbXml.match(new RegExp(`<sheet [^>]*name="${escaped}"[^>]*r:id="([^"]+)"`))
  if (!sheetM) throw new Error(`Sheet "${sheetName}" not found`)
  const relM = relsXml.match(new RegExp(`Id="${sheetM[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]+Target="([^"]+)"`))
  if (!relM) throw new Error(`Relationship for sheet "${sheetName}" not found`)
  const target = relM[1]
  return target.startsWith('worksheets/') ? `xl/${target}` : target
}

function patchCell(wsXml: string, cellRef: string, value: string | number): string {
  const r = cellRef.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  if (typeof value === 'number') {
    // Replace existing cell value, stripping type attr so Excel treats it as numeric
    return wsXml.replace(
      new RegExp(`<c r="${r}"([^>]*)>(?:<f>[^<]*</f>)?<v>[^<]*</v></c>`),
      (_, attrs) => `<c r="${r}"${attrs.replace(/ t="[^"]*"/, '')}><v>${value}</v></c>`,
    )
  }
  const esc = String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  return wsXml.replace(
    new RegExp(`<c r="${r}"[^>]*>(?:<f>[^<]*</f>)?(?:<v>[^<]*</v>|<is>[\s\S]*?<\/is>)</c>`),
    `<c r="${r}" t="inlineStr"><is><t xml:space="preserve">${esc}</t></is></c>`,
  )
}

/**
 * Apply character-level substitutions to a paragraph's XML.
 * `changes` is a list of { pos (in combined text), len, replacement } records,
 * applied last-to-first so earlier positions stay valid.
 */
function applyParaChanges(
  para: string,
  changes: Array<{ pos: number; len: number; replacement: string }>,
): string {
  if (!changes.length) return para

  const runRe = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g
  const runs: { open: string; text: string; close: string; xmlStart: number; xmlLen: number }[] = []
  let rm: RegExpExecArray | null
  while ((rm = runRe.exec(para)) !== null) {
    runs.push({ open: rm[1], text: rm[2], close: rm[3], xmlStart: rm.index, xmlLen: rm[0].length })
  }

  // Build per-character position map: combinedPos → { runIdx, offsetInRun }
  const posMap: { ri: number; ci: number }[] = []
  for (let ri = 0; ri < runs.length; ri++) {
    for (let ci = 0; ci < runs[ri].text.length; ci++) posMap.push({ ri, ci })
  }

  // Each change applies to a range [pos, pos+len). We store one replacement text
  // at the first character and empty strings for the rest.
  const charOps: { ri: number; ci: number; text: string }[] = []
  for (const { pos, len, replacement } of changes) {
    for (let i = 0; i < len; i++) {
      if (pos + i >= posMap.length) break
      const { ri, ci } = posMap[pos + i]
      charOps.push({ ri, ci, text: i === 0 ? replacement : '' })
    }
  }

  // Group by run
  const byRun = new Map<number, { ci: number; text: string }[]>()
  for (const op of charOps) {
    if (!byRun.has(op.ri)) byRun.set(op.ri, [])
    byRun.get(op.ri)!.push({ ci: op.ci, text: op.text })
  }

  // Apply within each run from last offset to first
  const modded = new Map<number, string>()
  for (const [ri, ops] of byRun.entries()) {
    ops.sort((a, b) => b.ci - a.ci)
    let text = runs[ri].text
    for (const { ci, text: rep } of ops) {
      text = text.slice(0, ci) + rep + text.slice(ci + 1)
    }
    modded.set(ri, text)
  }

  // Rebuild paragraph XML from last run to first (preserve xmlStart offsets)
  let result = para
  for (const [ri, newText] of [...modded.entries()].sort((a, b) => b[0] - a[0])) {
    const run = runs[ri]
    result =
      result.slice(0, run.xmlStart) +
      `${run.open}${newText}${run.close}` +
      result.slice(run.xmlStart + run.xmlLen)
  }
  return result
}


/**
 * Escape long-form [XX sentence...] placeholders and orphaned ] characters that
 * Docxtemplater can't handle. Uses full-width brackets ［ (U+FF3B) and ］ (U+FF3D)
 * as reversible escape characters that are restored after rendering.
 *
 * Two problem patterns:
 *  1. [XX ...long text...] where ] is >20 chars away — Docxtemplater's expression
 *     parser chokes on spaces/periods in the tag name.
 *  2. Orphaned ] with no matching [ in the same paragraph — the matching [ was in a
 *     different paragraph and Docxtemplater reports "unopened tag".
 */
function escapeLongXXTags(xml: string): string {
  const LESC = '［' // full-width [ U+FF3B
  const RESC = '］' // full-width ] U+FF3D
  const MAX_TAG_LEN = 20

  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
    type RunInfo = { open: string; text: string; close: string; xmlStart: number; xmlLen: number }
    const runs: RunInfo[] = []
    const runRe = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g
    let rm: RegExpExecArray | null
    while ((rm = runRe.exec(para)) !== null) {
      runs.push({ open: rm[1], text: rm[2], close: rm[3], xmlStart: rm.index, xmlLen: rm[0].length })
    }
    if (!runs.some(r => r.text.includes('[') || r.text.includes(']'))) return para

    // Build combined text and per-char position map
    const posMap: Array<{ runIdx: number; off: number }> = []
    for (let ri = 0; ri < runs.length; ri++) {
      for (let ci = 0; ci < runs[ri].text.length; ci++) posMap.push({ runIdx: ri, off: ci })
    }
    const combined = runs.map(r => r.text).join('')

    const escapeLPos = new Set<number>() // combined positions of [ to escape
    const escapeRPos = new Set<number>() // combined positions of ] to escape

    // Step 1: mark long-form [XX tags and their corresponding ]
    for (let i = 0; i < combined.length - 1; i++) {
      if (combined[i] !== '[') continue
      if (combined[i + 1] !== 'X' && combined[i + 1] !== 'x') continue
      const closeIdx = combined.indexOf(']', i + 1)
      if (closeIdx !== -1 && closeIdx - i <= MAX_TAG_LEN) continue // short tag, keep
      escapeLPos.add(i)
      if (closeIdx !== -1) {
        escapeRPos.add(closeIdx)
        i = closeIdx
      }
    }

    // Step 2: find orphaned ] (no unescaped [ precedes it in this paragraph)
    let depth = 0
    for (let i = 0; i < combined.length; i++) {
      if (combined[i] === '[' && !escapeLPos.has(i)) depth++
      else if (combined[i] === ']' && !escapeRPos.has(i)) {
        if (depth > 0) depth--
        else escapeRPos.add(i) // orphaned
      }
    }

    if (!escapeLPos.size && !escapeRPos.size) return para

    // Build modified run texts
    const modded = new Map<number, string>()
    const apply = (pos: number, esc: string) => {
      const { runIdx, off } = posMap[pos]
      const cur = modded.get(runIdx) ?? runs[runIdx].text
      modded.set(runIdx, cur.slice(0, off) + esc + cur.slice(off + 1))
    }
    for (const pos of escapeLPos) apply(pos, LESC)
    for (const pos of escapeRPos) apply(pos, RESC)

    // Rebuild para XML from last run to first (preserve xmlStart positions)
    let result = para
    for (const [ri, newText] of [...modded.entries()].sort((a, b) => b[0] - a[0])) {
      const run = runs[ri]
      result = result.slice(0, run.xmlStart) + `${run.open}${newText}${run.close}` + result.slice(run.xmlStart + run.xmlLen)
    }
    return result
  })
}

// ─── Word template filling ────────────────────────────────────────────────────

async function fillWordTemplate(
  templatePath: string,
  data: {
    giwref: string
    projectAddress: string
    client: string
    architect: string
    date: string
  },
  project: ReportProjectData,
  credits: ReportCreditData[],
): Promise<Buffer> {
  const content = fs.readFileSync(templatePath, 'binary')
  const zip = new PizZip(content)

  // Extract document paragraphs and criteria row texts for context
  const docXml = zip.file('word/document.xml')!.asText()
  const paragraphs = extractDocParagraphs(docXml)
  const criteriaNames = extractTableRowTexts(docXml)

  // Escape long-form [XX...] placeholders before Docxtemplater sees them
  // (Docxtemplater's expression parser can't handle spaces/periods in tag names)
  const escapedDocXml = escapeLongXXTags(docXml)
  if (escapedDocXml !== docXml) {
    zip.file('word/document.xml', escapedDocXml)
  }

  // Ask Claude to fill all [XX], [XXX], etc. and decide which rows to delete
  const { placeholders, rowsToDelete, council } = await getWordFillData(project, credits, paragraphs, criteriaNames)

  // Per-occurrence counters for generic tags
  const counters: Record<string, number> = {}

  const unwrapDocxtemplaterError = (err: unknown): never => {
    const e = err as any
    const details = e?.properties?.errors?.length
      ? (e.properties.errors as any[]).map((x: any) => x?.properties?.explanation ?? x?.message ?? String(x)).join(' | ')
      : e?.message ?? String(e)
    console.error('[report] Docxtemplater error:', details)
    throw new Error('Template render failed: ' + details)
  }

  let doc: Docxtemplater
  try {
    doc = new Docxtemplater(zip, {
      delimiters: { start: '[', end: ']' },
      paragraphLoop: true,
      linebreaks: true,
      nullGetter(part: { value: string }) {
        const tag = part.value

        const namedFallbacks: Record<string, string> = {
          'GIWREF': data.giwref,
          'Project Address': data.projectAddress,
          'Client': data.client,
          'Architect': data.architect,
          'Date': data.date,
          'SMP Visualisation. Copy from C2 to U25 from excel sheet. Paste as image':
            '[SMP Visualisation chart — open the attached Excel, copy range C2:U25 and paste as image here]',
        }
        if (namedFallbacks[tag] !== undefined) return namedFallbacks[tag]

        counters[tag] = (counters[tag] ?? 0) + 1
        const idx = counters[tag] - 1
        const vals = placeholders[tag]
        if (vals && idx < vals.length && vals[idx] !== undefined) {
          const v = String(vals[idx])
          if (v !== '') return v
        }

        return '[' + tag + ']'
      },
    })
  } catch (err) { unwrapDocxtemplaterError(err) }

  try {
    doc!.render({
      GIWREF: data.giwref,
      'Project Address': data.projectAddress,
      Client: data.client,
      Architect: data.architect,
      Date: data.date,
      'SMP Visualisation. Copy from C2 to U25 from excel sheet. Paste as image':
        '[SMP Visualisation chart — open the attached Excel, copy range C2:U25 and paste as image here]',
    })
  } catch (err) { unwrapDocxtemplaterError(err) }

  // Post-process rendered XML
  const renderedZip = doc!.getZip()
  const rawXml = renderedZip.file('word/document.xml')!.asText()
  // Restore escaped long-form brackets (full-width → ASCII)
  let renderedXml = rawXml.replace(/［/g, '[').replace(/］/g, ']')

  // Shield the Key ESD Initiatives table from all post-processing
  const { xml: shieldedForKeyESD, restore: restoreKeyESD } = shieldTable(renderedXml, 'Key ESD Initiatives')
  renderedXml = shieldedForKeyESD

  const safeApply = (label: string, op: () => string): void => {
    try {
      const result = op()
      const before = (renderedXml.match(/<\/w:p>/g) || []).length
      const after = (result.match(/<\/w:p>/g) || []).length
      if (after < before * 0.5) {
        console.error(`[report] "${label}" dropped paragraph count ${before}→${after}, reverting`)
      } else {
        renderedXml = result
      }
    } catch (e) {
      console.error(`[report] "${label}" threw:`, e)
    }
  }

  try {

  // ── Innovation table: sync BEFORE row deletion so template rows are available ──
  // Must run before deleteTableRows — Claude's rowsToDelete may include template
  // Innovation criteria names, which would remove the templateRow needed to build new rows.
  try {
    const innovCredit = credits.find(c => /^innovation\s*1\.1/i.test(c.creditId))
    if (innovCredit?.rawDataPoints) {
      const bessPoints = parseInnovationPoints(innovCredit.rawDataPoints)
      if (bessPoints.length > 0) {
        renderedXml = syncInnovationTable(renderedXml, bessPoints)
      }
    }
  } catch (e) {
    console.error('[report] Innovation table sync failed:', e)
  }

  if (rowsToDelete.length > 0) {
    // Materials & Waste rows must always remain regardless of credit status
    const NEVER_DELETE = /embodied\s*energy|structural.*steel|sustainable\s*timber|\bpvc\b|sustainable\s*products|building\s*re-?use/i
    const safeToDelete = rowsToDelete.filter(r => !NEVER_DELETE.test(r))
    if (safeToDelete.length > 0) {
      safeApply('deleteTableRows', () => {
        const result = deleteTableRows(renderedXml, safeToDelete)
        console.log(`[report] Deleted ${safeToDelete.length} table rows:`, safeToDelete)
        return result
      })
    }
  }

  // ── Council dropdowns ──────────────────────────────────────────────────────
  if (council) {
    console.log(`[report] Setting council dropdowns to: ${council}`)
    // Try council name variants: full name first, then without trailing " Shire"
    // (some dropdown options omit "Shire", e.g. "City of Yarra Ranges" vs "City of Yarra Ranges Shire")
    const councilVariants = [council]
    const noShire = council.replace(/\s+Shire$/, '')
    if (noShire !== council) councilVariants.push(noShire)

    let esdMatch: string | null = null
    let swMatch: string | null = null
    let shortMatch: string | null = null
    let sdappMatch: string | null = null

    for (const c of councilVariants) {
      const e = c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      if (!esdMatch) {
        esdMatch = (renderedXml.match(new RegExp(`displayText="(${e}[^"]*15\\.01[^"]*)"`, '')) ||
                   renderedXml.match(new RegExp(`displayText="(${e}[^"]*15\\.02[^"]*)"`, '')) ||
                   renderedXml.match(new RegExp(`displayText="(${e}[^"]*15\\.05[^"]*)"`, '')) ||
                   renderedXml.match(new RegExp(`displayText="(${e}[^"]*22\\.06[^"]*)"`, '')) ||
                   renderedXml.match(new RegExp(`displayText="(${e}[^"]*sustainability[^"]*)"`, 'i')))?.[1] ?? null
      }
      if (!swMatch) {
        swMatch = renderedXml.match(new RegExp(`displayText="(${e}[^"]*53\\.18[^"]*)"`, ''))?.[1] ?? null
        // Melbourne uses a different stormwater clause
        if (!swMatch) swMatch = renderedXml.match(new RegExp(`displayText="(${e}[^"]*19\\.03[^"]*)"`, ''))?.[1] ?? null
      }
      if (!shortMatch) {
        shortMatch = renderedXml.match(new RegExp(`displayText="(${e})"`, ''))?.[1] ?? null
      }
      if (!sdappMatch) {
        sdappMatch = renderedXml.match(new RegExp(`displayText="(${e}[^"]*(?:sdapp|sustainable design assessment)[^"]*)"`, 'i'))?.[1] ?? null
      }
    }

    if (esdMatch) {
      renderedXml = setDropdownContent(renderedXml, esdMatch, esdMatch)
      renderedXml = flattenDropdown(renderedXml, esdMatch)
    }
    if (swMatch) {
      renderedXml = setDropdownContent(renderedXml, swMatch, swMatch)
      renderedXml = flattenDropdown(renderedXml, swMatch)
    }
    if (shortMatch) {
      renderedXml = setDropdownContent(renderedXml, shortMatch, shortMatch)
      renderedXml = flattenDropdown(renderedXml, shortMatch)
    }
    if (sdappMatch) {
      renderedXml = setDropdownContent(renderedXml, sdappMatch, sdappMatch)
      renderedXml = flattenDropdown(renderedXml, sdappMatch)
    }
  }

  // ── BESS score plain-text replacement ─────────────────────────────────────
  if (project.bessScore != null) {
    const score = Math.round(project.bessScore)
    renderedXml = renderedXml.replace('BESS score of XX% with', () => `BESS score of ${score}% with`)
  }

  // ── Pre-application meeting dropdown ──────────────────────────────────────
  {
    const mgmt11 = credits.find(c => c.creditId.toLowerCase() === 'management 1.1')
    const preAppItem = mgmt11?.creditStatus === 'Y'
      ? 'GIW has been involved in a pre-application meeting with Council on XXX. '
      : 'GIW has been actively involved in the preliminary design stage, but has not been involved in a pre-application meeting with Council. '
    renderedXml = setDropdownContent(renderedXml, 'GIW has been involved in a pre-application meeting with Council on XXX. ', preAppItem)
  }

  // ── BADS cooling load dropdown ────────────────────────────────────────────
  {
    const oe11Raw = (findCredit(credits, 'oe 1.1')?.rawDataPoints ?? '').toLowerCase()
    const oe12Raw = (findCredit(credits, 'oe 1.2')?.rawDataPoints ?? '').toLowerCase()
    const combinedRaw = oe11Raw + ' ' + oe12Raw
    let badsItem: string
    if (/21\s*mj/i.test(combinedRaw) || /bads[^.]*21/i.test(combinedRaw)) {
      badsItem = '21MJ/M2'
    } else if (/22\s*mj/i.test(combinedRaw) || /bads[^.]*22/i.test(combinedRaw)) {
      badsItem = '22MJ/m2'
    } else {
      badsItem = '30 MJ/m2'
    }
    renderedXml = setDropdownContent(renderedXml, '30 MJ/m2', badsItem)
  }

  // ── Hot water system dropdown ─────────────────────────────────────────────
  {
    const raw = getHotWaterRaw(credits).toLowerCase()
    let hwItem: string
    if (raw.includes('heat pump')) {
      hwItem = 'a heat pump hot water system'
    } else if (raw.includes('gas') && (raw.includes('central') || raw.includes('communal'))) {
      hwItem = 'a centralised gas hot water system'
    } else if (raw.includes('gas')) {
      hwItem = 'inidividual gas hot water systems'
    } else if (raw.includes('electric') || raw.includes('instantaneous')) {
      hwItem = 'individual electric instantaneous hot water systems'
    } else {
      hwItem = 'a heat pump hot water system' // sensible default
    }
    renderedXml = setDropdownContent(renderedXml, 'a centralised gas hot water system', hwItem)
  }

  // ── Clothes drying dropdown (from OE dwelling profile "Clothes line") ─────
  {
    // Scan all OE 2.x rawDataPoints for clothes drying / clothes line mentions
    const oeRaw = credits
      .filter(c => /^oe\s*2/i.test(c.creditId))
      .map(c => c.rawDataPoints ?? '')
      .join(' ')
    const dryingRaw = (oeRaw + ' ' + credits
      .filter(c => /clothes|drying|laundry/i.test(c.creditId + ' ' + c.category + ' ' + (c.rawDataPoints ?? '')))
      .map(c => c.rawDataPoints ?? '')
      .join(' ')).toLowerCase()

    const noFacilities = /no (clothes|drying|laundry)|no (drying|clothes) facilit|clothes (line|drying).*none|none.*clothes/i.test(dryingRaw)

    if (noFacilities) {
      // Delete the Clothes Drying criteria row entirely
      renderedXml = deleteTableRows(renderedXml, ['Clothes Drying'])
    } else if (/communal|shared|rooftop|terrace/i.test(dryingRaw)) {
      renderedXml = setDropdownContent(renderedXml, 'clothes drying racks on the balcony',
        'Communal clothes drying facilities will be provided at rooftop terrace.')
    } else if (/indoor|internal|inside/i.test(dryingRaw)) {
      renderedXml = setDropdownContent(renderedXml, 'clothes drying racks on the balcony',
        'All apartments will be provided with indoor clothes drying rack / lines.')
    } else {
      // private / balcony / individual
      renderedXml = setDropdownContent(renderedXml, 'clothes drying racks on the balcony',
        'All apartments will be provided with clothes drying racks on the balcony.')
    }
  }

  // ── Landscape irrigation dropdown ─────────────────────────────────────────
  {
    const iwm31 = credits.find(c => c.creditId.toLowerCase() === 'iwm 3.1')
    let irrigationItem: string
    if (iwm31?.creditStatus === 'ScopedOut') {
      irrigationItem = 'Landscape irrigation demand will be connected to the rainwater tank. '
    } else {
      // Y (achieved) or absent/N → native vegetation (no irrigation needed)
      irrigationItem = 'The majority of landscaping is to be native vegetation with no irrigation demand after the initial establishment period.'
    }
    renderedXml = setDropdownContent(renderedXml, 'native vegetation with no irrigation demand', irrigationItem)
  }

  // ── Building Reuse dropdown ────────────────────────────────────────────────
  // Dropdown options: "None of the existing structure is re-used." (default) /
  //   "At least 30% of the existing structure is re-used." / "There is no existing building..."
  // Two dropdown instances exist (one in criteria table, one in a merged continuation row).
  {
    const BUILDING_REUSE_OPT = 'None of the existing structure is re-used.'
    const buildingReuse = credits.find(c => /building\s*re-?use/i.test(c.creditId))
    const brStatus = buildingReuse?.creditStatus ?? ''
    const isTargeted = /^(y|achieved|targeted|yes)$/i.test(brStatus)
    const isScopedOut = /scoped.?out/i.test(brStatus)

    if (isTargeted || isScopedOut) {
      const selectedValue = isTargeted
        ? 'At least 30% of the existing structure is re-used.'
        : 'There is no existing building on the proposed site.'
      // Set and flatten both dropdown instances (loop until none remain)
      let attempts = 0
      while (renderedXml.includes(`displayText="${BUILDING_REUSE_OPT}"`) && attempts < 4) {
        renderedXml = setDropdownContent(renderedXml, BUILDING_REUSE_OPT, selectedValue)
        renderedXml = flattenDropdown(renderedXml, BUILDING_REUSE_OPT)
        attempts++
      }
    } else {
      // Not targeted — delete all Building Reuse table rows
      renderedXml = deleteTableRows(renderedXml, ['Building Re-use', 'Building Reuse'])
      // Also remove merged-continuation rows that still contain the dropdown identifier
      renderedXml = deleteTableRowsContaining(renderedXml, BUILDING_REUSE_OPT)
    }
  }

  // ── Daylight DTS block deletion ────────────────────────────────────────────
  // Remove the Deemed-to-Satisfy block when daylight was assessed via simulation/modelling
  // OR when the BESS built-in calculator was used (both are alternatives to DTS).
  {
    const daylightRaw = credits
      .filter(c => /ieq 1\.[34]/i.test(c.creditId))
      .map(c => (c.rawDataPoints ?? '').toLowerCase())
      .join(' ')
    const usedModelling = /simulat|modell|daylight factor|\bdf\b|\d+\s*%\s*of\s*(living|bedroom)/i.test(daylightRaw)
    const usedBuiltInDaylight = /built[- ]?in\s*calculator|bess\s*calculator|bess\s*built[- ]?in/i.test(daylightRaw)
    if (usedModelling || usedBuiltInDaylight) {
      renderedXml = deleteParagraphsByText(renderedXml, [
        /^Or$/i,
        /Deemed-to-Satisfy method for IEQ/i,
        /All North.*West.*East.*8m/i,
        /floor-to-ceiling height of 2\.7/i,
        /60%.*visible light transmittance/i,
        /All living areas have an external facing window/i,
        /building separation tables/i,
      ])
    }
  }

  // ── Daylight method: BESS built-in calculator vs modelling ───────────────
  // If IEQ 1.3/1.4 OR IEQ 1.5 rawDataPoints indicate the BESS built-in calculator was used,
  // replace the modelling sentence with the built-in calculator sentence.
  {
    const daylightRaw = credits
      .filter(c => /ieq 1\.[34]/i.test(c.creditId))
      .map(c => (c.rawDataPoints ?? '').toLowerCase())
      .join(' ')
    const ieq15Raw = (findCredit(credits, 'ieq 1.5')?.rawDataPoints ?? '').toLowerCase()
    const usedBuiltIn = /built[- ]?in\s*calculator|bess\s*calculator|bess\s*built[- ]?in/i.test(daylightRaw + ' ' + ieq15Raw)
    if (usedBuiltIn) {
      const escaped = 'The BESS built-in calculator has been used to demonstrate compliance.'
      renderedXml = renderedXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para: string) => {
        const texts: string[] = []
        const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
        let m: RegExpExecArray | null
        while ((m = re.exec(para)) !== null) texts.push(m[1])
        const combined = texts.join('')
        if (!/daylight modelling has been conducted for a representative/i.test(combined)) return para
        let replaced = false
        return para.replace(/(<w:t[^>]*>)([^<]*)(<\/w:t>)/g, (_: string, open: string, content: string, close: string) => {
          if (!replaced && content.trim()) { replaced = true; return `${open}${escaped}${close}` }
          return `${open}${close}`
        })
      })
    }
  }

  // ── Bicycle spaces: delete ratio/increase sentence when credit not achieved ──
  // Check each bicycle credit and remove the follow-on sentence if not achieved.
  // Update the credit IDs below to match the BESS transport credit IDs for this project type.
  {
    const bikeCredits: Array<{
      creditId: string
      sentencePattern?: RegExp
      contextPattern?: RegExp
    }> = [
      { creditId: 'transport 1.1', sentencePattern: /ratio of approximately 1 resident bicycle space for every apartment/i },
      { creditId: 'transport 1.2', sentencePattern: /ratio of approximately 1 visitor bicycle space for every/i },
      { creditId: 'transport 1.3', sentencePattern: /bicycle spaces.*for employees/i },
      { creditId: 'transport 1.4', sentencePattern: /bicycle spaces.*for non-residential visitors/i },
    ]

    for (const { creditId, sentencePattern, contextPattern } of bikeCredits) {
      const credit = findCreditLike(credits, creditId)
      const notAchieved = !credit || credit.creditStatus === 'ScopedOut' ||
        credit.creditStatus === 'Not Achieved' || credit.creditStatus === 'N/A'
      if (notAchieved) {
        if (sentencePattern) renderedXml = deleteParagraphsByText(renderedXml, [sentencePattern])
        else if (contextPattern) renderedXml = deleteParagraphAfterMatch(renderedXml, contextPattern)
      }
    }
  }

  // ── IEQ 2.3: Non-residential ventilation — delete zero/unfilled options ─────
  {
    const ieq23 = findCredit(credits, 'ieq 2.3')
    const raw = ieq23?.rawDataPoints ?? ''

    const naturalPct = raw ? extractNumber(raw, [
      /natural(?:\s+ventilation)?[:\s]+(\d+(?:\.\d+)?)\s*%/i,
      /(\d+(?:\.\d+)?)\s*%[^.\n]*natural/i,
    ]) : null
    const outdoorAirPct = raw ? extractNumber(raw, [
      /outdoor\s+air[:\s]+(\d+(?:\.\d+)?)\s*%/i,
      /(\d+(?:\.\d+)?)\s*%[^.\n]*outdoor\s+air/i,
      /(\d+(?:\.\d+)?)\s*%[^.\n]*(?:increase|above)[^.\n]*AS\s*1668/i,
    ]) : null
    const co2Ppm = raw ? extractNumber(raw, [
      /(\d{3,4})\s*ppm/i,
      /CO2[^:]*:?\s*(\d{3,4})/i,
    ]) : null

    // Delete when value is confirmed 0 from rawDataPoints
    if (naturalPct === 0)
      renderedXml = deleteParagraphsByText(renderedXml, [/non-residential spaces is naturally ventilated/i])
    if (outdoorAirPct === 0)
      renderedXml = deleteParagraphsByText(renderedXml, [/increase in outdoor air rates/i])
    if (co2Ppm === 0 || co2Ppm === null)
      renderedXml = deleteParagraphsByText(renderedXml, [/CO2 sensors are to be installed/i])

    // Delete paragraphs that still have unfilled [XX%] placeholders (Claude couldn't determine value)
    // — runs regardless of whether rawDataPoints exist, catching any leftover placeholders
    renderedXml = deleteParagraphsByText(renderedXml, [
      /\[XX%\].*non-residential spaces is naturally ventilated/i,
      /\[XX%\].*increase in outdoor air rates/i,
    ])
  }

  // ── Ceiling fans: delete line when 0% or unfilled ─────────────────────────
  renderedXml = deleteParagraphsByText(renderedXml, [
    /(?:\b0%\b|\[XX\]%|\[XX%\]).*ceiling fan/i,
    /ceiling fan.*(?:\b0%\b|\[XX\]%|\[XX%\])/i,
  ])

  // ── Retail / commercial line items ────────────────────────────────────────
  const { hasRetail, hasCommercial } = detectNonResidential(credits, project.typology ?? null)
  safeApply('removeNonResidentialLines', () => removeNonResidentialLines(renderedXml, hasRetail, hasCommercial))

  // ── BESS rawDataPoints fallbacks (catches any unfilled placeholders) ──────
  safeApply('applyBESSFallbacks', () => applyBESSFallbacks(renderedXml, credits, project))

  // ── Category page breaks ──────────────────────────────────────────────────
  safeApply('addCategoryPageBreaks', () => addCategoryPageBreaks(renderedXml))

  // ── Appendix numbering ────────────────────────────────────────────────────
  safeApply('numberAppendices', () => numberAppendices(renderedXml))

  } catch (err) {
    console.error('[report] Post-processing error — falling back to raw render:', err)
    renderedXml = rawXml
  }

  // Restore the shielded Key ESD Initiatives table (no-op if error path took rawXml)
  renderedXml = restoreKeyESD(renderedXml)

  renderedZip.file('word/document.xml', renderedXml)
  const wordBuf = renderedZip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer
  // DEBUG: save to Desktop so we can inspect the exact generated file
  try {
    const debugPath = `C:\\Users\\inesb\\Desktop\\debug-report-${Date.now()}.docx`
    require('fs').writeFileSync(debugPath, wordBuf)
    console.log('[report] DEBUG saved to', debugPath)
  } catch { /* ignore */ }
  return wordBuf
}

// ─── Excel template filling ───────────────────────────────────────────────────

function fillExcelResidential(
  templatePath: string,
  project: ReportProjectData,
  credits: ReportCreditData[],
): Buffer {
  const zip = new PizZip(fs.readFileSync(templatePath, 'binary'))
  const sheetPath = findSheetXmlPath(zip, 'Inputs Residential & Mixed-Use')
  let ws = zip.file(sheetPath)!.asText()

  ws = patchCell(ws, 'D5', project.address ?? '')

  const oe12 = findCredit(credits, 'oe 1.2')
  if (oe12?.rawDataPoints) {
    const stars = extractNumber(oe12.rawDataPoints, [
      /(\d+\.?\d*)\s*star/i,
      /star\s*rating[^:]*:\s*(\d+\.?\d*)/i,
      /average[^:]*:\s*(\d+\.?\d*)/i,
    ])
    if (stars != null) ws = patchCell(ws, 'D23', stars)
  }

  const oe11 = findCredit(credits, 'oe 1.1')
  if (oe11?.rawDataPoints) {
    const heating = extractNumber(oe11.rawDataPoints, [
      /heating[^:]*:\s*([\d.,]+)\s*mj/i,
      /heating load[^:]*:\s*([\d.,]+)/i,
    ])
    const cooling = extractNumber(oe11.rawDataPoints, [
      /cooling[^:]*:\s*([\d.,]+)\s*mj/i,
      /cooling load[^:]*:\s*([\d.,]+)/i,
    ])
    if (heating != null) ws = patchCell(ws, 'D24', heating)
    if (cooling != null) ws = patchCell(ws, 'D25', cooling)
  }

  const oe27 = findCredit(credits, 'oe 2.7')
  if (oe27?.rawDataPoints) {
    const hvac = extractText(oe27.rawDataPoints, [
      /hvac[^:]*:\s*([^.;,\n]+)/i,
      /system type[^:]*:\s*([^.;,\n]+)/i,
      /(split system|vrv|vrf|heat pump|ducted|fan coil)[^\s,;.]*/i,
    ])
    if (hvac) { ws = patchCell(ws, 'D26', hvac); ws = patchCell(ws, 'D27', hvac) }
  }

  const lighting = findCredit(credits, 'oe 3.5') ?? findCredit(credits, 'oe 3.6')
  if (lighting?.rawDataPoints) {
    const pct = extractNumber(lighting.rawDataPoints, [/([\d.]+)%\s*reduction/i, /reduction[^:]*:\s*([\d.]+)%/i])
    if (pct != null) ws = patchCell(ws, 'D30', pct)
  }

  const oe26 = findCredit(credits, 'oe 2.6')
  if (oe26?.rawDataPoints) {
    const cooktop = extractText(oe26.rawDataPoints, [/cooktop[^:]*:\s*([^.;,\n]+)/i, /(induction|gas|electric)[^.;,\n]*/i])
    if (cooktop) ws = patchCell(ws, 'D32', cooktop)
  }

  const solar = findCredit(credits, 'oe 4.2') ?? findCredit(credits, 'oe 4.5')
  if (solar?.rawDataPoints) {
    const solarKwh = extractNumber(solar.rawDataPoints, [/([\d,]+)\s*kwh/i, /solar[^:]*:\s*([\d,]+)/i])
    if (solarKwh != null) ws = patchCell(ws, 'D40', solarKwh)
  }

  {
    const hwRaw = getHotWaterRaw(credits)
    const hwType = hwRaw ? extractText(hwRaw, [
      /hot water[^:]*:\s*([^.;,\n]+)/i,
      /type[^:]*:\s*([^.;,\n]+)/i,
      /(heat pump|electric|gas|solar|instantaneous)[^.;,\n]*/i,
    ]) : null
    if (hwType) ws = patchCell(ws, 'D51', hwType)
  }

  const iwm11 = findCredit(credits, 'iwm 1.1')
  if (iwm11?.rawDataPoints) {
    const refWater = extractNumber(iwm11.rawDataPoints, [/reference[^:]*:\s*([\d,]+)\s*kl/i, /([\d,]+)\s*kl\/yr[^.]*reference/i])
    const propWaterExcl = extractNumber(iwm11.rawDataPoints, [
      /proposed[^:]*excl[^:]*:\s*([\d,]+)\s*kl/i,
      /excluding rainwater[^:]*:\s*([\d,]+)\s*kl/i,
      /proposed[^:]*:\s*([\d,]+)\s*kl/i,
    ])
    const pctReduction = extractNumber(iwm11.rawDataPoints, [/([\d.]+)%\s*reduction/i, /reduction[^:]*:\s*([\d.]+)%/i])
    if (refWater != null) ws = patchCell(ws, 'Q51', refWater)
    if (propWaterExcl != null) ws = patchCell(ws, 'Q52', propWaterExcl)
    if (pctReduction != null) ws = patchCell(ws, 'D53', pctReduction)
  }

  const iwm21 = findCredit(credits, 'iwm 2.1')
  if (iwm21?.rawDataPoints) {
    const propWaterIncl = extractNumber(iwm21.rawDataPoints, [
      /including rainwater[^:]*:\s*([\d,]+)\s*kl/i,
      /proposed[^:]*incl[^:]*:\s*([\d,]+)\s*kl/i,
      /proposed[^:]*:\s*([\d,]+)\s*kl/i,
    ])
    if (propWaterIncl != null) ws = patchCell(ws, 'Q53', propWaterIncl)
  }

  zip.file(sheetPath, ws)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer
}

function fillExcelNonResidential(
  templatePath: string,
  project: ReportProjectData,
  credits: ReportCreditData[],
): Buffer {
  const zip = new PizZip(fs.readFileSync(templatePath, 'binary'))
  const sheetPath = findSheetXmlPath(zip, 'Inputs Non-Residential')
  let ws = zip.file(sheetPath)!.asText()

  ws = patchCell(ws, 'D5', project.address ?? '')

  const oe27 = findCredit(credits, 'oe 2.7')
  if (oe27?.rawDataPoints) {
    const hvac = extractText(oe27.rawDataPoints, [
      /hvac[^:]*:\s*([^.;,\n]+)/i,
      /(split system|vrv|vrf|heat pump|ducted|fan coil)[^\s,;.]*/i,
    ])
    if (hvac) {
      for (const cell of ['D24', 'I24', 'M24', 'Q24', 'U24']) ws = patchCell(ws, cell, hvac)
    }
  }

  const lighting = findCredit(credits, 'oe 3.5') ?? findCredit(credits, 'oe 3.6')
  if (lighting?.rawDataPoints) {
    const pct = extractNumber(lighting.rawDataPoints, [/([\d.]+)%\s*reduction/i, /reduction[^:]*:\s*([\d.]+)%/i])
    if (pct != null) {
      for (const cell of ['D26', 'I26', 'M26', 'Q26', 'U26']) ws = patchCell(ws, cell, pct)
    }
  }

  const solar = findCredit(credits, 'oe 4.2') ?? findCredit(credits, 'oe 4.5')
  if (solar?.rawDataPoints) {
    const solarKwh = extractNumber(solar.rawDataPoints, [/([\d,]+)\s*kwh/i, /solar[^:]*:\s*([\d,]+)/i])
    if (solarKwh != null) ws = patchCell(ws, 'D46', solarKwh)
  }

  {
    const hwRaw = getHotWaterRaw(credits)
    const hwType = hwRaw ? extractText(hwRaw, [
      /hot water[^:]*:\s*([^.;,\n]+)/i,
      /type[^:]*:\s*([^.;,\n]+)/i,
      /(heat pump|electric|gas|solar|instantaneous)[^.;,\n]*/i,
    ]) : null
    if (hwType) { ws = patchCell(ws, 'D53', hwType); ws = patchCell(ws, 'I53', hwType) }
  }

  const iwm11 = findCredit(credits, 'iwm 1.1')
  if (iwm11?.rawDataPoints) {
    const refWater = extractNumber(iwm11.rawDataPoints, [/reference[^:]*:\s*([\d,]+)\s*kl/i])
    const propWater = extractNumber(iwm11.rawDataPoints, [
      /proposed[^:]*excl[^:]*:\s*([\d,]+)\s*kl/i,
      /proposed[^:]*:\s*([\d,]+)\s*kl/i,
    ])
    if (refWater != null) ws = patchCell(ws, 'D62', refWater)
    if (propWater != null) ws = patchCell(ws, 'D63', propWater)
  }

  const iwm21 = findCredit(credits, 'iwm 2.1')
  if (iwm21?.rawDataPoints) {
    const propWaterIncl = extractNumber(iwm21.rawDataPoints, [
      /including rainwater[^:]*:\s*([\d,]+)\s*kl/i,
      /proposed[^:]*incl[^:]*:\s*([\d,]+)\s*kl/i,
      /proposed[^:]*:\s*([\d,]+)\s*kl/i,
    ])
    if (propWaterIncl != null) ws = patchCell(ws, 'D64', propWaterIncl)
  }

  zip.file(sheetPath, ws)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer
}

// ─── Typology inference from credit structure ─────────────────────────────────

/**
 * Infer typology from the credit data when the PDF extraction didn't capture it.
 *
 * Rules (matching what the BESS extraction prompt instructs Claude to detect):
 *  - OE 1.1 / OE 1.2 present → residential component exists
 *  - rawDataPoints mentioning "townhouse" or " TH " → townhouse inputs
 *  - Credits with category containing "retail", "office", "commercial" OR
 *    rawDataPoints mentioning those words → non-residential component exists
 */
function inferTypology(credits: ReportCreditData[]): string | null {
  const ids = credits.map(c => c.creditId.toLowerCase().replace(/\s+/g, ' '))
  const allRaw = credits.map(c => c.rawDataPoints ?? '').join(' ').toLowerCase()

  const hasResidentialCredits = ids.some(id => /^oe 1\.[12]/.test(id))
  const hasTownhouseInputs = /townhouse|\bth\b/.test(allRaw)
  const hasNonResidentialInputs =
    /retail|office|commercial|non.residential/.test(allRaw) ||
    credits.some(c => /retail|office|commercial/i.test(c.category))

  if (!hasResidentialCredits && !hasTownhouseInputs) return 'Non-Residential'
  if (hasResidentialCredits && hasNonResidentialInputs) return 'Mixed-Use'
  if (hasTownhouseInputs && !hasNonResidentialInputs) return 'Townhouse'
  if (hasResidentialCredits && !hasNonResidentialInputs) return 'Multi-Residential'
  return null
}

// ─── Public entry point ───────────────────────────────────────────────────────

export async function generateSMPReport(
  project: ReportProjectData,
  credits: ReportCreditData[],
  overrides?: { client?: string; architect?: string; giwref?: string; totalDwellings?: number },
): Promise<{
  wordBuffer: Buffer
  excelBuffer: Buffer
  wordFilename: string
  excelFilename: string
}> {
  const typology = project.typology || inferTypology(credits) || ''

  const wordTemplateName = WORD_TEMPLATES[typology]
  if (wordTemplateName === undefined) {
    throw new Error(
      `Unknown typology "${typology}". Expected one of: Mixed-Use, Multi-Residential, Townhouse, Non-Residential.`,
    )
  }
  const excelTemplateName = EXCEL_TEMPLATES[typology]
  const wordTemplatePath = path.join(TEMPLATES_DIR, wordTemplateName)
  const excelTemplatePath = path.join(TEMPLATES_DIR, excelTemplateName)

  if (!fs.existsSync(wordTemplatePath)) throw new Error(`Word template file missing: ${wordTemplateName}`)
  if (!fs.existsSync(excelTemplatePath)) throw new Error(`Excel template file missing: ${excelTemplateName}`)

  // Build output filenames following the existing naming convention
  const dateObj = project.date ? new Date(project.date) : new Date()
  const dd = String(dateObj.getDate()).padStart(2, '0')
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0')
  const yyyy = dateObj.getFullYear()
  const yy = String(yyyy).slice(-2)
  const dateStr = `${yy}${mm}${dd}`
  const giwref = overrides?.giwref || project.projectId || 'GIWREF'
  const safeAddr = (project.address ?? 'Address')
    .replace(/[^a-zA-Z0-9 ]/g, '')
    .replace(/\s+/g, '-')
    .substring(0, 40)
    .replace(/-+$/, '')
  const rev = project.revision ?? 'A'

  const typologySlug: Record<string, string> = {
    'Mixed-Use': 'MixUse', 'Multi-Residential': 'Multi-Apt',
    'Townhouse': 'TH', 'Non-Residential': 'Comm',
  }
  const wordFilename = `${dateStr}-${giwref}-${safeAddr}-SMP-${typologySlug[typology] ?? 'SMP'}-2022-${rev}.docx`
  const excelFilename = `${dateStr}-${giwref}-${safeAddr}-SMP-Visualisations-Rev${rev}.xlsx`

  const displayDate = `${dd}/${mm}/${yyyy}`
  const formattedAddress = formatAddress(project.address ?? '')

  // Override wins; then fall back to what was stored on the project record
  const resolvedTotalDwellings = overrides?.totalDwellings ?? project.totalDwellings ?? null

  const projectWithFormattedAddress = { ...project, address: formattedAddress, totalDwellings: resolvedTotalDwellings }

  const excelBuffer = typology === 'Non-Residential'
    ? fillExcelNonResidential(excelTemplatePath, projectWithFormattedAddress, credits)
    : fillExcelResidential(excelTemplatePath, projectWithFormattedAddress, credits)

  const wordBuffer = await fillWordTemplate(
    wordTemplatePath,
    {
      giwref,
      projectAddress: formattedAddress,
      client: overrides?.client ?? project.client ?? '',
      architect: overrides?.architect ?? project.architect ?? '',
      date: displayDate,
    },
    projectWithFormattedAddress,
    credits,
  )

  return { wordBuffer, excelBuffer, wordFilename, excelFilename }
}
