import * as fs from 'fs'
import * as path from 'path'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import Anthropic from '@anthropic-ai/sdk'
import ExcelJS from 'exceljs'

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

const TEMPLATES_DIR = path.resolve(process.cwd(), 'templates')

// Postcode → primary NatHERS climate zone, loaded once on first use
let _postcodeClimateZoneCache: Map<number, number> | null = null

async function getPostcodeClimateZoneMap(): Promise<Map<number, number>> {
  if (_postcodeClimateZoneCache) return _postcodeClimateZoneCache
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(path.join(TEMPLATES_DIR, 'NatHERSclimatezonesSept2025.xlsx'))
  const ws = wb.worksheets[0]
  const map = new Map<number, number>()
  ws.eachRow((row, rowNum) => {
    if (rowNum < 3) return // skip header rows
    const postcode = Number(row.getCell(1).value)
    const zone = Number(row.getCell(2).value)
    if (!isNaN(postcode) && !isNaN(zone) && zone > 0) map.set(postcode, zone)
  })
  _postcodeClimateZoneCache = map
  return map
}

function extractPostcode(address: string | null): number | null {
  if (!address) return null
  const match = address.match(/\b(\d{4})\b/)
  return match ? parseInt(match[1], 10) : null
}

// NatHERS climate zone → BADS cooling load dropdown value
const BADS_COOLING_LOAD: Record<number, string> = {
  21: '30 MJ/m2',
  60: '22MJ/m2',
  62: '21MJ/M2',
}

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
  buildingLevels?: number | null
  siteArea?: number | null
  rainwaterTankSize?: number | null
  gdft?: boolean | null
}

export interface ReportCreditData {
  creditId: string
  creditName?: string
  creditStatus: string
  rawDataPoints: string | null
  category: string
  commentsGIW?: string | null
  reviewerComments?: string[]
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

/** Add thousands comma to plain integers > 999 (e.g. 3356 → "3,356"). Leaves decimals and small numbers unchanged. */
function formatNum(val: string): string {
  const t = val.trim()
  if (/^\d{4,}$/.test(t)) return parseInt(t, 10).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return val
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

interface FillData {
  filledGIWComments: Record<string, string>
  rowsToDelete: string[]
  council: string | null
  namedValues: Record<string, string>
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
  // Replace the first <w:t> with the new value; clear any remaining <w:t> elements.
  // Some dropdowns store their placeholder as two runs: a space run and a grey "Choose an item" run.
  // Without clearing the second run, both texts appear in the output after flattenDropdown.
  let firstT = true
  content = content.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, (match) => {
    if (firstT) { firstT = false; return `<w:t>${escaped}</w:t>` }
    return match.replace(/>[^<]*<\/w:t>/, '></w:t>')
  })
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
 * Add yellow highlight to all runs inside a dropdown's sdtContent without
 * changing the selected value. Used to flag dropdowns that need manual review.
 */
function highlightDropdown(xml: string, identifyingOption: string): string {
  let pos = xml.indexOf(`displayText="${identifyingOption}"`)
  if (pos === -1) {
    const escaped = identifyingOption.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const m = xml.match(new RegExp(`displayText="[^"]*${escaped}[^"]*"`))
    if (!m) return xml
    pos = xml.indexOf(m[0])
    if (pos === -1) return xml
  }

  const contentStart = xml.indexOf('<w:sdtContent', pos)
  if (contentStart === -1) return xml
  const contentEnd = xml.indexOf('</w:sdtContent>', contentStart) + '</w:sdtContent>'.length

  let content = xml.slice(contentStart, contentEnd)
  // Inject yellow highlight into each run's <w:rPr>, creating one if absent
  content = content.replace(/<w:r(\s[^>]*)?>(?=[\s\S]*?<\/w:r>)/g, (match) => {
    if (match.includes('<w:rPr>')) return match
    return match.replace(/^(<w:r[^>]*>)/, '$1<w:rPr><w:highlight w:val="yellow"/></w:rPr>')
  })
  content = content.replace(/(<w:rPr>)(?![\s\S]*?<w:highlight )/g, '$1<w:highlight w:val="yellow"/>')

  return xml.slice(0, contentStart) + content + xml.slice(contentEnd)
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

  // Credits whose GIW comment contains [XX] placeholders that need filling
  const giwXXCredits = credits.filter(c => c.commentsGIW && c.commentsGIW.includes('[XX]'))
  const giwCommentSection = giwXXCredits.length > 0 ? `
GIW COMMENTS WITH [XX] PLACEHOLDERS (fill each [XX] with the correct value from the credit's rawDataPoints or project data):
${giwXXCredits.map(c => `${c.creditId} rawDataPoints: ${c.rawDataPoints ?? 'N/A'}\n${c.creditId} commentsGIW:\n${c.commentsGIW}`).join('\n\n')}
` : ''

  const prompt = `You are filling a Sustainability Management Plan (SMP) Word document template for a building project.

Your task is to provide named values for template tags, decide which criteria rows to delete, and fill [XX] placeholders in GIW comments.

PROJECT DATA:
- Project name: ${project.name}
- Address: ${project.address ?? 'Not provided'}
- BESS Score: ${project.bessScore != null ? project.bessScore + '%' : 'Not provided'}
- Client: ${project.client ?? 'Not provided'}
- Architect: ${project.architect ?? 'Not provided'}
- Revision: ${project.revision ?? 'A'}
- Date: ${project.date ? new Date(project.date).toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' }) : 'Not provided'}
- Site Area m²: ${project.siteArea != null ? project.siteArea : 'Not provided'}
- rainwaterTankSize (litres): ${project.rainwaterTankSize != null ? project.rainwaterTankSize : 'Not provided'}

BESS CREDIT DATA (creditId | status | rawDataPoints):
${creditSummary || 'No credit data available'}
${giwCommentSection}
ALL BESS CREDITS AND STATUSES (for row deletion):
${allCreditStatus || 'No credit data available'}

SMP REPORT TABLE CRITERIA ROWS (exact strings as they appear in the template):
${criteriaNames.join('\n')}

AVAILABLE MELBOURNE COUNCILS (pick the one matching the project address):
City of Boroondara, City of Yarra, City of Banyule, City of Bass Coast Shire, City of Darebin, City of Greater Dandenong, City of Hobsons Bay, City of Hume, City of Kingston, City of Knox, City of Manningham, City of Maribyrnong, City of Maroondah, City of Moonee Valley, City of Merri-bek, City of Port Phillip, City of Stonnington, City of Whitehorse, City of Whittlesea, City of Wyndham, City of Bayside, City of Brimbank, City of Glen Eira, City of Greater Bendigo, City of Greater Geelong, City of Melbourne, City of Monash, City of Yarra Ranges Shire
${giwCommentSection}
Return ONLY a valid JSON object with the following exact structure.

For "rowsToDelete": list the EXACT criteria names (from the list above) whose corresponding BESS credit is NOT claimed. A credit is not claimed if its status is "Scoped Out", "Not Achieved", "N/A", or if no matching credit exists in BESS at all. Do NOT include header rows (e.g. "Criteria", "Council Best Practice Standard", "Development Provision") or rows that don't correspond to a specific BESS credit. Use the exact criteria string from the list. IMPORTANT: Never include any Materials criteria (e.g. Embodied Energy, Structural and Reinforcing Steel, Sustainable Timber, PVC, Sustainable Products, Building Re-use) — these rows must always remain in the table.

For "council": return the EXACT council name from the list above that matches the project's suburb/address. Use your knowledge of Melbourne LGA boundaries. Return null if you cannot determine it.

{
  "council": "City of Yarra",${giwXXCredits.length > 0 ? `
  "giwCommentFills": {
${giwXXCredits.map(c => `    "${c.creditId}": "<fill all [XX] in the GIW comment using the credit rawDataPoints above>"`).join(',\n')}
  },` : ''}
  "rowsToDelete": ["exact criteria name 1", "exact criteria name 2", ...],
  "namedValues": {
    "total retail": "<number of retail tenancies from OE 2.x or null>",
    "Total Retail": "<same as total retail>",
    "total office": "<number of office tenancies from OE 2.x or null>",
    "Total Office": "<same as total office>",
    "total area retail": "<total retail floor area m² from OE 2.x or null>",
    "Total area retail": "<same as total area retail>",
    "total area office": "<total office floor area m² from OE 2.x or null>",
    "Total area office": "<same as total area office>",
    "distance to CBD": "<distance in km from project address to Melbourne CBD based on your knowledge of Melbourne geography>",
    "Distance to CBD": "<same as distance to CBD>",
    "toilet WELS": "<toilet WELS star rating from IWM 1.1 or null>",
    "Toilet WELS": "<same as toilet WELS>",
    "taps WELS": "<tap WELS star rating from IWM 1.1 or null>",
    "Taps WELS": "<same as taps WELS>",
    "shower WELS": "<shower WELS star rating from IWM 1.1 or null>",
    "Shower WELS": "<same as shower WELS>",
    "Dishwasher WELS": "<dishwasher WELS star rating from IWM 1.1 or null>",
    "Average star rating": "<average WELS star rating from IWM 1.1, e.g. '4.0', or null>",
    "improvement%": "<thermal performance improvement % above minimum from OE 1.x e.g. '15%', or null>",
    "Hot water": "<hot water system description e.g. 'a heat pump hot water system' from OE 3.x, or null>",
    "solar PV": "<solar PV system size as 'X kW' from OE 4.x, or null>",
    "Solar PV": "<same as solar PV>",
    "Solar PV output": "<annual solar generation in kWh (number only) from OE 4.x, or null>",
    "Solar output": "<same as Solar PV output>",
    "daylight living": "<% of living areas with adequate daylight from IEQ 1.x e.g. '75%', or null>",
    "daylight bedrooms": "<% of bedrooms with adequate daylight from IEQ 1.x e.g. '65%', or null>",
    "daylight non-resi": "<% of non-residential areas with adequate daylight from IEQ rawDataPoints e.g. '60%', or null>",
    "natural ventilation% (XX out of XX)": "<cross-ventilation formatted as 'N% (X out of Y)' from IEQ 2.3, or null>",
    "winter sunlight% (XX out of XX)": "<winter sunlight formatted as 'N% (X out of Y)' from IEQ 1.3, or null>",
    "ventilation non-resi": "<non-residential ventilation approach in one sentence from IEQ 2.3, or null>",
    "Ventilation Non-Resi": "<same as ventilation non-resi>",
    "Shading": "<shading description for residential from IEQ 3.2 or IEQ 3.4 GIW Comment — parse bullet points into a combined sentence, or null>",
    "Shading non-resi": "<shading description for non-residential areas from IEQ GIW comments, or null>",
    "fans": "<% of regular-use areas with ceiling fans from IEQ rawDataPoints (number only), or null>",
    "Fans": "<same as fans>",
    "residential bikes": "<long-stay/residential bicycle spaces from Transport rawDataPoints, or null>",
    "residential visitor bikes": "<short-stay/visitor residential bicycle spaces from Transport rawDataPoints, or null>",
    "employee bikes": "<employee bicycle spaces from Transport rawDataPoints, or null>",
    "Employee bikes": "<same as employee bikes>",
    "commercial visitor bikes": "<commercial visitor bicycle spaces from Transport rawDataPoints, or null>",
    "visitor bikes": "<same as commercial visitor bikes>",
    "EOT showers": "<end-of-trip showers from Transport rawDataPoints, or null>",
    "EOT lockers": "<end-of-trip lockers from Transport rawDataPoints, or null>",
    "motorbikes": "<motorbike spaces from Transport rawDataPoints, or null>",
    "Motorbikes": "<same as motorbikes>",
    "communal area": "<communal open space area m² (number only) from rawDataPoints, or null>",
    "Communal area": "<same as communal area>",
    "vegetation": "<vegetation/permeable coverage % of site (number only) from Urban Ecology rawDataPoints, or null>",
    "Vegetation": "<same as vegetation>",
    "food production area": "<food production garden area m² (number only) from Urban Ecology rawDataPoints, or null>",
    "Food production": "<same as food production area>",
    "Blue Factor score": "<Blue Factor stormwater quality score from IWM 2.1 rawDataPoints or GIW comments, or null>",
    "collection area": "<rainwater catchment/collection area m² from IWM 2.1 rawDataPoints or GIW comments, or null>",
    "raingarden size": "<rain garden area m² from IWM 2.1 rawDataPoints or GIW comments, or null>",
    "raingarden area": "<same as raingarden size>",
    "tap WELS": "<same as taps WELS — tap WELS star rating from IWM 1.1 or null>",
    "total townhouses": "<total number of townhouses from OE 2.x rawDataPoints or null>",
    "orientation% (XX out of XX)": "<percentage of townhouses with good solar orientation formatted as 'N% (X out of Y)' from IEQ 1.3 rawDataPoints, or null>"
  }
}

Return ONLY the JSON, no explanation.`

  const message = await anthropic.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 8192,
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
    // Extract named values — filter out null/empty strings from Claude's response
    const rawNamed = (typeof parsed.namedValues === 'object' && parsed.namedValues !== null)
      ? parsed.namedValues as Record<string, unknown>
      : {}
    const namedValues: Record<string, string> = {}
    for (const [k, v] of Object.entries(rawNamed)) {
      if (v != null && String(v).trim() !== '' && String(v) !== 'null') {
        namedValues[k] = String(v)
      }
    }
    const rawGIW = (typeof parsed.giwCommentFills === 'object' && parsed.giwCommentFills !== null)
      ? parsed.giwCommentFills as Record<string, unknown>
      : {}
    const filledGIWComments: Record<string, string> = {}
    for (const [k, v] of Object.entries(rawGIW)) {
      if (v != null && String(v).trim() !== '' && String(v) !== 'null') {
        filledGIWComments[k.toLowerCase()] = String(v)
      }
    }
    return { filledGIWComments, rowsToDelete, council, namedValues }
  } catch {
    console.error('[report] Claude fill-data response was not valid JSON:', jsonText.slice(0, 500))
    return { filledGIWComments: {}, rowsToDelete: [], council: null, namedValues: {} }
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

/** Return rawDataPoints from OE 3.x credits (hot water profile), falling back to OE 2.x. */
function getHotWaterRaw(credits: ReportCreditData[]): string {
  const oe3 = credits
    .filter(c => /^oe\s*3/i.test(c.creditId))
    .map(c => c.rawDataPoints ?? '')
    .filter(Boolean)
    .join(' ')
  if (oe3) return oe3
  return credits
    .filter(c => /^oe\s*2/i.test(c.creditId))
    .map(c => c.rawDataPoints ?? '')
    .filter(Boolean)
    .join(' ')
}

function extractHwDescription(raw: string): string | null {
  const r = raw.toLowerCase()

  // Prefer explicit "Type of Hot Water System: X" from BESS — use X verbatim (lowercased)
  const typeMatch = raw.match(/type of hot water system\s*:\s*([^\n.;]+)/i)
  if (typeMatch) {
    const bType = typeMatch[1].trim().toLowerCase()
    // Map known BESS dropdown values to report phrasing
    if (/individual.*electric.*instant|electric.*instant/i.test(bType)) return 'individual electric instantaneous hot water systems'
    if (/individual.*electric.*storage/i.test(bType)) return 'individual electric storage hot water systems'
    if (/individual.*gas.*instant|gas.*instant/i.test(bType)) return 'individual gas instantaneous hot water systems'
    if (/individual.*gas.*storage/i.test(bType)) return 'individual gas storage hot water systems'
    if (/central.*heat pump|heat pump.*central/i.test(bType)) return 'a centralised heat pump hot water system'
    if (/heat pump/i.test(bType)) return 'a heat pump hot water system'
    if (/(central|communal).*gas|gas.*(central|communal)/i.test(bType)) return 'a centralised gas hot water system'
    if (/solar.*gas backup|solar.*gas/i.test(bType)) return 'a solar hot water system with gas backup'
    if (/solar.*electric backup|solar.*electric/i.test(bType)) return 'a solar hot water system with electric backup'
    if (/solar/i.test(bType)) return 'a solar hot water system'
    if (/individual.*gas/i.test(bType)) return 'individual gas hot water systems'
    if (/individual.*electric/i.test(bType)) return 'individual electric hot water systems'
    // Fall through to generic keyword matching below if no pattern matched
  }

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
 * Replace all text content in the third cell of the first table row whose
 * first-cell text matches rowPattern. Preserves cell/paragraph/run formatting.
 * Multi-line newText is split into separate paragraphs.
 * If keepPattern is supplied, any existing paragraphs whose plain text matches
 * it are retained at the end of the cell (e.g. "Refer WSUD_Appendix").
 */
function replaceThirdCellContent(
  xml: string,
  rowPattern: RegExp,
  newText: string,
  keepPattern?: RegExp,
): string {
  let i = 0
  while (i < xml.length) {
    const trStart = nextTrOpen(xml, i)
    if (trStart === -1) break
    const trEnd = findRowEnd(xml, trStart)
    const rowXml = xml.slice(trStart, trEnd)
    if (rowPattern.test(getFirstCellText(rowXml))) {
      const cells = [...rowXml.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)]
      if (cells.length >= 3) {
        const thirdCell = cells[2][0]
        const tcPr    = (thirdCell.match(/<w:tcPr\b[\s\S]*?<\/w:tcPr>/) ?? [])[0] ?? ''
        const allParas = [...thirdCell.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map(p => p[0])
        const firstP  = allParas[0] ?? null
        const pPr     = firstP ? ((firstP.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/) ?? [])[0] ?? '') : ''
        const rPr     = firstP ? ((firstP.match(/<w:rPr\b[\s\S]*?<\/w:rPr>/) ?? [])[0] ?? '') : ''
        const kept    = keepPattern
          ? allParas.filter(p => keepPattern.test(p.replace(/<[^>]+>/g, ' ')))
          : []
        const lines   = newText.split('\n').map(l => l.trim()).filter(Boolean)
        const paras   = lines.map(line =>
          `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`
        ).join('')
        const newCell  = `<w:tc>${tcPr}${paras}${kept.join('')}</w:tc>`
        const thirdMatch = cells[2]
        const newRow   = rowXml.slice(0, thirdMatch.index!) + newCell +
                         rowXml.slice(thirdMatch.index! + thirdMatch[0].length)
        return xml.slice(0, trStart) + newRow + xml.slice(trEnd)
      }
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
  // ── Delete retail / office paragraphs and unfilled tags ─────────────────
  if (!hasRetail) {
    xml = deleteParagraphsByText(xml, [
      /^(?:\[XX\]|\d[\d,.]*)m2\s+retail$/i,
      /\[total retail\]/i,
      /\[Total Retail\]/i,
      /\[total area retail\]/i,
      /\[Total area retail\]/i,
    ])
  }
  if (!hasCommercial) {
    xml = deleteParagraphsByText(xml, [
      /^(?:\[XX\]|\d[\d,.]*)m2\s+(?:commercial|office)$/i,
      /\[total office\]/i,
      /\[Total Office\]/i,
      /\[total area office\]/i,
      /\[Total area office\]/i,
    ])
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
 * Post-render fallback: apply targeted fixes that don't rely on positional [XX] filling.
 * Handles: OE 1.1 improvement %, OE 1.3 electrification metering, OE 2.6 gas metering deletion,
 * hot water description correction, and [XX% (XX out of XX)] winter sunlight / natural ventilation.
 */
function applyBESSFallbacks(xml: string, credits: ReportCreditData[], project: ReportProjectData): string {
  // totalApts is used for [XX% (XX out of XX)] winter sunlight / natural ventilation blocks below
  const totalApts = project.totalDwellings ?? getTotalApartments(credits)

  // ── OE 1.1: fill or remove thermal performance improvement % ─────────────
  {
    const oe11Raw = findCredit(credits, 'oe 1.1')?.rawDataPoints ?? ''
    const pctMatch = oe11Raw.match(/(\d+(?:\.\d+)?)\s*%\s*(?:improvement|reduction|above|better)/i) ??
                     oe11Raw.match(/(?:improvement|reduction)[^:]*:\s*(\d+(?:\.\d+)?)\s*%/i)
    const improvement = pctMatch ? parseFloat(pctMatch[1]) : 0

    // Case 1: [by XX%] as a continuous string (template tag named "by XX%")
    if (xml.includes('[by XX%]')) {
      xml = improvement === 0
        ? xml.replace(/\s*\[by XX%\]/g, '')
        : xml.replace(/\[by XX%\]/g, `by ${improvement}%`)
    }

    // Case 2: paragraph-level — covers "by [XX%]"/"by [improvement%]" split across
    // runs, and "by 0%" from Claude. Uses applyParaChanges so "by " and the value
    // are removed as a unit across runs.
    xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
      const texts: string[] = []
      const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
      let m: RegExpExecArray | null
      while ((m = re.exec(para)) !== null) texts.push(m[1])
      const combined = texts.join('')
      if (!/(?:ncc|section j|thermal|envelope|J1|dts)/i.test(combined)) return para
      if (improvement === 0) {
        // Remove "by [improvement%]" or "by 0%" as a single unit
        const byHit = combined.match(/\s*\bby\s+(?:\[improvement%\]|0%)/i)
        if (!byHit || byHit.index === undefined) return para
        return applyParaChanges(para, [{ pos: byHit.index, len: byHit[0].length, replacement: '' }])
      } else {
        // Replace remaining placeholder with the improvement value
        const xxHit = combined.match(/\[improvement%\]/i)
        if (!xxHit || xxHit.index === undefined) return para
        return applyParaChanges(para, [{ pos: xxHit.index, len: xxHit[0].length, replacement: `${improvement}%` }])
      }
    })

  }

  // ── OE 1.3 Electrification: update metering row when all-electric ─────────
  // The metering description lives in Cell 3 (not Cell 2) of the Metering row,
  // so fillTableCellByHeader (which targets Cell 2) doesn't reach it.
  // Use a paragraph-level scan matching the existing "cold water…gas metering" text.
  {
    const electCredit = credits.find(c =>
      /^oe\s*1\.3/i.test(c.creditId) ||
      /electrif/i.test(c.creditId) ||
      /electrif/i.test(c.creditName ?? ''))
    if (electCredit?.creditStatus === 'Y') {
      const meterDesc = 'Individual electricity sub-meters are to be provided for each apartment and tenancy. The development will be all-electric with no gas connection required.'
      xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
        const texts: string[] = []
        const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
        let m: RegExpExecArray | null
        while ((m = re.exec(para)) !== null) texts.push(m[1])
        const combined = texts.join('')
        if (!/cold water.*hot water.*gas metering/i.test(combined)) return para
        return applyParaChanges(para, [{ pos: 0, len: combined.length, replacement: meterDesc }])
      })
    }
  }

  // ── OE 2.6: delete gas metering sentence for commercial tenancy when targeted ─
  {
    const oe26 = findCreditLike(credits, 'oe 2.6')
    if (/^(y|yes|achieved|targeted)$/i.test(oe26?.creditStatus ?? '')) {
      xml = deleteParagraphsByText(xml, [
        /gas metering is to be provided to the commercial tenancy/i,
        /gas metering is to be provided to each individual tenancy requiring a gas connection/i,
      ])
    }
  }

  // ── Hot water: "to utilise X" — correct wrong bare number from Claude ──────
  // Named tag [Hot water] is now filled by docxtemplater; this block only corrects
  // cases where Claude put a bare number into "to utilise N" instead of a description.
  {
    const hwRaw = getHotWaterRaw(credits)
    const allRaw = credits.map(c => c.rawDataPoints ?? '').join(' ')
    const desc = extractHwDescription(hwRaw) ?? extractHwDescription(allRaw)

    xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para) => {
      const texts: string[] = []
      const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
      let m: RegExpExecArray | null
      while ((m = re.exec(para)) !== null) texts.push(m[1])
      const combined = texts.join('')
      if (!/to utilise/i.test(combined)) return para
      const wrongNumMatch = combined.match(/\bto utilise\s+(\d+(?:\.\d+)?)\b/i)
      if (!wrongNumMatch) return para
      const effectiveDesc = desc ?? 'a heat pump hot water system'
      const escaped = escapeXml(effectiveDesc)
      const numStart = combined.search(/\bto utilise\s+\d+(?:\.\d+)?\b/i) + 'to utilise '.length
      return applyParaChanges(para, [{ pos: numStart, len: wrongNumMatch[1].length, replacement: escaped }])
    })
  }

  // ── Winter sunlight: [XX% (XX out of XX)] of apartments achieve … ────────
  if (xml.includes('[XX% (XX out of XX)]')) {
    const ieq13 = findCredit(credits, 'ieq 1.3')
    if (ieq13?.rawDataPoints) {
      const fmt = formatBESSPercentage(ieq13.rawDataPoints, totalApts)
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
 * Delete a section from a Heading2 matching headingPattern through (but not including)
 * the next Heading2. Removes everything in between: paragraphs, tables, etc.
 * Uses character-position slicing so nested tables are handled correctly.
 */
function deleteSectionByHeading(xml: string, headingPattern: RegExp): string {
  // Find the target Heading2 paragraph
  const paraRe = /<w:p\b[\s\S]*?<\/w:p>/g
  let hStart = -1, hEnd = -1
  let match: RegExpExecArray | null
  while ((match = paraRe.exec(xml)) !== null) {
    if (!match[0].includes('Heading2')) continue
    const texts: string[] = []
    const re2 = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re2.exec(match[0])) !== null) texts.push(m[1])
    if (headingPattern.test(texts.join('').trim())) {
      hStart = match.index
      hEnd = match.index + match[0].length
      break
    }
  }
  if (hStart === -1) return xml

  // Find the next Heading2 after the section start
  const afterRe = /<w:p\b[\s\S]*?<\/w:p>/g
  afterRe.lastIndex = hEnd
  let nextStart = xml.length
  while ((match = afterRe.exec(xml)) !== null) {
    if (match[0].includes('Heading2')) {
      nextStart = match.index
      break
    }
  }

  return xml.slice(0, hStart) + xml.slice(nextStart)
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

  const criteriaNames = extractTableRowTexts(docXml)
  const paragraphs = extractDocParagraphs(docXml)

  // Escape long-form [XX...] placeholders before Docxtemplater sees them
  // (Docxtemplater's expression parser can't handle spaces/periods in tag names)
  const escapedDocXml = escapeLongXXTags(docXml)
  zip.file('word/document.xml', escapedDocXml)

  // Ask Claude to fill named values, decide which rows to delete, and fill GIW comment [XX] placeholders
  const { filledGIWComments, rowsToDelete, council, namedValues } = await getWordFillData(project, credits, paragraphs, criteriaNames)

  // Authoritative database fields always override Claude's named values
  const dbNameMap: Array<[string[], string | null]> = [
    [['total apartments', 'total townhouses'], project.totalDwellings != null ? String(project.totalDwellings) : null],
    [['building height', 'Building Height'], project.buildingLevels != null ? String(project.buildingLevels) : null],
    [['site area', 'Site area'], project.siteArea != null ? String(project.siteArea) : null],
    [['BESS score'], project.bessScore != null ? String(Math.round(project.bessScore)) : null],
    [['rainwater tank size'], project.rainwaterTankSize != null ? String(project.rainwaterTankSize) : null],
  ]
  for (const [keys, val] of dbNameMap) {
    if (val != null) {
      for (const k of keys) namedValues[k] = val
    }
  }

  // ── [improvement%] tag: Non-Residential / MixUse templates (OE 1.1 Section J) ─
  // The templates use [improvement%] for the thermal performance improvement %;
  // extract directly from rawDataPoints so docxtemplater can fill it reliably.
  {
    const oe11ImprovRaw = findCredit(credits, 'oe 1.1')?.rawDataPoints ?? ''
    const oe11ImprovMatch =
      oe11ImprovRaw.match(/(\d+(?:\.\d+)?)\s*%\s*(?:improvement|reduction|above|better)/i) ??
      oe11ImprovRaw.match(/(?:improvement|reduction)[^:]*:\s*(\d+(?:\.\d+)?)\s*%/i)
    const oe11ImprovVal = oe11ImprovMatch ? parseFloat(oe11ImprovMatch[1]) : 0
    if (oe11ImprovVal > 0 && !namedValues['improvement%']) {
      namedValues['improvement%'] = `${oe11ImprovVal}%`
      console.log('[report] improvement% set from OE 1.1:', namedValues['improvement%'])
    }
  }

  // ── Shading: code-level extraction from IEQ 3.2 / 3.4 GIW comment ──────────
  // [Shading] / [shading] are standalone paragraph tags filled from commentsGIW.
  // Claude's extraction is unreliable for this — always override with direct parsing.
  {
    const shadingCredit = findCreditLike(credits, 'ieq 3.2') ?? findCreditLike(credits, 'ieq 3.4')
    console.log('[report] Shading credit found:', shadingCredit?.creditId ?? 'none', '| commentsGIW length:', shadingCredit?.commentsGIW?.length ?? 0)
    if (shadingCredit?.commentsGIW?.trim()) {
      const comment = shadingCredit.commentsGIW.trim()
      const bullets = comment
        .split('\n')
        .map(l => l.trim())
        .filter(l => /^[•\-*]/.test(l))
        .map(l => l.replace(/^[•\-*]\s*/, '').trim())
        .filter(Boolean)
      const shadingDesc = bullets.length > 0 ? bullets.join(' ') : comment
      if (shadingDesc) {
        namedValues['Shading'] = shadingDesc
        namedValues['shading'] = shadingDesc
        if (!namedValues['Shading non-resi']) namedValues['Shading non-resi'] = shadingDesc
        console.log('[report] Shading set from GIW comment:', shadingDesc.slice(0, 120))
      }
    }
  }

  // ── IWM 2.1: code-level extraction of stormwater values from GIW comment ─────
  // Blue Factor score, collection area, and raingarden size come from IWM 2.1
  // commentsGIW (or rawDataPoints as fallback). Only fills gaps Claude left empty.
  {
    const iwm21 = findCreditLike(credits, 'iwm 2.1')
    console.log('[report] IWM 2.1 credit found:', iwm21?.creditId ?? 'none', '| commentsGIW length:', iwm21?.commentsGIW?.length ?? 0, '| rawDataPoints length:', iwm21?.rawDataPoints?.length ?? 0)
    const iwm21Comment = iwm21?.commentsGIW?.trim() ?? ''
    const iwm21Raw = iwm21?.rawDataPoints?.trim() ?? ''
    const searchText = iwm21Comment || iwm21Raw

    if (searchText) {
      // When multiple options exist, pick based on reviewer selection
      let workingText = searchText
      const opt1M = searchText.match(/option\s*1\s*[:\-\n]([\s\S]*?)(?=\n\s*option\s*[23]|$)/i)
      if (opt1M) {
        const revText = (iwm21?.reviewerComments ?? []).join(' ').toLowerCase()
        const opt2M = searchText.match(/option\s*2\s*[:\-\n]([\s\S]*?)(?=\n\s*option\s*3|$)/i)
        workingText = opt2M && /option\s*2/i.test(revText) ? opt2M[1] : opt1M[1]
      }

      if (!namedValues['Blue Factor score']) {
        const m = workingText.match(/blue\s*factor[^:\n]*:\s*([\d.]+)/i) ??
                  workingText.match(/([\d.]+)\s*(?:blue\s*factor|bf\b)/i)
        if (m) {
          namedValues['Blue Factor score'] = m[1]
          console.log('[report] Blue Factor score from IWM 2.1:', m[1])
        }
      }
      if (!namedValues['collection area']) {
        const m = workingText.match(/collection\s+area[^:\n]*:\s*([\d,]+)/i) ??
                  workingText.match(/catchment\s+area[^:\n]*:\s*([\d,]+)/i) ??
                  workingText.match(/roof\s+area[^:\n]*:\s*([\d,]+)/i) ??
                  workingText.match(/([\d,]+)\s*m[²2]\s*(?:roof|collect|catch)/i)
        if (m) {
          namedValues['collection area'] = m[1].replace(/,/g, '')
          console.log('[report] Collection area from IWM 2.1:', m[1])
        }
      }
      if (!namedValues['raingarden size'] && !namedValues['raingarden area']) {
        const m = workingText.match(/rain\s*garden[^:\n]*:\s*([\d,]+)/i) ??
                  workingText.match(/([\d,]+)\s*m[²2]\s*rain\s*garden/i)
        if (m) {
          const val = m[1].replace(/,/g, '')
          namedValues['raingarden size'] = val
          namedValues['raingarden area'] = val
          console.log('[report] Raingarden size from IWM 2.1:', val)
        }
      }
    }
  }

  // ── Winter sunlight from IEQ 1.3 ─────────────────────────────────────────
  if (!namedValues['winter sunlight% (XX out of XX)']) {
    const ieq13 = findCreditLike(credits, 'ieq 1.3')
    if (ieq13?.rawDataPoints) {
      const totalApts = project.totalDwellings ?? getTotalApartments(credits)
      const fmt = formatBESSPercentage(ieq13.rawDataPoints, totalApts)
      if (fmt) {
        namedValues['winter sunlight% (XX out of XX)'] = fmt
        console.log('[report] Winter sunlight from IEQ 1.3:', fmt)
      }
    }
  }

  // ── Natural ventilation % from IEQ 2.3 ──────────────────────────────────
  if (!namedValues['natural ventilation% (XX out of XX)']) {
    const ieq23nv = findCreditLike(credits, 'ieq 2.3')
    if (ieq23nv?.rawDataPoints) {
      const totalApts = project.totalDwellings ?? getTotalApartments(credits)
      const fmt = formatBESSPercentage(ieq23nv.rawDataPoints, totalApts)
      if (fmt) {
        namedValues['natural ventilation% (XX out of XX)'] = fmt
        console.log('[report] Natural ventilation from IEQ 2.3:', fmt)
      }
    }
  }

  // ── Ceiling fans % from IEQ rawDataPoints ────────────────────────────────
  if (!namedValues['fans'] && !namedValues['Fans']) {
    const ieqRaw = credits.filter(c => /^ieq/i.test(c.creditId)).map(c => c.rawDataPoints ?? '').join('\n')
    const fansNum = extractNumber(ieqRaw, [
      /(\d+(?:\.\d+)?)\s*%[^.\n]*ceiling\s*fan/i,
      /ceiling\s*fan[^.\n]*?:\s*(\d+(?:\.\d+)?)\s*%/i,
      /(\d+(?:\.\d+)?)\s*%[^.\n]*regular[- ]?use\s*areas?/i,
    ])
    if (fansNum !== null) {
      namedValues['fans'] = String(fansNum)
      namedValues['Fans'] = String(fansNum)
      console.log('[report] Fans % from IEQ:', fansNum)
    }
  }

  // ── Bicycle spaces, EOT facilities, motorbikes from Transport rawDataPoints ─
  {
    const transportRaw = credits.filter(c => /^transport/i.test(c.creditId)).map(c => c.rawDataPoints ?? '').join('\n')
    if (transportRaw.trim()) {
      const set = (keys: string[], val: number) => {
        keys.forEach(k => { if (!namedValues[k]) namedValues[k] = String(val) })
      }

      const resBike = extractNumber(transportRaw, [
        /(\d+)\s*resident\s+bicycle/i,
        /resident\s+bicycle[^.\n]*?:\s*(\d+)/i,
        /(\d+)\s*long[- ]stay\s+bicycle/i,
      ])
      if (resBike !== null) { set(['residential bikes'], resBike); console.log('[report] residential bikes:', resBike) }

      const visResBike = extractNumber(transportRaw, [
        /(\d+)\s*residential\s+visitor\s+bicycle/i,
        /(\d+)\s*visitor\s+bicycle[^.\n]*?resident/i,
        /residential\s+visitor\s+bicycle[^.\n]*?:\s*(\d+)/i,
      ])
      if (visResBike !== null) { set(['residential visitor bikes'], visResBike); console.log('[report] residential visitor bikes:', visResBike) }

      const empBike = extractNumber(transportRaw, [
        /(\d+)\s*(?:employee|staff)\s+bicycle/i,
        /(?:employee|staff)\s+bicycle[^.\n]*?:\s*(\d+)/i,
        /bicycle[^.\n]*?for\s+(?:employees?|staff)[^.\n]*?:\s*(\d+)/i,
      ])
      if (empBike !== null) { set(['employee bikes', 'Employee bikes'], empBike); console.log('[report] employee bikes:', empBike) }

      const nrVisBike = extractNumber(transportRaw, [
        /(\d+)\s*non[- ]?residential\s+visitor\s+bicycle/i,
        /non[- ]?residential\s+visitor\s+bicycle[^.\n]*?:\s*(\d+)/i,
        /non[- ]?residential\s+visitors?[^.\n]*?:\s*(\d+)/i,
        /visitor[^.\n]*?non[- ]?residential[^.\n]*?:\s*(\d+)/i,
        /non[- ]?res(?:idential)?\s+visitor[^.\n]*?:\s*(\d+)/i,
        /(\d+)\s*(?:non[- ]?res|nr)\s+(?:visitor\s+)?bicycle/i,
      ])
      if (nrVisBike !== null) { set(['commercial visitor bikes', 'visitor bikes'], nrVisBike); console.log('[report] commercial visitor bikes:', nrVisBike) }

      const showers = extractNumber(transportRaw, [
        /(\d+)\s*shower/i,
        /shower[^.\n]*?:\s*(\d+)/i,
      ])
      if (showers !== null) { set(['EOT showers'], showers); console.log('[report] EOT showers:', showers) }

      const lockers = extractNumber(transportRaw, [
        /(\d+)\s*locker/i,
        /locker[^.\n]*?:\s*(\d+)/i,
      ])
      if (lockers !== null) { set(['EOT lockers'], lockers); console.log('[report] EOT lockers:', lockers) }

      const moto = extractNumber(transportRaw, [
        /(\d+)\s*motorbike/i,
        /(\d+)\s*moped/i,
        /motorbike[^.\n]*?:\s*(\d+)/i,
        /moped[^.\n]*?:\s*(\d+)/i,
      ])
      if (moto !== null) { set(['motorbikes', 'Motorbikes'], moto); console.log('[report] motorbikes:', moto) }
    }
  }

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

        // Check named values (case-insensitive) — covers all new named placeholders in MixUse/Comm templates
        const tagKey = tag.trim().toLowerCase()
        for (const [k, v] of Object.entries(namedValues)) {
          if (k.trim().toLowerCase() === tagKey) return formatNum(v)
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
    const bessPoints = innovCredit?.rawDataPoints ? parseInnovationPoints(innovCredit.rawDataPoints) : []
    if (bessPoints.length > 0) {
      renderedXml = syncInnovationTable(renderedXml, bessPoints)
    } else {
      // No innovation points detected — delete the entire Innovation section
      try {
        renderedXml = deleteSectionByHeading(renderedXml, /^innovation$/i)
      } catch (e2) {
        console.error('[report] deleteInnovationSection failed:', e2)
      }
    }
  } catch (e) {
    console.error('[report] Innovation table sync failed:', e)
  }

  if (rowsToDelete.length > 0) {
    // These rows must always remain regardless of credit status
    const NEVER_DELETE = /embodied\s*energy|structural.*steel|sustainable\s*timber|\bpvc\b|sustainable\s*products|building\s*re-?use|bicycle\s*parking|end\s*of\s*trip|motorbike/i
    const safeToDelete = rowsToDelete.filter(r => !NEVER_DELETE.test(r))
    if (safeToDelete.length > 0) {
      safeApply('deleteTableRows', () => {
        const result = deleteTableRows(renderedXml, safeToDelete)
        console.log(`[report] Deleted ${safeToDelete.length} table rows:`, safeToDelete)
        return result
      })
    }
  }

  // ── Bicycle/EOT row deletion — only when count is explicitly 0 ────────────
  {
    const transportRaw = credits
      .filter(c => /^transport/i.test(c.creditId))
      .map(c => c.rawDataPoints ?? '')
      .join('\n')

    if (transportRaw.trim()) {
      const resBike = extractNumber(transportRaw, [
        /(\d+)\s*resident\s+bicycle/i,
        /resident\s+bicycle[^.\n]*?:\s*(\d+)/i,
        /(\d+)\s*long[- ]stay\s+bicycle/i,
      ])
      if (resBike === 0) {
        renderedXml = deleteTableRows(renderedXml, ['Bicycle Parking – Residential & Residential Visitors', 'Bicycle Parking – Residential'])
        console.log('[report] Deleted residential bicycle row (count = 0)')
      }

      const empBike = extractNumber(transportRaw, [
        /(\d+)\s*(?:employee|staff)\s+bicycle/i,
        /(?:employee|staff)\s+bicycle[^.\n]*?:\s*(\d+)/i,
      ])
      if (empBike === 0) {
        renderedXml = deleteTableRows(renderedXml, ['Bicycle Parking – Non-Residential & Non-Residential Visitors', 'Bicycle Parking – Non-Residential'])
        console.log('[report] Deleted non-residential bicycle row (count = 0)')
      }

      const showers = extractNumber(transportRaw, [
        /(\d+)\s*shower/i,
        /shower[^.\n]*?:\s*(\d+)/i,
      ])
      if (showers === 0) {
        renderedXml = deleteTableRows(renderedXml, ['End of Trip Facilities – Non-Residential', 'End of Trip Facilities'])
        console.log('[report] Deleted end of trip row (count = 0)')
      }

      const moto = extractNumber(transportRaw, [
        /(\d+)\s*motorbike/i,
        /(\d+)\s*moped/i,
        /motorbike[^.\n]*?:\s*(\d+)/i,
        /moped[^.\n]*?:\s*(\d+)/i,
      ])
      if (moto === 0) {
        renderedXml = deleteTableRows(renderedXml, ['Motorbikes / Mopeds', 'Motorbikes', 'Mopeds'])
        console.log('[report] Deleted motorbike row (count = 0)')
      }
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

  // ── Pre-application meeting dropdown ──────────────────────────────────────
  {
    const mgmt11 = credits.find(c => c.creditId.toLowerCase() === 'management 1.1')
    const preAppAchieved = /^(y|yes|achieved|targeted)$/i.test(mgmt11?.creditStatus ?? '')
    const preAppItem = preAppAchieved
      ? 'GIW has been involved in a pre-application meeting with Council on XXX. '
      : 'GIW has been actively involved in the preliminary design stage, but has not been involved in a pre-application meeting with Council. '
    renderedXml = setDropdownContent(renderedXml, 'GIW has been involved in a pre-application meeting with Council on XXX. ', preAppItem)
  }

  // ── BADS cooling load dropdown ────────────────────────────────────────────
  {
    const postcodeClimateZones = await getPostcodeClimateZoneMap()
    const postcode = extractPostcode(project.address)
    const climateZone = postcode != null ? postcodeClimateZones.get(postcode) : undefined
    const badsItem = climateZone != null ? BADS_COOLING_LOAD[climateZone] : undefined
    if (badsItem != null) {
      renderedXml = setDropdownContent(renderedXml, '30 MJ/m2', badsItem)
    } else {
      renderedXml = highlightDropdown(renderedXml, '30 MJ/m2')
    }
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
    const isTH = /townhouse/i.test(project.typology ?? '')

    if (noFacilities) {
      renderedXml = deleteTableRows(renderedXml, ['Clothes Drying'])
    } else if (/communal|shared|rooftop|terrace/i.test(dryingRaw)) {
      renderedXml = setDropdownContent(renderedXml, 'clothes drying racks on the balcony',
        'Communal clothes drying facilities will be provided at rooftop terrace.')
    } else if (/indoor|internal|inside/i.test(dryingRaw)) {
      renderedXml = setDropdownContent(renderedXml, 'clothes drying racks on the balcony',
        isTH
          ? 'All townhouses will be provided with indoor clothes drying rack / lines.'
          : 'All apartments will be provided with indoor clothes drying rack / lines.')
    } else {
      renderedXml = setDropdownContent(renderedXml, 'clothes drying racks on the balcony',
        isTH
          ? 'All townhouses will be provided with clothes drying facilities in the private open space.'
          : 'All apartments will be provided with clothes drying racks on the balcony.')
    }
  }

  // ── Landscape irrigation dropdown ─────────────────────────────────────────
  {
    const iwm31 = findCreditLike(credits, 'iwm 3.1')
    const iwm31Status = iwm31?.creditStatus ?? ''
    const isScopedOut = /scoped.?out/i.test(iwm31Status)
    const irrigationItem = isScopedOut
      ? 'Landscape irrigation demand will be connected to the rainwater tank. '
      : 'The majority of landscaping is to be native vegetation with no irrigation demand after the initial establishment period.'
    renderedXml = setDropdownContent(renderedXml, 'native vegetation with no irrigation demand', irrigationItem)
  }

  // ── Building Reuse dropdown ────────────────────────────────────────────────
  // Dropdown options: "None of the existing structure is re-used." (default) /
  //   "At least 30% of the existing structure is re-used." / "There is no existing building..."
  // Two dropdown instances exist (one in criteria table, one in a merged continuation row).
  {
    const BUILDING_REUSE_OPT = 'None of the existing structure is re-used.'
    const buildingReuse = credits.find(c =>
      /building\s*re-?use/i.test(c.creditId) ||
      /building\s*re-?use/i.test(c.creditName ?? ''))
    const brStatus = buildingReuse?.creditStatus ?? ''
    const isTargeted = /^(y|achieved|targeted|yes)$/i.test(brStatus)

    if (isTargeted) {
      // Set and flatten both dropdown instances (loop until none remain)
      let attempts = 0
      while (renderedXml.includes(`displayText="${BUILDING_REUSE_OPT}"`) && attempts < 4) {
        renderedXml = setDropdownContent(renderedXml, BUILDING_REUSE_OPT, 'At least 30% of the existing structure is re-used.')
        renderedXml = flattenDropdown(renderedXml, BUILDING_REUSE_OPT)
        attempts++
      }
    } else {
      // Not targeted (including Scoped Out) — delete all Building Reuse table rows
      renderedXml = deleteTableRows(renderedXml, ['Building Re-use', 'Building Reuse'])
      renderedXml = deleteTableRowsContaining(renderedXml, BUILDING_REUSE_OPT)
    }
  }

  // ── GDFT: update star rating in Thermal Performance Rating – Residential ──
  if (project.gdft) {
    let i = 0
    while (i < renderedXml.length) {
      const trStart = nextTrOpen(renderedXml, i)
      if (trStart === -1) break
      const trEnd = findRowEnd(renderedXml, trStart)
      const rowXml = renderedXml.slice(trStart, trEnd)
      const firstCell = getFirstCellText(rowXml)
      if (/Thermal\s*Performance\s*Rating/i.test(firstCell) && !/Non.?Residential/i.test(firstCell)) {
        renderedXml = renderedXml.slice(0, trStart) +
          rowXml.replace(/6\.0 Stars/g, '6.5 Stars') +
          renderedXml.slice(trEnd)
        break
      }
      i = trEnd
    }
  }

  // ── GIW-comment direct cell replacements ─────────────────────────────────
  // For these rows the GIW comment is the authoritative content — replace the
  // entire third cell rather than relying on Claude to fill named placeholders.
  // Use filledGIWComments (Claude-filled [XX] placeholders) when available.
  {
    const waste21 = findCreditLike(credits, 'waste 2.1')
    if (waste21?.commentsGIW?.trim())
      renderedXml = replaceThirdCellContent(renderedXml, /Food\s*(?:&amp;|&|and)\s*Garden\s*Waste/i,
        filledGIWComments['waste 2.1'] ?? waste21.commentsGIW.trim())

    const waste22 = findCreditLike(credits, 'waste 2.2')
    if (waste22?.commentsGIW?.trim())
      renderedXml = replaceThirdCellContent(renderedXml, /Convenience\s*of\s*Recycl/i,
        filledGIWComments['waste 2.2'] ?? waste22.commentsGIW.trim())

    const iwm21 = findCreditLike(credits, 'iwm 2.1')
    if (iwm21?.commentsGIW?.trim())
      renderedXml = replaceThirdCellContent(renderedXml, /^Stormwater\s*Treatment$/i,
        filledGIWComments['iwm 2.1'] ?? iwm21.commentsGIW.trim(), /Refer\s*WSUD/i)

    const shadingCredit = findCreditLike(credits, 'ieq 3.2') ?? findCreditLike(credits, 'ieq 3.4')
    if (shadingCredit?.commentsGIW?.trim()) {
      const shadingId = shadingCredit.creditId.toLowerCase()
      const shadingComment = filledGIWComments[shadingId] ?? shadingCredit.commentsGIW.trim()
      renderedXml = replaceThirdCellContent(renderedXml, /^Thermal\s*Comfort$/i, shadingComment)
      renderedXml = replaceThirdCellContent(renderedXml, /^Thermal\s*Comfort\s*[–\-]/i, shadingComment)
    }
  }

  // ── Daylight pathway detection ────────────────────────────────────────────
  // BESS rawDataPoints key phrases per pathway:
  //   DTS:      "use the bess deemed to satisfy (dts) method?: yes"
  //   Built-in: "what calculation approach do you want to use?: use the built-in calculation tools"
  //   Modelling:"what calculation approach do you want to use?: provide your own calculations"
  {
    const daylightRaw = credits
      .filter(c => /ieq 1\.[12]/i.test(c.creditId))
      .map(c => c.rawDataPoints ?? '')
      .join('\n')
    // Parse table rows: "Question?:    Answer" — check the answer side, not just presence of keywords
    const usedDTS     = /deemed\s*to\s*satisfy[^:\n]*\?:\s*yes\b/i.test(daylightRaw)
    const usedBuiltIn = /calculation\s+approach[^:\n]*\?:\s*use\s+the\s+built[- ]?in\s+calculation/i.test(daylightRaw)

    // Delete DTS body paragraphs when modelling or built-in calculator was used
    if (!usedDTS) {
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

    // Delete the daylight modelling appendix and body reference for DTS and built-in pathways
    if (usedDTS || usedBuiltIn) {
      try {
        renderedXml = deleteSectionByHeading(renderedXml, /Appendix.*Daylight\s*Modelling/i)
      } catch (e) {
        console.error('[report] deleteDaylightAppendix failed:', e)
      }
      renderedXml = deleteParagraphsByText(renderedXml, [/Refer\s+Appendix.*Daylight\s*Modelling/i])
    }

    // DTS pathway: replace criteria table intro with DTS sentence, delete % result lines
    if (usedDTS) {
      const dtsSentence = 'The development complies with the BESS Deemed-to-Satisfy method for Daylight.'
      renderedXml = renderedXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para: string) => {
        const texts: string[] = []
        const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
        let m: RegExpExecArray | null
        while ((m = re.exec(para)) !== null) texts.push(m[1])
        const combined = texts.join('')
        if (/daylight modelling has been conducted for a representative/i.test(combined))
          return applyParaChanges(para, [{ pos: 0, len: combined.length, replacement: dtsSentence }])
        if (/% of (?:living|bedroom) floor area/i.test(combined))
          return ''
        return para
      })
    }

    // Built-in calculator pathway: replace criteria table intro only (% lines remain)
    if (usedBuiltIn) {
      const builtInSentence = 'The BESS built-in daylight calculator has been used to assess compliance. The summary result is as follows:'
      renderedXml = renderedXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para: string) => {
        const texts: string[] = []
        const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
        let m: RegExpExecArray | null
        while ((m = re.exec(para)) !== null) texts.push(m[1])
        const combined = texts.join('')
        if (!/daylight modelling has been conducted for a representative/i.test(combined)) return para
        return applyParaChanges(para, [{ pos: 0, len: combined.length, replacement: builtInSentence }])
      })
    }
  }

  // ── OE 1.1: Non-residential thermal — DTS vs J1V3 energy modelling pathway ──
  // BESS question: "Use the BESS Deem to Satisfy (DtS) method for Non-residential spaces?:"
  // DTS pathway    → delete the J1V3 energy consumption sentence and J1V3 appendix reference
  // J1V3 pathway   → delete DTS-specific insulation and façade calculator paragraphs
  {
    const oe11Raw = findCredit(credits, 'oe 1.1')?.rawDataPoints ?? ''
    const usedDTS = /use the bess deem to satisfy[^:\n]*non-residential spaces\?:\s*yes/i.test(oe11Raw)

    if (usedDTS) {
      renderedXml = deleteParagraphsByText(renderedXml, [
        /conditioned areas aim to reduce heating and cooling energy consumption/i,
        /refer\s+appendix.*j1v3\s*energy\s*modelling/i,
      ])
      try {
        renderedXml = deleteSectionByHeading(renderedXml, /appendix.*j1v3.*energy\s*modelling/i)
      } catch (e) {
        console.error('[report] deleteJ1V3Appendix failed:', e)
      }
    } else {
      renderedXml = deleteParagraphsByText(renderedXml, [
        /all exposed floors and ceilings.*envelope.*insulation/i,
        /all wall and glazing.*facade calculator/i,
        /refer\s+appendix.*j4d6/i,
      ])
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

    // Delete only when the value is confirmed 0 (or credit not achieved)
    const ieq23NotAchieved = !ieq23 || ieq23.creditStatus === 'ScopedOut' || ieq23.creditStatus === 'N'
    if (naturalPct === 0)
      renderedXml = deleteParagraphsByText(renderedXml, [/non-residential spaces is naturally ventilated/i])
    if (outdoorAirPct === 0)
      renderedXml = deleteParagraphsByText(renderedXml, [/increase in outdoor air rates/i])
    // Only delete CO2 sensors paragraph when we know CO2 monitoring isn't required:
    // either explicitly 0 ppm, credit not achieved, or no CO2 mention in rawDataPoints at all
    if (co2Ppm === 0 || (ieq23NotAchieved && co2Ppm === null) ||
        (ieq23 && raw && co2Ppm === null && !/CO2|carbon dioxide|sensor/i.test(raw)))
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
    /(?:\b0%\b|\[XX\]%|\[XX%\]|\[fans\]%|\[Fans\]%).*ceiling fan/i,
    /ceiling fan.*(?:\b0%\b|\[XX\]%|\[XX%\]|\[fans\]%|\[Fans\]%)/i,
  ])

  // ── Retail / commercial line items ────────────────────────────────────────
  const { hasRetail, hasCommercial } = detectNonResidential(credits, project.typology ?? null)
  safeApply('removeNonResidentialLines', () => removeNonResidentialLines(renderedXml, hasRetail, hasCommercial))

  // ── WELS star rating: strip [N] brackets left by Docxtemplater nullGetter ──
  // The template has [4]/[5] as tag-delimited placeholders (e.g. "WELS [4] Star - Toilets").
  // Each bracket is a separate run, so applyParaChanges handles the split correctly.
  renderedXml = renderedXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (para: string) => {
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(para)) !== null) texts.push(m[1])
    const combined = texts.join('')
    if (!combined.includes('WELS')) return para
    const changes: Array<{ pos: number; len: number; replacement: string }> = []
    const bracketRe = /\[(\d+)\]/g
    let bm: RegExpExecArray | null
    while ((bm = bracketRe.exec(combined)) !== null) {
      changes.push({ pos: bm.index, len: bm[0].length, replacement: bm[1] })
    }
    if (!changes.length) return para
    return applyParaChanges(para, changes)
  })

  // ── BESS rawDataPoints fallbacks (catches any unfilled placeholders) ──────
  safeApply('applyBESSFallbacks', () => applyBESSFallbacks(renderedXml, credits, project))

  // ── Solar PV: delete renewable energy appendix when no PV system ──────────
  {
    const solarRaw = credits
      .filter(c => /^oe\s*4/i.test(c.creditId))
      .map(c => c.rawDataPoints ?? '')
      .join(' ')
    // Match PV-specific terms; exclude solar thermal / solar hot water which are not PV
    const hasSolarPv = /\bsolar\s+pv\b|\bphotovoltaic\b|\bpv\s+system\b|\bpv\s+panel|\bsolar\s+panel|\bpv\s+array/i.test(solarRaw) ||
      (/\bsolar\b/i.test(solarRaw) && /\bkw\b/i.test(solarRaw) && !/solar\s+(?:hot\s+water|thermal|hws|dhw)/i.test(solarRaw))
    if (!hasSolarPv) {
      try {
        renderedXml = deleteSectionByHeading(renderedXml, /renewable\s+energy|solar\s+pv/i)
      } catch (e) {
        console.error('[report] deleteSolarPVAppendix failed:', e)
      }
    }
  }

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

  if (project.totalDwellings != null) ws = patchCell(ws, 'D6', project.totalDwellings)

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
    const oe34 = findCreditLike(credits, 'oe 3.4')
    if (oe34?.rawDataPoints) {
      const raw = oe34.rawDataPoints
      const dryRef = extractNumber(raw, [
        /reference[^:\n]*:\s*([\d,.]+)\s*kwh/i,
        /([\d,.]+)\s*kwh[^.\n]*reference/i,
      ])
      const dryProp = extractNumber(raw, [
        /proposed[^:\n]*:\s*([\d,.]+)\s*kwh/i,
        /([\d,.]+)\s*kwh[^.\n]*proposed/i,
      ])
      if (dryRef != null) ws = patchCell(ws, 'D33', dryRef)
      if (dryProp != null) ws = patchCell(ws, 'D34', dryProp)
    }
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
    const raw = iwm11.rawDataPoints
    const refWater = extractNumber(raw, [
      /reference\s*potable\s*water\s*use[^:\n]*:\s*([\d,.]+)\s*kl/i,
      /reference[^:\n]*:\s*([\d,.]+)\s*kl/i,
    ])
    const propWaterExcl = extractNumber(raw, [
      /proposed[^:\n]*excluding[^:\n]*:\s*([\d,.]+)\s*kl/i,
      /proposed[^:\n]*excl[^:\n]*:\s*([\d,.]+)\s*kl/i,
    ])
    const propWaterIncl = extractNumber(raw, [
      /proposed[^:\n]*including[^:\n]*:\s*([\d,.]+)\s*kl/i,
      /proposed[^:\n]*incl[^:\n]*:\s*([\d,.]+)\s*kl/i,
    ])
    const pctReduction = extractNumber(raw, [/([\d.]+)%\s*reduction/i, /reduction[^:\n]*:\s*([\d.]+)%/i])
    if (refWater != null) ws = patchCell(ws, 'Q51', refWater)
    if (propWaterExcl != null) ws = patchCell(ws, 'Q52', propWaterExcl)
    if (propWaterIncl != null) ws = patchCell(ws, 'Q53', propWaterIncl)
    if (pctReduction != null) ws = patchCell(ws, 'D53', pctReduction)
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
    const raw = iwm11.rawDataPoints
    const refWater = extractNumber(raw, [
      /reference\s*potable\s*water\s*use[^:\n]*:\s*([\d,.]+)\s*kl/i,
      /reference[^:\n]*:\s*([\d,.]+)\s*kl/i,
    ])
    const propWaterExcl = extractNumber(raw, [
      /proposed[^:\n]*excluding[^:\n]*:\s*([\d,.]+)\s*kl/i,
      /proposed[^:\n]*excl[^:\n]*:\s*([\d,.]+)\s*kl/i,
    ])
    const propWaterIncl = extractNumber(raw, [
      /proposed[^:\n]*including[^:\n]*:\s*([\d,.]+)\s*kl/i,
      /proposed[^:\n]*incl[^:\n]*:\s*([\d,.]+)\s*kl/i,
    ])
    const pctReduction = extractNumber(raw, [/([\d.]+)%\s*reduction/i, /reduction[^:\n]*:\s*([\d.]+)%/i])
    if (refWater != null) ws = patchCell(ws, 'D62', refWater)
    if (propWaterExcl != null) ws = patchCell(ws, 'D63', propWaterExcl)
    if (propWaterIncl != null) ws = patchCell(ws, 'D64', propWaterIncl)
    if (pctReduction != null) ws = patchCell(ws, 'D56', pctReduction)
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
  // Always use today as the generation date — project.date is the lodgement date, not the report date
  const dateObj = new Date()
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
