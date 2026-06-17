import * as fs from 'fs'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import { PrismaClient } from '@prisma/client'

// Inline the post-processing functions we want to test
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

function findRowEnd(xml: string, trStart: number): number {
  let depth = 0
  let i = trStart
  while (i < xml.length) {
    const nextOpen = nextTrOpen(xml, i)
    const nextClose = xml.indexOf('</w:tr>', i)
    if (nextClose === -1) return xml.length
    if (nextOpen !== -1 && nextOpen < nextClose) { depth++; i = nextOpen + 5 }
    else { depth--; if (depth === 0) return nextClose + 7; i = nextClose + 7 }
  }
  return xml.length
}

function getFirstCellText(rowXml: string): string {
  const tcMatch = rowXml.match(/<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/)
  if (!tcMatch) return ''
  const texts: string[] = []
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(tcMatch[1])) !== null) { if (m[1].trim()) texts.push(m[1]) }
  return texts.join('').trim()
}

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
    } else { result += rowXml }
    i = trEnd
  }
  if (orphanedBookmarkIds.size > 0) {
    result = result.replace(/<w:bookmarkEnd[^>]*\bw:id="(\d+)"[^/]*(\/?>)/g, (match, id) =>
      orphanedBookmarkIds.has(id) ? '' : match)
  }
  return result
}

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

function addCategoryPageBreaks(xml: string): string {
  const SEP = '</w:p>'
  let pastEsdAssessment = false
  return xml.split(SEP).map(para => {
    if (!para.includes('w:val="Heading2"')) return para
    const texts: string[] = []
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(para)) !== null) texts.push(m[1])
    const line = texts.join('').trim()
    if (!pastEsdAssessment) {
      if (line === 'ESD Assessment') pastEsdAssessment = true
      return para
    }
    if (para.includes('<w:pageBreakBefore')) return para
    if (para.includes('<w:pPr>')) {
      if (/<w:pStyle\b[^>]*\/>/.test(para)) {
        return para.replace(/<w:pStyle\b[^>]*\/>/, m => m + '<w:pageBreakBefore/>')
      }
      return para.replace('<w:pPr>', '<w:pPr><w:pageBreakBefore/>')
    }
    return para.replace(/(<w:r[ >])/, (_, run: string) => `<w:pPr><w:pageBreakBefore/></w:pPr>${run}`)
  }).join(SEP)
}

function numberAppendices(xml: string): string {
  const SEP = '</w:p>'
  const parts = xml.split(SEP)
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
  const keyToLetter = new Map(order.map((k, i) => [k, String.fromCharCode(66 + i)]))
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
    let p = part.replace(/Appendix X/g, `Appendix ${letter}`)
    p = p.replace(/<w:t([^>]*)>X<\/w:t>/g, `<w:t$1>${letter}</w:t>`)
    return p
  }).join(SEP)
}

function checkXml(label: string, xml: string) {
  const selfClose = (xml.match(/<w:p\b[^>]*\/>/g) || []).length
  const openP = (xml.match(/<w:p[ >]/g) || []).length - selfClose
  const closeP = (xml.match(/<\/w:p>/g) || []).length
  const openTr = (xml.match(/<w:tr[ >]/g) || []).length
  const closeTr = (xml.match(/<\/w:tr>/g) || []).length
  const openTbl = (xml.match(/<w:tbl[ >]/g) || []).length
  const closeTbl = (xml.match(/<\/w:tbl>/g) || []).length
  const pOk = openP === closeP ? '✓' : `✗ (${openP} vs ${closeP})`
  const trOk = openTr === closeTr ? '✓' : `✗ (${openTr} vs ${closeTr})`
  const tblOk = openTbl === closeTbl ? '✓' : `✗ (${openTbl} vs ${closeTbl})`
  const len = xml.length
  const status = (openP === closeP && openTr === closeTr && openTbl === closeTbl) ? 'OK' : 'BROKEN'
  console.log(`[${status}] ${label} (len=${len}): p=${pOk} tr=${trOk} tbl=${tblOk}`)
  return openP === closeP && openTr === closeTr && openTbl === closeTbl
}

async function main() {
  const prisma = new PrismaClient()
  const PROJECT_ID = 'cmpvnuyct0000w8dfx1akep9c' // Mixed-Use

  const buf = fs.readFileSync('./templates/SMP-MixUse.docx', 'binary')
  const zip = new PizZip(buf)
  const doc = new Docxtemplater(zip, {
    delimiters: { start: '[', end: ']' },
    paragraphLoop: true,
    linebreaks: true,
    nullGetter(part: { value: string }) { return '[' + part.value + ']' },
  })
  doc.render({ GIWREF: 'GIW-TEST', 'Project Address': '98 Riversdale Rd Hawthorn', Client: 'Test', Architect: 'Test', Date: '08/06/2026', 'SMP Visualisation. Copy from C2 to U25 from excel sheet. Paste as image': '[chart]' })
  const renderedZip = doc.getZip()
  let xml = renderedZip.file('word/document.xml')!.asText()

  checkXml('after docxtemplater render', xml)

  // Step 1: deleteTableRows
  const toDelete = [
    'Pre-Application Meeting', 'Clothes Drying', 'Daylight Access – Non-Residential',
    'Minimal Internal Bedrooms', 'Ventilation – Non-Residential', 'Thermal Comfort',
    'Thermal Comfort – Non-Residential', 'Air Quality – Non-Residential',
    'End of Trip Facilities – Non-Residential', 'Car Share Scheme', 'Motorbikes / Mopeds',
    'Embodied Energy', 'Structural and Reinforcing Steel', 'Sustainable Timber', 'PVC',
    'Sustainable Products', 'Building Re-use', 'Construction and Demolition Waste',
    'Communal Space', 'Green Walls / Roof', 'Food Production - Residential',
    'Heat Island Effect', 'Materials Exchange', 'Life Cycle Assessment',
    'Carbon Neutral Ready Development', 'ESD Checkpoint during Construction Phase',
  ]
  xml = deleteTableRows(xml, toDelete)
  checkXml('after deleteTableRows', xml)

  // Step 2: deleteParagraphsByText (daylight DTS)
  xml = deleteParagraphsByText(xml, [/^Or$/i, /Deemed-to-Satisfy method for IEQ/i, /All North.*West.*East.*8m/i, /floor-to-ceiling height of 2\.7/i, /60%.*visible light transmittance/i, /All living areas have an external facing window/i, /building separation tables/i])
  checkXml('after deleteParagraphsByText(DTS)', xml)

  // Step 3: addCategoryPageBreaks
  xml = addCategoryPageBreaks(xml)
  checkXml('after addCategoryPageBreaks', xml)

  // Step 4: numberAppendices
  xml = numberAppendices(xml)
  checkXml('after numberAppendices', xml)

  // Save final version
  renderedZip.file('word/document.xml', xml)
  const out = renderedZip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer
  fs.writeFileSync('C:\\Users\\inesb\\Desktop\\test-postprocessed.docx', out)
  console.log('Written to Desktop/test-postprocessed.docx')

  await prisma.$disconnect()
}

main().catch(e => { console.error(e.message); process.exit(1) })
