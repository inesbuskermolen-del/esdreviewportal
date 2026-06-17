import * as fs from 'fs'
import PizZip from 'pizzip'

function setDropdownContent(xml: string, identifyingOption: string, selectedValue: string): string {
  let pos = xml.indexOf(`displayText="${identifyingOption}"`)
  if (pos === -1) {
    const escaped = identifyingOption.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const m = xml.match(new RegExp(`displayText="[^"]*${escaped}[^"]*"`))
    if (!m) return xml
    pos = xml.indexOf(m[0])
    if (pos === -1) return xml
  }

  const sdtPrStart = xml.lastIndexOf('<w:sdtPr', pos)
  const sdtStart = sdtPrStart !== -1 ? xml.lastIndexOf('<w:sdt', sdtPrStart) : -1
  const contentStart = xml.indexOf('<w:sdtContent', pos)
  if (contentStart === -1) return xml
  const contentEnd = xml.indexOf('</w:sdtContent>', contentStart) + '</w:sdtContent>'.length
  const sdtEnd = xml.indexOf('</w:sdt>', contentEnd) + '</w:sdt>'.length

  const escaped = selectedValue.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

  let content = xml.slice(contentStart, contentEnd)
  content = content.replace(/<w:t[^>]*>[^<]*<\/w:t>/, `<w:t>${escaped}</w:t>`)
  content = content.replace(/<w:rStyle w:val="PlaceholderText"\/>/g, '')
  content = content.replace(/<w:color [^>]*\/>/g, '<w:color w:val="000000"/>')
  if (!content.includes('<w:color ')) {
    content = content.replace(/(<w:rPr>)/, '$1<w:color w:val="000000"/>')
  }

  const before = sdtStart !== -1 ? xml.slice(0, sdtStart) : xml.slice(0, contentStart)
  const sdtPrBlock = sdtStart !== -1 ? xml.slice(sdtStart, contentStart) : ''
  const cleanedSdtPr = sdtPrBlock.replace(/<w:showingPlcHdr\/>/g, '')
  const after = sdtEnd !== -1 ? xml.slice(sdtEnd) : xml.slice(contentEnd)
  const sdtClose = sdtEnd !== -1 ? xml.slice(contentEnd, sdtEnd) : ''

  return before + cleanedSdtPr + content + sdtClose + after
}

function check(label: string, xml: string) {
  const selfClose = (xml.match(/<w:p\b[^>]*\/>/g) || []).length
  const openP = (xml.match(/<w:p[ >]/g) || []).length - selfClose
  const closeP = (xml.match(/<\/w:p>/g) || []).length
  const openSdt = (xml.match(/<w:sdt[ >]/g) || []).length
  const closeSdt = (xml.match(/<\/w:sdt>/g) || []).length
  const ok = openP === closeP && openSdt === closeSdt
  const suffix = ok ? '' : ` p=${openP}/${closeP} sdt=${openSdt}/${closeSdt}`
  console.log(`[${ok ? 'OK' : 'BROKEN'}] ${label} (len=${xml.length})${suffix}`)
}

const buf = fs.readFileSync('C:\\Users\\inesb\\Desktop\\test-postprocessed.docx')
const zip = new PizZip(buf.toString('binary'))
let xml = zip.file('word/document.xml')!.asText()
check('initial', xml)

const council = 'City of Stonnington'
const esc = council.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
const esdMatch = (xml.match(new RegExp(`displayText="(${esc}[^"]*sustainability[^"]*)"`, 'i')) || [])[1] ?? null
console.log('esdMatch:', esdMatch)
if (esdMatch) xml = setDropdownContent(xml, esdMatch, esdMatch)
check('after council ESD dropdown', xml)

const shortMatch = (xml.match(new RegExp(`displayText="(${esc})"`, '')) || [])[1] ?? null
console.log('shortMatch:', shortMatch)
if (shortMatch) xml = setDropdownContent(xml, shortMatch, shortMatch)
check('after council short dropdown', xml)

xml = setDropdownContent(xml, 'a centralised gas hot water system', 'a heat pump hot water system')
check('after hot water dropdown', xml)

xml = setDropdownContent(xml, 'GIW has been involved in a pre-application meeting with Council on XXX. ', 'GIW has been actively involved in the preliminary design stage, but has not been involved in a pre-application meeting with Council. ')
check('after pre-app dropdown', xml)

xml = setDropdownContent(xml, 'clothes drying racks on the balcony', 'All apartments will be provided with clothes drying racks on the balcony.')
check('after clothes drying dropdown', xml)

// Landscape irrigation
xml = setDropdownContent(xml, 'native vegetation with no irrigation demand', 'The majority of landscaping is to be native vegetation with no irrigation demand after the initial establishment period.')
check('after landscape irrigation dropdown', xml)

// Write result
zip.file('word/document.xml', xml)
const out = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer
fs.writeFileSync('C:\\Users\\inesb\\Desktop\\test-with-dropdowns.docx', out)
console.log('Written test-with-dropdowns.docx, size:', out.length)
