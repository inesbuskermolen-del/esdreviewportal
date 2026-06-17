import * as fs from 'fs'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'

const buf = fs.readFileSync('./templates/SMP-MixUse.docx', 'binary')
const zip = new PizZip(buf)

const doc = new Docxtemplater(zip, {
  delimiters: { start: '[', end: ']' },
  paragraphLoop: true,
  linebreaks: true,
  nullGetter(part: { value: string }) { return '[' + part.value + ']' },
})

doc.render({
  GIWREF: 'GIW-TEST',
  'Project Address': '98 Riversdale Road, Hawthorn',
  Client: 'Test Client',
  Architect: 'Test Architect',
  Date: '08/06/2026',
  'SMP Visualisation. Copy from C2 to U25 from excel sheet. Paste as image': '[Chart placeholder]',
})

const out = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer
console.log('Docxtemplater render OK, size:', out.length)

// Validate: re-open as ZIP and check XML
const check = new PizZip(out.toString('binary'))
const docXml = check.file('word/document.xml')!.asText()
console.log('document.xml length:', docXml.length)

// Check for balanced tags in critical areas
const openTr = (docXml.match(/<w:tr[ >]/g) || []).length
const closeTr = (docXml.match(/<\/w:tr>/g) || []).length
const openP = (docXml.match(/<w:p[ >]/g) || []).length
const closeP = (docXml.match(/<\/w:p>/g) || []).length
console.log(`<w:tr>: ${openTr} open, ${closeTr} close`)
console.log(`<w:p>:  ${openP} open, ${closeP} close`)

// Check for XML validity issues in the rendered output
const textRuns = docXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []
let xmlIssues = 0
for (const run of textRuns) {
  const inner = run.replace(/<w:t[^>]*>/, '').replace(/<\/w:t>/, '')
  if (inner.includes('&') && !inner.match(/&(?:amp|lt|gt|quot|apos);/)) {
    console.log('Unescaped & in text run:', inner.slice(0, 100))
    xmlIssues++
  }
}
console.log('Text runs:', textRuns.length, ', XML char issues:', xmlIssues)

// Check if [SMP Visualisation...] value (which contains brackets) appears in XML
const vizTag = '[SMP Visualisation chart'
if (docXml.includes(vizTag)) {
  console.log('SMP vis placeholder found in XML (raw brackets in XML text)')
}

fs.writeFileSync('C:\\Users\\inesb\\Desktop\\test-raw.docx', out)
console.log('Written to Desktop/test-raw.docx')
