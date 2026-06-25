import Anthropic from '@anthropic-ai/sdk'
import { prisma } from './prisma'

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

/* ── Per-credit GIW comment templates (from ESD Review Matrix.xlsx) ── */

const CREDIT_COMMENT_TEMPLATES: Record<string, string[]> = {
  'management 1.1': ['GIW is to attend a pre-application meeting with Council DTP. Town planner to coordinate.'],
  'management 2.1': ['A prelimimary FirstRate assessment will be undertaken for the residence.'],
  'management 2.2': ['A sample of prelimimary FirstRate assessment will be undertaken for the development.'],
  'management 2.3': [
    'A preliminary facade assessment will be undertaken in accordance with NCC2022 Section J4D6.',
    'Preliminary J1V3 modelling will be undertaken for the development.',
  ],
  'management 3.1': ['Utility meters (electricity and water) will be provided for all individual dwellings.'],
  'management 3.2': ['Utility meters (electricity and water) will be provided for all individual tenancies.'],
  'management 3.3': ['Utility sub-meters will be provided for all major common area services.'],
  'management 4.1': ['A building user guide will be produced and issued to occupants.'],
  'iwm 1.1': [
    '• Showerheads: [stars] Star WELS\n• Kitchen Taps: [stars] Star WELS\n• Bathroom Taps: [stars] Star WELS\n• Dishwashers: [stars] Star WELS\n• WC: [stars] Star WELS\n• Urinals: Scope out\n• Washing Machine: Occupant to Install',
  ],
  'iwm 2.1': [
    'A compliant Blue Factor result can be achieved via the following options.\n\nOption 1:\n• Rainwater collection off [XX] is to be directed into a [XX]-litre rainwater tank connected to [XX] toilets and landscape irrigation.\n• Rainwater collection off [XX] is to be directed in a ≥[XX]m2 raingarden with 100mm of extended detention.\n\nOption 2:\n• Rainwater collection off [XX] is to be directed into a [XX]-litre rainwater tank connected to [XX] toilets and landscape irrigation.',
  ],
  'iwm 3.1': [
    'Landscaping is either native vegetation with no water demand after the initial establishment period OR landscape irrigation is to be connected to the rainwater tank. Please confirm preference?',
  ],
  'iwm 4.1': [
    '80% of fire system test water (e.g. hydrant pump test water or SCV annubar test) is to be reused on-site, either within the fire system or directed into the rainwater tank OR fire test water system does not expel water.',
  ],
  'oe 1.1': [
    'GIW has undertaken a preliminary facade assessment in accordance with NCC2022 Section J4D6 and recommend the application of the DtS pathway for Section J compliance.',
    'GIW has undertaken a preliminary facade assessment in accordance with NCC2022 Section J4D6 and recommend the application of the J1V3 pathway for Section J compliance. This will be undertaken during the DD stage.',
  ],
  'oe 1.2': [
    'The energy ratings are to achieve a [XX] Star average with no unit below 6 Stars and no unit exceeding the maximum allowed cooling loads as outlined under BADS.',
  ],
  'oe 2.1': ['Refer credit OE2.7 Energy Consumption'],
  'oe 2.2': ['>10% peak energy demand reduction is achieved.'],
  'oe 2.6': ['The development is all electric with induction cooktops and no gas connection.'],
  'oe 2.7': [
    'HVAC systems to have an Energy Efficiency Ratios (EER) not less than 85% of the EER of the most efficient equivalent unit available (type & capacity). Where VRV / VRF systems are proposed, a minimum COP of 3.4 is required.',
  ],
  'oe 3.1': ['Carpark ventilation fans are to be controlled by CO sensors.'],
  'oe 3.2': ['Centralised heat pump hot water system or individual electric instantaneous hot water systems. Please confirm?'],
  'oe 3.3': ['Operation of min. 50% of external lighting is controlled by a motion detector.'],
  'oe 3.4': ['[Individual / shared] clothes drying lines are to be introduced to the development.'],
  'oe 3.5': ['Maximum illumination power density for each dwelling is 4W/sqm or less.'],
  'oe 3.6': ['Lighting power density shall be as follows:\n• Dwellings: No greater than average 4W/m2\n• POS: No greater than average 3.2W/m2\n• Back of house and indoor car parks: No greater than average 1.6W/m2'],
  'oe 3.7': ['Lighting power density shall be as follows:\n• Retail: No greater than average 14W/m2\n• Office: No greater than average 4.5W/m2'],
  'oe 4.1': ['Combined heat and power system is to be introduced.'],
  'oe 4.2': ['A total [XX]kW solar PV system is to be installed at roof. The system is to be installed facing due [XX] at a [XX] inclination.'],
  'oe 4.4': ['A geothermal system is introduced to the development.'],
  'oe 4.5': ['A total [XX]kW solar PV system is to be installed at roof. The system is to be installed facing due [XX] at a [XX] inclination.'],
  'ieq 1.1': [
    'The daylight DtS pathway has been applied to demonstrate daylight compliance.',
    'The BESS built in daylight calculator has been applied to demonstrate compliance. [XX]% of the living areas achieve the BESS best practice daylight requirements.',
    '[XX]% of the living areas achieves the BESS best practice requirements.',
  ],
  'ieq 1.2': [
    'The daylight DtS pathway has been applied to demonstrate daylight compliance.',
    'The BESS built in daylight calculator has been applied to demonstrate compliance. [XX]% of the bedrooms achieve the BESS best practice daylight requirements.',
    '[XX]% of the bedrooms achieves the BESS best practice requirements.',
  ],
  'ieq 1.3': ['>70% of dwellings receive at least 3 hours of direct sunlight in all Living areas between 9am and 3pm in mid-winter.'],
  'ieq 1.4': ['The commercial areas are targeting a 2% DF to [33]% of the nominated area. This is deemed achievable based on the current design.'],
  'ieq 1.5': ['[XX]% of the floor area of the main living areas achieves adequate daylight.'],
  'ieq 1.6': ['[XX]% of the floor area of the secondary habitable rooms achieves adequate daylight.'],
  'ieq 2.1': ['[XX]% of the dwellings is naturally cross-ventilated with windows on opposite or adjacent facades.'],
  'ieq 2.2': ['All dwellings are naturally cross-ventilated with windows on opposite or adjacent facades.'],
  'ieq 2.3': [
    '[60%/100%] of the commercial area is to be naturally ventilated (operable windows / doors on adjacent facades).',
    'Outdoor air rates are to be [50%/100%] increased over AS1668.12 min. requirements.',
    'CO2 concentrations: the ventilation systems are designed to achieve, monitor and maintain CO2 concentrations below [700/800ppm].',
  ],
  'ieq 3.1': ['Double glazing (or better) is used for all habitable room windows.'],
  'ieq 3.2': ['We recommend to add additional shading to the following areas:\n• [XX]\n• [XX]'],
  'ieq 3.4': ['We recommend to add additional shading to the following areas:\n• [XX]\n• [XX]'],
  'ieq 3.5': ['Ceiling fans are to be provided to [XX]% of the tenancies.'],
  'ieq 4.1': ['Low VOC and / or formaldehyde products are to be used internally.'],
  'transport 1.1': ['[XX] secure bicycle spaces for residents.'],
  'transport 1.2': ['[XX] bicycle spaces for residential visitors.'],
  'transport 1.3': ['The majority of the residential bicycle parking spaces are located at ground or entry level.'],
  'transport 1.4': ['[XX] bicycle spaces for employees.'],
  'transport 1.5': ['[XX] bicycle spaces for commercial visitors.'],
  'transport 1.6': ['[XX] showers, [XX] lockers and change facilities are to be provided within the EOT facilities.'],
  'transport 2.1': ['Min. 1 electric vehicle charging station is integrated into the development.'],
  'transport 2.2': ['A formal car sharing scheme is integrated into the development.'],
  'transport 2.3': ['[XX] motor bike spaces are provided within the development.', 'Min. [XX]% of vehicle parking spaces are designated and labelled for motorbikes/mopeds.'],
  'waste 1.1': ['>30% of the existing development is re-used.'],
  'waste 2.1': ['Organic waste will be provided within the dedicated bin area.'],
  'waste 2.2': ['Annotate separate general, recycling, glass and green waste bins at bin area.'],
  'urban ecology 1.1': ['Min. [XX]m2 of communal space (currently satisfied).'],
  'urban ecology 2.1': ['[XX]% of the site area is provided with vegetation.'],
  'urban ecology 2.2': ['The development is provided with a green roof area.'],
  'urban ecology 2.3': ['The development is provided with a green wall / façade.'],
  'urban ecology 2.4': ['External taps and floor wastes are to be provided to all balconies and courtyards.'],
  'urban ecology 3.2': ['[XX]m2 of food production area will be provided.'],
}

function findTemplates(creditId: string): string[] {
  const key = creditId.toLowerCase().trim()

  // 1 — direct match
  if (CREDIT_COMMENT_TEMPLATES[key]) return CREDIT_COMMENT_TEMPLATES[key]

  // 2 — normalise verbose category names to short codes, then try exact match
  const normalised = key
    .replace(/^integrated water management/, 'iwm')
    .replace(/^operational energy/, 'oe')
    .replace(/^indoor environment(?:al)? quality/, 'ieq')
    .replace(/^indoor environment(?:al)?/, 'ieq')
    .replace(/^waste\s*(?:&|and)\s*resource recovery/, 'waste')
  if (normalised !== key && CREDIT_COMMENT_TEMPLATES[normalised]) return CREDIT_COMMENT_TEMPLATES[normalised]

  // 3 — extract just "category prefix + X.X" in case the creditId includes
  //     the full credit name (e.g. "Urban Ecology 2.1 Vegetation" or
  //     "Waste & Resource Recovery 1.1 Construction Waste") or uses an
  //     abbreviation the AI chose (e.g. "UE 1.1", "WRR 1.1", "W 1.1")
  const numMatch = key.match(/(\d+\.\d+)/)
  if (numMatch) {
    const num = numMatch[1]
    const categoryPatterns: [RegExp, string][] = [
      [/^(management|mgmt|mgt)\b/, 'management'],
      [/^(iwm\b|integrated water)/, 'iwm'],
      [/^(oe\b|operational energy)/, 'oe'],
      [/^(ieq\b|indoor environment)/, 'ieq'],
      [/^transport/, 'transport'],
      [/^(w\b|wrr\b|waste)/, 'waste'],
      [/^(ue\b|urban ecology|urban)/, 'urban ecology'],
      [/^innovation/, 'innovation'],
    ]
    for (const [pattern, prefix] of categoryPatterns) {
      if (pattern.test(key)) {
        const candidate = `${prefix} ${num}`
        if (CREDIT_COMMENT_TEMPLATES[candidate]) return CREDIT_COMMENT_TEMPLATES[candidate]
        break
      }
    }
  }

  return []
}

/* ── Innovation initiatives (from ESD_Review_Guide Innovation Credits sheet) ── */

export const INNOVATION_INITIATIVES: { name: string; desc: string; pts: string }[] = [
  { name: 'ESD As-built verification', desc: 'An ESD professional will be engaged throughout the design and construction process to perform a minimum of 2 site inspections during construction to ensure suitable implementation of ESD initiatives.', pts: '0.9' },
  { name: 'Design for Disassembly Plan', desc: 'Prepare design for disassembly plan.', pts: '0.9' },
  { name: 'Placemaking / Flexi-space', desc: 'A flexi-space is provided to the development. The function of this space is to be determined by the residents through a place-making process.', pts: '1.8' },
  { name: 'Community Development Program', desc: 'A community development program will be introduced and funded for the first 12 months.', pts: '0.9' },
  { name: 'Battery Storage', desc: 'The battery system will store the generated renewable solar energy during the day for use during peak electricity demand periods in the evening.', pts: '0.9' },
  { name: 'Triple Glazing', desc: 'Application of triple glazed windows throughout the development.', pts: '0.9' },
  { name: 'Grey water recycling', desc: 'Collecting grey water from sinks, showers and laundry for reuse for landscape irrigation.', pts: '1.8' },
  { name: 'Low GWP / No Refrigerants HHW', desc: 'Use heat pump systems with low GWP refrigerants or CO2.', pts: '0.9' },
  { name: 'Low GWP / No Refrigerants HVAC', desc: 'Use HVAC systems with low GWP refrigerants or CO2.', pts: '0.9' },
  { name: 'Carbon Neutral / Low Carbon Concrete', desc: 'Specify supplementary cementitious materials (SCMs) like fly ash or slag cement, and emerging carbon-negative concrete alternatives.', pts: '0.9' },
  { name: 'Carbon Neutral Power Agreement – Base Build', desc: '10 year carbon neutral power agreement between developer, owners corporation and electrical retailer to provide GreenPower to communal areas.', pts: '0.9' },
  { name: 'Carbon Neutral Power Agreement – Apartments', desc: 'All occupants will be connected to the embedded network which will be provided with 100% GreenPower.', pts: '0.9' },
  { name: 'Construction Waste Reduction', desc: 'Provide a construction waste management plan that commits to diverting at least 90% of construction and demolition waste from landfill.', pts: '0.9' },
  { name: 'Building User Engagement', desc: 'Commit to installing a system capable of capturing data produced by energy and water meters and displaying consumption trends for building user awareness.', pts: '0.9' },
  { name: 'Building Integrated Solar PV', desc: 'Solar PV will be integrated into the façade.', pts: '0.9' },
  { name: 'Airtightness testing (sample apartments)', desc: 'Air tightness testing for 2–3 sample apartments. Indicative cost: $5–7.5k.', pts: '0.9' },
  { name: 'Airtightness testing (whole building)', desc: 'Air tightness testing for the whole of the development. Indicative cost: $10–20k.', pts: '1.8' },
  { name: 'IEQ sensors', desc: 'IEQ sensors in all apartments measuring VOC levels, humidity, PPM and temperature. Approximately $500 AUD per apartment.', pts: '0.9' },
  { name: 'LCA', desc: 'Life cycle assessment to determine and reduce the whole of life carbon of the development. Approximately $15–20k.', pts: '1.8' },
  { name: 'Embodied Carbon Assessment', desc: 'Embodied carbon assessment to determine and reduce the embodied carbon of the development. Approximately $10–15k.', pts: '1.8' },
  { name: 'Material Passport', desc: 'Based on the life cycle / embodied carbon assessment a material passport will be developed outlining all material quantities used within the building and recommendations on potential recycling and/or reuse at end of life.', pts: '0.9' },
  { name: 'Micro grid', desc: 'Introduce a micro-grid to the development.', pts: '0.9' },
  { name: 'Building application', desc: 'Building smartphone app to facilitate communication between facility management team and occupants and between occupants.', pts: '0.9' },
  { name: 'Centralised HRV/ERV system', desc: 'Centralised HRV/ERV system. Indicative cost: ~$5–10k per apartment.', pts: '0.9' },
  { name: 'ESD Display', desc: 'A digital noticeboard in the lobby or lift displaying PTV timetable/map, weather forecast, environmental reminders, real-time solar PV output and rainwater harvested.', pts: '0.9' },
  { name: 'Composter', desc: 'The composter will significantly reduce organic waste generated on-site, diverting it from landfills.', pts: '0.9' },
  { name: 'Social Innovation', desc: 'The development will seek to implement a community portal and activity schedule to enhance the social interactions of its occupants within the communal spaces.', pts: '0.9' },
  { name: 'Electric Bike Fleet and E-Bike Charging Stations', desc: 'A fleet of electric bikes with parking spaces accessible to residents and employees (offered as a hire scheme) plus E-bike charging stations with access to GPO (min. 20% of bicycle parking spaces).', pts: '0.9' },
  { name: 'Water leak detection system', desc: 'Introduction of a water leak detection system including sensors in plant rooms, pump rooms and other critical areas, and ultrasonic flow meters for the main water users and each level of the building.', pts: '0.9' },
  { name: 'GreenFactor Tool and achievement of a 0.55 score', desc: 'Green Factor tool assessment to be undertaken by the Landscape Architect to demonstrate the achievement of a 0.55 score.', pts: '0.9' },
  { name: 'Heat Island Effect', desc: 'Meet Green Star Buildings V1 Credit 19: at least 75% of the whole site area uses vegetation, green roofs, high-SRI roofing, shaded hardscaping or water bodies to reduce the heat island effect.', pts: '0.9' },
  { name: 'Zero waste strategy', desc: 'A zero waste operational strategy within the Waste Management Plan outlining State/Council waste ambitions, local reduce/reuse/recycle opportunities, waste minimisation strategies and guidance on waste storage sizing.', pts: '0.9' },
  { name: 'Solshare system', desc: 'A Solshare system will be installed to equally divide on-site generated solar energy between tenants. Minimum 2.5kW per apartment recommended.', pts: '0.9' },
  { name: 'Recycled/reused/repurposed materials', desc: 'Application of 5–10% (of total project cost) recycled and reused materials within the development, such as recycled bricks, timber and flooring.', pts: '0.9' },
  { name: 'Share Economy', desc: 'Develop and implement a materials exchange initiative that facilitates the sharing and recycling of furniture, tools and other resources among community members.', pts: '0.9' },
  { name: 'Formal Pre-plaster inspection', desc: 'Commitment to undertake a formal pre-plaster insulation inspection by an EEC-accredited Certified Insulation Installer or building surveyor.', pts: '0.9' },
  { name: 'Bicycle Repair Station', desc: 'Inclusion of secure, convenient and accessible equipment including repair stands, pumps and tools for fixing bicycles.', pts: '0.9' },
  { name: 'Smart Grid Integration', desc: 'Install vehicle-to-grid (V2G) charging stations and battery storage to support grid stabilisation.', pts: '0.9' },
  { name: 'Natural Insulation Materials', desc: 'Specify sheep\'s wool, cellulose or hemp-based insulation materials for low embodied carbon.', pts: '0.9' },
  { name: 'Geothermal Systems', desc: 'Utilise ground-source heat pumps connected to shared geothermal loops.', pts: '1.8' },
  { name: 'Demand Response Systems', desc: 'Implement automated systems that shift energy loads to optimise renewable generation and grid pricing.', pts: '0.9' },
]

/* ── Innovation credits that must never appear (removed at user request) ── */
export const BLOCKED_INNOVATION_NAMES = new Set([
  '6 star acoustic treatment',
  '8 star energy ratings',
  'achieve 150% storm score',
  'air permeability testing',
  'demolition and construction waste reduction',
  'esd verification during construction',
  'micro grid and zero carbon strategy: building app',
  'micro grid and zero carbon strategy: building management system',
  'micro grid and zero carbon strategy: smart home technology',
  'micro grid and zero carbon strategy: net zero commitment',
  'night purge',
  'passive house standard',
  'public realm improvements',
  'solar pv system',
])

/* ── GIW Comment Generation ── */

function stripAILeakage(raw: string): string {
  // Patterns that identify a line as AI preamble / reasoning
  const preambleWord = /^(i |i'll |i will |let me |here is |here's |based on |according to |sure[,!]|certainly[,!]|looking at|given the|from the |the project|here we|analyzing|in reviewing|i need to|to fill|i can see|the data show|i've chosen|i'll choose|i will choose|choosing option|the template|filling the|filled template|as requested|below is|the following)/i
  // Lines that are pure intro headers (end with colon, no bullet)
  const introHeader = (l: string) => /^[^•\n*\-\d].*:\s*$/.test(l.trim()) && !l.trim().startsWith('•')
  // "Option N:" selector line
  const optionSelector = /^option\s+\d+\s*:?\s*$/i

  const lines = raw.trim().split('\n')
  let start = 0
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim()
    if (!t) continue
    if (preambleWord.test(t) || introHeader(t) || optionSelector.test(t)) continue
    start = i
    break
  }
  let text = lines.slice(start).join('\n').trim()

  // Strip trailing AI commentary paragraph (after a blank line)
  text = text.replace(/\n\n(this (completes|fills|provides|covers|addresses)|i (have|'ve) (replaced|filled|completed|chosen)|note that|please note|the above|as you can see|all placeholders|all \[|the values above)[\s\S]*/i, '').trim()

  // Strip a second paragraph only when it looks like leaked reasoning:
  // first paragraph is a complete sentence AND second paragraph is not a template continuation
  const paraBreak = text.indexOf('\n\n')
  if (paraBreak !== -1) {
    const first = text.slice(0, paraBreak).trim()
    const second = text.slice(paraBreak + 2).trimStart()
    const firstComplete = /[.!?]$/.test(first)
    const secondIsContinuation = /^(option |•|-|\d+\.)/.test(second.toLowerCase())
    if (firstComplete && !secondIsContinuation) return first
  }
  return text
}

export async function generateGIWComments(projectId: string): Promise<void> {
  const credits = await prisma.credit.findMany({
    where: { projectId },
    orderBy: [{ categoryOrder: 'asc' }, { creditId: 'asc' }],
  })

  for (const credit of credits) {
    // Manually edited comments are preserved across BESS revisions — never overwrite them
    if (credit.lastEditedBy) continue

    // Innovation umbrella credits are split into line items by generateInnovationLineItems — skip here
    if (credit.category.toLowerCase().includes('innovation')) continue

    // OE 4.2 / OE 4.5: extract solar PV regardless of credit status.
    // If a PV system is entered it must appear in the GIW comment even when not achieved / scoped out.
    if (/^oe\s+4\.[25]$/i.test(credit.creditId.trim())) {
      const raw = credit.rawDataPoints ?? ''
      const kwMatch =
        raw.match(/system\s+size[^:\n]*:\s*([\d.]+)\s*k[wW]?[ph]?/i) ??
        raw.match(/solar\s+photovoltaic[^:\n]*\n[^\n]*system\s+size[^:\n]*:\s*([\d.]+)/i) ??
        raw.match(/total\s+(?:size|capacity)[^:\n]*:\s*([\d.]+)\s*k[wW]/i) ??
        raw.match(/([\d.]+)\s*k[wW][ph]?\s+solar/i) ??
        raw.match(/solar[^:\n]*:\s*([\d.]+)\s*k[wW]/i)
      const kw = kwMatch ? kwMatch[1] : null
      if (kw) {
        const orientMatch =
          raw.match(/orientation[^:\n]*:\s*([^\n,;.]+)/i) ??
          raw.match(/facing\s+(?:due\s+)?([A-Za-z/ -]+)/i)
        const tiltMatch =
          raw.match(/(?:tilt|inclination|angle)[^:\n]*:\s*([\d.]+)\s*°?/i) ??
          raw.match(/([\d.]+)\s*°\s*(?:tilt|inclination)/i)
        const orientation = orientMatch ? orientMatch[1].trim() : '[XX]'
        const tilt = tiltMatch ? `${tiltMatch[1]}°` : '[XX]'
        const comment = `A total ${kw}kW solar PV system is to be installed at roof. The system is to be installed facing due ${orientation} at a ${tilt} inclination.`
        await prisma.credit.update({ where: { id: credit.id }, data: { commentsGIW: comment } })
        continue
      }
      // No PV data found — fall through to "Not targeted." / scoped-out handling below
    }

    // Not achieved with zero score = not targeted; skip AI generation
    if (credit.creditStatus === 'N' && credit.creditScore === 0) {
      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: 'Not targeted.' },
      })
      continue
    }

    // OE 3.7 special case: show only the space types present in the BESS assessment
    if (/^oe\s+3\.7$/i.test(credit.creditId.trim()) && credit.creditStatus !== 'ScopedOut') {
      const allRaw = credits.map(c => c.rawDataPoints ?? '').join(' ')
      const hasRetail = /\bretail\b|\bshop\b/i.test(allRaw)
      const hasOffice = /\boffice\b/i.test(allRaw)
      const lines: string[] = ['Lighting power density shall be as follows:']
      if (hasRetail) lines.push('• Retail: No greater than average 14W/m2')
      if (hasOffice) lines.push('• Office: No greater than average 4.5W/m2')
      // Fall back to both if neither detected
      if (!hasRetail && !hasOffice) {
        lines.push('• Retail: No greater than average 14W/m2')
        lines.push('• Office: No greater than average 4.5W/m2')
      }
      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: lines.join('\n') },
      })
      continue
    }

    // OE 2.6 special case
    if (/^oe\s+2\.6$/i.test(credit.creditId.trim()) && credit.creditStatus === 'ScopedOut') {
      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: 'Not targeted.' },
      })
      continue
    }
    if (/^oe\s+2\.6$/i.test(credit.creditId.trim()) && credit.creditStatus === 'Y') {
      const oe11 = credits.find(c => /^oe\s+1\.1$/i.test(c.creditId.trim()))
      const raw = oe11?.rawDataPoints ?? ''

      // Detect office-only buildings (non-default NLA, mentions office, no retail/residential)
      const areaNumbers = [...raw.matchAll(/\b(\d{1,3}(?:,\d{3})*|\d+(?:\.\d+)?)\b(?!\s*%)/g)]
        .map(m => parseFloat(m[1].replace(/,/g, '')))
        .filter(n => !isNaN(n) && n >= 100 && n <= 999999)
      const hasNonDefaultNLA = areaNumbers.some(n => n !== 1000 && n !== 10000)
      const mentionsOffice = /\boffice\b/i.test(raw)
      const mentionsOtherUse = /\bretail\b|\bresidential\b|\bdwelling\b|\bapartment\b|\bshop\b/i.test(raw)
      const isOfficeOnly = hasNonDefaultNLA && mentionsOffice && !mentionsOtherUse

      // Detect retail component across all credits' rawDataPoints
      const allRaw = credits.map(c => c.rawDataPoints ?? '').join(' ')
      const hasRetail = /\bretail\b|\bshop\b/i.test(allRaw)
      const retailSuffix = hasRetail
        ? ' Please confirm if retail requires a gas cooktop? Note that this will have a significant negative impact on the BESS score.'
        : ''

      const base = isOfficeOnly
        ? 'The development will be all electric with no gas connection.'
        : 'The development is all electric with induction cooktops and no gas connection.'

      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: base + retailSuffix },
      })
      continue
    }

    // IWM 3.1 scoped-out: fixed annotation confirming rainwater tank connection
    if (/^iwm\s+3\.1$/i.test(credit.creditId.trim()) && credit.creditStatus === 'ScopedOut') {
      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: 'Landscape irrigation will be connected to the rainwater tank.' },
      })
      continue
    }

    // OE 3.1 scoped-out: fixed annotation for CO sensor carpark ventilation
    if (/^oe\s+3\.1$/i.test(credit.creditId.trim()) && credit.creditStatus === 'ScopedOut') {
      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: 'Carpark ventilation fans are to be controlled by CO sensors.' },
      })
      continue
    }

    // Management 3.2 / 3.3 scoped-out: single commercial tenancy
    if (/^management\s+3\.[23]$/i.test(credit.creditId.trim()) && credit.creditStatus === 'ScopedOut') {
      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: 'N/A - only one commercial tenancy.' },
      })
      continue
    }

    // OE 1.1: always mark as Achieved (mandatory credit, compliance assumed)
    if (/^oe\s+1\.1$/i.test(credit.creditId.trim()) && credit.creditStatus !== 'ScopedOut') {
      await prisma.credit.update({ where: { id: credit.id }, data: { creditStatus: 'Y' } })
    }

    // OE 1.2: mark as Achieved only if average NatHERS star rating ≥ 7
    if (/^oe\s+1\.2$/i.test(credit.creditId.trim()) && credit.creditStatus !== 'ScopedOut') {
      const raw = credit.rawDataPoints ?? ''
      const starsMatch =
        raw.match(/average[^:\n]*:\s*(\d+(?:\.\d+)?)\s*star/i) ??
        raw.match(/(\d+(?:\.\d+)?)\s*star[^s]*average/i) ??
        raw.match(/(\d+(?:\.\d+)?)\s*stars?\s*average/i) ??
        raw.match(/average\s*(?:star\s*)?rating[^:\n]*:\s*(\d+(?:\.\d+)?)/i)
      const avgStars = starsMatch ? parseFloat(starsMatch[1]) : null
      if (avgStars !== null && avgStars >= 7) {
        await prisma.credit.update({ where: { id: credit.id }, data: { creditStatus: 'Y' } })
      }
    }

    // IEQ 1.1 / IEQ 1.2: deterministically fill the daylight percentage
    // The template [XX]% means the % of living areas / bedrooms that comply, NOT the credit score.
    // BESS scores 66% when 80% of areas pass, 100% when all pass — Claude can't infer this.
    // Three pathways: DtS | Built-in calculator | Own calculations (modelling)
    if (/^ieq\s+1\.[12]$/i.test(credit.creditId.trim()) && credit.creditStatus !== 'ScopedOut') {
      const raw = credit.rawDataPoints ?? ''
      const isLiving = /^ieq\s+1\.1/i.test(credit.creditId.trim())
      const areaLabel = isLiving ? 'living areas' : 'bedrooms'

      // Pathway detection (mirrors report.ts logic)
      const usedDtS = /deemed\s*to\s*satisfy[^:\n]*\?:\s*yes\b/i.test(raw) || /\bdts\b|dts.*path/i.test(raw)
      const usedBuiltIn =
        /calculation\s+approach[^:\n]*\?:\s*use\s+the\s+built[- ]?in\s+calculation/i.test(raw) ||
        /approach[^:\n]*daylight[^:\n]*\?:\s*use\s+the\s+built[- ]?in/i.test(raw) ||
        /use\s+the\s+(?:bess\s+)?built[- ]?in\s+calculation/i.test(raw)
      const usedModelling =
        /calculation\s+approach[^:\n]*\?:\s*provide\s+(?:your|our)\s+own\s+calculations/i.test(raw) ||
        /approach[^:\n]*daylight[^:\n]*\?:\s*provide\s+(?:your|our)\s+own\s+calculations/i.test(raw) ||
        /provide\s+(?:your|our)\s+own\s+(?:daylight\s+)?calculations/i.test(raw) ||
        /daylight\s+modell(?:ing|ed)/i.test(raw) ||
        /own\s+(?:daylight\s+)?calculations?\s+(?:have\s+been|provided|submitted|uploaded)/i.test(raw) ||
        /(?:modell(?:ing|ed)|simulation)\s+(?:has\s+been\s+)?(?:provided|submitted|undertaken|used)/i.test(raw)

      // DtS pathway — fixed comment, no percentage needed
      if (usedDtS) {
        await prisma.credit.update({
          where: { id: credit.id },
          data: { commentsGIW: 'The daylight DtS pathway has been applied to demonstrate daylight compliance.' },
        })
        continue
      }

      // Extract compliance % from rawDataPoints
      const rawPctMatch = raw.match(/(\d{1,3}(?:\.\d+)?)\s*%[^,.\n]*(?:living|bedroom|habitable|area|space|comply|achieve)/i)
        ?? raw.match(/(?:living|bedroom|habitable|area|space)[^,.\n]*?(\d{1,3}(?:\.\d+)?)\s*%/i)
      let areaPct: string | null = rawPctMatch ? rawPctMatch[1] : null

      // Fall back to creditScore → area percentage mapping
      if (!areaPct && credit.creditScore != null) {
        const s = credit.creditScore
        if (s >= 95) areaPct = '100'
        else if (s >= 60) areaPct = '80'
        else if (s >= 30) areaPct = '40'
      }

      if (areaPct) {
        let comment: string
        if (usedBuiltIn && !usedModelling) {
          // Built-in calculator pathway explicitly detected
          comment = `The BESS built in daylight calculator has been applied to demonstrate compliance. ${areaPct}% of the ${areaLabel} achieve the BESS best practice daylight requirements.`
        } else {
          // Own calculations / daylight modelling pathway (default when no specific pathway detected, mirrors report.ts)
          comment = `${areaPct}% of the ${areaLabel} achieves the BESS best practice requirements.`
        }
        await prisma.credit.update({
          where: { id: credit.id },
          data: { commentsGIW: comment },
        })
        continue
      }
      // If we still can't determine the percentage, fall through to AI with an enriched prompt
    }

    // IEQ 1.5 / 1.6 scoped-out: hide from admin panel (leave comment null so soft-delete fires)
    if (/^ieq\s+1\.[56]$/i.test(credit.creditId.trim()) && credit.creditStatus === 'ScopedOut') continue

    // OE 4.x scoped-out: mark as not targeted
    const NOT_TARGETED_SCOPED = new Set(['oe 4.1', 'oe 4.2', 'oe 4.4', 'oe 4.5'])
    if (NOT_TARGETED_SCOPED.has(credit.creditId.trim().toLowerCase()) && credit.creditStatus === 'ScopedOut') {
      await prisma.credit.update({
        where: { id: credit.id },
        data: { commentsGIW: 'Not targeted.' },
      })
      continue
    }

    // Transport 1.1: extract residential bicycle count — show comment even if not achieved
    if (/^transport\s+1\.1$/i.test(credit.creditId.trim()) && credit.creditStatus !== 'ScopedOut') {
      const raw = credit.rawDataPoints ?? ''
      const countMatch =
        raw.match(/(\d+)\s*(?:secure\s+)?(?:resident(?:ial)?\s+)?bicycle\s+spaces?\s+(?:for\s+)?residents/i) ??
        raw.match(/resident(?:ial)?\s+bicycle[^:\n]*?:\s*(\d+)/i) ??
        raw.match(/(\d+)\s*long[- ]?stay\s+bicycle/i) ??
        raw.match(/how many[^?]*bicycle[^?]*\?[^:\n]*:\s*(\d+)/i) ??
        raw.match(/(\d+)\s*resident(?:ial)?\s+bicycle/i)
      const count = countMatch ? parseInt(countMatch[1], 10) : 0
      if (count > 0) {
        await prisma.credit.update({
          where: { id: credit.id },
          data: { commentsGIW: `${count} secure bicycle spaces for residents.` },
        })
        continue
      }
    }

    // Transport 1.2: extract residential visitor bicycle count — show comment even if not achieved
    if (/^transport\s+1\.2$/i.test(credit.creditId.trim()) && credit.creditStatus !== 'ScopedOut') {
      const raw = credit.rawDataPoints ?? ''
      const countMatch =
        raw.match(/(\d+)\s*residential\s+visitor\s+bicycle/i) ??
        raw.match(/(\d+)\s*visitor\s+bicycle[^.\n]*?resident/i) ??
        raw.match(/residential\s+visitor\s+bicycle[^:\n]*?:\s*(\d+)/i) ??
        raw.match(/how many[^?]*visitor[^?]*bicycle[^?]*\?[^:\n]*:\s*(\d+)/i) ??
        raw.match(/(\d+)\s*short[- ]?stay\s+bicycle/i)
      const count = countMatch ? parseInt(countMatch[1], 10) : 0
      if (count > 0) {
        await prisma.credit.update({
          where: { id: credit.id },
          data: { commentsGIW: `${count} bicycle spaces for residential visitors.` },
        })
        continue
      }
    }

    const templates = findTemplates(credit.creditId)

    if (templates.length > 0) {
      // Check if any template has placeholders that need filling
      const hasPlaceholders = templates.some(t => /\[[^\]]{1,80}\]/.test(t))

      if (!hasPlaceholders && templates.length === 1) {
        // No placeholders, single template — write verbatim, no AI needed
        await prisma.credit.update({
          where: { id: credit.id },
          data: { commentsGIW: templates[0] },
        })
        continue
      }

      // Template has placeholders or multiple options — use AI to fill/choose
      const templateList = templates.map((t, i) =>
        templates.length > 1 ? `Option ${i + 1}:\n${t}` : t,
      ).join('\n\n')

      const systemPrompt =
        'You are a text completion function. Output ONLY the completed template — nothing else. Zero preamble, zero explanation, zero reasoning, zero trailing notes. Start your response with the exact first character of the chosen template. Never write "I", "Based on", "Looking at", "Note", "Option", or any commentary. If choosing between options, output the chosen option text only, not its label. Always use "retail" instead of "shop".'

      const userPrompt = `Project data:\n${credit.rawDataPoints ?? 'No specific data recorded.'}\nStatus: ${credit.creditStatus}   Score: ${credit.creditScore != null ? `${credit.creditScore}%` : 'N/A'}\n\n${templates.length > 1 ? `Choose the single most appropriate option and fill its [placeholders] with values from the project data. Output the chosen option text only:\n\n${templateList}` : `Fill every [placeholder] in this template with the matching value from the project data. If a specific value is not in the data, leave the placeholder as-is:\n\n${templateList}`}`

      try {
        const message = await anthropic.messages.create({
          model: 'claude-sonnet-4-6',
          max_tokens: 512,
          system: systemPrompt,
          messages: [{ role: 'user', content: userPrompt }],
        })

        const block = message.content[0]
        if (block.type === 'text') {
          const text = stripAILeakage(block.text)
            .replace(/\[[^\]]*[A-Za-z][^\]]*\]/g, '')  // remove letter-containing unfilled placeholders [XX], [value]
            .replace(/\[\s*[\d.,°%/\s-]*\s*\]/g, '')   // remove numeric/symbol-only brackets [10], [15°]
            .replace(/\[\s*\]/g, '')                    // remove any remaining empty brackets []
            .replace(/[ \t]{2,}/g, ' ')
            .replace(/\n{3,}/g, '\n\n')
            .trim()
          await prisma.credit.update({
            where: { id: credit.id },
            data: { commentsGIW: text },
          })
        }
      } catch (err) {
        console.error(`[generate] Comment failed for credit ${credit.creditId}:`, err)
      }
    } else {
      // No template: write a short factual comment, max 150 characters
      const systemPrompt =
        'You are writing a concise factual note for a BESS building assessment. Output only the comment text — no preamble, no reasoning, no first-person, no "Based on", no "Looking at". Start directly with the fact. Maximum 150 characters. Always use "retail" instead of "shop".'

      const userPrompt = `Write a factual one-line comment for this BESS credit. Start with the key fact from the project data.\nStatus: ${credit.creditStatus}   Score: ${credit.creditScore != null ? `${credit.creditScore}%` : 'N/A'}\nProject data: ${credit.rawDataPoints ?? 'No specific data recorded.'}\nIf ScopedOut, briefly note what is outside scope. Max 150 characters.`

      try {
        const message = await anthropic.messages.create({
          model: 'claude-sonnet-4-6',
          max_tokens: 80,
          system: systemPrompt,
          messages: [{ role: 'user', content: userPrompt }],
        })

        const block = message.content[0]
        if (block.type === 'text') {
          const text = stripAILeakage(block.text)
            .replace(/[ \t]{2,}/g, ' ')
            .trim()
            .slice(0, 150)
          await prisma.credit.update({
            where: { id: credit.id },
            data: { commentsGIW: text },
          })
        }
      } catch (err) {
        console.error(`[generate] Comment failed for credit ${credit.creditId}:`, err)
      }
    }
  }

}

/* ── ESD Excellence Opportunity Generation ── */

export async function generateExcellenceOpportunities(projectId: string): Promise<void> {
  const project = await prisma.project.findUniqueOrThrow({
    where: { id: projectId },
    select: { id: true, name: true, address: true },
  })

  const credits = await prisma.credit.findMany({
    where: { projectId },
    orderBy: [{ categoryOrder: 'asc' }, { creditId: 'asc' }],
  })

  // Clear existing opportunities
  await prisma.eSDExcellenceOpportunity.deleteMany({ where: { projectId } })

  const EXCELLENCE_EXCLUDED = ['oe 4.4', 'ieq 1.1', 'ieq 1.2', 'ieq 1.3', 'innovation']

  const eligible = credits.filter((c) => {
    const id = c.creditId.toLowerCase().trim()
    if (c.creditStatus === 'ScopedOut') return false
    if (c.creditScore != null && c.creditScore >= 100) return false
    if (c.creditScore === 0) return false
    if (EXCELLENCE_EXCLUDED.some((ex) => id === ex || id.startsWith(ex + ' ') || c.category.toLowerCase().includes('innovation'))) return false
    return true
  })

  const EXCELLENCE_FIXED_SUFFIX: Record<string, string> = {
    'ieq 2.1':            'This can be achieved through the introduction of mechanically assisted natural ventilation to non-cross ventilated apartments.',
    'urban ecology 2.2':  'Significant planters at roof or large terraces can be claimed as green roofs.',
    'urban ecology 2.3':  'This can be creepers, green wall systems or hanging plants cascading down.',
  }

  for (const credit of eligible) {
    const userPrompt = `Generate a short ESD Excellence Opportunity description for this credit.

Project: ${project.name}${project.address ? `, ${project.address}` : ''}
Credit: ${credit.creditId} ${credit.creditName}
Current score: ${credit.creditScore != null ? `${credit.creditScore}%` : 'N/A'}  Max: ${credit.creditId.toLowerCase().trim() === 'ieq 1.4' ? '60%' : '100%'}
Status: ${credit.creditStatus}
Requirement: ${credit.creditRequirement ?? 'N/A'}
Project data:
${credit.rawDataPoints ?? 'No specific data recorded.'}

Score thresholds (include only if relevant):
- OE 1.1 NatHERS Houses/Townhouses: 12.5% credit score at 0% improvement, 37.5% at 10% improvement, 50% at 20% improvement, 100% at 60% improvement
- OE 1.2 NatHERS Apartments: 50% credit score at 7.5 stars average, 75% at 8.0 stars, 100% at 8.5 stars average
- OE 2.1 GHG emissions: continuous — score increases per % improvement, max at 20% reduction
- OE 2.7 Energy consumption: continuous — score increases per % improvement, max at 20% reduction
- OE 4.2 Solar Apartments: minimum 5% of apartment energy met by solar
- OE 4.5 Solar Townhouses: min 30% for any points, 100% for maximum
- IEQ 1.1 Daylight Living: 66% credit score when 80% of areas pass, 100% when 100% pass
- IEQ 1.2 Daylight Bedrooms: 66% credit score when 80% of bedrooms pass, 100% when 100% pass
- IEQ 2.1 Natural Ventilation: 66% credit score when 60% of apartments pass, 100% when 100% pass. This can be achieved through the introduction of mechanically assisted natural ventilation to non-cross ventilated apartments.
- IEQ 3.2 Thermal Comfort: continuous — score increases proportionally from 33%
- IEQ 3.4 Noise: 66% credit score when 50% of areas achieve target, 100% when 100% achieve target
- IEQ 1.4 Daylight Commercial: maximum achievable score is 60% — do not suggest exceeding this
- IWM 1.1 Potable Water: continuous — score increases proportionally with % reduction. Pathways to improve: increase WELS star ratings on fixtures, connect toilets to rainwater/recycled water (if not already), connect landscape irrigation to rainwater/recycled water (if not already)
- Urban Ecology 2.1: 25% credit score at >5% site area, 50% at >10%, 75% at >20%, 100% at >30%
- OE 3.4 Clothes Drying: suggest individual clothes lines (per dwelling) or communal clothes lines, only if not already targeted in the project data
- Transport 2.1 EV Charging: minimum 1 EV charging station must be installed at practical completion to claim full points
- Urban Ecology 2.2 Green Roof: Significant planters at roof or large terraces can be claimed as green roofs.
- Urban Ecology 2.3 Green Wall: This can be creepers, green wall systems or hanging plants cascading down.
- Innovation 1.1: 1 point per confirmed initiative, 10 points max

${/^iwm\s*1\.1$/i.test(credit.creditId.trim())
  ? `Write 1–3 sentences describing how to improve this credit. Check the project data above — for each pathway not already in use (higher WELS ratings, toilet connection to rainwater/recycled water, landscape irrigation connection to rainwater/recycled water), mention it specifically. Do not mention pathways already achieved. Do not mention the credit weight.`
  : `Write exactly 1 sentence describing the specific action needed to improve this credit, with the exact threshold or target value. Do not mention the current score. Do not mention the credit weight. Do not add extra sentences. Keep the response under 150 characters.`}`

    try {
      const message = await anthropic.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: /^iwm\s*1\.1$/i.test(credit.creditId.trim()) ? 300 : 150,
        system:
          'You are an ESD consultant writing concise improvement notes for a BESS assessment. Be specific — cite actual numbers and thresholds from the project data. No generic advice. No AI references. One sentence only, under 150 characters: describe the specific improvement needed with an exact number or target. Never mention credit weights. Always use "retail" instead of "shop".',
        messages: [{ role: 'user', content: userPrompt }],
      })

      const block = message.content[0]
      if (block.type === 'text') {
        const suffix = EXCELLENCE_FIXED_SUFFIX[credit.creditId.toLowerCase().trim()]
        const cleaned = stripAILeakage(block.text)
        const description = suffix ? `${cleaned} ${suffix}` : cleaned
        await prisma.eSDExcellenceOpportunity.create({
          data: {
            projectId,
            creditId: credit.id,
            creditReference: credit.creditId,
            creditName: credit.creditName,
            currentScore: credit.creditScore,
            improvementDescription: description,
          },
        })
      }
    } catch (err) {
      console.error(
        `[generate] Excellence opportunity failed for credit ${credit.creditId}:`,
        err,
      )
    }
  }

  // Add all innovation initiatives as fixed cards
  for (const initiative of INNOVATION_INITIATIVES) {
    await prisma.eSDExcellenceOpportunity.create({
      data: {
        projectId,
        creditReference: 'Innovation',
        creditName: initiative.name,
        improvementDescription: initiative.desc || null,
        bessPoints: initiative.pts,
      },
    })
  }
}

/* ── Trigger wrappers (called from create-from-pdf) ── */

/* ── Innovation credit line items ── */

export async function generateInnovationLineItems(projectId: string): Promise<void> {
  // Restore umbrella credits hidden by a previous generation run
  await prisma.credit.updateMany({
    where: {
      projectId,
      category: { contains: 'innovation', mode: 'insensitive' },
      creditId: { not: 'Innovation' },
      hiddenFromPortal: true,
      rawDataPoints: { not: null },
    },
    data: { hiddenFromPortal: false },
  })

  // Remove line items from a previous generation run
  await prisma.credit.deleteMany({
    where: { projectId, creditId: 'Innovation' },
  })

  // Find Innovation umbrella credits that have raw data
  const umbrellaCredits = await prisma.credit.findMany({
    where: {
      projectId,
      category: { contains: 'innovation', mode: 'insensitive' },
      creditId: { not: 'Innovation' },
      creditStatus: { not: 'ScopedOut' },
      deletedByGIW: false,
      rawDataPoints: { not: null },
    },
  })

  for (const credit of umbrellaCredits) {
    if (!credit.rawDataPoints?.trim()) continue

    // Extract the list of claimed initiative names and descriptions from the raw BESS data
    let items: { name: string; desc: string | null }[] = []
    try {
      const msg = await anthropic.messages.create({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 800,
        system: 'Extract innovation initiative names and descriptions from BESS assessment data. Return ONLY a valid JSON array of objects with "name" and "desc" string fields — no other text. If no description is available for an initiative, set "desc" to null.',
        messages: [{
          role: 'user',
          content: `List the innovation initiatives claimed in this BESS data as a JSON array of {name, desc} objects:\n\n${credit.rawDataPoints}`,
        }],
      })
      const raw = msg.content[0]?.type === 'text' ? msg.content[0].text.trim() : '[]'
      const match = raw.match(/\[[\s\S]*\]/)
      items = match ? (JSON.parse(match[0]) as { name: string; desc: string | null }[]) : []
    } catch (err) {
      console.error('[generate] Innovation extraction failed:', err)
      continue
    }

    if (items.length === 0) continue

    for (const item of items) {
      const trimmed = String(item.name).trim()
      if (!trimmed) continue

      // Skip permanently blocked initiatives
      const lower = trimmed.toLowerCase()
      if (BLOCKED_INNOVATION_NAMES.has(lower)) continue

      // Match against INNOVATION_INITIATIVES for description and points; fall back to extracted desc
      const initiative = INNOVATION_INITIATIVES.find(
        (i) =>
          i.name.toLowerCase() === lower ||
          i.name.toLowerCase().includes(lower) ||
          lower.includes(i.name.toLowerCase()),
      )

      const pts = initiative?.pts ?? '0.9'
      const ptsNum = parseFloat(pts) || 0.9
      const desc = initiative?.desc ?? item.desc ?? null

      await prisma.credit.create({
        data: {
          projectId,
          category: credit.category,
          categoryOrder: credit.categoryOrder,
          creditId: 'Innovation',
          creditName: trimmed,
          commentsGIW: desc,
          creditScore: 100,
          creditWeight: 90,
          creditStatus: 'Y',
          mandatory: false,
        },
      })
    }

    // Hide the umbrella from the matrix so only line items appear, but keep it for BESS calculation
    await prisma.credit.update({
      where: { id: credit.id },
      data: { hiddenFromPortal: true },
    })
  }
}

/* ── Auto visibility rules ── */

export async function applyAutoVisibilityRules(projectId: string): Promise<void> {
  // Fixed-comment scoped-out credits: restore if previously deleted with no comment
  const fixedScopedComments: Array<{ pattern: RegExp; comment: string }> = [
    { pattern: /^iwm\s*3\.1$/i,          comment: 'Landscape irrigation will be connected to the rainwater tank.' },
    { pattern: /^oe\s*3\.1$/i,           comment: 'Carpark ventilation fans are to be controlled by CO sensors.' },
    { pattern: /^oe\s*4\.[124]$/i,        comment: 'Not targeted.' },
    { pattern: /^oe\s*4\.5$/i,           comment: 'Not targeted.' },
    { pattern: /^management\s*3\.[23]$/i, comment: 'N/A - only one commercial tenancy.' },
  ]
  const scopedCredits = await prisma.credit.findMany({
    where: { projectId, creditStatus: 'ScopedOut' },
    select: { id: true, creditId: true, commentsGIW: true },
  })
  for (const sc of scopedCredits) {
    const rule = fixedScopedComments.find(r => r.pattern.test(sc.creditId.trim()))
    if (rule && (!sc.commentsGIW || sc.commentsGIW === '')) {
      await prisma.credit.update({
        where: { id: sc.id },
        data: { deletedByGIW: false, commentsGIW: rule.comment },
      })
    }
  }

  // Soft-delete scoped-out credits that have no GIW comment
  await prisma.credit.updateMany({
    where: {
      projectId,
      creditStatus: 'ScopedOut',
      deletedByGIW: false,
      OR: [{ commentsGIW: null }, { commentsGIW: '' }],
    },
    data: { deletedByGIW: true },
  })

  // Hide not-achieved credits from the reviewer portal
  await prisma.credit.updateMany({
    where: { projectId, creditStatus: 'N', deletedByGIW: false },
    data: { hiddenFromPortal: true },
  })

  // Hide scoped-out credits from the reviewer portal, except OE 3.1
  await prisma.credit.updateMany({
    where: {
      projectId,
      creditStatus: 'ScopedOut',
      deletedByGIW: false,
      NOT: { creditId: { equals: 'OE 3.1', mode: 'insensitive' } },
    },
    data: { hiddenFromPortal: true },
  })

  // Ensure OE 3.1 scoped-out is visible in the reviewer portal
  await prisma.credit.updateMany({
    where: { projectId, creditStatus: 'ScopedOut', creditId: { equals: 'OE 3.1', mode: 'insensitive' } },
    data: { hiddenFromPortal: false },
  })

  // Un-hide credits that are now achieved (not N or ScopedOut)
  // Exclude innovation umbrella credits — intentionally hidden by generateInnovationLineItems
  await prisma.credit.updateMany({
    where: {
      projectId,
      creditStatus: { notIn: ['N', 'ScopedOut'] },
      hiddenFromPortal: true,
      NOT: { category: { contains: 'innovation', mode: 'insensitive' } },
    },
    data: { hiddenFromPortal: false },
  })

  // Transport 1.1 and 1.2: always show in the reviewer portal when bicycle counts > 0,
  // even if the credit is not achieved (the comment was filled above only when count > 0)
  await prisma.credit.updateMany({
    where: {
      projectId,
      deletedByGIW: false,
      OR: [
        { creditId: { equals: 'Transport 1.1', mode: 'insensitive' } },
        { creditId: { equals: 'Transport 1.2', mode: 'insensitive' } },
      ],
      commentsGIW: { not: null },
    },
    data: { hiddenFromPortal: false },
  })
}

export async function triggerCommentGeneration(projectId: string): Promise<void> {
  try {
    await generateGIWComments(projectId)
    await generateInnovationLineItems(projectId)
    await generateExcellenceOpportunities(projectId)
    await applyAutoVisibilityRules(projectId)
    await prisma.project.update({
      where: { id: projectId },
      data: { generationStatus: 'complete' },
    })
  } catch (err) {
    console.error(`[generate] Generation failed for project ${projectId}:`, err)
    await prisma.project
      .update({ where: { id: projectId }, data: { generationStatus: 'error' } })
      .catch(console.error)
  }
}

export async function triggerDrawingGeneration(projectId: string): Promise<void> {
  console.log(`[generate] Drawing requirement generation pending for project ${projectId}`)
}
