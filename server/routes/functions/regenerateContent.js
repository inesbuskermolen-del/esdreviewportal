import Anthropic from '@anthropic-ai/sdk';
import { supabase } from '../../config/supabase.js';
import { attachUser } from '../../middleware/auth.js';

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const CAT_ABBREV = {
  'Management': 'Mgmt',
  'Integrated Water Management': 'IWM',
  'Operational Energy': 'OE',
  'Indoor Environment Quality': 'IEQ',
  'Transport': 'Transport',
  'Waste & Resource Recovery': 'Waste',
  'Urban Ecology': 'UE',
  'Innovation': 'Innovation',
};

function buildShortRef(category, creditRef) {
  const abbrev = CAT_ABBREV[category] || (category || '').split(' ').map(w => w[0]).join('').toUpperCase();
  return `${abbrev} ${creditRef || ''}`.trim();
}

function parseCreditRef(creditName) {
  const m = creditName?.match(/^([\d.]+[a-z]?)\s*[—\-]/i);
  return m ? m[1] : '';
}

const CREDIT_TEMPLATES = {
  'Management': {
    '1.1': 'GIW is to attend a pre-application meeting with Council DTP. Town planner to coordinate.',
    '2.1': 'A preliminary FirstRate assessment will be undertaken for the residence.',
    '2.2': 'A sample of preliminary FirstRate assessments will be undertaken for the development.',
    '2.3': 'Preliminary J1V3 modelling will be undertaken for the development.',
    '3.1': 'Utility meters (electricity and water) will be provided for all individual dwellings.',
    '3.2': 'Utility meters (electricity and water) will be provided for all individual tenancies.',
    '3.3': 'Utility sub-meters will be provided for all major common area services.',
    '4.1': 'A building user guide will be produced and issued to occupants.',
  },
  'Integrated Water Management': {
    '1.1': '• Showerheads: 4 Star WELS (>6 but <7.5 l/min)\n• Kitchen Taps: 5 Star WELS\n• Bathroom Taps: 5 Star WELS\n• Dishwashers: 5 Star WELS\n• WC: 4 Star WELS\n• Urinals: Scope out\n• Washing Machine: Occupant to Install',
    '2.1': 'A compliant Blue Factor result can be achieved via the following options.\n\nOption 1:\n• Rainwater collection off [XX] is to be directed into a [XX]-litre rainwater tank connected to [XX] toilets and landscape irrigation.\n• Rainwater collection off [XX] is to be directed into a ≥[XX]m2 raingarden with 100mm of extended detention.',
    '3.1': 'Landscaping is either native vegetation with no water demand after the initial establishment period OR landscape irrigation is to be connected to the rainwater tank. Please confirm preference?',
    '4.1': '80% of fire system test water (e.g. hydrant pump test water or SCV annubar test) is to be reused on-site, either within the fire system or directed into the rainwater tank OR the fire test water system does not expel water.',
  },
  'Operational Energy': {
    '1.1': 'GIW has undertaken a preliminary facade assessment in accordance with NCC2022 Section J4D6 and recommends the application of the J1V3 pathway for Section J compliance. This will be undertaken during the DD stage.',
    '1.2': 'The energy ratings are to achieve a [7] Star average with no unit below 6 Stars and no unit exceeding the maximum allowed cooling loads as outlined under BADS.',
    '2.6': 'The development is all electric with induction cooktops and no gas connection.',
    '2.7': '• Apartment HVAC systems are to be within one star of the best available unit of the same capacity.\n• Commercial HVAC systems are to have a COP of 3.4.',
    '3.1': 'Carpark ventilation fans are to be controlled by CO sensors.',
    '3.2': 'Centralised heat pump hot water system or individual electric instantaneous hot water systems. Please confirm?',
    '3.3': 'Operation of min. 50% of external lighting is controlled by a motion detector.',
    '3.4': '[Individual / shared] clothes drying lines are to be introduced to the development.',
    '3.5': 'Maximum illumination power density for each dwelling is 4W/sqm or less.',
    '3.6': 'Maximum illumination power density for each dwelling is 4W/sqm or less.',
    '3.7': 'Maximum illumination power density (W/sqm) of the relevant building class is at least 20% lower than current NCC requirements.',
    '4.1': 'Combined heat and power system is to be introduced.',
    '4.2': 'A [XX]kW solar PV system is to be installed on the roof, facing due [XX] at a [XX]° inclination.',
    '4.4': 'A geothermal system is to be introduced to the development.',
    '4.5': 'A total [XX]kW solar PV system is to be installed on the roof, facing due [XX] at a [XX]° inclination.',
  },
  'Indoor Environment Quality': {
    '1.1': '[XX]% of living areas achieves the BESS best practice daylight requirements.',
    '1.2': '[XX]% of bedrooms achieves the BESS best practice daylight requirements.',
    '1.3': '>70% of dwellings receive at least 3 hours of direct sunlight in all living areas between 9am and 3pm in mid-winter.',
    '1.4': 'The commercial areas are targeting a 2% DF to [33]% of the nominated area. This is deemed achievable based on the current design.',
    '1.5': '[XX]% of the floor area of the main living areas achieves adequate daylight.',
    '1.6': '[XX]% of the floor area of the secondary habitable rooms achieves adequate daylight.',
    '2.1': '[XX]% of the dwellings are naturally cross-ventilated with windows on opposite or adjacent facades.',
    '2.2': 'All dwellings are naturally cross-ventilated with windows on opposite or adjacent facades.',
    '2.3': '[60%/100%] of the commercial area is to be naturally ventilated (operable windows/doors on adjacent facades).',
    '3.1': 'Double glazing (or better) is used for all habitable room windows.',
    '3.2': 'Additional shading is recommended to the following areas:\n• [XX]\n• [XX]',
    '3.4': 'Additional shading is recommended to the following areas:\n• [XX]\n• [XX]',
    '3.5': 'Ceiling fans are to be provided to [XX]% of the tenancies.',
    '4.1': 'Low VOC products are to be used internally.',
  },
  'Transport': {
    '1.1': '[XX] secure bicycle spaces are provided for residents.',
    '1.2': '[XX] bicycle spaces are provided for residential visitors.',
    '1.3': 'The majority of the residential bicycle parking spaces are located at ground or entry level.',
    '1.4': '[XX] bicycle spaces are provided for employees.',
    '1.5': '[XX] bicycle spaces are provided for commercial visitors.',
    '1.6': '[XX] showers, [XX] lockers and change facilities are to be provided within the EOT facilities.',
    '2.1': 'Min. 1 electric vehicle charging station is integrated into the development.',
    '2.2': 'A formal car sharing scheme is integrated into the development.',
    '2.3': '[XX] motorbike spaces are provided within the development.',
  },
  'Waste & Resource Recovery': {
    '1.1': '>30% of the existing development is re-used.',
    '2.1': 'Organic waste facilities will be provided within the bin room.',
    '2.2': 'Separate general, recycling, glass and green waste bins are to be annotated at the bin area.',
  },
  'Urban Ecology': {
    '1.1': 'Min. [XX]m2 of communal space is provided.',
    '2.1': '[XX]% of the site area is provided with vegetation.',
    '2.2': 'The development is provided with a green roof area.',
    '2.3': 'The development is provided with a green wall/façade.',
    '3.2': '[XX]m2 of food production area will be provided.',
  },
};

function getCategoryTemplate(category, creditRef) {
  const ref = (creditRef || '').replace(/[a-z]$/i, '').trim();
  return CREDIT_TEMPLATES[category]?.[ref] || null;
}

function getScoreTierContext(category, creditRef) {
  const ref = (creditRef || '').replace(/[a-z]$/i, '').trim();
  const cat = (category || '').toLowerCase();
  if (cat.includes('energy')) {
    if (ref === '1.2') return 'NatHERS score tiers: 50% at weighted average 7.5 stars, higher at 8.0 and 8.5 stars.';
    if (ref === '2.1') return 'Points increase for each % GHG improvement; maximum score at 20% reduction.';
    if (ref === '2.7') return 'Points increase for each % energy consumption improvement; maximum score at 20% reduction.';
    if (ref === '4.2') return 'Score based on % of apartment building energy met by solar; minimum threshold for any points is 5%.';
    if (ref === '4.5') return 'Score based on % of townhouse energy met by solar; 30% minimum for any points, 100% for maximum score.';
  }
  if (cat.includes('indoor') || cat.includes('ieq')) {
    if (ref === '1.1') return '66% score at 80% of living areas achieving daylight criteria; 100% score at 100% of living areas.';
    if (ref === '1.2') return '66% score at 80% of bedrooms achieving daylight criteria; 100% score at 100% of bedrooms.';
    if (ref === '2.1') return 'Score bands at >60%, >75%, >90% of apartments naturally ventilated.';
  }
  if (cat.includes('urban') || cat.includes('ecology')) {
    if (ref === '2.1') return 'Score tiers at >5%, >10%, >20%, >30% of site area vegetated.';
  }
  if (cat.includes('water')) {
    if (ref === '1.1') return 'Points awarded above 25% potable water reduction; higher bands at 30%, 40%, 50%+ reduction.';
  }
  if (cat.includes('innovation')) {
    if (ref === '1.1') return 'Points per confirmed initiative; maximum 10 points total.';
  }
  return '';
}

const GIW_COMMENT_SYSTEM = `You are an ESD consultant writing internal project review notes for a BESS (Built Environment Sustainability Scorecard) assessment. Your notes will be read by the design team. Write in plain, direct professional language. Every sentence must reference specific numbers, ratings, or facts from the project data provided. Do not write generic statements that could apply to any project. Do not mention AI, automation, or report generation. Do not use phrases like 'the project achieves' as an opener for every sentence — vary the sentence structure. Write in third person. Length should be 3–6 sentences.`;

const OPP_SYSTEM = `You are an ESD consultant identifying specific design improvement opportunities for a BESS assessment. Write in plain, direct professional language. Every recommendation must be based on the specific project data provided — cite actual numbers, ratings, thresholds and targets from the BESS data. Do not write generic advice that could apply to any project. Do not mention AI, automation, or report generation. Reference the BESS credit ID at the start of the description. Be specific about what needs to change and by how much.`;

// POST /api/functions/regenerateContent
export default [attachUser, async (req, res) => {
  try {
    const { projectId } = req.body;
    if (!projectId) return res.status(400).json({ error: 'Missing projectId' });

    const [{ data: projectsArr }, { data: credits }] = await Promise.all([
      supabase.from('projects').select('*').eq('id', projectId),
      supabase.from('credits').select('*').eq('project_id', projectId),
    ]);
    const project = projectsArr?.[0];
    if (!project) return res.status(404).json({ error: 'Project not found' });

    let commentsGenerated = 0, oppsGenerated = 0;

    // 1. Regenerate GIW Comments
    for (const credit of (credits || [])) {
      try {
        const creditRef = parseCreditRef(credit.credit_name);
        const shortRef = buildShortRef(credit.category, creditRef);
        const categoryTemplate = getCategoryTemplate(credit.category, creditRef);

        const statusDesc =
          credit.credit_status === 'Y' ? `Achieved at ${credit.credit_score}%` :
          credit.credit_status === 'ScopedOut' ? `Scoped Out — ${credit.scoped_out_reason || 'see BESS report'}` :
          `Not Achieved (score: ${credit.credit_score ?? 0}%)`;

        const userPrompt = categoryTemplate
          ? `Generate a Comments GIW entry for the following BESS credit for this project.

Project: ${project.name || ''}, ${project.address || ''}
Credit: ${shortRef} ${credit.credit_name}
Credit status: ${credit.credit_status}
Credit requirement: ${credit.credit_requirement || ''}
${credit.scoped_out_reason ? `Scoped-out reason: ${credit.scoped_out_reason}` : ''}

Use the following template as your answer:
"${categoryTemplate}"

Rules:
- If status is Y (achieved): state the template text as a confirmed fact, filling in any specific numbers from the credit requirement
- If status is N (not achieved): rewrite the template to state what is not provided or missing
- If status is ScopedOut: state the scoped-out reason
- Replace any [X] placeholders with the actual numbers extracted from the credit requirement
- Keep to 1–2 sentences only. Do not add analytical content, score gaps, or extra context.`
          : `Generate a Comments GIW entry for the following BESS credit for this project.

Project: ${project.name || ''}, ${project.address || ''}
Credit: ${shortRef} ${credit.credit_name}
Credit score: ${credit.credit_score ?? 0}%
Credit status: ${credit.credit_status}
Credit requirement: ${credit.credit_requirement || ''}
${credit.scoped_out_reason ? `Scoped-out reason: ${credit.scoped_out_reason}` : ''}

Write a 3–6 sentence comment that:
- States what this credit requires (briefly, one clause only)
- States what the project has achieved or not achieved, citing any specific numbers from the credit requirement above
- If score is 100%: ends with a confirmation the credit is fully met
- If score is between 1% and 99%: states the current achievement level and what the numerical gap is to the next score threshold, without suggesting how to fix it
- If score is 0% and status is N: states what is missing relative to the credit requirement
- If status is ScopedOut: states the specific reason this credit was scoped out, as extracted from the BESS data`;

        const message = await anthropic.messages.create({
          model: 'claude-sonnet-4-6',
          max_tokens: 1024,
          system: GIW_COMMENT_SYSTEM,
          messages: [{ role: 'user', content: userPrompt }],
        });

        const generated = message.content[0]?.text || '';
        await supabase.from('credits').update({
          comments_giw: generated,
          last_edited_by: req.user?.email || 'system',
          last_edited_at: new Date().toISOString(),
        }).eq('id', credit.id);
        commentsGenerated++;
      } catch (_) { /* Non-fatal */ }
    }

    // 2. Delete existing excellence opportunities
    const { data: existingOpps } = await supabase.from('esd_excellence_opportunities').select('id').eq('project_id', projectId);
    for (const o of (existingOpps || [])) await supabase.from('esd_excellence_opportunities').delete().eq('id', o.id);

    // 3. Regenerate ESD Excellence Opportunities
    const improvableCredits = (credits || []).filter(
      c => c.credit_status !== 'ScopedOut' && (c.credit_score ?? 0) < 100
    );

    for (const credit of improvableCredits) {
      try {
        const creditRef = parseCreditRef(credit.credit_name);
        const shortRef = buildShortRef(credit.category, creditRef);
        const scoreTierContext = getScoreTierContext(credit.category, creditRef);

        const userPrompt = `Generate an ESD Excellence Opportunity description for the following BESS credit.

Project: ${project.name || ''}, ${project.address || ''}
Credit: ${shortRef} ${credit.credit_name}
Current credit score: ${credit.credit_score ?? 0}%
Maximum possible score: 100%
Credit status: ${credit.credit_status}
Credit requirement: ${credit.credit_requirement || ''}
Credit weight in BESS: ${credit.credit_weight ?? 0}% of category score
${scoreTierContext ? `\nScore tier thresholds for this credit:\n${scoreTierContext}` : ''}

Write a 2–4 sentence opportunity description that:
- Starts with the credit ID and credit name (e.g. '${shortRef} ${credit.credit_name}:')
- States the current score and what the maximum available score is
- Identifies specifically what design change, documentation, or target adjustment would increase the score for this project
- Where a credit has defined score tiers or thresholds in BESS, states the next threshold that would increase the score and by how much
- Does not use phrases like 'consider' or 'you may wish to' — write as a direct identification of the gap and the specific change required`;

        const message = await anthropic.messages.create({
          model: 'claude-sonnet-4-6',
          max_tokens: 1024,
          system: OPP_SYSTEM,
          messages: [{ role: 'user', content: userPrompt }],
        });

        const description = message.content[0]?.text || '';
        await supabase.from('esd_excellence_opportunities').insert({
          project_id: projectId,
          credit_id: credit.id,
          credit_reference: shortRef,
          credit_name: credit.credit_name || '',
          category: credit.category || '',
          current_score: credit.credit_score ?? 0,
          max_score: 100,
          improvement_description: description,
          flag: 'Unflagged',
          flagged_by: '',
          deleted_by_giw: false,
        });
        oppsGenerated++;
      } catch (_) { /* Non-fatal */ }
    }

    return res.json({ success: true, commentsGenerated, oppsGenerated });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}];
