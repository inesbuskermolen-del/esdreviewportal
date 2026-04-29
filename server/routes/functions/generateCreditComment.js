import Anthropic from '@anthropic-ai/sdk';
import { supabase } from '../../config/supabase.js';
import { attachUser, requireAuth } from '../../middleware/auth.js';

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const SYSTEM_PROMPT = `You are an ESD consultant writing internal project review notes for a BESS (Built Environment Sustainability Scorecard) assessment. Your notes will be read by the design team. Write in plain, direct professional language. Every sentence must reference specific numbers, ratings, or facts from the project data provided. Do not write generic statements that could apply to any project. Do not mention AI, automation, or report generation. Do not use phrases like 'the project achieves' as an opener for every sentence — vary the sentence structure. Write in third person. Length should be 3–6 sentences.`;

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
  // creditRef may arrive as "Mgmt 1.1" or "1.1" — extract the numeric part only
  const ref = (creditRef || '').replace(/^[a-z]+\s+/i, '').replace(/[a-z]$/i, '').trim();
  return CREDIT_TEMPLATES[category]?.[ref] || null;
}

// POST /api/functions/generateCreditComment
export default [attachUser, requireAuth, async (req, res) => {
  try {
    const { creditId, creditName, creditRequirement, creditStatus, creditScore, scopedOutReason, projectId, creditRef, category } = req.body;

    let projectName = '', projectAddress = '';
    if (projectId) {
      const { data } = await supabase.from('projects').select('name,address').eq('id', projectId).single();
      if (data) { projectName = data.name || ''; projectAddress = data.address || ''; }
    }

    const categoryTemplate = getCategoryTemplate(category, creditRef);

    const userPrompt = categoryTemplate
      ? `Generate a Comments GIW entry for the following BESS credit for this project.

Project: ${projectName}${projectAddress ? ', ' + projectAddress : ''}
Credit: ${creditRef || ''} ${creditName}
Credit status: ${creditStatus}
Credit requirement: ${creditRequirement || ''}
${scopedOutReason ? `Scoped-out reason: ${scopedOutReason}` : ''}

Use the following template as your answer:
"${categoryTemplate}"

Rules:
- If status is Y (achieved): state the template text as a confirmed fact, filling in any specific numbers from the credit requirement
- If status is N (not achieved): rewrite the template to state what is not provided or missing
- If status is ScopedOut: state the scoped-out reason
- Replace any [X] placeholders with the actual numbers extracted from the credit requirement
- Keep to 1–2 sentences only. Do not add analytical content, score gaps, or extra context.`
      : `Generate a Comments GIW entry for the following BESS credit for this project.

Project: ${projectName}${projectAddress ? ', ' + projectAddress : ''}
Credit: ${creditRef || ''} ${creditName}
Credit score: ${creditScore ?? 0}%
Credit status: ${creditStatus}
Credit requirement: ${creditRequirement || ''}
${scopedOutReason ? `Scoped-out reason: ${scopedOutReason}` : ''}

Write a 3–6 sentence comment that:
- States what this credit requires (briefly, one clause only)
- States what the project has achieved or not achieved, citing any specific numbers from the credit requirement above
- If score is 100%: ends with a confirmation the credit is fully met
- If score is between 1% and 99%: states the current achievement level and what the numerical gap is to the next score threshold, without suggesting how to fix it
- If score is 0% and status is N: states what is missing relative to the credit requirement
- If status is ScopedOut: states the specific reason this credit was scoped out`;

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 1024,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userPrompt }],
    });

    const comment = message.content[0]?.text || '';
    return res.json({ comment });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}];
