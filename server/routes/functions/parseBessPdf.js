import Anthropic from '@anthropic-ai/sdk';
import { supabase } from '../../config/supabase.js';
import { attachUser } from '../../middleware/auth.js';

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const GIW_MATRIX = {
  'Management 1.1': { mandatory: false, responsibleParty: 'ESD Consultant' },
  'Management 2.1': { mandatory: true,  responsibleParty: 'ESD Consultant' },
  'Management 2.2': { mandatory: true,  responsibleParty: 'ESD Consultant' },
  'Management 2.3': { mandatory: false, responsibleParty: 'ESD Consultant' },
  'Management 3.1': { mandatory: false, responsibleParty: 'Services / Developer' },
  'Management 3.2': { mandatory: false, responsibleParty: 'Services / Developer' },
  'Management 3.3': { mandatory: false, responsibleParty: 'Services / Developer' },
  'Management 4.1': { mandatory: false, responsibleParty: 'ESD Consultant' },
  'Integrated Water Management 1.1': { mandatory: true,  responsibleParty: 'Architect / Developer' },
  'Integrated Water Management 2.1': { mandatory: true,  responsibleParty: 'Civil / Services / Architect / Developer' },
  'Integrated Water Management 3.1': { mandatory: false, responsibleParty: 'Landscape / Architect / Developer' },
  'Integrated Water Management 4.1': { mandatory: false, responsibleParty: 'Services / Developer' },
  'Operational Energy 1.1': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Operational Energy 1.2': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Operational Energy 2.1': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 2.2': { mandatory: false, responsibleParty: 'Developer / Services' },
  'Operational Energy 2.6': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 2.7': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 3.1': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 3.2': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 3.3': { mandatory: false, responsibleParty: 'Developer / Services' },
  'Operational Energy 3.4': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Operational Energy 3.5': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 3.6': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 3.7': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 4.1': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 4.2': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 4.4': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Operational Energy 4.5': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Indoor Environment Quality 1.1': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 1.2': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 1.3': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 1.4': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 1.5': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 1.6': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 2.1': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 2.2': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 2.3': { mandatory: true,  responsibleParty: 'Developer / Architect / Services' },
  'Indoor Environment Quality 3.1': { mandatory: true,  responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 3.2': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 3.3': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 3.4': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 3.5': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Indoor Environment Quality 4.1': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.1': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.2': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.3': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.4': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.5': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 1.6': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 2.1': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Transport 2.2': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Transport 2.3': { mandatory: false, responsibleParty: 'Developer / Architect / Services' },
  'Waste & Resource Recovery 1.1': { mandatory: false, responsibleParty: 'Developer / Architect / Waste Consultant' },
  'Waste & Resource Recovery 2.1': { mandatory: false, responsibleParty: 'Developer / Architect / Waste Consultant' },
  'Waste & Resource Recovery 2.2': { mandatory: false, responsibleParty: 'Developer / Architect / Waste Consultant' },
  'Urban Ecology 1.1': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Urban Ecology 2.1': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 2.2': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 2.3': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 2.4': { mandatory: false, responsibleParty: 'Developer / Architect' },
  'Urban Ecology 3.1': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Urban Ecology 3.2': { mandatory: false, responsibleParty: 'Developer / Architect / Landscape' },
  'Innovation 1.1': { mandatory: false, responsibleParty: 'ESD Consultant / Developer' },
};

function findMatrixMatch(category, creditRef) {
  if (!category || !creditRef) return null;
  const key = `${category} ${creditRef}`;
  if (GIW_MATRIX[key]) return GIW_MATRIX[key];
  const baseKey = `${category} ${creditRef.replace(/[a-z]$/i, '')}`;
  if (GIW_MATRIX[baseKey]) return GIW_MATRIX[baseKey];
  return null;
}

// The actual processing — runs in background after HTTP response is sent
async function runParseBessPdf(fileUrl, projectId, userEmail, authHeader) {
  if (!process.env.ANTHROPIC_API_KEY) throw new Error('ANTHROPIC_API_KEY is not set in server/.env');

  // Fetch the PDF and convert to base64
  const pdfResponse = await fetch(fileUrl);
  if (!pdfResponse.ok) throw new Error('Could not fetch PDF from URL');
  const pdfBuffer = await pdfResponse.arrayBuffer();
  const pdfBase64 = Buffer.from(pdfBuffer).toString('base64');

  const prompt = `Parse this BESS (Built Environment Sustainability Scorecard) PDF report and extract structured data.

Extract:
1. PROJECT METADATA: projectName, address, projectId, date (YYYY-MM-DD), bessScore (number, e.g. 67)
2. CATEGORY WEIGHTS from the "Performance by category" table (e.g. Management=5, Transport=9). Use as creditWeight for every credit in that category.
3. ALL CREDITS (skip any with status "Disabled"):
   - creditRef (e.g. "1.1", "2.3a"), category, creditName, creditRequirement, rawStatus, creditStatus (Y/N/ScopedOut), creditScore (0-100), creditWeight, scopedOutReason
4. SUPPORTING EVIDENCE tables:
   "Shown on Floor Plans" → drawingType="FloorPlan"
   "Supporting Documentation" → drawingType="SupportingDoc"
   Fields: creditReference, drawingType, requirement, discipline=""

Return JSON only. No markdown. Schema:
{
  "projectName": string, "address": string, "projectId": string, "date": string, "bessScore": number,
  "credits": [{ "creditRef": string, "category": string, "creditName": string, "creditRequirement": string, "rawStatus": string, "creditStatus": string, "creditScore": number, "creditWeight": number, "scopedOutReason": string }],
  "supportingEvidence": [{ "creditReference": string, "drawingType": string, "requirement": string, "discipline": string }]
}`;

  const message = await anthropic.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: 8192,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'document',
          source: {
            type: 'base64',
            media_type: 'application/pdf',
            data: pdfBase64,
          },
        },
        {
          type: 'text',
          text: prompt,
        },
      ],
    }],
  });

  const rawText = message.content[0].text.trim();
  // Strip markdown code fences if present
  const jsonText = rawText.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '');
  const pdfData = JSON.parse(jsonText);

  const credits = pdfData.credits || [];
  const supportingEvidence = pdfData.supportingEvidence || [];

  // Clear old data
  const { data: existingCredits } = await supabase.from('credits').select('id').eq('project_id', projectId);
  for (const c of (existingCredits || [])) await supabase.from('credits').delete().eq('id', c.id);
  const { data: existingDrawings } = await supabase.from('drawing_requirements').select('id').eq('project_id', projectId);
  for (const d of (existingDrawings || [])) await supabase.from('drawing_requirements').delete().eq('id', d.id);

  // Save credits
  let savedCredits = 0, disabledSkipped = 0;
  for (const credit of credits) {
    if (credit.rawStatus?.toLowerCase() === 'disabled' || credit.creditStatus == null) { disabledSkipped++; continue; }
    const match = findMatrixMatch(credit.category, credit.creditRef);
    await supabase.from('credits').insert({
      project_id: projectId,
      category: credit.category || '',
      credit_name: `${credit.creditRef ? credit.creditRef + ' — ' : ''}${credit.creditName || ''}`,
      credit_requirement: credit.creditRequirement || '',
      mandatory: match?.mandatory === true,
      responsible_party: match?.responsibleParty || '',
      credit_status: credit.creditStatus,
      credit_score: typeof credit.creditScore === 'number' ? credit.creditScore : 0,
      credit_weight: typeof credit.creditWeight === 'number' ? credit.creditWeight : 0,
      scoped_out_reason: credit.scopedOutReason || '',
      is_excellence_target: false,
      last_edited_by: userEmail,
      last_edited_at: new Date().toISOString(),
    });
    savedCredits++;
  }

  // Save drawing requirements
  for (const item of supportingEvidence) {
    await supabase.from('drawing_requirements').insert({
      project_id: projectId,
      credit_reference: item.creditReference || '',
      drawing_type: item.drawingType || 'SupportingDoc',
      requirement: item.requirement || '',
      discipline: item.discipline || '',
      status: 'NotStarted',
      notes: '',
    });
  }

  // Update project metadata
  const projectUpdate = {};
  if (pdfData.bessScore != null) projectUpdate.bess_score = pdfData.bessScore;
  if (pdfData.address) projectUpdate.address = pdfData.address;
  if (pdfData.date) projectUpdate.date = pdfData.date;
  if (pdfData.projectId) projectUpdate.project_id = pdfData.projectId;
  if (Object.keys(projectUpdate).length > 0) {
    await supabase.from('projects').update(projectUpdate).eq('id', projectId);
  }

  // Mark parse as complete in project record
  await supabase.from('projects').update({ parse_status: 'complete' }).eq('id', projectId);

  console.log(`[parseBessPdf] Done — ${savedCredits} credits saved, ${disabledSkipped} disabled skipped for project ${projectId}`);

  // Fire-and-forget AI comment + opportunity generation
  fetch(`${process.env.SERVER_URL || 'http://localhost:3001'}/api/functions/regenerateContent`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: authHeader || '' },
    body: JSON.stringify({ projectId }),
  }).catch(() => {});
}

// POST /api/functions/parseBessPdf
// Returns immediately — processing happens in the background.
export default [attachUser, async (req, res) => {
  const { fileUrl, projectId } = req.body;
  if (!fileUrl || !projectId) return res.status(400).json({ error: 'Missing fileUrl or projectId' });

  // Mark as parsing so the frontend knows it started
  await supabase.from('projects').update({ parse_status: 'parsing' }).eq('id', projectId);

  // Start processing in background — do NOT await
  runParseBessPdf(fileUrl, projectId, req.user?.email, req.headers.authorization)
    .catch(async (err) => {
      console.error('[parseBessPdf] Error:', err.message);
      await supabase.from('projects')
        .update({ parse_status: 'error', parse_error: err.message })
        .eq('id', projectId);
    });

  // Respond immediately so the browser isn't waiting
  return res.json({ started: true });
}];
