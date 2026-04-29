import Anthropic from '@anthropic-ai/sdk';
import { supabase } from '../../config/supabase.js';
import { attachUser, requireGIW } from '../../middleware/auth.js';

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const normalize = (s) =>
  (s || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();

// POST /api/functions/parseMatrixExcel
export default [attachUser, requireGIW, async (req, res) => {
  try {
    const { fileUrl, projectId } = req.body;
    if (!fileUrl || !projectId) return res.status(400).json({ error: 'Missing fileUrl or projectId' });

    // Fetch the Excel file and convert to base64
    const fileResponse = await fetch(fileUrl);
    if (!fileResponse.ok) return res.status(400).json({ error: 'Could not fetch Excel file' });
    const fileBuffer = await fileResponse.arrayBuffer();
    const fileBase64 = Buffer.from(fileBuffer).toString('base64');

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'document',
            source: { type: 'base64', media_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', data: fileBase64 },
          },
          {
            type: 'text',
            text: 'Extract all credit rows from this Excel spreadsheet. For each row return: the full credit name (including any category/reference prefix), whether the Mandatory column is Y (true) or N/blank (false), and the exact text from the Responsible Party column. Return JSON only with a "rows" array. Schema: { "rows": [{ "creditName": string, "mandatory": boolean, "responsibleParty": string }] }',
          },
        ],
      }],
    });

    let rows = [];
    try {
      const text = message.content[0]?.text?.replace(/^```json\s*/,'').replace(/\s*```$/,'') || '{}';
      rows = JSON.parse(text).rows || [];
    } catch (_) {
      return res.status(500).json({ error: 'Failed to parse Claude response' });
    }

    if (rows.length === 0) return res.status(400).json({ error: 'No rows extracted from Excel' });

    const { data: existingCredits } = await supabase.from('credits').select('id,credit_name').eq('project_id', projectId);

    let updatedCount = 0, unmatchedCount = 0;

    for (const row of rows) {
      if (!row.creditName) continue;
      const normRow = normalize(row.creditName);
      let best = null, bestScore = 0;
      for (const credit of (existingCredits || [])) {
        const normCredit = normalize(credit.credit_name);
        const words = normRow.split(' ').filter(w => w.length > 3);
        const matchCount = words.filter(w => normCredit.includes(w)).length;
        const score = words.length > 0 ? matchCount / words.length : 0;
        if (score > bestScore) { bestScore = score; best = credit; }
      }
      if (best && bestScore >= 0.5) {
        await supabase.from('credits').update({ mandatory: row.mandatory === true, responsible_party: row.responsibleParty || '' }).eq('id', best.id);
        updatedCount++;
      } else {
        unmatchedCount++;
      }
    }

    return res.json({ success: true, summary: { rowsInExcel: rows.length, creditsUpdated: updatedCount, unmatched: unmatchedCount } });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}];
