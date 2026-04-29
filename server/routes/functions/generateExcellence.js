import Anthropic from '@anthropic-ai/sdk';
import { attachUser, requireAuth } from '../../middleware/auth.js';

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// POST /api/functions/generateExcellence
export default [attachUser, requireAuth, async (req, res) => {
  try {
    const { projectId, bessScore, credits } = req.body;

    const creditSummary = (credits || []).map(c =>
      `- [${c.id}] ${c.credit_name} | Status: ${c.credit_status} | Score: ${c.credit_score ?? 0}% | Weight: ${c.credit_weight ?? 0}% | Category: ${c.category} | Req: ${c.credit_requirement || ''} | ScopedOutReason: ${c.scoped_out_reason || ''}`
    ).join('\n');

    const prompt = `You are an ESD consultant analysing a BESS (Built Environment Sustainability Scorecard) credit matrix for a project.

Current BESS score: ${bessScore ?? 'unknown'}%

Full credit list (format: [creditId] name | status | score | weight | category | requirement):
${creditSummary}

Generate a list of ESD Excellence Opportunities — specific, actionable improvements that would increase the project's BESS score beyond its current level.

For each opportunity:
- Reference the specific credit name
- State the current score and the achievable improvement (e.g. "Currently 50% score — upgrading to X would achieve 100%")
- Describe SPECIFICALLY what design change would achieve it, with numbers where applicable
- Focus on credits that are Not Achieved (N) or partially achieved (Y with score < 100%)
- Also flag disabled/scoped-out credits that could be unlocked if assumptions change

Return JSON only. No markdown.`;

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      messages: [{ role: 'user', content: prompt }],
    });

    let opportunities = [];
    try {
      const text = message.content[0]?.text || '{}';
      const parsed = JSON.parse(text);
      opportunities = parsed.opportunities || [];
    } catch (_) {
      opportunities = [];
    }

    return res.json({ opportunities });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}];
