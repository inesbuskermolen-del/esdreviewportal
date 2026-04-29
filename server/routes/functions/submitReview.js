import nodemailer from 'nodemailer';
import { supabase } from '../../config/supabase.js';
import { attachUser, requireAuth } from '../../middleware/auth.js';

function getTransport() {
  return nodemailer.createTransport({
    host: process.env.SMTP_HOST || 'smtp.gmail.com',
    port: parseInt(process.env.SMTP_PORT || '587'),
    secure: process.env.SMTP_SECURE === 'true',
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS,
    },
  });
}

// POST /api/functions/submitReview
export default [attachUser, requireAuth, async (req, res) => {
  try {
    const { projectId, reviewerId } = req.body;

    const { data: reviewer } = await supabase.from('reviewers').select('*').eq('id', reviewerId).single();
    if (!reviewer) return res.status(404).json({ error: 'Reviewer not found' });
    if (reviewer.email !== req.user.email) return res.status(403).json({ error: 'Forbidden' });

    const { data: project } = await supabase.from('projects').select('*').eq('id', projectId).single();
    if (!project) return res.status(404).json({ error: 'Project not found' });

    const now = new Date().toISOString();
    await supabase.from('reviewers').update({ has_submitted: true, submitted_at: now }).eq('id', reviewerId);

    const dateStr = new Date(now).toLocaleDateString('en-AU', {
      day: '2-digit', month: 'long', year: 'numeric',
      hour: '2-digit', minute: '2-digit', timeZone: 'Australia/Sydney',
    });

    const projectLink = `${process.env.FRONTEND_URL || 'http://localhost:5173'}/project/${projectId}/matrix`;
    const discipline = reviewer.discipline || 'Unknown';
    const reviewerName = req.user.user_metadata?.full_name || reviewer.email;

    if (process.env.SMTP_USER) {
      const transport = getTransport();

      // Notify GIW
      await transport.sendMail({
        from: `"GIW ESD Review Portal" <${process.env.SMTP_USER}>`,
        to: process.env.GIW_NOTIFY_EMAIL || 'info@giw.com.au',
        subject: `[${discipline}] review submitted – ${project.name}`,
        text: `${reviewerName} (${reviewer.email}) (${discipline}) has completed their ESD review for ${project.name} on ${dateStr}.\n\nLog in to view their comments: ${projectLink}`,
      });

      // Confirm to reviewer
      await transport.sendMail({
        from: `"GIW Consultancy – ESD Review Portal" <${process.env.SMTP_USER}>`,
        to: reviewer.email,
        subject: `Your review for ${project.name} has been received`,
        text: `Hi ${reviewerName},\n\nThank you — your ${discipline} review for ${project.name} has been successfully submitted on ${dateStr}.\n\nGIW Consultancy has been notified and will review your comments.\n\nProject: ${project.name}\nDiscipline: ${discipline}\nSubmitted: ${dateStr}\n\nRegards,\nGIW Consultancy`,
      });
    }

    return res.json({ success: true });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}];
