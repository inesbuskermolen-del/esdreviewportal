import { Router } from 'express';
import { supabase, getUserFromToken } from '../config/supabase.js';

const router = Router();

// GET /api/auth/me — return the current user from their Bearer token
router.get('/me', async (req, res) => {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  const user = await getUserFromToken(token);
  if (!user) return res.status(401).json({ error: 'auth_required' });
  return res.json({ user });
});

// POST /api/auth/magic-link — send a magic link email
router.post('/magic-link', async (req, res) => {
  const { email, redirectTo } = req.body;
  if (!email) return res.status(400).json({ error: 'email required' });
  const { error } = await supabase.auth.signInWithOtp({
    email,
    options: { emailRedirectTo: redirectTo || process.env.FRONTEND_URL },
  });
  if (error) return res.status(400).json({ error: error.message });
  return res.json({ ok: true });
});

// POST /api/auth/logout — sign out the user
router.post('/logout', async (req, res) => {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (token) {
    // Revoke the session
    await supabase.auth.admin.signOut(token).catch(() => {});
  }
  return res.json({ ok: true });
});

export default router;
