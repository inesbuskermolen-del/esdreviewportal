import { getUserFromToken } from '../config/supabase.js';

// Attach req.user if a valid Bearer token is present.
// Routes that require auth should call requireAuth middleware.
export async function attachUser(req, res, next) {
  try {
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    req.user = await getUserFromToken(token);
  } catch (_) {
    req.user = null;
  }
  next();
}

// Hard-require an authenticated user.
export function requireAuth(req, res, next) {
  if (!req.user) return res.status(401).json({ error: 'Unauthorized' });
  next();
}

// Require the user to be a GIW admin (@giw.com.au).
export function requireGIW(req, res, next) {
  if (!req.user) return res.status(401).json({ error: 'Unauthorized' });
  if (!req.user.email?.endsWith('@giw.com.au')) {
    return res.status(403).json({ error: 'GIW access only' });
  }
  next();
}
