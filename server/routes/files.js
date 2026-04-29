/**
 * File upload route — replaces base44.integrations.Core.UploadFile()
 * Returns a public URL that can be passed to backend functions.
 */
import { Router } from 'express';
import multer from 'multer';
import { supabase } from '../config/supabase.js';
import { attachUser } from '../middleware/auth.js';

const router = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

router.use(attachUser);

// POST /api/files/upload
router.post('/upload', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file provided' });

  const ext = req.file.originalname.split('.').pop();
  const fileName = `${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`;

  const { error } = await supabase.storage
    .from('uploads')
    .upload(fileName, req.file.buffer, { contentType: req.file.mimetype });

  if (error) return res.status(400).json({ error: error.message });

  const { data } = supabase.storage.from('uploads').getPublicUrl(fileName);
  return res.json({ file_url: data.publicUrl });
});

export default router;
