/**
 * Generic entity CRUD router.
 *
 * Maps the Base44 entity method patterns onto Supabase table queries:
 *   base44.entities.Project.list()           → GET  /api/entities/projects
 *   base44.entities.Project.filter({id})     → GET  /api/entities/projects?id=xxx
 *   base44.entities.Project.create({...})    → POST /api/entities/projects
 *   base44.entities.Project.update(id,{...}) → PUT  /api/entities/projects/:id
 *   base44.entities.Project.delete(id)       → DELETE /api/entities/projects/:id
 *
 * Entity names (URL segment → Supabase table):
 *   projects, credits, credit-comments, drawing-requirements,
 *   esd-excellence-opportunities, reviewers
 */
import { Router } from 'express';
import { supabase } from '../config/supabase.js';
import { attachUser } from '../middleware/auth.js';

const router = Router();
router.use(attachUser);

// Map URL slugs → Supabase table names
const TABLE_MAP = {
  projects:                    'projects',
  credits:                     'credits',
  'credit-comments':           'credit_comments',
  'drawing-requirements':      'drawing_requirements',
  'esd-excellence-opportunities': 'esd_excellence_opportunities',
  reviewers:                   'reviewers',
};

function getTable(entity) {
  return TABLE_MAP[entity] || null;
}

// GET /api/entities/:entity — list or filter
router.get('/:entity', async (req, res) => {
  const table = getTable(req.params.entity);
  if (!table) return res.status(404).json({ error: 'Unknown entity' });

  const { sort, ...filters } = req.query;

  let query = supabase.from(table).select('*');

  // Apply equality filters from query string
  Object.entries(filters).forEach(([key, value]) => {
    query = query.eq(key, value);
  });

  // Support sort param (e.g. ?sort=-created_date)
  if (sort) {
    const desc = sort.startsWith('-');
    query = query.order(desc ? sort.slice(1) : sort, { ascending: !desc });
  }

  const { data, error } = await query;
  if (error) return res.status(400).json({ error: error.message });
  return res.json(data);
});

// POST /api/entities/:entity — create
router.post('/:entity', async (req, res) => {
  const table = getTable(req.params.entity);
  if (!table) return res.status(404).json({ error: 'Unknown entity' });

  const { data, error } = await supabase
    .from(table)
    .insert(req.body)
    .select()
    .single();

  if (error) {
    console.error(`[entities POST ${table}]`, error.message);
    return res.status(400).json({ error: error.message });
  }
  return res.status(201).json(data);
});

// PUT /api/entities/:entity/:id — update
router.put('/:entity/:id', async (req, res) => {
  const table = getTable(req.params.entity);
  if (!table) return res.status(404).json({ error: 'Unknown entity' });

  const { data, error } = await supabase
    .from(table)
    .update(req.body)
    .eq('id', req.params.id)
    .select()
    .single();

  if (error) return res.status(400).json({ error: error.message });
  return res.json(data);
});

// DELETE /api/entities/:entity/:id — delete
router.delete('/:entity/:id', async (req, res) => {
  const table = getTable(req.params.entity);
  if (!table) return res.status(404).json({ error: 'Unknown entity' });

  const { error } = await supabase
    .from(table)
    .delete()
    .eq('id', req.params.id);

  if (error) return res.status(400).json({ error: error.message });
  return res.status(204).send();
});

export default router;
