const express = require('express');
const router = express.Router();
const supabase = require('../db');
const { requireAuth, requireRole } = require('../auth');

const CATEGORIES = ['materials', 'labor', 'equipment', 'fuel', 'other'];

// Create a transaction (site clerk logs the day's activity/expense)
router.post('/', requireAuth, requireRole('clerk', 'admin'), async (req, res) => {
  const { site_id, category, description, amount, transaction_date } = req.body;

  if (!site_id || !category || !description || !amount) {
    return res.status(400).json({ error: 'site_id, category, description and amount are required' });
  }
  if (!CATEGORIES.includes(category)) {
    return res.status(400).json({ error: `category must be one of: ${CATEGORIES.join(', ')}` });
  }

  const { data, error } = await supabase
    .from('cst_transactions')
    .insert({
      site_id,
      category,
      description,
      amount,
      transaction_date: transaction_date || new Date().toISOString().slice(0, 10),
      created_by: req.user.id
    })
    .select()
    .single();

  if (error) return res.status(500).json({ error: error.message });
  res.status(201).json(data);
});

// List transactions, optionally filtered by status/site/category
router.get('/', requireAuth, async (req, res) => {
  const { status, site_id, category } = req.query;

  let query = supabase.from('cst_transactions').select('*').order('created_at', { ascending: false });
  if (status) query = query.eq('status', status);
  if (site_id) query = query.eq('site_id', site_id);
  if (category) query = query.eq('category', category);

  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

// Manager approval
router.post('/:id/approve/manager', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { data: txn, error: fetchErr } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fetchErr || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'pending') return res.status(400).json({ error: `Cannot approve a transaction with status "${txn.status}"` });

  const updates = {
    manager_approved: true,
    manager_approved_by: req.user.id,
    manager_approved_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
  if (txn.finance_approved) updates.status = 'approved';

  const { data, error } = await supabase.from('cst_transactions').update(updates).eq('id', id).select().single();
  if (error) return res.status(500).json({ error: error.message });

  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'approve_manager' });
  res.json(data);
});

// Finance approval
router.post('/:id/approve/finance', requireAuth, requireRole('finance', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { data: txn, error: fetchErr } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fetchErr || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'pending') return res.status(400).json({ error: `Cannot approve a transaction with status "${txn.status}"` });

  const updates = {
    finance_approved: true,
    finance_approved_by: req.user.id,
    finance_approved_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
  if (txn.manager_approved) updates.status = 'approved';

  const { data, error } = await supabase.from('cst_transactions').update(updates).eq('id', id).select().single();
  if (error) return res.status(500).json({ error: error.message });

  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'approve_finance' });
  res.json(data);
});

// Reject (either manager or finance can reject)
router.post('/:id/reject', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { reason } = req.body;

  const { data: txn, error: fetchErr } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fetchErr || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'pending') return res.status(400).json({ error: `Cannot reject a transaction with status "${txn.status}"` });

  const { data, error } = await supabase
    .from('cst_transactions')
    .update({
      status: 'rejected',
      rejected_by: req.user.id,
      rejected_at: new Date().toISOString(),
      rejection_reason: reason || null,
      updated_at: new Date().toISOString()
    })
    .eq('id', id)
    .select()
    .single();

  if (error) return res.status(500).json({ error: error.message });

  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'reject', note: reason || null });
  res.json(data);
});

// Finance settles/pays an approved transaction
router.post('/:id/pay', requireAuth, requireRole('finance', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { data: txn, error: fetchErr } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fetchErr || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'approved') return res.status(400).json({ error: 'Only fully approved transactions can be marked paid' });

  const { data, error } = await supabase
    .from('cst_transactions')
    .update({ status: 'paid', paid_by: req.user.id, paid_at: new Date().toISOString(), updated_at: new Date().toISOString() })
    .eq('id', id)
    .select()
    .single();

  if (error) return res.status(500).json({ error: error.message });

  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'mark_paid' });
  res.json(data);
});

module.exports = router;
