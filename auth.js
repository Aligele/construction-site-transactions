const express = require('express');
const router = express.Router();
const supabase = require('../db');
const { comparePassword, signToken } = require('../auth');

router.post('/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password required' });
  }

  const { data: user, error } = await supabase
    .from('cst_users')
    .select('*')
    .eq('email', email.toLowerCase().trim())
    .single();

  if (error || !user || !comparePassword(password, user.password_hash)) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }

  const token = signToken(user);
  res.json({
    token,
    user: { id: user.id, full_name: user.full_name, email: user.email, role: user.role, site_id: user.site_id }
  });
});

module.exports = router;
