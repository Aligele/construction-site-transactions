require('dotenv').config();
const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const ExcelJS = require('exceljs');
const { createClient } = require('@supabase/supabase-js');

const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change-me';
const CATEGORIES = ['materials', 'labor', 'equipment', 'fuel', 'other'];

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY, {
  auth: { autoRefreshToken: false, persistSession: false }
});

const app = express();
app.use(cors());
app.use(express.json());

// ---------- Auth helpers ----------
function signToken(user) {
  return jwt.sign(
    { id: user.id, role: user.role, site_id: user.site_id, full_name: user.full_name },
    JWT_SECRET,
    { expiresIn: '12h' }
  );
}
function requireAuth(req, res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (!token) return res.status(401).json({ error: 'Missing token' });
  try {
    req.user = jwt.verify(token, JWT_SECRET);
    next();
  } catch {
    return res.status(401).json({ error: 'Invalid or expired token' });
  }
}
function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.user || !roles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Not permitted for this role' });
    }
    next();
  };
}

// ---------- Frontend (embedded, no static folder needed) ----------
const INDEX_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Site Transactions</title>
<style>
:root{
  --sidebar:#16302278;--sidebar-bg:#173B29;--sidebar-active:#2f9e5c;--sidebar-text:#cfe3d6;--sidebar-text-dim:#7fa891;
  --navy:#1f2d3d;--amber:#d98c2b;--green:#2f7d4f;--red:#b03a3a;--bg:#f5f6f8;--card:#fff;--border:#dde1e6;
}
*{box-sizing:border-box;}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:var(--bg);color:var(--navy);}
#app{min-height:100vh;}

/* ---- App shell / sidebar ---- */
.shell{display:flex;min-height:100vh;}
.sidebar{
  position:fixed;top:0;left:0;bottom:0;width:250px;background:linear-gradient(180deg,#173B29,#12301F);
  color:var(--sidebar-text);z-index:40;transform:translateX(-100%);transition:transform .22s ease;
  overflow-y:auto;padding-bottom:24px;
}
.sidebar.open{transform:translateX(0);}
.brand{display:flex;align-items:center;gap:12px;padding:20px 18px;border-bottom:1px solid rgba(255,255,255,.08);}
.brand-logo{width:44px;height:44px;border-radius:50%;background:#eaf7ee;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 0 0 2px rgba(255,255,255,.15);}
.brand-title{font-weight:700;color:#fff;font-size:15px;line-height:1.2;}
.brand-sub{font-size:11px;color:var(--sidebar-text-dim);margin-top:2px;}
.nav-section-label{font-size:11px;letter-spacing:.08em;color:var(--sidebar-text-dim);padding:16px 18px 6px;}
.nav-item{display:flex;align-items:center;gap:10px;margin:2px 10px;padding:10px 12px;border-radius:8px;color:var(--sidebar-text);text-decoration:none;font-size:14px;cursor:pointer;}
.nav-item:hover{background:rgba(255,255,255,.06);}
.nav-item.active{background:var(--sidebar-active);color:#fff;font-weight:600;}
.nav-icon{font-size:16px;width:18px;text-align:center;}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:30;display:none;}
.overlay.open{display:block;}
.main{flex:1;min-width:0;}
.topbar-app{display:flex;align-items:center;justify-content:space-between;background:#fff;border-bottom:1px solid var(--border);padding:12px 16px;position:sticky;top:0;z-index:20;}
.menu-btn{background:transparent;border:none;color:var(--navy);font-size:20px;padding:4px 8px;cursor:pointer;}
.topbar-right{display:flex;align-items:center;gap:14px;}
.avatar{width:34px;height:34px;border-radius:50%;background:var(--sidebar-active);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;}
.bell{font-size:18px;position:relative;}
.bell .dot{position:absolute;top:-3px;right:-4px;background:var(--red);color:#fff;font-size:9px;border-radius:10px;padding:1px 4px;}
#content{max-width:900px;margin:0 auto;padding:20px;}

.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:20px;margin-bottom:16px;}
h1{font-size:22px;margin:0 0 4px;} h2{font-size:17px;margin:0 0 12px;} .muted{color:#6b7684;font-size:13px;}
input,select,textarea,button{font-family:inherit;font-size:14px;padding:9px 10px;border-radius:6px;border:1px solid var(--border);}
label{display:block;font-size:13px;margin:10px 0 4px;font-weight:600;}
button{background:var(--navy);color:#fff;border:none;cursor:pointer;font-weight:600;}
button:hover{opacity:.9;} button.secondary{background:#eef0f3;color:var(--navy);}
button.danger{background:var(--red);} button.success{background:var(--green);}
.row{display:flex;gap:10px;flex-wrap:wrap;} .row>*{flex:1;min-width:160px;}
table{width:100%;border-collapse:collapse;margin-top:10px;font-size:13px;}
th,td{text-align:left;padding:8px 6px;border-bottom:1px solid var(--border);} th{color:#6b7684;font-weight:600;}
.badge{padding:2px 8px;border-radius:20px;font-size:12px;font-weight:600;}
.badge.pending{background:#fdf0dc;color:var(--amber);} .badge.approved{background:#e2f1e8;color:var(--green);}
.badge.paid{background:#dfeee5;color:var(--green);} .badge.rejected{background:#f6e2e2;color:var(--red);}
.actions button{margin-right:4px;padding:5px 8px;font-size:12px;}
.topbar{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;}
.error{color:var(--red);font-size:13px;margin-top:8px;}
</style>
</head>
<body>
<div id="app"></div>
<script>
const API = '/api';
let state = { token: localStorage.getItem('cst_token'), user: JSON.parse(localStorage.getItem('cst_user') || 'null') };
const appEl = document.getElementById('app');
const TRUCK_SVG = '<svg width="26" height="26" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="18" width="24" height="14" rx="2" fill="#1f2d3d"/><path d="M26 22h9l7 7v3a2 2 0 0 1-2 2h-1" stroke="#1f2d3d" stroke-width="2.5" fill="none" stroke-linejoin="round"/><circle cx="13" cy="35" r="4" fill="#173B29" stroke="#1f2d3d" stroke-width="1.5"/><circle cx="34" cy="35" r="4" fill="#173B29" stroke="#1f2d3d" stroke-width="1.5"/></svg>';
function initials(name){ return (name||'?').split(' ').filter(Boolean).slice(0,2).map(w=>w[0].toUpperCase()).join(''); }
function save(token, user){ state={token,user}; localStorage.setItem('cst_token',token); localStorage.setItem('cst_user', JSON.stringify(user)); }
function logout(){ localStorage.removeItem('cst_token'); localStorage.removeItem('cst_user'); state={token:null,user:null}; render(); }
async function api(path, opts={}) {
  const res = await fetch(API+path, { ...opts, headers: { 'Content-Type':'application/json', ...(state.token?{Authorization:'Bearer '+state.token}:{}), ...(opts.headers||{}) } });
  if (!res.ok) { const b = await res.json().catch(()=>({})); throw new Error(b.error||'Request failed'); }
  return res.json();
}
function money(n){ return 'KES ' + Number(n).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}); }
function badge(s){ return '<span class="badge '+s+'">'+s+'</span>'; }
function render(){ state.token ? renderDashboard() : renderLogin(); }
function renderLogin(){
  appEl.innerHTML = '<div class="card" style="max-width:360px;margin:60px auto;"><div style="display:flex;align-items:center;gap:10px;margin-bottom:6px;"><div class="brand-logo" style="box-shadow:none;background:#eaf7ee;">'+TRUCK_SVG+'</div><div><h1 style="margin:0;">Site Transactions</h1><p class="muted" style="margin:0;">Construction Portal</p></div></div><p class="muted">Sign in to continue</p><label>Email</label><input id="email" type="email" placeholder="you@example.com" /><label>Password</label><input id="password" type="password" placeholder="********" /><div style="margin-top:16px;"><button id="loginBtn" style="width:100%;">Log in</button></div><div class="error" id="err"></div></div>';
  document.getElementById('loginBtn').onclick = async () => {
    const email = document.getElementById('email').value.trim();
    const password = document.getElementById('password').value;
    try { const {token,user} = await api('/auth/login',{method:'POST',body:JSON.stringify({email,password})}); save(token,user); render(); }
    catch(e){ document.getElementById('err').textContent = e.message; }
  };
}
async function renderDashboard(){
  const {user} = state;
  appEl.innerHTML =
    '<div class="shell">' +
      '<div class="overlay" id="overlay"></div>' +
      '<aside class="sidebar" id="sidebar">' +
        '<div class="brand"><div class="brand-logo">'+TRUCK_SVG+'</div><div><div class="brand-title">Site Transactions</div><div class="brand-sub">Construction Portal</div></div></div>' +
        '<div class="nav-section-label">MAIN</div>' +
        '<a class="nav-item active" id="navDashboard"><span class="nav-icon">&#8962;</span> Dashboard</a>' +
        '<div class="nav-section-label">ACCOUNT</div>' +
        '<a class="nav-item" id="navNotifications"><span class="nav-icon">&#128276;</span> Notifications</a>' +
        '<a class="nav-item" id="navLogout"><span class="nav-icon">&#8618;</span> Log out</a>' +
      '</aside>' +
      '<main class="main">' +
        '<div class="topbar-app">' +
          '<button class="menu-btn" id="menuBtn">&#9776;</button>' +
          '<div class="topbar-right"><span class="bell" id="bellIcon">&#128276;<span class="dot" id="bellDot" style="display:none;"></span></span><div class="avatar">'+initials(user.full_name)+'</div><div><div style="font-weight:600;font-size:13px;">'+user.full_name+'</div><div class="muted" style="font-size:11px;text-transform:capitalize;">'+user.role+'</div></div></div>' +
        '</div>' +
        '<div id="content"></div>' +
      '</main>' +
    '</div>';

  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('overlay');
  document.getElementById('menuBtn').onclick = () => { sidebar.classList.toggle('open'); overlay.classList.toggle('open'); };
  overlay.onclick = () => { sidebar.classList.remove('open'); overlay.classList.remove('open'); };
  document.getElementById('navLogout').onclick = logout;
  document.getElementById('navDashboard').onclick = () => { document.getElementById('content').scrollIntoView({behavior:'smooth'}); sidebar.classList.remove('open'); overlay.classList.remove('open'); };
  document.getElementById('navNotifications').onclick = () => { const n=document.getElementById('notifList'); if(n) n.scrollIntoView({behavior:'smooth',block:'center'}); sidebar.classList.remove('open'); overlay.classList.remove('open'); };
  document.getElementById('bellIcon').onclick = document.getElementById('navNotifications').onclick;

  const content = document.getElementById('content');
  let contentHtml = '<div class="card"><div class="topbar"><h2>Notifications</h2><span class="muted" id="notifCount"></span></div><div id="notifList" class="muted">Loading...</div></div>';
  if (user.role==='clerk' || user.role==='admin') {
    contentHtml += '<div class="card"><h2>Log a transaction</h2><div class="row"><div><label>Site ID</label><input id="site_id" placeholder="site UUID" value="'+(user.site_id||'')+'" /></div><div><label>Category</label><select id="category"><option value="materials">Materials</option><option value="labor">Labor</option><option value="equipment">Equipment</option><option value="fuel">Fuel</option><option value="other">Other</option></select></div><div><label>Amount (KES)</label><input id="amount" type="number" step="0.01" /></div></div><label>Description</label><textarea id="description" rows="2" style="width:100%;"></textarea><div style="margin-top:12px;"><button id="submitTxn">Submit</button></div><div class="error" id="submitErr"></div></div>';
  }
  if (user.role==='manager' || user.role==='admin') {
    contentHtml += '<div class="card"><h2>Add a clerk or finance login</h2><div class="row"><div><label>Full name</label><input id="new_name" placeholder="Full name" /></div><div><label>Email</label><input id="new_email" type="email" placeholder="person@example.com" /></div><div><label>Role</label><select id="new_role"><option value="clerk">Clerk</option><option value="finance">Finance</option></select></div></div><div style="margin-top:12px;"><button id="createUserBtn">Create login</button></div><div class="error" id="createUserErr"></div><div id="createUserResult"></div></div>';
    contentHtml += '<div class="card"><h2>Message templates</h2><p class="muted">Tokens: {{description}} {{amount}} {{site_name}} {{paid_date}} {{reason_suffix}}</p><div id="templateList" class="muted">Loading...</div></div>';
  }
  contentHtml += '<div class="card"><div class="topbar"><h2>Transactions</h2><div>'+(['manager','finance','admin'].includes(user.role)?'<button id="exportBtn" class="secondary">Download Excel</button>':'')+(['manager','admin'].includes(user.role)?' <button id="clearAllBtn" class="danger">Clear all data</button>':'')+'</div></div><div id="table"></div></div>';
  content.innerHTML = contentHtml;

  loadNotifications();
  if (document.getElementById('templateList')) loadTemplates();

  if (document.getElementById('createUserBtn')) {
    document.getElementById('createUserBtn').onclick = async () => {
      const full_name = document.getElementById('new_name').value.trim();
      const email = document.getElementById('new_email').value.trim();
      const role = document.getElementById('new_role').value;
      try {
        const result = await api('/users', {method:'POST', body: JSON.stringify({full_name, email, role})});
        document.getElementById('createUserErr').textContent = '';
        document.getElementById('createUserResult').innerHTML = '<div class="card" style="background:#e2f1e8;margin-top:10px;"><strong>Login created.</strong><br/>Email: '+result.email+'<br/>Temporary password: <strong>'+result.temporary_password+'</strong><br/><span class="muted">Share this with them directly — it will not be shown again.</span></div>';
        document.getElementById('new_name').value=''; document.getElementById('new_email').value='';
      } catch(e) { document.getElementById('createUserErr').textContent = e.message; }
    };
  }

  if (document.getElementById('submitTxn')) {
    document.getElementById('submitTxn').onclick = async () => {
      const site_id = document.getElementById('site_id').value.trim();
      const category = document.getElementById('category').value;
      const amount = document.getElementById('amount').value;
      const description = document.getElementById('description').value.trim();
      try { await api('/transactions',{method:'POST',body:JSON.stringify({site_id,category,amount,description})}); document.getElementById('description').value=''; document.getElementById('amount').value=''; loadTable(); }
      catch(e){ document.getElementById('submitErr').textContent = e.message; }
    };
  }
  if (document.getElementById('exportBtn')) {
    document.getElementById('exportBtn').onclick = () => {
      fetch(API+'/export', {headers:{Authorization:'Bearer '+state.token}}).then(r=>r.blob()).then(blob=>{
        const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url;
        a.download = 'site-transactions-'+new Date().toISOString().slice(0,10)+'.xlsx'; a.click();
      });
    };
  }
  if (document.getElementById('clearAllBtn')) {
    document.getElementById('clearAllBtn').onclick = async () => {
      if (!confirm('This will permanently delete ALL transaction data for your site. Are you sure?')) return;
      if (!confirm('Really sure? This cannot be undone.')) return;
      try { await api('/transactions?confirm=yes', {method:'DELETE'}); loadTable(); }
      catch(e){ alert(e.message); }
    };
  }
  loadTable();
}
async function loadTable(){
  const {user} = state;
  const rows = await api('/transactions');
  let html = '<table><thead><tr><th>Date</th><th>Category</th><th>Description</th><th>Amount</th><th>Status</th><th>Actions</th></tr></thead><tbody>';
  rows.forEach(t => {
    html += '<tr><td>'+t.transaction_date+'</td><td>'+t.category+'</td><td>'+t.description+'</td><td>'+money(t.amount)+'</td><td>'+badge(t.status)+'</td><td class="actions">'+actionButtons(t,user)+' <button class="secondary" data-details="'+t.id+'">Details</button></td></tr>';
    html += '<tr id="details-'+t.id+'" style="display:none;"><td colspan="6"><div id="details-body-'+t.id+'" class="muted">Loading...</div></td></tr>';
  });
  html += '</tbody></table>';
  document.getElementById('table').innerHTML = html;
  document.querySelectorAll('[data-action]').forEach(btn => {
    btn.onclick = async () => {
      const {action,id} = btn.dataset;
      try { await api('/transactions/'+id+'/'+action,{method:'POST',body:JSON.stringify({})}); loadTable(); }
      catch(e){ alert(e.message); }
    };
  });
  document.querySelectorAll('[data-receipt]').forEach(btn => {
    btn.onclick = () => {
      fetch(API+'/transactions/'+btn.dataset.receipt+'/receipt', {headers:{Authorization:'Bearer '+state.token}})
        .then(r=>r.text()).then(html=>{ const w = window.open('', '_blank'); w.document.write(html); w.document.close(); });
    };
  });
  document.querySelectorAll('[data-delete]').forEach(btn => {
    btn.onclick = async () => {
      if (!confirm('Delete this transaction permanently? This cannot be undone.')) return;
      try { await api('/transactions/'+btn.dataset.delete, {method:'DELETE'}); loadTable(); }
      catch(e){ alert(e.message); }
    };
  });
  document.querySelectorAll('[data-details]').forEach(btn => {
    btn.onclick = () => toggleDetails(btn.dataset.details);
  });
  document.querySelectorAll('[data-edit]').forEach(btn => {
    btn.onclick = async () => {
      const row = rows.find(r => r.id === btn.dataset.edit);
      const newDesc = prompt('Description:', row.description);
      if (newDesc === null) return;
      const newAmount = prompt('Amount (KES):', row.amount);
      if (newAmount === null) return;
      const newStatus = prompt('Status (pending/approved/rejected/paid):', row.status);
      if (newStatus === null) return;
      try {
        await api('/transactions/'+row.id, {method:'PUT', body: JSON.stringify({description:newDesc, amount:newAmount, status:newStatus})});
        loadTable();
      } catch(e){ alert(e.message); }
    };
  });
}
function actionButtons(t,user){
  const canAct = ['manager','finance','admin'].includes(user.role);
  let b = '';
  if (t.status==='pending' && canAct) {
    if ((user.role==='manager'||user.role==='admin') && !t.manager_approved) b += '<button class="success" data-action="approve/manager" data-id="'+t.id+'">Approve (Manager)</button>';
    if ((user.role==='finance'||user.role==='admin') && !t.finance_approved) b += '<button class="success" data-action="approve/finance" data-id="'+t.id+'">Approve (Finance)</button>';
    b += '<button class="danger" data-action="reject" data-id="'+t.id+'">Reject</button>';
  }
  if (t.status==='approved' && (user.role==='finance'||user.role==='admin')) b += '<button data-action="pay" data-id="'+t.id+'">Mark Paid</button>';
  if (t.status==='paid') b += '<button class="secondary" data-receipt="'+t.id+'">Receipt</button>';
  if (user.role==='manager' || user.role==='admin') {
    b += '<button class="secondary" data-edit="'+t.id+'">Edit</button>';
    b += '<button class="danger" data-delete="'+t.id+'">Delete</button>';
  }
  return b;
}
let openDetails = {};
async function toggleDetails(txnId){
  const row = document.getElementById('details-'+txnId);
  if (openDetails[txnId]) { row.style.display='none'; openDetails[txnId]=false; return; }
  row.style.display='table-row'; openDetails[txnId]=true;
  const body = document.getElementById('details-body-'+txnId);
  body.textContent = 'Loading...';
  try {
    const [vendorLinks, allVendors] = await Promise.all([
      api('/transactions/'+txnId+'/vendors'),
      api('/vendors')
    ]);
    const canManage = ['manager','finance','admin'].includes(state.user.role);
    let html = '<strong>Assigned partners/vendors</strong>';
    html += vendorLinks.length ? '<ul style="padding-left:18px;margin:6px 0;">' + vendorLinks.map(v =>
      '<li>'+v.vendor.name+' <span class="muted">('+v.vendor.role_type+')</span>'+(canManage?' <button class="danger" style="padding:2px 6px;font-size:11px;" data-remove-vendor="'+v.id+'">remove</button>':'')+'</li>'
    ).join('') + '</ul>' : '<div class="muted">No vendors assigned.</div>';

    if (canManage) {
      html += '<div class="row" style="margin-top:8px;"><select id="vendorSelect-'+txnId+'">' +
        allVendors.map(v => '<option value="'+v.id+'">'+v.name+' ('+v.role_type+')</option>').join('') +
        '</select><button data-assign-vendor="'+txnId+'">Assign</button></div>';
      html += '<div class="muted" style="margin-top:6px;">New vendor: <input id="newVendorName-'+txnId+'" placeholder="Name" style="width:120px;" /> <select id="newVendorType-'+txnId+'"><option value="vendor">Vendor</option><option value="inspector">Inspector</option><option value="legal_counsel">Legal Counsel</option><option value="underwriter">Underwriter</option><option value="other">Other</option></select> <button class="secondary" data-add-vendor="'+txnId+'">Add</button></div>';
    }
    body.innerHTML = html;

    body.querySelectorAll('[data-remove-vendor]').forEach(btn => {
      btn.onclick = async () => { try { await api('/transaction-vendors/'+btn.dataset.removeVendor, {method:'DELETE'}); toggleDetails(txnId); openDetails[txnId]=false; toggleDetails(txnId); } catch(e){ alert(e.message); } };
    });
    const assignBtn = body.querySelector('[data-assign-vendor]');
    if (assignBtn) assignBtn.onclick = async () => {
      const vendor_id = document.getElementById('vendorSelect-'+txnId).value;
      try { await api('/transactions/'+txnId+'/vendors', {method:'POST', body: JSON.stringify({vendor_id})}); openDetails[txnId]=false; toggleDetails(txnId); }
      catch(e){ alert(e.message); }
    };
    const addBtn = body.querySelector('[data-add-vendor]');
    if (addBtn) addBtn.onclick = async () => {
      const name = document.getElementById('newVendorName-'+txnId).value.trim();
      const role_type = document.getElementById('newVendorType-'+txnId).value;
      if (!name) return;
      try {
        const v = await api('/vendors', {method:'POST', body: JSON.stringify({name, role_type})});
        await api('/transactions/'+txnId+'/vendors', {method:'POST', body: JSON.stringify({vendor_id: v.id})});
        openDetails[txnId]=false; toggleDetails(txnId);
      } catch(e){ alert(e.message); }
    };
  } catch(e) { body.textContent = 'Could not load details: ' + e.message; }
}
async function loadTemplates(){
  const el = document.getElementById('templateList');
  try {
    const templates = await api('/templates');
    el.innerHTML = templates.map(t =>
      '<div style="margin-bottom:12px;"><label>'+t.event_type+'</label><textarea data-template="'+t.event_type+'" rows="2" style="width:100%;">'+t.template+'</textarea><button class="secondary" data-save-template="'+t.event_type+'" style="margin-top:4px;">Save</button></div>'
    ).join('');
    el.querySelectorAll('[data-save-template]').forEach(btn => {
      btn.onclick = async () => {
        const eventType = btn.dataset.saveTemplate;
        const value = el.querySelector('[data-template="'+eventType+'"]').value;
        try { await api('/templates/'+eventType, {method:'PUT', body: JSON.stringify({template:value})}); btn.textContent='Saved'; setTimeout(()=>btn.textContent='Save',1200); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) { el.textContent = 'Could not load templates.'; }
}
async function loadNotifications(){
  try {
    const notes = await api('/notifications');
    const unread = notes.filter(n => !n.read_at).length;
    document.getElementById('notifCount').textContent = unread ? (unread+' new') : '';
    const dot = document.getElementById('bellDot');
    if (dot) { dot.style.display = unread ? 'inline-block' : 'none'; dot.textContent = unread; }
    if (!notes.length) { document.getElementById('notifList').textContent = 'No notifications yet.'; return; }
    document.getElementById('notifList').innerHTML = notes.slice(0,8).map(n =>
      '<div style="padding:8px 0;border-bottom:1px solid #eee;'+(n.read_at?'':'font-weight:600;')+'">'+n.message+'<div class="muted" style="font-weight:normal;">'+new Date(n.created_at).toLocaleString()+'</div></div>'
    ).join('');
  } catch(e) { document.getElementById('notifList').textContent = 'Could not load notifications.'; }
}
render();
</script>
</body>
</html>`;

app.get('/', (req, res) => res.type('html').send(INDEX_HTML));

// ---------- Auth routes ----------
app.post('/api/auth/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ error: 'Email and password required' });

  const { data: user, error } = await supabase.from('cst_users').select('*').eq('email', email.toLowerCase().trim()).single();
  if (error || !user || !bcrypt.compareSync(password, user.password_hash)) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }
  const token = signToken(user);
  res.json({ token, user: { id: user.id, full_name: user.full_name, email: user.email, role: user.role, site_id: user.site_id } });
});

// ---------- Transaction routes ----------
app.post('/api/transactions', requireAuth, requireRole('clerk', 'admin'), async (req, res) => {
  const { site_id, category, description, amount, transaction_date } = req.body;
  if (!site_id || !category || !description || !amount) {
    return res.status(400).json({ error: 'site_id, category, description and amount are required' });
  }
  if (!CATEGORIES.includes(category)) return res.status(400).json({ error: `category must be one of: ${CATEGORIES.join(', ')}` });

  const { data, error } = await supabase.from('cst_transactions').insert({
    site_id, category, description, amount,
    transaction_date: transaction_date || new Date().toISOString().slice(0, 10),
    created_by: req.user.id
  }).select().single();

  if (error) return res.status(500).json({ error: error.message });
  res.status(201).json(data);
});

app.get('/api/transactions', requireAuth, async (req, res) => {
  const { status, site_id, category } = req.query;
  let query = supabase.from('cst_transactions').select('*').order('created_at', { ascending: false });
  if (status) query = query.eq('status', status);
  if (site_id) query = query.eq('site_id', site_id);
  if (category) query = query.eq('category', category);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.post('/api/transactions/:id/approve/manager', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { data: txn, error: fe } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fe || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'pending') return res.status(400).json({ error: `Cannot approve a transaction with status "${txn.status}"` });

  const updates = { manager_approved: true, manager_approved_by: req.user.id, manager_approved_at: new Date().toISOString(), updated_at: new Date().toISOString() };
  if (txn.finance_approved) updates.status = 'approved';

  const { data, error } = await supabase.from('cst_transactions').update(updates).eq('id', id).select().single();
  if (error) return res.status(500).json({ error: error.message });
  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'approve_manager' });

  if (data.status === 'approved') {
    await notify(data.created_by, id, 'fully_approved', await renderTemplate('fully_approved', data));
  } else {
    await notify(data.created_by, id, 'approved_manager', await renderTemplate('approved_manager', data));
  }
  res.json(data);
});

app.post('/api/transactions/:id/approve/finance', requireAuth, requireRole('finance', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { data: txn, error: fe } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fe || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'pending') return res.status(400).json({ error: `Cannot approve a transaction with status "${txn.status}"` });

  const updates = { finance_approved: true, finance_approved_by: req.user.id, finance_approved_at: new Date().toISOString(), updated_at: new Date().toISOString() };
  if (txn.manager_approved) updates.status = 'approved';

  const { data, error } = await supabase.from('cst_transactions').update(updates).eq('id', id).select().single();
  if (error) return res.status(500).json({ error: error.message });
  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'approve_finance' });

  if (data.status === 'approved') {
    await notify(data.created_by, id, 'fully_approved', await renderTemplate('fully_approved', data));
  } else {
    await notify(data.created_by, id, 'approved_finance', await renderTemplate('approved_finance', data));
  }
  res.json(data);
});

app.post('/api/transactions/:id/reject', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { reason } = req.body;
  const { data: txn, error: fe } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fe || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'pending') return res.status(400).json({ error: `Cannot reject a transaction with status "${txn.status}"` });

  const { data, error } = await supabase.from('cst_transactions').update({
    status: 'rejected', rejected_by: req.user.id, rejected_at: new Date().toISOString(),
    rejection_reason: reason || null, updated_at: new Date().toISOString()
  }).eq('id', id).select().single();

  if (error) return res.status(500).json({ error: error.message });
  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'reject', note: reason || null });
  await notify(data.created_by, id, 'rejected', await renderTemplate('rejected', data, { reason }));
  res.json(data);
});

app.post('/api/transactions/:id/pay', requireAuth, requireRole('finance', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { data: txn, error: fe } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (fe || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'approved') return res.status(400).json({ error: 'Only fully approved transactions can be marked paid' });

  const { data, error } = await supabase.from('cst_transactions').update({
    status: 'paid', paid_by: req.user.id, paid_at: new Date().toISOString(), updated_at: new Date().toISOString()
  }).eq('id', id).select().single();

  if (error) return res.status(500).json({ error: error.message });
  await supabase.from('cst_approval_log').insert({ transaction_id: id, actor_id: req.user.id, action: 'mark_paid' });

  const receiptMsg = await renderTemplate('paid_receipt', data);
  const recipients = new Set([data.created_by, data.manager_approved_by, data.finance_approved_by].filter(Boolean));
  for (const uid of recipients) await notify(uid, id, 'paid_receipt', receiptMsg);

  res.json(data);
});

// ---------- Manager/admin direct edit & delete (full data control) ----------
app.put('/api/transactions/:id', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { category, description, amount, transaction_date, status } = req.body;

  const updates = { updated_at: new Date().toISOString() };
  if (category) {
    if (!CATEGORIES.includes(category)) return res.status(400).json({ error: `category must be one of: ${CATEGORIES.join(', ')}` });
    updates.category = category;
  }
  if (description) updates.description = description;
  if (amount) updates.amount = amount;
  if (transaction_date) updates.transaction_date = transaction_date;
  if (status) {
    if (!['pending', 'approved', 'rejected', 'paid'].includes(status)) return res.status(400).json({ error: 'Invalid status' });
    updates.status = status;
  }

  const { data, error } = await supabase.from('cst_transactions').update(updates).eq('id', id).select().single();
  if (error) return res.status(500).json({ error: error.message });
  if (!data) return res.status(404).json({ error: 'Transaction not found' });
  res.json(data);
});

app.delete('/api/transactions/:id', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { error } = await supabase.from('cst_transactions').delete().eq('id', id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

// Bulk-clear all transaction data for a site — destructive, requires explicit confirm=yes
app.delete('/api/transactions', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  if (req.query.confirm !== 'yes') return res.status(400).json({ error: 'Add ?confirm=yes to confirm this destructive action' });

  let query = supabase.from('cst_transactions').delete();
  const site_id = req.user.role === 'admin' ? req.query.site_id : req.user.site_id;
  if (site_id) query = query.eq('site_id', site_id);
  else if (req.user.role !== 'admin') return res.status(400).json({ error: 'No site associated with your account' });

  const { error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

// ---------- Excel export ----------
const HEADER = [
  { header: 'Date', key: 'transaction_date', width: 14 },
  { header: 'Category', key: 'category', width: 14 },
  { header: 'Description', key: 'description', width: 40 },
  { header: 'Amount', key: 'amount', width: 14 },
  { header: 'Status', key: 'status', width: 12 },
  { header: 'Created At', key: 'created_at', width: 20 },
  { header: 'Paid At', key: 'paid_at', width: 20 }
];
function styleSheet(sheet) {
  sheet.columns = HEADER;
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
}

app.get('/api/export', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { site_id } = req.query;
  let query = supabase.from('cst_transactions').select('*').order('transaction_date', { ascending: false });
  if (site_id) query = query.eq('site_id', site_id);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Construction Site Transactions';
  workbook.created = new Date();

  const paid = data.filter(t => t.status === 'paid');
  const pending = data.filter(t => t.status === 'pending' || t.status === 'approved');
  const rejected = data.filter(t => t.status === 'rejected');

  const paidSheet = workbook.addWorksheet('Paid'); styleSheet(paidSheet); paidSheet.addRows(paid);
  const pendingSheet = workbook.addWorksheet('Pending'); styleSheet(pendingSheet); pendingSheet.addRows(pending);
  const rejectedSheet = workbook.addWorksheet('Rejected'); styleSheet(rejectedSheet); rejectedSheet.addRows(rejected);

  const summarySheet = workbook.addWorksheet('Summary');
  summarySheet.columns = [{ header: 'Metric', key: 'metric', width: 24 }, { header: 'Value', key: 'value', width: 18 }];
  summarySheet.getRow(1).font = { bold: true };
  summarySheet.addRows([
    { metric: 'Total Paid', value: paid.reduce((s, t) => s + Number(t.amount), 0) },
    { metric: 'Total Pending', value: pending.reduce((s, t) => s + Number(t.amount), 0) },
    { metric: 'Total Rejected', value: rejected.reduce((s, t) => s + Number(t.amount), 0) },
    { metric: 'Count Paid', value: paid.length },
    { metric: 'Count Pending', value: pending.length }
  ]);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="site-transactions-${new Date().toISOString().slice(0, 10)}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
});

// ---------- User management (manager/admin creates clerk & finance logins) ----------
function randomPassword() {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  let out = '';
  for (let i = 0; i < 10; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}

app.post('/api/users', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { full_name, email, role } = req.body;
  if (!full_name || !email || !role) return res.status(400).json({ error: 'full_name, email and role are required' });

  const allowedRoles = req.user.role === 'admin' ? ['clerk', 'manager', 'finance', 'admin'] : ['clerk', 'finance'];
  if (!allowedRoles.includes(role)) return res.status(403).json({ error: `You can only create: ${allowedRoles.join(', ')}` });

  const site_id = req.body.site_id || req.user.site_id;
  const plainPassword = randomPassword();
  const password_hash = bcrypt.hashSync(plainPassword, 10);

  const { data, error } = await supabase.from('cst_users').insert({
    full_name, email: email.toLowerCase().trim(), password_hash, role, site_id
  }).select('id, full_name, email, role, site_id').single();

  if (error) {
    if (error.message.includes('duplicate')) return res.status(409).json({ error: 'A user with this email already exists' });
    return res.status(500).json({ error: error.message });
  }
  // Return the generated password once — the manager must relay it to the new user.
  res.status(201).json({ ...data, temporary_password: plainPassword });
});

app.get('/api/users', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  let query = supabase.from('cst_users').select('id, full_name, email, role, site_id, created_at').order('created_at', { ascending: false });
  if (req.user.role !== 'admin') query = query.eq('site_id', req.user.site_id);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

// ---------- Notifications (template-driven with token substitution) ----------
async function renderTemplate(eventType, txn, extra = {}) {
  const { data: tmpl } = await supabase.from('cst_message_templates').select('template').eq('event_type', eventType).single();
  let site_name = '';
  try {
    const { data: site } = await supabase.from('cst_sites').select('name').eq('id', txn.site_id).single();
    site_name = site ? site.name : '';
  } catch {}
  const tokens = {
    description: txn.description,
    amount: 'KES ' + Number(txn.amount).toLocaleString(undefined, { minimumFractionDigits: 2 }),
    site_name,
    paid_date: txn.paid_at ? new Date(txn.paid_at).toISOString().slice(0, 10) : '',
    reason_suffix: extra.reason ? ' Reason: ' + extra.reason : ''
  };
  let text = tmpl ? tmpl.template : `${eventType}: {{description}} ({{amount}})`;
  for (const [k, v] of Object.entries(tokens)) text = text.split('{{' + k + '}}').join(v || '');
  return text;
}

async function notify(userId, transactionId, type, message) {
  if (!userId) return;
  await supabase.from('cst_notifications').insert({ user_id: userId, transaction_id: transactionId, type, message });
}

app.get('/api/notifications', requireAuth, async (req, res) => {
  const { data, error } = await supabase
    .from('cst_notifications')
    .select('*')
    .eq('user_id', req.user.id)
    .order('created_at', { ascending: false })
    .limit(50);
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.post('/api/notifications/:id/read', requireAuth, async (req, res) => {
  const { error } = await supabase.from('cst_notifications').update({ read_at: new Date().toISOString() }).eq('id', req.params.id).eq('user_id', req.user.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

// ---------- Printable receipt (evidence after payment) ----------
app.get('/api/transactions/:id/receipt', requireAuth, async (req, res) => {
  const { id } = req.params;
  const { data: txn, error } = await supabase.from('cst_transactions').select('*').eq('id', id).single();
  if (error || !txn) return res.status(404).send('Transaction not found');

  const allowed = req.user.role === 'admin' || req.user.id === txn.created_by || req.user.id === txn.manager_approved_by || req.user.id === txn.finance_approved_by || ['manager', 'finance'].includes(req.user.role);
  if (!allowed) return res.status(403).send('Not permitted to view this receipt');
  if (txn.status !== 'paid') return res.status(400).send('Receipt is only available once a transaction is paid');

  res.type('html').send(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Receipt ${id}</title>
  <style>body{font-family:Arial,sans-serif;max-width:480px;margin:40px auto;color:#1f2d3d;}
  .box{border:1px solid #dde1e6;border-radius:10px;padding:24px;}
  h1{font-size:18px;margin:0 0 4px;} .muted{color:#6b7684;font-size:12px;}
  table{width:100%;border-collapse:collapse;margin-top:16px;}
  td{padding:8px 0;border-bottom:1px solid #eee;font-size:14px;}
  td:first-child{color:#6b7684;} .amt{font-size:22px;font-weight:700;margin-top:16px;}
  .stamp{margin-top:20px;padding:8px 12px;background:#dfeee5;color:#2f7d4f;display:inline-block;border-radius:6px;font-weight:700;font-size:13px;}
  @media print{button{display:none;}}</style></head>
  <body><div class="box">
  <h1>Payment Receipt</h1>
  <p class="muted">Transaction ID: ${txn.id}</p>
  <table>
    <tr><td>Date</td><td>${txn.transaction_date}</td></tr>
    <tr><td>Category</td><td>${txn.category}</td></tr>
    <tr><td>Description</td><td>${txn.description}</td></tr>
    <tr><td>Paid At</td><td>${txn.paid_at || ''}</td></tr>
  </table>
  <div class="amt">KES ${Number(txn.amount).toLocaleString(undefined, { minimumFractionDigits: 2 })}</div>
  <div class="stamp">PAID &check;</div>
  <p style="margin-top:24px;"><button onclick="window.print()">Print / Save as PDF</button></p>
  </div></body></html>`);
});

// ---------- Vendors / partners ----------
app.get('/api/vendors', requireAuth, async (req, res) => {
  let query = supabase.from('cst_vendors').select('*').order('name');
  if (req.user.role !== 'admin' && req.user.site_id) query = query.eq('site_id', req.user.site_id);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.post('/api/vendors', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { name, role_type, contact_phone, contact_email } = req.body;
  if (!name) return res.status(400).json({ error: 'name is required' });
  const { data, error } = await supabase.from('cst_vendors').insert({
    name, role_type: role_type || 'vendor', contact_phone, contact_email, site_id: req.user.site_id
  }).select().single();
  if (error) return res.status(500).json({ error: error.message });
  res.status(201).json(data);
});

app.get('/api/transactions/:id/vendors', requireAuth, async (req, res) => {
  const { data, error } = await supabase.from('cst_transaction_vendors').select('id, vendor:vendor_id(id, name, role_type, contact_phone, contact_email)').eq('transaction_id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.post('/api/transactions/:id/vendors', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { vendor_id } = req.body;
  if (!vendor_id) return res.status(400).json({ error: 'vendor_id is required' });
  const { data, error } = await supabase.from('cst_transaction_vendors').insert({ transaction_id: req.params.id, vendor_id }).select().single();
  if (error) {
    if (error.message.includes('duplicate')) return res.status(409).json({ error: 'That vendor is already assigned to this transaction' });
    return res.status(500).json({ error: error.message });
  }
  res.status(201).json(data);
});

app.delete('/api/transaction-vendors/:linkId', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { error } = await supabase.from('cst_transaction_vendors').delete().eq('id', req.params.linkId);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

// ---------- Message templates (custom tokens: {{description}} {{amount}} {{site_name}} {{paid_date}} {{reason_suffix}}) ----------
app.get('/api/templates', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { data, error } = await supabase.from('cst_message_templates').select('*').order('event_type');
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.put('/api/templates/:eventType', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { template } = req.body;
  if (!template) return res.status(400).json({ error: 'template is required' });
  const { data, error } = await supabase.from('cst_message_templates').update({ template, updated_at: new Date().toISOString() }).eq('event_type', req.params.eventType).select().single();
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.get('/api/health', (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
if (require.main === module) {
  app.listen(PORT, () => console.log(`Construction site transactions API running on :${PORT}`));
}

module.exports = app;
