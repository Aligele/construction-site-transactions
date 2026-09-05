require('dotenv').config();
const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');
const ExcelJS = require('exceljs');
const { createClient } = require('@supabase/supabase-js');

const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change-me';
const CATEGORIES = ['materials', 'labor', 'equipment', 'fuel', 'other'];

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY, {
  auth: { autoRefreshToken: false, persistSession: false }
});

// ---------- M-Pesa Daraja B2C integration ----------
const MPESA_ENV = process.env.MPESA_ENV || 'sandbox';
const MPESA_BASE_URL = MPESA_ENV === 'production' ? 'https://api.safaricom.co.ke' : 'https://sandbox.safaricom.co.ke';
let mpesaTokenCache = { token: null, expiresAt: 0 };

function mpesaConfigured() {
  return !!(process.env.MPESA_CONSUMER_KEY && process.env.MPESA_CONSUMER_SECRET && process.env.MPESA_SECURITY_CREDENTIAL);
}

async function getMpesaToken() {
  if (mpesaTokenCache.token && Date.now() < mpesaTokenCache.expiresAt) return mpesaTokenCache.token;
  const auth = Buffer.from(`${process.env.MPESA_CONSUMER_KEY}:${process.env.MPESA_CONSUMER_SECRET}`).toString('base64');
  const r = await fetch(`${MPESA_BASE_URL}/oauth/v1/generate?grant_type=client_credentials`, {
    headers: { Authorization: `Basic ${auth}` }
  });
  const data = await r.json();
  if (!r.ok || !data.access_token) throw new Error('Failed to get M-Pesa token: ' + JSON.stringify(data));
  mpesaTokenCache = { token: data.access_token, expiresAt: Date.now() + 55 * 60 * 1000 };
  return data.access_token;
}

function normalizePhone(phone) {
  let p = (phone || '').replace(/\s|-/g, '');
  if (p.startsWith('+')) p = p.slice(1);
  if (p.startsWith('0')) p = '254' + p.slice(1);
  if (p.startsWith('7') || p.startsWith('1')) p = '254' + p;
  return p;
}

async function sendB2CPayment({ phone, amount, remarks, occasion, callbackHost }) {
  const token = await getMpesaToken();
  const shortcode = process.env.MPESA_SHORTCODE || '600000';
  const body = {
    OriginatorConversationID: crypto.randomUUID(),
    InitiatorName: process.env.MPESA_INITIATOR_NAME || 'testapi',
    SecurityCredential: process.env.MPESA_SECURITY_CREDENTIAL,
    CommandID: 'BusinessPayment',
    Amount: Math.round(Number(amount)),
    PartyA: shortcode,
    PartyB: normalizePhone(phone),
    Remarks: (remarks || 'Wage payment').slice(0, 100),
    QueueTimeOutURL: `https://${callbackHost}/api/mpesa/b2c/timeout`,
    ResultURL: `https://${callbackHost}/api/mpesa/b2c/result`,
    Occasion: (occasion || '').slice(0, 100)
  };
  const r = await fetch(`${MPESA_BASE_URL}/mpesa/b2c/v3/paymentrequest`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const data = await r.json();
  return { ok: r.ok && data.ResponseCode === '0', status: r.status, data };
}

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
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --sidebar:#16302278;--sidebar-bg:#173B29;--sidebar-active:#2f9e5c;--sidebar-text:#cfe3d6;--sidebar-text-dim:#7fa891;
  --navy:#1f2d3d;--amber:#d98c2b;--green:#2f7d4f;--red:#b03a3a;--bg:#f5f6f8;--card:#fff;--border:#dde1e6;
  --forest:#0f2818;--forest2:#1c4a30;--moss:#7ab98a;--paper:#f6f4ee;
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

/* ---- Premium login page ---- */
.auth-shell{min-height:100vh;background:var(--paper);}
.auth-hero{
  position:relative;overflow:hidden;padding:36px 24px 30px;
  background:
    repeating-linear-gradient(0deg, rgba(255,255,255,.05) 0px, rgba(255,255,255,.05) 1px, transparent 1px, transparent 28px),
    repeating-linear-gradient(90deg, rgba(255,255,255,.05) 0px, rgba(255,255,255,.05) 1px, transparent 1px, transparent 28px),
    linear-gradient(160deg, var(--forest) 0%, var(--forest2) 75%);
}
.auth-hero-inner{position:relative;z-index:2;max-width:420px;margin:0 auto;opacity:0;animation:heroIn .6s ease-out forwards;}
@keyframes heroIn{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:translateY(0);}}
.auth-logo{width:52px;height:52px;border-radius:14px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.18);display:flex;align-items:center;justify-content:center;margin-bottom:16px;}
.auth-title{font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:28px;line-height:1.15;color:#fff;margin:0 0 8px;letter-spacing:-0.01em;}
.auth-tagline{font-family:'Inter',sans-serif;color:var(--moss);font-size:14px;line-height:1.5;margin:0 0 20px;max-width:320px;}
.flow-chips{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}
.flow-chip{font-family:'Inter',sans-serif;font-size:12px;font-weight:500;color:#eaf3ec;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.14);padding:5px 11px;border-radius:20px;transition:background .15s,transform .15s;}
.flow-chip:hover{background:rgba(255,255,255,.16);transform:translateY(-1px);}
.flow-arrow{color:var(--moss);font-size:12px;}
.auth-card-wrap{max-width:420px;margin:-18px auto 0;padding:0 24px 40px;position:relative;z-index:3;}
.auth-card{background:#fff;border-radius:14px;padding:26px 22px;box-shadow:0 12px 32px rgba(15,40,24,.14);}
.auth-card h2{font-family:'Space Grotesk',sans-serif;font-size:18px;margin:0 0 4px;color:var(--forest);}
.auth-card .muted{margin-bottom:16px;}
.auth-field{margin-bottom:14px;}
.auth-field label{font-family:'Inter',sans-serif;}
.auth-field input{
  width:100%;box-sizing:border-box;border:1.5px solid var(--border);border-radius:9px;padding:11px 12px;
  font-family:'Inter',sans-serif;font-size:15px;transition:border-color .15s,box-shadow .15s;
}
.auth-field input:focus{outline:none;border-color:var(--moss);box-shadow:0 0 0 3px rgba(122,185,138,.22);}
.auth-submit{
  width:100%;background:var(--forest);color:#fff;border:none;border-radius:9px;padding:13px;
  font-family:'Inter',sans-serif;font-weight:600;font-size:15px;cursor:pointer;transition:background .15s,transform .1s;
}
.auth-submit:hover{background:var(--forest2);}
.auth-submit:active{transform:scale(.98);}
.auth-forgot{display:block;text-align:center;margin-top:14px;font-family:'Inter',sans-serif;font-size:13px;color:var(--forest2);text-decoration:none;}
.auth-forgot:hover{text-decoration:underline;}

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
  appEl.innerHTML =
    '<div class="auth-shell">' +
      '<div class="auth-hero"><div class="auth-hero-inner">' +
        '<div class="auth-logo">'+TRUCK_SVG+'</div>' +
        '<h1 class="auth-title">Site Transactions</h1>' +
        '<p class="auth-tagline">Every shilling on site, tracked from the day it\\'s logged to the day it\\'s paid.</p>' +
        '<div class="flow-chips"><span class="flow-chip">Log</span><span class="flow-arrow">&#8594;</span><span class="flow-chip">Approve</span><span class="flow-arrow">&#8594;</span><span class="flow-chip">Pay</span><span class="flow-arrow">&#8594;</span><span class="flow-chip">Reconcile</span></div>' +
      '</div></div>' +
      '<div class="auth-card-wrap"><div class="auth-card">' +
        '<h2>Sign in</h2><p class="muted">Use the credentials your manager gave you.</p>' +
        '<div class="auth-field"><label>Email</label><input id="email" type="email" placeholder="you@example.com" /></div>' +
        '<div class="auth-field"><label>Password</label><input id="password" type="password" placeholder="********" /></div>' +
        '<button class="auth-submit" id="loginBtn">Log in</button>' +
        '<a href="#" id="forgotLink" class="auth-forgot">Forgot password?</a>' +
        '<div class="error" id="err"></div><div class="muted" id="forgotMsg"></div>' +
      '</div></div>' +
    '</div>';
  document.getElementById('loginBtn').onclick = async () => {
    const email = document.getElementById('email').value.trim();
    const password = document.getElementById('password').value;
    try { const {token,user} = await api('/auth/login',{method:'POST',body:JSON.stringify({email,password})}); save(token,user); render(); }
    catch(e){ document.getElementById('err').textContent = e.message; }
  };
  document.getElementById('forgotLink').onclick = async (ev) => {
    ev.preventDefault();
    const email = document.getElementById('email').value.trim();
    if (!email) { document.getElementById('err').textContent = 'Enter your email above first, then tap "Forgot password?".'; return; }
    try {
      await api('/auth/forgot-password', {method:'POST', body: JSON.stringify({email})});
      document.getElementById('err').textContent = '';
      document.getElementById('forgotMsg').textContent = 'If that email is registered, a reset link has been sent — check your inbox.';
    } catch(e) { document.getElementById('err').textContent = e.message; }
  };
}
function buildOverviewHtml(user){
  const canExport = ['manager','finance','admin'].includes(user.role);
  return '<div class="card"><div class="topbar"><h2>Overview</h2>'+(canExport?'<button class="success" id="overviewExportBtn">Download Excel</button>':'')+'</div><div id="summaryBox" class="muted">Loading...</div></div>';
}
function buildTransactionsHtml(user){
  let html = '';
  if (user.role==='clerk' || user.role==='admin') {
    html += '<div class="card"><h2>Log transactions</h2><p class="muted">Add materials, labor, fuel, etc. as separate line items, then submit them together — the total goes to your manager and finance in one notification.</p><div class="row"><div><label>Site ID</label><input id="site_id" placeholder="site UUID" value="'+(user.site_id||'')+'" /></div></div><div id="itemRows"></div><div style="margin-top:8px;"><button class="secondary" id="addItemBtn" type="button">+ Add item</button></div><div class="topbar" style="margin-top:14px;"><strong>Total: <span id="batchTotal">KES 0.00</span></strong><button id="submitBatch">Submit all</button></div><div class="error" id="submitErr"></div></div>';
  }
  html += '<div class="card"><div class="topbar"><h2>Transactions</h2><div>'+(['manager','finance','admin'].includes(user.role)?'<button id="exportBtn" class="secondary">Download Excel</button>':'')+(['manager','admin'].includes(user.role)?' <button id="clearAllBtn" class="danger">Clear all data</button>':'')+'</div></div>' +
    '<div class="row" style="margin-top:6px;"><input id="filterSearch" placeholder="Search description..." /><select id="filterStatus"><option value="">All statuses</option><option value="pending">Pending</option><option value="approved">Approved</option><option value="paid">Paid</option><option value="rejected">Rejected</option></select><input id="filterFrom" type="date" /><input id="filterTo" type="date" /><button class="secondary" id="filterApply">Filter</button></div>' +
    '<div id="table"></div><div id="pagination" class="row" style="margin-top:10px;justify-content:flex-end;"></div></div>';
  return html;
}
function buildRegistryHtml(user){
  let html = '';
  if (user.role==='clerk' || user.role==='admin') {
    html += '<div class="card"><h2>Enroll a worker</h2><div class="row"><div><label>Full name</label><input id="w_name" placeholder="Full name" /></div><div><label>ID number</label><input id="w_id" placeholder="National ID number" /></div></div><div class="row"><div><label>Phone number</label><input id="w_phone" placeholder="07xxxxxxxx" /></div><div><label>Designation</label><input id="w_designation" placeholder="e.g. Mason, Fundi, Laborer" /></div><div><label>Daily rate (KES)</label><input id="w_rate" type="number" step="0.01" placeholder="e.g. 800" /></div></div><div style="margin-top:12px;"><button id="enrollWorkerBtn">Enroll worker</button></div><div class="error" id="enrollWorkerErr"></div></div>';
    html += '<div class="card"><div class="topbar"><h2>Daily attendance</h2><input id="attDate" type="date" /></div><div id="workerAttendance" class="muted">Loading workers...</div><div style="margin-top:12px;"><button id="submitAttendance">Submit today\\'s attendance</button></div><div class="error" id="attendanceErr"></div><div id="attendanceOk" class="muted"></div></div>';
  }
  if (user.role==='manager' || user.role==='finance') {
    html += '<div class="card"><div class="topbar"><h2>Workers & attendance</h2><input id="attViewDate" type="date" /></div><div id="workerAttendanceView" class="muted">Loading...</div></div>';
  }
  if (user.role==='finance' || user.role==='manager' || user.role==='admin') {
    html += '<div class="card"><div class="topbar"><h2>Weekly wages</h2><div><label style="display:inline;margin-right:6px;">Week starting</label><input id="wagesWeekStart" type="date" /></div></div><div id="weeklyWages" class="muted">Loading...</div><div style="margin-top:10px;text-align:left;">'+(['finance','admin'].includes(user.role)?'<button class="success" id="payAllWagesBtn">Pay all wages</button> ':'')+'<button class="success" id="downloadWagesBtn">Download Excel</button></div><div class="error" id="payWagesErr"></div><div class="muted" id="payWagesOk"></div></div>';
  }
  return html;
}
function buildNotificationsHtml(user){
  return '<div class="card"><div class="topbar"><h2>Notifications</h2><span class="muted" id="notifCount"></span></div><div id="notifList" class="muted">Loading...</div></div>';
}
function buildSettingsHtml(user){
  let html = '<div class="card"><h2>Change password</h2><div class="row"><div><label>Current password</label><input id="curPw" type="password" /></div><div><label>New password</label><input id="newPw" type="password" /></div></div><div style="margin-top:12px;"><button id="changePwBtn">Update password</button></div><div class="error" id="changePwErr"></div><div id="changePwOk" class="muted"></div></div>';
  if (user.role==='manager' || user.role==='admin') {
    html += '<div class="card"><h2>Add a clerk or finance login</h2><div class="row"><div><label>Full name</label><input id="new_name" placeholder="Full name" /></div><div><label>Email</label><input id="new_email" type="email" placeholder="person@example.com" /></div><div><label>Role</label><select id="new_role"><option value="clerk">Clerk</option><option value="finance">Finance</option></select></div></div><div style="margin-top:12px;"><button id="createUserBtn">Create login</button></div><div class="error" id="createUserErr"></div><div id="createUserResult"></div></div>';
    html += '<div class="card"><h2>Team logins</h2><div id="userList" class="muted">Loading...</div></div>';
  }
  return html;
}

async function renderDashboard(){
  const {user} = state;
  appEl.innerHTML =
    '<div class="shell">' +
      '<div class="overlay" id="overlay"></div>' +
      '<aside class="sidebar" id="sidebar">' +
        '<div class="brand"><div class="brand-logo">'+TRUCK_SVG+'</div><div><div class="brand-title">Site Transactions</div><div class="brand-sub">Construction Portal</div></div></div>' +
        '<div class="nav-section-label">MAIN</div>' +
        '<a class="nav-item active" id="navDashboard" data-view="dashboard"><span class="nav-icon">&#8962;</span> Dashboard</a>' +
        '<a class="nav-item" id="navTransactions" data-view="transactions"><span class="nav-icon">&#128203;</span> Transactions</a>' +
        (['clerk','finance','manager','admin'].includes(user.role) ? '<a class="nav-item" id="navRegistry" data-view="registry"><span class="nav-icon">&#128101;</span> Registry</a>' : '') +
        '<div class="nav-section-label">ACCOUNT</div>' +
        '<a class="nav-item" id="navNotifications" data-view="notifications"><span class="nav-icon">&#128276;</span> Notifications</a>' +
        '<a class="nav-item" id="navSettings" data-view="settings"><span class="nav-icon">&#9881;</span> Settings</a>' +
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
  function setActiveNav(navEl){ document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active')); if (navEl) navEl.classList.add('active'); }
  function closeSidebar(){ sidebar.classList.remove('open'); overlay.classList.remove('open'); }

  function switchView(view, navEl){
    setActiveNav(navEl);
    const content = document.getElementById('content');
    let html = '';
    if (view === 'dashboard') html = buildOverviewHtml(user);
    else if (view === 'transactions') html = buildTransactionsHtml(user);
    else if (view === 'registry') html = buildRegistryHtml(user);
    else if (view === 'notifications') html = buildNotificationsHtml(user);
    else if (view === 'settings') html = buildSettingsHtml(user);
    content.innerHTML = html;
    wireContent();
    closeSidebar();
  }

  document.querySelectorAll('[data-view]').forEach(navEl => {
    navEl.onclick = () => switchView(navEl.dataset.view, navEl);
  });
  document.getElementById('bellIcon').onclick = () => switchView('notifications', document.getElementById('navNotifications'));

  function wireContent(){
    if (document.getElementById('summaryBox')) loadSummary();
    if (document.getElementById('notifList')) loadNotifications();
    if (document.getElementById('userList')) loadUsers();
    if (document.getElementById('workerAttendance')) {
      const dateInput = document.getElementById('attDate');
      dateInput.value = new Date().toISOString().slice(0,10);
      dateInput.onchange = loadWorkerAttendance;
      loadWorkerAttendance();
    }
    if (document.getElementById('workerAttendanceView')) {
      const dateInput = document.getElementById('attViewDate');
      dateInput.value = new Date().toISOString().slice(0,10);
      dateInput.onchange = loadWorkerAttendanceView;
      loadWorkerAttendanceView();
    }
    if (document.getElementById('weeklyWages')) {
      const wsInput = document.getElementById('wagesWeekStart');
      const today = new Date();
      const day = today.getDay();
      const diffToMonday = day === 0 ? 6 : day - 1;
      const monday = new Date(today);
      monday.setDate(today.getDate() - diffToMonday);
      wsInput.value = monday.toISOString().slice(0,10);
      wsInput.onchange = loadWeeklyWages;
      loadWeeklyWages();
      document.getElementById('downloadWagesBtn').onclick = () => {
        const start = document.getElementById('wagesWeekStart').value;
        fetch(API+'/wages/weekly/export?start='+start, {headers:{Authorization:'Bearer '+state.token}})
          .then(r=>r.blob()).then(blob=>{
            const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url;
            a.download = 'weekly-wages-'+start+'.xlsx'; a.click();
          });
      };
      if (document.getElementById('payAllWagesBtn')) {
        document.getElementById('payAllWagesBtn').onclick = async () => {
          const start = document.getElementById('wagesWeekStart').value;
          if (!confirm('Pay all outstanding wages for the week starting '+start+'? This settles every worker who hasn\\'t been paid yet for that week.')) return;
          try {
            const result = await api('/wages/weekly/pay-all', {method:'POST', body: JSON.stringify({start})});
            document.getElementById('payWagesErr').textContent = '';
            if (result.paid_count > 0) {
              let msg = 'Recorded payment for '+result.paid_count+' worker(s), total '+money(result.total_amount)+'.';
              if (result.mpesa_enabled) {
                const sent = (result.results||[]).filter(r=>r.disbursement_status==='sent').length;
                const failed = (result.results||[]).filter(r=>r.disbursement_status==='failed').length;
                msg += ' M-Pesa: '+sent+' sent, '+failed+' failed.';
                const failedNames = (result.results||[]).filter(r=>r.disbursement_status==='failed').map(r=>r.name+' ('+(r.mpesa_result_desc||'error')+')');
                if (failedNames.length) msg += ' Failed: '+failedNames.join('; ');
              } else {
                msg += ' (M-Pesa not configured — recorded as paid manually.)';
              }
              document.getElementById('payWagesOk').textContent = msg;
            } else {
              document.getElementById('payWagesOk').textContent = result.message || 'Nothing to pay.';
            }
            loadWeeklyWages(); loadSummary(); if (document.getElementById('table')) loadTable();
          } catch(e) { document.getElementById('payWagesErr').textContent = e.message; document.getElementById('payWagesOk').textContent=''; }
        };
      }
    }

    if (document.getElementById('enrollWorkerBtn')) {
      document.getElementById('enrollWorkerBtn').onclick = async () => {
        const name = document.getElementById('w_name').value.trim();
        const id_number = document.getElementById('w_id').value.trim();
        const phone_number = document.getElementById('w_phone').value.trim();
        const designation = document.getElementById('w_designation').value.trim();
        const daily_rate = document.getElementById('w_rate').value;
        try {
          await api('/workers', {method:'POST', body: JSON.stringify({name, id_number, phone_number, designation, daily_rate})});
          document.getElementById('enrollWorkerErr').textContent = '';
          document.getElementById('w_name').value=''; document.getElementById('w_id').value=''; document.getElementById('w_phone').value=''; document.getElementById('w_designation').value=''; document.getElementById('w_rate').value='';
          loadWorkerAttendance();
        } catch(e) { document.getElementById('enrollWorkerErr').textContent = e.message; }
      };
    }
    if (document.getElementById('submitAttendance')) {
      document.getElementById('submitAttendance').onclick = async () => {
        const attendance_date = document.getElementById('attDate').value;
        const worker_ids = Array.from(document.querySelectorAll('[data-worker-check]:checked')).map(cb => cb.dataset.workerCheck);
        if (worker_ids.length === 0) { document.getElementById('attendanceErr').textContent = 'Select at least one worker.'; return; }
        try {
          const result = await api('/attendance', {method:'POST', body: JSON.stringify({attendance_date, worker_ids})});
          document.getElementById('attendanceErr').textContent = '';
          document.getElementById('attendanceOk').textContent = 'Attendance recorded for '+result.recorded+' worker(s) on '+result.date+'.';
        } catch(e) { document.getElementById('attendanceErr').textContent = e.message; document.getElementById('attendanceOk').textContent=''; }
      };
    }

    if (document.getElementById('changePwBtn')) {
      document.getElementById('changePwBtn').onclick = async () => {
        const current_password = document.getElementById('curPw').value;
        const new_password = document.getElementById('newPw').value;
        try {
          await api('/auth/change-password', {method:'POST', body: JSON.stringify({current_password, new_password})});
          document.getElementById('changePwErr').textContent = '';
          document.getElementById('changePwOk').textContent = 'Password updated.';
          document.getElementById('curPw').value=''; document.getElementById('newPw').value='';
        } catch(e) { document.getElementById('changePwErr').textContent = e.message; document.getElementById('changePwOk').textContent=''; }
      };
    }

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
          loadUsers();
        } catch(e) { document.getElementById('createUserErr').textContent = e.message; }
      };
    }

    if (document.getElementById('itemRows')) {
      let itemCounter = 0;
      function itemRowHtml(n){
        return '<div class="row" id="itemRow-'+n+'" style="align-items:flex-end;">' +
          '<div><label>Category</label><select id="itemCat-'+n+'"><option value="materials">Materials</option><option value="labor">Labor</option><option value="equipment">Equipment</option><option value="fuel">Fuel</option><option value="other">Other</option></select></div>' +
          '<div style="flex:2;"><label>Description</label><input id="itemDesc-'+n+'" placeholder="What was this for?" /></div>' +
          '<div><label>Amount (KES)</label><input id="itemAmt-'+n+'" type="number" step="0.01" data-item-amt /></div>' +
          '<div><button class="danger" type="button" data-remove-item="'+n+'" style="padding:9px 10px;">&times;</button></div>' +
        '</div>';
      }
      function addItemRow(){
        itemCounter++;
        document.getElementById('itemRows').insertAdjacentHTML('beforeend', itemRowHtml(itemCounter));
        const removeBtn = document.querySelector('[data-remove-item="'+itemCounter+'"]');
        removeBtn.onclick = () => { document.getElementById('itemRow-'+itemCounter).remove(); updateBatchTotal(); };
        document.getElementById('itemAmt-'+itemCounter).addEventListener('input', updateBatchTotal);
      }
      function updateBatchTotal(){
        let total = 0;
        document.querySelectorAll('[data-item-amt]').forEach(inp => { total += Number(inp.value) || 0; });
        document.getElementById('batchTotal').textContent = money(total);
      }
      document.getElementById('addItemBtn').onclick = addItemRow;
      addItemRow(); // start with one row

      document.getElementById('submitBatch').onclick = async () => {
        const site_id = document.getElementById('site_id').value.trim();
        const rows = document.querySelectorAll('#itemRows > div');
        const items = [];
        let hasError = false;
        rows.forEach(row => {
          const n = row.id.replace('itemRow-','');
          const category = document.getElementById('itemCat-'+n).value;
          const description = document.getElementById('itemDesc-'+n).value.trim();
          const amount = document.getElementById('itemAmt-'+n).value;
          if (!description || !amount) { hasError = true; return; }
          items.push({category, description, amount});
        });
        if (!site_id) { document.getElementById('submitErr').textContent = 'Site ID is required.'; return; }
        if (hasError || items.length === 0) { document.getElementById('submitErr').textContent = 'Fill in description and amount for every item.'; return; }
        try {
          await api('/batches', {method:'POST', body: JSON.stringify({site_id, items})});
          document.getElementById('submitErr').textContent = '';
          document.getElementById('itemRows').innerHTML = '';
          itemCounter = 0;
          addItemRow();
          updateBatchTotal();
          if (document.getElementById('table')) loadTable();
          loadSummary();
        } catch(e){ document.getElementById('submitErr').textContent = e.message; }
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
    if (document.getElementById('overviewExportBtn')) {
      document.getElementById('overviewExportBtn').onclick = () => {
        fetch(API+'/export', {headers:{Authorization:'Bearer '+state.token}}).then(r=>r.blob()).then(blob=>{
          const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url;
          a.download = 'overview-'+new Date().toISOString().slice(0,10)+'.xlsx'; a.click();
        });
      };
    }
    if (document.getElementById('clearAllBtn')) {
      document.getElementById('clearAllBtn').onclick = async () => {
        if (!confirm('This will permanently delete ALL transaction data for your site. Are you sure?')) return;
        if (!confirm('Really sure? This cannot be undone.')) return;
        try { await api('/transactions?confirm=yes', {method:'DELETE'}); loadTable(); loadSummary(); }
        catch(e){ alert(e.message); }
      };
    }
    if (document.getElementById('filterApply')) {
      document.getElementById('filterApply').onclick = () => { tablePage = 0; loadTable(); };
    }
    if (document.getElementById('table')) loadTable();
  }

  switchView('dashboard', document.getElementById('navDashboard'));
}
let tablePage = 0;
const PAGE_SIZE = 20;
async function loadTable(){
  const {user} = state;
  const search = document.getElementById('filterSearch') ? document.getElementById('filterSearch').value.trim() : '';
  const status = document.getElementById('filterStatus') ? document.getElementById('filterStatus').value : '';
  const from_date = document.getElementById('filterFrom') ? document.getElementById('filterFrom').value : '';
  const to_date = document.getElementById('filterTo') ? document.getElementById('filterTo').value : '';
  const params = new URLSearchParams({ limit: PAGE_SIZE, offset: tablePage*PAGE_SIZE });
  if (search) params.set('search', search);
  if (status) params.set('status', status);
  if (from_date) params.set('from_date', from_date);
  if (to_date) params.set('to_date', to_date);
  const { rows, total } = await api('/transactions?'+params.toString());
  let html = '<table><thead><tr><th>Date</th><th>Category</th><th>Description</th><th>Amount</th><th>Status</th><th>Actions</th></tr></thead><tbody>';
  rows.forEach(t => {
    html += '<tr><td>'+t.transaction_date+'</td><td>'+t.category+'</td><td>'+t.description+'</td><td>'+money(t.amount)+'</td><td>'+badge(t.status)+'</td><td class="actions">'+actionButtons(t,user)+' <button class="secondary" data-details="'+t.id+'">Details</button></td></tr>';
    html += '<tr id="details-'+t.id+'" style="display:none;"><td colspan="6"><div id="details-body-'+t.id+'" class="muted">Loading...</div></td></tr>';
    html += '<tr id="edit-'+t.id+'" style="display:none;"><td colspan="6"><div class="row"><div><label>Description</label><input id="editDesc-'+t.id+'" value="'+t.description.replace(/"/g,'&quot;')+'" /></div><div><label>Amount</label><input id="editAmount-'+t.id+'" type="number" step="0.01" value="'+t.amount+'" /></div><div><label>Status</label><select id="editStatus-'+t.id+'"><option value="pending"'+(t.status==='pending'?' selected':'')+'>Pending</option><option value="approved"'+(t.status==='approved'?' selected':'')+'>Approved</option><option value="rejected"'+(t.status==='rejected'?' selected':'')+'>Rejected</option><option value="paid"'+(t.status==='paid'?' selected':'')+'>Paid</option></select></div></div><div style="margin-top:8px;"><button data-save-edit="'+t.id+'">Save</button> <button class="secondary" data-cancel-edit="'+t.id+'">Cancel</button></div></td></tr>';
  });
  html += '</tbody></table>';
  document.getElementById('table').innerHTML = html;

  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  document.getElementById('pagination').innerHTML = '<span class="muted" style="align-self:center;">Page '+(tablePage+1)+' of '+totalPages+' ('+total+' total)</span><button class="secondary" id="prevPage" '+(tablePage===0?'disabled':'')+'>Prev</button><button class="secondary" id="nextPage" '+(tablePage>=totalPages-1?'disabled':'')+'>Next</button>';
  const prevBtn = document.getElementById('prevPage'); if (prevBtn) prevBtn.onclick = () => { if (tablePage>0){ tablePage--; loadTable(); } };
  const nextBtn = document.getElementById('nextPage'); if (nextBtn) nextBtn.onclick = () => { if (tablePage<totalPages-1){ tablePage++; loadTable(); } };

  document.querySelectorAll('[data-action]').forEach(btn => {
    btn.onclick = async () => {
      const {action,id} = btn.dataset;
      try { await api('/transactions/'+id+'/'+action,{method:'POST',body:JSON.stringify({})}); loadTable(); loadSummary(); }
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
      try { await api('/transactions/'+btn.dataset.delete, {method:'DELETE'}); loadTable(); loadSummary(); }
      catch(e){ alert(e.message); }
    };
  });
  document.querySelectorAll('[data-details]').forEach(btn => {
    btn.onclick = () => toggleDetails(btn.dataset.details);
  });
  document.querySelectorAll('[data-edit]').forEach(btn => {
    btn.onclick = () => { const row = document.getElementById('edit-'+btn.dataset.edit); row.style.display = row.style.display==='none' ? 'table-row' : 'none'; };
  });
  document.querySelectorAll('[data-cancel-edit]').forEach(btn => {
    btn.onclick = () => { document.getElementById('edit-'+btn.dataset.cancelEdit).style.display = 'none'; };
  });
  document.querySelectorAll('[data-save-edit]').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.saveEdit;
      const description = document.getElementById('editDesc-'+id).value;
      const amount = document.getElementById('editAmount-'+id).value;
      const status = document.getElementById('editStatus-'+id).value;
      try { await api('/transactions/'+id, {method:'PUT', body: JSON.stringify({description, amount, status})}); loadTable(); loadSummary(); }
      catch(e){ alert(e.message); }
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
    const [vendorLinks, allVendors, batchItems] = await Promise.all([
      api('/transactions/'+txnId+'/vendors'),
      api('/vendors'),
      api('/transactions/'+txnId+'/batch-items')
    ]);
    const canManage = ['manager','finance','admin'].includes(state.user.role);
    let html = '';
    if (batchItems.length) {
      html += '<strong>Item breakdown</strong><ul style="padding-left:18px;margin:6px 0;">' + batchItems.map(i =>
        '<li>'+i.category+': '+i.description+' — '+money(i.amount)+'</li>'
      ).join('') + '</ul>';
    }
    html += '<strong>Assigned partners/vendors</strong>';
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
async function loadSummary(){
  const el = document.getElementById('summaryBox');
  try {
    const s = await api('/summary');
    const cards = [
      {label:'Pending', amt:s.pending, count:s.pending_count, color:'#d98c2b'},
      {label:'Approved', amt:s.approved, count:s.approved_count, color:'#2f7d4f'},
      {label:'Paid', amt:s.paid, count:s.paid_count, color:'#2f7d4f'},
      {label:'Rejected', amt:s.rejected, count:s.rejected_count, color:'#b03a3a'}
    ];
    el.innerHTML = '<div class="row">' + cards.map(c =>
      '<div style="background:#f5f6f8;border-radius:8px;padding:12px;"><div class="muted">'+c.label+' ('+c.count+')</div><div style="font-size:18px;font-weight:700;color:'+c.color+';">'+money(c.amt)+'</div></div>'
    ).join('') + '</div>';
  } catch(e) { el.textContent = 'Could not load summary.'; }
}
async function loadWeeklyWages(){
  const el = document.getElementById('weeklyWages');
  const start = document.getElementById('wagesWeekStart').value;
  try {
    const { start: s, end, rows } = await api('/wages/weekly?start='+start);
    if (!rows.length) { el.textContent = 'No workers enrolled yet.'; return; }
    const grandTotal = rows.reduce((sum, r) => sum + r.total_pay, 0);
    const outstanding = rows.filter(r => !r.paid).reduce((sum, r) => sum + r.total_pay, 0);
    el.innerHTML = '<p class="muted">Week: '+s+' to '+end+'</p><table><thead><tr><th>Name</th><th>Designation</th><th>Daily Rate</th><th>Days Present</th><th>Total Pay</th><th>Status</th></tr></thead><tbody>' +
      rows.map(r => '<tr><td>'+r.name+'</td><td>'+r.designation+'</td><td>'+money(r.daily_rate)+'</td><td>'+r.days_present+'</td><td><strong>'+money(r.total_pay)+'</strong></td><td>'+(r.paid?'<span class="badge paid">paid</span>':(r.total_pay>0?'<span class="badge pending">outstanding</span>':'<span class="muted">—</span>'))+'</td></tr>').join('') +
      '</tbody></table><p style="text-align:right;margin-top:8px;"><strong>Grand total: '+money(grandTotal)+'</strong><br/><span class="muted">Outstanding: '+money(outstanding)+'</span></p>';
  } catch(e) { el.textContent = 'Could not load wages.'; }
}
async function loadWorkerAttendanceView(){
  const el = document.getElementById('workerAttendanceView');
  const date = document.getElementById('attViewDate').value;
  try {
    const [workers, marked] = await Promise.all([ api('/workers'), api('/attendance?date='+date) ]);
    if (!workers.length) { el.textContent = 'No workers enrolled yet.'; return; }
    const markedIds = new Set(marked.map(m => m.worker_id));
    el.innerHTML = '<p class="muted">'+marked.length+' of '+workers.length+' present on '+date+'</p><table><thead><tr><th>Name</th><th>ID number</th><th>Designation</th><th>Status</th></tr></thead><tbody>' +
      workers.map(w => '<tr><td>'+w.name+'</td><td>'+w.id_number+'</td><td>'+w.designation+'</td><td>'+(markedIds.has(w.id)?'<span class="badge paid">present</span>':'<span class="badge rejected">absent</span>')+'</td></tr>').join('') +
      '</tbody></table>';
  } catch(e) { el.textContent = 'Could not load attendance.'; }
}
async function loadWorkerAttendance(){
  const el = document.getElementById('workerAttendance');
  const date = document.getElementById('attDate').value;
  try {
    const [workers, marked] = await Promise.all([ api('/workers'), api('/attendance?date='+date) ]);
    if (!workers.length) { el.textContent = 'No workers enrolled yet — add one above.'; return; }
    const markedIds = new Set(marked.map(m => m.worker_id));
    el.innerHTML = '<table><thead><tr><th></th><th>Name</th><th>ID number</th><th>Designation</th><th></th></tr></thead><tbody>' +
      workers.map(w => '<tr><td><input type="checkbox" data-worker-check="'+w.id+'" '+(markedIds.has(w.id)?'checked':'')+' /></td><td>'+w.name+'</td><td>'+w.id_number+'</td><td>'+w.designation+'</td><td><button class="danger" style="padding:4px 8px;font-size:12px;" data-remove-worker="'+w.id+'">remove</button></td></tr>').join('') +
      '</tbody></table>';
    el.querySelectorAll('[data-remove-worker]').forEach(btn => {
      btn.onclick = async () => {
        if (!confirm('Remove this worker from the registry?')) return;
        try { await api('/workers/'+btn.dataset.removeWorker, {method:'DELETE'}); loadWorkerAttendance(); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) { el.textContent = 'Could not load workers.'; }
}
async function loadUsers(){
  const el = document.getElementById('userList');
  try {
    const users = await api('/users');
    if (!users.length) { el.textContent = 'No team logins yet.'; return; }
    el.innerHTML = '<table><thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead><tbody>' +
      users.map(u => '<tr><td>'+u.full_name+'</td><td>'+u.email+'</td><td style="text-transform:capitalize;">'+u.role+'</td><td>'+(u.is_active===false?'<span class="badge rejected">inactive</span>':'<span class="badge paid">active</span>')+'</td><td class="actions">'+(u.id!==state.user.id?'<button class="'+(u.is_active===false?'success':'danger')+'" data-toggle-user="'+u.id+'" data-next="'+(u.is_active===false?'true':'false')+'">'+(u.is_active===false?'Activate':'Deactivate')+'</button>':'<span class="muted">you</span>')+'</td></tr>').join('') +
      '</tbody></table>';
    el.querySelectorAll('[data-toggle-user]').forEach(btn => {
      btn.onclick = async () => {
        try { await api('/users/'+btn.dataset.toggleUser+'/status', {method:'PATCH', body: JSON.stringify({is_active: btn.dataset.next==='true'})}); loadUsers(); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) { el.textContent = 'Could not load team logins.'; }
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
const MAX_FAILED_ATTEMPTS = 5;
const LOCKOUT_MINUTES = 15;

app.post('/api/auth/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ error: 'Email and password required' });

  const { data: user, error } = await supabase.from('cst_users').select('*').eq('email', email.toLowerCase().trim()).single();
  if (error || !user) return res.status(401).json({ error: 'Invalid email or password' });

  if (user.is_active === false) return res.status(403).json({ error: 'This account has been deactivated. Contact your manager.' });

  if (user.locked_until && new Date(user.locked_until) > new Date()) {
    const mins = Math.ceil((new Date(user.locked_until) - new Date()) / 60000);
    return res.status(423).json({ error: `Too many failed attempts. Try again in ${mins} minute(s).` });
  }

  if (!bcrypt.compareSync(password, user.password_hash)) {
    const attempts = (user.failed_attempts || 0) + 1;
    const updates = { failed_attempts: attempts };
    if (attempts >= MAX_FAILED_ATTEMPTS) {
      updates.locked_until = new Date(Date.now() + LOCKOUT_MINUTES * 60000).toISOString();
      updates.failed_attempts = 0;
    }
    await supabase.from('cst_users').update(updates).eq('id', user.id);
    if (updates.locked_until) return res.status(423).json({ error: `Too many failed attempts. Account locked for ${LOCKOUT_MINUTES} minutes.` });
    return res.status(401).json({ error: 'Invalid email or password' });
  }

  if (user.failed_attempts || user.locked_until) {
    await supabase.from('cst_users').update({ failed_attempts: 0, locked_until: null }).eq('id', user.id);
  }

  const token = signToken(user);
  res.json({ token, user: { id: user.id, full_name: user.full_name, email: user.email, role: user.role, site_id: user.site_id } });
});

app.post('/api/auth/change-password', requireAuth, async (req, res) => {
  const { current_password, new_password } = req.body;
  if (!current_password || !new_password) return res.status(400).json({ error: 'current_password and new_password are required' });
  if (new_password.length < 8) return res.status(400).json({ error: 'New password must be at least 8 characters' });

  const { data: user, error } = await supabase.from('cst_users').select('*').eq('id', req.user.id).single();
  if (error || !user || !bcrypt.compareSync(current_password, user.password_hash)) {
    return res.status(401).json({ error: 'Current password is incorrect' });
  }
  const { error: updErr } = await supabase.from('cst_users').update({ password_hash: bcrypt.hashSync(new_password, 10) }).eq('id', user.id);
  if (updErr) return res.status(500).json({ error: updErr.message });
  res.json({ ok: true });
});

// ---------- Forgot / reset password ----------
app.post('/api/auth/forgot-password', async (req, res) => {
  const { email } = req.body;
  if (!email) return res.status(400).json({ error: 'email is required' });

  const { data: user } = await supabase.from('cst_users').select('id, email, full_name').eq('email', email.toLowerCase().trim()).single();
  // Always respond the same way, whether or not the email exists, to avoid leaking which emails are registered.
  if (user) {
    const rawToken = crypto.randomBytes(32).toString('hex');
    const tokenHash = crypto.createHash('sha256').update(rawToken).digest('hex');
    await supabase.from('cst_users').update({
      reset_token_hash: tokenHash,
      reset_token_expires: new Date(Date.now() + 30 * 60000).toISOString()
    }).eq('id', user.id);

    const proto = req.headers['x-forwarded-proto'] || req.protocol;
    const resetUrl = `${proto}://${req.get('host')}/reset?token=${rawToken}`;
    await sendEmail(user.email, 'Reset your Site Transactions password',
      `Hi ${user.full_name},\n\nClick the link below to reset your password. This link expires in 30 minutes.\n\n${resetUrl}\n\nIf you didn't request this, you can ignore this email.`);
  }
  res.json({ ok: true, message: 'If that email is registered, a reset link has been sent.' });
});

app.post('/api/auth/reset-password', async (req, res) => {
  const { token, new_password } = req.body;
  if (!token || !new_password) return res.status(400).json({ error: 'token and new_password are required' });
  if (new_password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters' });

  const tokenHash = crypto.createHash('sha256').update(token).digest('hex');
  const { data: user, error } = await supabase.from('cst_users').select('*').eq('reset_token_hash', tokenHash).single();
  if (error || !user || !user.reset_token_expires || new Date(user.reset_token_expires) < new Date()) {
    return res.status(400).json({ error: 'This reset link is invalid or has expired. Request a new one.' });
  }

  await supabase.from('cst_users').update({
    password_hash: bcrypt.hashSync(new_password, 10),
    reset_token_hash: null,
    reset_token_expires: null,
    failed_attempts: 0,
    locked_until: null
  }).eq('id', user.id);

  res.json({ ok: true });
});

app.get('/reset', (req, res) => {
  const token = (req.query.token || '').replace(/[^a-f0-9]/gi, '');
  res.type('html').send(`<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Reset password</title>
  <style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:#f5f6f8;color:#1f2d3d;margin:0;}
  .card{max-width:360px;margin:60px auto;background:#fff;border:1px solid #dde1e6;border-radius:10px;padding:20px;}
  label{display:block;font-size:13px;margin:10px 0 4px;font-weight:600;}
  input{width:100%;font-size:14px;padding:9px 10px;border-radius:6px;border:1px solid #dde1e6;box-sizing:border-box;}
  button{width:100%;margin-top:16px;background:#1f2d3d;color:#fff;border:none;padding:10px;border-radius:6px;font-weight:600;cursor:pointer;}
  .error{color:#b03a3a;font-size:13px;margin-top:8px;} .ok{color:#2f7d4f;font-size:13px;margin-top:8px;}</style></head>
  <body><div class="card"><h2>Reset your password</h2>
  <label>New password</label><input id="pw" type="password" placeholder="At least 8 characters" />
  <button id="btn">Set new password</button>
  <div class="error" id="err"></div><div class="ok" id="ok"></div>
  <script>
  document.getElementById('btn').onclick = async () => {
    const new_password = document.getElementById('pw').value;
    try {
      const res = await fetch('/api/auth/reset-password', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({token:'${token}', new_password})});
      const body = await res.json();
      if (!res.ok) throw new Error(body.error || 'Could not reset password');
      document.getElementById('err').textContent='';
      document.getElementById('ok').textContent='Password updated. You can close this page and log in.';
    } catch(e) { document.getElementById('err').textContent = e.message; }
  };
  </script></div></body></html>`);
});

// ---------- Batch submission (multiple line items at once, e.g. materials + labor + fuel) ----------
app.post('/api/batches', requireAuth, requireRole('clerk', 'admin'), async (req, res) => {
  const { site_id, items } = req.body;
  if (!site_id || !Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: 'site_id and a non-empty items array are required' });
  }
  for (const item of items) {
    if (!item.category || !item.description || !item.amount) {
      return res.status(400).json({ error: 'Each item needs category, description and amount' });
    }
    if (!CATEGORIES.includes(item.category)) return res.status(400).json({ error: `category must be one of: ${CATEGORIES.join(', ')}` });
  }

  const total = items.reduce((s, i) => s + Number(i.amount), 0);
  const { data: batch, error: batchErr } = await supabase.from('cst_batches').insert({
    site_id, created_by: req.user.id, item_count: items.length, total_amount: total
  }).select().single();
  if (batchErr) return res.status(500).json({ error: batchErr.message });

  const itemRows = items.map(i => ({ batch_id: batch.id, category: i.category, description: i.description, amount: i.amount }));
  const { error: itemErr } = await supabase.from('cst_batch_items').insert(itemRows);
  if (itemErr) return res.status(500).json({ error: itemErr.message });

  // Single transaction covering the whole batch — this is what manager/finance see and act on.
  const uniqueCategories = [...new Set(items.map(i => i.category))];
  const combinedCategory = uniqueCategories.length === 1 ? uniqueCategories[0] : 'mixed';
  const combinedDescription = items.map(i => `${i.category}: ${i.description} (KES ${Number(i.amount).toLocaleString()})`).join('; ');

  const { data: txn, error: txnErr } = await supabase.from('cst_transactions').insert({
    site_id, category: combinedCategory, description: combinedDescription, amount: total,
    transaction_date: new Date().toISOString().slice(0, 10),
    created_by: req.user.id, batch_id: batch.id
  }).select().single();
  if (txnErr) return res.status(500).json({ error: txnErr.message });

  // Notify every manager and finance user at this site with the combined total.
  const { data: recipients } = await supabase.from('cst_users').select('id').eq('site_id', site_id).in('role', ['manager', 'finance']);
  const msg = `New transaction from a clerk: ${items.length} items totaling KES ${total.toLocaleString(undefined, { minimumFractionDigits: 2 })}. Breakdown: ${combinedDescription}`;
  for (const r of (recipients || [])) await notify(r.id, txn.id, 'batch_submitted', msg);

  res.status(201).json({ batch, transaction: txn });
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
  const { status, site_id, category, search, from_date, to_date, limit, offset } = req.query;
  let query = supabase.from('cst_transactions').select('*', { count: 'exact' }).order('created_at', { ascending: false });
  if (status) query = query.eq('status', status);
  if (site_id) query = query.eq('site_id', site_id);
  if (category) query = query.eq('category', category);
  if (search) query = query.ilike('description', `%${search}%`);
  if (from_date) query = query.gte('transaction_date', from_date);
  if (to_date) query = query.lte('transaction_date', to_date);
  const lim = Math.min(parseInt(limit) || 50, 200);
  const off = parseInt(offset) || 0;
  query = query.range(off, off + lim - 1);
  const { data, error, count } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json({ rows: data, total: count });
});

app.get('/api/summary', requireAuth, async (req, res) => {
  let query = supabase.from('cst_transactions').select('status, amount, category');
  if (req.user.role !== 'admin' && req.user.site_id) query = query.eq('site_id', req.user.site_id);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });

  const summary = { pending: 0, approved: 0, paid: 0, rejected: 0, pending_count: 0, approved_count: 0, paid_count: 0, rejected_count: 0, by_category: {} };
  for (const t of data) {
    const amt = Number(t.amount) || 0;
    if (summary[t.status] !== undefined) { summary[t.status] += amt; summary[t.status + '_count']++; }
    summary.by_category[t.category] = (summary.by_category[t.category] || 0) + amt;
  }
  res.json(summary);
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
    await notify(data.created_by, id, 'fully_approved', `Your transaction "${data.description}" (KES ${data.amount}) was fully approved and is ready to be paid.`);
  } else {
    await notify(data.created_by, id, 'approved_manager', `Manager approved your transaction "${data.description}" (KES ${data.amount}). Waiting on finance.`);
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
    await notify(data.created_by, id, 'fully_approved', `Your transaction "${data.description}" (KES ${data.amount}) was fully approved and is ready to be paid.`);
  } else {
    await notify(data.created_by, id, 'approved_finance', `Finance approved your transaction "${data.description}" (KES ${data.amount}). Waiting on manager.`);
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
  await notify(data.created_by, id, 'rejected', `Your transaction "${data.description}" (KES ${data.amount}) was rejected.${reason ? ' Reason: ' + reason : ''}`);
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

  const receiptMsg = `Payment settled for "${data.description}" (KES ${data.amount}). Receipt available.`;
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
  const pendingOnly = data.filter(t => t.status === 'pending');
  const approved = data.filter(t => t.status === 'approved');
  const rejected = data.filter(t => t.status === 'rejected');
  const pendingAndApproved = data.filter(t => t.status === 'pending' || t.status === 'approved');

  const paidSheet = workbook.addWorksheet('Paid'); styleSheet(paidSheet); paidSheet.addRows(paid);
  const pendingSheet = workbook.addWorksheet('Pending'); styleSheet(pendingSheet); pendingSheet.addRows(pendingAndApproved);
  const rejectedSheet = workbook.addWorksheet('Rejected'); styleSheet(rejectedSheet); rejectedSheet.addRows(rejected);

  const totalPaid = paid.reduce((s, t) => s + Number(t.amount), 0);
  const totalPending = pendingOnly.reduce((s, t) => s + Number(t.amount), 0);
  const totalRejected = rejected.reduce((s, t) => s + Number(t.amount), 0);
  const remaining = approved.reduce((s, t) => s + Number(t.amount), 0);

  const summarySheet = workbook.addWorksheet('Summary');
  summarySheet.columns = [{ header: 'Metric', key: 'metric', width: 26 }, { header: 'Value', key: 'value', width: 18 }];
  summarySheet.getRow(1).font = { bold: true };
  summarySheet.addRows([
    { metric: 'Total Paid', value: totalPaid },
    { metric: 'Total Pending', value: totalPending },
    { metric: 'Total Rejected', value: totalRejected },
    { metric: 'Remaining (Approved, awaiting payment)', value: remaining },
    { metric: 'Pending (count)', value: pendingOnly.length },
    { metric: 'Paid (count)', value: paid.length },
    { metric: 'Approved (count)', value: approved.length },
    { metric: 'Rejected (count)', value: rejected.length }
  ]);
  summarySheet.getColumn('value').numFmt = '#,##0.00';
  summarySheet.getCell('B6').numFmt = '0';
  summarySheet.getCell('B7').numFmt = '0';
  summarySheet.getCell('B8').numFmt = '0';
  summarySheet.getCell('B9').numFmt = '0';

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
  let query = supabase.from('cst_users').select('id, full_name, email, role, site_id, is_active, created_at').order('created_at', { ascending: false });
  if (req.user.role !== 'admin') query = query.eq('site_id', req.user.site_id);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.patch('/api/users/:id/status', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { is_active } = req.body;
  if (typeof is_active !== 'boolean') return res.status(400).json({ error: 'is_active (boolean) is required' });
  if (req.params.id === req.user.id) return res.status(400).json({ error: "You can't deactivate your own account" });

  const { data, error } = await supabase.from('cst_users').update({ is_active }).eq('id', req.params.id).select('id, full_name, email, role, is_active').single();
  if (error) return res.status(500).json({ error: error.message });
  if (!data) return res.status(404).json({ error: 'User not found' });
  res.json(data);
});

// ---------- Notifications (in-app + real email via Resend) ----------
async function sendEmail(to, subject, text) {
  if (!process.env.RESEND_API_KEY || !to) return;
  try {
    const r = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + process.env.RESEND_API_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify({ from: 'Site Transactions <notifications@cxm.co.ke>', to: [to], subject, text })
    });
    if (!r.ok) { const body = await r.text(); console.error('Resend rejected email:', r.status, body); }
  } catch (e) { console.error('email send failed:', e.message); }
}

async function notify(userId, transactionId, type, message) {
  if (!userId) return;
  await supabase.from('cst_notifications').insert({ user_id: userId, transaction_id: transactionId, type, message });
  try {
    const { data: u } = await supabase.from('cst_users').select('email, full_name').eq('id', userId).single();
    if (u && u.email) await sendEmail(u.email, 'Site Transactions update', message);
  } catch {}
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
app.get('/api/transactions/:id/batch-items', requireAuth, async (req, res) => {
  const { data: txn, error: txnErr } = await supabase.from('cst_transactions').select('batch_id').eq('id', req.params.id).single();
  if (txnErr || !txn) return res.status(404).json({ error: 'Transaction not found' });
  if (!txn.batch_id) return res.json([]);
  const { data, error } = await supabase.from('cst_batch_items').select('*').eq('batch_id', txn.batch_id).order('created_at');
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

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

// ---------- Worker registry & daily attendance ----------
app.post('/api/workers', requireAuth, requireRole('clerk', 'admin'), async (req, res) => {
  const { name, id_number, phone_number, designation, daily_rate } = req.body;
  if (!name || !id_number || !designation) return res.status(400).json({ error: 'name, id_number and designation are required' });
  const site_id = req.user.site_id;
  if (!site_id) return res.status(400).json({ error: 'Your account has no site associated with it' });

  const { data, error } = await supabase.from('cst_workers').insert({
    site_id, name, id_number, phone_number: phone_number || null, designation, daily_rate: daily_rate || 0, enrolled_by: req.user.id
  }).select().single();
  if (error) {
    if (error.message.includes('duplicate')) return res.status(409).json({ error: 'A worker with this ID number is already enrolled at this site' });
    return res.status(500).json({ error: error.message });
  }
  res.status(201).json(data);
});

app.get('/api/workers', requireAuth, async (req, res) => {
  let query = supabase.from('cst_workers').select('*').order('name');
  if (req.user.role !== 'admin') query = query.eq('site_id', req.user.site_id);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.delete('/api/workers/:id', requireAuth, requireRole('clerk', 'manager', 'admin'), async (req, res) => {
  const { error } = await supabase.from('cst_workers').delete().eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

app.post('/api/attendance', requireAuth, requireRole('clerk', 'admin'), async (req, res) => {
  const { attendance_date, worker_ids } = req.body;
  if (!Array.isArray(worker_ids) || worker_ids.length === 0) return res.status(400).json({ error: 'worker_ids (non-empty array) is required' });
  const site_id = req.user.site_id;
  if (!site_id) return res.status(400).json({ error: 'Your account has no site associated with it' });
  const date = attendance_date || new Date().toISOString().slice(0, 10);

  const rows = worker_ids.map(worker_id => ({ worker_id, site_id, attendance_date: date, submitted_by: req.user.id }));
  const { data, error } = await supabase.from('cst_attendance').upsert(rows, { onConflict: 'worker_id,attendance_date', ignoreDuplicates: true }).select();
  if (error) return res.status(500).json({ error: error.message });
  res.status(201).json({ ok: true, recorded: data.length, date });
});

app.get('/api/attendance', requireAuth, async (req, res) => {
  const date = req.query.date || new Date().toISOString().slice(0, 10);
  let query = supabase.from('cst_attendance').select('worker_id, attendance_date').eq('attendance_date', date);
  if (req.user.role !== 'admin') query = query.eq('site_id', req.user.site_id);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

// ---------- Weekly wages (attendance days x daily rate) ----------
async function computeWeeklyWages(req) {
  const start = req.query.start || req.body?.start || (() => {
    const d = new Date();
    const day = d.getDay(); // 0=Sun
    const diffToMonday = day === 0 ? 6 : day - 1;
    d.setDate(d.getDate() - diffToMonday);
    return d.toISOString().slice(0, 10);
  })();
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  const endStr = end.toISOString().slice(0, 10);

  let workerQuery = supabase.from('cst_workers').select('*').order('name');
  if (req.user.role !== 'admin') workerQuery = workerQuery.eq('site_id', req.user.site_id);
  const { data: workers, error: wErr } = await workerQuery;
  if (wErr) throw wErr;

  let attQuery = supabase.from('cst_attendance').select('worker_id, attendance_date').gte('attendance_date', start).lte('attendance_date', endStr);
  if (req.user.role !== 'admin') attQuery = attQuery.eq('site_id', req.user.site_id);
  const { data: attendance, error: aErr } = await attQuery;
  if (aErr) throw aErr;

  const { data: payments, error: pErr } = await supabase.from('cst_wage_payments').select('worker_id, paid_at, disbursement_status').eq('week_start', start).neq('disbursement_status', 'failed');
  if (pErr) throw pErr;
  const paidMap = {};
  for (const p of payments) paidMap[p.worker_id] = p.paid_at;

  const daysByWorker = {};
  for (const a of attendance) daysByWorker[a.worker_id] = (daysByWorker[a.worker_id] || 0) + 1;

  const rows = workers.map(w => {
    const daysPresent = daysByWorker[w.id] || 0;
    return {
      worker_id: w.id, name: w.name, id_number: w.id_number, designation: w.designation, phone_number: w.phone_number,
      daily_rate: Number(w.daily_rate) || 0, days_present: daysPresent, total_pay: daysPresent * (Number(w.daily_rate) || 0),
      paid: !!paidMap[w.id], paid_at: paidMap[w.id] || null
    };
  });
  return { start, end: endStr, rows };
}

app.get('/api/wages/weekly', requireAuth, requireRole('finance', 'manager', 'admin'), async (req, res) => {
  try { res.json(await computeWeeklyWages(req)); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/wages/weekly/export', requireAuth, requireRole('finance', 'manager', 'admin'), async (req, res) => {
  try {
    const { start, end, rows } = await computeWeeklyWages(req);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Construction Site Transactions';
    workbook.created = new Date();
    const sheet = workbook.addWorksheet('Weekly Wages');
    sheet.columns = [
      { header: 'Name', key: 'name', width: 24 },
      { header: 'ID Number', key: 'id_number', width: 16 },
      { header: 'Designation', key: 'designation', width: 18 },
      { header: 'Daily Rate', key: 'daily_rate', width: 14 },
      { header: 'Days Present', key: 'days_present', width: 14 },
      { header: 'Total Pay', key: 'total_pay', width: 16 },
      { header: 'Paid', key: 'paid_label', width: 10 }
    ];
    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
    sheet.addRows(rows.map(r => ({ ...r, paid_label: r.paid ? 'Yes' : 'No' })));
    const grandTotal = rows.reduce((s, r) => s + r.total_pay, 0);
    sheet.addRow({});
    const totalRow = sheet.addRow({ name: 'TOTAL', total_pay: grandTotal });
    totalRow.font = { bold: true };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="weekly-wages-${start}-to-${end}.xlsx"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/wages/weekly/pay-all', requireAuth, requireRole('finance', 'admin'), async (req, res) => {
  try {
    const { start, end, rows } = await computeWeeklyWages(req);
    const site_id = req.user.site_id;
    const toPay = rows.filter(r => !r.paid && r.total_pay > 0);
    if (toPay.length === 0) return res.json({ ok: true, paid_count: 0, message: 'Nothing to pay — all workers already settled or no attendance this week.' });

    const mpesaOn = mpesaConfigured();
    const callbackHost = req.get('host');
    const results = [];

    for (const r of toPay) {
      let disbursement_status = 'not_attempted';
      let mpesa_conversation_id = null;
      let mpesa_result_desc = null;

      if (mpesaOn && r.phone_number) {
        try {
          const b2c = await sendB2CPayment({
            phone: r.phone_number, amount: r.total_pay,
            remarks: `Wages ${start} to ${end}`, occasion: r.name, callbackHost
          });
          disbursement_status = b2c.ok ? 'sent' : 'failed';
          mpesa_conversation_id = b2c.data.ConversationID || null;
          mpesa_result_desc = b2c.data.ResponseDescription || b2c.data.errorMessage || JSON.stringify(b2c.data);
        } catch (e) {
          disbursement_status = 'failed';
          mpesa_result_desc = e.message;
        }
      }

      const { error: insErr } = await supabase.from('cst_wage_payments').upsert({
        site_id, worker_id: r.worker_id, week_start: start, week_end: end,
        days_present: r.days_present, daily_rate: r.daily_rate, total_amount: r.total_pay, paid_by: req.user.id,
        disbursement_status, mpesa_conversation_id, mpesa_result_desc, paid_at: new Date().toISOString()
      }, { onConflict: 'worker_id,week_start' });
      if (insErr) { results.push({ name: r.name, ok: false, error: insErr.message }); continue; }
      results.push({ name: r.name, ok: true, disbursement_status, mpesa_result_desc });
    }

    // Log a single reconciliation transaction covering the whole payroll run, already marked paid.
    const grandTotal = toPay.reduce((s, r) => s + r.total_pay, 0);
    const breakdown = toPay.map(r => `${r.name}: ${r.days_present}d x KES ${r.daily_rate} = KES ${r.total_pay.toLocaleString()}`).join('; ');
    await supabase.from('cst_transactions').insert({
      site_id, category: 'labor', description: `Weekly wages ${start} to ${end} (${toPay.length} workers): ${breakdown}`,
      amount: grandTotal, transaction_date: new Date().toISOString().slice(0, 10),
      status: 'paid', created_by: req.user.id, paid_by: req.user.id, paid_at: new Date().toISOString(),
      manager_approved: true, finance_approved: true
    });

    res.json({ ok: true, paid_count: toPay.length, total_amount: grandTotal, mpesa_enabled: mpesaOn, results });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ---------- M-Pesa B2C callbacks ----------
app.post('/api/mpesa/b2c/result', express.json(), async (req, res) => {
  try {
    const result = req.body.Result || {};
    const conversationId = result.ConversationID;
    const resultCode = result.ResultCode;
    const resultDesc = result.ResultDesc;
    let receipt = null;
    const params = result.ResultParameters && result.ResultParameters.ResultParameter;
    if (Array.isArray(params)) {
      const txnParam = params.find(p => p.Key === 'TransactionReceipt' || p.Key === 'TransactionID');
      if (txnParam) receipt = txnParam.Value;
    }
    if (conversationId) {
      await supabase.from('cst_wage_payments').update({
        disbursement_status: resultCode === 0 || resultCode === '0' ? 'confirmed' : 'failed',
        mpesa_receipt: receipt, mpesa_result_desc: resultDesc
      }).eq('mpesa_conversation_id', conversationId);
    }
  } catch (e) { console.error('M-Pesa result callback error:', e.message); }
  res.json({ ResultCode: 0, ResultDesc: 'Accepted' });
});

app.post('/api/mpesa/b2c/timeout', express.json(), (req, res) => {
  console.error('M-Pesa B2C request timed out:', JSON.stringify(req.body));
  res.json({ ResultCode: 0, ResultDesc: 'Accepted' });
});

app.get('/api/health', (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
if (require.main === module) {
  app.listen(PORT, () => console.log(`Construction site transactions API running on :${PORT}`));
}

module.exports = app;
