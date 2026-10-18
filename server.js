require('dotenv').config();
const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');
const PDFDocument = require('pdfkit');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } });

// ---------- Shared PDF form styling helpers ----------
const PDF_FOREST = '#0f2818';
const PDF_MOSS = '#7ab98a';
const PDF_LEFT = 40, PDF_RIGHT = 555;

function drawDefaultLogo(doc) {
  doc.roundedRect(PDF_LEFT, 40, 44, 44, 8).fillAndStroke('#eaf7ee', PDF_FOREST);
  doc.rect(PDF_LEFT + 9, 57, 20, 13).fill(PDF_FOREST);
  doc.circle(PDF_LEFT + 16, 74, 3.5).fill(PDF_FOREST);
  doc.circle(PDF_LEFT + 29, 74, 3.5).fill(PDF_FOREST);
}

function drawFormHeader(doc, title, subtitle, branding) {
  let logoDrawn = false;
  if (branding && branding.logoBuffer) {
    try {
      doc.save();
      doc.roundedRect(PDF_LEFT, 40, 44, 44, 8).clip();
      doc.image(branding.logoBuffer, PDF_LEFT, 40, { width: 44, height: 44 });
      doc.restore();
      logoDrawn = true;
    } catch { logoDrawn = false; }
  }
  if (!logoDrawn) drawDefaultLogo(doc);

  const brandName = (branding && branding.name) || 'Site Transactions';
  doc.fillColor(PDF_FOREST).fontSize(20).font('Helvetica-Bold').text(brandName, PDF_LEFT + 56, 44, { width: PDF_RIGHT - PDF_LEFT - 56 });
  doc.fontSize(10).font('Helvetica').fillColor('#555').text(subtitle, PDF_LEFT + 56, 68);

  doc.moveTo(PDF_LEFT, 96).lineTo(PDF_RIGHT, 96).lineWidth(1.5).strokeColor(PDF_FOREST).stroke();
  doc.lineWidth(1);

  doc.fontSize(14).font('Helvetica-Bold').fillColor(PDF_FOREST).text(title, PDF_LEFT, 106);
}

async function fetchOrgBranding(req) {
  if (!req.user.organization_id) return null;
  try {
    const { data: org } = await supabase.from('cst_organizations').select('name, logo_url').eq('id', req.user.organization_id).single();
    if (!org) return null;
    let logoBuffer = null;
    if (org.logo_url) {
      try {
        const r = await fetch(org.logo_url);
        if (r.ok) logoBuffer = Buffer.from(await r.arrayBuffer());
      } catch {}
    }
    return { name: org.name, logoBuffer };
  } catch { return null; }
}

function drawInfoBox(doc, y, fields) {
  const boxHeight = fields.length * 20 + 12;
  doc.roundedRect(PDF_LEFT, y, PDF_RIGHT - PDF_LEFT, boxHeight, 4).strokeColor('#ccc').stroke();
  doc.font('Helvetica').fontSize(9).fillColor('#333');
  fields.forEach((line, i) => doc.text(line, PDF_LEFT + 10, y + 8 + i * 20));
  return y + boxHeight;
}

function drawGridTable(doc, { startY, columns, headerHeight = 22, rowHeight = 22, numRows }) {
  const tableWidth = PDF_RIGHT - PDF_LEFT;
  let colX = [PDF_LEFT];
  columns.forEach(c => colX.push(colX[colX.length - 1] + c.width));
  const tableBottom = startY + headerHeight + numRows * rowHeight;

  // Header row background
  doc.rect(PDF_LEFT, startY, tableWidth, headerHeight).fillAndStroke('#e2f1e8', PDF_FOREST);
  doc.font('Helvetica-Bold').fontSize(9).fillColor(PDF_FOREST);
  columns.forEach((c, i) => doc.text(c.label, colX[i] + 6, startY + 7, { width: c.width - 10 }));

  // Outer border
  doc.rect(PDF_LEFT, startY, tableWidth, tableBottom - startY).strokeColor(PDF_FOREST).lineWidth(1.2).stroke();
  doc.lineWidth(1);

  // Column dividers
  colX.slice(1, -1).forEach(x => {
    doc.moveTo(x, startY).lineTo(x, tableBottom).strokeColor('#bbb').stroke();
  });

  // Row dividers
  for (let i = 1; i <= numRows; i++) {
    const y = startY + headerHeight + i * rowHeight;
    doc.moveTo(PDF_LEFT, y).lineTo(PDF_RIGHT, y).strokeColor('#ddd').stroke();
  }
  doc.moveTo(PDF_LEFT, startY + headerHeight).lineTo(PDF_RIGHT, startY + headerHeight).strokeColor(PDF_FOREST).lineWidth(1.2).stroke();
  doc.lineWidth(1);

  return tableBottom;
}

function drawFormFooter(doc, y, text) {
  doc.moveTo(PDF_LEFT, y).lineTo(PDF_RIGHT, y).strokeColor('#ccc').stroke();
  doc.font('Helvetica-Oblique').fontSize(8).fillColor('#777').text(text, PDF_LEFT, y + 10, { width: PDF_RIGHT - PDF_LEFT });
}

const ExcelJS = require('exceljs');
const { createClient } = require('@supabase/supabase-js');

const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change-me';
const CATEGORIES = ['materials', 'equipment', 'fuel', 'other'];

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY, {
  auth: { autoRefreshToken: false, persistSession: false }
});

// ---------- Multi-tenancy: organization-scoped visibility ----------
// 'admin' sees every site within their own organization (not globally).
// 'superadmin' sees everything, across all organizations.
// Everyone else is scoped to their own single site, as before.
async function getOrgSiteIds(organizationId) {
  if (!organizationId) return [];
  const { data } = await supabase.from('cst_sites').select('id').eq('organization_id', organizationId);
  return (data || []).map(s => s.id);
}
// Applies the correct site-scoping filter to a Supabase query builder, given the requesting user.
async function scopeToSites(query, user) {
  if (user.role === 'superadmin') return query;
  if (user.role === 'admin') return query.in('site_id', await getOrgSiteIds(user.organization_id));
  return user.site_id ? query.eq('site_id', user.site_id) : query;
}

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
    { id: user.id, role: user.role, site_id: user.site_id, organization_id: user.organization_id || null, full_name: user.full_name },
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
<link rel="manifest" href="/manifest.webmanifest" />
<meta name="theme-color" content="#173B29" />
<meta name="mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-title" content="Site Transactions" />
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
<link rel="apple-touch-icon" href="/apple-touch-icon.png" />
<link rel="icon" type="image/png" href="/icon-192.png" />
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

.card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:22px;margin-bottom:16px;box-shadow:0 1px 3px rgba(15,40,24,.04);transition:box-shadow .2s ease,transform .2s ease;}
.card:hover{box-shadow:0 6px 20px rgba(15,40,24,.08);}
h1{font-family:'Space Grotesk',sans-serif;font-size:22px;margin:0 0 4px;letter-spacing:-0.01em;}
h2{font-family:'Space Grotesk',sans-serif;font-size:16px;margin:0 0 12px;font-weight:600;color:var(--forest);letter-spacing:-0.01em;}
.muted{color:#6b7684;font-size:13px;}
input,select,textarea,button{font-family:'Inter',-apple-system,sans-serif;font-size:14px;padding:10px 12px;border-radius:8px;border:1.5px solid var(--border);}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--moss);box-shadow:0 0 0 3px rgba(122,185,138,.18);}
label{display:block;font-size:12px;margin:10px 0 5px;font-weight:600;color:#4a5a50;text-transform:uppercase;letter-spacing:.03em;}
button{background:var(--forest);color:#fff;border:none;cursor:pointer;font-weight:600;transition:background .15s ease,transform .08s ease,box-shadow .15s ease;}
button:hover{background:var(--forest2);box-shadow:0 3px 10px rgba(15,40,24,.18);}
button:active{transform:scale(.97);}
button.secondary{background:#eef0f3;color:var(--navy);} button.secondary:hover{background:#e2e5ea;box-shadow:none;}
button.danger{background:var(--red);} button.danger:hover{background:#8f2e2e;}
button.success{background:var(--sidebar-active);} button.success:hover{background:#268049;}
.row{display:flex;gap:10px;flex-wrap:wrap;} .row>*{flex:1;min-width:160px;}
table{width:100%;border-collapse:collapse;margin-top:10px;font-size:13px;}
th,td{text-align:left;padding:9px 8px;border-bottom:1px solid var(--border);}
th{color:#6b7684;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.03em;}
tbody tr{transition:background .12s ease;} tbody tr:hover{background:#f5f9f6;}
.badge{padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.02em;}
.badge.pending{background:#fdf0dc;color:var(--amber);} .badge.approved{background:#e2f1e8;color:var(--green);}
.badge.paid{background:#dfeee5;color:var(--green);} .badge.rejected{background:#f6e2e2;color:var(--red);}
.actions button{margin-right:4px;padding:6px 10px;font-size:12px;border-radius:6px;}
.topbar{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;flex-wrap:wrap;gap:8px;}
.error{color:var(--red);font-size:13px;margin-top:8px;}

/* ---- View transitions & loading states ---- */
#content > *{animation:viewFadeIn .35s ease both;}
#content > *:nth-child(2){animation-delay:.04s;} #content > *:nth-child(3){animation-delay:.08s;}
#content > *:nth-child(4){animation-delay:.12s;} #content > *:nth-child(5){animation-delay:.16s;}
@keyframes viewFadeIn{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:translateY(0);}}
.skeleton{background:linear-gradient(90deg,#eef1ef 25%,#e2e8e4 37%,#eef1ef 63%);background-size:400% 100%;animation:skeletonShine 1.4s ease infinite;border-radius:6px;height:14px;margin:6px 0;}
@keyframes skeletonShine{0%{background-position:100% 50%;}100%{background-position:0 50%;}}
.stat-card{background:#f5f6f8;border-radius:10px;padding:14px;transition:transform .18s ease,box-shadow .18s ease;}
.stat-card:hover{transform:translateY(-2px);box-shadow:0 6px 16px rgba(15,40,24,.1);}
.stat-num{font-family:'Space Grotesk',sans-serif;font-size:19px;font-weight:700;}
.nav-item{transition:background .15s ease,transform .1s ease;}
.nav-item:active{transform:scale(.97);}

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
// ----- Install as an app (PWA) -----
var deferredInstall = null;
if ('serviceWorker' in navigator) { window.addEventListener('load', function(){ navigator.serviceWorker.register('/sw.js').catch(function(){}); }); }
function isStandalone(){ return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true; }
function showInstallBar(){
  if (isStandalone() || document.getElementById('installBar')) return;
  try { if (localStorage.getItem('installDismissed') === '1') return; } catch(e){}
  var isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
  if (!deferredInstall && !isIOS) return;
  var bar = document.createElement('div');
  bar.id = 'installBar';
  bar.style.cssText = 'position:fixed;left:12px;right:12px;bottom:12px;z-index:9999;background:#173B29;color:#fff;border-radius:14px;padding:12px 14px;display:flex;align-items:center;gap:10px;box-shadow:0 8px 24px rgba(0,0,0,.3);font:500 14px Inter,system-ui,sans-serif;';
  bar.innerHTML = '<div style="flex:1;">' + (deferredInstall ? 'Install Site Transactions on your phone for quick access.' : 'To install: tap the Share button, then \\"Add to Home Screen\\".') + '</div>' +
    (deferredInstall ? '<button id="installBtn" style="background:#2f9e5c;color:#fff;border:0;border-radius:10px;padding:9px 14px;font-weight:600;">Install</button>' : '') +
    '<button id="installClose" aria-label="Close" style="background:transparent;color:#cfe3d6;border:0;font-size:20px;">&times;</button>';
  document.body.appendChild(bar);
  document.getElementById('installClose').onclick = function(){ bar.remove(); try { localStorage.setItem('installDismissed','1'); } catch(e){} };
  var b = document.getElementById('installBtn');
  if (b) b.onclick = function(){ deferredInstall.prompt(); deferredInstall.userChoice.then(function(){ deferredInstall = null; bar.remove(); }); };
}
window.addEventListener('beforeinstallprompt', function(e){ e.preventDefault(); deferredInstall = e; showInstallBar(); });
window.addEventListener('appinstalled', function(){ var b = document.getElementById('installBar'); if (b) b.remove(); });
window.addEventListener('load', function(){ setTimeout(showInstallBar, 2500); });
const API = '/api';
let state = { token: localStorage.getItem('cst_token'), user: JSON.parse(localStorage.getItem('cst_user') || 'null') };
const appEl = document.getElementById('app');
const TRUCK_SVG = '<svg width="26" height="26" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="18" width="24" height="14" rx="2" fill="#1f2d3d"/><path d="M26 22h9l7 7v3a2 2 0 0 1-2 2h-1" stroke="#1f2d3d" stroke-width="2.5" fill="none" stroke-linejoin="round"/><circle cx="13" cy="35" r="4" fill="#173B29" stroke="#1f2d3d" stroke-width="1.5"/><circle cx="34" cy="35" r="4" fill="#173B29" stroke="#1f2d3d" stroke-width="1.5"/></svg>';
function initials(name){ return (name||'?').split(' ').filter(Boolean).slice(0,2).map(w=>w[0].toUpperCase()).join(''); }
function save(token, user){ state={token,user}; localStorage.setItem('cst_token',token); localStorage.setItem('cst_user', JSON.stringify(user)); }
function logout(){ window.__brandLoaded = false; localStorage.removeItem('cst_token'); localStorage.removeItem('cst_user'); state={token:null,user:null}; render(); }
async function api(path, opts={}) {
  const res = await fetch(API+path, { ...opts, headers: { 'Content-Type':'application/json', ...(state.token?{Authorization:'Bearer '+state.token}:{}), ...(opts.headers||{}) } });
  if (!res.ok) { const b = await res.json().catch(()=>({})); throw new Error(b.error||'Request failed'); }
  return res.json();
}
function money(n){ return 'KES ' + Number(n).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}); }
function badge(s){ return '<span class="badge '+s+'">'+s+'</span>'; }
function render(){
  if (!state.token) return renderLogin();
  if (state.user.role === 'superadmin') return renderSuperadminPlatform();
  return renderDashboard();
}
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
        '<div class="auth-field"><label>Email</label><input id="email" type="email" placeholder="you@example.com" autocapitalize="none" autocorrect="off" spellcheck="false" autocomplete="username" /></div>' +
        '<div class="auth-field"><label>Password</label><input id="password" type="password" placeholder="********" autocapitalize="none" autocomplete="current-password" /><label style="display:flex;gap:6px;align-items:center;font-size:13px;margin-top:6px;"><input type="checkbox" id="showPw" style="width:auto;" /> Show password</label></div>' +
        '<button class="auth-submit" id="loginBtn">Log in</button>' +
        '<a href="#" id="forgotLink" class="auth-forgot">Forgot password?</a>' +
        '<div class="error" id="err"></div><div class="muted" id="forgotMsg"></div>' +
      '</div></div>' +
    '</div>';
  document.getElementById('showPw').onchange = function(){ document.getElementById('password').type = this.checked ? 'text' : 'password'; };
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
  let html = '<div class="card"><div class="topbar"><h2>Overview</h2>'+(canExport?'<button class="success" id="overviewExportBtn">Download Excel</button>':'')+'</div><div id="summaryBox"><div class="skeleton" style="width:90%;"></div><div class="skeleton" style="width:70%;"></div><div class="skeleton" style="width:80%;"></div></div></div>';
  if (['finance','manager','admin'].includes(user.role)) {
    html += '<div class="card"><h2>Waiting for approval</h2><p class="muted">Pending transactions that need a manager and/or finance sign-off.</p><div id="queuePending"><div class="skeleton" style="width:80%;"></div></div></div>';
    html += '<div class="card"><h2>Waiting for payment</h2><p class="muted">Fully approved — ready for finance to settle.</p><div id="queueApproved"><div class="skeleton" style="width:80%;"></div></div></div>';
  }
  return html;
}
function buildTransactionsHtml(user){
  let html = '';
  html += '<div class="card"><h2>Offline form</h2><p class="muted">No network on site? Download a blank printable form, fill it by hand, and enter it into the app (or upload a photo below) once you\\'re back online.</p><button class="secondary" id="downloadFormBtn">Download blank form (PDF)</button></div>';
  if (user.role==='clerk' || user.role==='admin') {
    html += '<div class="card"><h2>Log transactions</h2><p class="muted">Add materials, labor, fuel, etc. as separate line items, then submit them together — the total goes to your manager and finance in one notification.</p>'+(user.site_id ? '<input id="site_id" type="hidden" value="'+user.site_id+'" />' : '<div class="row"><div><label>Site ID</label><input id="site_id" placeholder="site UUID" /></div></div>')+'<div id="itemRows"></div><div style="margin-top:8px;"><button class="secondary" id="addItemBtn" type="button">+ Add item</button></div><div class="topbar" style="margin-top:14px;"><strong>Total: <span id="batchTotal">KES 0.00</span></strong><button id="submitBatch">Submit all</button></div><div class="muted" id="outboxNote"></div><div class="error" id="submitErr"></div></div>';
  }
  html += '<div class="card"><div class="topbar"><h2>'+(user.role==='clerk'?'My submissions':'Transactions')+'</h2><div>'+(['manager','finance','admin'].includes(user.role)?'<button id="exportBtn" class="secondary">Download Excel</button>':'')+(['manager','admin'].includes(user.role)?' <button id="clearAllBtn" class="danger">Clear all data</button>':'')+'</div></div>' +
    '<div class="row" style="margin-top:6px;"><input id="filterSearch" placeholder="Search description..." /><select id="filterStatus"><option value="">All statuses</option><option value="pending">Pending</option><option value="approved">Approved</option><option value="paid">Paid</option><option value="rejected">Rejected</option></select><input id="filterFrom" type="date" /><input id="filterTo" type="date" /><button class="secondary" id="filterApply">Filter</button></div>' +
    '<div id="table"></div><div id="pagination" class="row" style="margin-top:10px;justify-content:flex-end;"></div></div>';
  html += '<div class="card"><h2>Documents</h2><p class="muted">Upload photos or scans of filled forms, receipts, or other paper records.</p><div class="row"><input id="docFile" type="file" accept="image/*,.pdf" /><input id="docDesc" placeholder="Short description (optional)" /></div><div style="margin-top:10px;"><button id="uploadDocBtn">Upload</button></div><div class="error" id="docErr"></div><div id="docList" class="muted" style="margin-top:12px;">Loading...</div></div>';
  return html;
}
function buildRegistryHtml(user){
  let html = '';
  if (user.role==='clerk' || user.role==='manager' || user.role==='admin') {
    html += '<div class="card"><h2>Offline worker registration form</h2><p class="muted">No network? Download this, fill it by hand for each worker, then enter them into the app once you\\'re back online.</p><button class="secondary" id="downloadWorkerFormBtn">Download blank form (PDF)</button></div>';
    html += '<div class="card"><h2>Enroll a worker</h2><div class="row"><div><label>Full name</label><input id="w_name" placeholder="Full name" /></div><div><label>ID number</label><input id="w_id" placeholder="National ID number" /></div></div><div class="row"><div><label>Phone number</label><input id="w_phone" placeholder="07xxxxxxxx" /></div><div><label>Designation</label><input id="w_designation" placeholder="e.g. Mason, Fundi, Laborer" /></div><div><label>Daily rate (KES)</label><input id="w_rate" type="number" step="0.01" placeholder="e.g. 800" /></div></div><div style="margin-top:12px;"><button id="enrollWorkerBtn">Enroll worker</button></div><div class="error" id="enrollWorkerErr"></div></div>';
  }
  if (user.role==='clerk' || user.role==='admin') {
    html += '<div class="card"><div class="topbar"><h2>Daily attendance</h2><input id="attDate" type="date" /></div><div id="workerAttendance" class="muted">Loading workers...</div><div style="margin-top:12px;"><button id="submitAttendance">Submit today\\'s attendance</button></div><div class="error" id="attendanceErr"></div><div id="attendanceOk" class="muted"></div></div>';
  }
  if (user.role==='manager' || user.role==='finance') {
    html += '<div class="card"><div class="topbar"><h2>Workers & attendance</h2><input id="attViewDate" type="date" /></div><div id="workerAttendanceView"><div class="skeleton" style="width:85%;"></div><div class="skeleton" style="width:60%;"></div></div></div>';
  }
  if (user.role==='finance' || user.role==='manager' || user.role==='admin') {
    html += '<div class="card"><div class="topbar"><h2>Weekly wages</h2><div><label style="display:inline;margin-right:6px;">Week starting</label><input id="wagesWeekStart" type="date" /></div></div><div id="weeklyWages"><div class="skeleton" style="width:85%;"></div><div class="skeleton" style="width:60%;"></div></div><div style="margin-top:10px;text-align:left;">'+(['finance','admin'].includes(user.role)?'<button class="success" id="payAllWagesBtn">Pay all wages</button> ':'')+'<button class="success" id="downloadWagesBtn">Download Excel</button></div><div class="error" id="payWagesErr"></div><div class="muted" id="payWagesOk"></div></div>';
  }
  html += '<div class="card"><h2>Weekly attendance report</h2><p class="muted">Present shown in green, absent in red — printable at the end of the week.</p><div class="row"><div><label>Week starting</label><input id="attReportWeekStart" type="date" /></div></div><div style="margin-top:12px;"><button class="success" id="downloadAttendanceReportBtn">Download PDF report</button></div></div>';
  return html;
}
function buildNotificationsHtml(user){
  return '<div class="card"><div class="topbar"><h2>Notifications</h2><span class="muted" id="notifCount"></span></div><div id="notifList"><div class="skeleton" style="width:85%;"></div><div class="skeleton" style="width:60%;"></div></div></div>';
}
function buildSettingsHtml(user){
  let html = '<div class="card"><h2>Change password</h2><div class="row"><div><label>Current password</label><input id="curPw" type="password" /></div><div><label>New password</label><input id="newPw" type="password" /></div></div><div style="margin-top:12px;"><button id="changePwBtn">Update password</button></div><div class="error" id="changePwErr"></div><div id="changePwOk" class="muted"></div></div>';
  if ((user.role==='manager' || user.role==='admin') && user.organization_id) {
    html += '<div class="card"><h2>Company logo</h2><p class="muted">Your logo appears in the menu and on every PDF form, report and receipt. Use a square PNG or JPG under 2 MB.</p><div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;"><div id="logoPreview" style="width:72px;height:72px;border-radius:12px;background:#eee;display:flex;align-items:center;justify-content:center;color:#999;font-size:12px;overflow:hidden;">No logo</div><div><input id="logoFile" type="file" accept="image/png,image/jpeg" /><div style="margin-top:8px;"><button id="uploadLogoBtn">Upload logo</button></div></div></div><div class="error" id="logoErr"></div><div class="muted" id="logoOk"></div></div>';
  }
  if (user.role==='manager' || user.role==='admin') {
    html += '<div class="card"><h2>Add a team login</h2><div class="row"><div><label>Full name</label><input id="new_name" placeholder="Full name" /></div><div><label>Email</label><input id="new_email" type="email" placeholder="person@example.com" /></div><div><label>Role</label><select id="new_role"><option value="clerk">Clerk</option><option value="finance">Finance</option><option value="storekeeper">Storekeeper</option></select></div></div><div style="margin-top:12px;"><button id="createUserBtn">Create login</button></div><div class="error" id="createUserErr"></div><div id="createUserResult"></div></div>';
    html += '<div class="card"><h2>Team logins</h2><div id="userList"><div class="skeleton" style="width:85%;"></div><div class="skeleton" style="width:60%;"></div></div></div>';
  }
  return html;
}

function buildStoreHtml(user){
  const canManage = ['storekeeper','admin'].includes(user.role);
  const canReceive = ['storekeeper','admin'].includes(user.role);
  let html = '';
  if (['manager','finance'].includes(user.role)) {
    html += '<div class="card"><p class="muted">You have view access here — the storekeeper records deliveries and stock changes. This is the record of what suppliers have brought to site and what\\'s owed.</p></div>';
  }
  if (canReceive) {
    html += '<div class="card"><h2>Record a delivery</h2><p class="muted">When a supplier brings materials to site, record it all in one go — the supplier, what they brought, and what you owe them. This creates a pending payment for manager and finance to approve.</p>' +
      '<div class="row"><div><label>Supplier (existing)</label><select id="dv_supplier"><option value="">— or type a new one below —</option></select></div><div><label>New supplier name</label><input id="dv_new_supplier" placeholder="Only if not in the list" /></div></div>' +
      '<div class="row"><div><label>Material (existing)</label><select id="dv_item"><option value="">— or type a new one below —</option></select></div><div><label>New material name</label><input id="dv_new_item" placeholder="e.g. Cement (50kg bags)" /></div><div><label>Unit (for new material)</label><input id="dv_new_unit" placeholder="e.g. bags, liters" /></div></div>' +
      '<div class="row"><div><label>Quantity delivered</label><input id="dv_qty" type="number" step="0.01" /></div><div><label>Amount owed (KES)</label><input id="dv_amount" type="number" step="0.01" /></div></div>' +
      '<label>Note (optional)</label><input id="dv_note" placeholder="e.g. delivery note number" style="width:100%;" />' +
      '<label style="margin-top:8px;">Delivery note photo (optional)</label><input id="dv_photo" type="file" accept="image/*" />' +
      '<div style="margin-top:12px;"><button id="recordDeliveryBtn">Record delivery</button></div><div class="error" id="deliveryErr"></div><div class="muted" id="deliveryOk"></div></div>';
  }
  if (canManage) {
    html += '<div class="card"><h2>Add a material manually</h2><p class="muted">Use this to start tracking a material with no delivery yet (e.g. zero starting stock).</p><div class="row"><div><label>Material name</label><input id="item_name" placeholder="e.g. Cement (50kg bags)" /></div><div><label>Unit</label><input id="item_unit" placeholder="e.g. bags, liters, pieces" /></div></div><div class="row"><div><label>Initial quantity</label><input id="item_qty" type="number" step="0.01" placeholder="0" /></div><div><label>Reorder level (alert below this)</label><input id="item_reorder" type="number" step="0.01" placeholder="optional" /></div></div><div style="margin-top:12px;"><button id="addItemMaterialBtn">Add material</button></div><div class="error" id="addItemErr"></div></div>';
  }
  html += '<div class="card"><div class="topbar"><h2>Material stock</h2><button class="success" id="downloadStockPdfBtn">Download PDF report</button></div><div id="storeList"><div class="skeleton" style="width:85%;"></div><div class="skeleton" style="width:60%;"></div></div></div>';
  return html;
}

function buildAnalyticsHtml(user){
  const isManager = ['manager','admin'].includes(user.role);
  let html = '<div class="card"><h2>'+(isManager?'Analytics — everything':'Financial analytics')+'</h2><p class="muted">'+(isManager?'A full picture of transactions, workers, store, and documents at your site.':'Money-related analytics only — transactions, wages, and payment status.')+'</p></div>';
  html += '<div class="card"><h2>Transactions by status</h2><div id="analyticsStatusChart" class="muted"><div class="skeleton" style="width:80%;"></div></div></div>';
  html += '<div class="card"><h2>Transactions by category</h2><div id="analyticsCategoryChart" class="muted"><div class="skeleton" style="width:80%;"></div></div></div>';
  html += '<div class="card"><h2>Weekly transaction trend (last 6 weeks)</h2><div id="analyticsTrendChart" class="muted"><div class="skeleton" style="width:80%;"></div></div></div>';
  html += '<div class="card"><h2>Wages paid to date</h2><div id="analyticsWages" class="muted"><div class="skeleton" style="width:50%;"></div></div></div>';
  if (isManager) {
    html += '<div class="card"><h2>Site operations</h2><div id="analyticsOperational" class="muted"><div class="skeleton" style="width:70%;"></div></div></div>';
  }
  return html;
}

function buildSuppliersHtml(user){
  let html = '<div class="card"><h2>Add a supplier</h2><div class="row"><div><label>Supplier name</label><input id="sup_name" placeholder="Company or person name" /></div><div><label>Phone</label><input id="sup_phone" placeholder="07xxxxxxxx" /></div><div><label>Email</label><input id="sup_email" type="email" placeholder="optional" /></div></div><div style="margin-top:12px;"><button id="addSupplierBtn">Add supplier</button></div><div class="error" id="addSupplierErr"></div></div>';
  html += '<div class="card"><h2>Suppliers & their payments</h2><div id="supplierList"><div class="skeleton" style="width:85%;"></div><div class="skeleton" style="width:60%;"></div></div></div>';
  return html;
}

async function renderSuperadminPlatform(){
  appEl.innerHTML =
    '<div style="min-height:100vh;background:var(--bg);">' +
      '<div class="topbar-app"><div style="display:flex;align-items:center;gap:10px;"><div class="brand-logo" style="box-shadow:none;width:36px;height:36px;">'+TRUCK_SVG+'</div><div><strong>Platform Admin</strong><div class="muted" style="font-size:11px;">'+state.user.full_name+'</div></div></div><button class="secondary" id="platformLogoutBtn">Log out</button></div>' +
      '<div style="max-width:900px;margin:0 auto;padding:20px;">' +
        '<div class="card"><h2>Add a construction company</h2><p class="muted">Onboard a new company with its logo and details. Optionally set up its first site and admin login at the same time.</p>' +
          '<div class="row"><div><label>Company name</label><input id="org_name" placeholder="e.g. Acme Builders Ltd" /></div><div><label>Logo</label><input id="org_logo" type="file" accept="image/*" /></div></div>' +
          '<div class="row"><div><label>Contact email</label><input id="org_email" type="email" /></div><div><label>Contact phone</label><input id="org_phone" /></div></div>' +
          '<div class="row"><div><label>Address</label><input id="org_address" /></div><div><label>Registration number</label><input id="org_regnum" /></div></div>' +
          '<h2 style="margin-top:16px;">Optional: first site + admin login</h2>' +
          '<div class="row"><div><label>First site name</label><input id="org_site_name" placeholder="e.g. Main Site" /></div></div>' +
          '<div class="row"><div><label>Admin full name</label><input id="org_admin_name" /></div><div><label>Admin email</label><input id="org_admin_email" type="email" /></div></div>' +
          '<div style="margin-top:12px;"><button id="createOrgBtn">Create company</button></div><div class="error" id="createOrgErr"></div><div id="createOrgResult"></div>' +
        '</div>' +
        '<div class="card"><h2>Companies on this platform</h2><div id="orgList"><div class="skeleton" style="width:85%;"></div><div class="skeleton" style="width:60%;"></div></div></div>' +
      '</div>' +
    '</div>';

  document.getElementById('platformLogoutBtn').onclick = logout;
  document.getElementById('createOrgBtn').onclick = async () => {
    const formData = new FormData();
    formData.append('name', document.getElementById('org_name').value.trim());
    if (document.getElementById('org_logo').files[0]) formData.append('logo', document.getElementById('org_logo').files[0]);
    formData.append('contact_email', document.getElementById('org_email').value.trim());
    formData.append('contact_phone', document.getElementById('org_phone').value.trim());
    formData.append('address', document.getElementById('org_address').value.trim());
    formData.append('registration_number', document.getElementById('org_regnum').value.trim());
    formData.append('first_site_name', document.getElementById('org_site_name').value.trim());
    formData.append('admin_full_name', document.getElementById('org_admin_name').value.trim());
    formData.append('admin_email', document.getElementById('org_admin_email').value.trim());
    if (!document.getElementById('org_name').value.trim()) { document.getElementById('createOrgErr').textContent = 'Company name is required.'; return; }
    try {
      const res = await fetch(API+'/superadmin/organizations', { method:'POST', headers:{Authorization:'Bearer '+state.token}, body: formData });
      const result = await res.json();
      if (!res.ok) throw new Error(result.error || 'Could not create company');
      document.getElementById('createOrgErr').textContent = '';
      let msg = '<div class="card" style="background:#e2f1e8;margin-top:10px;"><strong>Company created.</strong>';
      if (result.admin) msg += '<br/>Admin login: '+result.admin.email+'<br/>Temporary password: <strong>'+result.admin.temporary_password+'</strong><br/><span class="muted">Share this with them directly — it will not be shown again.</span>';
      msg += '</div>';
      document.getElementById('createOrgResult').innerHTML = msg;
      ['org_name','org_email','org_phone','org_address','org_regnum','org_site_name','org_admin_name','org_admin_email'].forEach(id => document.getElementById(id).value = '');
      document.getElementById('org_logo').value = '';
      loadOrganizations();
    } catch(e) { document.getElementById('createOrgErr').textContent = e.message; }
  };
  loadOrganizations();
}
async function loadOrganizations(){
  const el = document.getElementById('orgList');
  try {
    const orgs = await api('/superadmin/organizations');
    if (!orgs.length) { el.textContent = 'No companies yet.'; return; }
    el.innerHTML = orgs.map(o =>
      '<div class="card" style="display:flex;align-items:center;gap:14px;margin-bottom:10px;padding:14px;">' +
        (o.logo_url ? '<img src="'+o.logo_url+'" style="width:48px;height:48px;border-radius:8px;object-fit:cover;" />' : '<div style="width:48px;height:48px;border-radius:8px;background:#eee;display:flex;align-items:center;justify-content:center;color:#999;font-size:11px;">No logo</div>') +
        '<div style="flex:1;"><strong>'+o.name+'</strong>'+(o.is_active===false?' <span class="badge rejected">inactive</span>':' <span class="badge paid">active</span>')+
        '<div class="muted">'+(o.contact_email||'')+(o.contact_phone?' · '+o.contact_phone:'')+'</div>' +
        '<div class="muted">'+o.site_count+' site(s) · '+o.user_count+' user(s)</div></div>' +
        '<label class="secondary" style="cursor:pointer;padding:8px 12px;border-radius:8px;border:1px solid #ccc;font-size:13px;margin-right:6px;">Change logo<input type="file" accept="image/png,image/jpeg" data-logo-org="'+o.id+'" style="display:none;" /></label>' +
        '<button class="'+(o.is_active===false?'success':'danger')+'" data-toggle-org="'+o.id+'" data-next="'+(o.is_active===false?'true':'false')+'">'+(o.is_active===false?'Activate':'Deactivate')+'</button>' +
      '</div>'
    ).join('');
    el.querySelectorAll('[data-logo-org]').forEach(inp => {
      inp.onchange = async () => {
        if (!inp.files[0]) return;
        const fd = new FormData(); fd.append('logo', inp.files[0]);
        try {
          const res = await fetch(API+'/superadmin/organizations/'+inp.dataset.logoOrg+'/logo', { method:'POST', headers:{Authorization:'Bearer '+state.token}, body: fd });
          const b = await res.json().catch(()=>({}));
          if (!res.ok) throw new Error(b.error || 'Upload failed');
          loadOrganizations();
        } catch(e) { alert(e.message); }
      };
    });
    el.querySelectorAll('[data-toggle-org]').forEach(btn => {
      btn.onclick = async () => {
        try { await api('/superadmin/organizations/'+btn.dataset.toggleOrg+'/status', {method:'PATCH', body: JSON.stringify({is_active: btn.dataset.next==='true'})}); loadOrganizations(); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) { el.textContent = 'Could not load companies.'; }
}
async function renderDashboard(){
  const {user} = state;
  appEl.innerHTML =
    '<div class="shell">' +
      '<div class="overlay" id="overlay"></div>' +
      '<aside class="sidebar" id="sidebar">' +
        '<div class="brand"><div class="brand-logo">'+TRUCK_SVG+'</div><div><div class="brand-title">Site Transactions</div><div class="brand-sub">Construction Portal</div></div></div>' +
        '<div class="nav-section-label">MAIN</div>' +
        (['manager','finance','admin'].includes(user.role) ? '<a class="nav-item" id="navDashboard" data-view="dashboard"><span class="nav-icon">&#8962;</span> Dashboard</a>' : '') +
        (['manager','finance','admin'].includes(user.role) ? '<a class="nav-item" id="navAnalytics" data-view="analytics"><span class="nav-icon">&#128202;</span> Analytics</a>' : '') +
        (['manager','finance','admin'].includes(user.role) ? '<a class="nav-item" id="navSuppliers" data-view="suppliers"><span class="nav-icon">&#127981;</span> Suppliers</a>' : '') +
        (['manager','finance','admin'].includes(user.role) ? '<a class="nav-item" id="navPayments" data-view="payments"><span class="nav-icon">&#128179;</span> Payments</a>' : '') +
        (['manager','admin'].includes(user.role) ? '<a class="nav-item" id="navApprovals" data-view="approvals"><span class="nav-icon">&#9989;</span> Approvals</a>' : '') +
        (['manager','admin'].includes(user.role) ? '<a class="nav-item" id="navReports" data-view="reports"><span class="nav-icon">&#128196;</span> Reports</a>' : '') +
        (['manager','admin'].includes(user.role) ? '<a class="nav-item" id="navAudit" data-view="audit"><span class="nav-icon">&#128214;</span> Audit trail</a>' : '') +
        (['clerk','finance','manager','admin'].includes(user.role) ? '<a class="nav-item" id="navTransactions" data-view="transactions"><span class="nav-icon">&#128203;</span> Transactions</a>' : '') +
        (['clerk','finance','manager','admin'].includes(user.role) ? '<a class="nav-item" id="navRegistry" data-view="registry"><span class="nav-icon">&#128101;</span> Registry</a>' : '') +
        (['storekeeper','finance','manager','admin'].includes(user.role) ? '<a class="nav-item" id="navStore" data-view="store"><span class="nav-icon">&#128230;</span> Store</a>' : '') +
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
    else if (view === 'store') html = buildStoreHtml(user);
    else if (view === 'analytics') html = buildAnalyticsHtml(user);
    else if (view === 'suppliers') html = buildSuppliersHtml(user);
    else if (view === 'payments') html = buildPaymentsHtml(user);
    else if (view === 'approvals') html = buildApprovalsHtml(user);
    else if (view === 'reports') html = buildReportsHtml(user);
    else if (view === 'audit') html = buildAuditHtml(user);
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
    if (document.getElementById('ledgerBox')) wirePayments();
    if (document.getElementById('approvalsBox') || document.getElementById('reportsBox') || document.getElementById('auditBox')) wireManagerTools();
    if (document.getElementById('queuePending')) loadApprovalQueues();
    if (document.getElementById('notifList')) loadNotifications();
    if (document.getElementById('userList')) loadUsers();
    if (document.getElementById('uploadLogoBtn')) {
      (async () => {
        try { const o = await api('/organization'); if (o.logo_url) document.getElementById('logoPreview').innerHTML = '<img src="'+o.logo_url+'" style="width:100%;height:100%;object-fit:cover;" />'; } catch(e){}
      })();
      document.getElementById('uploadLogoBtn').onclick = async () => {
        const f = document.getElementById('logoFile').files[0];
        const err = document.getElementById('logoErr'), ok = document.getElementById('logoOk');
        err.textContent = ''; ok.textContent = '';
        if (!f) { err.textContent = 'Choose an image first.'; return; }
        const fd = new FormData(); fd.append('logo', f);
        try {
          const res = await fetch(API+'/organization/logo', { method:'POST', headers:{Authorization:'Bearer '+state.token}, body: fd });
          const b = await res.json().catch(()=>({}));
          if (!res.ok) throw new Error(b.error || 'Upload failed');
          document.getElementById('logoPreview').innerHTML = '<img src="'+b.logo_url+'" style="width:100%;height:100%;object-fit:cover;" />';
          const bl = document.querySelector('.sidebar .brand-logo'); if (bl) bl.innerHTML = '<img src="'+b.logo_url+'" style="width:100%;height:100%;object-fit:cover;border-radius:50%;" />';
          ok.textContent = 'Logo updated. It will appear on new PDFs and receipts.';
        } catch(e) { err.textContent = e.message; }
      };
    }
    if (document.getElementById('sidebar') && state.user.organization_id && !window.__brandLoaded) {
      window.__brandLoaded = true;
      api('/organization').then(o => { const bl = document.querySelector('.sidebar .brand-logo'); if (o.logo_url && bl) bl.innerHTML = '<img src="'+o.logo_url+'" style="width:100%;height:100%;object-fit:cover;border-radius:50%;" />'; }).catch(()=>{});
    }
    if (document.getElementById('storeList')) loadStoreItems();
    if (document.getElementById('recordDeliveryBtn')) {
      (async () => {
        try {
          const [suppliers, items] = await Promise.all([
            api('/vendors').then(v => v.filter(x => x.role_type === 'vendor')),
            api('/store/items')
          ]);
          document.getElementById('dv_supplier').innerHTML += suppliers.map(s => '<option value="'+s.id+'">'+s.name+'</option>').join('');
          document.getElementById('dv_item').innerHTML += items.map(i => '<option value="'+i.id+'" data-unit="'+i.unit+'">'+i.name+' ('+i.unit+')</option>').join('');
        } catch {}
      })();
      document.getElementById('recordDeliveryBtn').onclick = async () => {
        const supplier_id = document.getElementById('dv_supplier').value;
        const new_supplier_name = document.getElementById('dv_new_supplier').value.trim();
        const item_id = document.getElementById('dv_item').value;
        const new_item_name = document.getElementById('dv_new_item').value.trim();
        const new_item_unit = document.getElementById('dv_new_unit').value.trim();
        const quantity = document.getElementById('dv_qty').value;
        const amount = document.getElementById('dv_amount').value;
        const description = document.getElementById('dv_note').value.trim();
        if (!supplier_id && !new_supplier_name) { document.getElementById('deliveryErr').textContent = 'Pick a supplier or type a new supplier name.'; return; }
        if (!item_id && !new_item_name) { document.getElementById('deliveryErr').textContent = 'Pick a material or type a new material name.'; return; }
        try {
          await api('/deliveries', {method:'POST', body: JSON.stringify({supplier_id, new_supplier_name, item_id, new_item_name, new_item_unit, quantity, amount, description})});
          document.getElementById('deliveryErr').textContent = '';
          const photo = document.getElementById('dv_photo');
          let photoMsg = '';
          if (photo && photo.files[0]) {
            const supSel = document.getElementById('dv_supplier'), itSel = document.getElementById('dv_item');
            const supName = new_supplier_name || (supSel.selectedOptions[0] ? supSel.selectedOptions[0].text : '');
            const itName = new_item_name || (itSel.selectedOptions[0] ? itSel.selectedOptions[0].text : '');
            const fd = new FormData(); fd.append('file', photo.files[0]); fd.append('description', 'Delivery note - ' + supName + ' - ' + itName + (description ? ' - ' + description : ''));
            try { const up = await fetch(API+'/documents', {method:'POST', headers:{Authorization:'Bearer '+state.token}, body: fd}); photoMsg = up.ok ? ' Delivery note photo saved under Documents.' : ' (The photo could not be saved - upload it under Documents.)'; } catch(e2) { photoMsg = ' (The photo could not be saved - upload it under Documents.)'; }
            photo.value = '';
          }
          document.getElementById('deliveryOk').textContent = 'Delivery recorded - stock updated and a pending payment was sent for approval.' + photoMsg;
          ['dv_new_supplier','dv_new_item','dv_new_unit','dv_qty','dv_amount','dv_note'].forEach(id => document.getElementById(id).value = '');
          loadStoreItems();
        } catch(e) { document.getElementById('deliveryErr').textContent = e.message; document.getElementById('deliveryOk').textContent=''; }
      };
    }
    if (document.getElementById('analyticsStatusChart')) loadAnalytics();
    if (document.getElementById('supplierList')) loadSuppliers();
    if (document.getElementById('addSupplierBtn')) {
      document.getElementById('addSupplierBtn').onclick = async () => {
        const name = document.getElementById('sup_name').value.trim();
        const contact_phone = document.getElementById('sup_phone').value.trim();
        const contact_email = document.getElementById('sup_email').value.trim();
        if (!name) { document.getElementById('addSupplierErr').textContent = 'Supplier name is required.'; return; }
        try {
          await api('/suppliers', {method:'POST', body: JSON.stringify({name, contact_phone, contact_email})});
          document.getElementById('addSupplierErr').textContent = '';
          document.getElementById('sup_name').value=''; document.getElementById('sup_phone').value=''; document.getElementById('sup_email').value='';
          loadSuppliers();
        } catch(e) { document.getElementById('addSupplierErr').textContent = e.message; }
      };
    }
    if (document.getElementById('downloadStockPdfBtn')) {
      document.getElementById('downloadStockPdfBtn').onclick = () => {
        fetch(API+'/store/report.pdf', {headers:{Authorization:'Bearer '+state.token}}).then(r=>r.blob()).then(blob=>{
          const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url;
          a.download = 'stock-report.pdf'; a.click();
        });
      };
    }
    if (document.getElementById('addItemMaterialBtn')) {
      document.getElementById('addItemMaterialBtn').onclick = async () => {
        const name = document.getElementById('item_name').value.trim();
        const unit = document.getElementById('item_unit').value.trim();
        const initial_quantity = document.getElementById('item_qty').value;
        const reorder_level = document.getElementById('item_reorder').value;
        try {
          await api('/store/items', {method:'POST', body: JSON.stringify({name, unit, initial_quantity, reorder_level})});
          document.getElementById('addItemErr').textContent = '';
          document.getElementById('item_name').value=''; document.getElementById('item_unit').value=''; document.getElementById('item_qty').value=''; document.getElementById('item_reorder').value='';
          loadStoreItems();
        } catch(e) { document.getElementById('addItemErr').textContent = e.message; }
      };
    }
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
    if (document.getElementById('downloadAttendanceReportBtn')) {
      const wsInput = document.getElementById('attReportWeekStart');
      const today = new Date();
      const day = today.getDay();
      const diffToMonday = day === 0 ? 6 : day - 1;
      const monday = new Date(today);
      monday.setDate(today.getDate() - diffToMonday);
      wsInput.value = monday.toISOString().slice(0,10);
      document.getElementById('downloadAttendanceReportBtn').onclick = () => {
        const start = document.getElementById('attReportWeekStart').value;
        fetch(API+'/registry/attendance-report.pdf?start='+start, {headers:{Authorization:'Bearer '+state.token}})
          .then(r=>r.blob()).then(blob=>{
            const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url;
            a.download = 'attendance-report-'+start+'.pdf'; a.click();
          });
      };
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
          if (document.getElementById('attDate')) loadWorkerAttendance();
          if (document.getElementById('attViewDate')) loadWorkerAttendanceView();
        } catch(e) { document.getElementById('enrollWorkerErr').textContent = e.message; }
      };
    }
    if (document.getElementById('submitAttendance')) {
      document.getElementById('submitAttendance').onclick = async () => {
        const attendance_date = document.getElementById('attDate').value;
        const worker_ids = Array.from(document.querySelectorAll('[data-worker-check]:checked')).map(cb => cb.dataset.workerCheck);
        if (worker_ids.length === 0 && !confirm('No one is marked present. Clear all attendance for this date?')) return;
        try {
          const result = await api('/attendance', {method:'POST', body: JSON.stringify({attendance_date, worker_ids, replace: true})});
          document.getElementById('attendanceErr').textContent = '';
          document.getElementById('attendanceOk').textContent = 'Attendance saved: '+result.recorded+' worker(s) present on '+result.date+'. You can change it and save again to correct mistakes.';
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
          '<div><label>Category</label><select id="itemCat-'+n+'"><option value="materials">Materials</option><option value="equipment">Equipment</option><option value="fuel">Fuel</option><option value="other">Other</option></select></div>' +
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
        saveDraft();
      }
      document.getElementById('addItemBtn').onclick = addItemRow;
      addItemRow(); // start with one row
      try {
        const dr = JSON.parse(localStorage.getItem(draftKey()) || 'null');
        if (dr && dr.length) {
          document.getElementById('itemRows').innerHTML = ''; itemCounter = 0;
          dr.forEach(function(it){ addItemRow(); document.getElementById('itemCat-'+itemCounter).value = it.c; document.getElementById('itemDesc-'+itemCounter).value = it.d; document.getElementById('itemAmt-'+itemCounter).value = it.a; });
          updateBatchTotal();
          document.getElementById('submitErr').textContent = 'Your unsent entries were restored.';
        }
      } catch(e){}
      document.getElementById('itemRows').addEventListener('input', saveDraft);
      document.getElementById('itemRows').addEventListener('change', saveDraft);
      showOutboxNote(); flushOutbox();

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
        } catch(e){
          if (e instanceof TypeError || !navigator.onLine) {
            const q = outboxGet(); q.push({site_id, items}); outboxSet(q);
            document.getElementById('itemRows').innerHTML = ''; itemCounter = 0; addItemRow(); updateBatchTotal();
            document.getElementById('submitErr').textContent = '';
            showOutboxNote();
          } else { document.getElementById('submitErr').textContent = e.message; }
        }
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
    if (document.getElementById('downloadFormBtn')) {
      document.getElementById('downloadFormBtn').onclick = () => {
        fetch(API+'/forms/transaction-log.pdf', {headers:{Authorization:'Bearer '+state.token}}).then(r=>r.blob()).then(blob=>{
          const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url;
          a.download = 'site-transaction-log-form.pdf'; a.click();
        });
      };
    }
    if (document.getElementById('downloadWorkerFormBtn')) {
      document.getElementById('downloadWorkerFormBtn').onclick = () => {
        fetch(API+'/forms/worker-registration.pdf', {headers:{Authorization:'Bearer '+state.token}}).then(r=>r.blob()).then(blob=>{
          const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url;
          a.download = 'worker-registration-form.pdf'; a.click();
        });
      };
    }
    if (document.getElementById('docList')) {
      loadDocuments();
      document.getElementById('uploadDocBtn').onclick = async () => {
        const fileInput = document.getElementById('docFile');
        const description = document.getElementById('docDesc').value.trim();
        if (!fileInput.files.length) { document.getElementById('docErr').textContent = 'Choose a file first.'; return; }
        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('description', description);
        const btn = document.getElementById('uploadDocBtn');
        btn.disabled = true; btn.textContent = 'Uploading & reading document...';
        try {
          const res = await fetch(API+'/documents', { method:'POST', headers:{Authorization:'Bearer '+state.token}, body: formData });
          if (!res.ok) { const b = await res.json().catch(()=>({})); throw new Error(b.error||'Upload failed'); }
          document.getElementById('docErr').textContent = '';
          fileInput.value=''; document.getElementById('docDesc').value='';
          loadDocuments();
        } catch(e) { document.getElementById('docErr').textContent = e.message; }
        finally { btn.disabled = false; btn.textContent = 'Upload'; }
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

  const defaultView = ['manager','finance','admin'].includes(user.role) ? 'dashboard' : (user.role === 'storekeeper' ? 'store' : 'transactions');
  const defaultNavId = { dashboard: 'navDashboard', store: 'navStore', transactions: 'navTransactions' }[defaultView];
  switchView(defaultView, document.getElementById(defaultNavId));
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
  window.__txnCache = {}; rows.forEach(t => { window.__txnCache[t.id] = t; });
  let html = '<table><thead><tr><th>Date</th><th>Category</th><th>Description</th><th>Amount</th><th>Status</th><th>Actions</th></tr></thead><tbody>';
  rows.forEach(t => {
    html += '<tr><td>'+t.transaction_date+'</td><td>'+t.category+'</td><td>'+escHtml(t.description)+(t.status==='rejected'&&t.rejection_reason?'<div style="color:#b03a3a;font-size:12px;margin-top:4px;">Rejected: '+escHtml(t.rejection_reason)+'</div>':'')+'</td><td>'+money(t.amount)+'</td><td>'+badge(t.status)+'</td><td class="actions">'+actionButtons(t,user)+' <button class="secondary" data-details="'+t.id+'">Details</button></td></tr>';
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
  document.querySelectorAll('[data-resubmit]').forEach(btn => {
    btn.onclick = async () => {
      const t = window.__txnCache[btn.dataset.resubmit]; if (!t) return;
      const description = prompt('Correct the description:', t.description); if (description === null) return;
      const amount = prompt('Correct the amount (KES):', t.amount); if (amount === null) return;
      try { await api('/transactions/'+t.id+'/resubmit', {method:'POST', body: JSON.stringify({description, amount})}); loadTable(); loadSummary(); }
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
  if (user.role==='clerk' && t.status==='rejected') b += '<button data-resubmit="'+t.id+'">Edit and resubmit</button>';
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
      '<li>'+v.vendor.name+' <span class="muted">('+v.vendor.role_type+')</span>'+(v.supplied_item?' — <em>'+v.supplied_item+'</em>':'')+(canManage?' <button class="danger" style="padding:2px 6px;font-size:11px;" data-remove-vendor="'+v.id+'">remove</button>':'')+'</li>'
    ).join('') + '</ul>' : '<div class="muted">No vendors assigned.</div>';

    if (canManage) {
      html += '<div class="row" style="margin-top:8px;"><select id="vendorSelect-'+txnId+'">' +
        allVendors.map(v => '<option value="'+v.id+'">'+v.name+' ('+v.role_type+')</option>').join('') +
        '</select><input id="suppliedItem-'+txnId+'" placeholder="What was supplied? (e.g. 50 bags cement)" /><button data-assign-vendor="'+txnId+'">Assign</button></div>';
      html += '<div class="muted" style="margin-top:6px;">New vendor: <input id="newVendorName-'+txnId+'" placeholder="Name" style="width:120px;" /> <select id="newVendorType-'+txnId+'"><option value="vendor">Vendor</option><option value="inspector">Inspector</option><option value="legal_counsel">Legal Counsel</option><option value="underwriter">Underwriter</option><option value="other">Other</option></select> <button class="secondary" data-add-vendor="'+txnId+'">Add</button></div>';
    }
    body.innerHTML = html;

    body.querySelectorAll('[data-remove-vendor]').forEach(btn => {
      btn.onclick = async () => { try { await api('/transaction-vendors/'+btn.dataset.removeVendor, {method:'DELETE'}); toggleDetails(txnId); openDetails[txnId]=false; toggleDetails(txnId); } catch(e){ alert(e.message); } };
    });
    const assignBtn = body.querySelector('[data-assign-vendor]');
    if (assignBtn) assignBtn.onclick = async () => {
      const vendor_id = document.getElementById('vendorSelect-'+txnId).value;
      const supplied_item = document.getElementById('suppliedItem-'+txnId).value.trim();
      try { await api('/transactions/'+txnId+'/vendors', {method:'POST', body: JSON.stringify({vendor_id, supplied_item})}); openDetails[txnId]=false; toggleDetails(txnId); }
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
function animateCount(el, target){
  const start = 0;
  const duration = 700;
  const startTime = performance.now();
  function step(now){
    const progress = Math.min((now - startTime) / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    el.textContent = money(start + (target - start) * eased);
    if (progress < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}
function donutChart(segments, size){
  size = size || 120;
  const r = size/2 - 10, cx = size/2, cy = size/2, circ = 2*Math.PI*r;
  const total = segments.reduce((s,seg)=>s+seg.value, 0) || 1;
  let offset = 0;
  const arcs = segments.filter(s=>s.value>0).map(seg => {
    const frac = seg.value/total;
    const dash = frac*circ;
    const el = '<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="'+seg.color+'" stroke-width="14" stroke-dasharray="'+dash+' '+(circ-dash)+'" stroke-dashoffset="'+(-offset)+'" transform="rotate(-90 '+cx+' '+cy+')" stroke-linecap="butt"><animate attributeName="stroke-dashoffset" from="0" to="'+(-offset)+'" dur=".6s" fill="freeze"/></circle>';
    offset += dash;
    return el;
  }).join('');
  return '<svg width="'+size+'" height="'+size+'" viewBox="0 0 '+size+' '+size+'">'+
    '<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="#eee" stroke-width="14"/>'+arcs+'</svg>';
}
function simpleBarChart(rows, maxWidth){
  maxWidth = maxWidth || 100;
  const max = Math.max(...rows.map(r=>r.value), 1);
  return rows.map(r =>
    '<div style="margin-bottom:8px;"><div class="muted" style="display:flex;justify-content:space-between;font-size:12px;"><span>'+r.label+'</span><span>'+money(r.value)+'</span></div>' +
    '<div style="background:#eee;border-radius:6px;height:10px;overflow:hidden;"><div style="width:'+Math.max((r.value/max)*maxWidth,2)+'%;height:100%;background:'+(r.color||'#2f7d4f')+';border-radius:6px;transition:width .6s ease;"></div></div></div>'
  ).join('');
}
let openSupplierDetails = {};
function downloadAuth(path, filename){
  fetch(API+path, {headers:{Authorization:'Bearer '+state.token}}).then(function(r){ if(!r.ok) throw new Error('Download failed'); return r.blob(); }).then(function(blob){
    var url = URL.createObjectURL(blob); var a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function(){ URL.revokeObjectURL(url); }, 2000);
  }).catch(function(e){ alert(e.message); });
}
function escHtml(s){ return String(s==null?'':s).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function buildPaymentsHtml(user){
  var canBudget = ['manager','admin'].indexOf(user.role) >= 0;
  var h = '';
  h += '<div class="card"><h2>Site budget</h2><div id="budgetBox"><div class="skeleton" style="width:70%;"></div></div>';
  if (canBudget) h += '<div class="row" style="margin-top:12px;"><div><label>Total budget (KES)</label><input id="budgetAmt" type="number" step="0.01" placeholder="e.g. 5000000" /></div><div><label>Warn me at (% used)</label><input id="budgetWarn" type="number" min="1" max="100" value="80" /></div></div><div style="margin-top:10px;"><button id="saveBudgetBtn">Save budget</button></div><div class="error" id="budgetErr"></div>';
  h += '</div>';
  h += '<div class="card" id="ledgerBox"><h2>Payments ledger</h2><p class="muted">Every payment made: expenses and worker wages. Search by supplier, worker or description, and filter by date.</p>' +
    '<div class="row"><div><label>From</label><input id="ledgerFrom" type="date" /></div><div><label>To</label><input id="ledgerTo" type="date" /></div><div><label>Search</label><input id="ledgerQ" placeholder="supplier, worker, cement..." /></div></div>' +
    '<div style="margin:10px 0;display:flex;gap:8px;flex-wrap:wrap;"><button id="ledgerSearchBtn">Search</button><button class="secondary" id="ledgerXlsBtn">Download Excel</button><button class="secondary" id="ledgerPdfBtn">Download PDF</button></div>' +
    '<div id="ledgerTable"><div class="skeleton" style="width:85%;"></div></div></div>';
  h += '<div class="card"><h2>Wage payment results</h2><p class="muted">See who was paid by M-Pesa and who failed.</p>' +
    '<div class="row"><div><label>Week starting (Monday) - leave empty for this week</label><input id="wrStart" type="date" /></div></div>' +
    '<div style="margin:10px 0;display:flex;gap:8px;flex-wrap:wrap;"><button class="secondary" id="wrRefreshBtn">Show results</button><button id="wrRetryBtn" style="display:none;">Retry failed payments</button></div>' +
    '<div id="wrBox"><div class="skeleton" style="width:80%;"></div></div></div>';
  h += '<div class="card"><h2>Finance summary report (PDF)</h2><p class="muted">Pending, approved, paid and rejected totals for a day, week or month.</p>' +
    '<div class="row"><div><label>Report type</label><select id="sumPeriod"><option value="daily">Daily</option><option value="weekly" selected>Weekly (last 7 days)</option><option value="monthly">Monthly</option></select></div><div><label>Up to date</label><input id="sumDate" type="date" /></div></div>' +
    '<div style="margin-top:10px;"><button id="sumPdfBtn">Download summary PDF</button></div></div>';
  return h;
}
async function loadBudget(){
  var el = document.getElementById('budgetBox'); if (!el) return;
  try {
    var b = await api('/budget');
    var amtEl = document.getElementById('budgetAmt'), warnEl = document.getElementById('budgetWarn');
    if (amtEl && b.budget_amount != null) { amtEl.value = b.budget_amount; warnEl.value = b.warn_percent; }
    if (b.budget_amount == null) { el.innerHTML = '<p class="muted">No budget set for ' + escHtml(b.site_name) + ' yet.' + (amtEl ? ' Enter one below.' : ' Ask your manager to set it.') + '</p>'; return; }
    var pct = Math.min(100, b.percent_used || 0);
    var color = b.level === 'over' ? '#b03a3a' : b.level === 'warning' ? '#d98c2b' : '#2f7d4f';
    var html = '';
    if (b.level === 'over') html += '<div style="background:#fbeaea;color:#b03a3a;padding:10px 12px;border-radius:8px;margin-bottom:10px;font-weight:600;">Over budget: spending has passed the budget for ' + escHtml(b.site_name) + '.</div>';
    else if (b.level === 'warning') html += '<div style="background:#fdf3e3;color:#a36413;padding:10px 12px;border-radius:8px;margin-bottom:10px;font-weight:600;">Warning: ' + b.percent_used.toFixed(0) + '% of the budget is used up.</div>';
    html += '<div style="background:#e8ece9;border-radius:8px;height:18px;overflow:hidden;"><div style="width:' + pct + '%;height:100%;background:' + color + ';"></div></div>';
    html += '<div class="row" style="margin-top:12px;">' +
      '<div><div class="muted">Budget</div><strong>' + money(b.budget_amount) + '</strong></div>' +
      '<div><div class="muted">Paid</div><strong>' + money(b.paid) + '</strong></div>' +
      '<div><div class="muted">Approved, awaiting payment</div><strong>' + money(b.approved_unpaid) + '</strong></div>' +
      '<div><div class="muted">Pending approval</div><strong>' + money(b.pending) + '</strong></div>' +
      '<div><div class="muted">Remaining</div><strong style="color:' + color + ';">' + money(b.remaining) + '</strong></div></div>';
    el.innerHTML = html;
  } catch(e) { el.textContent = 'Could not load budget.'; }
}
async function loadLedger(){
  var el = document.getElementById('ledgerTable'); if (!el) return;
  var qs = ledgerQuery();
  try {
    var d = await api('/ledger' + qs);
    if (!d.rows.length) { el.innerHTML = '<p class="muted">No payments found.</p>'; return; }
    el.innerHTML = '<div class="muted" style="margin-bottom:6px;">' + d.rows.length + ' payments, total <strong>' + money(d.total) + '</strong></div><div style="overflow-x:auto;"><table><thead><tr><th>Date</th><th>Type</th><th>Paid to</th><th>Description</th><th>Method</th><th style="text-align:right;">Amount</th></tr></thead><tbody>' +
      d.rows.map(function(r){ return '<tr><td>' + escHtml(r.date ? String(r.date).slice(0,10) : '') + '</td><td>' + escHtml(r.type) + '</td><td>' + escHtml(r.payee) + '</td><td>' + escHtml(r.description) + '</td><td>' + escHtml(r.method) + (r.reference ? '<br/><span class="muted">' + escHtml(r.reference) + '</span>' : '') + '</td><td style="text-align:right;font-weight:600;">' + money(r.amount) + '</td></tr>'; }).join('') +
      '</tbody></table></div>';
  } catch(e) { el.textContent = 'Could not load the ledger.'; }
}
function ledgerQuery(){
  var p = [];
  var f = document.getElementById('ledgerFrom').value, t = document.getElementById('ledgerTo').value, q = document.getElementById('ledgerQ').value.trim();
  if (f) p.push('from=' + encodeURIComponent(f)); if (t) p.push('to=' + encodeURIComponent(t)); if (q) p.push('q=' + encodeURIComponent(q));
  return p.length ? '?' + p.join('&') : '';
}
async function loadWageResults(){
  var el = document.getElementById('wrBox'); if (!el) return;
  var s = document.getElementById('wrStart').value;
  try {
    var d = await api('/wages/results' + (s ? '?start=' + encodeURIComponent(s) : ''));
    var retryBtn = document.getElementById('wrRetryBtn');
    var canRetry = ['finance','admin'].indexOf(state.user.role) >= 0;
    if (retryBtn) { retryBtn.style.display = (canRetry && d.counts.failed > 0) ? 'inline-block' : 'none'; retryBtn.dataset.start = d.start; }
    if (!d.rows.length) { el.innerHTML = '<p class="muted">No wage payments recorded for the week of ' + d.start + '.</p>'; return; }
    var label = { confirmed:['paid','Paid - M-Pesa confirmed'], sent:['pending','Sent - awaiting confirmation'], failed:['rejected','FAILED'], not_attempted:['approved','Recorded - no M-Pesa sent'] };
    el.innerHTML = '<div class="muted" style="margin-bottom:6px;">Week ' + d.start + ' to ' + d.end + ' - Confirmed: ' + (d.counts.confirmed||0) + ', awaiting: ' + (d.counts.sent||0) + ', failed: ' + (d.counts.failed||0) + ', recorded only: ' + (d.counts.not_attempted||0) + (d.counts.failed ? ' (' + money(d.failed_total) + ' not delivered)' : '') + '</div>' +
      '<div style="overflow-x:auto;"><table><thead><tr><th>Worker</th><th>Phone</th><th>Days</th><th>Amount</th><th>Status</th><th>Receipt / note</th></tr></thead><tbody>' +
      d.rows.map(function(r){ var l = label[r.disbursement_status] || ['pending', r.disbursement_status]; return '<tr><td>' + escHtml(r.worker ? r.worker.name : '') + '</td><td>' + escHtml(r.worker ? r.worker.phone_number : '') + '</td><td>' + r.days_present + '</td><td>' + money(r.total_amount) + '</td><td><span class="badge ' + l[0] + '">' + l[1] + '</span></td><td class="muted">' + escHtml(r.mpesa_receipt || r.mpesa_result_desc || '') + '</td></tr>'; }).join('') + '</tbody></table></div>';
  } catch(e) { el.textContent = 'Could not load wage results.'; }
}
function wirePayments(){
  var today = new Date().toISOString().slice(0,10);
  var sd = document.getElementById('sumDate'); if (sd) sd.value = today;
  loadBudget(); loadLedger(); loadWageResults();
  var sb = document.getElementById('saveBudgetBtn');
  if (sb) sb.onclick = async function(){
    var err = document.getElementById('budgetErr'); err.textContent = '';
    try { await api('/budget', {method:'PUT', body: JSON.stringify({budget_amount: document.getElementById('budgetAmt').value, warn_percent: document.getElementById('budgetWarn').value})}); loadBudget(); }
    catch(e){ err.textContent = e.message; }
  };
  document.getElementById('ledgerSearchBtn').onclick = loadLedger;
  document.getElementById('ledgerXlsBtn').onclick = function(){ downloadAuth('/ledger/export' + ledgerQuery(), 'payments-ledger.xlsx'); };
  document.getElementById('ledgerPdfBtn').onclick = function(){ downloadAuth('/ledger.pdf' + ledgerQuery(), 'payments-ledger.pdf'); };
  document.getElementById('wrRefreshBtn').onclick = loadWageResults;
  document.getElementById('wrRetryBtn').onclick = async function(){
    if (!confirm('Retry the failed M-Pesa payments for this week? Workers who were already paid will not be paid again.')) return;
    try {
      var r = await api('/wages/weekly/retry-failed', {method:'POST', body: JSON.stringify({start: this.dataset.start})});
      alert(r.message || ('Retried ' + r.retried + ' payment(s); ' + r.sent + ' sent. Check the status column.'));
      loadWageResults();
    } catch(e){ alert(e.message); }
  };
  document.getElementById('sumPdfBtn').onclick = function(){
    var p = document.getElementById('sumPeriod').value, d = document.getElementById('sumDate').value || today;
    downloadAuth('/reports/finance-summary.pdf?period=' + p + '&date=' + d, 'finance-summary-' + p + '-' + d + '.pdf');
  };
}
// ----- Clerk: drafts and offline outbox -----
function draftKey(){ return 'cst_draft_' + (state.user ? state.user.id : ''); }
function outboxKey(){ return 'cst_outbox_' + (state.user ? state.user.id : ''); }
function saveDraft(){
  try {
    var items = [];
    document.querySelectorAll('#itemRows > div').forEach(function(row){
      var n = row.id.replace('itemRow-','');
      items.push({c: document.getElementById('itemCat-'+n).value, d: document.getElementById('itemDesc-'+n).value, a: document.getElementById('itemAmt-'+n).value});
    });
    var any = items.some(function(i){ return i.d || i.a; });
    if (any) localStorage.setItem(draftKey(), JSON.stringify(items)); else localStorage.removeItem(draftKey());
  } catch(e){}
}
function outboxGet(){ try { return JSON.parse(localStorage.getItem(outboxKey()) || '[]'); } catch(e){ return []; } }
function outboxSet(a){ try { localStorage.setItem(outboxKey(), JSON.stringify(a)); } catch(e){} }
function showOutboxNote(){
  var el = document.getElementById('outboxNote'); if (!el) return;
  var n = outboxGet().length;
  el.textContent = n ? (n + ' entr' + (n === 1 ? 'y is' : 'ies are') + ' saved on this phone and will be sent automatically when the network is back.') : '';
}
async function flushOutbox(){
  if (!state.token || !state.user || !navigator.onLine) return;
  var q = outboxGet(); if (!q.length) return;
  var left = [], sent = 0;
  for (var i = 0; i < q.length; i++) {
    try { await api('/batches', {method:'POST', body: JSON.stringify(q[i])}); sent++; }
    catch(e) {
      if (e instanceof TypeError) { left = left.concat(q.slice(i)); break; }
      alert('A saved entry could not be sent: ' + e.message);
    }
  }
  outboxSet(left); showOutboxNote();
  if (sent) { if (document.getElementById('table')) loadTable(); }
}
window.addEventListener('online', flushOutbox);

// ----- Workers: edit and history -----
async function editWorker(id, done){
  var w = (window.__workers || []).filter(function(x){ return x.id === id; })[0]; if (!w) return;
  var name = prompt('Full name:', w.name); if (name === null) return;
  var phone = prompt('Phone number (07xxxxxxxx):', w.phone_number || ''); if (phone === null) return;
  var designation = prompt('Designation:', w.designation || ''); if (designation === null) return;
  var rate = prompt('Daily rate (KES):', w.daily_rate || 0); if (rate === null) return;
  try { await api('/workers/' + id, {method:'PATCH', body: JSON.stringify({name: name, phone_number: phone, designation: designation, daily_rate: rate})}); if (done) done(); }
  catch(e){ alert(e.message); }
}
async function showWorkerHistory(id){
  try {
    var h = await api('/workers/' + id + '/history');
    var w = h.worker;
    var ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9998;display:flex;align-items:flex-start;justify-content:center;overflow:auto;padding:20px;';
    var wages = h.wages.length ? '<div style="overflow-x:auto;"><table><thead><tr><th>Week</th><th>Days</th><th>Amount</th><th>Status</th></tr></thead><tbody>' + h.wages.map(function(x){ return '<tr><td>' + x.week_start + '</td><td>' + x.days_present + '</td><td>' + money(x.total_amount) + '</td><td>' + escHtml(x.disbursement_status || '') + (x.mpesa_receipt ? '<br/><span class="muted">' + escHtml(x.mpesa_receipt) + '</span>' : '') + '</td></tr>'; }).join('') + '</tbody></table></div>' : '<p class="muted">No wage payments yet.</p>';
    var recent = h.attendance.slice(0, 60).map(function(d){ return '<span class="badge paid" style="margin:2px;display:inline-block;">' + d + '</span>'; }).join('');
    ov.innerHTML = '<div class="card" style="max-width:640px;width:100%;margin-top:20px;"><div class="topbar"><h2>' + escHtml(w.name) + '</h2><button class="secondary" id="closeHist">Close</button></div>' +
      '<p class="muted">' + escHtml(w.designation || '') + ' - ID ' + escHtml(w.id_number || '') + ' - ' + escHtml(w.phone_number || 'no phone') + ' - ' + money(w.daily_rate) + ' per day</p>' +
      '<div class="row"><div><div class="muted">Days present (on record)</div><strong>' + h.days_total + '</strong></div><div><div class="muted">Total paid</div><strong>' + money(h.paid_total) + '</strong></div></div>' +
      '<h3 style="margin-top:14px;">Wage payments</h3>' + wages + '<h3 style="margin-top:14px;">Recent days present</h3><div>' + (recent || '<span class="muted">None yet.</span>') + '</div></div>';
    document.body.appendChild(ov);
    ov.querySelector('#closeHist').onclick = function(){ ov.remove(); };
    ov.onclick = function(e){ if (e.target === ov) ov.remove(); };
  } catch(e){ alert(e.message); }
}

// ----- Manager: approvals inbox, reports, audit trail -----
function buildApprovalsHtml(user){
  return '<div class="card" id="approvalsBox"><h2>Approvals inbox</h2><p class="muted">Everything waiting for your approval. Tick items, then approve or reject them together.</p><div id="approvalsList"><div class="skeleton" style="width:80%;"></div></div>' +
    '<div class="topbar" style="margin-top:12px;"><strong>Selected: <span id="apSelTotal">KES 0.00</span></strong><div style="display:flex;gap:8px;flex-wrap:wrap;"><button class="success" id="apApproveBtn">Approve selected</button><button class="danger" id="apRejectBtn">Reject selected</button></div></div><div class="error" id="apErr"></div><div class="muted" id="apOk"></div></div>';
}
async function loadApprovals(){
  var el = document.getElementById('approvalsList'); if (!el) return;
  try {
    var r = await api('/transactions?status=pending&limit=200');
    var rows = r.rows.filter(function(t){ return !t.manager_approved; });
    window.__apRows = rows;
    if (!rows.length) { el.innerHTML = '<p class="muted">Nothing is waiting for your approval.</p>'; apTotal(); return; }
    var all = rows.reduce(function(s, t){ return s + Number(t.amount); }, 0);
    el.innerHTML = '<div class="muted" style="margin-bottom:6px;">' + rows.length + ' item(s), total ' + money(all) + '</div><div style="overflow-x:auto;"><table><thead><tr><th><input type="checkbox" id="apAll" /></th><th>Date</th><th>Category</th><th>Description</th><th style="text-align:right;">Amount</th></tr></thead><tbody>' +
      rows.map(function(t){ return '<tr><td><input type="checkbox" data-ap="' + t.id + '" /></td><td>' + t.transaction_date + '</td><td>' + escHtml(t.category) + '</td><td>' + escHtml(t.description) + '</td><td style="text-align:right;font-weight:600;">' + money(t.amount) + '</td></tr>'; }).join('') + '</tbody></table></div>';
    document.getElementById('apAll').onchange = function(){ var on = this.checked; el.querySelectorAll('[data-ap]').forEach(function(c){ c.checked = on; }); apTotal(); };
    el.querySelectorAll('[data-ap]').forEach(function(c){ c.onchange = apTotal; });
    apTotal();
  } catch(e){ el.textContent = 'Could not load approvals.'; }
}
function apSelected(){ return Array.prototype.slice.call(document.querySelectorAll('[data-ap]:checked')).map(function(c){ return c.dataset.ap; }); }
function apTotal(){
  var ids = apSelected(), sum = 0;
  (window.__apRows || []).forEach(function(t){ if (ids.indexOf(t.id) >= 0) sum += Number(t.amount); });
  var el = document.getElementById('apSelTotal'); if (el) el.textContent = money(sum) + ' (' + ids.length + ')';
}
async function apRun(action){
  var ids = apSelected(), err = document.getElementById('apErr'), ok = document.getElementById('apOk');
  err.textContent = ''; ok.textContent = '';
  if (!ids.length) { err.textContent = 'Tick at least one item first.'; return; }
  var body = {};
  if (action === 'reject') { var reason = prompt('Reason for rejecting (the clerk will see this):'); if (reason === null) return; body.reason = reason; }
  else if (!confirm('Approve ' + ids.length + ' item(s)?')) return;
  var done = 0, failed = 0;
  for (var i = 0; i < ids.length; i++) {
    try { await api('/transactions/' + ids[i] + '/' + (action === 'reject' ? 'reject' : 'approve/manager'), {method:'POST', body: JSON.stringify(body)}); done++; } catch(e){ failed++; }
  }
  ok.textContent = done + ' item(s) ' + (action === 'reject' ? 'rejected' : 'approved') + (failed ? ', ' + failed + ' could not be updated.' : '.');
  loadApprovals();
}
function buildReportsHtml(user){
  return '<div class="card" id="reportsBox"><h2>Site activity report (PDF)</h2><p class="muted">Attendance, expenses, deliveries and stock used, in one printable report.</p>' +
    '<div class="row"><div><label>Report type</label><select id="actPeriod"><option value="daily">Daily</option><option value="weekly" selected>Weekly (last 7 days)</option></select></div><div><label>Up to date</label><input id="actDate" type="date" /></div></div>' +
    '<div style="margin-top:10px;"><button id="actPdfBtn">Download site activity report</button></div></div>';
}
function buildAuditHtml(user){
  return '<div class="card" id="auditBox"><h2>Audit trail</h2><p class="muted">Who approved, edited, deleted or changed what, and when.</p>' +
    '<div class="row"><div><label>From</label><input id="auFrom" type="date" /></div><div><label>To</label><input id="auTo" type="date" /></div><div><label>Search</label><input id="auQ" placeholder="name, action or detail" /></div></div>' +
    '<div style="margin:10px 0;"><button id="auGo">Search</button></div><div id="auList"><div class="skeleton" style="width:80%;"></div></div></div>';
}
async function loadAudit(){
  var el = document.getElementById('auList'); if (!el) return;
  var p = [], f = document.getElementById('auFrom').value, t = document.getElementById('auTo').value, q = document.getElementById('auQ').value.trim();
  if (f) p.push('from=' + encodeURIComponent(f)); if (t) p.push('to=' + encodeURIComponent(t)); if (q) p.push('q=' + encodeURIComponent(q));
  try {
    var rows = await api('/audit' + (p.length ? '?' + p.join('&') : ''));
    if (!rows.length) { el.innerHTML = '<p class="muted">No activity found. Actions are recorded from now on.</p>'; return; }
    el.innerHTML = '<div style="overflow-x:auto;"><table><thead><tr><th>When</th><th>Who</th><th>Action</th><th>Details</th></tr></thead><tbody>' +
      rows.map(function(r){ return '<tr><td>' + escHtml(new Date(r.created_at).toLocaleString()) + '</td><td>' + escHtml(r.actor_name || '') + '<br/><span class="muted">' + escHtml(r.actor_role || '') + '</span></td><td>' + escHtml(String(r.action).split('_').join(' ')) + '</td><td class="muted">' + escHtml(r.detail || '') + '</td></tr>'; }).join('') + '</tbody></table></div>';
  } catch(e){ el.textContent = 'Could not load the audit trail.'; }
}
function wireManagerTools(){
  var ab = document.getElementById('apApproveBtn');
  if (ab) { loadApprovals(); ab.onclick = function(){ apRun('approve'); }; document.getElementById('apRejectBtn').onclick = function(){ apRun('reject'); }; }
  var ap = document.getElementById('actPdfBtn');
  if (ap) {
    document.getElementById('actDate').value = new Date().toISOString().slice(0,10);
    ap.onclick = function(){ var p = document.getElementById('actPeriod').value, d = document.getElementById('actDate').value; downloadAuth('/reports/site-activity.pdf?period=' + p + '&date=' + d, 'site-activity-' + p + '-' + d + '.pdf'); };
  }
  var ag = document.getElementById('auGo');
  if (ag) { loadAudit(); ag.onclick = loadAudit; }
}
async function loadSuppliers(){
  const el = document.getElementById('supplierList');
  const canDelete = ['manager','admin'].includes(state.user.role);
  try {
    const suppliers = await api('/suppliers');
    if (!suppliers.length) { el.textContent = 'No suppliers added yet.'; return; }
    let html = '<table><thead><tr><th>Supplier</th><th>Contact</th><th>Total Paid</th><th>Total Pending</th><th>Transactions</th><th>Actions</th></tr></thead><tbody>';
    suppliers.forEach(s => {
      html += '<tr><td>'+s.name+'</td><td class="muted">'+(s.contact_phone||'')+(s.contact_email?' / '+s.contact_email:'')+'</td><td style="color:#2f7d4f;font-weight:600;">'+money(s.total_paid)+'</td><td style="color:#d98c2b;font-weight:600;">'+money(s.total_pending)+'</td><td>'+s.transaction_count+'</td><td class="actions"><button class="secondary" data-supplier-details="'+s.id+'">Details</button> '+'<button class="secondary" data-supplier-statement="'+s.id+'" data-supplier-name="'+String(s.name).replace(/"/g,'')+'">Statement PDF</button> '+(canDelete?'<button class="danger" data-del-supplier="'+s.id+'">Remove</button>':'')+'</td></tr>';
      html += '<tr id="supplier-details-'+s.id+'" style="display:none;"><td colspan="6"><div id="supplier-details-body-'+s.id+'" class="muted">Loading...</div></td></tr>';
    });
    html += '</tbody></table>';
    el.innerHTML = html;
    el.querySelectorAll('[data-del-supplier]').forEach(btn => {
      btn.onclick = async () => {
        if (!confirm('Remove this supplier? It will be unlinked from any transactions.')) return;
        try { await api('/suppliers/'+btn.dataset.delSupplier, {method:'DELETE'}); loadSuppliers(); }
        catch(e){ alert(e.message); }
      };
    });
    el.querySelectorAll('[data-supplier-statement]').forEach(btn => {
      btn.onclick = () => downloadAuth('/suppliers/'+btn.dataset.supplierStatement+'/statement.pdf', 'supplier-statement-'+btn.dataset.supplierName+'.pdf');
    });
    el.querySelectorAll('[data-supplier-details]').forEach(btn => {
      btn.onclick = () => toggleSupplierDetails(btn.dataset.supplierDetails);
    });
  } catch(e) { el.textContent = 'Could not load suppliers.'; }
}
async function toggleSupplierDetails(id){
  const row = document.getElementById('supplier-details-'+id);
  if (openSupplierDetails[id]) { row.style.display='none'; openSupplierDetails[id]=false; return; }
  row.style.display='table-row'; openSupplierDetails[id]=true;
  const body = document.getElementById('supplier-details-body-'+id);
  body.innerHTML = '<div class="skeleton" style="width:70%;"></div>';
  try {
    const [txns, deliveries] = await Promise.all([
      api('/suppliers/'+id+'/transactions'),
      api('/suppliers/'+id+'/deliveries')
    ]);
    let html = '';
    html += '<strong>Deliveries to site</strong>';
    html += deliveries.length ? '<table style="margin-top:6px;margin-bottom:14px;"><thead><tr><th>Date</th><th>Material</th><th>Quantity</th><th>Recorded by</th><th>Note</th></tr></thead><tbody>' +
      deliveries.map(d => '<tr><td>'+new Date(d.created_at).toLocaleDateString()+'</td><td>'+d.item.name+'</td><td>'+d.quantity+' '+d.item.unit+'</td><td class="muted">'+(d.recorder?d.recorder.full_name:'—')+'</td><td class="muted">'+(d.reason||'')+'</td></tr>').join('') +
      '</tbody></table>' : '<div class="muted" style="margin-bottom:14px;">No deliveries recorded yet — record one under Store, "Stock in (delivery)".</div>';

    html += '<strong>Linked transactions</strong>';
    html += txns.length ? '<table style="margin-top:6px;"><thead><tr><th>Date</th><th>Supplied</th><th>Category</th><th>Amount</th><th>Status</th></tr></thead><tbody>' +
      txns.map(t => '<tr><td>'+t.transaction_date+'</td><td>'+(t.supplied_item?'<strong>'+t.supplied_item+'</strong><br/><span class="muted">'+t.description+'</span>':t.description)+'</td><td>'+t.category+'</td><td>'+money(t.amount)+'</td><td>'+badge(t.status)+'</td></tr>').join('') +
      '</tbody></table>' : '<div class="muted">Nothing linked yet — assign this supplier under a transaction\\'s Details panel.</div>';

    body.innerHTML = html;
  } catch(e) { body.textContent = 'Could not load supplier details.'; }
}
async function loadAnalytics(){
  try {
    const { financial, operational } = await api('/analytics');

    const statusColors = { pending:'#d98c2b', approved:'#4a9d68', paid:'#2f7d4f', rejected:'#b03a3a' };
    const statusRows = Object.entries(financial.by_status).map(([k,v]) => ({label:k, value:v, color:statusColors[k]}));
    document.getElementById('analyticsStatusChart').innerHTML =
      '<div style="display:flex;gap:20px;flex-wrap:wrap;align-items:center;"><div style="flex-shrink:0;">'+donutChart(statusRows.map(r=>({value:r.value,color:r.color})))+'</div><div style="flex:1;min-width:200px;">'+simpleBarChart(statusRows)+'</div></div>';

    const catColors = ['#2f7d4f','#d98c2b','#4a7fb0','#8a5fb0','#b0895f'];
    const catRows = Object.entries(financial.by_category).map(([k,v],i) => ({label:k, value:v, color:catColors[i%catColors.length]}));
    document.getElementById('analyticsCategoryChart').innerHTML = catRows.length ? simpleBarChart(catRows) : '<p class="muted">No transactions recorded yet.</p>';

    const trendMax = Math.max(...financial.weekly_trend.map(w=>w.total), 1);
    document.getElementById('analyticsTrendChart').innerHTML = '<div style="display:flex;align-items:flex-end;gap:8px;height:120px;">' +
      financial.weekly_trend.map(w => {
        const h = Math.max((w.total/trendMax)*100, 2);
        return '<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;"><div style="width:100%;background:#2f7d4f;border-radius:4px 4px 0 0;height:'+h+'%;transition:height .6s ease;" title="'+money(w.total)+'"></div><div class="muted" style="font-size:10px;margin-top:4px;">'+w.start.slice(5)+'</div></div>';
      }).join('') + '</div>';

    document.getElementById('analyticsWages').innerHTML = '<div class="stat-num" style="color:#2f7d4f;">'+money(financial.wages_paid_total)+'</div><p class="muted">Total wages disbursed across all recorded weekly payments.</p>';

    if (operational && document.getElementById('analyticsOperational')) {
      document.getElementById('analyticsOperational').innerHTML =
        '<div class="row">' +
        '<div class="stat-card"><div class="muted">Workers enrolled</div><div class="stat-num">'+operational.worker_count+'</div></div>' +
        '<div class="stat-card"><div class="muted">Present today</div><div class="stat-num">'+operational.present_today+'/'+operational.worker_count+'</div></div>' +
        '<div class="stat-card"><div class="muted">Materials tracked</div><div class="stat-num">'+operational.store_item_count+'</div></div>' +
        '<div class="stat-card"><div class="muted">Low stock items</div><div class="stat-num" style="color:'+(operational.low_stock_count?'#b03a3a':'#2f7d4f')+';">'+operational.low_stock_count+'</div></div>' +
        '<div class="stat-card"><div class="muted">Documents uploaded</div><div class="stat-num">'+operational.documents_count+'</div></div>' +
        '<div class="stat-card"><div class="muted">Documents read by AI</div><div class="stat-num">'+operational.documents_read_count+'</div></div>' +
        '</div>';
    }
  } catch(e) {
    ['analyticsStatusChart','analyticsCategoryChart','analyticsTrendChart','analyticsWages','analyticsOperational'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = 'Could not load analytics.';
    });
  }
}
async function loadApprovalQueues(){
  const user = state.user;
  const canApproveManager = ['manager','admin'].includes(user.role);
  const canApproveFinance = ['finance','admin'].includes(user.role);
  try {
    const { rows } = await api('/transactions?status=pending&limit=100');
    const pendingEl = document.getElementById('queuePending');
    if (!rows.length) { pendingEl.textContent = 'Nothing waiting for approval.'; }
    else {
      pendingEl.innerHTML = '<table><thead><tr><th>Date</th><th>Category</th><th>Description</th><th>Amount</th><th>Status</th><th>Actions</th></tr></thead><tbody>' +
        rows.map(t => {
          let actions = '';
          if (canApproveManager && !t.manager_approved) actions += '<button class="success" data-q-action="approve/manager" data-q-id="'+t.id+'">Approve (Mgr)</button> ';
          if (canApproveFinance && !t.finance_approved) actions += '<button class="success" data-q-action="approve/finance" data-q-id="'+t.id+'">Approve (Fin)</button> ';
          actions += '<button class="danger" data-q-action="reject" data-q-id="'+t.id+'">Reject</button>';
          return '<tr><td>'+t.transaction_date+'</td><td>'+t.category+'</td><td>'+t.description+'</td><td>'+money(t.amount)+'</td><td>'+(t.manager_approved?'Mgr &check; ':'')+(t.finance_approved?'Fin &check;':'')+'</td><td class="actions">'+actions+'</td></tr>';
        }).join('') + '</tbody></table>';
    }

    const { rows: approvedRows } = await api('/transactions?status=approved&limit=100');
    const approvedEl = document.getElementById('queueApproved');
    if (!approvedRows.length) { approvedEl.textContent = 'Nothing waiting for payment.'; }
    else {
      approvedEl.innerHTML = '<table><thead><tr><th>Date</th><th>Category</th><th>Description</th><th>Amount</th><th>Actions</th></tr></thead><tbody>' +
        approvedRows.map(t => '<tr><td>'+t.transaction_date+'</td><td>'+t.category+'</td><td>'+t.description+'</td><td>'+money(t.amount)+'</td><td class="actions">'+(canApproveFinance?'<button data-q-action="pay" data-q-id="'+t.id+'">Mark Paid</button>':'<span class="muted">Awaiting finance</span>')+'</td></tr>').join('') +
        '</tbody></table>';
    }

    document.querySelectorAll('[data-q-action]').forEach(btn => {
      btn.onclick = async () => {
        try { await api('/transactions/'+btn.dataset.qId+'/'+btn.dataset.qAction, {method:'POST', body: JSON.stringify({})}); loadApprovalQueues(); loadSummary(); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) {
    document.getElementById('queuePending').textContent = 'Could not load.';
    if (document.getElementById('queueApproved')) document.getElementById('queueApproved').textContent = 'Could not load.';
  }
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
    const chartSegments = [
      {value:s.pending, color:'#d98c2b'}, {value:s.approved, color:'#4a9d68'},
      {value:s.paid, color:'#2f7d4f'}, {value:s.rejected, color:'#b03a3a'}
    ];
    const hasAny = chartSegments.some(c=>c.value>0);
    let html = '<div style="display:flex;gap:20px;flex-wrap:wrap;align-items:center;">';
    if (hasAny) html += '<div style="flex-shrink:0;">'+donutChart(chartSegments)+'</div>';
    html += '<div class="row" style="flex:1;min-width:220px;">' + cards.map((c,i) =>
      '<div class="stat-card"><div class="muted">'+c.label+' ('+c.count+')</div><div class="stat-num" id="statNum'+i+'" style="color:'+c.color+';">KES 0.00</div></div>'
    ).join('') + '</div></div>';
    if ((s.documents_paid||0) + (s.documents_pending||0) + (s.documents_unclear||0) > 0) {
      html += '<h2 style="margin-top:18px;">From uploaded documents</h2><div class="row">' +
        '<div class="stat-card"><div class="muted">Paid (from docs)</div><div class="stat-num" style="color:#2f7d4f;">'+money(s.documents_paid||0)+'</div></div>' +
        '<div class="stat-card"><div class="muted">Pending (from docs)</div><div class="stat-num" style="color:#d98c2b;">'+money(s.documents_pending||0)+'</div></div>' +
        (s.documents_unclear ? '<div class="stat-card"><div class="muted">Unclear (from docs)</div><div class="stat-num" style="color:#777;">'+money(s.documents_unclear)+'</div></div>' : '') +
        '</div><p class="muted" style="margin-top:6px;">These totals come from amounts detected in uploaded documents — separate from recorded transactions above.</p>';
    }
    el.innerHTML = html;
    cards.forEach((c,i) => animateCount(document.getElementById('statNum'+i), c.amt));
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
    window.__workers = workers;
    const canEditW = ['manager','admin'].includes(state.user.role);
    el.innerHTML = '<div class="row" style="margin-bottom:8px;"><input id="workerSearchV" placeholder="Search by name or ID number..." /></div><p class="muted">'+marked.length+' of '+workers.length+' present on '+date+'</p><table><thead><tr><th>Name</th><th>ID number</th><th>Designation</th><th>Status</th><th></th></tr></thead><tbody>' +
      workers.map(w => '<tr data-search="'+escHtml((w.name+' '+w.id_number).toLowerCase())+'"><td>'+escHtml(w.name)+'</td><td>'+escHtml(w.id_number)+'</td><td>'+escHtml(w.designation)+'</td><td>'+(markedIds.has(w.id)?'<span class="badge paid">present</span>':'<span class="badge rejected">absent</span>')+'</td><td class="actions"><button class="secondary" data-hist-worker="'+w.id+'">History</button>'+(canEditW?' <button class="secondary" data-edit-workerv="'+w.id+'">Edit</button> <button class="danger" data-del-workerv="'+w.id+'">Remove</button>':'')+'</td></tr>').join('') +
      '</tbody></table>';
    document.getElementById('workerSearchV').oninput = function(){ const q = this.value.toLowerCase().trim(); el.querySelectorAll('tr[data-search]').forEach(r => { r.style.display = r.dataset.search.indexOf(q) >= 0 ? '' : 'none'; }); };
    el.querySelectorAll('[data-hist-worker]').forEach(b => { b.onclick = () => showWorkerHistory(b.dataset.histWorker); });
    el.querySelectorAll('[data-edit-workerv]').forEach(b => { b.onclick = () => editWorker(b.dataset.editWorkerv, loadWorkerAttendanceView); });
    el.querySelectorAll('[data-del-workerv]').forEach(b => { b.onclick = async () => {
      if (!confirm('Remove this worker from the registry? Their attendance records will also be removed.')) return;
      try { await api('/workers/'+b.dataset.delWorkerv, {method:'DELETE'}); loadWorkerAttendanceView(); } catch(e){ alert(e.message); }
    }; });
  } catch(e) { el.textContent = 'Could not load attendance.'; }
}
async function loadWorkerAttendance(){
  const el = document.getElementById('workerAttendance');
  const date = document.getElementById('attDate').value;
  try {
    const [workers, marked] = await Promise.all([ api('/workers'), api('/attendance?date='+date) ]);
    if (!workers.length) { el.textContent = 'No workers enrolled yet — add one above.'; return; }
    const markedIds = new Set(marked.map(m => m.worker_id));
    window.__workers = workers;
    el.innerHTML = '<div class="row" style="margin-bottom:8px;"><input id="workerSearch" placeholder="Search by name or ID number..." /></div><div style="margin-bottom:8px;display:flex;gap:8px;flex-wrap:wrap;"><button class="secondary" id="markAllBtn" type="button">Mark all present</button><button class="secondary" id="sameYesterdayBtn" type="button">Same as yesterday</button><button class="secondary" id="clearMarksBtn" type="button">Clear all</button></div><table><thead><tr><th></th><th>Name</th><th>ID number</th><th>Designation</th><th></th></tr></thead><tbody>' +
      workers.map(w => '<tr data-search="'+escHtml((w.name+' '+w.id_number).toLowerCase())+'"><td><input type="checkbox" data-worker-check="'+w.id+'" '+(markedIds.has(w.id)?'checked':'')+' /></td><td>'+w.name+'</td><td>'+w.id_number+'</td><td>'+w.designation+'</td><td><button class="secondary" style="padding:4px 8px;font-size:12px;" data-edit-worker="'+w.id+'">edit</button> <button class="danger" style="padding:4px 8px;font-size:12px;" data-remove-worker="'+w.id+'">remove</button></td></tr>').join('') +
      '</tbody></table>';
    document.getElementById('workerSearch').oninput = function(){ const q = this.value.toLowerCase().trim(); el.querySelectorAll('tr[data-search]').forEach(r => { r.style.display = r.dataset.search.indexOf(q) >= 0 ? '' : 'none'; }); };
    document.getElementById('markAllBtn').onclick = () => el.querySelectorAll('[data-worker-check]').forEach(c => { if (c.closest('tr').style.display !== 'none') c.checked = true; });
    document.getElementById('clearMarksBtn').onclick = () => el.querySelectorAll('[data-worker-check]').forEach(c => { c.checked = false; });
    document.getElementById('sameYesterdayBtn').onclick = async () => {
      const d = new Date(document.getElementById('attDate').value + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() - 1);
      try {
        const prev = await api('/attendance?date=' + d.toISOString().slice(0,10));
        const ids = new Set(prev.map(m => m.worker_id));
        if (!ids.size) { alert('No attendance was recorded for the day before.'); return; }
        el.querySelectorAll('[data-worker-check]').forEach(c => { c.checked = ids.has(c.dataset.workerCheck); });
      } catch(e){ alert(e.message); }
    };
    el.querySelectorAll('[data-edit-worker]').forEach(btn => { btn.onclick = () => editWorker(btn.dataset.editWorker, loadWorkerAttendance); });
    el.querySelectorAll('[data-remove-worker]').forEach(btn => {
      btn.onclick = async () => {
        if (!confirm('Remove this worker from the registry?')) return;
        try { await api('/workers/'+btn.dataset.removeWorker, {method:'DELETE'}); loadWorkerAttendance(); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) { el.textContent = 'Could not load workers.'; }
}
async function loadDocuments(){
  const el = document.getElementById('docList');
  const canDelete = ['manager','admin'].includes(state.user.role);
  try {
    const docs = await api('/documents');
    if (!docs.length) { el.textContent = 'No documents uploaded yet.'; return; }
    const readCount = docs.filter(d=>d.extraction_status==='done').length;
    const pendingCount = docs.filter(d=>d.extraction_status==='pending').length;
    const failedCount = docs.filter(d=>d.extraction_status==='failed').length;
    let html = '<p class="muted"><strong>'+docs.length+'</strong> document(s) uploaded — '+readCount+' read'+(pendingCount?', '+pendingCount+' analyzing':'')+(failedCount?', '+failedCount+' unreadable':'')+'.</p>';
    html += docs.map(d => {
      const statusBadge = d.extraction_status==='done' ? '<span class="badge paid">read</span>' : d.extraction_status==='failed' ? '<span class="badge rejected">unreadable</span>' : d.extraction_status==='skipped' ? '<span class="muted">not analyzed</span>' : '<span class="badge pending">analyzing</span>';
      const uploaderLine = d.uploader ? d.uploader.full_name+' ('+d.uploader.role+')' : 'Unknown';
      return '<div class="card" style="margin-bottom:10px;padding:14px;">' +
        '<div class="topbar"><strong>'+d.file_name+'</strong>'+statusBadge+'</div>' +
        (d.description?'<div class="muted">'+d.description+'</div>':'') +
        '<div class="muted" style="margin:6px 0;">Uploaded by '+uploaderLine+' on '+new Date(d.created_at).toLocaleDateString()+'</div>' +
        (d.extracted_text?'<div style="background:#f5f6f8;border-radius:6px;padding:8px;font-size:13px;margin:6px 0;">'+d.extracted_text+(d.extracted_amount?'<div style="margin-top:6px;font-weight:600;">Amount: '+money(d.extracted_amount)+' — '+(d.extracted_payment_status||'unclear')+'</div>':'')+'</div>':'') +
        '<div class="actions">'+(d.download_url?'<a href="'+d.download_url+'" target="_blank"><button class="secondary">Download</button></a>':'')+(canDelete?' <button class="danger" data-del-doc="'+d.id+'">Delete</button>':'')+'</div>' +
      '</div>';
    }).join('');
    el.innerHTML = html;
    el.querySelectorAll('[data-del-doc]').forEach(btn => {
      btn.onclick = async () => {
        if (!confirm('Delete this document?')) return;
        try { await api('/documents/'+btn.dataset.delDoc, {method:'DELETE'}); loadDocuments(); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) { el.textContent = 'Could not load documents.'; }
}
async function loadStoreItems(){
  const el = document.getElementById('storeList');
  const canManage = ['storekeeper','admin'].includes(state.user.role);
  const canReceive = ['storekeeper','admin'].includes(state.user.role);
  const canDelete = ['storekeeper','admin'].includes(state.user.role);
  try {
    const items = await api('/store/items');
    if (!items.length) { el.textContent = 'No materials tracked yet.'; return; }
    const lowItems = items.filter(i => i.reorder_level && Number(i.quantity_in_stock) <= Number(i.reorder_level));
    const lowBanner = lowItems.length ? '<div style="background:#fbeaea;color:#b03a3a;padding:10px 12px;border-radius:8px;margin-bottom:10px;font-weight:600;">Low stock: ' + lowItems.map(i => escHtml(i.name) + ' (' + i.quantity_in_stock + ' ' + escHtml(i.unit) + ' left)').join(', ') + '</div>' : '';
    el.innerHTML = lowBanner + '<table><thead><tr><th>Material</th><th>Stock In</th><th>Stock Out</th><th>Remaining</th><th>Reorder at</th><th>Actions</th></tr></thead><tbody>' +
      items.map(i => {
        const low = i.reorder_level && Number(i.quantity_in_stock) <= Number(i.reorder_level);
        return '<tr><td>'+i.name+'</td><td>'+i.total_in+' '+i.unit+'</td><td>'+i.total_out+' '+i.unit+'</td><td>'+(low?'<span class="badge rejected">':'<strong>')+i.quantity_in_stock+' '+i.unit+(low?'</span>':'</strong>')+'</td><td>'+(i.reorder_level||'—')+'</td>' +
          '<td class="actions">'+(canManage?'<button class="success" data-move-in="'+i.id+'">Stock in</button> ':'')+(low&&canReceive?'<button class="secondary" data-reorder="'+i.id+'">Request reorder</button> ':'')+(canManage?'<button class="danger" data-move-out="'+i.id+'">Stock out</button> <button class="secondary" data-count-item="'+i.id+'">Count</button>':'')+(canDelete?' <button class="secondary" data-del-item="'+i.id+'">Remove</button>':'')+'</td></tr>';
      }).join('') +
      '</tbody></table>';

    el.querySelectorAll('[data-move-in]').forEach(btn => {
      btn.onclick = async () => {
        const id = btn.dataset.moveIn;
        const quantity = prompt('Quantity to add to stock (e.g. a count correction, or a delivery with no cost to track):');
        if (!quantity) return;
        const reason = prompt('Reason / note (optional):') || '';
        try { await api('/store/items/'+id+'/movement', {method:'POST', body: JSON.stringify({movement_type:'in', quantity, reason})}); loadStoreItems(); }
        catch(e){ alert(e.message); }
      };
    });
    el.querySelectorAll('[data-reorder]').forEach(btn => {
      btn.onclick = async () => {
        const id = btn.dataset.reorder;
        const quantity = prompt('How much do you want to reorder?');
        if (!quantity) return;
        const estimated_amount = prompt('Estimated cost (KES) — this goes to manager and finance for approval:');
        if (!estimated_amount) return;
        const note = prompt('Note (optional):') || '';
        try {
          await api('/store/items/'+id+'/reorder-request', {method:'POST', body: JSON.stringify({quantity, estimated_amount, note})});
          alert('Reorder request submitted — manager and finance will review it.');
        } catch(e){ alert(e.message); }
      };
    });
    el.querySelectorAll('[data-move-out]').forEach(btn => {
      btn.onclick = async () => {
        const id = btn.dataset.moveOut;
        const quantity = prompt('Quantity to record OUT (used/issued):');
        if (!quantity) return;
        const issued_to = prompt('Issued to (worker or team name):'); if (!issued_to) return;
        const used_for = prompt('Used for (task or area, e.g. Block A foundation):') || '';
        try { await api('/store/items/'+id+'/movement', {method:'POST', body: JSON.stringify({movement_type:'out', quantity, issued_to, used_for})}); loadStoreItems(); }
        catch(e){ alert(e.message); }
      };
    });
    el.querySelectorAll('[data-count-item]').forEach(btn => {
      btn.onclick = async () => {
        const counted = prompt('Physical count: how many are actually in the store right now?'); if (counted === null || counted === '') return;
        const reason = prompt('Note (optional, e.g. damaged, miscounted earlier):') || '';
        try {
          const r = await api('/store/items/'+btn.dataset.countItem+'/count', {method:'POST', body: JSON.stringify({counted, reason})});
          alert(r.difference === 0 ? 'The count matches the system. No change needed.' : 'Stock corrected by ' + (r.difference > 0 ? '+' : '') + r.difference + '. The difference was logged.');
          loadStoreItems();
        } catch(e){ alert(e.message); }
      };
    });
    el.querySelectorAll('[data-del-item]').forEach(btn => {
      btn.onclick = async () => {
        if (!confirm('Remove this material from tracking? Its movement history will also be deleted.')) return;
        try { await api('/store/items/'+btn.dataset.delItem, {method:'DELETE'}); loadStoreItems(); }
        catch(e){ alert(e.message); }
      };
    });
  } catch(e) { el.textContent = 'Could not load store items.'; }
}
async function loadUsers(){
  const el = document.getElementById('userList');
  try {
    const users = await api('/users');
    if (!users.length) { el.textContent = 'No team logins yet.'; return; }
    el.innerHTML = '<table><thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead><tbody>' +
      users.map(u => '<tr><td>'+u.full_name+'</td><td>'+u.email+'</td><td style="text-transform:capitalize;">'+u.role+'</td><td>'+(u.is_active===false?'<span class="badge rejected">inactive</span>':'<span class="badge paid">active</span>')+'</td><td class="actions">'+(u.id!==state.user.id?'<button class="secondary" data-reset-user="'+u.id+'">Reset password</button> <button class="'+(u.is_active===false?'success':'danger')+'" data-toggle-user="'+u.id+'" data-next="'+(u.is_active===false?'true':'false')+'">'+(u.is_active===false?'Activate':'Deactivate')+'</button>':'<span class="muted">you</span>')+'</td></tr>').join('') +
      '</tbody></table>';
    el.querySelectorAll('[data-reset-user]').forEach(btn => {
      btn.onclick = async () => {
        if (!confirm('Reset this person\\'s password? A new temporary password will be shown once.')) return;
        try { const r = await api('/users/'+btn.dataset.resetUser+'/reset-password', {method:'POST'}); alert('New temporary password for '+r.email+':\\n\\n'+r.temporary_password+'\\n\\nGive it to them; they can change it in Settings.'); }
        catch(e){ alert(e.message); }
      };
    });
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

// ---------- Installable app (PWA) ----------
const zlib = require('zlib');
const CRC_TABLE = (() => { const t = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(buf) { let c = 0xffffffff; for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function pngChunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
// Draws a green rounded tile with a white tipper-truck mark, as a PNG (no image libraries needed).
function makeIconPng(size, maskable) {
  const px = Buffer.alloc(size * size * 4);
  const inRect = (x, y, x0, y0, x1, y1) => x >= x0 && x <= x1 && y >= y0 && y <= y1;
  const inCircle = (x, y, cx, cy, r) => (x - cx) * (x - cx) + (y - cy) * (y - cy) <= r * r;
  const scale = maskable ? 0.72 : 1; // keep the mark inside the safe zone for maskable icons
  for (let j = 0; j < size; j++) for (let i = 0; i < size; i++) {
    const u = i / size, v = j / size;
    // background (rounded square, or full-bleed when maskable)
    let bg = true;
    if (!maskable) {
      const r = 0.18, cx = Math.min(Math.max(u, r), 1 - r), cy = Math.min(Math.max(v, r), 1 - r);
      bg = (u - cx) * (u - cx) + (v - cy) * (v - cy) <= r * r;
    }
    let rgba = [0, 0, 0, 0];
    if (bg) {
      rgba = [23, 59, 41, 255];
      const x = (u - 0.5) / scale + 0.5, y = (v - 0.5) / scale + 0.5;
      const body = inRect(x, y, 0.16, 0.36, 0.58, 0.64);
      const cab = inRect(x, y, 0.60, 0.46, 0.84, 0.64) || (inRect(x, y, 0.60, 0.46, 0.70, 0.64));
      const win = inRect(x, y, 0.68, 0.50, 0.78, 0.56);
      const wheel = inCircle(x, y, 0.30, 0.68, 0.075) || inCircle(x, y, 0.70, 0.68, 0.075);
      const hub = inCircle(x, y, 0.30, 0.68, 0.03) || inCircle(x, y, 0.70, 0.68, 0.03);
      if (wheel) rgba = [255, 255, 255, 255];
      if (hub) rgba = [47, 158, 92, 255];
      if (body || cab) rgba = [255, 255, 255, 255];
      if (win) rgba = [47, 158, 92, 255];
    }
    const o = (j * size + i) * 4; px[o] = rgba[0]; px[o + 1] = rgba[1]; px[o + 2] = rgba[2]; px[o + 3] = rgba[3];
  }
  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let j = 0; j < size; j++) { raw[j * (size * 4 + 1)] = 0; px.copy(raw, j * (size * 4 + 1) + 1, j * size * 4, (j + 1) * size * 4); }
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4); ihdr[8] = 8; ihdr[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), pngChunk('IHDR', ihdr), pngChunk('IDAT', zlib.deflateSync(raw)), pngChunk('IEND', Buffer.alloc(0))]);
}
const ICON_CACHE = {};
function sendIcon(res, size, maskable) {
  const key = size + (maskable ? 'm' : '');
  if (!ICON_CACHE[key]) ICON_CACHE[key] = makeIconPng(size, maskable);
  res.setHeader('Cache-Control', 'public, max-age=86400');
  res.type('png').send(ICON_CACHE[key]);
}
app.get('/icon-192.png', (req, res) => sendIcon(res, 192, false));
app.get('/icon-512.png', (req, res) => sendIcon(res, 512, false));
app.get('/icon-maskable-512.png', (req, res) => sendIcon(res, 512, true));
app.get('/apple-touch-icon.png', (req, res) => sendIcon(res, 180, true));
app.get('/favicon.ico', (req, res) => sendIcon(res, 192, false));

app.get('/manifest.webmanifest', (req, res) => {
  res.setHeader('Cache-Control', 'public, max-age=3600');
  res.type('application/manifest+json').send(JSON.stringify({
    name: 'Site Transactions',
    short_name: 'SiteTx',
    description: 'Construction site expenses, store, workers and wages.',
    id: '/',
    start_url: '/',
    scope: '/',
    display: 'standalone',
    orientation: 'portrait',
    background_color: '#173B29',
    theme_color: '#173B29',
    icons: [
      { src: '/icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any' },
      { src: '/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any' },
      { src: '/icon-maskable-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' }
    ]
  }));
});

// Service worker: keeps the app shell available with no network; API calls and PDFs always go to the server.
const SW_JS = `
const CACHE = 'sitetx-shell-v1';
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.add('/')).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin || url.pathname.startsWith('/api/')) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => { const copy = r.clone(); if (r.ok && url.pathname === '/') caches.open(CACHE).then(c => c.put('/', copy)); return r; }).catch(() => caches.match('/').then(r => r || new Response('You are offline. Reconnect and reopen the app.', { status: 503, headers: { 'Content-Type': 'text/plain' } }))));
    return;
  }
  if (/^\\/(icon-|apple-touch|manifest)/.test(url.pathname)) {
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res; })));
  }
});
`;
app.get('/sw.js', (req, res) => {
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Service-Worker-Allowed', '/');
  res.type('application/javascript').send(SW_JS);
});

// Basic security headers
app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('Referrer-Policy', 'same-origin');
  next();
});

app.get('/', (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0');
  res.type('html').send(INDEX_HTML);
});

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
  res.json({ token, user: { id: user.id, full_name: user.full_name, email: user.email, role: user.role, site_id: user.site_id, organization_id: user.organization_id } });
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
// =====================================================================
// Role improvements: tenant-isolation guard, audit trail, clerk resubmit,
// worker edit/history, attendance corrections, stock count, site activity PDF.
// =====================================================================

// Every /:id route on these resources first checks that the record belongs to the caller's site/company.
function scopeGuard(table) {
  return (req, res, next) => requireAuth(req, res, async () => {
    try {
      const id = req.params.id;
      if (!/^[0-9a-f-]{36}$/i.test(id)) return next();
      const { data: row } = await supabase.from(table).select('site_id').eq('id', id).maybeSingle();
      if (!row) return next();
      if (req.user.role === 'superadmin') return next();
      if (req.user.role === 'admin') {
        const ids = await getOrgSiteIds(req.user.organization_id);
        return ids.includes(row.site_id) ? next() : res.status(404).json({ error: 'Not found' });
      }
      if (row.site_id && row.site_id === req.user.site_id) return next();
      return res.status(404).json({ error: 'Not found' });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  });
}
app.use('/api/transactions/:id', scopeGuard('cst_transactions'));
app.use('/api/workers/:id', scopeGuard('cst_workers'));
app.use('/api/store/items/:id', scopeGuard('cst_store_items'));
app.use('/api/suppliers/:id', scopeGuard('cst_vendors'));

// Best-effort audit trail: never blocks or fails the real action.
function audit(req, action, entity, entityId, detail) {
  try {
    Promise.resolve(supabase.from('cst_audit_log').insert({
      organization_id: req.user.organization_id || null, site_id: req.user.site_id || null, actor_id: req.user.id,
      actor_name: req.user.full_name || null, actor_role: req.user.role, action, entity: entity || null,
      entity_id: entityId ? String(entityId) : null, detail: detail ? String(detail).slice(0, 500) : null
    })).catch(() => {});
  } catch {}
}

async function lowStockAlert(item, newQty) {
  if (item.reorder_level && newQty <= Number(item.reorder_level)) {
    const { data: recipients } = await supabase.from('cst_users').select('id').eq('site_id', item.site_id).in('role', ['manager', 'storekeeper']);
    for (const r of (recipients || [])) await notify(r.id, null, 'batch_submitted', 'Low stock alert: "' + item.name + '" is down to ' + newQty + ' ' + item.unit + ' (reorder level: ' + item.reorder_level + ').');
  }
}

// ----- Clerk: fix and resubmit a rejected transaction -----
app.post('/api/transactions/:id/resubmit', requireAuth, requireRole('clerk'), async (req, res) => {
  const { data: txn } = await supabase.from('cst_transactions').select('*').eq('id', req.params.id).maybeSingle();
  if (!txn || txn.created_by !== req.user.id) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status !== 'rejected') return res.status(400).json({ error: 'Only rejected transactions can be resubmitted' });
  const description = String(req.body.description || txn.description).trim();
  const amount = Number(req.body.amount || txn.amount);
  if (!description || !(amount > 0)) return res.status(400).json({ error: 'Enter a description and an amount above 0' });
  const { data, error } = await supabase.from('cst_transactions').update({
    description, amount, status: 'pending', manager_approved: false, manager_approved_by: null, manager_approved_at: null,
    finance_approved: false, finance_approved_by: null, finance_approved_at: null,
    rejected_by: null, rejected_at: null, rejection_reason: null, updated_at: new Date().toISOString()
  }).eq('id', txn.id).select().single();
  if (error) return res.status(500).json({ error: error.message });
  const { data: recipients } = await supabase.from('cst_users').select('id').eq('site_id', txn.site_id).in('role', ['manager', 'finance']);
  for (const r of (recipients || [])) await notify(r.id, txn.id, 'batch_submitted', 'A clerk corrected and resubmitted a rejected transaction: "' + description + '" (KES ' + amount.toLocaleString() + ').');
  audit(req, 'resubmit_transaction', 'transaction', txn.id, description + ' KES ' + amount);
  res.json(data);
});

// ----- Workers: edit + history -----
app.patch('/api/workers/:id', requireAuth, requireRole('clerk', 'manager', 'admin'), async (req, res) => {
  const updates = {};
  for (const k of ['name', 'id_number', 'phone_number', 'designation']) if (req.body[k] !== undefined) updates[k] = String(req.body[k]).trim() || null;
  if (req.body.daily_rate !== undefined) {
    const r = Number(req.body.daily_rate);
    if (!(r >= 0) || !isFinite(r)) return res.status(400).json({ error: 'Enter a valid daily rate' });
    updates.daily_rate = r;
  }
  if ((updates.name === null) || (updates.designation === null) || (updates.id_number === null)) return res.status(400).json({ error: 'Name, ID number and designation cannot be empty' });
  if (!Object.keys(updates).length) return res.status(400).json({ error: 'Nothing to update' });
  const { data, error } = await supabase.from('cst_workers').update(updates).eq('id', req.params.id).select().single();
  if (error) {
    if (error.message.includes('duplicate')) return res.status(409).json({ error: 'Another worker already has that ID number' });
    return res.status(500).json({ error: error.message });
  }
  audit(req, 'edit_worker', 'worker', req.params.id, JSON.stringify(updates));
  res.json(data);
});

app.get('/api/workers/:id/history', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  try {
    const { data: worker } = await supabase.from('cst_workers').select('*').eq('id', req.params.id).maybeSingle();
    if (!worker) return res.status(404).json({ error: 'Worker not found' });
    const { data: att } = await supabase.from('cst_attendance').select('attendance_date').eq('worker_id', worker.id).order('attendance_date', { ascending: false }).limit(366);
    const { data: wages } = await supabase.from('cst_wage_payments').select('week_start, week_end, days_present, daily_rate, total_amount, paid_at, disbursement_status, mpesa_receipt').eq('worker_id', worker.id).order('week_start', { ascending: false }).limit(52);
    const paid = (wages || []).filter(w => w.disbursement_status !== 'failed').reduce((s, w) => s + (Number(w.total_amount) || 0), 0);
    res.json({ worker, attendance: (att || []).map(a => a.attendance_date), days_total: (att || []).length, wages: wages || [], paid_total: paid });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ----- Stock count (physical count vs system) -----
app.post('/api/store/items/:id/count', requireAuth, requireRole('storekeeper', 'admin'), async (req, res) => {
  const counted = Number(req.body.counted);
  const note = String(req.body.reason || '').trim();
  if (req.body.counted === '' || req.body.counted == null || !(counted >= 0) || !isFinite(counted)) return res.status(400).json({ error: 'Enter the counted quantity (0 or more)' });
  const { data: item } = await supabase.from('cst_store_items').select('*').eq('id', req.params.id).maybeSingle();
  if (!item) return res.status(404).json({ error: 'Item not found' });
  const before = Number(item.quantity_in_stock), diff = counted - before;
  if (diff === 0) return res.json({ item, difference: 0, message: 'The count matches the system. No change needed.' });
  const { error: updErr } = await supabase.from('cst_store_items').update({ quantity_in_stock: counted, updated_at: new Date().toISOString() }).eq('id', item.id);
  if (updErr) return res.status(500).json({ error: updErr.message });
  await supabase.from('cst_store_movements').insert({
    item_id: item.id, site_id: item.site_id, movement_type: diff > 0 ? 'in' : 'out', quantity: Math.abs(diff),
    reason: 'Stock count adjustment: system ' + before + ', counted ' + counted + (note ? ' - ' + note : ''), recorded_by: req.user.id
  });
  audit(req, 'stock_count', 'store_item', item.id, item.name + ': system ' + before + ' -> counted ' + counted + (note ? ' (' + note + ')' : ''));
  await lowStockAlert(item, counted);
  res.json({ item: { ...item, quantity_in_stock: counted }, difference: diff });
});

// ----- Audit trail -----
app.get('/api/audit', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  let q = supabase.from('cst_audit_log').select('*').order('created_at', { ascending: false }).limit(300);
  if (req.user.role === 'admin') q = q.eq('organization_id', req.user.organization_id);
  else q = q.eq('site_id', req.user.site_id);
  if (req.query.from) q = q.gte('created_at', req.query.from + 'T00:00:00Z');
  if (req.query.to) q = q.lte('created_at', req.query.to + 'T23:59:59Z');
  const text = String(req.query.q || '').replace(/[^a-zA-Z0-9 _-]/g, '').trim();
  if (text) q = q.or('action.ilike.%' + text + '%,detail.ilike.%' + text + '%,actor_name.ilike.%' + text + '%');
  const { data, error } = await q;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data || []);
});

// ----- Site activity report (PDF): attendance, expenses, deliveries, stock issued -----
function reportRange(period, dateStr) {
  const endDate = /^\d{4}-\d{2}-\d{2}$/.test(dateStr || '') ? dateStr : new Date().toISOString().slice(0, 10);
  const s = new Date(endDate + 'T00:00:00Z');
  if (period === 'weekly') s.setUTCDate(s.getUTCDate() - 6);
  return { from: s.toISOString().slice(0, 10), to: endDate };
}
app.get('/api/reports/site-activity.pdf', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  try {
    const period = req.query.period === 'daily' ? 'daily' : 'weekly';
    const { from, to } = reportRange(period, req.query.date);
    const n2 = n => Number(n).toLocaleString('en-KE', { minimumFractionDigits: 2 });

    let wq = supabase.from('cst_workers').select('id, name, designation'); wq = await scopeToSites(wq, req.user);
    const { data: workers } = await wq;
    let aq = supabase.from('cst_attendance').select('worker_id, attendance_date').gte('attendance_date', from).lte('attendance_date', to); aq = await scopeToSites(aq, req.user);
    const { data: att } = await aq;
    let tq = supabase.from('cst_transactions').select('category, description, amount, status, transaction_date').gte('transaction_date', from).lte('transaction_date', to).order('transaction_date'); tq = await scopeToSites(tq, req.user);
    const { data: txns } = await tq;
    let mq = supabase.from('cst_store_movements').select('created_at, movement_type, quantity, reason, issued_to, used_for, item:item_id(name, unit), supplier:supplier_id(name)').gte('created_at', from + 'T00:00:00Z').lte('created_at', to + 'T23:59:59Z').order('created_at'); mq = await scopeToSites(mq, req.user);
    const { data: moves } = await mq;

    const nameOf = {}; for (const w of (workers || [])) nameOf[w.id] = w;
    const perDay = {}, perWorker = {};
    for (const a of (att || [])) { perDay[a.attendance_date] = (perDay[a.attendance_date] || 0) + 1; perWorker[a.worker_id] = (perWorker[a.worker_id] || 0) + 1; }

    const branding = await fetchOrgBranding(req);
    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="site-activity-' + period + '-' + to + '.pdf"');
    doc.pipe(res);
    drawFormHeader(doc, (period === 'daily' ? 'Daily' : 'Weekly') + ' Site Activity Report', 'Attendance, expenses, deliveries and stock used', branding);
    let y = drawInfoBox(doc, 130, ['Period:  ' + from + '  to  ' + to + '          Workers enrolled: ' + (workers || []).length, 'Generated:  ' + new Date().toLocaleString('en-KE')]) + 16;
    const heading = t => { if (y > 730) { doc.addPage(); y = 50; } doc.font('Helvetica-Bold').fontSize(11).fillColor(PDF_FOREST).text(t, PDF_LEFT, y); y += 16; };

    heading('Attendance');
    const aRows = Object.keys(perDay).sort().map(d => ({ cells: [d, perDay[d] + ' present', Math.max(0, (workers || []).length - perDay[d]) + ' absent'] }));
    if (!aRows.length) aRows.push({ cells: ['No attendance recorded', '', ''] });
    y = pdfTable(doc, y, [{ label: 'Date', width: 150 }, { label: 'Present', width: 180 }, { label: 'Absent', width: 185 }], aRows) + 8;
    const wRows = (workers || []).filter(w => perWorker[w.id]).sort((a, b) => perWorker[b.id] - perWorker[a.id]).map(w => ({ cells: [w.name, w.designation || '', perWorker[w.id] + ' day(s)'] }));
    if (wRows.length) { y = pdfTable(doc, y, [{ label: 'Worker', width: 220 }, { label: 'Designation', width: 160 }, { label: 'Days present', width: 135 }], wRows) + 18; } else y += 10;

    heading('Expenses');
    const exp = (txns || []).filter(t => t.status !== 'rejected');
    const eRows = exp.map(t => ({ cells: [t.transaction_date, t.category, t.description, t.status.toUpperCase(), n2(t.amount)] }));
    eRows.push({ bold: true, cells: ['', '', 'TOTAL (excluding rejected)', '', n2(exp.reduce((s, t) => s + (Number(t.amount) || 0), 0))] });
    y = pdfTable(doc, y, [{ label: 'Date', width: 62 }, { label: 'Category', width: 62 }, { label: 'Description', width: 218 }, { label: 'Status', width: 55 }, { label: 'Amount (KES)', width: 118, align: 'right' }], eRows) + 18;

    heading('Deliveries received');
    const dl = (moves || []).filter(m => m.movement_type === 'in' && m.supplier);
    const dRows = dl.map(m => ({ cells: [String(m.created_at).slice(0, 10), m.supplier.name, m.item ? m.item.name : '', Number(m.quantity) + ' ' + (m.item ? m.item.unit : '')] }));
    if (!dRows.length) dRows.push({ cells: ['No deliveries in this period', '', '', ''] });
    y = pdfTable(doc, y, [{ label: 'Date', width: 70 }, { label: 'Supplier', width: 170 }, { label: 'Material', width: 170 }, { label: 'Quantity', width: 105 }], dRows) + 18;

    heading('Stock issued / used');
    const out = (moves || []).filter(m => m.movement_type === 'out');
    const oRows = out.map(m => ({ cells: [String(m.created_at).slice(0, 10), m.item ? m.item.name : '', Number(m.quantity) + ' ' + (m.item ? m.item.unit : ''), m.issued_to || '', m.used_for || m.reason || ''] }));
    if (!oRows.length) oRows.push({ cells: ['No stock issued in this period', '', '', '', ''] });
    y = pdfTable(doc, y, [{ label: 'Date', width: 62 }, { label: 'Material', width: 140 }, { label: 'Quantity', width: 80 }, { label: 'Issued to', width: 105 }, { label: 'Used for / note', width: 128 }], oRows);
    drawFormFooter(doc, Math.min(y + 14, 800), 'Generated automatically from attendance, transaction and store records.');
    doc.end();
  } catch (e) { if (!res.headersSent) res.status(500).send('Could not generate report: ' + e.message); else res.end(); }
});

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
  if (req.user.role === 'clerk') query = query.eq('created_by', req.user.id);
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
  query = await scopeToSites(query, req.user);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });

  const summary = { pending: 0, approved: 0, paid: 0, rejected: 0, pending_count: 0, approved_count: 0, paid_count: 0, rejected_count: 0, by_category: {} };
  for (const t of data) {
    const amt = Number(t.amount) || 0;
    if (summary[t.status] !== undefined) { summary[t.status] += amt; summary[t.status + '_count']++; }
    summary.by_category[t.category] = (summary.by_category[t.category] || 0) + amt;
  }

  let docQuery = supabase.from('cst_documents').select('extracted_amount, extracted_payment_status').eq('extraction_status', 'done').not('extracted_amount', 'is', null);
  docQuery = await scopeToSites(docQuery, req.user);
  const { data: docs } = await docQuery;
  summary.documents_paid = 0;
  summary.documents_pending = 0;
  summary.documents_unclear = 0;
  for (const d of (docs || [])) {
    const amt = Number(d.extracted_amount) || 0;
    if (d.extracted_payment_status === 'paid') summary.documents_paid += amt;
    else if (d.extracted_payment_status === 'pending') summary.documents_pending += amt;
    else summary.documents_unclear += amt;
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
  audit(req, 'approve_manager', 'transaction', id, null);

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
  audit(req, 'approve_finance', 'transaction', id, null);

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
  audit(req, 'reject_transaction', 'transaction', id, null);
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
  audit(req, 'mark_paid', 'transaction', id, null);

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
  audit(req, 'edit_transaction', 'transaction', id, JSON.stringify(updates));
  res.json(data);
});

app.delete('/api/transactions/:id', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { id } = req.params;
  const { data: before } = await supabase.from('cst_transactions').select('description, amount, status').eq('id', id).maybeSingle();
  const { error } = await supabase.from('cst_transactions').delete().eq('id', id);
  if (error) return res.status(500).json({ error: error.message });
  audit(req, 'delete_transaction', 'transaction', id, before ? before.description + ' KES ' + before.amount + ' (' + before.status + ')' : null);
  res.json({ ok: true });
});

// Bulk-clear all transaction data for a site — destructive, requires explicit confirm=yes
app.delete('/api/transactions', requireAuth, requireRole('manager', 'admin', 'superadmin'), async (req, res) => {
  if (req.query.confirm !== 'yes') return res.status(400).json({ error: 'Add ?confirm=yes to confirm this destructive action' });

  let query = supabase.from('cst_transactions').delete();
  if (req.query.site_id) {
    // Explicit site requested — only allow it if it's actually within the requester's reach.
    if (req.user.role !== 'superadmin') {
      const allowedSites = req.user.role === 'admin' ? await getOrgSiteIds(req.user.organization_id) : [req.user.site_id];
      if (!allowedSites.includes(req.query.site_id)) return res.status(403).json({ error: 'That site is not in your organization' });
    }
    query = query.eq('site_id', req.query.site_id);
  } else if (req.user.role === 'superadmin') {
    return res.status(400).json({ error: 'Superadmin must specify ?site_id= explicitly — refusing to wipe every organization at once' });
  } else {
    query = await scopeToSites(query, req.user);
    if (req.user.role !== 'admin' && !req.user.site_id) return res.status(400).json({ error: 'No site associated with your account' });
  }

  const { error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  audit(req, 'clear_all_transactions', 'transaction', null, 'All transaction data was cleared');
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

  const allowedRoles = req.user.role === 'admin' ? ['clerk', 'manager', 'finance', 'admin', 'storekeeper'] : ['clerk', 'finance', 'storekeeper'];
  if (!allowedRoles.includes(role)) return res.status(403).json({ error: `You can only create: ${allowedRoles.join(', ')}` });

  const site_id = req.body.site_id || req.user.site_id;
  const plainPassword = randomPassword();
  const password_hash = bcrypt.hashSync(plainPassword, 10);

  const { data, error } = await supabase.from('cst_users').insert({
    full_name, email: email.toLowerCase().trim(), password_hash, role, site_id, organization_id: req.user.organization_id
  }).select('id, full_name, email, role, site_id').single();

  if (error) {
    if (error.message.includes('duplicate')) return res.status(409).json({ error: 'A user with this email already exists' });
    return res.status(500).json({ error: error.message });
  }
  // Return the generated password once — the manager must relay it to the new user.
  audit(req, 'create_user', 'user', data.id, data.full_name + ' (' + data.role + ')');
  res.status(201).json({ ...data, temporary_password: plainPassword });
});

app.get('/api/users', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  let query = supabase.from('cst_users').select('id, full_name, email, role, site_id, is_active, created_at').order('created_at', { ascending: false });
  query = await scopeToSites(query, req.user);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

// Loads a user the caller is allowed to manage (same site/organisation); null if not allowed.
async function getManageableUser(req, id) {
  const { data: target } = await supabase.from('cst_users').select('id, full_name, email, role, site_id, organization_id, is_active').eq('id', id).single();
  if (!target) return null;
  if (req.user.role === 'admin') {
    if ((target.organization_id || null) !== (req.user.organization_id || null)) return null;
  } else {
    if (target.site_id !== req.user.site_id || ['manager', 'admin', 'superadmin'].includes(target.role)) return null;
  }
  return target;
}

app.patch('/api/users/:id/status', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { is_active } = req.body;
  if (typeof is_active !== 'boolean') return res.status(400).json({ error: 'is_active (boolean) is required' });
  if (req.params.id === req.user.id) return res.status(400).json({ error: "You can't deactivate your own account" });
  const target = await getManageableUser(req, req.params.id);
  if (!target) return res.status(404).json({ error: 'User not found' });

  const { data, error } = await supabase.from('cst_users').update({ is_active }).eq('id', req.params.id).select('id, full_name, email, role, is_active').single();
  if (error) return res.status(500).json({ error: error.message });
  audit(req, is_active ? 'activate_user' : 'deactivate_user', 'user', req.params.id, target.full_name + ' (' + target.role + ')');
  res.json(data);
});

// Manager/admin resets a team member's password and unlocks the account; the new temporary password is shown once.
app.post('/api/users/:id/reset-password', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  if (req.params.id === req.user.id) return res.status(400).json({ error: 'Use Settings > Change password for your own account' });
  const target = await getManageableUser(req, req.params.id);
  if (!target) return res.status(404).json({ error: 'User not found' });
  const plainPassword = randomPassword();
  const { error } = await supabase.from('cst_users').update({ password_hash: bcrypt.hashSync(plainPassword, 10), failed_attempts: 0, locked_until: null }).eq('id', target.id);
  if (error) return res.status(500).json({ error: error.message });
  audit(req, 'reset_password', 'user', target.id, target.full_name + ' (' + target.role + ')');
  res.json({ id: target.id, email: target.email, full_name: target.full_name, temporary_password: plainPassword });
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

  let allowed = req.user.id === txn.created_by || req.user.id === txn.manager_approved_by || req.user.id === txn.finance_approved_by || ['manager', 'finance'].includes(req.user.role);
  if (!allowed && req.user.role === 'admin') {
    const orgSites = await getOrgSiteIds(req.user.organization_id);
    allowed = orgSites.includes(txn.site_id);
  }
  if (!allowed && req.user.role === 'superadmin') allowed = true;
  if (!allowed) return res.status(403).send('Not permitted to view this receipt');
  if (txn.status !== 'paid') return res.status(400).send('Receipt is only available once a transaction is paid');

  let orgName = 'Site Transactions', orgLogoUrl = null;
  if (req.user.organization_id) {
    const { data: org } = await supabase.from('cst_organizations').select('name, logo_url').eq('id', req.user.organization_id).single();
    if (org) { orgName = org.name; orgLogoUrl = org.logo_url; }
  }

  res.type('html').send(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Receipt ${id}</title>
  <style>body{font-family:Arial,sans-serif;max-width:480px;margin:40px auto;color:#1f2d3d;}
  .box{border:1px solid #dde1e6;border-radius:10px;padding:24px;}
  .brand{display:flex;align-items:center;gap:10px;margin-bottom:14px;}
  .brand img{width:40px;height:40px;border-radius:8px;object-fit:cover;}
  .brand-name{font-weight:700;font-size:15px;}
  h1{font-size:18px;margin:0 0 4px;} .muted{color:#6b7684;font-size:12px;}
  table{width:100%;border-collapse:collapse;margin-top:16px;}
  td{padding:8px 0;border-bottom:1px solid #eee;font-size:14px;}
  td:first-child{color:#6b7684;} .amt{font-size:22px;font-weight:700;margin-top:16px;}
  .stamp{margin-top:20px;padding:8px 12px;background:#dfeee5;color:#2f7d4f;display:inline-block;border-radius:6px;font-weight:700;font-size:13px;}
  @media print{button{display:none;}}</style></head>
  <body><div class="box">
  <div class="brand">${orgLogoUrl ? `<img src="${orgLogoUrl}" alt="logo" />` : ''}<span class="brand-name">${orgName}</span></div>
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
  query = await scopeToSites(query, req.user);
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
  const { data, error } = await supabase.from('cst_transaction_vendors').select('id, supplied_item, vendor:vendor_id(id, name, role_type, contact_phone, contact_email)').eq('transaction_id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.post('/api/transactions/:id/vendors', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { vendor_id, supplied_item } = req.body;
  if (!vendor_id) return res.status(400).json({ error: 'vendor_id is required' });
  const { data, error } = await supabase.from('cst_transaction_vendors').insert({ transaction_id: req.params.id, vendor_id, supplied_item: supplied_item || null }).select().single();
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
app.post('/api/workers', requireAuth, requireRole('clerk', 'manager', 'admin'), async (req, res) => {
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
  audit(req, 'enroll_worker', 'worker', data.id, data.name + ' (' + data.id_number + ')');
  res.status(201).json(data);
});

app.get('/api/workers', requireAuth, async (req, res) => {
  let query = supabase.from('cst_workers').select('*').order('name');
  query = await scopeToSites(query, req.user);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.delete('/api/workers/:id', requireAuth, requireRole('clerk', 'manager', 'admin'), async (req, res) => {
  const { data: w } = await supabase.from('cst_workers').select('name, id_number').eq('id', req.params.id).maybeSingle();
  const { error } = await supabase.from('cst_workers').delete().eq('id', req.params.id);
  if (error) {
    if (/foreign key|violates/i.test(error.message)) return res.status(409).json({ error: 'This worker already has wage payment records, so they cannot be removed.' });
    return res.status(500).json({ error: error.message });
  }
  audit(req, 'remove_worker', 'worker', req.params.id, w ? w.name + ' (' + w.id_number + ')' : null);
  res.json({ ok: true });
});

app.post('/api/attendance', requireAuth, requireRole('clerk', 'admin'), async (req, res) => {
  const { attendance_date, worker_ids, replace } = req.body;
  if (!Array.isArray(worker_ids) || (worker_ids.length === 0 && !replace)) return res.status(400).json({ error: 'worker_ids (non-empty array) is required' });
  const site_id = req.user.site_id;
  if (!site_id) return res.status(400).json({ error: 'Your account has no site associated with it' });
  const date = attendance_date || new Date().toISOString().slice(0, 10);

  // Only workers that belong to this site can be marked present.
  let validIds = [];
  if (worker_ids.length) {
    const { data: ws } = await supabase.from('cst_workers').select('id').eq('site_id', site_id).in('id', worker_ids);
    validIds = (ws || []).map(w => w.id);
  }
  if (replace) {
    const { error: delErr } = await supabase.from('cst_attendance').delete().eq('site_id', site_id).eq('attendance_date', date);
    if (delErr) return res.status(500).json({ error: delErr.message });
  }
  let recorded = 0;
  if (validIds.length) {
    const rows = validIds.map(worker_id => ({ worker_id, site_id, attendance_date: date, submitted_by: req.user.id }));
    const { data, error } = await supabase.from('cst_attendance').upsert(rows, { onConflict: 'worker_id,attendance_date', ignoreDuplicates: true }).select();
    if (error) return res.status(500).json({ error: error.message });
    recorded = replace ? validIds.length : data.length;
  }
  audit(req, 'record_attendance', 'attendance', null, date + ': ' + validIds.length + ' present' + (replace ? ' (corrected)' : ''));
  res.status(201).json({ ok: true, recorded, date });
});

app.get('/api/attendance', requireAuth, async (req, res) => {
  const date = req.query.date || new Date().toISOString().slice(0, 10);
  let query = supabase.from('cst_attendance').select('worker_id, attendance_date').eq('attendance_date', date);
  query = await scopeToSites(query, req.user);
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
  workerQuery = await scopeToSites(workerQuery, req.user);
  const { data: workers, error: wErr } = await workerQuery;
  if (wErr) throw wErr;

  let attQuery = supabase.from('cst_attendance').select('worker_id, attendance_date').gte('attendance_date', start).lte('attendance_date', endStr);
  attQuery = await scopeToSites(attQuery, req.user);
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

// ---------- Printable blank transaction form (for offline/no-network use) ----------
app.get('/api/forms/transaction-log.pdf', requireAuth, async (req, res) => {
  const branding = await fetchOrgBranding(req);
  const doc = new PDFDocument({ margin: 40, size: 'A4' });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename="site-transaction-log-form.pdf"');
  doc.pipe(res);

  drawFormHeader(doc, 'Manual Transaction Log', 'Construction Portal', branding);

  const infoBottom = drawInfoBox(doc, 130, [
    'Site name / ID:  ____________________________________________          Date:  ______________________',
    'Clerk name:  ____________________________________________          Signature:  ______________________'
  ]);

  const tableBottom = drawGridTable(doc, {
    startY: infoBottom + 16,
    columns: [
      { label: 'Date', width: 65 },
      { label: 'Category', width: 80 },
      { label: 'Description', width: 275 },
      { label: 'Amount (KES)', width: 95 }
    ],
    headerHeight: 22,
    rowHeight: 22,
    numRows: 20
  });

  doc.font('Helvetica-Bold').fontSize(10).fillColor(PDF_FOREST)
    .text('TOTAL', PDF_LEFT + 65 + 80, tableBottom + 10)
    .text('KES ________________', PDF_LEFT + 65 + 80 + 275, tableBottom + 10);

  drawFormFooter(doc, tableBottom + 45,
    'Fill this out by hand when there is no network connection. Once back online, submit these entries in the app and upload a photo of this sheet under "Documents" for the record.'
  );

  doc.end();
});

app.get('/api/forms/worker-registration.pdf', requireAuth, async (req, res) => {
  const branding = await fetchOrgBranding(req);
  const doc = new PDFDocument({ margin: 40, size: 'A4' });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename="worker-registration-form.pdf"');
  doc.pipe(res);

  drawFormHeader(doc, 'Worker Registration Form', 'Construction Portal', branding);

  const infoBottom = drawInfoBox(doc, 130, [
    'Site name / ID:  ____________________________________________          Date:  ______________________',
    'Recorded by:  ____________________________________________          Signature:  ______________________'
  ]);

  const tableBottom = drawGridTable(doc, {
    startY: infoBottom + 16,
    columns: [
      { label: 'Full Name', width: 135 },
      { label: 'ID Number', width: 90 },
      { label: 'Phone Number', width: 95 },
      { label: 'Designation', width: 100 },
      { label: 'Daily Rate', width: 95 }
    ],
    headerHeight: 22,
    rowHeight: 24,
    numRows: 20
  });

  drawFormFooter(doc, tableBottom + 20,
    'Fill this out by hand when there is no network connection. Once back online, enter each worker under "Enroll a worker" in the app, and upload a photo of this sheet under "Documents" for the record.'
  );

  doc.end();
});

// ---------- Manual document uploads (photos/scans of offline forms, receipts, etc.) ----------
// ---------- AI document extraction (OCR-style reading of uploaded forms/receipts) ----------
async function extractDocumentInfo(buffer, mimetype) {
  if (!process.env.ANTHROPIC_API_KEY) return { status: 'skipped', text: null, amount: null, payment_status: null };
  try {
    const base64 = buffer.toString('base64');
    const isPdf = mimetype === 'application/pdf';
    const contentBlock = isPdf
      ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } }
      : { type: 'image', source: { type: 'base64', media_type: mimetype, data: base64 } };

    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-5',
        max_tokens: 700,
        messages: [{
          role: 'user',
          content: [
            contentBlock,
            { type: 'text', text: 'This is a photo or scan of a construction site record — it may be a handwritten transaction log, a receipt, or an invoice. ' +
              'Read every transaction line you can find: date, category (materials/labor/equipment/fuel/other), description, and amount (KES). Also note any worker names, vendor names, site name, or other relevant details.\n\n' +
              'Respond with ONLY a JSON object (no markdown fences, no extra text) in exactly this shape:\n' +
              '{"summary": "<concise plain-text summary of what the document shows, suitable for a notification to a manager and finance officer, or a one-sentence note if unreadable>", ' +
              '"total_amount": <sum of every monetary amount found in the document as a plain number, or null if none found>, ' +
              '"payment_status": "<one of: paid, pending, unclear — paid if the document is a receipt/proof that money already changed hands, pending if it is an invoice/bill/unpaid claim, unclear if you cannot tell>"}' }
          ]
        }]
      })
    });
    const data = await r.json();
    if (!r.ok) return { status: 'failed', text: data.error?.message || 'Extraction failed', amount: null, payment_status: null };
    const raw = (data.content || []).filter(c => c.type === 'text').map(c => c.text).join('\n').trim();
    try {
      const cleaned = raw.replace(/^```json\s*|```$/g, '').trim();
      const parsed = JSON.parse(cleaned);
      return {
        status: 'done',
        text: parsed.summary || 'No readable content found.',
        amount: typeof parsed.total_amount === 'number' ? parsed.total_amount : null,
        payment_status: ['paid', 'pending', 'unclear'].includes(parsed.payment_status) ? parsed.payment_status : 'unclear'
      };
    } catch {
      // Fell back to plain text if the model didn't return valid JSON — still useful, just no structured totals.
      return { status: 'done', text: raw || 'No readable content found.', amount: null, payment_status: 'unclear' };
    }
  } catch (e) {
    return { status: 'failed', text: e.message, amount: null, payment_status: null };
  }
}

app.post('/api/documents', requireAuth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const site_id = req.user.site_id;
  if (!site_id) return res.status(400).json({ error: 'Your account has no site associated with it' });

  const safeName = req.file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
  const storagePath = `${site_id}/${Date.now()}-${safeName}`;

  const { error: upErr } = await supabase.storage.from('site-documents').upload(storagePath, req.file.buffer, {
    contentType: req.file.mimetype
  });
  if (upErr) return res.status(500).json({ error: upErr.message });

  const { data, error } = await supabase.from('cst_documents').insert({
    site_id, uploaded_by: req.user.id, file_name: req.file.originalname, storage_path: storagePath,
    description: req.body.description || null, transaction_id: req.body.transaction_id || null
  }).select().single();
  if (error) return res.status(500).json({ error: error.message });

  // Extract content and notify manager/finance. Awaited (not fire-and-forget) since Vercel
  // serverless functions don't reliably continue work after the response has been sent.
  const extraction = await extractDocumentInfo(req.file.buffer, req.file.mimetype);
  await supabase.from('cst_documents').update({
    extracted_text: extraction.text, extraction_status: extraction.status,
    extracted_amount: extraction.amount, extracted_payment_status: extraction.payment_status
  }).eq('id', data.id);

  if (extraction.status === 'done') {
    const { data: recipients } = await supabase.from('cst_users').select('id').eq('site_id', site_id).in('role', ['manager', 'finance']);
    const amountNote = extraction.amount ? ` Amount: KES ${extraction.amount.toLocaleString()} (${extraction.payment_status}).` : '';
    const msg = `Document uploaded: "${req.file.originalname}". Extracted: ${extraction.text.slice(0, 400)}${amountNote}`;
    for (const r of (recipients || [])) await notify(r.id, null, 'batch_submitted', msg);
  }

  res.status(201).json({ ...data, extracted_text: extraction.text, extraction_status: extraction.status, extracted_amount: extraction.amount, extracted_payment_status: extraction.payment_status });
});

app.get('/api/documents', requireAuth, async (req, res) => {
  let query = supabase.from('cst_documents').select('*, uploader:uploaded_by(full_name, role)').order('created_at', { ascending: false });
  query = await scopeToSites(query, req.user);
  const { data, error } = await query;
  if (error) return res.status(500).json({ error: error.message });

  const withUrls = await Promise.all(data.map(async d => {
    const { data: signed } = await supabase.storage.from('site-documents').createSignedUrl(d.storage_path, 3600);
    return { ...d, download_url: signed ? signed.signedUrl : null };
  }));
  res.json(withUrls);
});

app.delete('/api/documents/:id', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { data: doc, error: fe } = await supabase.from('cst_documents').select('storage_path').eq('id', req.params.id).single();
  if (fe || !doc) return res.status(404).json({ error: 'Document not found' });
  await supabase.storage.from('site-documents').remove([doc.storage_path]);
  const { error } = await supabase.from('cst_documents').delete().eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

// ---------- Store / material inventory ----------
async function getStoreItemsWithTotals(req) {
  let query = supabase.from('cst_store_items').select('*').order('name');
  query = await scopeToSites(query, req.user);
  const { data, error } = await query;
  if (error) throw new Error(error.message);

  let movQuery = supabase.from('cst_store_movements').select('item_id, movement_type, quantity');
  movQuery = await scopeToSites(movQuery, req.user);
  const { data: movements } = await movQuery;

  const totals = {};
  for (const m of (movements || [])) {
    if (!totals[m.item_id]) totals[m.item_id] = { total_in: 0, total_out: 0 };
    totals[m.item_id][m.movement_type === 'in' ? 'total_in' : 'total_out'] += Number(m.quantity);
  }

  return data.map(i => ({
    ...i,
    total_in: totals[i.id]?.total_in || 0,
    total_out: totals[i.id]?.total_out || 0
  }));
}

app.get('/api/store/items', requireAuth, async (req, res) => {
  try {
    const withTotals = await getStoreItemsWithTotals(req);
    res.json(withTotals);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/registry/attendance-report.pdf', requireAuth, async (req, res) => {
  try {
    const start = req.query.start || (() => {
      const d = new Date();
      const day = d.getDay();
      const diffToMonday = day === 0 ? 6 : day - 1;
      d.setDate(d.getDate() - diffToMonday);
      return d.toISOString().slice(0, 10);
    })();
    const startDate = new Date(start);
    const dayLabels = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    const days = dayLabels.map((label, i) => {
      const d = new Date(startDate);
      d.setDate(startDate.getDate() + i);
      return { label, date: d.toISOString().slice(0, 10) };
    });
    const endStr = days[6].date;

    let workerQuery = supabase.from('cst_workers').select('*').order('name');
    workerQuery = await scopeToSites(workerQuery, req.user);
    const { data: workers } = await workerQuery;

    let attQuery = supabase.from('cst_attendance').select('worker_id, attendance_date').gte('attendance_date', start).lte('attendance_date', endStr);
    attQuery = await scopeToSites(attQuery, req.user);
    const { data: attendance } = await attQuery;

    const presentSet = new Set((attendance || []).map(a => a.worker_id + '_' + a.attendance_date));
    const branding = await fetchOrgBranding(req);

    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="attendance-report-${start}-to-${endStr}.pdf"`);
    doc.pipe(res);

    drawFormHeader(doc, 'Weekly Attendance Report', 'Construction Portal', branding);
    const infoBottom = drawInfoBox(doc, 130, [
      `Week: ${start} to ${endStr}`,
      `Total workers: ${(workers || []).length}`
    ]);

    if (!workers || !workers.length) {
      doc.font('Helvetica').fontSize(11).fillColor('#555').text('No workers enrolled yet.', PDF_LEFT, infoBottom + 20);
    } else {
      const nameW = 105, desigW = 65, dayW = (PDF_RIGHT - PDF_LEFT - nameW - desigW - 40) / 7, totalW = 40;
      const headerH = 24, rowH = 22;
      let y = infoBottom + 16;
      const tableTop = y;

      // Header row
      doc.rect(PDF_LEFT, y, PDF_RIGHT - PDF_LEFT, headerH).fillAndStroke('#e2f1e8', PDF_FOREST);
      doc.font('Helvetica-Bold').fontSize(8).fillColor(PDF_FOREST);
      let x = PDF_LEFT;
      doc.text('Name', x + 4, y + 8, { width: nameW - 8 }); x += nameW;
      doc.text('Designation', x + 4, y + 8, { width: desigW - 8 }); x += desigW;
      days.forEach(d => { doc.text(d.label, x + 2, y + 8, { width: dayW - 4, align: 'center' }); x += dayW; });
      doc.text('Days', x + 2, y + 8, { width: totalW - 4, align: 'center' });
      y += headerH;

      workers.forEach(w => {
        let daysPresent = 0;
        let rowX = PDF_LEFT;
        doc.rect(PDF_LEFT, y, PDF_RIGHT - PDF_LEFT, rowH).strokeColor('#ddd').stroke();
        doc.font('Helvetica').fontSize(8.5).fillColor('#222');
        doc.text(w.name, rowX + 4, y + 6, { width: nameW - 8 }); rowX += nameW;
        doc.text(w.designation, rowX + 4, y + 6, { width: desigW - 8 }); rowX += desigW;
        days.forEach(d => {
          const present = presentSet.has(w.id + '_' + d.date);
          if (present) daysPresent++;
          doc.rect(rowX + 2, y + 3, dayW - 4, rowH - 6).fill(present ? '#dfeee5' : '#f6e2e2');
          doc.fillColor(present ? '#2f7d4f' : '#b03a3a').font('Helvetica-Bold').fontSize(8)
            .text(present ? 'P' : 'A', rowX + 2, y + 7, { width: dayW - 4, align: 'center' });
          rowX += dayW;
        });
        doc.fillColor('#222').font('Helvetica-Bold').fontSize(8.5).text(String(daysPresent), rowX + 2, y + 6, { width: totalW - 4, align: 'center' });
        y += rowH;
      });

      // Outer border for whole table
      doc.rect(PDF_LEFT, tableTop, PDF_RIGHT - PDF_LEFT, y - tableTop).strokeColor(PDF_FOREST).lineWidth(1.2).stroke();
      doc.lineWidth(1);

      // Legend
      y += 16;
      doc.rect(PDF_LEFT, y, 12, 12).fill('#dfeee5');
      doc.fillColor('#2f7d4f').font('Helvetica-Bold').fontSize(9).text('P', PDF_LEFT + 2, y + 2);
      doc.fillColor('#333').font('Helvetica').fontSize(9).text('Present', PDF_LEFT + 18, y + 2);
      doc.rect(PDF_LEFT + 90, y, 12, 12).fill('#f6e2e2');
      doc.fillColor('#b03a3a').font('Helvetica-Bold').fontSize(9).text('A', PDF_LEFT + 92, y + 2);
      doc.fillColor('#333').font('Helvetica').fontSize(9).text('Absent', PDF_LEFT + 108, y + 2);

      drawFormFooter(doc, y + 30, 'Generated automatically from daily attendance records.');
    }

    doc.end();
  } catch (e) { res.status(500).send('Could not generate report: ' + e.message); }
});

app.get('/api/store/report.pdf', requireAuth, async (req, res) => {
  try {
    const items = await getStoreItemsWithTotals(req);
    const branding = await fetchOrgBranding(req);
    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="stock-report.pdf"');
    doc.pipe(res);

    drawFormHeader(doc, 'Material Stock Report', 'Construction Portal', branding);
    const infoBottom = drawInfoBox(doc, 130, [
      'Generated: ' + new Date().toLocaleString(),
      'Total materials tracked: ' + items.length
    ]);

    if (!items.length) {
      doc.font('Helvetica').fontSize(11).fillColor('#555').text('No materials are being tracked yet.', PDF_LEFT, infoBottom + 20);
    } else {
      const rowsPerPage = 22;
      let idx = 0;
      let startY = infoBottom + 16;
      while (idx < items.length) {
        const pageItems = items.slice(idx, idx + rowsPerPage);
        const tableBottom = drawGridTable(doc, {
          startY,
          columns: [
            { label: 'Material', width: 165 },
            { label: 'Stock In', width: 90 },
            { label: 'Stock Out', width: 90 },
            { label: 'Remaining', width: 100 },
            { label: 'Reorder At', width: 70 }
          ],
          headerHeight: 22,
          rowHeight: 20,
          numRows: pageItems.length
        });
        doc.font('Helvetica').fontSize(9).fillColor('#222');
        pageItems.forEach((it, i) => {
          const rowY = startY + 22 + i * 20 + 5;
          doc.text(it.name, PDF_LEFT + 6, rowY, { width: 155 });
          doc.text(`${it.total_in} ${it.unit}`, PDF_LEFT + 171, rowY, { width: 84 });
          doc.text(`${it.total_out} ${it.unit}`, PDF_LEFT + 261, rowY, { width: 84 });
          doc.text(`${it.quantity_in_stock} ${it.unit}`, PDF_LEFT + 351, rowY, { width: 94 });
          doc.text(it.reorder_level ? String(it.reorder_level) : '—', PDF_LEFT + 451, rowY, { width: 64 });
        });
        idx += rowsPerPage;
        if (idx < items.length) { doc.addPage(); startY = 40; }
        else startY = tableBottom;
      }
      drawFormFooter(doc, startY + 20, 'Generated automatically from live stock records — reorder levels are highlighted in the app when stock runs low.');
    }

    doc.end();
  } catch (e) { res.status(500).send('Could not generate report: ' + e.message); }
});

app.post('/api/store/items', requireAuth, requireRole('storekeeper', 'admin'), async (req, res) => {
  const { name, unit, reorder_level, initial_quantity } = req.body;
  if (!name || !unit) return res.status(400).json({ error: 'name and unit are required' });
  const site_id = req.user.site_id;
  if (!site_id) return res.status(400).json({ error: 'Your account has no site associated with it' });

  const { data, error } = await supabase.from('cst_store_items').insert({
    site_id, name, unit, reorder_level: reorder_level || 0, quantity_in_stock: initial_quantity || 0, created_by: req.user.id
  }).select().single();
  if (error) {
    if (error.message.includes('duplicate')) return res.status(409).json({ error: 'An item with this name already exists at this site' });
    return res.status(500).json({ error: error.message });
  }

  if (initial_quantity && Number(initial_quantity) > 0) {
    await supabase.from('cst_store_movements').insert({
      item_id: data.id, site_id, movement_type: 'in', quantity: initial_quantity, reason: 'Initial stock', recorded_by: req.user.id
    });
  }
  res.status(201).json(data);
});

app.delete('/api/store/items/:id', requireAuth, requireRole('storekeeper', 'admin'), async (req, res) => {
  const { error } = await supabase.from('cst_store_items').delete().eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

app.post('/api/store/items/:id/movement', requireAuth, requireRole('storekeeper', 'admin'), async (req, res) => {
  const { movement_type, quantity, supplier_id } = req.body;
  let { reason } = req.body;
  const issued_to = String(req.body.issued_to || '').trim() || null, used_for = String(req.body.used_for || '').trim() || null;
  if (movement_type === 'out' && !issued_to) return res.status(400).json({ error: 'Say who the stock was issued to' });
  if (movement_type === 'out') reason = [reason, 'Issued to ' + issued_to, used_for ? 'for ' + used_for : null].filter(Boolean).join(' - ');
  if (!['in', 'out'].includes(movement_type)) return res.status(400).json({ error: 'movement_type must be "in" or "out"' });
  if (req.user.role === 'clerk' && movement_type !== 'in') return res.status(403).json({ error: 'Clerks can only record deliveries (stock in), not stock out' });
  if (!quantity || Number(quantity) <= 0) return res.status(400).json({ error: 'quantity must be greater than 0' });

  const { data: item, error: itemErr } = await supabase.from('cst_store_items').select('*').eq('id', req.params.id).single();
  if (itemErr || !item) return res.status(404).json({ error: 'Item not found' });

  const newQty = movement_type === 'in' ? Number(item.quantity_in_stock) + Number(quantity) : Number(item.quantity_in_stock) - Number(quantity);
  if (newQty < 0) return res.status(400).json({ error: `Not enough stock — only ${item.quantity_in_stock} ${item.unit} available` });

  const { error: updErr } = await supabase.from('cst_store_items').update({ quantity_in_stock: newQty, updated_at: new Date().toISOString() }).eq('id', item.id);
  if (updErr) return res.status(500).json({ error: updErr.message });

  const { data: movement, error: movErr } = await supabase.from('cst_store_movements').insert({
    item_id: item.id, site_id: item.site_id, movement_type, quantity, reason: reason || null, recorded_by: req.user.id,
    supplier_id: movement_type === 'in' ? (supplier_id || null) : null,
    issued_to: movement_type === 'out' ? issued_to : null, used_for: movement_type === 'out' ? used_for : null
  }).select().single();
  if (movErr) return res.status(500).json({ error: movErr.message });

  if (item.reorder_level && newQty <= Number(item.reorder_level)) {
    const { data: recipients } = await supabase.from('cst_users').select('id').eq('site_id', item.site_id).in('role', ['manager', 'storekeeper']);
    for (const r of (recipients || [])) await notify(r.id, null, 'batch_submitted', `Low stock alert: "${item.name}" is down to ${newQty} ${item.unit} (reorder level: ${item.reorder_level}).`);
  }

  audit(req, movement_type === 'out' ? 'stock_out' : 'stock_in', 'store_item', item.id, item.name + ' ' + quantity + ' ' + item.unit + (reason ? ' - ' + reason : ''));
  res.status(201).json({ item: { ...item, quantity_in_stock: newQty }, movement });
});

app.get('/api/store/items/:id/movements', requireAuth, async (req, res) => {
  const { data, error } = await supabase.from('cst_store_movements').select('*, recorder:recorded_by(full_name)').eq('item_id', req.params.id).order('created_at', { ascending: false }).limit(50);
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

// ---------- Analytics ----------
app.get('/api/analytics', requireAuth, requireRole('manager', 'finance', 'admin', 'superadmin'), async (req, res) => {
  try {
    const siteFilter = async (q) => await scopeToSites(q, req.user);

    const { data: txns } = await siteFilter(supabase.from('cst_transactions').select('status, category, amount, transaction_date'));
    const byCategory = {};
    const byStatus = { pending: 0, approved: 0, paid: 0, rejected: 0 };
    for (const t of (txns || [])) {
      const amt = Number(t.amount) || 0;
      byCategory[t.category] = (byCategory[t.category] || 0) + amt;
      if (byStatus[t.status] !== undefined) byStatus[t.status] += amt;
    }

    // Weekly trend: last 6 ISO weeks (Monday start)
    const weeks = [];
    const today = new Date();
    const day = today.getDay();
    const diffToMonday = day === 0 ? 6 : day - 1;
    const thisMonday = new Date(today);
    thisMonday.setDate(today.getDate() - diffToMonday);
    for (let i = 5; i >= 0; i--) {
      const start = new Date(thisMonday);
      start.setDate(thisMonday.getDate() - i * 7);
      const end = new Date(start);
      end.setDate(start.getDate() + 6);
      weeks.push({ start: start.toISOString().slice(0, 10), end: end.toISOString().slice(0, 10), total: 0, paid: 0 });
    }
    for (const t of (txns || [])) {
      const w = weeks.find(w => t.transaction_date >= w.start && t.transaction_date <= w.end);
      if (w) { w.total += Number(t.amount) || 0; if (t.status === 'paid') w.paid += Number(t.amount) || 0; }
    }

    const { data: wagePayments } = await siteFilter(supabase.from('cst_wage_payments').select('total_amount'));
    const wagesPaidTotal = (wagePayments || []).reduce((s, w) => s + Number(w.total_amount), 0);

    const financial = { by_category: byCategory, by_status: byStatus, weekly_trend: weeks, wages_paid_total: wagesPaidTotal };

    let operational = null;
    if (['manager', 'admin', 'superadmin'].includes(req.user.role)) {
      const { data: workers } = await siteFilter(supabase.from('cst_workers').select('id'));
      const todayStr = new Date().toISOString().slice(0, 10);
      const { data: todayAtt } = await siteFilter(supabase.from('cst_attendance').select('worker_id').eq('attendance_date', todayStr));
      const { data: items } = await siteFilter(supabase.from('cst_store_items').select('quantity_in_stock, reorder_level'));
      const lowStock = (items || []).filter(i => i.reorder_level && Number(i.quantity_in_stock) <= Number(i.reorder_level)).length;
      const { data: docs } = await siteFilter(supabase.from('cst_documents').select('extraction_status'));
      const docsRead = (docs || []).filter(d => d.extraction_status === 'done').length;

      operational = {
        worker_count: (workers || []).length,
        present_today: (todayAtt || []).length,
        store_item_count: (items || []).length,
        low_stock_count: lowStock,
        documents_count: (docs || []).length,
        documents_read_count: docsRead
      };
    }

    res.json({ financial, operational });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ---------- Suppliers (vendors with payment tracking) ----------
app.get('/api/suppliers', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  let query = supabase.from('cst_vendors').select('*').eq('role_type', 'vendor').order('name');
  query = await scopeToSites(query, req.user);
  const { data: suppliers, error } = await query;
  if (error) return res.status(500).json({ error: error.message });
  if (!suppliers.length) return res.json([]);

  const supplierIds = suppliers.map(s => s.id);
  const { data: links } = await supabase.from('cst_transaction_vendors').select('vendor_id, transaction_id').in('vendor_id', supplierIds);
  const txnIds = [...new Set((links || []).map(l => l.transaction_id))];
  let txnMap = {};
  if (txnIds.length) {
    const { data: txns } = await supabase.from('cst_transactions').select('id, amount, status').in('id', txnIds);
    for (const t of (txns || [])) txnMap[t.id] = t;
  }

  const result = suppliers.map(s => {
    const myLinks = (links || []).filter(l => l.vendor_id === s.id);
    let total_paid = 0, total_pending = 0, transaction_count = myLinks.length;
    for (const l of myLinks) {
      const t = txnMap[l.transaction_id];
      if (!t) continue;
      const amt = Number(t.amount) || 0;
      if (t.status === 'paid') total_paid += amt;
      else if (t.status === 'pending' || t.status === 'approved') total_pending += amt;
    }
    return { ...s, total_paid, total_pending, transaction_count };
  });
  res.json(result);
});

app.post('/api/suppliers', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { name, contact_phone, contact_email } = req.body;
  if (!name) return res.status(400).json({ error: 'name is required' });
  const { data, error } = await supabase.from('cst_vendors').insert({
    name, role_type: 'vendor', contact_phone, contact_email, site_id: req.user.site_id
  }).select().single();
  if (error) return res.status(500).json({ error: error.message });
  res.status(201).json(data);
});

app.get('/api/suppliers/:id/transactions', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { data: links, error } = await supabase.from('cst_transaction_vendors').select('transaction_id, supplied_item').eq('vendor_id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  if (!links.length) return res.json([]);
  const txnIds = links.map(l => l.transaction_id);
  const { data: txns } = await supabase.from('cst_transactions').select('*').in('id', txnIds).order('transaction_date', { ascending: false });
  const suppliedMap = {};
  for (const l of links) suppliedMap[l.transaction_id] = l.supplied_item;
  const merged = (txns || []).map(t => ({ ...t, supplied_item: suppliedMap[t.id] || null }));
  res.json(merged);
});

app.delete('/api/suppliers/:id', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  const { error } = await supabase.from('cst_vendors').delete().eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

app.get('/api/suppliers/:id/deliveries', requireAuth, requireRole('manager', 'finance', 'admin'), async (req, res) => {
  const { data: movements, error } = await supabase.from('cst_store_movements').select('*, item:item_id(name, unit), recorder:recorded_by(full_name)').eq('supplier_id', req.params.id).order('created_at', { ascending: false });
  if (error) return res.status(500).json({ error: error.message });
  res.json(movements || []);
});

// ---------- Combined delivery recording: supplier + materials + payment, all at once ----------
app.post('/api/deliveries', requireAuth, requireRole('storekeeper', 'admin'), async (req, res) => {
  const { supplier_id, new_supplier_name, item_id, new_item_name, new_item_unit, quantity, amount, description } = req.body;
  const site_id = req.user.site_id;
  if (!site_id) return res.status(400).json({ error: 'Your account has no site associated with it' });
  if (!quantity || Number(quantity) <= 0) return res.status(400).json({ error: 'quantity must be greater than 0' });
  if (!amount || Number(amount) <= 0) return res.status(400).json({ error: 'amount must be greater than 0' });

  // Resolve or create supplier
  let supplierId = supplier_id || null;
  if (!supplierId && new_supplier_name) {
    const { data: sup, error: supErr } = await supabase.from('cst_vendors').insert({
      name: new_supplier_name, role_type: 'vendor', site_id
    }).select().single();
    if (supErr) return res.status(500).json({ error: supErr.message });
    supplierId = sup.id;
  }

  // Resolve or create store item
  let item;
  if (item_id) {
    const { data, error: itemErr } = await supabase.from('cst_store_items').select('*').eq('id', item_id).single();
    if (itemErr || !data) return res.status(404).json({ error: 'Material not found' });
    item = data;
  } else if (new_item_name) {
    if (!new_item_unit) return res.status(400).json({ error: 'unit is required for a new material' });
    const { data, error: createErr } = await supabase.from('cst_store_items').insert({
      site_id, name: new_item_name, unit: new_item_unit, quantity_in_stock: 0, created_by: req.user.id
    }).select().single();
    if (createErr) {
      if (createErr.message.includes('duplicate')) return res.status(409).json({ error: 'A material with this name already exists — pick it from the list instead' });
      return res.status(500).json({ error: createErr.message });
    }
    item = data;
  } else {
    return res.status(400).json({ error: 'Select an existing material or provide a new material name and unit' });
  }

  // Stock in
  const newQty = Number(item.quantity_in_stock) + Number(quantity);
  await supabase.from('cst_store_items').update({ quantity_in_stock: newQty, updated_at: new Date().toISOString() }).eq('id', item.id);
  await supabase.from('cst_store_movements').insert({
    item_id: item.id, site_id, movement_type: 'in', quantity, reason: 'Delivery recorded', recorded_by: req.user.id, supplier_id: supplierId
  });

  // Create the payment transaction (pending — goes through the usual manager+finance approval)
  const desc = description || `Delivery: ${quantity} ${item.unit} of ${item.name}`;
  const { data: txn, error: txnErr } = await supabase.from('cst_transactions').insert({
    site_id, category: 'materials', description: desc, amount,
    transaction_date: new Date().toISOString().slice(0, 10), created_by: req.user.id
  }).select().single();
  if (txnErr) return res.status(500).json({ error: txnErr.message });

  if (supplierId) {
    await supabase.from('cst_transaction_vendors').insert({
      transaction_id: txn.id, vendor_id: supplierId, supplied_item: `${quantity} ${item.unit} of ${item.name}`
    });
  }

  const { data: recipients } = await supabase.from('cst_users').select('id').eq('site_id', site_id).in('role', ['manager', 'finance']);
  for (const r of (recipients || [])) await notify(r.id, txn.id, 'batch_submitted', `New delivery recorded: ${quantity} ${item.unit} of ${item.name} (KES ${Number(amount).toLocaleString()}) — awaiting approval.`);

  res.status(201).json({ item: { ...item, quantity_in_stock: newQty }, transaction: txn, supplier_id: supplierId });
});

// ---------- Material reorder request (goes through normal transaction approval) ----------
app.post('/api/store/items/:id/reorder-request', requireAuth, requireRole('storekeeper', 'admin'), async (req, res) => {
  const { quantity, estimated_amount, note } = req.body;
  if (!quantity || Number(quantity) <= 0) return res.status(400).json({ error: 'quantity must be greater than 0' });
  if (!estimated_amount || Number(estimated_amount) <= 0) return res.status(400).json({ error: 'estimated_amount must be greater than 0' });

  const { data: item, error: itemErr } = await supabase.from('cst_store_items').select('*').eq('id', req.params.id).single();
  if (itemErr || !item) return res.status(404).json({ error: 'Material not found' });

  const description = `Reorder request: ${quantity} ${item.unit} of ${item.name}${note ? ' — ' + note : ''} (current stock: ${item.quantity_in_stock} ${item.unit})`;
  const { data: txn, error: txnErr } = await supabase.from('cst_transactions').insert({
    site_id: item.site_id, category: 'materials', description, amount: estimated_amount,
    transaction_date: new Date().toISOString().slice(0, 10), created_by: req.user.id
  }).select().single();
  if (txnErr) return res.status(500).json({ error: txnErr.message });

  const { data: recipients } = await supabase.from('cst_users').select('id').eq('site_id', item.site_id).in('role', ['manager', 'finance']);
  for (const r of (recipients || [])) await notify(r.id, txn.id, 'batch_submitted', `Reorder request: ${quantity} ${item.unit} of ${item.name} (est. KES ${Number(estimated_amount).toLocaleString()}) — awaiting approval.`);

  res.status(201).json(txn);
});

// ---------- Superadmin: platform-level company (organization) management ----------
app.get('/api/superadmin/organizations', requireAuth, requireRole('superadmin'), async (req, res) => {
  const { data: orgs, error } = await supabase.from('cst_organizations').select('*').order('created_at', { ascending: false });
  if (error) return res.status(500).json({ error: error.message });

  const withCounts = await Promise.all(orgs.map(async o => {
    const { count: siteCount } = await supabase.from('cst_sites').select('id', { count: 'exact', head: true }).eq('organization_id', o.id);
    const { count: userCount } = await supabase.from('cst_users').select('id', { count: 'exact', head: true }).eq('organization_id', o.id);
    return { ...o, site_count: siteCount || 0, user_count: userCount || 0 };
  }));
  res.json(withCounts);
});

// ---------- Company logo (upload / replace) ----------
async function storeOrgLogo(file) {
  const okTypes = ['image/png', 'image/jpeg'];
  if (!file) throw new Error('Choose a logo image first');
  if (!okTypes.includes(file.mimetype)) throw new Error('Logo must be a PNG or JPG image');
  if (file.size > 2 * 1024 * 1024) throw new Error('Logo must be smaller than 2 MB');
  const safeName = file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
  const path = `${Date.now()}-${safeName}`;
  const { error: upErr } = await supabase.storage.from('org-logos').upload(path, file.buffer, { contentType: file.mimetype });
  if (upErr) throw new Error('Logo upload failed: ' + upErr.message);
  return supabase.storage.from('org-logos').getPublicUrl(path).data.publicUrl;
}

app.get('/api/organization', requireAuth, async (req, res) => {
  if (!req.user.organization_id) return res.json({ id: null, name: null, logo_url: null });
  const { data, error } = await supabase.from('cst_organizations').select('id, name, logo_url').eq('id', req.user.organization_id).single();
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.post('/api/organization/logo', requireAuth, requireRole('manager', 'admin'), upload.single('logo'), async (req, res) => {
  if (!req.user.organization_id) return res.status(400).json({ error: 'Your account is not linked to a company' });
  try {
    const logo_url = await storeOrgLogo(req.file);
    const { error } = await supabase.from('cst_organizations').update({ logo_url }).eq('id', req.user.organization_id);
    if (error) return res.status(500).json({ error: error.message });
    res.json({ logo_url });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.post('/api/superadmin/organizations/:id/logo', requireAuth, requireRole('superadmin'), upload.single('logo'), async (req, res) => {
  try {
    const logo_url = await storeOrgLogo(req.file);
    const { data, error } = await supabase.from('cst_organizations').update({ logo_url }).eq('id', req.params.id).select('id').single();
    if (error || !data) return res.status(404).json({ error: 'Company not found' });
    res.json({ logo_url });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.post('/api/superadmin/organizations', requireAuth, requireRole('superadmin'), upload.single('logo'), async (req, res) => {
  const { name, contact_email, contact_phone, address, registration_number, first_site_name, admin_full_name, admin_email } = req.body;
  if (!name) return res.status(400).json({ error: 'Company name is required' });

  let logo_url = null;
  if (req.file) {
    const safeName = req.file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
    const path = `${Date.now()}-${safeName}`;
    const { error: upErr } = await supabase.storage.from('org-logos').upload(path, req.file.buffer, { contentType: req.file.mimetype });
    if (upErr) return res.status(500).json({ error: 'Logo upload failed: ' + upErr.message });
    const { data: pub } = supabase.storage.from('org-logos').getPublicUrl(path);
    logo_url = pub.publicUrl;
  }

  const { data: org, error: orgErr } = await supabase.from('cst_organizations').insert({
    name, logo_url, contact_email: contact_email || null, contact_phone: contact_phone || null,
    address: address || null, registration_number: registration_number || null, created_by: req.user.id
  }).select().single();
  if (orgErr) return res.status(500).json({ error: orgErr.message });

  let site = null, adminAccount = null;
  if (first_site_name) {
    const { data: siteData, error: siteErr } = await supabase.from('cst_sites').insert({
      name: first_site_name, organization_id: org.id
    }).select().single();
    if (siteErr) return res.status(500).json({ error: 'Organization created, but site setup failed: ' + siteErr.message });
    site = siteData;

    if (admin_full_name && admin_email) {
      const plainPassword = randomPassword();
      const { data: adminData, error: adminErr } = await supabase.from('cst_users').insert({
        full_name: admin_full_name, email: admin_email.toLowerCase().trim(), password_hash: bcrypt.hashSync(plainPassword, 10),
        role: 'admin', site_id: site.id, organization_id: org.id
      }).select('id, full_name, email, role').single();
      if (adminErr) return res.status(500).json({ error: 'Organization and site created, but admin account failed: ' + adminErr.message });
      adminAccount = { ...adminData, temporary_password: plainPassword };
    }
  }

  res.status(201).json({ organization: org, site, admin: adminAccount });
});

app.patch('/api/superadmin/organizations/:id/status', requireAuth, requireRole('superadmin'), async (req, res) => {
  const { is_active } = req.body;
  if (typeof is_active !== 'boolean') return res.status(400).json({ error: 'is_active (boolean) is required' });
  const { data, error } = await supabase.from('cst_organizations').update({ is_active }).eq('id', req.params.id).select().single();
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

// =====================================================================
// Finance tools: payments ledger, supplier statements, site budgets,
// wage payment results (with retry) and finance summary PDF.
// =====================================================================
const FIN_ROLES = ['finance', 'manager', 'admin'];
const kes = n => 'KES ' + Number(n || 0).toLocaleString('en-KE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// Run an .in() query in chunks so long id lists never exceed URL limits.
async function inChunks(table, select, column, ids, extra) {
  const out = [];
  for (let i = 0; i < ids.length; i += 100) {
    let q = supabase.from(table).select(select).in(column, ids.slice(i, i + 100));
    if (extra) q = extra(q);
    const { data, error } = await q;
    if (error) throw new Error(error.message);
    out.push(...(data || []));
  }
  return out;
}

// Simple paginated PDF table (header repeats on each page).
function pdfTable(doc, startY, cols, rows, opts = {}) {
  const rowH = opts.rowHeight || 20, fontSize = opts.fontSize || 8;
  const tableW = cols.reduce((s, c) => s + c.width, 0);
  let y = startY;
  const drawHeader = () => {
    doc.rect(PDF_LEFT, y, tableW, 20).fillAndStroke('#e2f1e8', PDF_FOREST);
    doc.font('Helvetica-Bold').fontSize(fontSize).fillColor(PDF_FOREST);
    let x = PDF_LEFT;
    cols.forEach(c => { doc.text(c.label, x + 4, y + 6, { width: c.width - 8, align: c.align || 'left', lineBreak: false }); x += c.width; });
    y += 20;
  };
  drawHeader();
  rows.forEach((r, idx) => {
    if (y + rowH > 800) { doc.addPage(); y = 50; drawHeader(); }
    if (idx % 2 === 1) doc.rect(PDF_LEFT, y, tableW, rowH).fill('#f7faf8');
    doc.font(r.bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(fontSize).fillColor('#222');
    let x = PDF_LEFT;
    cols.forEach((c, i) => {
      doc.text(String(r.cells[i] == null ? '' : r.cells[i]), x + 4, y + 6, { width: c.width - 8, align: c.align || 'left', height: rowH - 4, ellipsis: true, lineBreak: false });
      x += c.width;
    });
    doc.moveTo(PDF_LEFT, y + rowH).lineTo(PDF_LEFT + tableW, y + rowH).strokeColor('#e3e3e3').lineWidth(0.5).stroke();
    y += rowH;
  });
  doc.lineWidth(1);
  return y;
}

// ---------- 1. Payments ledger ----------
async function getLedger(req) {
  const from = req.query.from || null, to = req.query.to || null;
  const q = String(req.query.q || '').toLowerCase().trim();

  let tq = supabase.from('cst_transactions').select('id, category, description, amount, paid_at, transaction_date').eq('status', 'paid').order('paid_at', { ascending: false });
  tq = await scopeToSites(tq, req.user);
  const { data: txns, error } = await tq;
  if (error) throw new Error(error.message);
  // Weekly wage runs are shown per worker below, so skip the summary transaction to avoid double counting.
  const expenses = (txns || []).filter(t => !(t.category === 'labor' && /^Weekly wages /.test(t.description || '')));

  const links = expenses.length ? await inChunks('cst_transaction_vendors', 'transaction_id, vendor:vendor_id(name)', 'transaction_id', expenses.map(t => t.id)) : [];
  const payeeByTxn = {};
  for (const l of links) if (l.vendor && l.vendor.name) payeeByTxn[l.transaction_id] = l.vendor.name;

  let wq = supabase.from('cst_wage_payments').select('id, total_amount, paid_at, week_start, week_end, days_present, disbursement_status, mpesa_receipt, worker:worker_id(name, phone_number)').neq('disbursement_status', 'failed').order('paid_at', { ascending: false });
  wq = await scopeToSites(wq, req.user);
  const { data: wages, error: wErr } = await wq;
  if (wErr) throw new Error(wErr.message);

  let rows = [];
  for (const t of expenses) rows.push({
    date: t.paid_at || t.transaction_date, type: 'Expense', payee: payeeByTxn[t.id] || '', category: t.category,
    description: t.description || '', amount: Number(t.amount) || 0, method: 'Recorded', reference: String(t.id).slice(0, 8)
  });
  for (const w of (wages || [])) rows.push({
    date: w.paid_at, type: 'Wages', payee: (w.worker && w.worker.name) || '', category: 'labor',
    description: 'Wages ' + w.week_start + ' to ' + w.week_end + ' (' + w.days_present + ' days)', amount: Number(w.total_amount) || 0,
    method: w.disbursement_status === 'confirmed' || w.disbursement_status === 'sent' ? 'M-Pesa' : 'Recorded', reference: w.mpesa_receipt || ''
  });
  rows = rows.filter(r => {
    const d = r.date ? String(r.date).slice(0, 10) : '';
    if (from && d < from) return false;
    if (to && d > to) return false;
    if (q && !(r.payee + ' ' + r.description + ' ' + r.category + ' ' + r.reference + ' ' + r.type).toLowerCase().includes(q)) return false;
    return true;
  });
  rows.sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
  return { rows, total: rows.reduce((s, r) => s + r.amount, 0) };
}

app.get('/api/ledger', requireAuth, requireRole(...FIN_ROLES), async (req, res) => {
  try { res.json(await getLedger(req)); } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/ledger/export', requireAuth, requireRole(...FIN_ROLES), async (req, res) => {
  try {
    const { rows, total } = await getLedger(req);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Construction Site Transactions';
    const ws = wb.addWorksheet('Payments Ledger');
    ws.columns = [
      { header: 'Date', key: 'date', width: 14 }, { header: 'Type', key: 'type', width: 10 }, { header: 'Paid to', key: 'payee', width: 26 },
      { header: 'Category', key: 'category', width: 14 }, { header: 'Description', key: 'description', width: 50 },
      { header: 'Method', key: 'method', width: 12 }, { header: 'Reference', key: 'reference', width: 16 }, { header: 'Amount (KES)', key: 'amount', width: 16 }
    ];
    ws.getRow(1).font = { bold: true };
    rows.forEach(r => ws.addRow({ ...r, date: r.date ? String(r.date).slice(0, 10) : '' }));
    ws.getColumn('amount').numFmt = '#,##0.00';
    const t = ws.addRow({ description: 'TOTAL PAID', amount: total }); t.font = { bold: true };
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="payments-ledger.xlsx"');
    await wb.xlsx.write(res); res.end();
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/ledger.pdf', requireAuth, requireRole(...FIN_ROLES), async (req, res) => {
  try {
    const { rows, total } = await getLedger(req);
    const branding = await fetchOrgBranding(req);
    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="payments-ledger.pdf"');
    doc.pipe(res);
    drawFormHeader(doc, 'Payments Ledger', 'Every payment made', branding);
    const info = ['Period:  ' + (req.query.from || 'all time') + '  to  ' + (req.query.to || 'today') + (req.query.q ? '      Search: ' + req.query.q : ''),
      'Payments: ' + rows.length + '          Total paid: ' + kes(total) + '          Generated: ' + new Date().toLocaleString('en-KE')];
    const y = drawInfoBox(doc, 130, info) + 14;
    const cols = [{ label: 'Date', width: 55 }, { label: 'Type', width: 45 }, { label: 'Paid to', width: 105 }, { label: 'Description', width: 175 }, { label: 'Method', width: 50 }, { label: 'Amount (KES)', width: 85, align: 'right' }];
    const body = rows.map(r => ({ cells: [r.date ? String(r.date).slice(0, 10) : '', r.type, r.payee, r.description, r.method, Number(r.amount).toLocaleString('en-KE', { minimumFractionDigits: 2 })] }));
    body.push({ bold: true, cells: ['', '', '', 'TOTAL PAID', '', Number(total).toLocaleString('en-KE', { minimumFractionDigits: 2 })] });
    const end = pdfTable(doc, y, cols, body);
    drawFormFooter(doc, Math.min(end + 14, 800), 'Generated automatically from payment records.');
    doc.end();
  } catch (e) { if (!res.headersSent) res.status(500).send('Could not generate ledger: ' + e.message); else res.end(); }
});

// ---------- 2. Supplier statement PDF ----------
app.get('/api/suppliers/:id/statement.pdf', requireAuth, requireRole(...FIN_ROLES), async (req, res) => {
  try {
    const { data: vrows } = await scopeToSites(supabase.from('cst_vendors').select('*').eq('id', req.params.id), req.user);
    const supplier = vrows && vrows[0];
    if (!supplier) return res.status(404).send('Supplier not found');

    const { data: moves } = await supabase.from('cst_store_movements').select('created_at, quantity, reason, item:item_id(name, unit)').eq('supplier_id', supplier.id).eq('movement_type', 'in').order('created_at');
    const { data: links } = await supabase.from('cst_transaction_vendors').select('transaction_id, supplied_item').eq('vendor_id', supplier.id);
    const txns = (links && links.length) ? await inChunks('cst_transactions', 'id, description, amount, status, transaction_date, paid_at', 'id', links.map(l => l.transaction_id)) : [];
    const suppliedBy = {}; for (const l of (links || [])) suppliedBy[l.transaction_id] = l.supplied_item;
    txns.sort((a, b) => String(a.transaction_date).localeCompare(String(b.transaction_date)));

    let billed = 0, paid = 0, owed = 0;
    for (const t of txns) {
      const a = Number(t.amount) || 0;
      if (t.status === 'paid') { paid += a; billed += a; }
      else if (t.status === 'pending' || t.status === 'approved') { owed += a; billed += a; }
    }

    const branding = await fetchOrgBranding(req);
    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="supplier-statement-' + supplier.name.replace(/[^a-zA-Z0-9]+/g, '-') + '.pdf"');
    doc.pipe(res);
    drawFormHeader(doc, 'Supplier Statement', 'Deliveries, payments and balance', branding);
    let y = drawInfoBox(doc, 130, [
      'Supplier:  ' + supplier.name + '          Contact:  ' + [supplier.contact_phone, supplier.contact_email].filter(Boolean).join(' / '),
      'Statement date:  ' + new Date().toLocaleDateString('en-KE') + '          Total billed: ' + kes(billed) + '          Paid: ' + kes(paid) + '          OUTSTANDING: ' + kes(owed)
    ]) + 16;

    doc.font('Helvetica-Bold').fontSize(11).fillColor(PDF_FOREST).text('Deliveries to site', PDF_LEFT, y); y += 16;
    const dRows = (moves || []).map(m => ({ cells: [String(m.created_at).slice(0, 10), m.item ? m.item.name : '', Number(m.quantity) + ' ' + (m.item ? m.item.unit : ''), m.reason || ''] }));
    if (!dRows.length) dRows.push({ cells: ['', 'No deliveries recorded', '', ''] });
    y = pdfTable(doc, y, [{ label: 'Date', width: 70 }, { label: 'Material', width: 200 }, { label: 'Quantity', width: 90 }, { label: 'Note', width: 155 }], dRows) + 18;

    if (y > 700) { doc.addPage(); y = 50; }
    doc.font('Helvetica-Bold').fontSize(11).fillColor(PDF_FOREST).text('Invoices and payments', PDF_LEFT, y); y += 16;
    const tRows = txns.map(t => ({ cells: [String(t.transaction_date).slice(0, 10), (suppliedBy[t.id] ? suppliedBy[t.id] + ' - ' : '') + (t.description || ''), t.status.toUpperCase(), t.status === 'paid' ? String(t.paid_at || '').slice(0, 10) : '', Number(t.amount).toLocaleString('en-KE', { minimumFractionDigits: 2 })] }));
    if (!tRows.length) tRows.push({ cells: ['', 'No transactions linked yet', '', '', ''] });
    tRows.push({ bold: true, cells: ['', 'Total billed (excluding rejected)', '', '', Number(billed).toLocaleString('en-KE', { minimumFractionDigits: 2 })] });
    tRows.push({ bold: true, cells: ['', 'Paid', '', '', Number(paid).toLocaleString('en-KE', { minimumFractionDigits: 2 })] });
    tRows.push({ bold: true, cells: ['', 'Still owed (pending + approved)', '', '', Number(owed).toLocaleString('en-KE', { minimumFractionDigits: 2 })] });
    y = pdfTable(doc, y, [{ label: 'Date', width: 65 }, { label: 'Description', width: 210 }, { label: 'Status', width: 60 }, { label: 'Paid on', width: 65 }, { label: 'Amount (KES)', width: 115, align: 'right' }], tRows);
    drawFormFooter(doc, Math.min(y + 14, 800), 'Generated automatically. Rejected items are not counted in the balance.');
    doc.end();
  } catch (e) { if (!res.headersSent) res.status(500).send('Could not generate statement: ' + e.message); else res.end(); }
});

// ---------- 3. Site budget ----------
async function resolveBudgetSite(req) {
  if (req.user.site_id) return req.user.site_id;
  const ids = await getOrgSiteIds(req.user.organization_id);
  const wanted = req.query.site_id || req.body?.site_id;
  if (wanted) return ids.includes(wanted) ? wanted : null;
  return ids[0] || null;
}
async function computeBudget(siteId) {
  const [{ data: site }, { data: b }, { data: txns }] = await Promise.all([
    supabase.from('cst_sites').select('id, name').eq('id', siteId).maybeSingle(),
    supabase.from('cst_site_budgets').select('*').eq('site_id', siteId).maybeSingle(),
    supabase.from('cst_transactions').select('amount, status').eq('site_id', siteId)
  ]);
  let paid = 0, approved = 0, pending = 0;
  for (const t of (txns || [])) {
    const a = Number(t.amount) || 0;
    if (t.status === 'paid') paid += a; else if (t.status === 'approved') approved += a; else if (t.status === 'pending') pending += a;
  }
  const budget = b ? Number(b.budget_amount) : null, warn = b ? b.warn_percent : 80;
  const used = paid + approved;
  const percent = budget ? (used / budget) * 100 : null;
  let level = 'none';
  if (budget) level = percent >= 100 ? 'over' : percent >= warn ? 'warning' : 'ok';
  return { site_id: siteId, site_name: site ? site.name : '', budget_amount: budget, warn_percent: warn, paid, approved_unpaid: approved, pending, used, remaining: budget == null ? null : budget - used, percent_used: percent, level };
}
app.get('/api/budget', requireAuth, requireRole(...FIN_ROLES), async (req, res) => {
  try {
    const siteId = await resolveBudgetSite(req);
    if (!siteId) return res.status(400).json({ error: 'No site found for your account' });
    res.json(await computeBudget(siteId));
  } catch (e) { res.status(500).json({ error: e.message }); }
});
app.put('/api/budget', requireAuth, requireRole('manager', 'admin'), async (req, res) => {
  try {
    const siteId = await resolveBudgetSite(req);
    if (!siteId) return res.status(400).json({ error: 'No site found for your account' });
    const amount = Number(req.body.budget_amount), warn = Number(req.body.warn_percent || 80);
    if (!(amount >= 0) || !isFinite(amount)) return res.status(400).json({ error: 'Enter a valid budget amount' });
    if (!(warn >= 1 && warn <= 100)) return res.status(400).json({ error: 'Warning level must be between 1 and 100' });
    const { error } = await supabase.from('cst_site_budgets').upsert({ site_id: siteId, budget_amount: amount, warn_percent: Math.round(warn), updated_by: req.user.id, updated_at: new Date().toISOString() });
    if (error) return res.status(500).json({ error: error.message });
    audit(req, 'set_budget', 'budget', siteId, 'Budget set to KES ' + amount + ', warn at ' + Math.round(warn) + '%');
    res.json(await computeBudget(siteId));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ---------- 4. Wage payment results + retry failed ----------
app.get('/api/wages/results', requireAuth, requireRole(...FIN_ROLES), async (req, res) => {
  try {
    const { start, end } = await computeWeeklyWages(req);
    let q = supabase.from('cst_wage_payments').select('id, total_amount, days_present, paid_at, disbursement_status, mpesa_receipt, mpesa_result_desc, worker:worker_id(name, phone_number, designation)').eq('week_start', start).order('paid_at', { ascending: false });
    q = await scopeToSites(q, req.user);
    const { data, error } = await q;
    if (error) return res.status(500).json({ error: error.message });
    const counts = { confirmed: 0, sent: 0, failed: 0, not_attempted: 0 };
    let failedTotal = 0;
    for (const r of (data || [])) { counts[r.disbursement_status] = (counts[r.disbursement_status] || 0) + 1; if (r.disbursement_status === 'failed') failedTotal += Number(r.total_amount) || 0; }
    res.json({ start, end, rows: data || [], counts, failed_total: failedTotal, mpesa_enabled: mpesaConfigured() });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/wages/weekly/retry-failed', requireAuth, requireRole('finance', 'admin'), async (req, res) => {
  try {
    if (!mpesaConfigured()) return res.status(400).json({ error: 'M-Pesa is not set up on this system, so there is nothing to retry.' });
    const { start, end } = await computeWeeklyWages(req);
    let q = supabase.from('cst_wage_payments').select('id, total_amount, worker:worker_id(name, phone_number)').eq('week_start', start).eq('disbursement_status', 'failed');
    q = await scopeToSites(q, req.user);
    const { data: failed, error } = await q;
    if (error) return res.status(500).json({ error: error.message });
    if (!failed.length) return res.json({ ok: true, retried: 0, results: [], message: 'No failed payments for this week.' });

    const results = [];
    for (const p of failed) {
      const phone = p.worker && p.worker.phone_number;
      if (!phone) { results.push({ name: p.worker ? p.worker.name : '?', ok: false, error: 'No phone number on file' }); continue; }
      let status = 'failed', convId = null, desc = null;
      try {
        const b2c = await sendB2CPayment({ phone, amount: Number(p.total_amount), remarks: 'Wages ' + start + ' to ' + end + ' (retry)', occasion: p.worker.name, callbackHost: req.get('host') });
        status = b2c.ok ? 'sent' : 'failed';
        convId = b2c.data.ConversationID || null;
        desc = b2c.data.ResponseDescription || b2c.data.errorMessage || JSON.stringify(b2c.data);
      } catch (e) { desc = e.message; }
      await supabase.from('cst_wage_payments').update({ disbursement_status: status, mpesa_conversation_id: convId, mpesa_result_desc: desc, paid_by: req.user.id, paid_at: new Date().toISOString() }).eq('id', p.id);
      results.push({ name: p.worker.name, ok: status === 'sent', status, detail: desc });
    }
    res.json({ ok: true, retried: failed.length, sent: results.filter(r => r.ok).length, results });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ---------- 5. Finance summary PDF (daily / weekly / monthly) ----------
app.get('/api/reports/finance-summary.pdf', requireAuth, requireRole(...FIN_ROLES), async (req, res) => {
  try {
    const period = ['daily', 'weekly', 'monthly'].includes(req.query.period) ? req.query.period : 'weekly';
    const endDate = /^\d{4}-\d{2}-\d{2}$/.test(req.query.date || '') ? req.query.date : new Date().toISOString().slice(0, 10);
    const e = new Date(endDate + 'T00:00:00Z');
    const s = new Date(e);
    if (period === 'weekly') s.setUTCDate(s.getUTCDate() - 6);
    else if (period === 'monthly') s.setUTCDate(1);
    const from = s.toISOString().slice(0, 10), to = endDate;

    let tq = supabase.from('cst_transactions').select('category, description, amount, status, transaction_date, paid_at').gte('transaction_date', from).lte('transaction_date', to);
    tq = await scopeToSites(tq, req.user);
    const { data: txns, error } = await tq;
    if (error) throw new Error(error.message);
    let oq = supabase.from('cst_transactions').select('amount, status').in('status', ['pending', 'approved']);
    oq = await scopeToSites(oq, req.user);
    const { data: open } = await oq;
    let wq = supabase.from('cst_wage_payments').select('total_amount, disbursement_status').neq('disbursement_status', 'failed').gte('paid_at', from + 'T00:00:00Z').lte('paid_at', to + 'T23:59:59Z');
    wq = await scopeToSites(wq, req.user);
    const { data: wages } = await wq;

    const by = { pending: [0, 0], approved: [0, 0], paid: [0, 0], rejected: [0, 0] };
    const cat = {};
    for (const t of (txns || [])) {
      const a = Number(t.amount) || 0;
      if (by[t.status]) { by[t.status][0]++; by[t.status][1] += a; }
      if (t.status === 'paid') cat[t.category] = (cat[t.category] || 0) + a;
    }
    const openTotal = (open || []).reduce((s2, t) => s2 + (Number(t.amount) || 0), 0);
    const wagesTotal = (wages || []).reduce((s2, w) => s2 + (Number(w.total_amount) || 0), 0);

    let budget = null;
    try { const sid = await resolveBudgetSite(req); if (sid) budget = await computeBudget(sid); } catch {}

    const branding = await fetchOrgBranding(req);
    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="finance-summary-' + period + '-' + to + '.pdf"');
    doc.pipe(res);
    drawFormHeader(doc, period.charAt(0).toUpperCase() + period.slice(1) + ' Finance Summary', 'Pending, approved, paid and rejected totals', branding);
    let y = drawInfoBox(doc, 130, ['Period:  ' + from + '  to  ' + to, 'Generated:  ' + new Date().toLocaleString('en-KE') + '          Waiting to be paid (all time, pending + approved): ' + kes(openTotal)]) + 16;

    doc.font('Helvetica-Bold').fontSize(11).fillColor(PDF_FOREST).text('Transactions by status', PDF_LEFT, y); y += 16;
    const money2 = n => Number(n).toLocaleString('en-KE', { minimumFractionDigits: 2 });
    const stRows = ['pending', 'approved', 'paid', 'rejected'].map(k => ({ cells: [k.toUpperCase(), by[k][0], money2(by[k][1])] }));
    stRows.push({ bold: true, cells: ['TOTAL', Object.values(by).reduce((s2, v) => s2 + v[0], 0), money2(Object.values(by).reduce((s2, v) => s2 + v[1], 0))] });
    y = pdfTable(doc, y, [{ label: 'Status', width: 200 }, { label: 'Count', width: 100, align: 'right' }, { label: 'Amount (KES)', width: 215, align: 'right' }], stRows, { rowHeight: 22, fontSize: 9 }) + 18;

    doc.font('Helvetica-Bold').fontSize(11).fillColor(PDF_FOREST).text('Paid, by category', PDF_LEFT, y); y += 16;
    const cRows = Object.entries(cat).map(([k, v]) => ({ cells: [k, money2(v)] }));
    if (!cRows.length) cRows.push({ cells: ['No payments in this period', ''] });
    y = pdfTable(doc, y, [{ label: 'Category', width: 300 }, { label: 'Paid (KES)', width: 215, align: 'right' }], cRows, { rowHeight: 20, fontSize: 9 }) + 18;

    doc.font('Helvetica').fontSize(10).fillColor('#222').text('Worker wages paid this period: ' + kes(wagesTotal) + ' (' + (wages || []).length + ' payments)', PDF_LEFT, y); y += 22;
    if (budget && budget.budget_amount != null) {
      const lvl = budget.level === 'over' ? 'OVER BUDGET' : budget.level === 'warning' ? 'WARNING - nearly used up' : 'On track';
      doc.font('Helvetica-Bold').fontSize(11).fillColor(PDF_FOREST).text('Site budget (' + budget.site_name + ')', PDF_LEFT, y); y += 16;
      doc.font('Helvetica').fontSize(10).fillColor('#222').text('Budget: ' + kes(budget.budget_amount) + '     Used (paid + approved): ' + kes(budget.used) + ' (' + budget.percent_used.toFixed(1) + '%)     Remaining: ' + kes(budget.remaining) + '     Status: ' + lvl, PDF_LEFT, y, { width: PDF_RIGHT - PDF_LEFT }); y += 30;
    }
    drawFormFooter(doc, Math.min(y + 10, 800), 'Generated automatically from transaction and wage records.');
    doc.end();
  } catch (e) { if (!res.headersSent) res.status(500).send('Could not generate summary: ' + e.message); else res.end(); }
});

app.get('/api/health', (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
if (require.main === module) {
  app.listen(PORT, () => console.log(`Construction site transactions API running on :${PORT}`));
}

module.exports = app;
