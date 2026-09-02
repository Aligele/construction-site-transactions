const API = '/api';
let state = { token: localStorage.getItem('cst_token'), user: JSON.parse(localStorage.getItem('cst_user') || 'null') };

const app = document.getElementById('app');

function save(token, user) {
  state = { token, user };
  localStorage.setItem('cst_token', token);
  localStorage.setItem('cst_user', JSON.stringify(user));
}
function logout() {
  localStorage.removeItem('cst_token');
  localStorage.removeItem('cst_user');
  state = { token: null, user: null };
  render();
}

async function api(path, opts = {}) {
  const res = await fetch(API + path, {
    ...opts,
    headers: {
      'Content-Type': 'application/json',
      ...(state.token ? { Authorization: 'Bearer ' + state.token } : {}),
      ...(opts.headers || {})
    }
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || 'Request failed');
  }
  return res.json();
}

function money(n) {
  return 'KES ' + Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function badge(status) {
  return `<span class="badge ${status}">${status}</span>`;
}

async function render() {
  if (!state.token) return renderLogin();
  return renderDashboard();
}

function renderLogin() {
  app.innerHTML = `
    <div class="card" style="max-width:360px;margin:60px auto;">
      <h1>Site Transactions</h1>
      <p class="muted">Sign in to continue</p>
      <label>Email</label>
      <input id="email" type="email" placeholder="you@example.com" />
      <label>Password</label>
      <input id="password" type="password" placeholder="••••••••" />
      <div style="margin-top:16px;"><button id="loginBtn" style="width:100%;">Log in</button></div>
      <div class="error" id="err"></div>
    </div>
  `;
  document.getElementById('loginBtn').onclick = async () => {
    const email = document.getElementById('email').value.trim();
    const password = document.getElementById('password').value;
    try {
      const { token, user } = await api('/auth/login', { method: 'POST', body: JSON.stringify({ email, password }) });
      save(token, user);
      render();
    } catch (e) {
      document.getElementById('err').textContent = e.message;
    }
  };
}

async function renderDashboard() {
  const { user } = state;
  app.innerHTML = `
    <div class="topbar">
      <div>
        <h1>Site Transactions</h1>
        <p class="muted">${user.full_name} · ${user.role}</p>
      </div>
      <button class="secondary" id="logoutBtn">Log out</button>
    </div>
    <div id="content"></div>
  `;
  document.getElementById('logoutBtn').onclick = logout;

  const content = document.getElementById('content');

  if (user.role === 'clerk' || user.role === 'admin') {
    content.innerHTML += `
      <div class="card">
        <h2>Log a transaction</h2>
        <div class="row">
          <div><label>Site ID</label><input id="site_id" placeholder="site UUID" value="${user.site_id || ''}" /></div>
          <div><label>Category</label>
            <select id="category">
              <option value="materials">Materials</option>
              <option value="labor">Labor</option>
              <option value="equipment">Equipment</option>
              <option value="fuel">Fuel</option>
              <option value="other">Other</option>
            </select>
          </div>
          <div><label>Amount (KES)</label><input id="amount" type="number" step="0.01" /></div>
        </div>
        <label>Description</label>
        <textarea id="description" rows="2" style="width:100%;"></textarea>
        <div style="margin-top:12px;"><button id="submitTxn">Submit</button></div>
        <div class="error" id="submitErr"></div>
      </div>
    `;
    document.getElementById('submitTxn').onclick = async () => {
      const site_id = document.getElementById('site_id').value.trim();
      const category = document.getElementById('category').value;
      const amount = document.getElementById('amount').value;
      const description = document.getElementById('description').value.trim();
      try {
        await api('/transactions', { method: 'POST', body: JSON.stringify({ site_id, category, amount, description }) });
        document.getElementById('description').value = '';
        document.getElementById('amount').value = '';
        loadTable();
      } catch (e) {
        document.getElementById('submitErr').textContent = e.message;
      }
    };
  }

  content.innerHTML += `
    <div class="card">
      <div class="topbar">
        <h2>Transactions</h2>
        ${['manager', 'finance', 'admin'].includes(user.role) ? '<button id="exportBtn" class="secondary">Download Excel</button>' : ''}
      </div>
      <div id="table"></div>
    </div>
  `;

  if (document.getElementById('exportBtn')) {
    document.getElementById('exportBtn').onclick = () => {
      fetch(`${API}/export`, { headers: { Authorization: 'Bearer ' + state.token } })
        .then(r => r.blob())
        .then(blob => {
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `site-transactions-${new Date().toISOString().slice(0, 10)}.xlsx`;
          a.click();
        });
    };
  }

  loadTable();
}

async function loadTable() {
  const { user } = state;
  const rows = await api('/transactions');
  const canAct = ['manager', 'finance', 'admin'].includes(user.role);

  const html = `
    <table>
      <thead>
        <tr>
          <th>Date</th><th>Category</th><th>Description</th><th>Amount</th><th>Status</th>${canAct ? '<th>Actions</th>' : ''}
        </tr>
      </thead>
      <tbody>
        ${rows.map(t => `
          <tr>
            <td>${t.transaction_date}</td>
            <td>${t.category}</td>
            <td>${t.description}</td>
            <td>${money(t.amount)}</td>
            <td>${badge(t.status)}</td>
            ${canAct ? `<td class="actions">${actionButtons(t, user)}</td>` : ''}
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
  document.getElementById('table').innerHTML = html;

  document.querySelectorAll('[data-action]').forEach(btn => {
    btn.onclick = async () => {
      const { action, id } = btn.dataset;
      try {
        await api(`/transactions/${id}/${action}`, { method: 'POST', body: JSON.stringify({}) });
        loadTable();
      } catch (e) {
        alert(e.message);
      }
    };
  });
}

function actionButtons(t, user) {
  if (t.status !== 'pending' && t.status !== 'approved') return '';
  let buttons = '';
  if (t.status === 'pending') {
    if ((user.role === 'manager' || user.role === 'admin') && !t.manager_approved) {
      buttons += `<button class="success" data-action="approve/manager" data-id="${t.id}">Approve (Manager)</button>`;
    }
    if ((user.role === 'finance' || user.role === 'admin') && !t.finance_approved) {
      buttons += `<button class="success" data-action="approve/finance" data-id="${t.id}">Approve (Finance)</button>`;
    }
    buttons += `<button class="danger" data-action="reject" data-id="${t.id}">Reject</button>`;
  }
  if (t.status === 'approved' && (user.role === 'finance' || user.role === 'admin')) {
    buttons += `<button data-action="pay" data-id="${t.id}">Mark Paid</button>`;
  }
  return buttons;
}

render();
