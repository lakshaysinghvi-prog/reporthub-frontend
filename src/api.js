// src/api.js  —  All calls to the ReportHub backend
// VITE_API_URL is set in .env (e.g. https://reporthub-api.railway.app)

const BASE = import.meta.env.VITE_API_URL || 'http://localhost:3001';

function getToken() {
  return localStorage.getItem('rh_token');
}

function authHeaders() {
  return {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${getToken()}`
  };
}

async function api(path, options = {}) {
  const res = await fetch(`${BASE}${path}`, {
    headers: authHeaders(),
    ...options,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: res.statusText }));
    if (res.status === 401 && err.error !== 'needs_auth') {
      // Only clear session for ReportHub JWT failures (Invalid token / No token).
      // Do NOT clear for third-party OAuth failures (needs_auth from Microsoft/Google).
      localStorage.removeItem('rh_token');
      throw new Error('Session expired. Please log in again.');
    }
    throw new Error(err.message || err.error || `HTTP ${res.status}`);
  }
  return res.json();
}

// ── Auth ─────────────────────────────────────────────────────────────────────
export async function login(username, password) {
  const res = await fetch(`${BASE}/api/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password })
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || 'Login failed');
  }
  const data = await res.json();
  localStorage.setItem('rh_token', data.token);
  if (data.id) localStorage.setItem('rh_user_id', data.id);
  return data; // { token, role, username, id }
}

export function logout() {
  localStorage.removeItem('rh_token');
}

export function isLoggedIn() {
  return !!getToken();
}

// ── Users ────────────────────────────────────────────────────────────────────
export const getUsers = () => api('/api/users');

export const createUser = (username, password, role) =>
  api('/api/users', { method: 'POST', body: JSON.stringify({ username, password, role }) });

export const updatePassword = (id, password) =>
  api(`/api/users/${id}/password`, { method: 'PATCH', body: JSON.stringify({ password }) });

export const updateRole = (id, role) =>
  api(`/api/users/${id}/role`, { method: 'PATCH', body: JSON.stringify({ role }) });

export const approveUser = (id) =>
  api(`/api/users/${id}/approve`, { method: 'PATCH' });

export const getRefreshSchedule = (id) =>
  api(`/api/reports/${id}/refresh-schedule`);

export const setRefreshSchedule = (id, intervalMinutes, enabled) =>
  api(`/api/reports/${id}/refresh-schedule`, {
    method: 'PUT',
    body: JSON.stringify({ intervalMinutes, enabled }),
  });

export const deleteUser = (id) =>
  api(`/api/users/${id}`, { method: 'DELETE' });

// ── Reports ──────────────────────────────────────────────────────────────────
export const getReports = () => api('/api/reports');

export const createReport = (payload) =>
  api('/api/reports', { method: 'POST', body: JSON.stringify(payload) });

export const updateReport = (id, payload) =>
  api(`/api/reports/${id}`, { method: 'PUT', body: JSON.stringify(payload) });

export const updateReportConfig = (id, config, name) =>
  api(`/api/reports/${id}/config`, { method: 'PATCH', body: JSON.stringify({ config, ...(name?{name}:{}) }) });

export const deleteReport = (id) =>
  api(`/api/reports/${id}`, { method: 'DELETE' });

export const publishReport = (id) =>
  api(`/api/reports/${id}/publish`, { method: 'PATCH' });

export const unpublishReport = (id) =>
  api(`/api/reports/${id}/unpublish`, { method: 'PATCH' });

export const getReportData = (id) =>
  api(`/api/reports/${id}/data`);

export const refreshReportUrl = (id, url, sheetName) =>
  api(`/api/reports/${id}/refresh-url`, {
    method: 'POST',
    body: JSON.stringify({ url, sheetName })
  });

// Proxy-fetch a URL via backend (bypasses browser CORS)
// Works with OneDrive, Dropbox, Google Drive, SharePoint
export const fetchUrlViaProxy = (url, sheetName) =>
  api('/api/fetch-url', {
    method: 'POST',
    body: JSON.stringify({ url, sheetName })
  });

// ── Public API (no auth — for mobile/shared link access) ──────────────────────
export const getPublishedReports = () => {
  return fetch(`${BASE}/api/public/reports`)
    .then(r => r.ok ? r.json() : Promise.reject(new Error('HTTP ' + r.status)));
};

export const getPublishedReportData = (id) => {
  return fetch(`${BASE}/api/public/reports/${id}/data`)
    .then(r => r.ok ? r.json() : Promise.reject(new Error('HTTP ' + r.status)));
};

// ── OAuth connection management ─────────────────────────────────────────────────
export const getOAuthStatus = () => api('/auth/status');

export const startMicrosoftAuth = () => api('/auth/microsoft/start');
export const startGoogleAuth = () => api('/auth/google/start');

export const disconnectOAuth = (provider) =>
  api(`/auth/${provider}`, { method: 'DELETE' });

// ── Report access management ───────────────────────────────────────────────────
export const getReportAccess = (id) => api(`/api/reports/${id}/access`);
export const setReportAccess = (id, userIds) =>
  api(`/api/reports/${id}/access`, { method: 'PUT', body: JSON.stringify({ userIds }) });

// ── Custom app credentials management ─────────────────────────────────────────
export const getCustomCredentials = () => api('/api/custom-credentials');

export const saveCustomCredentials = (provider, clientId, clientSecret, tenantId) =>
  api(`/api/custom-credentials/${provider}`, {
    method: 'PUT',
    body: JSON.stringify({ clientId, clientSecret, tenantId }),
  });

export const testCustomCredentials = (provider, clientId, clientSecret, tenantId) =>
  api(`/api/custom-credentials/${provider}/test`, {
    method: 'POST',
    body: JSON.stringify({ clientId, clientSecret, tenantId }),
  });

export const deleteCustomCredentials = (provider) =>
  api(`/api/custom-credentials/${provider}`, { method: 'DELETE' });

// ── Phase 3: Collaborative Workflow ───────────────────────────────────────────

// Lightweight field-name list for collab setup dropdowns (no rows loaded)
export const getReportFields = (id) => api(`/api/reports/${id}/fields`);

export const toggleCollab = (id, enabled) =>
  api(`/api/reports/${id}/collab-toggle`, { method: 'PATCH', body: JSON.stringify({ enabled }) });

export const getCollabColumns = (id) => api(`/api/reports/${id}/collab-columns`);
export const createCollabColumn = (id, payload) =>
  api(`/api/reports/${id}/collab-columns`, { method: 'POST', body: JSON.stringify(payload) });
export const updateCollabColumn = (id, colId, payload) =>
  api(`/api/reports/${id}/collab-columns/${colId}`, { method: 'PUT', body: JSON.stringify(payload) });
export const deleteCollabColumn = (id, colId) =>
  api(`/api/reports/${id}/collab-columns/${colId}`, { method: 'DELETE' });

export const getCollabCycles = (id) => api(`/api/reports/${id}/collab-cycles`);
export const openCollabCycle = (id, periodLabel, historyViewerIds) =>
  api(`/api/reports/${id}/collab-cycles`, { method: 'POST', body: JSON.stringify({ period_label: periodLabel, history_viewer_ids: historyViewerIds }) });
export const closeCollabCycle = (id, cycleId) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/close`, { method: 'PATCH' });

export const renameCollabCycle = (id, cycleId, label) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/rename`, { method: 'PATCH', body: JSON.stringify({ label }) });

export const deleteCollabCycle = (id, cycleId) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}`, { method: 'DELETE' });

export const reopenCollabCycle = (id, cycleId) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/reopen`, { method: 'PATCH' });

export const exportCollabCycle = (id, cycleId) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/export`);

export const getCollabValues = (id, cycleId) => api(`/api/reports/${id}/collab-cycles/${cycleId}/values`);
export const upsertCollabValue = (id, cycleId, payload) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/values`, { method: 'PUT', body: JSON.stringify(payload) });
export const submitCollabValue = (id, cycleId, row_key, col_id) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/values/submit`, { method: 'PATCH', body: JSON.stringify({ row_key, col_id }) });
export const reviewCollabValue = (id, cycleId, row_key, col_id, action, remarks, modified_value) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/values/review`, { method: 'PATCH', body: JSON.stringify({ row_key, col_id, action, remarks, modified_value }) });

export const getCollabAudit = (id, cycleId, row_key) =>
  api(`/api/reports/${id}/collab-cycles/${cycleId}/audit${row_key ? '?row_key=' + encodeURIComponent(row_key) : ''}`);
export const getCollabHistory = (id) => api(`/api/reports/${id}/collab-history`);
