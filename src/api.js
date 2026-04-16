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
    throw new Error(err.error || `HTTP ${res.status}`);
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
  return data; // { token, role, username }
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

export const deleteUser = (id) =>
  api(`/api/users/${id}`, { method: 'DELETE' });

// ── Reports ──────────────────────────────────────────────────────────────────
export const getReports = () => api('/api/reports');

export const createReport = (payload) =>
  api('/api/reports', { method: 'POST', body: JSON.stringify(payload) });

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
