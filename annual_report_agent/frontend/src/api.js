// Empty base = same origin (served by the backend / Vite dev proxy). Set VITE_API_URL for Firebase Hosting.
const BASE = (import.meta.env.VITE_API_URL || '').replace(/\/$/, '')

async function request(path, options) {
  const res = await fetch(`${BASE}${path}`, options)
  if (!res.ok) {
    let message = `Request failed (${res.status})`
    try {
      const body = await res.json()
      message = body.detail || message
    } catch { /* not JSON */ }
    throw new Error(message)
  }
  return res
}

export async function startJob(form) {
  const res = await request('/api/jobs', { method: 'POST', body: form })
  return res.json()
}

export async function getJob(id) {
  const res = await request(`/api/jobs/${id}`)
  return res.json()
}

export async function getDefaultSpec() {
  const res = await request('/api/default-spec')
  return res.text()
}

export const downloadUrl = (id) => `${BASE}/api/jobs/${id}/download`
export const jsonUrl = (id) => `${BASE}/api/jobs/${id}/json`
