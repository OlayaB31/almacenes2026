import type { Store } from './types'

const API_URL = (import.meta.env.VITE_API_URL || '').replace(/\/$/, '')

function token() {
  return localStorage.getItem('wsiga_token') || ''
}

async function request<T>(path: string, options: RequestInit = {}): Promise<T> {
  const headers = new Headers(options.headers)
  headers.set('Content-Type', 'application/json')
  if (token()) headers.set('Authorization', `Bearer ${token()}`)
  const res = await fetch(`${API_URL}${path}`, { ...options, headers })
  const body = await res.json().catch(() => ({}))
  if (!res.ok) throw new Error(body.message || 'Error de servidor')
  return body as T
}

export async function login(username: string, password: string) {
  return request<{ token: string; user: { username: string } }>('/login.php', {
    method: 'POST', body: JSON.stringify({ username, password })
  })
}
export async function listStores() { return request<{ data: Store[] }>('/stores.php') }
export async function createStore(store: Store) {
  return request<{ data: Store }>('/stores.php', { method: 'POST', body: JSON.stringify(store) })
}
export async function updateStore(id: number, store: Store) {
  return request<{ data: Store }>(`/stores.php?id=${id}`, { method: 'PUT', body: JSON.stringify(store) })
}
export async function deleteStore(id: number) {
  return request<{ message: string }>(`/stores.php?id=${id}`, { method: 'DELETE' })
}
export async function importStores(stores: Store[]) {
  return request<{ inserted: number; updated: number; errors: string[] }>('/import.php', {
    method: 'POST', body: JSON.stringify({ stores })
  })
}
