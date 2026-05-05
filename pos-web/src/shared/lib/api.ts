/** 开发环境通过 Vite 代理访问同源 `/api`；生产可设置 VITE_API_URL */
export function apiUrl(path: string): string {
  const base = import.meta.env.VITE_API_URL ?? ''
  const p = path.startsWith('/') ? path : `/${path}`
  if (!base) return p
  return `${base.replace(/\/$/, '')}${p}`
}

export async function getJson<T>(path: string): Promise<T> {
  const r = await fetch(apiUrl(path))
  if (!r.ok) {
    const t = await r.text()
    throw new Error(t || `HTTP ${r.status}`)
  }
  return r.json() as Promise<T>
}

export async function postJson<T>(path: string, body: unknown): Promise<T> {
  const r = await fetch(apiUrl(path), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  if (!r.ok) {
    const t = await r.text()
    throw new Error(t || `HTTP ${r.status}`)
  }
  return r.json() as Promise<T>
}
