import type { CartLine } from '../types'

export const CART_STORAGE_KEY = 'pos-cart-state-v1'

export type PersistedCartState = {
  v: 1
  lines: CartLine[]
  appliedDiscountCode: string | null
  manualTotalCents: number | null
  lineSeq: number
}

function isCartLine(x: unknown): x is CartLine {
  if (!x || typeof x !== 'object') return false
  const o = x as Record<string, unknown>
  return (
    typeof o.id === 'string' &&
    typeof o.productId === 'string' &&
    typeof o.name === 'string' &&
    typeof o.unitPriceCents === 'number' &&
    typeof o.quantity === 'number' &&
    typeof o.maxQty === 'number'
  )
}

export function maxLineIndex(lines: CartLine[]): number {
  let max = 0
  for (const line of lines) {
    const m = /^line-(\d+)$/.exec(line.id)
    if (m) max = Math.max(max, Number(m[1]))
  }
  return max
}

export function inferLineSeq(lines: CartLine[], storedSeq: number): number {
  return Math.max(Number.isFinite(storedSeq) ? storedSeq : 0, maxLineIndex(lines))
}

export function loadPersistedCart(): PersistedCartState | null {
  if (typeof localStorage === 'undefined') return null
  try {
    const raw = localStorage.getItem(CART_STORAGE_KEY)
    if (!raw) return null
    const data = JSON.parse(raw) as unknown
    if (!data || typeof data !== 'object') return null
    const o = data as Record<string, unknown>
    if (o.v !== 1 || !Array.isArray(o.lines)) return null
    const lines = o.lines.filter(isCartLine)
    const storedSeq =
      typeof o.lineSeq === 'number' && Number.isFinite(o.lineSeq)
        ? Math.max(0, Math.floor(o.lineSeq))
        : 0
    const lineSeq = inferLineSeq(lines, storedSeq)
    return {
      v: 1,
      lines,
      appliedDiscountCode:
        typeof o.appliedDiscountCode === 'string' ? o.appliedDiscountCode : null,
      manualTotalCents:
        typeof o.manualTotalCents === 'number' && Number.isFinite(o.manualTotalCents)
          ? Math.max(0, Math.round(o.manualTotalCents))
          : null,
      lineSeq,
    }
  } catch {
    return null
  }
}

export function savePersistedCart(state: PersistedCartState): void {
  if (typeof localStorage === 'undefined') return
  try {
    localStorage.setItem(CART_STORAGE_KEY, JSON.stringify(state))
  } catch {
    /* 存储已满或隐私模式 */
  }
}
