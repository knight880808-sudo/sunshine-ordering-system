import { useCallback, useEffect, useMemo, useState } from 'react'
import type { Product } from '@/features/product-management/types'
import { resolveDiscount } from '../lib/discountCodes'
import { inferLineSeq, loadPersistedCart, savePersistedCart } from '../lib/cartStorage'
import type { CartLine } from '../types'

let lineSeq = 0

const BOOT = (() => {
  const p = loadPersistedCart()
  if (!p) {
    return {
      lines: [] as CartLine[],
      appliedDiscountCode: null as string | null,
      manualTotalCents: null as number | null,
    }
  }
  lineSeq = inferLineSeq(p.lines, p.lineSeq)
  return {
    lines: p.lines,
    appliedDiscountCode: p.appliedDiscountCode,
    manualTotalCents: p.manualTotalCents,
  }
})()

function nextLineId() {
  lineSeq += 1
  return `line-${lineSeq}`
}

function clampMoney(cents: number, maxCents: number) {
  if (!Number.isFinite(cents)) return 0
  const rounded = Math.round(cents)
  return Math.max(0, Math.min(rounded, maxCents))
}

export function useCart() {
  const [lines, setLines] = useState<CartLine[]>(() => BOOT.lines)
  const [appliedDiscountCode, setAppliedDiscountCode] = useState<string | null>(
    () => BOOT.appliedDiscountCode,
  )
  const [manualTotalCents, setManualTotalCents] = useState<number | null>(
    () => BOOT.manualTotalCents,
  )

  useEffect(() => {
    savePersistedCart({
      v: 1,
      lines,
      appliedDiscountCode,
      manualTotalCents,
      lineSeq,
    })
  }, [lines, appliedDiscountCode, manualTotalCents])

  const subtotalCents = useMemo(
    () => lines.reduce((sum, l) => sum + l.unitPriceCents * l.quantity, 0),
    [lines],
  )

  const discountCents = useMemo(() => {
    if (!appliedDiscountCode) return 0
    const r = resolveDiscount(subtotalCents, appliedDiscountCode)
    return r.ok ? r.discountCents : 0
  }, [subtotalCents, appliedDiscountCode])

  const computedTotalCents = useMemo(
    () => Math.max(0, subtotalCents - discountCents),
    [subtotalCents, discountCents],
  )

  const totalCents = useMemo(() => {
    if (manualTotalCents === null) return computedTotalCents
    return clampMoney(manualTotalCents, subtotalCents)
  }, [manualTotalCents, computedTotalCents, subtotalCents])

  const applyDiscountCode = useCallback((rawCode: string) => {
    const r = resolveDiscount(subtotalCents, rawCode)
    if (!r.ok) return { ok: false as const, message: r.message }
    setAppliedDiscountCode(rawCode.trim().toUpperCase())
    return { ok: true as const }
  }, [subtotalCents])

  const clearDiscountCode = useCallback(() => {
    setAppliedDiscountCode(null)
  }, [])

  const setManualTotal = useCallback((cents: number | null) => {
    if (cents === null) {
      setManualTotalCents(null)
      return
    }
    if (!Number.isFinite(cents)) {
      setManualTotalCents(null)
      return
    }
    setManualTotalCents(Math.round(cents))
  }, [])

  const clearManualTotal = useCallback(() => {
    setManualTotalCents(null)
  }, [])

  const addProduct = useCallback((product: Product) => {
    setLines((prev) => {
      const existing = prev.find((l) => l.productId === product.id)
      const cap = product.stock
      if (existing) {
        if (existing.quantity >= cap) return prev
        return prev.map((l) =>
          l.productId === product.id
            ? { ...l, quantity: Math.min(cap, l.quantity + 1), maxQty: cap }
            : l,
        )
      }
      if (cap < 1) return prev
      return [
        ...prev,
        {
          id: nextLineId(),
          productId: product.id,
          name: product.name,
          unitPriceCents: product.priceCents,
          quantity: 1,
          maxQty: cap,
        },
      ]
    })
  }, [])

  const setQuantity = useCallback((lineId: string, quantity: number) => {
    const q = Math.max(0, Math.floor(quantity))
    setLines((prev) => {
      const line = prev.find((l) => l.id === lineId)
      if (!line) return prev
      const capped = Math.min(q, line.maxQty)
      if (capped === 0) return prev.filter((l) => l.id !== lineId)
      return prev.map((l) => (l.id === lineId ? { ...l, quantity: capped } : l))
    })
  }, [])

  const incrementQuantity = useCallback((lineId: string) => {
    setLines((prev) => {
      const line = prev.find((l) => l.id === lineId)
      if (!line || line.quantity >= line.maxQty) return prev
      return prev.map((l) =>
        l.id === lineId ? { ...l, quantity: l.quantity + 1 } : l,
      )
    })
  }, [])

  const decrementQuantity = useCallback((lineId: string) => {
    setLines((prev) => {
      const line = prev.find((l) => l.id === lineId)
      if (!line) return prev
      if (line.quantity <= 1) return prev.filter((l) => l.id !== lineId)
      return prev.map((l) =>
        l.id === lineId ? { ...l, quantity: l.quantity - 1 } : l,
      )
    })
  }, [])

  const removeLine = useCallback((lineId: string) => {
    setLines((prev) => prev.filter((l) => l.id !== lineId))
  }, [])

  const clear = useCallback(() => {
    setLines([])
    setAppliedDiscountCode(null)
    setManualTotalCents(null)
  }, [])

  return useMemo(
    () => ({
      lines,
      addProduct,
      setQuantity,
      incrementQuantity,
      decrementQuantity,
      removeLine,
      clear,
      subtotalCents,
      discountCents,
      computedTotalCents,
      totalCents,
      appliedDiscountCode,
      applyDiscountCode,
      clearDiscountCode,
      manualTotalCents,
      setManualTotal,
      clearManualTotal,
      isEmpty: lines.length === 0,
    }),
    [
      lines,
      addProduct,
      setQuantity,
      incrementQuantity,
      decrementQuantity,
      removeLine,
      clear,
      subtotalCents,
      discountCents,
      computedTotalCents,
      totalCents,
      appliedDiscountCode,
      applyDiscountCode,
      clearDiscountCode,
      manualTotalCents,
      setManualTotal,
      clearManualTotal,
    ],
  )
}
