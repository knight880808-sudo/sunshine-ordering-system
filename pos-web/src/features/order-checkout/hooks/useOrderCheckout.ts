import { useCallback, useMemo } from 'react'
import type { CartLine } from '@/features/cart/types'
import { postJson } from '@/shared/lib/api'
import { simulateReceiptPrint } from '../lib/simulateReceiptPrint'
import type { Order, PaymentMethod } from '../types'

type CheckoutResponse = {
  orderId: string
  createdAt: string
}

type Params = {
  lines: CartLine[]
  subtotalCents: number
  discountCents: number
  totalCents: number
  appliedDiscountCode: string | null
  manualPriceOverride: boolean
  clearCart: () => void
  onAfterCheckout?: () => void | Promise<void>
}

export function useOrderCheckout({
  lines,
  subtotalCents,
  discountCents,
  totalCents,
  appliedDiscountCode,
  manualPriceOverride,
  clearCart,
  onAfterCheckout,
}: Params) {
  const completeOrder = useCallback(
    async (payment: PaymentMethod) => {
      if (lines.length === 0) return
      const snapshotLines = lines.map((l) => ({ ...l }))
      const body = {
        lines: snapshotLines.map((l) => ({
          id: l.id,
          productId: l.productId,
          name: l.name,
          unitPriceCents: l.unitPriceCents,
          quantity: l.quantity,
          maxQty: l.maxQty,
        })),
        subtotalCents,
        discountCents,
        totalCents,
        discountCode: appliedDiscountCode,
        manualPriceOverride,
        payment,
      }
      const res = await postJson<CheckoutResponse>('/api/pos/checkout', body)
      const order: Order = {
        id: res.orderId,
        createdAt: new Date(res.createdAt).toISOString(),
        lines: snapshotLines,
        subtotalCents,
        discountCents,
        totalCents,
        discountCode: appliedDiscountCode,
        manualPriceOverride,
        payment,
      }
      simulateReceiptPrint(order)
      clearCart()
      await onAfterCheckout?.()
    },
    [
      lines,
      subtotalCents,
      discountCents,
      totalCents,
      appliedDiscountCode,
      manualPriceOverride,
      clearCart,
      onAfterCheckout,
    ],
  )

  return useMemo(
    () => ({
      completeOrder,
    }),
    [completeOrder],
  )
}
