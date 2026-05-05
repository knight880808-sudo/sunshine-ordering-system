import { useMemo } from 'react'
import type { Order } from '@/features/order-checkout/types'

export type SalesSummary = {
  orderCount: number
  revenueCents: number
  averageTicketCents: number
}

export function useSalesStats(orders: Order[]): SalesSummary {
  return useMemo(() => {
    const orderCount = orders.length
    const revenueCents = orders.reduce((sum, o) => sum + o.totalCents, 0)
    const averageTicketCents =
      orderCount === 0 ? 0 : Math.round(revenueCents / orderCount)
    return { orderCount, revenueCents, averageTicketCents }
  }, [orders])
}
