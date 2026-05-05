import type { MoneyCents } from '@/shared/types'
import type { CartLine } from '@/features/cart/types'

/** 现金 / 扫码 / 信用卡 */
export type PaymentMethod = 'cash' | 'scan' | 'credit'

export type Order = {
  id: string
  createdAt: string
  lines: CartLine[]
  subtotalCents: MoneyCents
  discountCents: MoneyCents
  totalCents: MoneyCents
  discountCode?: string | null
  manualPriceOverride: boolean
  payment: PaymentMethod
}
