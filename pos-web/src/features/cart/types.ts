import type { MoneyCents } from '@/shared/types'
import type { ProductId } from '@/features/product-management/types'

export type CartLineId = string

export type CartLine = {
  id: CartLineId
  productId: ProductId
  name: string
  unitPriceCents: MoneyCents
  quantity: number
  /** 加入时的库存上限，用于修改数量时不超卖 */
  maxQty: number
}
