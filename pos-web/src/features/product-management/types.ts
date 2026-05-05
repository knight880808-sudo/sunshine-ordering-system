import type { MoneyCents } from '@/shared/types'

export type ProductId = string

export type Product = {
  id: ProductId
  name: string
  sku: string
  priceCents: MoneyCents
  /** 商品主图 URL（可用占位图或自有 CDN） */
  imageUrl: string
  /** 当前可售库存 */
  stock: number
  category?: string
}
