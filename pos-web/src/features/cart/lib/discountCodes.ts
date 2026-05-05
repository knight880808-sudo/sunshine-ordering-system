export type DiscountResolve =
  | { ok: true; discountCents: number; description: string }
  | { ok: false; message: string }

/** 演示用折扣码，后续可改为后端校验 */
export function resolveDiscount(subtotalCents: number, rawCode: string): DiscountResolve {
  const code = rawCode.trim().toUpperCase()
  if (!code) {
    return { ok: false, message: '请输入折扣码' }
  }
  if (subtotalCents <= 0) {
    return { ok: false, message: '购物车为空时无法应用折扣' }
  }

  switch (code) {
    case 'SAVE10': {
      const discountCents = Math.floor((subtotalCents * 10) / 100)
      return {
        ok: true,
        discountCents: Math.min(discountCents, subtotalCents),
        description: '整单 9 折',
      }
    }
    case 'VIP88': {
      const discountCents = Math.floor((subtotalCents * 12) / 100)
      return {
        ok: true,
        discountCents: Math.min(discountCents, subtotalCents),
        description: '会员 8.8 折（减 12%）',
      }
    }
    case 'MINUS500': {
      const discountCents = Math.min(500, subtotalCents)
      return { ok: true, discountCents, description: '立减 ¥5.00' }
    }
    default:
      return { ok: false, message: '无效的折扣码' }
  }
}
