import { useMemo, useState } from 'react'
import type { CartLine } from '@/features/cart/types'
import type { Product } from '../types'

function cartQtyForProduct(lines: CartLine[], productId: string): number {
  return lines.find((l) => l.productId === productId)?.quantity ?? 0
}

type Props = {
  products: Product[]
  cartLines: CartLine[]
  onAddToCart: (product: Product) => void
}

export function ProductList({ products, cartLines, onAddToCart }: Props) {
  const [query, setQuery] = useState('')

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    if (!q) return products
    return products.filter((p) => p.name.toLowerCase().includes(q))
  }, [products, query])

  return (
    <div className="flex min-h-0 flex-1 flex-col gap-4">
      <div className="shrink-0">
        <label htmlFor="product-search" className="sr-only">
          按名称搜索商品
        </label>
        <input
          id="product-search"
          type="search"
          placeholder="搜索商品名称…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          className="w-full rounded-xl border border-zinc-700/90 bg-zinc-950/80 px-4 py-2.5 text-sm text-zinc-100 shadow-inner shadow-black/20 outline-none ring-emerald-500/0 transition placeholder:text-zinc-600 focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/25"
          autoComplete="off"
        />
      </div>

      {filtered.length === 0 ? (
        <p className="py-12 text-center text-sm text-zinc-500">没有匹配的商品。</p>
      ) : (
        <ul className="grid min-h-0 flex-1 grid-cols-1 gap-4 overflow-y-auto pb-1 sm:grid-cols-2">
          {filtered.map((product) => {
            const inCart = cartQtyForProduct(cartLines, product.id)
            const remaining = Math.max(0, product.stock - inCart)
            const canAdd = remaining > 0

            return (
              <li key={product.id}>
                <button
                  type="button"
                  disabled={!canAdd}
                  onClick={() => canAdd && onAddToCart(product)}
                  className={`flex w-full gap-3 rounded-2xl border p-4 text-left transition-all duration-200 ease-out will-change-transform ${
                    canAdd
                      ? 'border-zinc-700/80 bg-zinc-900/60 shadow-md shadow-black/30 ring-1 ring-white/[0.06] hover:border-emerald-500/35 hover:bg-zinc-800/50 hover:shadow-lg hover:shadow-black/40 hover:ring-emerald-500/20 active:scale-[0.98] active:brightness-95'
                      : 'cursor-not-allowed border-zinc-800 bg-zinc-900/30 opacity-60 shadow-sm shadow-black/20'
                  }`}
                >
                  <div className="relative h-[5.25rem] w-[5.25rem] shrink-0 overflow-hidden rounded-xl bg-zinc-800 ring-1 ring-white/5">
                    <img
                      src={product.imageUrl}
                      alt={product.name}
                      className="h-full w-full object-cover"
                      loading="lazy"
                    />
                    {!canAdd ? (
                      <span className="absolute inset-0 flex items-center justify-center bg-black/60 text-xs font-medium text-zinc-200">
                        已售罄
                      </span>
                    ) : null}
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="line-clamp-2 font-medium leading-snug text-zinc-100">
                      {product.name}
                    </div>
                    <div className="mt-2 flex flex-wrap items-baseline gap-x-3 gap-y-1 text-sm">
                      <span className="text-lg font-semibold tabular-nums text-emerald-400">
                        ¥{(product.priceCents / 100).toFixed(2)}
                      </span>
                      <span className="tabular-nums text-zinc-500">
                        库存{' '}
                        <strong className="font-medium text-zinc-400">{product.stock}</strong>
                        {inCart > 0 ? (
                          <span className="text-zinc-600">（购物车 {inCart}）</span>
                        ) : null}
                      </span>
                    </div>
                    {remaining > 0 && remaining <= 5 ? (
                      <p className="mt-2 text-xs text-amber-400/90">仅剩可售 {remaining} 件</p>
                    ) : null}
                  </div>
                </button>
              </li>
            )
          })}
        </ul>
      )}
    </div>
  )
}
