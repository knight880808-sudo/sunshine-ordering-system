import type { CartLine } from '@/features/cart/types'
import { ProductList } from './ProductList'
import type { Product } from '../types'

type Props = {
  products: Product[]
  cartLines: CartLine[]
  onAddToCart: (product: Product) => void
  loading?: boolean
  error?: string | null
}

export function ProductManagementPanel({
  products,
  cartLines,
  onAddToCart,
  loading,
  error,
}: Props) {
  return (
    <section className="flex min-h-0 flex-col rounded-2xl border border-zinc-800/90 bg-zinc-900/40 p-5 shadow-[inset_0_1px_0_0_rgba(255,255,255,0.04)] backdrop-blur-sm">
      <header className="mb-4 shrink-0 border-b border-zinc-800/80 pb-3">
        <h2 className="text-base font-semibold tracking-tight text-zinc-100">商品</h2>
        <p className="mt-1 text-sm text-zinc-500">
          数据来自 products.xlsx；默认加载前 600 条，可在下方本地搜索筛选。
        </p>
      </header>

      {loading ? (
        <div className="flex items-center gap-3 py-16 text-sm text-zinc-500">
          <span
            className="inline-block h-5 w-5 animate-spin rounded-full border-2 border-zinc-600 border-t-emerald-500"
            aria-hidden
          />
          正在加载商品…
        </div>
      ) : error ? (
        <div className="rounded-xl border border-rose-500/30 bg-rose-500/10 px-4 py-3 text-sm text-rose-300">
          <p className="font-medium">无法连接商品服务</p>
          <p className="mt-1 text-rose-300/80">{error}</p>
          <p className="mt-2 text-xs text-rose-300/60">
            请先在本机启动 API：<code className="rounded bg-black/30 px-1">run_pos_api.bat</code> 或{' '}
            <code className="rounded bg-black/30 px-1">uvicorn pos_api:app --port 5055</code>
          </p>
        </div>
      ) : (
        <ProductList products={products} cartLines={cartLines} onAddToCart={onAddToCart} />
      )}
    </section>
  )
}
