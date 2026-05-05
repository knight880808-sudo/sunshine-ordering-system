import { useCallback, useEffect, useMemo, useState } from 'react'
import { getJson } from '@/shared/lib/api'
import type { Product, ProductId } from '../types'

type ProductsResponse = {
  items: Product[]
}

export function useProductCatalog() {
  const [products, setProducts] = useState<Product[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  const refresh = useCallback(async () => {
    setLoading(true)
    setError(null)
    try {
      const data = await getJson<ProductsResponse>('/api/products')
      setProducts(data.items ?? [])
    } catch (e) {
      setError(e instanceof Error ? e.message : '商品加载失败')
      setProducts([])
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    const t = window.setTimeout(() => {
      void refresh()
    }, 0)
    return () => window.clearTimeout(t)
  }, [refresh])

  const getById = useCallback(
    (id: ProductId) => products.find((p) => p.id === id),
    [products],
  )

  const upsertProduct = useCallback((product: Product) => {
    setProducts((prev) => {
      const idx = prev.findIndex((p) => p.id === product.id)
      if (idx === -1) return [...prev, product]
      const next = [...prev]
      next[idx] = product
      return next
    })
  }, [])

  const removeProduct = useCallback((id: ProductId) => {
    setProducts((prev) => prev.filter((p) => p.id !== id))
  }, [])

  return useMemo(
    () => ({
      products,
      loading,
      error,
      refresh,
      getById,
      upsertProduct,
      removeProduct,
    }),
    [products, loading, error, refresh, getById, upsertProduct, removeProduct],
  )
}
