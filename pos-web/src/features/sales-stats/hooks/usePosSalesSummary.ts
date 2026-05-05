import { useCallback, useEffect, useState } from 'react'
import { getJson } from '@/shared/lib/api'
import type { SalesSummary } from './useSalesStats'

export function usePosSalesSummary() {
  const [summary, setSummary] = useState<SalesSummary>({
    orderCount: 0,
    revenueCents: 0,
    averageTicketCents: 0,
  })
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  const refresh = useCallback(async () => {
    setError(null)
    try {
      const data = await getJson<SalesSummary>('/api/pos/stats/summary')
      setSummary(data)
    } catch (e) {
      setError(e instanceof Error ? e.message : '统计加载失败')
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

  return { summary, loading, error, refresh }
}
