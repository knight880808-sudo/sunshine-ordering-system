import type { SalesSummary } from '../hooks/useSalesStats'

type Props = {
  summary: SalesSummary
  loading?: boolean
  error?: string | null
}

export function SalesStatsPanel({ summary, loading, error }: Props) {
  return (
    <section className="rounded-2xl border border-zinc-800/90 bg-zinc-900/40 p-5 shadow-lg shadow-black/20 ring-1 ring-white/[0.04] backdrop-blur-sm">
      <header className="mb-4 border-b border-zinc-800/90 pb-3">
        <h2 className="text-base font-semibold tracking-tight text-zinc-100">销售统计</h2>
        <p className="mt-1 text-sm text-zinc-500">来自数据库 pos_orders（本会话结账累加）</p>
      </header>

      {error ? (
        <p className="text-sm text-rose-400/90">{error}</p>
      ) : loading ? (
        <div className="flex items-center gap-3 py-8 text-sm text-zinc-500">
          <span
            className="inline-block h-5 w-5 animate-spin rounded-full border-2 border-zinc-600 border-t-emerald-500"
            aria-hidden
          />
          加载统计…
        </div>
      ) : (
        <dl className="grid grid-cols-1 gap-3 sm:grid-cols-3">
          <div className="rounded-xl border border-zinc-800/60 bg-zinc-950/50 px-4 py-3 shadow-md shadow-black/20">
            <dt className="text-[11px] font-medium uppercase tracking-wider text-zinc-500">
              订单数
            </dt>
            <dd className="mt-1 text-2xl font-semibold tabular-nums text-zinc-100">
              {summary.orderCount}
            </dd>
          </div>
          <div className="rounded-xl border border-zinc-800/60 bg-zinc-950/50 px-4 py-3 shadow-md shadow-black/20">
            <dt className="text-[11px] font-medium uppercase tracking-wider text-zinc-500">
              销售额
            </dt>
            <dd className="mt-1 text-2xl font-semibold tabular-nums text-emerald-400">
              ¥{(summary.revenueCents / 100).toFixed(2)}
            </dd>
          </div>
          <div className="rounded-xl border border-zinc-800/60 bg-zinc-950/50 px-4 py-3 shadow-md shadow-black/20">
            <dt className="text-[11px] font-medium uppercase tracking-wider text-zinc-500">
              客单价
            </dt>
            <dd className="mt-1 text-2xl font-semibold tabular-nums text-zinc-100">
              ¥{(summary.averageTicketCents / 100).toFixed(2)}
            </dd>
          </div>
        </dl>
      )}
    </section>
  )
}
