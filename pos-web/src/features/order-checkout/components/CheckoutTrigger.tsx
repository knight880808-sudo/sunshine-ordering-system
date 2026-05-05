type Props = {
  totalCents: number
  disabled: boolean
  onOpen: () => void
}

export function CheckoutTrigger({ totalCents, disabled, onOpen }: Props) {
  return (
    <section className="rounded-2xl border border-zinc-800/90 bg-zinc-900/50 p-5 shadow-lg shadow-black/25 ring-1 ring-white/[0.04] backdrop-blur-md">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <div className="text-xs font-medium uppercase tracking-wider text-zinc-500">
            待结账
          </div>
          <div className="mt-1 text-3xl font-semibold tabular-nums tracking-tight text-zinc-50">
            ¥{(totalCents / 100).toFixed(2)}
          </div>
        </div>
        <button
          type="button"
          disabled={disabled}
          onClick={onOpen}
          className="rounded-xl bg-emerald-500 px-8 py-3.5 text-base font-semibold text-zinc-950 shadow-lg shadow-emerald-500/25 transition-all duration-150 ease-out hover:bg-emerald-400 hover:shadow-emerald-400/30 active:scale-[0.96] active:brightness-95 disabled:cursor-not-allowed disabled:bg-zinc-800 disabled:text-zinc-500 disabled:shadow-none"
        >
          结账
        </button>
      </div>
      <p className="mt-4 text-xs leading-relaxed text-zinc-600">
        结账后选择支付方式；确认后将模拟打印小票并清空购物车。
      </p>
    </section>
  )
}
