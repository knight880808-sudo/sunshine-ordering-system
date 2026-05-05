import { useRef, useState } from 'react'
import type { CartLine } from '../types'

type DiscountResult = { ok: true } | { ok: false; message: string }

const btnBase =
  'rounded-xl font-medium transition-all duration-150 ease-out active:scale-[0.96] disabled:pointer-events-none disabled:opacity-40'

type Props = {
  lines: CartLine[]
  subtotalCents: number
  discountCents: number
  computedTotalCents: number
  totalCents: number
  appliedDiscountCode: string | null
  manualTotalCents: number | null
  onIncrement: (lineId: string) => void
  onDecrement: (lineId: string) => void
  onChangeQty: (lineId: string, quantity: number) => void
  onRemove: (lineId: string) => void
  onApplyDiscount: (code: string) => DiscountResult
  onClearDiscount: () => void
  onApplyManualTotalCents: (cents: number | null) => void
  onClearManualTotal: () => void
}

export function CartPanel({
  lines,
  subtotalCents,
  discountCents,
  computedTotalCents,
  totalCents,
  appliedDiscountCode,
  manualTotalCents,
  onIncrement,
  onDecrement,
  onChangeQty,
  onRemove,
  onApplyDiscount,
  onClearDiscount,
  onApplyManualTotalCents,
  onClearManualTotal,
}: Props) {
  const discountInputRef = useRef<HTMLInputElement>(null)
  const manualInputRef = useRef<HTMLInputElement>(null)
  const [discountError, setDiscountError] = useState<string | null>(null)

  const handleApplyDiscount = () => {
    const raw = discountInputRef.current?.value ?? ''
    const r = onApplyDiscount(raw)
    if (r.ok) {
      setDiscountError(null)
      return
    }
    setDiscountError(r.message)
  }

  const handleApplyManual = () => {
    const trimmed = manualInputRef.current?.value.trim() ?? ''
    if (trimmed === '') {
      onApplyManualTotalCents(null)
      return
    }
    const yuan = Number.parseFloat(trimmed)
    if (!Number.isFinite(yuan) || yuan < 0) {
      return
    }
    const cents = Math.round(yuan * 100)
    onApplyManualTotalCents(Math.min(cents, subtotalCents))
  }

  const manualActive = manualTotalCents !== null

  const inputCls =
    'rounded-xl border border-zinc-700/90 bg-zinc-950/70 px-3 py-2 text-sm text-zinc-100 shadow-inner shadow-black/30 outline-none ring-emerald-500/0 transition placeholder:text-zinc-600 focus:border-emerald-500/45 focus:ring-2 focus:ring-emerald-500/20'

  return (
    <section className="flex min-h-0 flex-col rounded-2xl border border-zinc-800/90 bg-zinc-900/50 p-5 shadow-lg shadow-black/25 ring-1 ring-white/[0.04] backdrop-blur-md">
      <header className="mb-4 shrink-0 border-b border-zinc-800/90 pb-3">
        <h2 className="text-base font-semibold tracking-tight text-zinc-100">购物车</h2>
        <p className="mt-1 text-sm text-zinc-500">本地持久化 · 调整数量与优惠</p>
      </header>

      {lines.length === 0 ? (
        <p className="text-sm text-zinc-500">购物车为空，请添加商品。</p>
      ) : (
        <ul className="min-h-0 flex-1 space-y-2 overflow-y-auto overscroll-contain">
          {lines.map((line) => (
            <li
              key={line.id}
              className="rounded-xl border border-zinc-800/80 bg-zinc-950/40 px-3 py-3 shadow-sm shadow-black/20"
            >
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <div className="truncate font-medium text-zinc-100">{line.name}</div>
                  <div className="text-xs text-zinc-500">
                    单价 ¥{(line.unitPriceCents / 100).toFixed(2)}
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => onRemove(line.id)}
                  className="shrink-0 rounded-lg px-1.5 py-0.5 text-xs text-rose-400/90 underline-offset-2 transition hover:bg-rose-500/10 hover:text-rose-300 hover:underline active:scale-95"
                >
                  移除
                </button>
              </div>
              <div className="mt-3 flex items-center justify-between gap-2">
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    aria-label="减少数量"
                    onClick={() => onDecrement(line.id)}
                    className={`flex h-9 w-9 items-center justify-center rounded-lg border border-zinc-600/90 bg-zinc-900 text-lg leading-none text-zinc-200 shadow-sm hover:bg-zinc-800 hover:text-white active:scale-90 ${btnBase}`}
                  >
                    −
                  </button>
                  <input
                    type="number"
                    min={1}
                    max={line.maxQty}
                    className={`${inputCls} w-14 text-center tabular-nums`}
                    value={line.quantity}
                    onChange={(e) => onChangeQty(line.id, Number(e.target.value))}
                  />
                  <button
                    type="button"
                    aria-label="增加数量"
                    disabled={line.quantity >= line.maxQty}
                    onClick={() => onIncrement(line.id)}
                    className={`flex h-9 w-9 items-center justify-center rounded-lg border border-zinc-600/90 bg-zinc-900 text-lg leading-none text-zinc-200 shadow-sm hover:bg-zinc-800 hover:text-white active:scale-90 disabled:border-zinc-800 disabled:bg-zinc-900/50 ${btnBase}`}
                  >
                    +
                  </button>
                </div>
                <span className="tabular-nums font-medium text-zinc-200">
                  ¥{((line.unitPriceCents * line.quantity) / 100).toFixed(2)}
                </span>
              </div>
            </li>
          ))}
        </ul>
      )}

      <div className="mt-4 shrink-0 space-y-4 border-t border-zinc-800/90 pt-4">
        <div>
          <div className="mb-2 text-[11px] font-medium uppercase tracking-wider text-zinc-500">
            折扣码
          </div>
          <div className="flex flex-wrap gap-2">
            <input
              ref={discountInputRef}
              key={appliedDiscountCode ?? 'discount-draft'}
              type="text"
              placeholder="输入后点击应用"
              defaultValue={appliedDiscountCode ?? ''}
              onChange={() => setDiscountError(null)}
              className={`${inputCls} min-w-[140px] flex-1`}
              autoComplete="off"
            />
            <button
              type="button"
              onClick={handleApplyDiscount}
              disabled={lines.length === 0}
              className={`${btnBase} bg-zinc-100 px-4 py-2 text-sm text-zinc-950 shadow-md shadow-black/30 hover:bg-white`}
            >
              应用
            </button>
            {appliedDiscountCode ? (
              <button
                type="button"
                onClick={() => {
                  onClearDiscount()
                  setDiscountError(null)
                }}
                className={`${btnBase} border border-zinc-700 bg-transparent px-4 py-2 text-sm text-zinc-300 hover:border-zinc-600 hover:bg-zinc-800/80`}
              >
                清除折扣
              </button>
            ) : null}
          </div>
          {discountError ? (
            <p className="mt-2 text-xs text-rose-400">{discountError}</p>
          ) : (
            <p className="mt-2 text-xs text-zinc-600">
              演示：SAVE10 · VIP88 · MINUS500
            </p>
          )}
        </div>

        <div>
          <div className="mb-2 text-[11px] font-medium uppercase tracking-wider text-zinc-500">
            手动应收
          </div>
          <div className="flex flex-wrap gap-2">
            <div className="relative min-w-[120px] flex-1">
              <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-zinc-500">
                ¥
              </span>
              <input
                ref={manualInputRef}
                key={manualTotalCents === null ? 'manual-off' : `manual-${manualTotalCents}`}
                type="text"
                inputMode="decimal"
                placeholder="留空则自动合计"
                defaultValue={
                  manualTotalCents === null ? '' : (manualTotalCents / 100).toFixed(2)
                }
                className={`${inputCls} w-full py-2.5 pl-8 pr-3`}
              />
            </div>
            <button
              type="button"
              onClick={handleApplyManual}
              disabled={lines.length === 0}
              className={`${btnBase} border border-emerald-500/40 bg-emerald-500/15 px-4 py-2 text-sm text-emerald-300 hover:bg-emerald-500/25`}
            >
              应用改价
            </button>
            {manualActive ? (
              <button
                type="button"
                onClick={() => {
                  onClearManualTotal()
                }}
                className={`${btnBase} border border-zinc-700 px-4 py-2 text-sm text-zinc-300 hover:bg-zinc-800/80`}
              >
                清除改价
              </button>
            ) : null}
          </div>
          <p className="mt-2 text-xs text-zinc-600">
            不超过商品小计 ¥{(subtotalCents / 100).toFixed(2)}
          </p>
        </div>
      </div>

      <footer className="mt-4 shrink-0 space-y-2 border-t border-zinc-800/90 pt-4">
        <div className="flex items-center justify-between text-sm text-zinc-400">
          <span>小计 Subtotal</span>
          <span className="tabular-nums font-medium text-zinc-200">
            ¥{(subtotalCents / 100).toFixed(2)}
          </span>
        </div>
        {discountCents > 0 ? (
          <div className="flex items-center justify-between text-sm text-emerald-400/90">
            <span>折扣</span>
            <span className="tabular-nums font-medium">
              −¥{(discountCents / 100).toFixed(2)}
            </span>
          </div>
        ) : null}
        {manualActive ? (
          <div className="flex items-center justify-between text-xs text-zinc-500">
            <span>自动应收（折后）</span>
            <span className="tabular-nums">¥{(computedTotalCents / 100).toFixed(2)}</span>
          </div>
        ) : null}
        <div className="flex items-center justify-between pt-1 text-base font-semibold text-zinc-100">
          <span>总计 Total</span>
          <span className="tabular-nums text-xl text-emerald-400">
            ¥{(totalCents / 100).toFixed(2)}
          </span>
        </div>
        {manualActive ? (
          <p className="text-xs text-amber-400/85">已手动改价，结账以总计为准。</p>
        ) : null}
      </footer>
    </section>
  )
}
