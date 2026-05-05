import { useEffect, useId, useRef, useState } from 'react'
import type { PaymentMethod } from '../types'
import { paymentLabel } from '../paymentLabels'

const METHODS: PaymentMethod[] = ['cash', 'scan', 'credit']

const btnGhost =
  'rounded-xl border border-zinc-600/90 bg-zinc-800/50 px-4 py-2.5 text-sm font-medium text-zinc-200 transition-all duration-150 ease-out hover:border-zinc-500 hover:bg-zinc-800 active:scale-[0.97]'
const btnPrimary =
  'rounded-xl bg-emerald-500 px-5 py-2.5 text-sm font-semibold text-zinc-950 shadow-lg shadow-emerald-500/20 transition-all duration-150 ease-out hover:bg-emerald-400 active:scale-[0.97] active:brightness-95'

type Props = {
  totalCents: number
  onClose: () => void
  onConfirm: (payment: PaymentMethod) => void
}

export function CheckoutModal({ totalCents, onClose, onConfirm }: Props) {
  const titleId = useId()
  const panelRef = useRef<HTMLDivElement>(null)
  const [payment, setPayment] = useState<PaymentMethod>('cash')

  useEffect(() => {
    const prev = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    panelRef.current?.focus()
    return () => {
      document.body.style.overflow = prev
    }
  }, [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const handleConfirm = () => {
    onConfirm(payment)
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
      <button
        type="button"
        aria-label="关闭结账窗口"
        className="absolute inset-0 bg-black/70 backdrop-blur-sm transition-opacity"
        onClick={onClose}
      />
      <div
        ref={panelRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        tabIndex={-1}
        className="relative z-10 w-full max-w-md rounded-2xl border border-zinc-700/90 bg-zinc-900 p-6 shadow-2xl shadow-black/60 ring-1 ring-white/10 outline-none backdrop-blur-xl"
      >
        <h2 id={titleId} className="text-lg font-semibold text-zinc-100">
          结账
        </h2>
        <p className="mt-1 text-sm text-zinc-500">选择支付方式并确认收款</p>

        <div className="mt-6 rounded-2xl border border-zinc-800 bg-zinc-950/80 px-4 py-5 text-center shadow-inner shadow-black/40">
          <div className="text-[11px] font-medium uppercase tracking-wider text-zinc-500">
            应付金额
          </div>
          <div className="mt-2 text-4xl font-bold tabular-nums text-emerald-400">
            ¥{(totalCents / 100).toFixed(2)}
          </div>
        </div>

        <fieldset className="mt-6">
          <legend className="mb-3 text-sm font-medium text-zinc-400">支付方式</legend>
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
            {METHODS.map((m) => (
              <label
                key={m}
                className={`flex cursor-pointer flex-col items-center rounded-xl border-2 px-3 py-3 text-center transition-all duration-150 ease-out active:scale-[0.98] ${
                  payment === m
                    ? 'border-emerald-500/70 bg-emerald-500/15 text-emerald-200 shadow-md shadow-emerald-500/10'
                    : 'border-zinc-700/90 bg-zinc-950/40 text-zinc-400 hover:border-zinc-600 hover:bg-zinc-800/50'
                }`}
              >
                <input
                  type="radio"
                  name="checkout-payment"
                  className="sr-only"
                  checked={payment === m}
                  onChange={() => setPayment(m)}
                />
                <span className="text-sm font-semibold text-zinc-100">{paymentLabel(m)}</span>
                <span className="mt-1 text-[11px] text-zinc-500">
                  {m === 'cash' && '现金收款'}
                  {m === 'scan' && '微信 / 支付宝'}
                  {m === 'credit' && '刷卡或闪付'}
                </span>
              </label>
            ))}
          </div>
        </fieldset>

        <div className="mt-8 flex justify-end gap-3 border-t border-zinc-800 pt-5">
          <button type="button" onClick={onClose} className={btnGhost}>
            取消
          </button>
          <button type="button" onClick={handleConfirm} className={btnPrimary}>
            确认支付
          </button>
        </div>
      </div>
    </div>
  )
}
