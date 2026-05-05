import type { PaymentMethod } from './types'

const LABELS: Record<PaymentMethod, string> = {
  cash: '现金',
  scan: '扫码',
  credit: '信用卡',
}

export function paymentLabel(method: PaymentMethod): string {
  return LABELS[method] ?? method
}
