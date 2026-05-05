import type { Order } from '../types'
import { paymentLabel } from '../paymentLabels'

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function formatTime(iso: string): string {
  try {
    const d = new Date(iso)
    return d.toLocaleString('zh-CN', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    })
  } catch {
    return iso
  }
}

/** 生成小票 HTML 并调用浏览器打印（模拟热敏打印机） */
export function simulateReceiptPrint(order: Order): void {
  const linesHtml = order.lines
    .map((l) => {
      const name = escapeHtml(l.name)
      const lineTotal = ((l.unitPriceCents * l.quantity) / 100).toFixed(2)
      return `<div class="row"><span>${name} ×${l.quantity}</span><span>¥${lineTotal}</span></div>`
    })
    .join('')

  const discountRow =
    order.discountCents > 0
      ? `<div class="row muted"><span>折扣</span><span>−¥${(order.discountCents / 100).toFixed(2)}</span></div>`
      : ''

  const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>小票 ${escapeHtml(order.id)}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: ui-monospace, "Cascadia Mono", "Courier New", monospace;
      font-size: 12px;
      padding: 16px;
      margin: 0 auto;
      max-width: 280px;
      color: #111;
    }
    h1 { font-size: 14px; text-align: center; margin: 0 0 8px; font-weight: 700; }
    .sub { text-align: center; font-size: 11px; color: #444; margin-bottom: 12px; }
    .row {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      margin: 4px 0;
      word-break: break-all;
    }
    .row.muted { color: #555; }
    hr.dash {
      border: none;
      border-top: 1px dashed #333;
      margin: 10px 0;
    }
    .total { font-weight: 700; font-size: 14px; margin-top: 8px; }
    .foot { text-align: center; margin-top: 16px; font-size: 11px; color: #555; }
  </style>
</head>
<body>
  <h1>简易收银台</h1>
  <div class="sub">感谢光临 · 模拟小票</div>
  <div class="row"><span>订单号</span><span>${escapeHtml(order.id)}</span></div>
  <div class="row muted"><span>时间</span><span>${escapeHtml(formatTime(order.createdAt))}</span></div>
  <hr class="dash" />
  ${linesHtml}
  <hr class="dash" />
  <div class="row"><span>小计</span><span>¥${(order.subtotalCents / 100).toFixed(2)}</span></div>
  ${discountRow}
  <div class="row total"><span>应付</span><span>¥${(order.totalCents / 100).toFixed(2)}</span></div>
  <div class="row muted"><span>支付方式</span><span>${escapeHtml(paymentLabel(order.payment))}</span></div>
  ${order.manualPriceOverride ? `<div class="row muted"><span>备注</span><span>手动改价</span></div>` : ''}
  <p class="foot">—— 打印完成 ——</p>
</body>
</html>`

  const iframe = document.createElement('iframe')
  iframe.setAttribute(
    'style',
    'position:fixed;right:0;bottom:0;width:0;height:0;border:0;opacity:0;pointer-events:none',
  )
  document.body.appendChild(iframe)

  const doc = iframe.contentDocument
  const win = iframe.contentWindow
  if (!doc || !win) {
    document.body.removeChild(iframe)
    return
  }

  doc.open()
  doc.write(html)
  doc.close()

  const cleanup = () => {
    if (iframe.parentNode) iframe.parentNode.removeChild(iframe)
  }

  setTimeout(() => {
    try {
      win.focus()
      win.print()
    } finally {
      setTimeout(cleanup, 500)
    }
  }, 150)
}
