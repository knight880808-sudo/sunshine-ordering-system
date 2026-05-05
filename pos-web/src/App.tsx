import { useState } from 'react'
import {
  ProductManagementPanel,
  useProductCatalog,
} from '@/features/product-management'
import { CartPanel, useCart } from '@/features/cart'
import {
  CheckoutModal,
  CheckoutTrigger,
  useOrderCheckout,
  type PaymentMethod,
} from '@/features/order-checkout'
import { SalesStatsPanel, usePosSalesSummary } from '@/features/sales-stats'

function App() {
  const [checkoutOpen, setCheckoutOpen] = useState(false)
  const catalog = useProductCatalog()
  const stats = usePosSalesSummary()
  const cart = useCart()
  const checkout = useOrderCheckout({
    lines: cart.lines,
    subtotalCents: cart.subtotalCents,
    discountCents: cart.discountCents,
    totalCents: cart.totalCents,
    appliedDiscountCode: cart.appliedDiscountCode,
    manualPriceOverride: cart.manualTotalCents !== null,
    clearCart: cart.clear,
    onAfterCheckout: stats.refresh,
  })

  const handleConfirmPay = async (payment: PaymentMethod) => {
    try {
      await checkout.completeOrder(payment)
      setCheckoutOpen(false)
    } catch (e) {
      window.alert(e instanceof Error ? e.message : '结账失败，请确认已启动 POS API (端口 5055)')
    }
  }

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 antialiased">
      <header className="sticky top-0 z-40 border-b border-zinc-800/90 bg-zinc-950/85 backdrop-blur-xl">
        <div className="mx-auto flex h-14 max-w-[1600px] items-center justify-between px-4 sm:px-6">
          <div className="flex items-baseline gap-3">
            <h1 className="text-lg font-semibold tracking-tight text-zinc-100">收银台</h1>
            <span className="hidden text-xs font-normal text-zinc-500 sm:inline">POS</span>
          </div>
          <span className="text-xs text-zinc-500">API · products.xlsx</span>
        </div>
      </header>

      <div className="relative mx-auto max-w-[1600px]">
        <main className="px-4 py-6 pb-10 sm:px-6 lg:mr-[380px] lg:pb-8">
          <div className="space-y-6">
            <ProductManagementPanel
              products={catalog.products}
              cartLines={cart.lines}
              onAddToCart={cart.addProduct}
              loading={catalog.loading}
              error={catalog.error}
            />
            <SalesStatsPanel
              summary={stats.summary}
              loading={stats.loading}
              error={stats.error}
            />
          </div>
        </main>

        <aside
          className="mx-4 mt-2 flex flex-col gap-4 sm:mx-6 lg:fixed lg:right-0 lg:top-14 lg:z-30 lg:mx-0 lg:mt-0 lg:h-[calc(100vh-3.5rem)] lg:w-[380px] lg:overflow-y-auto lg:overscroll-contain lg:border-l lg:border-zinc-800/90 lg:bg-zinc-950/75 lg:p-4 lg:shadow-[-16px_0_48px_-12px_rgba(0,0,0,0.65)] lg:backdrop-blur-2xl"
          aria-label="购物车"
        >
          <CartPanel
            lines={cart.lines}
            subtotalCents={cart.subtotalCents}
            discountCents={cart.discountCents}
            computedTotalCents={cart.computedTotalCents}
            totalCents={cart.totalCents}
            appliedDiscountCode={cart.appliedDiscountCode}
            manualTotalCents={cart.manualTotalCents}
            onIncrement={cart.incrementQuantity}
            onDecrement={cart.decrementQuantity}
            onChangeQty={cart.setQuantity}
            onRemove={cart.removeLine}
            onApplyDiscount={cart.applyDiscountCode}
            onClearDiscount={cart.clearDiscountCode}
            onApplyManualTotalCents={cart.setManualTotal}
            onClearManualTotal={cart.clearManualTotal}
          />
          <CheckoutTrigger
            totalCents={cart.totalCents}
            disabled={cart.isEmpty}
            onOpen={() => setCheckoutOpen(true)}
          />
        </aside>
      </div>

      {checkoutOpen ? (
        <CheckoutModal
          totalCents={cart.totalCents}
          onClose={() => setCheckoutOpen(false)}
          onConfirm={handleConfirmPay}
        />
      ) : null}
    </div>
  )
}

export default App
