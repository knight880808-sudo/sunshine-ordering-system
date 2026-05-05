# SUNSHINE Ordering System (Beginner Guide)

This guide is for first-time users.  
Goal: follow the steps and complete tasks without technical knowledge.

---

## 1. Know the 3 Roles

- **Branch Staff**: place orders, check my orders, confirm receipt, view shortage updates
- **Warehouse Staff**: process pending dispatch, enter dispatch quantities, handle shortages
- **Administrator**: monitor overall status, export reports, publish stock arrival notices, manage email/images/messages

---

## 2. Login and Basic Navigation

1. Open the system and choose language (EN/中文).
2. Select your role:
   - Branch Staff: choose your branch and enter
   - Warehouse/Admin: enter password to login
3. Use the left sidebar to switch pages.
4. Use top-right language toggle anytime.

---

## 3. Branch Staff Workflow

## 3.1 Place an Order

Go to `🛒 Place Order`:

1. Search by **product name / barcode / item code** (scanner supported).
2. Enter quantity on each product card (cartons / pcs).
3. Click `📥 Add Selected to Cart`.
4. Click `📝 Review & Send Order`, check items, then click `📤 Send Order`.
5. When you see “Order has been sent successfully”, it is done.

Tips:
- Order page shows 5 products per page; use pagination to continue.
- If you see a “new stock arrival” notice, prioritize those items first.

## 3.2 Check Orders and Confirm Receipt

Go to `📋 My Orders`:

1. Find orders with status `Dispatched`.
2. Expand the order and review shipped quantities.
3. Enter actual received quantities and click `✅ Confirm Receipt`.

Result:
- Normal receipt changes status to `Received`.
- If actual < dispatched, shortage records are auto-created and warehouse is notified.

## 3.3 Check Shortage Progress

Go to `🔔 My Shortages`:

- View warehouse handling status (`Resending` / `Out of Stock` / `Resolved`) and replies.

## 3.4 Check System Messages

Go to `🔔 Messages`:

- View notifications for orders, dispatch, receipt, shortages, stock arrivals.
- Use `Mark as read` / `Mark all as read`.

---

## 4. Warehouse Staff Workflow

## 4.1 Pending Dispatch

Go to `📦 Pending Dispatch`:

1. Filter by branch, date, keyword if needed.
2. Expand an order; default dispatch qty equals ordered qty.
3. Adjust quantities based on real stock.
4. Click `🚚 Mark as Dispatched`.

Result:
- Branch can confirm receipt in `My Orders`.
- System sends both email and in-app notification.

## 4.2 Handle Shortage Notifications

Go to `🔔 Shortage Notifications`:

- For each shortage record, choose:
  - `♻️ Resend` (set to Resending)
  - `❌ Mark Out of Stock`
- Add warehouse reply; branch can see it in shortages/messages.

## 4.3 Dispatch History

Go to `📜 Dispatch History`:

- Review dispatched orders by date and receipt status.

---

## 5. Administrator Workflow

## 5.1 Dashboard

Go to `📊 Dashboard` to monitor:

- Orders today
- Pending dispatch
- Dispatched
- Open shortages
- Branch status and latest records

## 5.2 Stock Arrivals (Important)

Go to `📦 Stock Arrivals`:

1. Fill arrival title and notice.
2. Enter arrival item list (one per line).
3. You can search/select products and append to list (to avoid manual typing).
4. Click `Publish arrival notice`.

Effect:
- Branch sees the notice before placing orders.
- Same notice is pushed to branch message center.
- New notice replaces the previous active one.

## 5.3 Export Reports

Go to `📥 Export Reports`:

- Export picking list, reconciliation, shortage report (Excel).

## 5.4 Email Settings

Go to `📧 Email`:

- Configure SMTP, event recipients, branch emails.
- Send test email.
- Check email activity log.

---

## 6. Notification Rules

- In-app notifications cover new order, dispatched, received, shortage, stock arrival.
- `Messages` in sidebar shows unread count.
- Check messages first every shift to avoid missing tasks.

---

## 7. FAQ (Quick)

## Q1: Cannot find product in search?

- Confirm keyword (name/barcode/item code)
- Make sure you are on `🛒 Place Order`
- Clear keyword and search again

## Q2: Branch cannot see dispatched order?

- Ask branch to refresh `📋 My Orders`
- Confirm warehouse actually clicked `🚚 Mark as Dispatched`
- Admin can verify via email log and messages

## Q3: Why shortage appears?

- During receipt, actual qty is lower than dispatched qty
- System auto-creates shortage and notifies warehouse

---

## 8. Recommended Daily Routine

- **Branch**: check `Messages` first, then place orders
- **Warehouse**: process `Pending Dispatch` first, then `Shortage Notifications`
- **Admin**: publish `Stock Arrivals` immediately when new stock comes

---

For training, you can share this file directly with new staff.

