# RTLH Admin

One phone app for orders, events (quotations & invoices), stock and expenses.
Replaces the separate rtlh-inventory app.

- **Products, prices, colours, photos** come from the website: `rtlh.art/catalog-all.json`
  (edit `rtlh-site/src/data/products.js` in the site repo — admin and kiosk update by themselves).
- **Data** lives in the "RTLH" Google Sheet: Orders, Expenses, Events, Stock, Moves, Settings.
  A green row in Orders = done (same as highlighting by hand).
- **Sheet code**: `backend/Code.gs` — steps to install are at the top of that file.

Tabs: Stats · Orders (To do / Done / Events / Expenses) · Add (Order / Event / Expense) · Stock (Stock / Quick order / Restock / Incoming / History) · Settings.

Event revenue counts when an event is **Confirmed** or **Paid**. Kiosk pieces count in product stats
but not in revenue (the event's invoice is the revenue). Group orders count in revenue only.
