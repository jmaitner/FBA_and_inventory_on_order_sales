# Distribution Solutions — Reorder & Sales Velocity Dashboard

A self-contained web app that turns two reports into actionable replenishment insights:

1. **Apprise ERP "Inventory On Order and Sales"** (`.xlsx`) → **PO / reorder** (supplier → warehouse)
2. **Amazon FBA restock / inventory health** (`.csv`) → **ship-in to FBA** (warehouse → Amazon)

It runs entirely in your browser — no install, no server, no internet, and your data
never leaves your machine.

---

## Access / password

The site opens to a password gate. Current password: **`Plainwell1!`** (asked once per
browser session). To change it, recompute the hash and update both spots in `index.html`:
run the snippet below with your new password and paste the result into the `EXPECT` value
in the inline gate script.

```
node -e 'const s="YOUR_NEW_PASSWORD";let x=5381;for(let i=0;i<s.length;i++){x=(((x<<5)+x)^s.charCodeAt(i))>>>0;}console.log(x.toString(16));'
```

> ⚠️ This is a **client-side** gate — fine for keeping casual viewers out, but anyone who
> reads the page source can bypass it. For real security when you deploy to Cloudflare, put
> **Cloudflare Access** in front of the site (proper authentication, no code changes needed).

## How to run it (local)

**Easiest:** double-click `index.html`. It opens in your browser and works offline.

Click **Upload report** (or drag a file onto the page). Upload either report first, then
add the other anytime — they're combined automatically. The FBA file is matched to the ERP
by stripping the leading `FBA` from each Amazon SKU (e.g. `FBA5530619` → `5530619`).

### Focus list (cuts the ERP noise)

The ERP report has ~1,800 SKUs, most of which are noise. The dashboard shows only your
**focus list** in the Reorder/PO view: a SKU is included if it's in the move-forward list,
its code is **≥ 5,530,847** (your new-items block and anything newer), **or** it's an FBA
item (FBA is always kept). Everything else is hidden — tick **"Show all SKUs"** to reveal it.

The list is bundled in `whitelist.js` (move-forward codes + the `NEW_ITEM_MIN` threshold).
To refresh it, just **upload an updated `move forward items.xlsx`** (the one with the
`Continue Relationship Y/N` column) — it replaces the move-forward set for that session.
FBA is never filtered.

### Two views (top toggle, appears once both files are loaded)
- **📦 Reorder / PO** — what to order from suppliers into your warehouse.
- **🚚 FBA restock** — what to ship from your warehouse into Amazon FBA. The **FBA Action**
  badge ties the two together:
  - **FBA out** — out of stock at Amazon (losing sales)
  - **Ship to FBA** — needs a shipment and your warehouse can cover it
  - **Ship + reorder** — needs a shipment but your warehouse *can't* cover it → also cut a PO
  - **FBA excess** — too much sitting at Amazon (consider a removal/sale)

Each view shows only the columns it needs, so the table fits your screen with no horizontal
scrolling. Click any row to expand it; matched rows show both warehouse history and Amazon
FBA detail (available, inbound, FBA velocity, days of supply, Amazon's recommended ship-in
qty/date, price). SKUs listed twice on Amazon (two ASINs) are aggregated automatically.

> If your browser blocks the local `xlsx` library when opened directly, run a tiny
> local server instead: open a terminal in this folder and run
> `python -m http.server 8770`, then visit `http://localhost:8770`.

---

## What it shows

- **KPI summary** — active SKUs, units sold (12 mo), stockouts with demand, SKUs to
  reorder, total units to order, over/dead stock.
- **Insight cards** — most urgent to reorder, fastest growing, slowing down, dead stock.
- **Planning assumptions** (live) — change any value and every projection and reorder
  quantity recalculates instantly:
  - **Velocity window** — last 3 / 6 / 12 months
  - **Weighting** — even average, or weight recent months higher
  - **Growth adjustment %** — bump all projections up/down for seasonality or expected demand
  - **Lead time, Coverage target, Safety stock** (days)
  - **Round order up to** — case/carton pack multiple
- **Main table** — sortable on any column, filter by supplier / category / status, free-text
  search, group by supplier, and "only items needing action". Each row shows a 12-month
  sales sparkline and trend arrow.
- **Editable Order Qty** — auto-filled with the suggested quantity; override any cell and the
  override (highlighted amber) is what gets exported.

## Exports

- **Export reorder / PO** → `Reorder_PO_<date>.xlsx` — items with an order qty > 0,
  grouped/sorted by supplier, plus a **Supplier Summary** sheet. Flags which items are also
  sold on FBA. Use this to cut POs.
- **Export FBA ship-in** → `Ship_to_FBA_<date>.xlsx` — items with a ship qty > 0: FBA
  available, inbound, days of supply, health, warehouse available, Amazon's recommendation,
  your ship qty, and a **Warehouse Short?** flag. Use this to build FBA shipping plans.
- **Export full analysis** → `Inventory_Analysis_<date>.xlsx` — every SKU with all computed
  metrics, the full monthly sales history, and (when loaded) the FBA columns. Auto-filter on.

---

## How the numbers are calculated

- **Velocity / month** = average monthly sales over the selected window, using only
  *complete* months. The current month is partial (mid-month) so it's excluded from velocity
  but still shown in the sparkline. Negative months (returns) are floored at 0 for demand.
- **Projections** = velocity/day × 30/60/90, × (1 + growth%).
- **Inventory position** = Available + On Order + In Transit.
- **Suggested order** = `velocity/day × (lead time + coverage + safety) − inventory position`,
  rounded up to the case pack, floored at 0. Discontinued and non-selling items are never
  auto-ordered.
- **Status:** Stockout (selling, none available) · Critical (runs out before a reorder lands) ·
  Reorder soon · Healthy · Overstock (>365 days cover) · Dead stock (inventory, no sales) ·
  No sales · Discontinued.
- **Urgency** (default sort) puts active stockouts on top, then items most likely to run out
  before replenishment, weighted by sales volume.

---

## Customizing the brand

- **Logo:** drop a file named `logo.png` in this folder — it appears in the header automatically.
- **Colors:** edit the hex values in the `:root` block at the top of `styles.css`.

---

## Notes & limitations

- The source report has **no cost or selling price**, so the app works in **units** (velocity,
  coverage, reorder qty). Feed the exported reorder quantities into your existing margin
  spreadsheet for dollar figures.
- Column and month detection is dynamic — when you upload next month's report, the months
  shift automatically. It locates columns by their header names ("Product Code", "On Hand",
  "Available", etc.) and any `M/YYYY` columns, so it keeps working as periods roll forward.

## Deploying to Cloudflare later

This is a plain static site. To publish to Cloudflare Pages, upload this whole folder
(`index.html`, `styles.css`, `app.js`, `vendor/`) as a Pages project — no build step needed.
