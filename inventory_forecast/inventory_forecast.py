"""
NetSuite Inventory Forecasting Tool
====================================
Inputs:
  - NetSuite "Current Inventory Snapshot" export (.xls / SpreadsheetML)
  - NetSuite "Sales by Item Summary" export (.xls / SpreadsheetML)

Outputs:
  - Console summary by urgency tier
  - inventory_forecast_report.csv  (full detail)
  - reorder_now.csv                (only items needing action ≤ REORDER_DAYS)

Usage:
  python inventory_forecast.py \
      --inventory  "FCSCurrentInventorySnapshot.xls" \
      --sales      "FCSSalesbyItemSummary.xls" \
      --sales-days 61 \
      --lead-time  30 \
      --safety-days 14
"""

import argparse
import csv
import math
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

# ── Column indices in the Inventory Snapshot (0-based) ───────────────────────
# Row 6 (header1): Item | Class | Description | Vendor | DateCreated | GA-01... | Total...
# Row 7 (header2):  \xa0 | Reorder Pt | Pref Stock | On Hand | On Order | Committed | To Order | In Transit | Curr Qty Avail | (repeat for Total)
INV_ITEM        = 0
INV_CLASS       = 1
INV_DESC        = 2
INV_VENDOR      = 3
# GA-01 section
INV_REORDER_PT  = 5
INV_PREF_STOCK  = 6
INV_ON_HAND     = 7
INV_ON_ORDER    = 8
INV_COMMITTED   = 9
INV_TO_ORDER    = 10
INV_IN_TRANSIT  = 11
INV_CURR_AVAIL  = 12
# Total section (same layout, offset by 7)
TOTAL_ON_HAND   = 15
TOTAL_ON_ORDER  = 16
TOTAL_COMMITTED = 17
TOTAL_CURR_AVAIL= 20

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}


def _cell_text(cell) -> str:
    d = cell.find("ss:Data", NS)
    return (d.text or "").strip() if d is not None else ""


def _safe_float(val: str) -> float:
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def parse_inventory(path: str) -> dict[str, dict]:
    """Return {item_id: {...}} from the Inventory Snapshot XLS."""
    tree = ET.parse(path)
    root = tree.getroot()
    ws = root.findall(".//ss:Worksheet", NS)[0]
    table = ws.find(".//ss:Table", NS)
    rows = table.findall("ss:Row", NS)

    items = {}
    # Data starts at row index 9 (0-based); skip header rows 0-8
    for row in rows[9:]:
        cells = row.findall("ss:Cell", NS)
        if not cells:
            continue
        vals = [_cell_text(c) for c in cells]
        # Pad to at least 21 columns
        while len(vals) < 21:
            vals.append("")

        item_id = vals[INV_ITEM].strip()
        if not item_id or item_id in ("Inventory Item", "Non-Inventory Item", "Kit/Package"):
            continue

        items[item_id] = {
            "item":        item_id,
            "class":       vals[INV_CLASS],
            "description": vals[INV_DESC],
            "vendor":      vals[INV_VENDOR],
            # Use Total section for multi-location accuracy
            "on_hand":     _safe_float(vals[TOTAL_ON_HAND]),
            "on_order":    _safe_float(vals[TOTAL_ON_ORDER]),
            "committed":   _safe_float(vals[TOTAL_COMMITTED]),
            "curr_avail":  _safe_float(vals[TOTAL_CURR_AVAIL]),
            # GA-01 reorder point if set
            "reorder_pt":  _safe_float(vals[INV_REORDER_PT]),
            "pref_stock":  _safe_float(vals[INV_PREF_STOCK]),
        }
    return items


def parse_sales(path: str) -> dict[str, float]:
    """Return {item_id: qty_sold} from the Sales by Item Summary XLS."""
    tree = ET.parse(path)
    root = tree.getroot()
    ws = root.findall(".//ss:Worksheet", NS)[0]
    table = ws.find(".//ss:Table", NS)
    rows = table.findall("ss:Row", NS)

    sales = {}
    # Data starts at row index 8
    for row in rows[8:]:
        cells = row.findall("ss:Cell", NS)
        if len(cells) < 3:
            continue
        vals = [_cell_text(c) for c in cells]
        item_id = vals[0].strip()
        if not item_id:
            continue
        qty = _safe_float(vals[2])
        # Accumulate in case of duplicates
        sales[item_id] = sales.get(item_id, 0.0) + qty
    return sales


@dataclass
class ForecastRow:
    item: str
    description: str
    class_name: str
    vendor: str
    on_hand: float
    on_order: float
    committed: float
    curr_avail: float
    reorder_pt: float         # from NetSuite (0 = not set)
    pref_stock: float
    qty_sold: float           # in the sales window
    daily_sales: float        # qty_sold / sales_days
    days_of_stock: float      # curr_avail / daily_sales  (inf if no sales)
    reorder_signal: float     # when stock will hit reorder_pt
    suggested_order: float    # units to bring back to pref_stock
    status: str               # OUT_OF_STOCK / REORDER_NOW / LOW / WATCH / OK / SLOW_MOVER / NO_SALES


def classify(
    row_data: dict,
    sales: dict[str, float],
    sales_days: int,
    lead_time: int,
    safety_days: int,
) -> ForecastRow:
    item = row_data["item"]
    qty_sold = sales.get(item, 0.0)
    daily = qty_sold / sales_days if qty_sold > 0 else 0.0

    curr_avail = row_data["curr_avail"]
    on_hand = row_data["on_hand"]
    reorder_pt = row_data["reorder_pt"]
    pref_stock = row_data["pref_stock"]

    # Days of stock remaining (use curr_avail = on_hand - committed + in_transit)
    if daily > 0:
        days_of_stock = curr_avail / daily
    else:
        days_of_stock = float("inf")

    # If NetSuite reorder point not set, derive from lead time + safety
    effective_reorder_pt = reorder_pt if reorder_pt > 0 else daily * (lead_time + safety_days)

    # Days until we hit the reorder point
    if daily > 0 and curr_avail > effective_reorder_pt:
        reorder_signal = (curr_avail - effective_reorder_pt) / daily
    elif daily > 0:
        reorder_signal = 0.0  # already below reorder point
    else:
        reorder_signal = float("inf")

    # Suggested order: bring stock up to preferred stock level (or 60-day supply)
    target = pref_stock if pref_stock > 0 else daily * 60
    suggested_order = max(0.0, target - (curr_avail + row_data["on_order"]))

    # Status tiers
    if on_hand <= 0 and qty_sold > 0:
        status = "OUT_OF_STOCK"
    elif days_of_stock <= lead_time and qty_sold > 0:
        status = "REORDER_NOW"
    elif days_of_stock <= lead_time + safety_days and qty_sold > 0:
        status = "LOW_STOCK"
    elif days_of_stock <= 90 and qty_sold > 0:
        status = "WATCH"
    elif qty_sold == 0 and on_hand > 0:
        status = "SLOW_MOVER"
    elif qty_sold == 0 and on_hand <= 0:
        status = "NO_SALES_NO_STOCK"
    else:
        status = "OK"

    return ForecastRow(
        item=item,
        description=row_data["description"],
        class_name=row_data["class"],
        vendor=row_data["vendor"],
        on_hand=on_hand,
        on_order=row_data["on_order"],
        committed=row_data["committed"],
        curr_avail=curr_avail,
        reorder_pt=reorder_pt,
        pref_stock=pref_stock,
        qty_sold=qty_sold,
        daily_sales=round(daily, 3),
        days_of_stock=round(days_of_stock, 1) if days_of_stock != float("inf") else 9999,
        reorder_signal=round(reorder_signal, 1) if reorder_signal != float("inf") else 9999,
        suggested_order=math.ceil(suggested_order),
        status=status,
    )


STATUS_ORDER = {
    "OUT_OF_STOCK": 0,
    "REORDER_NOW": 1,
    "LOW_STOCK": 2,
    "WATCH": 3,
    "OK": 4,
    "SLOW_MOVER": 5,
    "NO_SALES_NO_STOCK": 6,
}

STATUS_LABELS = {
    "OUT_OF_STOCK":      "🔴 OUT OF STOCK",
    "REORDER_NOW":       "🟠 REORDER NOW",
    "LOW_STOCK":         "🟡 LOW STOCK",
    "WATCH":             "🟢 WATCH",
    "OK":                "✅ OK",
    "SLOW_MOVER":        "⚪ SLOW MOVER",
    "NO_SALES_NO_STOCK": "⬛ NO SALES/NO STOCK",
}


def run(
    inventory_path: str,
    sales_path: str,
    sales_days: int = 61,
    lead_time: int = 30,
    safety_days: int = 14,
    output_full: str = "inventory_forecast_report.csv",
    output_action: str = "reorder_now.csv",
):
    print(f"\nLoading inventory:  {inventory_path}")
    inventory = parse_inventory(inventory_path)
    print(f"  → {len(inventory):,} SKUs loaded")

    print(f"Loading sales data: {sales_path}")
    sales = parse_sales(sales_path)
    print(f"  → {len(sales):,} SKUs with sales (over {sales_days} days)")

    print("\nCalculating forecasts...")
    rows = []
    for item_data in inventory.values():
        rows.append(classify(item_data, sales, sales_days, lead_time, safety_days))

    rows.sort(key=lambda r: (STATUS_ORDER.get(r.status, 9), r.days_of_stock))

    # ── Console summary ───────────────────────────────────────────────────────
    from collections import Counter
    counts = Counter(r.status for r in rows)

    print("\n" + "=" * 60)
    print("  INVENTORY FORECAST SUMMARY")
    print(f"  Lead time: {lead_time}d  |  Safety stock: {safety_days}d  |  Sales window: {sales_days}d")
    print("=" * 60)
    for status, label in STATUS_LABELS.items():
        n = counts.get(status, 0)
        if n:
            print(f"  {label:<30} {n:>5} SKUs")
    print("=" * 60)

    # Top urgent items
    urgent = [r for r in rows if r.status in ("OUT_OF_STOCK", "REORDER_NOW")]
    if urgent:
        print(f"\n  Top {min(20, len(urgent))} URGENT items (reorder immediately):")
        print(f"  {'Item':<15} {'Description':<35} {'On Hand':>8} {'Daily':>7} {'Days Left':>10} {'Suggest':>8}")
        print(f"  {'-'*15} {'-'*35} {'-'*8} {'-'*7} {'-'*10} {'-'*8}")
        for r in urgent[:20]:
            print(f"  {r.item:<15} {r.description[:35]:<35} {r.on_hand:>8.0f} {r.daily_sales:>7.2f} {r.days_of_stock:>10.0f} {r.suggested_order:>8}")

    # ── Write full CSV ────────────────────────────────────────────────────────
    fieldnames = [
        "status", "item", "description", "class", "vendor",
        "on_hand", "on_order", "committed", "curr_avail",
        "reorder_pt_netsuite", "pref_stock_netsuite",
        "qty_sold_period", "daily_sales_rate",
        "days_of_stock", "days_until_reorder",
        "suggested_order_qty",
    ]
    with open(output_full, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({
                "status": r.status,
                "item": r.item,
                "description": r.description,
                "class": r.class_name,
                "vendor": r.vendor,
                "on_hand": r.on_hand,
                "on_order": r.on_order,
                "committed": r.committed,
                "curr_avail": r.curr_avail,
                "reorder_pt_netsuite": r.reorder_pt,
                "pref_stock_netsuite": r.pref_stock,
                "qty_sold_period": r.qty_sold,
                "daily_sales_rate": r.daily_sales,
                "days_of_stock": r.days_of_stock,
                "days_until_reorder": r.reorder_signal,
                "suggested_order_qty": r.suggested_order,
            })
    print(f"\n  Full report  → {output_full}  ({len(rows):,} rows)")

    # ── Write action CSV ──────────────────────────────────────────────────────
    action_rows = [r for r in rows if r.status in ("OUT_OF_STOCK", "REORDER_NOW", "LOW_STOCK")]
    with open(output_action, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in action_rows:
            w.writerow({
                "status": r.status,
                "item": r.item,
                "description": r.description,
                "class": r.class_name,
                "vendor": r.vendor,
                "on_hand": r.on_hand,
                "on_order": r.on_order,
                "committed": r.committed,
                "curr_avail": r.curr_avail,
                "reorder_pt_netsuite": r.reorder_pt,
                "pref_stock_netsuite": r.pref_stock,
                "qty_sold_period": r.qty_sold,
                "daily_sales_rate": r.daily_sales,
                "days_of_stock": r.days_of_stock,
                "days_until_reorder": r.reorder_signal,
                "suggested_order_qty": r.suggested_order,
            })
    print(f"  Action list  → {output_action}  ({len(action_rows):,} rows need attention)")
    print()


def main():
    parser = argparse.ArgumentParser(description="NetSuite Inventory Forecasting Tool")
    parser.add_argument("--inventory",   required=True,  help="Path to FCSCurrentInventorySnapshot.xls")
    parser.add_argument("--sales",       required=True,  help="Path to FCSSalesbyItemSummary.xls")
    parser.add_argument("--sales-days",  type=int, default=61,
                        help="Number of days covered by the sales report (default: 61)")
    parser.add_argument("--lead-time",   type=int, default=30,
                        help="Supplier lead time in days (default: 30)")
    parser.add_argument("--safety-days", type=int, default=14,
                        help="Safety stock buffer in days (default: 14)")
    parser.add_argument("--output-full",   default="inventory_forecast_report.csv")
    parser.add_argument("--output-action", default="reorder_now.csv")
    args = parser.parse_args()

    run(
        inventory_path=args.inventory,
        sales_path=args.sales,
        sales_days=args.sales_days,
        lead_time=args.lead_time,
        safety_days=args.safety_days,
        output_full=args.output_full,
        output_action=args.output_action,
    )


if __name__ == "__main__":
    main()
