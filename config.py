"""
config.py — Central configuration for Order Consolidation Checker.

Copy .env.example to .env and fill in your values, or edit CONFIG directly.
"""

import os

CONFIG = {
    # ── Google Sheet ──────────────────────────────────────────────────────────
    "sheet_id":         os.getenv("SHEET_ID",        "1kARjfsVweCBXcaS77a_CB3KiE92DgIOajoQ6IJ8wSCY"),
    "tracking_tab":     os.getenv("TRACKING_TAB",    "2026"),
    "if_history_tab":   os.getenv("IF_HISTORY_TAB",  "IF_history"),

    # ── File paths ────────────────────────────────────────────────────────────
    "if_xls_path":      os.getenv("IF_XLS_PATH",      "ItemFulfillments.xls"),
    "credentials_file": os.getenv("CREDENTIALS_FILE", "credentials.json"),

    # ── Tracking sheet column names (must match row 3 headers exactly) ────────
    "col_order_date":   os.getenv("COL_ORDER_DATE",   "Order Date"),
    "col_acct":         os.getenv("COL_ACCT",         "ACCT#"),
    "col_zip":          os.getenv("COL_ZIP",          "ZIP CODE"),
    "col_so":           os.getenv("COL_SO",           "SO#"),
    "col_status":       os.getenv("COL_STATUS",       "Status"),
    "col_note":         os.getenv("COL_NOTE",         "Note"),
    "col_shipped_date": os.getenv("COL_SHIPPED_DATE", "normal shipped date"),
    "col_send_date":    os.getenv("COL_SEND_DATE",    "SEND DATE"),
    "col_qty":          os.getenv("COL_QTY",          "QTY"),

    # ── Accounts excluded from consolidation ──────────────────────────────────
    # e.g. internal warehouse accounts that should never be merged with others
    "exclude_accts": [
        a.strip()
        for a in os.getenv("EXCLUDE_ACCTS", "5042").split(",")
        if a.strip()
    ],

    # ── Consolidation parameters ──────────────────────────────────────────────
    "expiry_days":         int(os.getenv("EXPIRY_DAYS",          7)),
    "urgent_days":         int(os.getenv("URGENT_DAYS",          2)),
    "consolidate_window":  int(os.getenv("CONSOLIDATE_WINDOW",   3)),
    "large_qty_threshold": int(os.getenv("LARGE_QTY_THRESHOLD",  300)),
}
