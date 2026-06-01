# Order Consolidation Checker

Automatically identifies which open orders can be shipped together, and writes the recommendations back to a Google Sheet — saving daily manual review time for warehouse and logistics teams.

---

## How It Works

```
NetSuite IF Export (.xls)         Google Sheet Tracker
         │                                │
         └──────────────┬─────────────────┘
                        ▼
              Load & Normalize Data
                        │
                        ▼
           Upsert IF_history Tab
         (track pick dates per SO)
                        │
                        ▼
          Eligibility Filter
    (status whitelist + exclude accounts)
                        │
                        ▼
       Group by ACCT# + ZIP Code
                        │
                        ▼
      Pairwise Time Window Check
   (ship deadline diff OR pick date diff ≤ 3 days)
                        │
                        ▼
      Large-Order Safety Check
    (no-IF order with QTY > 300 → warn & skip)
                        │
                        ▼
        Clique Validation
  (every pair in a group must be compatible)
                        │
                        ▼
    Write Recommendations to Note Column
         + Export Local CSV Report
```

---

## Features

- **Smart grouping** — groups by account + ZIP, then validates every pair in a group before merging (no false positives from transitive matches)
- **Dual time signal** — uses either the promised ship date or the NetSuite pick date, whichever is available
- **Large-order guard** — prevents merging a tracked order with an untracked high-QTY order that might delay shipment
- **Non-destructive writes** — human notes in the sheet are never overwritten; only the `[Auto]` section is refreshed
- **IF history tracking** — maintains a running `IF_history` tab so pick dates persist across daily runs
- **Urgency flags** — highlights orders due within 2 days (🔴) or 4 days (🟡)

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Set up Google credentials

Create a [Google Service Account](https://docs.gspread.org/en/latest/oauth2.html), download the JSON key, and save it as `credentials.json` in the project root.

Share your Google Sheet with the service account email.

### 3. Configure

Copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
```

Or edit `config.py` directly.

### 4. Export from NetSuite

Export today's Item Fulfillment report from NetSuite as **Excel** (SpreadsheetML `.xls` format — not CSV or binary xlsx).

Save it to the project root as `ItemFulfillments.xls` (or update `IF_XLS_PATH` in your config).

### 5. Run

```bash
# Interactive — asks before writing back to Sheet
python consolidation_checker.py

# Write results directly without prompt
python consolidation_checker.py --write

# Analyze only, do not modify Sheet
python consolidation_checker.py --dry-run
```

See [`examples/outputs/sample_terminal_output.md`](examples/outputs/sample_terminal_output.md) for what to expect.

---

## Google Sheet Requirements

| Requirement | Detail |
|-------------|--------|
| Header row | Row **3** contains column names; data starts at row 4 |
| Required columns | `Order Date`, `ACCT#`, `ZIP CODE`, `SO#`, `Status`, `Note`, `normal shipped date`, `QTY` |
| Optional column | `SEND DATE` — used to detect same-day send orders |

---

## Project Structure

```
consolidation_checker/
├── consolidation_checker.py     # Main script
├── config.py                    # Centralized configuration (env-aware)
├── requirements.txt
├── .env.example                 # Config template
├── .gitignore
├── docs/
│   └── consolidation_logic.md  # Detailed logic explanation
└── examples/
    └── outputs/
        └── sample_terminal_output.md
```

> `credentials.json` and `ItemFulfillments.xls` are excluded by `.gitignore` — never commit them.

---

## Consolidation Rules (Summary)

Two orders can be consolidated when:

1. They share the same `ACCT#` and `ZIP CODE`
2. Their ship deadlines **or** NetSuite pick dates are within **3 days** of each other
3. Neither order is blocked by the large-order safety check
4. Every other order in the proposed group also satisfies rule 2 with them (clique requirement)

For the full logic with examples, see [`docs/consolidation_logic.md`](docs/consolidation_logic.md).

---

## Configuration Reference

| Key | Default | Description |
|-----|---------|-------------|
| `consolidate_window` | `3` | Max days between deadlines/pick dates to allow merging |
| `expiry_days` | `7` | Fallback deadline = order date + this value |
| `urgent_days` | `2` | Orders due within N days are flagged 🔴 |
| `large_qty_threshold` | `300` | No-IF orders above this QTY are excluded from merging |
| `exclude_accts` | `["5042"]` | Account numbers never included in consolidation |

---

## Troubleshooting

| Error | Cause & Fix |
|-------|-------------|
| `❌ IF table file not found` | `ItemFulfillments.xls` missing — check path or `IF_XLS_PATH` in config |
| `❌ IF table parse failed` | Wrong export format — re-export from NetSuite as Excel (SpreadsheetML) |
| `❌ credentials.json not found` | Missing service account key — see [gspread auth docs](https://docs.gspread.org/en/latest/oauth2.html) |
| `⚠️ Tracker missing columns` | Column name mismatch — compare config keys against row 3 of your Sheet |
