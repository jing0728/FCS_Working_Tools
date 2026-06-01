# Consolidation Logic Deep Dive

This document explains exactly how the script decides whether two orders can be shipped together.

---

## Eligibility Filter

Only orders in the following statuses enter the analysis pool:

| Status | Meaning |
|--------|---------|
| `stock order` | Unprocessed, waiting to be picked |
| `fcsrepleshing` | Inventory replenishment in progress |
| `hold` | On hold, still shippable |
| `send` *(today only)* | Handed to shipping today — still time to consolidate |

Orders with an empty `SO#`, or belonging to excluded accounts (e.g. internal warehouse account `5042`), are filtered out before analysis.

---

## Time Window Check

Two orders are candidates for consolidation if their shipment timelines overlap within a configurable window (default **3 days**).

The script checks **two independent date signals**:

```
Deadline diff  = |deadline_A  − deadline_B|  ≤ window   (3 days)
   OR
IF date diff   = |pick_date_A − pick_date_B| ≤ window   (3 days)
```

**Deadline** is resolved in priority order:
1. `normal shipped date` column (explicit promised ship date)
2. `Order Date + expiry_days` (fallback, default 7 days)

**IF date** is the most recent Item Fulfillment pick date for that SO, sourced from the `IF_history` tab.

---

## Large-Order Safety Check

When one order **has** an IF record and the other **does not**, there is a risk that the untracked order requires a long pick time that would delay the consolidated shipment.

Rule: if the no-IF order's `QTY > large_qty_threshold` (default 300), the pair is **excluded** from consolidation and a warning is written to the Note instead.

```
A (has IF) + B (no IF, QTY=450)  →  ⚠️ QTY=450 no pick record, merge with A abandoned
```

Crucially, excluding B does **not** prevent A from merging with a different order C if A–C passes all checks independently.

---

## Clique Grouping

Simple pair-wise merging can produce invalid groups where A–B and B–C are compatible, but A–C are not. The script avoids this with a **clique check**: a group is only formed when **every pair** within it satisfies the time window.

```
Example:
  A–B: ship dates 3 days apart  ✅
  B–C: ship dates 2 days apart  ✅
  A–C: ship dates 6 days apart  ❌

Result:
  Group {A, B}   ← valid clique
  C alone        ← not merged
```

The algorithm uses a greedy maximum-clique search, starting from the node with the most connections.

---

## Note Writing Strategy

The script appends suggestions to the `Note` column using an `[Auto]` sentinel:

```
<human content>  [Auto]<machine content>
```

On each run, only the `[Auto]...` portion is refreshed. Content before `[Auto]` is read-only for the script, ensuring human annotations are never lost.

Cells that already have the exact same `[Auto]` content are skipped (no unnecessary API writes).

---

## Configuration Reference

| Parameter | Default | Effect |
|-----------|---------|--------|
| `consolidate_window` | `3` | Max days between deadlines or pick dates to allow merging |
| `expiry_days` | `7` | Days after order date used as fallback deadline |
| `urgent_days` | `2` | Orders due within this many days are flagged 🔴 Urgent |
| `large_qty_threshold` | `300` | QTY above this for a no-IF order triggers a merge warning |
