# Sample Terminal Output

Below is an example of what the script prints when run against a real tracking sheet.

```
=======================================================
  Order Consolidation Checker
  2026-06-01 09:14
=======================================================
📥 Reading Google Sheet tracker...
   Loaded 87 rows, columns: ['Order Date', 'ACCT#', 'ZIP CODE', 'SO#', 'Status', 'Note', 'normal shipped date', 'QTY']...
📥 Reading IF table: ItemFulfillments.xls
   Parsed 63 valid IF records

📂 Managing IF_history...
🔄 Updating IF_history...
   IF table: 63 records → 41 matched in tracker, rest ignored
   Updated existing records: 38
   Added new records: 3
   Cleaned up 2 expired IF records

📊 Analysis pool: 24 SOs
   stock order: 18  |  FcsRepleshing: 3  |  hold: 2  |  Today-SEND: 1
   Excluded account ['5042']: 1 SO (not eligible for consolidation)

=======================================================
  ✅ Consolidatable SOs: 14  across 6 groups
  ⚠️  Merge warnings: 1
  🔴 Urgent / Expired: 2
=======================================================

Consolidation detail:
  Group 1  🔴 Urgent
    📦 IF    QTY:120  Due:6/2  SO:100234
       No IF  QTY:80   Due:6/3  SO:100287

  Group 2
    📦 IF    QTY:200  Due:6/5  SO:100301
    📦 IF    QTY:150  Due:6/6  SO:100344

  Group 3
    📦 IF    QTY:95   Due:6/7  SO:100198
       No IF  QTY:60   Due:6/7  SO:100412

  Group 4  🟡 Attention
    📦 IF    QTY:310  Due:6/4  SO:100089
    📦 IF    QTY:275  Due:6/5  SO:100103

  Group 5
       No IF  QTY:45   Due:6/8  SO:100455
       No IF  QTY:30   Due:6/9  SO:100478

  Group 6
    📦 IF    QTY:180  Due:6/6  SO:100512
    📦 IF    QTY:220  Due:6/6  SO:100533
    📦 IF    QTY:140  Due:6/7  SO:100561

Is this correct? Write results back to Google Sheet Note column? (y/n): y
   Write progress: 14/14
✅ Updated 14 notes in Google Sheet (human notes preserved)
📄 Local report: consolidation_report_20260601.csv (14 records)

🎉 Done!
```

---

## What Gets Written to the Note Column

| Before (human note) | After script runs |
|---|---|
| *(empty)* | `[Auto] Can consolidate with SO100287 \| IF picked: 5/30` |
| `Check with warehouse` | `Check with warehouse  [Auto] Can consolidate with SO100344 \| IF picked: 5/31` |
| `[Auto] Can consolidate with SO100287` | `[Auto] Can consolidate with SO100287 \| IF picked: 5/30` *(refreshed)* |

The `[Auto]` tag marks machine-generated content. Human notes before the tag are always preserved and never overwritten.
