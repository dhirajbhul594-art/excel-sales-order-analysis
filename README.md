# Sales Order Analysis (Excel Project)

A small Excel project analyzing 50 retail orders across 3 products — built to demonstrate core Excel skills for business analysis and reporting.

## What this project does
- Raw order data (Order ID, Product, Quantity, Status, City) across 50 orders
- Looks up product prices using **VLOOKUP** and **INDEX-MATCH** (two different methods, same result)
- Wraps VLOOKUP in **IFERROR** to handle missing matches safely
- Calculates Total Value and flags orders using **IF**
- Summarizes totals using **SUM**, **COUNT/COUNTA**, **SUMIFS**, and **COUNTIFS**
- Uses **Data Validation** for a clean Status dropdown (Pending/Completed/Cancelled)
- Uses **Conditional Formatting** to highlight low-quantity orders
- Includes a **Pivot Table** and **chart** summarizing total quantity by product

## Sheets
| Sheet | Purpose |
|---|---|
| `Orders` | Raw + calculated order data |
| `PriceList` | Product price reference table (used for lookups) |
| `Pivot Table` | Summary view with chart |

## Skills demonstrated
VLOOKUP · INDEX-MATCH · IF · IFERROR · SUM · COUNT/COUNTA · SUMIFS · COUNTIFS · Sort · Filter · Data Validation · Conditional Formatting · Pivot Table · Chart
