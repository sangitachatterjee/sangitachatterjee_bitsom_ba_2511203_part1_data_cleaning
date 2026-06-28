# Part 1 - Data Cleaning (Retail Orders)

**Capstone Project | Business Analytics**  
**Author:** Sangita Chatterjee

---

## Problem Summary

Our retail company exported order-level sales data from multiple internal systems into a single Excel file. Before any reporting or analysis could happen, the data needed serious cleanup - inconsistent text formatting, mixed date formats, duplicate records, missing values, invalid discounts, and calculation errors were all present in the raw export.

My job was to produce a clean, validated, analysis-ready dataset along with quality reports and pivot summaries that management can actually use for business review.

---

## Dataset Description

| Item | Detail |
|------|--------|
| **Source** | Google Drive shared folder - Part 1 dataset |
| **Raw file** | `data/raw_orders.xlsx` |
| **Records** | 932 orders (raw) → 912 after removing exact duplicates |
| **Columns** | 21 original fields covering order details, customer info, product category, pricing, and status |
| **Key fields** | order_id, order_date, ship_date, customer_name, segment, region, category, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status |
| **Cleaned file** | `data/cleaned_orders.xlsx` - includes 10 additional calculated/flag columns |

---

## Tools Used

- **Microsoft Excel-compatible workbooks** (.xlsx) for all deliverables
- **Python 3.9** with pandas, openpyxl, and matplotlib for cleaning automation and screenshot generation
- Excel techniques mirrored in code: TRIM, SUBSTITUTE, standardized Find-and-Replace mappings, date parsing, and pivot-style aggregations
- Reproducible script: `clean_orders.py`

---

## Cleaning Steps Performed

1. **Preserved raw data** - `raw_orders.xlsx` was never modified; all work saved to `cleaned_orders.xlsx`
2. **Text cleaning** - Trimmed spaces, fixed casing, standardized category/region/segment/ship mode/status labels
3. **Date standardization** - Converted mixed formats to `YYYY-MM-DD`; calculated shipping delay in days
4. **Duplicate handling** - Removed 20 exact duplicates; flagged 12 conflicting order_id groups for review
5. **Missing value treatment** - Filled missing region/ship_mode as "Unknown"; handled missing discounts per business rules
6. **Business rule validation** - Flagged invalid discounts, bad shipping dates, and calculation mismatches
7. **Calculated columns** - Added cleaned_discount, calculated_sales, calculated_profit, profit_margin, order_month/year, data_quality_flag
8. **Reports** - Built data quality report, pivot summary, cleaning log, and screenshots

---

## Business Rules Applied

| Rule | What I did |
|------|------------|
| Missing region | Filled as "Unknown", flagged in quality report |
| Missing ship_mode | Filled as "Unknown", flagged in quality report |
| Missing discount | Treated as 0 only when other sales fields were valid |
| Negative discount | Flagged as invalid |
| Discount > 100% | Flagged as invalid |
| Cancelled orders | Excluded from completed sales pivots |
| Failed payments | Excluded from completed sales pivots |
| Refunded orders | Summarized separately by region |
| Ship date before order date | Flagged as invalid shipping record |

Full details in `outputs/cleaning_log.md`.

---

## Data Quality Issues Found

| Issue | Count |
|-------|-------|
| Exact duplicate rows | 20 (removed) |
| Conflicting duplicate order_ids | 12 groups (flagged) |
| Missing region | 26 |
| Missing ship_mode | 22 |
| Missing discount | 18 |
| Invalid discount (negative or >100%) | 15 |
| Ship date before order date | 93 |
| Sales/profit calculation mismatch | 64 |
| **Final flag breakdown** | 709 clean / 188 warning / 15 invalid |

See `outputs/data_quality_report.xlsx` for sheet-by-sheet breakdown.

---

## Pivot Reports Summary

File: `outputs/pivot_summary.xlsx`

| Sheet | What it shows |
|-------|---------------|
| Sales by Region | Total sales & profit for completed+paid orders, sorted descending (with filter) |
| Sales by Category | Sales & profit by category and sub-category, sorted by profit (with filter) |
| Orders by Ship Mode | Order count by shipping method |
| Margin by Segment | Average profit margin and total sales by customer segment |
| Problem Orders | Cancelled, returned, failed, and refunded orders by region |
| Monthly Trend | Monthly sales trend for completed+paid orders |
| Refunded by Region | Separate refund summary as required |

---

## Key Business Insights

1. **South leads in completed sales** (~₹15.5L), followed closely by West (~₹15.1L) and East (~₹13.8L). North is slightly behind at ~₹12.8L.
2. **Technology > Copiers** is the top category-subcategory combo by profit (~₹1.95L profit), ahead of Furniture > Chairs.
3. **Home Office segment** has the highest average profit margin (~29.4%), while Consumer segment has the lowest (~22.0%) - worth investigating pricing or product mix.
4. **Standard Class** is the most used ship mode (242 orders), but all four main modes are fairly balanced.
5. **Monthly sales are relatively stable** in 2024 (~₹2.1L–₹3.1L per month), giving management a reliable baseline for forecasting.
6. **Refunds are concentrated in South** (20 refunded orders, ~₹2.25L) - may need operational follow-up.
7. **~21% of records (188)** carry warning flags - mostly shipping date and calculation issues - and should be reviewed before using in executive dashboards.

---

## Assumptions and Limitations

- Valid discount range assumed as 0–100% (0.0 to 1.0 decimal)
- Ambiguous dates (e.g., 06/08/2024) parsed as MM/DD/YYYY; some ship-before-order flags may be parsing artifacts
- Conflicting duplicate order_ids were flagged, not auto-resolved - manual review recommended
- Completed sales pivots only include Paid + Completed orders
- See full assumptions list in `outputs/cleaning_log.md`

---

## Screenshots Included

| File | Description |
|------|-------------|
| `screenshots/raw_data_preview.png` | Raw dataset before any cleaning |
| `screenshots/cleaned_data_preview.png` | Cleaned data with calculated columns |
| `screenshots/pivot_summary_1.png` | Sales & profit by region (bar chart) |
| `screenshots/pivot_summary_2.png` | Monthly sales trend (line chart) |

---

## Repository Structure

```
├── data/
│   ├── raw_orders.xlsx          ← Original data (untouched)
│   └── cleaned_orders.xlsx      ← Cleaned + calculated columns
├── outputs/
│   ├── data_quality_report.xlsx ← Quality checks by issue type
│   ├── pivot_summary.xlsx       ← Business pivot summaries
│   └── cleaning_log.md          ← Detailed cleaning documentation
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   └── pivot_summary_2.png
├── clean_orders.py              ← Reproducible cleaning script
└── README.md
```

---

## How to Reproduce

```bash
pip install pandas openpyxl matplotlib python-dateutil gdown
python3 clean_orders.py
```

This regenerates `cleaned_orders.xlsx`, both output reports, and all four screenshots from the raw file.
