# Cleaning Log - Retail Orders Dataset

**Analyst:** Sangita Chatterjee  
**Date:** June 28, 2026  
**Source file:** `data/raw_orders.xlsx` (932 rows, 21 columns)  
**Output file:** `data/cleaned_orders.xlsx` (912 rows after deduplication)

---

## 1. Issues Found

### Text formatting problems
- Extra leading/trailing spaces in `region`, `segment`, `ship_mode`, `category`, and `order_status`
- Inconsistent casing (e.g., `NORTH`, `north`, `  North ` all referring to North)
- Category names written differently (`OFFICE SUPPLIES`, `Office  Supplies`, `  Furniture `)
- Payment and order status values with mixed case (`Paid `, `PENDING`, `completed`, `  Cancelled `)

### Date problems
- Multiple date formats in the raw file: `21 Jul 2024`, `08/31/2024`, `28-11-2024`, `2024-05-24`, etc.
- **93 records** where ship date was earlier than order date (likely due to ambiguous day/month parsing in the raw export)
- No completely blank date text values, but some dates needed careful parsing

### Duplicate records
- **20 exact duplicate rows** (identical across all columns)
- **31 order_id values** appeared more than once
- **12 order_id groups** had conflicting information (same order_id but different status, sales, or other fields)

### Missing values
- **26 missing region** values in raw data
- **22 missing ship_mode** values in raw data
- **18 missing discount** values in raw data

### Discount issues
- Some discounts stored as percentages (`70%`, `85%`) instead of decimals
- **Negative discounts** found (e.g., -0.19)
- **Discounts above 100%** found (values > 1.0 after conversion)
- Missing discounts where other sales fields were still present

### Calculation mismatches
- **64 records** where `quantity × unit_price × (1 - discount)` did not match the reported `sales` value (difference > ₹1)
- Similar gaps between reported profit and recalculated profit on those rows

### Order and payment status issues
- Cancelled orders (143 in raw), returned orders (164), failed payments (68), and refunded payments (72) mixed into the same dataset
- These need to be handled differently for sales reporting

---

## 2. Cleaning Actions Performed

### Text standardization (Task 2)
Applied TRIM-equivalent logic (strip spaces, collapse internal whitespace) on:
`customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`

Mapped known variants to standard labels:
| Field | Examples fixed |
|-------|----------------|
| region | `NORTH`, `west`, `  East ` → `North`, `West`, `East` |
| segment | `  Small Business ` → `Small Business` |
| category | `OFFICE SUPPLIES`, `  Furniture ` → `Office Supplies`, `Furniture` |
| ship_mode | `STANDARD CLASS`, `Standard  Class` → `Standard Class` |
| payment_status | `PENDING`, `failed` → `Pending`, `Failed` |
| order_status | `completed`, `  Completed ` → `Completed` |

Removed unwanted special characters from text fields where they appeared.

### Date cleaning (Task 3)
- Parsed all date strings into a consistent `YYYY-MM-DD` format
- Created `shipping_delay_days` = ship_date minus order_date
- Flagged records where ship date is before order date

### Duplicate handling (Task 4)
| Action | Count |
|--------|-------|
| Exact duplicates found | 20 |
| Duplicate order_id groups | 31 |
| Conflicting order_id groups | 12 |
| Records removed | 20 (exact duplicates only, kept first occurrence) |
| Records flagged for review | 18 rows tied to conflicting order_ids |

**Logic:** I removed exact duplicate rows because they add no new information. For conflicting duplicates (same order_id, different details), I kept all rows and flagged them in `data_quality_flag` instead of deleting them silently.

### Calculated columns added (Task 6)
| Column | Formula / Logic |
|--------|-----------------|
| `cleaned_discount` | Standardized decimal discount (converted `%` values) |
| `calculated_sales` | `quantity × unit_price × (1 - cleaned_discount)` |
| `calculated_profit` | `calculated_sales - cost` |
| `profit_margin` | `calculated_profit / calculated_sales` |
| `shipping_delay_days` | Days between order and ship date |
| `order_month` | Month from order_date |
| `order_year` | Year from order_date |
| `data_quality_flag` | `clean`, `warning`, or `invalid` |
| `quality_flag_detail` | Specific reasons for the flag |

---

## 3. Business Rules Applied

| Rule | Action taken |
|------|--------------|
| Missing region | Filled as `Unknown`; flagged as `region_was_missing` |
| Missing ship_mode | Filled as `Unknown`; flagged as `ship_mode_was_missing` |
| Missing discount | Set to `0` only when quantity, unit_price, and sales were all valid; otherwise flagged invalid |
| Negative discount | Kept value but flagged as `discount_negative` → **invalid** |
| Discount above 100% | Flagged as `discount_above_range` → **invalid** |
| Cancelled orders | Excluded from completed sales pivot summaries |
| Failed payments | Excluded from completed sales pivot summaries |
| Refunded orders | Summarized separately in `Refunded by Region` pivot sheet |
| Ship date before order date | Flagged as `ship_before_order` → **warning** |

### Flag assignment logic
- **clean** - no issues detected after cleaning
- **warning** - minor or review-worthy issues (duplicate conflict, missing value filled, ship date issue, sales mismatch)
- **invalid** - serious data problems (bad discount, unparseable/missing critical dates)

**Final counts:** 709 clean | 188 warning | 15 invalid

---

## 4. Assumptions Made

1. **Discount range:** Valid discounts are between 0 and 1 (0% to 100%). Anything above 1.0 is treated as invalid.
2. **Date parsing:** US-style dates (`MM/DD/YYYY`) were assumed when ambiguous. Some ship-before-order flags may be due to day/month swap in the original export - I flagged these rather than auto-correcting.
3. **Missing discount = 0:** Applied only when quantity, unit_price, and sales were all present and numeric.
4. **Completed sales definition:** Only orders with `order_status = Completed` AND `payment_status = Paid` count toward revenue pivots.
5. **Duplicate conflicts:** When the same order_id had different statuses, I assumed the latest/most complete record might be correct but flagged all conflicting rows for manual review rather than picking one automatically.
6. **Sales mismatch tolerance:** Differences of ₹1 or less between calculated and reported sales were ignored (rounding).

---

## 5. Records Removed

- **20 rows** removed - all were exact duplicates (identical in every column)
- **0 rows** removed for conflicting order_ids (these were flagged instead)

---

## 6. Records Flagged

| Flag type | Approximate count |
|-----------|-------------------|
| Ship before order | 93 |
| Sales calculation mismatch | 64 |
| Duplicate order_id conflict | 18 |
| Region was missing (filled Unknown) | 25 |
| Ship mode was missing (filled Unknown) | 21 |
| Discount missing (filled as 0) | 3 |
| Invalid discount | 15 |

---

## 7. Limitations

1. **Ambiguous dates:** Without knowing the source system's locale, some dates may have been parsed with the wrong day/month. I flagged ship-before-order cases but did not swap dates automatically.
2. **Conflicting duplicates:** Manual review is still needed for 12 order_id groups where records disagree.
3. **Sales mismatches:** The original `sales` and `profit` columns were kept as-is; I added calculated columns alongside them. Reporting uses calculated values for consistency.
4. **Automation vs Excel:** Core cleaning was done programmatically (Python/pandas) to ensure consistency across 900+ rows, but the output files are standard Excel workbooks that can be opened and validated manually.
5. **Unknown region/ship_mode:** Filling with "Unknown" keeps rows in analysis but may slightly distort regional summaries (₹190K sales in Unknown region in pivot).

---

## 8. Tools Used

- **Python 3.9** with pandas, openpyxl, matplotlib, python-dateutil
- Excel output reviewed in compatible spreadsheet format
- Cleaning script: `clean_orders.py` (reproducible pipeline)
