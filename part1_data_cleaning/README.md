# Part 1: Business Data Cleaning, Validation & Excel Reporting

## Problem Summary

Raw retail order data for an Indian e-commerce business contained multiple data quality issues that made it unsuitable for direct analysis. These issues included inconsistent text casing and whitespace, mixed date formats, invalid and missing discount values, duplicate records, and orders with failed or cancelled statuses that could inflate sales figures. This project performs end-to-end data cleaning, applies business rules, and produces pivot-based summaries suitable for business reporting.

---

## Dataset Description  

The dataset consists of e-commerce sales records, detailing transaction-level information. Key fields include:
* **Identifiers:** Order ID
* **Temporal Data:** Order Date, Ship Date
* **Categorical Data:** Region, Ship Mode, Customer Segment, Category, Sub-Category
* **Financial Data:** Unit Price, Quantity, Discount, Sales, Profit
* **Status Flags:** Order Status (e.g., Completed, Cancelled), Payment Status (e.g., Paid, Failed, Refunded)

| Attribute | Detail |
|---|---|
| Source file | `raw_orders.xlsx` |
| Total raw records | 932 |
| Records after cleaning | 912 (20 exact duplicates removed) |
| Date range | January 2024 – October 2025 |
| Columns (raw) | 21 |
| Columns (cleaned) | 35 (21 original + 14 calculated/flag columns) |
| Geography | Pan-India (North, South, East, West regions) |
| Categories | Technology, Furniture, Office Supplies |
| Customer segments | Consumer, Corporate, Small Business, Home Office |

Key columns include `order_id`, `order_date`, `ship_date`, `customer_name`, `segment`, `region`, `category`, `sub_category`, `quantity`, `unit_price`, `discount`, `sales`, `cost`, `profit`, `payment_status`, and `order_status`.

---

## Tools Used

* **Microsoft Excel:** Utilized for data cleaning, text standardization, complex logical/mathematical formulas, and building interactive Pivot Table dashboards.
* **Markdown:** Utilized for project documentation and reporting.

---

## Cleaning Steps Performed

### Task 1 — Dataset Inspection
- Reviewed all 21 columns for data types, formats, and anomalies
- Identified 5 date formats, invalid discounts, blank fields, and casing issues

### Task 2 — Text Field Cleaning
All 10 text fields (`region`, `segment`, `category`, `sub_category`, `ship_mode`, `customer_name`, `state`, `city`, `payment_status`, `order_status`) were cleaned using a 4-step pipeline:
1. **TRIM** — removed leading/trailing spaces

2. **PROPER** — standardised to Title Case

Helper columns (e.g. `clean_region`) were created, verified, pasted back as values, then deleted.

### Task 3 — Date Standardisation
- 5 mixed date formats found (`DD Mon YYYY`, `MM/DD/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD`, `DD/MM/YYYY`)
- All dates normalised to `DD Mon YYYY` using DateValue and than formatting it to date.
- `shipping_delay_days` column created: `= ship_date - order_date`
- 21 records with `ship_date < order_date` flagged as invalid

### Task 4 — Duplicate Handling

- **20 exact duplicate rows removed** via Data → Remove Duplicates
- 24 conflicting duplicate `order_id` records (12 pairs) flagged for review — not deleted

### Task 5 — Business Rules Applied
- Missing `region` (25 records) and `ship_mode` (21 records) filled with `Unknown`
- Invalid discounts flagged; `data_quality_flag` column created per record
- Cancelled, failed, and refunded orders retained but excluded from completed sales summaries

### Task 6 — Calculated Columns Added

| Column | Formula |
|---|---|
| `cleaned_discount` | Text % converted to decimal; negatives and >0.5 set to 0 |
| `calculated_sales` | `quantity × unit_price × (1 − cleaned_discount)` |
| `calculated_profit` | `calculated_sales − cost` |
| `profit_margin` | `calculated_profit / calculated_sales` |
| `shipping_delay_days` | `ship_date − order_date` |
| `order_month` | `TEXT(order_date, "mmm")` |
| `order_year` | `YEAR(order_date)` |
| `data_quality_flag` | Nested IF — flags Clean / Warning / Invalid per row |

---

## Business Rules Applied

| Rule | Action Taken |
|---|---|
| Missing `region` | Filled with `Unknown`; flagged in quality report |
| Missing `ship_mode` | Filled with `Unknown`; flagged in quality report |
| Missing `discount` | Treated as 0 in `cleaned_discount` where other fields valid (18 records) |
| Negative discount | `cleaned_discount` set to 0; flagged `Invalid` (15 records) |
| Discount above 50% | `cleaned_discount` set to 0; flagged `Invalid` (15 records) |
| Discount as text % | Converted to decimal first, then range check applied (8 records) |
| Cancelled orders | Retained in dataset; excluded from completed sales pivot summaries (145 records) |
| Failed payments | Retained in dataset; excluded from completed sales pivot summaries (69 records) |
| Refunded orders | Retained; summarised separately in  pivot sheet (71 records) |
| Ship date before order date | Flagged as `Invalid - Ship Before Order Date` (21 records) |

All business rule decisions are fully documented in `outputs/cleaning_log.md`.

---

## Summary of Data Quality Issues Found

### Record-Level Summary

| Status | Count | Definition |
|---|---|---|
| Original Records | 932 | Total raw entries |
| Duplicate Rows Removed | 20 | Exact full-row duplicates |
| Final Records | 912 | After duplicate removal |
| Clean Records | 797 | Passed all business rules |
| Warning Records | 64 | Missing region/ship_mode or conflicting duplicate order_ids |
| Invalid Records | 51 | Invalid dates or invalid discounts |

### Issue Breakdown

| Issue Type | Count | Action |
|---|---|---|
| Missing `region` | 25 | Filled with Unknown |
| Missing `ship_mode` | 21 | Filled with Unknown |
| Missing `discount` | 18 | Set to 0 in `cleaned_discount` |
| Negative discount | 15 | Set to 0 in `cleaned_discount`, flagged Invalid |
| Discount above 50% | 15 | Set to 0 in `cleaned_discount`, flagged Invalid |
| Discount as text % | 8 | Converted to decimal, range-checked |
| Ship date before order date | 21 | Flagged Invalid, records retained |
| Exact duplicate rows | 20 | Removed |
| Conflicting duplicate order_ids | 24 | Flagged for review, retained |
| Sales/profit calc mismatch | 73 | Flagged Mismatch, `calculated_sales` used for analysis |
| Cancelled orders | 145 | Excluded from completed sales summaries |
| Failed payments | 69 | Excluded from completed sales summaries |
| Refunded orders | 71 | Summarised separately |

> Full details available in `outputs/data_quality_report.xlsx`.

---

## Summary of Final Pivot Reports

All pivot tables are in `outputs/pivot_summary.xlsx`. Cancelled orders and failed/refunded payments are excluded from completed sales figures in all relevant pivots.

### P1 — Sales and Profit by Region:    
Gives High-level geographic performance.
Sorted descending by `calculated_sales`. South leads in revenue.

| Region | Calculated Sales | Calculated Profit |
|---|---|---|
| South | ₹18,15,177.55 | ₹5,22,707.13 |
| West | ₹17,21,696.38 | ₹4,98,316.59 |
| East | ₹17,01,688.75 | ₹5,21,231.69 |
| North | ₹15,32,434.30 | ₹4,64,251.97 |
| Unknown | ₹1,95,616.68 | ₹66,305.48 |
| **Grand Total** | **₹69,66,613.66** | **₹20,72,812.86** |

### P2 — Sales and Profit by Category and Sub-Category: 
Gives Product portfolio performance hierarchy.
Sorted descending by profit within each category. Technology is the top revenue category.

| Category | Calculated Sales | Calculated Profit |
|---|---|---|
| Technology | ₹24,91,299.71 | ₹7,53,380.30 |
| Furniture | ₹22,98,541.71 | ₹6,83,464.03 |
| Office Supplies | ₹21,76,772.25 | ₹6,35,968.54 |

Top sub-categories by profit: Copiers (₹2,17,541), Chairs (₹2,05,061), Storage (₹1,73,646).

### P3 — Order Count by Ship Mode:  
It gives Volume of logistical channels utilized.
Filtered for Corporate segment, excluding cancelled orders.

| Ship Mode | Order Count |
|---|---|
| Standard Class | 54 |
| First Class | 43 |
| Second Class | 42 |
| Same Day | 32 |
| Unknown | 2 |

### P4 — Profit Margin by Customer Segment: 
Accurate, recalculated profit margins per customer segment using calculated fields.
All segments maintain ~30% profit margin on completed orders.

| Segment | Profit Margin |
|---|---|
| Consumer | 30% |
| Corporate | 30% |
| Home Office | 30% |
| Small Business | 29% |

### P5 — Refunded/Cancelled/Failed Orders by Region:    
A comprehensive breakdown of Refunded, Cancelled, and Failed orders by Region.
North has the highest count of problem orders (25), followed by South (18) and West (13).

### P6 — Monthly Sales Trend:   
Time-series analysis of completed revenue.
Filtered for completed paid orders only. 2024 total: ₹38,27,767. 2025 total (Jan–Oct): ₹31,38,846. Peak months: June 2024 (₹4,23,635) and February 2025 (₹4,36,513).

---

## Key Business Insights

1. **South is the top-performing region** with ₹18.15 lakh in sales, but East is very close in profit (₹5.21 lakh vs South's ₹5.22 lakh), suggesting East has a slightly better cost structure.

2. **Technology drives the most revenue** (₹24.9 lakh, 36% of total), with Copiers being the single most profitable sub-category at ₹2.17 lakh.

3. **Profit margins are uniform across all customer segments** at ~30%, indicating pricing and discount policies are consistent regardless of customer type.

4. **Standard Class is the most used shipping mode** in the Corporate segment (54 orders), while Same Day is least used (32 orders), suggesting cost-consciousness among corporate buyers.

5. **North region has the most problem orders** (25 cancelled/refunded/failed), which is worth investigating for operational or fulfilment issues.

6. **73 records (8% of dataset) have sales/profit calculation mismatches** between the raw data and the formula-derived values — these likely reflect manual adjustments or data entry errors in the source system.

7. **2024 revenue (₹38.3 lakh) outpaced 2025 YTD (₹31.4 lakh through October)**, but this is expected since 2025 data only covers 10 months. Monthly averages are comparable.

8. **18% of records are either cancelled, failed, or refunded** (285 out of 912), representing significant lost revenue that warrants operational review.

---

## Assumptions and Limitations

### Assumptions
1. Valid discount range is **0% to 50%** based on the `business_rules` reference tab
2. Blank discount = no discount applied (not unknown) — treated as 0
3. `DD Mon YYYY` chosen as the standard date format (Indian locale convention)
4. Title Case is the correct casing standard for all text fields
5. Cancelled, failed, and refunded records are retained for audit — excluded only from sales summaries
6. Conflicting duplicate `order_id` records cannot be auto-resolved without source system access
7. `calculated_sales` and `calculated_profit` supersede raw `sales` and `profit` where mismatches are detected

### Limitations
1. **Single-flag logic** — `data_quality_flag` captures only the first issue per row; records with multiple issues show only one flag
2. **Date ambiguity** — dates like `03/12/2024` could be interpreted as MM/DD or DD/MM; edge cases may have been misassigned
3. **No region lookup** — missing regions filled as `Unknown` since no state-to-region mapping was available
4. **Unresolved duplicate order_ids** — 24 flagged records could not be resolved without source system access
5. **Sales mismatch root cause unknown** — 73 mismatches flagged but not auto-corrected; may represent legitimate adjustments
6. **No change history in Excel** — this log is the only audit trail; a production pipeline would use version control
7. **2025 data is incomplete** — dataset ends in October 2025, so year-over-year comparisons should account for missing months

