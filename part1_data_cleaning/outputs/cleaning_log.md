# Cleaning Log  

## 1. Issues Found in Raw Dataset

### 1.1 Text Field Issues

| Field | Issue |
|---|---|
| `region` | Leading/trailing spaces, ALL CAPS (e.g. `NORTH`, `EAST`) |
| `segment` | Leading/trailing spaces, mixed case (e.g. `  Small Business `) |
| `category` | Leading/trailing spaces, ALL CAPS (e.g. `OFFICE SUPPLIES`) |
| `ship_mode` | ALL CAPS, trailing spaces (e.g. `STANDARD CLASS`, `Second Class `) |
| `customer_name` | ALL CAPS (e.g. `PRIYA MENON`), double spaces (e.g. `Vikram  Iyer`) |
| `sub_category` | Trailing spaces, inconsistent casing |
| `payment_status` | Trailing spaces (e.g. `Paid `) |
| `order_status` | Trailing spaces, lowercase (e.g. `completed`, `  Completed`) |
| `state` | No issues found |
| `city` | No issues found |

### 1.2 Date Issues

| Issue | Description |
|---|---|
| Mixed date formats | 4 formats found: `DD Mon YYYY`, `MM/DD/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD` |
| Ship date before order date | ship_date earlier than order_date (e.g. ORD-2024-10073, ORD-2024-10143, ORD-2025-10128) |
| Text dates not parsed by Excel | All date strings stored as left-aligned text |

### 1.3 Discount Issues

| Issue | Examples | Count  |
|---|---|---|
| Negative discount | -0.24, -0.15, -0.09 | 15 records |
| Discount above 50% | 0.55, 0.65 | 15 records |
| Discount stored as text % | `70%`,  `85%` | 8 records |
| Blank discount | No discount value present | 18 records |

### 1.4 Missing Value Issues

| Field | Missing Count | Action |
|---|---|---|
| `region` | 25 records | Filled with `Unknown` |
| `ship_mode` | 21 records | Filled with `Unknown` |
| `discount` | 18 records | Treated as 0 in `cleaned_discount` where other fields valid |

### 1.5 Duplicate Issues

| Type | Description |
|---|---|
| Exact duplicate rows | Rows where all 20 columns are identical |
| Conflicting duplicate order_ids | Same order_id, different values in other columns |

### 1.6 Order Status & Payment Issues

| Issue | Description |
|---|---|
| Cancelled orders | Present in dataset; must not contribute to completed sales summary |
| Failed payments | Present in dataset; must not contribute to completed sales summary |
| Refunded orders | Present in dataset; must be summarized separately |
| Inconsistent status combos | E.g. `Failed` payment with `Completed` status |

### 1.7 Sales/Profit Calculation Mismatches

Several records show `sales` values that do not match `quantity x unit_price x (1 - discount)`. These were identified by comparing the raw `sales` column against the newly calculated `calculated_sales` column.

---

## 2. Cleaning Actions Performed

### 2.1 Text Field Cleaning (Task 2)

All 10 text fields were cleaned using a 4-step pipeline applied via Excel helper columns:

1. **TRIM** — removed leading and trailing spaces
2. **SUBSTITUTE** (run twice) — collapsed internal double/triple spaces to single space
3. **Removed special characters** — using Find & Replace where applicable
4. **PROPER** — standardised all values to Title Case

Helper columns were created temporarily (e.g. `clean_region`, `clean_segment`), verified, then values were pasted back over the original columns using Paste Special > Values. Helper columns were deleted after copy-paste.

**Result:** All 10 text fields now have clean, consistent, Title Case values with no extra spaces.

### 2.2 Date Cleaning (Task 3)

- All `order_date` and `ship_date` values standardised to `DD Mon YYYY` format

- converted all textual dates to numeric date-values using `datevalue`  
formula used:   
`=IF(
   ISNUMBER(SEARCH("/",B2)),
   DATE(RIGHT(B2,4),LEFT(B2,2),MID(B2,4,2)),
IF(
   MID(B2,5,1)="-",
   DATE(LEFT(B2,4),MID(B2,6,2),RIGHT(B2,2)),
IFERROR(
   DATEVALUE(B2),
   DATE(RIGHT(B2,4),MID(B2,4,2),LEFT(B2,2))
)))`

- than converted the date value to date format `DD Mon YYYY`

- `shipping_delay_days` column created: `= ship_date - order_date`, formatted as Number and `negative` values were used to identify ship before order


### 2.3 Duplicate Handling (Task 4)


- Excel **Data > Remove Duplicates (all columns)** used to remove exact duplicate rows
- temporary column `order_id_dup_count` column uisng formula `=COUNTIF(A:A, A2)` used to find order_ids appearing more than once
- Conflicting duplicate order_ids cells flagged with `orange` colour.
- Conflicting records **retained** — not deleted silently

### 2.4 Missing Value Fills (Task 5)

- **Missing `region`:** Selected blanks using Ctrl+G > Special > Blanks, filled with `Unknown` via Ctrl+Enter
- **Missing `ship_mode`:** Same method, filled with `Unknown`
- **Missing `discount`:** Left blank in original column; treated as 0 in `cleaned_discount` column where other fields are valid

### 2.5 Discount Standardisation (Task 6 — cleaned_discount)

- Text percentage values (e.g. `70%`) converted to decimal using: `= VALUE(LEFT(P2, LEN(P2)-1)) / 100`
- Negative values converted to 0
- Values above 0.5 (50%) converted to 0
- Valid values (0.0 to 0.5) retained as-is
- Result stored in new `cleaned_discount` column

### 2.6 Calculated Columns Added (Task 6)

| Column | Formula Used |
|---|---|
| `cleaned_discount` | Standardised decimal discount (see above) |
| `calculated_sales` | `= quantity x unit_price x (1 - cleaned_discount)` |
| `calculated_profit` | `= calculated_sales - cost` |
| `profit_margin` | `= calculated_profit / calculated_sales` (with divide-by-zero guard) |
| `shipping_delay_days` | `= ship_date - order_date` |
| `order_month` | `= TEXT(order_date, "mmm")` |
| `order_year` | `= YEAR(order_date)` |
| `data_quality_flag` | Nested IF logic — see Section 3 |

---

## 3. Business Rules Applied

### Rule 1: Missing Region — Fill as Unknown
- **Action:** All blank `region` values filled with the string `Unknown`
- **Flag:** Records noted in `data_quality_report.xlsx` > Missing Values sheet


### Rule 2: Missing Ship Mode — Fill as Unknown
- **Action:** All blank `ship_mode` values filled with the string `Unknown`
- **Flag:** Records noted in `data_quality_report.xlsx` > Missing Values sheet


### Rule 3: Missing Discount — Treat as 0 if Other Fields Valid
- **Action:** Where `discount` was blank and `quantity`, `unit_price`, and `sales` were all present and valid, `cleaned_discount` was set to 0
- **Visual:** Light Green conditional formatting on original `discount` column
- **Assumption:** A missing discount means no discount was applied, not that the discount is unknown
- **Exception:** If other sales fields were also missing or invalid, the record was flagged 

### Rule 4: Negative Discount — Flag as Invalid
- **Action:** Negative discount values retained in original `discount` column; `cleaned_discount` set to 0 for calculation purposes

- **Visual:** Red conditional formatting on original `discount` column
- **Assumption:** Negative discounts are data entry errors, not intentional price increases

### Rule 5: Discount Above Allowed Range — Flag as Invalid
- **Threshold:** 0.5 (50%) based on `business_rules` reference tab in source file
- **Action:** Discounts above 0.5 retained in original column; `cleaned_discount` set to 0
- **Visual:** Light Blue conditional formatting on original `discount` column
- **Text percentages:** Values like `70%`, `80%` converted to 0.70, 0.80 first, then flagged and set to 0 in `cleaned_discount`

### Rule 6: Cancelled Orders — Exclude from Completed Sales Summary
- **Action:** Records with `order_status = "Cancelled"` retained in dataset
- **Exclusion:** Filtered out when building completed sales pivot tables (Task 8)

- **Assumption:** Cancelled orders represent orders that never fulfilled; including them would overstate revenue

### Rule 7: Failed Payments — Exclude from Completed Sales Summary
- **Action:** Records with `payment_status = "Failed"` retained in dataset
- **Exclusion:** Filtered out from all completed sales summaries

- **Assumption:** Failed payments mean revenue was not collected; these records should not inflate sales figures

### Rule 8: Refunded Orders — Summarise Separately
- **Action:** Records with `payment_status = "Refunded"` retained in main dataset
- **Separate summary:** Dedicated pivot table created in `pivot_summary.xlsx`  sheet
- **Exclusion:** Not included in completed sales totals
- **Assumption:** Refunded orders represent revenue reversals and require separate business review

### Rule 9: Ship Date Before Order Date — Flag as Invalid Shipping Record
- **Action:** Records where `ship_date < order_date` flagged in `date_flag` column as `Ship Before Order`
- **Records retained:** Not deleted; kept for audit and review

- **Assumption:** These are data entry or system export errors; the actual shipment timeline cannot be verified

---

## 4. Assumptions Made

1. **Discount range of 0% to 50%** is the valid business range, based on the `business_rules` reference tab in the source file
2. **Blank discount = no discount applied**, not unknown — treated as 0 where other fields are valid
3. **Cancelled and failed records are kept** in the dataset for audit trail purposes; excluded only from summary calculations
4. **Conflicting duplicate order_ids are not deleted** — may represent legitimate amendments or system errors requiring human review
5. **Title Case is the correct standard** for all text fields, based on the majority pattern in clean records
6. **`DD Mon YYYY`** selected as the standard date format for consistency with Indian locale conventions
7. **`calculated_sales` and `calculated_profit` supersede raw `sales` and `profit`** where a mismatch is detected, as the calculated versions follow verified business formula logic
8. **Missing region/ship_mode cannot be inferred** from other columns — no state-to-region mapping was available, so `Unknown` was the only viable fill value

---

## 5. Records Removed

| Reason | Count |
|---|---|
| Exact duplicate rows (all 20 columns identical) | 20 |
| **Total records removed** | 20|

> No conflicting duplicate records were deleted. Only exact full-row duplicates were removed using Excel's built-in Remove Duplicates tool with all columns selected.

---

## 6. Records Flagged

| Flag Value | Meaning | Approx. Count |
|---|---|---|
| `Invalid - Negative Discount` | Discount value below 0 | 15 |
| `Invalid - Discount Above Range` | Discount above 0.5 (50%) | 15 |
| `Invalid - Ship Before Order Date` | ship_date earlier than order_date | 21 |
| `Warning - Missing Region` | Region was blank, filled as Unknown | 25 |
| `Warning - Missing Ship Mode` | Ship mode was blank, filled as Unknown | 21 |
| `Warning - Failed Payment` | Payment status is Failed | 69 |
| `Warning - Cancelled Order` | Order status is Cancelled | 145 |
| `DUPLICATE - REVIEW` | Same order_id with conflicting data in other columns | 24 |
| `Clean` | Passed all business rules | 797 |

>  Exact counts are documented in `data_quality_report.xlsx`.

---

## 7. Limitations of the Cleaning Process

1. **Single-flag per record:** The `data_quality_flag` formula uses nested IF logic and captures only the first issue found per row. Records with multiple issues (e.g. negative discount AND missing region) will show only one flag. A concatenated multi-flag approach would provide full issue coverage.

2. **Date format ambiguity:** Some dates like `03/12/2024` could be March 12 (MM/DD) or December 3 (DD/MM). Disambiguation relied on surrounding records and context. A small number of edge cases may have been incorrectly parsed.

3. **No external reference data for region:** Missing `region` values were filled as `Unknown` because no state-to-region lookup table was available. If such a mapping existed, VLOOKUP could assign correct regions automatically.

4. **Conflicting duplicate order_ids unresolved:** Flagged records with the same `order_id` but different values are kept for review but not resolved. The correct record cannot be determined without access to the originating source system.

5. **Sales mismatch does not guarantee error:** Some `calculated_sales` vs raw `sales` differences may result from legitimate adjustments (partial fulfilment, rounding, returns) not captured in the dataset. Mismatches are flagged but not automatically corrected.

6. **Refunded order ambiguity:** Some records have `payment_status = Refunded` but `order_status = Completed` or `Cancelled`. The exact business meaning of these combinations is unclear and has been flagged rather than corrected.

7. **No native change history in Excel:** Excel does not maintain a row-level change log automatically. This cleaning log is the only record of what was changed and why. A version control system (e.g. Git) would provide stronger auditability in a production data pipeline.



