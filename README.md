# Retail POS Sales Analysis — Excel Capstone (FY 2024-25)

Cleaning and analysing one fiscal year of point-of-sale transactions from a retail
business, starting from a raw Tally export and ending with a validated dataset and
dashboard.

The headline of this project is not the dashboard — it is the **data quality audit**.
The source file looks like a clean table and is not one. Roughly 92% of its date
values and 91% of its period labels are unusable. Finding that, proving it, and
documenting what can and cannot be concluded from the data is the analytical work.

---

## Dataset

| Property | Value |
|---|---|
| Source | Tally POS export (`Sales_Transaction.xlsx`) |
| Raw rows | 3,413 |
| Valid transactions after cleaning | 3,387 |
| Voucher series | TSP/001\24-25 to TSP/3388\24-25 |
| Voucher type | TSP POS (single type throughout) |
| Period | FY 2024-25 (April 2024 – March 2025) |
| Total revenue | ₹2,92,30,332.71 |
| Columns | Date, TxnDate, Vch Type, Vch No., Amount, Month, Year, Quarter |

---

## Data quality findings

Each issue below was identified in the raw file and handled during cleaning.

**1. Stacked exports, not a single table**
The file is a set of monthly exports pasted end to end. Header rows ("Vch Type",
"Amount", "Debit") appear *inside* the data, and 10 block subtotal rows sit in the
Amount column alongside real transactions — values from ₹17.4 lakh to ₹40.3 lakh
that would otherwise be read as individual sales.

**2. The Date column is broken after row 266**
Vouchers TSP/001–266 carry genuine transaction dates spread evenly across the
fiscal year (16–26 transactions per month). Every row from TSP/267 onward is
stamped `2025-03-31`, which is a fill-down artifact rather than a real date.
**Only 7.9% of transactions have a trustworthy date.**

**3. Month / Year / Quarter columns are unreliable**
These derived columns disagree with the Date column in 3,075 of 3,389 rows (91%).
Transactions dated April 2024 are labelled "January 2024", and the labels scramble
entirely in later blocks. The Year column contains 2026, which is outside the
fiscal year covered. The Quarter column applies calendar-quarter logic to Indian
fiscal-year data, so it is wrong on its own terms as well.

**4. Amount stored as text**
The column mixes numbers with the strings "Debit" and "Amount", plus 14 blanks,
forcing a type coercion step before any aggregation.

**5. TxnDate is fully redundant**
`Date` and `TxnDate` are identical in all 3,413 rows. Dropped.

**6. Duplicate and blank voucher numbers**
25 duplicated values in `Vch No.`, concentrated in the 24 structural blank rows
that separate export blocks.

---

## Validation

The cleaned dataset is provably complete. The 10 subtotal rows extracted from the
raw file sum to **₹2,92,30,332.71**, which matches the sum of the 3,387 cleaned
transaction rows to the paisa. No revenue was lost or double-counted during
cleaning.

---

## Analysis and findings

### Transaction value distribution (all 3,387 transactions)

| Metric | Value |
|---|---|
| Mean | ₹8,630 |
| Median | ₹7,380 |
| Minimum | ₹400 |
| 90th percentile | ₹15,172 |
| 99th percentile | ₹32,814 |
| Maximum | ₹47,320 |

Mean sits above median, indicating a right-skewed distribution: a small number of
high-value sales pull the average up while typical sales cluster around ₹7,000.
Median is therefore the more honest measure of a "normal" transaction here.

### Time patterns (verified subset only, n = 266)

Monthly transaction counts in the date-verified subset are stable, ranging from 16
in July 2024 to 26 in April, September and February. Because this subset is only
8% of the data and was not randomly sampled, these patterns are reported as
indicative rather than conclusive.

### What this dataset cannot answer

The export contains no product, category, customer, quantity, cost, salesperson or
location fields. Product mix, margin analysis, customer segmentation (RFM), basket
analysis and regional comparison are all out of scope — not by choice, but because
the underlying columns do not exist. Full-year seasonality and forecasting are
similarly out of reach until a source export with intact dates is obtained.

Stating these limits is deliberate. Inventing columns to fill a dashboard template
would produce a more colourful project and a less truthful one.

---

## Repository structure

```
├── data/
│   ├── raw/Sales_Transaction.xlsx        # Original untouched export
│   └── clean/Sales_Clean.xlsx            # Validated transaction table
├── docs/
│   ├── data_quality_log.md               # Issue-by-issue cleaning record
│   └── findings.md                       # Full written analysis
├── dashboard/Sales_Dashboard.xlsx        # Pivot tables and charts
└── README.md
```

---

## Tools and techniques

- **Power Query** — removing embedded headers and subtotal rows, type coercion,
  column removal, structural blank handling
- **Excel formulas** — `XLOOKUP`, `SUMIFS`, `COUNTIFS`, `IFERROR`, `PERCENTILE.INC`
  for validation checks and summary metrics
- **PivotTables and PivotCharts** — distribution and period summaries
- **Slicers** — dashboard interactivity
- **Reconciliation testing** — subtotal-to-detail tie-out as a correctness proof

---

## How to use

1. Open `data/raw/Sales_Transaction.xlsx` to see the source in its original state.
2. Open `dashboard/Sales_Dashboard.xlsx` and refresh all queries
   (**Data → Refresh All**) to re-run the cleaning pipeline end to end.
3. Read `docs/data_quality_log.md` for the reasoning behind each cleaning decision.

---

## Next steps

- Obtain a re-export from the source Tally system with the date field intact, which
  would unlock full-year seasonality, day-of-week staffing analysis and forecasting
- Request item-level detail to enable product and margin analysis
- Add a rolling reconciliation check so future monthly exports validate automatically

---

## Author

**Dharshini **
Excel Data Analytics Capstone Project

