# Zopper BI Internship Assignment — Car Insurance Analytics

## Overview
Simulation and analysis of 1,000,000 car insurance policies sold in 2024, with claims modeled for 2025–2026. Built using Python (pandas, numpy, openpyxl) with results delivered as a formatted Excel workbook and written report.

---

## Files Included

| File | Description |
|------|-------------|
| `Zopper_BI_Assignment.xlsx` | Main Excel workbook — dashboard, datasets, all analytical queries |
| `Zopper_BI_Approach_Report.docx` | Written report — methodology, assumptions, insights |
| `policy_sales_data.csv` | Full 1,000,000 policy records |
| `claims_data.csv` | Full 49,340 claim records |


---

## How to Run

### Requirements
```bash
pip install pandas numpy openpyxl
npm install -g docx
```

### Generate Data & Run Queries
```bash
python simulate.py
```

### Rebuild Excel Workbook
```bash
python build_excel.py
```

### Rebuild Word Report
```bash
node build_report.js
```

---

## Dataset Summary

### Policy Sales Data (Table 1)
- **1,000,000** policies sold Jan 1 – Dec 31, 2024
- Spread evenly across all 366 days (2024 is a leap year)
- Tenure distribution: 1yr (20%) / 2yr (30%) / 3yr (40%) / 4yr (10%)
- Vehicle value: ₹1,00,000 | Premium: ₹100 per year of tenure
- Policy starts 365 days after purchase date

### Claims Data (Table 2)
- **2025 claims:** 30% of vehicles purchased on 7th/14th/21st/28th of each month → claim filed on policy start date
- **2026 claims (Jan 1 – Feb 28):** 10% of all 4-year tenure policies, distributed evenly across 59 days
- Vehicles that claimed in 2025 and re-qualify in 2026 are marked `Claim_Type = 2`

---

## Analytical Query Results (Part 3)

| # | Question | Answer |
|---|----------|--------|
| Q1 | Total premium collected 2024 | ₹24,01,10,800 (~₹24 Cr) |
| Q2 | Total claim cost 2025 | ₹39,34,40,000 (~₹39.3 Cr) |
| Q2 | Total claim cost 2026 | ₹9,99,60,000 (~₹10 Cr) |
| Q3 | Best loss ratio tenure | 3-Year (1.31x) |
| Q3 | Worst loss ratio tenure | 1-Year (3.94x) |
| Q4 | Loss ratio by sale month | Consistent ~2.00x–2.14x across all months |
| Q5 | Potential claim liability | ₹9,51,02,30,000 (~₹951 Cr) |
| Q6a | Earned premium to Feb 28, 2026 | ₹6,59,02,532 |
| Q6b | Est. monthly premium (46 months) | ₹3,78,714 / month |

---

## Loss Ratio by Tenure

| Tenure | Total Premium (₹) | Total Claims (₹) | Loss Ratio |
|--------|-------------------|------------------|------------|
| 1 Year | 1,99,37,600 | 7,84,60,000 | 3.94x |
| 2 Years | 6,00,21,400 | 11,89,00,000 | 1.98x |
| 3 Years | 12,01,65,000 | 15,78,20,000 | **1.31x** ✅ |
| 4 Years | 3,99,86,800 | 13,82,20,000 | 3.46x |

---

## Key Insights (Part 4)

1. **Most profitable tenure: 3-Year** — lowest loss ratio (1.31x), highest premium volume, 40% of all sales. 1-year and 4-year policies are significantly under-priced.

2. **Overall portfolio loss ratio: 2.05x** — for every ₹1 in premium, ₹2.05 is paid out in claims. A healthy ratio is 60–80%. Current pricing is unsustainable.

3. **Claim trends:** 2025 claims spike on cohorts tied to special purchase dates (7/14/21/28). 2026 claims are smooth and evenly spread across Jan–Feb.

4. **+5% annual claim frequency:** Adds ~₹1.97 Cr/year to losses. By Year 5, the loss ratio would approach 2.5x without repricing.

5. **Potential liability of ₹951 Cr** dwarfs the ₹24 Cr premium collected — urgent need to build adequate reserves.

---

## Assumptions

- 2024 treated as a 366-day leap year
- "Evenly distributed" = integer floor per day, remainder assigned to earliest days
- 30% claim rule applied as random sample for realism (seeded for reproducibility)
- Daily premium = Total Premium ÷ (Tenure × 365) for earned premium calculations
- 46 months used as remaining policy period after Feb 28, 2026
- Policy active check enforced on all claims (claim date must be within start–end window)

---

## Tools Used
- **Python** — pandas, numpy, openpyxl
- **Node.js** — docx library
- **Excel** — output workbook with charts and formatted tables
