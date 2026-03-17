# Client-Data-Transformation-Import-Readiness-Framework-45K-Records

## About project

Provide an overview of the project’s goals and context

This project simulates a real-world client onboarding scenario where a large, messy dataset must be transformed into a structured, validated, import-ready format.

The dataset contained **45,000 project-level records** with significant data integrity issues including duplicates, inconsistent naming, mixed formats, and numeric contamination.

The objective was to design and execute a systematic data transformation workflow using advanced Excel techniques and structured QA controls.

## Objectives

- Transform 45,000 messy records into a clean dataset
- Remove duplicate Project IDs
- Standardize inconsistent vendor names
- Clean numeric financial fields
- Normalize mixed date formats
- Validate data integrity before import
- Build a QA framework for ongoing verification

## Raw Data Characteristics

The dataset included:

| Column | Issues Identified |
| --- | --- |
| Project ID | Heavy duplication (~70%) |
| Vendor Name | 36 inconsistent variations (casing, spacing, misspellings) |
| Budget | Text contamination ("TBD"), blanks |
| Actual Cost | Text contamination ("N/A"), blanks |
| Start Date | Mixed formats |
| End Date | Mixed formats |
| Status | Missing values |
| Notes | Free-text inconsistencies |

### Raw Data Sample:
<img width="941" height="706" alt="Screenshot 2026-02-17 211643" src="https://github.com/user-attachments/assets/798f9f7d-a069-460e-a6a7-3e7589f9573e" />


## Phase 1: Data Profiling (Pre-Cleaning)

Before transformation, a dedicated profiling sheet was built to quantify data quality issues.

### Initial Findings

- Total Records: **45,000**
- Unique Project IDs: **13,417**
- Duplicate Rows: **31,583** (Total Records - Unique Project IDs)
- Budget Blanks: **6,654**
- Actual Cost Blanks: **6,654**
- Non-Numeric Budget Values: **3,476**
- Non-Numeric Actual Cost Values: **3,409**

### Interpretation

- ~70% structural duplication
- Over half of financial data unusable
- Mixed format corruption across date fields

This confirmed the dataset was not import-ready.

<img width="396" height="205" alt="Screenshot 2026-02-20 015724" src="https://github.com/user-attachments/assets/49dfcee1-9a94-47af-8102-185b6aaf8ea1" />


## Phase 2: Deduplication Strategy

### Step 1: Structural Deduplication

- Removed duplicate rows based on **Project ID**
- Business rule: Keep first occurrence

Steps:

1. Click inside the table
2. Data → Remove Duplicates
3. Click “Unselect All”
4. Check only:
    
    ✔ Project ID
    
5. Click OK

Result:

- 31,583 duplicate Project IDs removed
- Dataset reduced from 45,000 to 13,417 records

### Results

- Records reduced from **45,000 → 13,417**
- Duplicate Rows: **0**
- Duplicate Flags: **0**

Impact:

- Reduced dataset size by ~70%
- Improved processing efficiency
- Stabilized dataset structure

---

## Phase 3: Vendor Name Standardization

### Problem

36 inconsistent vendor variations representing 10 actual vendors.

### Solution

1. Extract unique vendor list
2. Create mapping table (`tbl_vendor_map`)
3. Standardize via XLOOKUP:

```
=IFERROR(
XLOOKUP([@[Vendor Name]],
tbl_vendor_map[Raw Vendor],
tbl_vendor_map[Standard Vendor]),
"Unmapped"
)
```

<img width="1369" height="755" alt="Screenshot 2026-02-18 091511" src="https://github.com/user-attachments/assets/f4071829-6ad4-4eec-98fa-2dcd737e7dd6" />


<img width="322" height="502" alt="Screenshot 2026-02-19 222101" src="https://github.com/user-attachments/assets/874111eb-2ac6-405a-a236-10eab6091956" />

### Validation

Unmapped Vendors: **0**

Result:

- All vendors standardized to consistent naming
- Reporting accuracy improved

---

## Phase 4: Financial Data Normalization

### Created Clean Columns (Preserved Raw Data)

```
=IFERROR(VALUE([@Budget]),"")
=IFERROR(VALUE([@[Actual Cost]]),"")
```

### Results

| Metric | Count |
| --- | --- |
| Valid Budget Values | 9,941 |
| Budget Blanks | 3,476 |
| Valid Actual Cost | 10,008 |
| Actual Cost Blanks | 3,409 |

All numeric contamination removed successfully.

---

## Phase 5: Variance & Budget Control Logic

### Budget Variance

```
=IF(OR([@[Budget Clean]]="",[@[Actual Cost Clean]]=""),"",
[@[Budget Clean]]-[@[Actual Cost Clean]])
```

### Over Budget Flag

```
=IF([@[Budget Variance]]="","",
IF([@[Budget Variance]]<0,"Over Budget","Within Budget"))
```

### Results

- Over Budget Projects: **2,135**
- Within Budget Projects: **5,222**
- Projects With Variance: **7,357**

Integrity check confirmed:

Over Budget + Within Budget = Total Variance Rows

---

## Phase 6: Date Normalization & Validation

### Clean Date Columns

```
=IFERROR(DATEVALUE([@[Start Date]]),"")
=IFERROR(DATEVALUE([@[End Date]]),"")
```

Standardized format: `yyyy-mm-dd`

### Date Validation

```
=IF(OR([@[Start Date Clean]]="",[@[End Date Clean]]=""),"",
IF([@[End Date Clean]]<[@[Start Date Clean]],
"Invalid Date Range","OK"))
```

### Results

- Invalid Date Ranges: **0**
- Clean Date Blanks: **0**
- All Records Validated

### Phase 7: Data Validation & QA Controls

Implemented:

- Status dropdown validation
- Vendor dropdown validation
- Conditional formatting for:
    - Blank financial values
    - Over Budget flags
    - Error conditions

Built final QA Summary sheet with reconciliation checks:

- Total Records
- Missing Counts
- Duplicate Check
- Variance Breakdown Check
- Date Validation Check
- Vendor Mapping Check

<img width="790" height="507" alt="Screenshot 2026-02-19 230819" src="https://github.com/user-attachments/assets/12dfe13c-b5a1-4fe9-9573-f75d36e2fb47" />


All integrity checks returned valid results.

---

### Conditional Formatting using new set of rules:

<img width="1117" height="502" alt="Screenshot 2026-02-19 222751" src="https://github.com/user-attachments/assets/0c204c0c-f91a-4c6e-8ede-8f3baf14a10f" />


---

Conditional formatting was applied to visually surface data quality issues and financial risk indicators within the cleaned dataset.

The following rules were implemented:

- **Budget Clean blanks** highlighted to quickly identify missing financial inputs.
- **Actual Cost Clean blanks** highlighted to expose incomplete cost records.
- **Over Budget Flag = "Over Budget"** highlighted to immediately surface projects exceeding budget.

The formatting rules were applied using relative cell references aligned with the first row of the selected range to ensure accurate row-level evaluation.

### Result

- Missing financial data is instantly visible.
- Over-budget projects are clearly flagged.
- Data quality issues can be identified without running additional formulas.
- The sheet functions as both a working model and a live QA dashboard.

This approach improves speed of review, reduces manual error detection time, and supports operational decision-making under scale.

## Final Dataset Summary

| Metric | Value |
| --- | --- |
| Clean Records | 13,417 |
| Duplicate IDs | 0 |
| Unmapped Vendors | 0 |
| Invalid Dates | 0 |
| Financial Contamination | Eliminated |
| QA Reconciliation | Verified |

Dataset is fully structured and import-ready.

---

## Python Automation (pandas)

Core logic replicated in Python for scalability:

```
import pandas as pd

df=pd.read_excel("raw_client_data.xlsx")

df=df.drop_duplicates(subset="Project ID")
df["Budget Clean"]=pd.to_numeric(df["Budget"],errors="coerce")
df["Actual Cost Clean"]=pd.to_numeric(df["Actual Cost"],errors="coerce")
df["Budget Variance"]=df["Budget Clean"]-df["Actual Cost Clean"]

df["Over Budget Flag"]=df["Budget Variance"].apply(
lambda x:"Over Budget"ifpd.notnull(x)andx<0else"Within Budget"
)

df.to_excel("clean_import_ready.xlsx",index=False)
```

Demonstrates automation readiness and scalability beyond Excel.

---

## Skills Demonstrated

- Advanced Excel modeling (XLOOKUP, structured tables, nested logic)
- Large-scale deduplication strategy
- Numeric contamination removal
- Data validation framework design
- QA reconciliation modeling
- Conditional formatting diagnostics
- pandas-based automation
- Structured operational workflow design

---

## Business Relevance

This project mirrors real-world client onboarding:

- Intake messy multi-format data
- Impose structure rapidly
- Validate integrity under scale
- Deliver QA-verified import-ready datasets

The workflow reflects ownership, structured thinking, and operational readiness required in high-growth data operations environments.

## **DATA:**

- Raw data file and the Final results files are available at: GitHub
