Got it 👍 I’ll keep that mindset for the whole session: **Excel 365 for Mac, no Power Pivot / Data Model, best-practice worksheet engineering, and explanations that assume solid Excel basics but not advanced design patterns.**

Below is a **structured, detailed description of the uploaded workbook**, written explicitly so it can serve as **functional and technical requirements** for rebuilding it in a cleaner, more robust, and more maintainable way.

---

## 1. Overall Purpose of the Workbook

The workbook is a **long-term cost and performance tracking model** for a single aircraft (*Mooney LX-JCO*), covering **2013–2025**.

It has three logical layers:

1. **Raw transactional data** (expenses, quantities, hours)
2. **Yearly aggregations and statistics** (by cost type and year)
3. **High-level KPIs** derived from those statistics

The current implementation works, but:

* Logic is **hard-coded into layouts**
* Year handling is **manual and brittle**
* Calculations are **spread across many ad-hoc formulas**
* Readability and maintainability are low

---

## 2. Sheet: `OLD-Data` (Raw Data Layer)

### 2.1 Role of the Sheet

This is the **single source of truth** for all calculations downstream.

Each row represents **one financial or operational event** (expense, fuel purchase, labor, etc.).

### 2.2 Structure (Tabular Data)

| Column | Name        | Meaning                                    |
| -----: | ----------- | ------------------------------------------ |
|      A | Date        | Transaction date                           |
|      B | Year        | Calendar year extracted from Date          |
|      C | Description | Free-text description                      |
|      D | Category    | Cost category (used implicitly downstream) |
|      E | AmountHT    | Net amount (excl. VAT)                     |
|      F | TVARate     | VAT rate (decimal, e.g. 0.2)               |
|      G | VATAmount   | Calculated VAT                             |
|      H | AmountTTC   | Gross amount                               |
|      I | AVGASLiters | Fuel volume (liters)                       |
|      J | man-h       | Man-hours (labor)                          |

### 2.3 Calculations

* **Year**

  ```excel
  =YEAR(Date)
  ```

* **VATAmount**

  ```excel
  =AmountHT * TVARate
  ```

* **AmountTTC**

  ```excel
  =AmountHT + VATAmount
  ```

### 2.4 Observations / Issues

* The sheet is *almost* a proper Excel Table, but:

  * Not formally defined as one
  * Calculations are copied row-by-row instead of column formulas
* Categories are free text → **no validation**
* No explicit separation between:

  * Variable costs
  * Fixed recurring costs
  * Fixed non-recurring costs

### 2.5 Functional Requirements (Rebuild)

* Convert to a **single Excel Table** (`tblTransactions`)
* Use **structured references**
* Add:

  * Data validation for Category
  * Named VAT rates if applicable
* No formulas outside calculated columns
* This sheet must contain **zero aggregations**

---

## 3. Sheet: `OLD-Stats` (Yearly Aggregation & Cost Model)

### 3.1 Role of the Sheet

This is the **core analytical engine** of the workbook.

It:

* Aggregates raw data **by year**
* Splits costs into:

  * Variable recurring
  * Fixed recurring
  * Fixed non-recurring
* Computes:

  * Annual totals
  * Hourly rates
  * Long-term averages
  * Provisions

### 3.2 Layout Logic

* **Columns C → O** represent **years (2013–2025)**
* **Rows** represent:

  * Inputs (Hobbs hours, flight hours)
  * Cost categories
  * Intermediate subtotals
  * Derived rates

This is effectively a **manual pivot table implemented with formulas**.

### 3.3 Major Sections (Logical, Not Visual)

#### A. Flight Activity

* Hobbs start / end
* Flight hours per year
* Cumulative flight hours since 2013

#### B. Variable Costs

Includes:

* Oil & consumables
* 50h / 100h visits
* Fuel & taxes
* Landing taxes
* AVGAS volume and cost

Derived metrics:

* Total variable cost per year
* Variable cost per flight hour
* Variable cost excluding fuel

#### C. Fixed Recurring Costs

Includes:

* Insurance
* Annual inspections
* Avionics
* Hangar
* Regulatory fees
* Scheduled maintenance
* Management / manpower (man-hours)

Derived metrics:

* Annual fixed cost
* Fixed cost per flight hour
* Long-term average since 2013
* Monthly provisioning recommendation

#### D. Fixed Non-Recurring Costs

Includes:

* Oxygen bottle changes
* Amortized capital expenses
* Irregular maintenance events

Derived metrics:

* Annual totals
* Hourly equivalent charges

### 3.4 Calculation Style

* Extensive use of:

  * Row-based summations
  * Cross-row dependencies
  * Hard-coded cell references
* Year columns are **manually extended**
* No abstraction of:

  * “Year”
  * “Cost type”
  * “Aggregation rule”

### 3.5 Observations / Issues

* Extremely **layout-dependent**
* Adding a new year requires:

  * Copying columns
  * Verifying dozens of formulas
* Difficult to audit
* Logic and presentation are tightly coupled

### 3.6 Functional Requirements (Rebuild)

* Replace with:

  * Formula-driven yearly aggregation from `tblTransactions`
* Use:

  * `SUMIFS`
  * Dynamic year lists
  * Named ranges for year vectors
* One calculation rule = one clearly labeled block
* No duplicated logic per year

---

## 4. Sheet: `OLD-KPI` (High-Level Indicators)

### 4.1 Role of the Sheet

This sheet summarizes **key performance indicators** by summing across years.

It answers questions like:

* Total variable costs over period
* Total fixed costs over period
* Fuel consumption totals
* Cumulative averages

### 4.2 Structure

* Single-column KPI values
* Each KPI is:

  ```excel
  =SUM('OLD-Stats'!Cxx:Oxx)
  ```

### 4.3 Observations / Issues

* KPIs are:

  * Hard-wired to specific rows
  * Indirectly dependent on layout of `OLD-Stats`
* Very fragile if `OLD-Stats` changes

### 4.4 Functional Requirements (Rebuild)

* KPIs must:

  * Reference **named calculations**, not cell ranges
  * Be independent of year count
* Prefer:

  * `SUMIFS` over raw data
  * Or summation over dynamic helper ranges

---

## 5. Cross-Sheet Dependencies (Current)

```
OLD-Data  →  OLD-Stats  →  OLD-KPI
```

Problems:

* No abstraction layer
* Layout changes cascade into breakage
* No documentation of assumptions

---

## 6. Target Architecture (Implicit Requirements)

When rebuilt, the workbook should have:

1. **Data Layer**

   * One clean transaction table
2. **Calculation Layer**

   * Clearly named intermediate calculations
   * Year-driven logic (not column-driven)
3. **Presentation Layer**

   * KPIs and summaries that reference named results

All formulas should be:

* Readable
* Auditable
* Commented (cell comments or adjacent documentation)

---

## 7. What I Can Do Next (Suggested Next Steps)

Tell me how you’d like to proceed:

1. **Propose a clean target architecture** (sheet-by-sheet)
2. **Define a normalized category model** (variable vs fixed)
3. **Design the replacement for `OLD-Stats`**
4. **Refactor everything using Tables + named formulas**
5. **Rebuild step-by-step with explanations**

If you want, we can treat this like a **professional Excel refactoring project**, not just a rewrite.
