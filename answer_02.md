## Thoughtful, detailed description & requirements (for rebuilding the workbook)

Nice — I’ll treat this as my internal “thinking” about the file and lay out a clear, implementable specification you can use to rebuild the workbook on **Excel 365 for Mac** (no Power Pivot / Data Model). Below I give a precise, non-layout-dependent description of the file contents, the data model we’ll build instead, the exact columns / tables / named formulas, and example formulas and checks. This is written as requirements and implementation notes you (or I) can directly follow to produce a cleaner workbook with the same functionality.

---

## Executive summary (one-sentence)

The workbook is a **transaction-level aircraft cost & usage tracker (Mooney LX-JCO)** for 2013–2025 that aggregates transactions into yearly variable/fixed costs, computes per-hour rates, long-term averages and provisioning/amortization schedules — but is currently layout-bound and brittle; rebuild it as a single transactions table + lookup tables + calculation sheet + KPI/dashboard, using structured tables, named formulas, `SUMIFS`/`XLOOKUP`/`LET`/`LAMBDA`, and strong validation.

---

## File inventory (logical sheets the uploaded file contains or implies)

1. `OLD-Data` — raw transactions (date, description, category, amounts, fuel liters, man-hours, VAT, etc.)
2. `OLD-Stats` — year-by-year calculations (2013–2025 across columns) that split costs into variable/fixed/one-off and compute per-hour rates
3. `OLD-KPI` — single-column KPIs that sum or reference rows in `OLD-Stats`
4. (implicit) category lists / lookup tables and assumptions (VAT rates, category classification)
5. (implicit) some audit checks and cumulative Hobbs/flight hours tracking

---

## High-level design goals for the rebuild

* **Single source of truth:** 1 structured Excel Table `tblTransactions` (no duplicate row-level calculations elsewhere).
* **Separation of concerns:** Data layer (tables), Calculation layer (formulas, named formulas), Presentation layer (KPIs + Dashboard).
* **Year-driven, not column-driven:** produce dynamic year lists using `UNIQUE`/`SORT`/`SEQUENCE` and `SUMIFS`, so adding a new year (new transactions) requires zero manual formula edits.
* **Human-readable names & comments:** named ranges and named formulas (e.g. `GetYearCategoryAmount`) and a `README`/`Documentation` sheet describing every variable.
* **Excel 365 for Mac compatible:** avoid Power Pivot, Data Model, Windows-only features; use `LET`, `LAMBDA`, `XLOOKUP`, `SUMIFS`, structured references and implicit intersection-friendly formulas.
* **Robust validation & audit:** data validation for categories, VAT rates table, automated data quality checks.

---

## Detailed sheet-by-sheet specification (exact fields, constraints, and formulas)

### Sheet: `Data` (implementation name)

**Purpose:** raw transaction table (single canonical table).

Create an Excel Table named `tblTransactions`.

Columns (exact names — capitalization matters for structured refs):

* `Date` (Date) — transaction date. Validation: date only.
* `Year` (Number, calculated) — `=YEAR([@Date])` as a calculated column.
* `Source` (Text) — free text / supplier / note (optional).
* `Description` (Text) — transaction description.
* `Category` (Text) — validated drop-down from `tblCategories[Category]`.
* `SubCategory` (Text) — optional (e.g., "Oil", "50h", "Avionics").
* `Type` (Text) — validated: `Variable`, `FixedRecurring`, `FixedNonRecurring`, `Capital` (used for aggregation rules).
* `AmountHT` (Number) — net amount (ex-VAT).
* `TVARate` (Number) — VAT decimal (e.g. `0.20`). Default from `tblVATRates` via `XLOOKUP` in calculated column if blank.
* `VATAmount` (Number, calculated) — `=[@AmountHT]*[@TVARate]`
* `AmountTTC` (Number, calculated) — `=[@AmountHT]+[@VATAmount]`
* `AVGASLiters` (Number) — liters purchased (blank for non-fuel rows).
* `ManHours` (Number) — man-hours charged (blank if none).
* `HobbsStart` (Number) — optional, only used if transactions record Hobbs meter updates
* `HobbsEnd` (Number) — optional
* `FlightHours` (Number, calculated if Hobbs change recorded) — `=IF(AND([@[HobbsStart]]<>"",[@[HobbsEnd]]<>""),[@[HobbsEnd]]-[@[HobbsStart]],[@FlightHours])` — keep flight/hobbs entries consistent (or store flight/hobbs in separate table if preferred).
* `Notes` (Text) — free comment.

Implementation notes:

* Make `Category` a validated pick-list (see `tblCategories` sheet).
* Make `TVARate` auto-populate with `XLOOKUP([@SomeKey], tblVATRates[Key], tblVATRates[Rate], 0.0)` if user leaves blank.
* All column calculations use table-calculated columns (no copied formulas).

---

### Sheet: `Lookups`

Contain small tables:

1. `tblCategories` — columns: `Category`, `Type` (Variable/FixedRecurring/FixedNonRecurring/Capital), `CostGroup` (e.g., Fuel, Maintenance, Insurance), `DefaultTVAKey` (optional).
2. `tblVATRates` — `VATKey`, `Rate` (e.g. `FRStandard`, `0.20`).
3. `tblAmortization` — rows describing capital items: `Item`, `PurchaseDate`, `Amount`, `YearsToAmortize`.
4. `tblRates` — other configurable per-year or per-unit rates (e.g., hourly provisioning target, fuel price benchmark) if desired.

Use these tables for data validation and mapping.

---

### Sheet: `Calc_Yearly` (calculation layer)

**Purpose:** compute year-by-year aggregates from `tblTransactions` using formulas (not manual columns per year).

Structure:

* A left column `YearList` created dynamically:

  ```excel
  =LET(years,SORT(UNIQUE(tblTransactions[Year])), FILTER(years,years<>""))
  ```

  Put `YearList` vertically (one cell per year) using dynamic array or spilled formula.

* For each important aggregation we compute a column:

  * `Total_AmountHT` per year:

    ```excel
    =SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], $A2)
    ```
  * `Total_VAT`:

    ```excel
    =SUMIFS(tblTransactions[VATAmount], tblTransactions[Year], $A2)
    ```
  * `Total_TTC`:

    ```excel
    =SUMIFS(tblTransactions[AmountTTC], tblTransactions[Year], $A2)
    ```
  * `Fuel_Liters`:

    ```excel
    =SUMIFS(tblTransactions[AVGASLiters], tblTransactions[Year], $A2)
    ```
  * `ManHours`:

    ```excel
    =SUMIFS(tblTransactions[ManHours], tblTransactions[Year], $A2)
    ```
  * `Variable_Costs`:

    ```excel
    =SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], $A2, tblTransactions[Type], "Variable")
    ```
  * `FixedRecurring_Costs`:

    ```excel
    =SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], $A2, tblTransactions[Type], "FixedRecurring")
    ```
  * `FixedNonRecurring_Costs`:

    ```excel
    =SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], $A2, tblTransactions[Type], "FixedNonRecurring")
    ```
  * `Capital_Amortization` — computed by spreading rows in `tblAmortization` across amortization years. Implementation example below.

* Derived metrics per year (use `LET` to keep formulas readable):

  * `Total_Costs = Total_AmountHT + Total_VAT` (or `Total_TTC`)
  * `Variable_perFlightHour = Variable_Costs / FlightHours` (guard division by zero)
  * `Fixed_perFlightHour = FixedRecurring_Costs / FlightHours`
  * `Provision_Monthly = (FixedRecurring_Costs / 12)` or a smoothed provisioning approach.

**Amortization schedule (example implementation):**

* In `Calc_Yearly`, compute `Capital_Amortization` as sum over `tblAmortization` of the portion of each item that falls into the year.
* A simple method: in `tblAmortization` add `StartYear` and `EndYear` (StartYear = YEAR(PurchaseDate); EndYear = StartYear + YearsToAmortize - 1). Then:

  ```excel
  =SUMPRODUCT( (tblAmortization[StartYear] <= $A2) * (tblAmortization[EndYear] >= $A2) * (tblAmortization[Amount] / tblAmortization[YearsToAmortize]) )
  ```

  This spreads cost evenly year-by-year.

**Reusable named formula (recommended):**

* Name: `GetYearCategoryAmount`
* Definition (LAMBDA):

  ```
  =LAMBDA(year, cat, SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], year, tblTransactions[Category], cat))
  ```
* Use `GetYearCategoryAmount($A2, "Fuel")`.

---

### Sheet: `KPIs` (presentation)

**Purpose:** single place with human-friendly KPIs, each KPI referencing named formulas or `Calc_Yearly` cells, not raw `OLD-Stats` row numbers.

Suggested KPIs (examples):

* `TotalCost_AllYears`:

  ```excel
  =SUM(Calc_Yearly[Total_TTC])
  ```
* `AverageVariablePerHour`:

  ```excel
  =AVERAGE( FILTER( Calc_Yearly[Variable_perFlightHour], Calc_Yearly[FlightHours] > 0) )
  ```
* `FuelConsumed_Total`:

  ```excel
  =SUM(Calc_Yearly[Fuel_Liters])
  ```

All KPIs should reference named column headers or named ranges.

---

### Sheet: `Dashboard`

**Purpose:** visual summary for decision makers.

* Use sparklines, small charts, per-year bars for Total Cost, Variable Cost/Hour, Fuel liters.
* Add slicers only if they work on Tables (Excel for Mac supports slicers on Tables in recent versions — if not, use drop-downs & formulas).
* Ensure all visuals are driven by dynamic year list; charts refer to dynamic ranges.

---

### Sheet: `Audit` / `Checks`

**Purpose:** automated quality checks.

Examples (cells with clear PASS/FAIL):

* `MissingCategoryCount = COUNTBLANK(tblTransactions[Category])`
* `NegativeAmountCount = COUNTIFS(tblTransactions[AmountHT],"<0")`
* `VATMismatchCount = COUNTIFS(tblTransactions[VATAmount], "<>" , ROUND(tblTransactions[AmountHT]*tblTransactions[TVARate], 2))`
* `YearContinuityCheck` — compares year list against expected span (2013 to MAX Year) and flags missing years.
* `TransactionsNotAssignedType` = COUNTIFS(tblTransactions[Type], "")

Show detailed lists when checks fail using `FILTER` to display offending rows.

---

## Naming conventions, comments, and documentation

* Table names: `tblTransactions`, `tblCategories`, `tblVATRates`, `tblAmortization`, `tblRates`.
* Named formulas: `Years`, `GetYearTotal`, `GetYearCategoryAmount`, `GetTotalForType`.
* All named formulas must be short, descriptive, and recorded in `Documentation` sheet with:

  * Name, signature (parameters), description, example use.
* For any complex formula cell, add a cell comment (right-click → Insert Comment) explaining:

  * What it calculates
  * Inputs it depends on (table/columns)
  * Why it's written that way (e.g., amortization uses SUMPRODUCT to avoid helper columns)
* Keep a `README` sheet with:

  * Purpose of workbook
  * How to add a new Transaction
  * How to add a new Category / VAT rate
  * Release/version history (ChangeLog)

---

## Example formulas (copy-paste-ready) — Excel 365 for Mac compatible

1. **Dynamic Years list (single cell that spills down):**

```excel
=LET(yrs, SORT(UNIQUE(tblTransactions[Year])), FILTER(yrs, yrs<>""))
```

2. **Sum of AmountHT for category "Fuel" in year at A2:**

```excel
=SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], $A2, tblTransactions[Category], "Fuel")
```

3. **Variable costs for year (using Type column):**

```excel
=SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], $A2, tblTransactions[Type], "Variable")
```

4. **Named LAMBDA (define in Name Manager):**
   Name: `GetYearCategoryAmount`
   Refers to:

```excel
=LAMBDA(y,cat, SUMIFS(tblTransactions[AmountHT], tblTransactions[Year], y, tblTransactions[Category], cat))
```

Use:

```excel
=GetYearCategoryAmount($A2,"Avionics")
```

5. **Amortization contribution for year in cell with year in $A2 (polls tblAmortization):**

```excel
=SUMPRODUCT( (tblAmortization[StartYear] <= $A2) * (tblAmortization[EndYear] >= $A2) * (tblAmortization[Amount] / tblAmortization[YearsToAmortize]) )
```

6. **Guarded per-hour metric:**

```excel
=IF(Calc_Yearly[@FlightHours]>0, Calc_Yearly[@Variable_Costs] / Calc_Yearly[@FlightHours], NA())
```

---

## UX & practical recommendations

* Use **structured table references** everywhere (no whole-column references like A:A).
* Avoid volatile functions (e.g., `INDIRECT`, `OFFSET`) where possible — they slow recalculation.
* Keep raw transactions table on its own sheet and freeze header row.
* Use **Data Validation** (List) for `Category`, `Type`, `TVARateKey` to avoid typos.
* Protect the `Lookups`, `Calc_Yearly`, `KPIs` sheets (lock formulas), allow editing of `Data` only.
* Use `XLOOKUP` for mapping defaults (`XLOOKUP(lookup, lookup_array, return_array, 0)`).

---

## Migration plan (how to convert the old workbook to the new structure)

1. Copy `OLD-Data` rows into `tblTransactions` — map columns and correct types.
2. Create `tblCategories` from unique Category texts, classify them into `Type`.
3. Recompute `TVARate` using `tblVATRates`.
4. Populate `tblAmortization` from any capital/large purchase rows.
5. Populate `Calc_Yearly` dynamic years; validate totals against old `OLD-Stats` for a few years to confirm parity.
6. Build `KPIs` and `Dashboard`.
7. Run `Audit` checks and fix data problems.

---

## Edge cases & special handling

* **Free-text categories:** map them to normalized categories via `tblCategories` and consider creating a mapping table for historical data where values differ by spelling.
* **Multiple currencies:** if transactions have more than one currency, add a `Currency` column and conversion rates table; convert to base currency in `tblTransactions`.
* **Hobbs vs FlightHours:** prefer storing `FlightHours` explicitly per transaction to avoid ambiguity; if using Hobbs meter deltas, implement a small helper table `tblHobbs` with ordered meter events and compute deltas with `INDEX`/`LAG` style logic or a pivot on sorted events.
* **VAT rounding differences:** calculate VAT at row level and sum; if jurisdiction requires VAT on invoice-level rounding, add a column `VAT_RoundingRule` and store reported VAT to detect mismatches.

---

## What I will produce if you want me to implement it next

If you want, I can now:

1. Produce an exact, sheet-by-sheet Excel file (rebuild) implementing the design above and populate it with the data from the uploaded workbook (I already see the file path you uploaded).
2. Or produce the workbook skeleton (empty tables, named formulas, documentation sheet) that you can paste data into safely.

Say which one you prefer and I’ll proceed with the rebuild plan immediately (I already have everything required from the uploaded workbook to map fields).

---

Would you like the **full rebuilt workbook** (I’ll migrate data, build the amortization schedule, produce Dashboard + Audit) or start with a **skeleton template** and we’ll migrate stepwise?
