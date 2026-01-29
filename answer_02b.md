# Concrete example of replacement of Pivot tables

Good catch — you’re absolutely right to call that out 👍
From now on I’ll **strictly assume**:

* **EU locale**

  * List separator: **`;`**
  * Decimal separator: **`,`**
* **Excel 365 for Mac**

  * No Power Pivot / Data Model
  * Modern functions OK (`LET`, `LAMBDA`, `XLOOKUP`, dynamic arrays)
* You know Excel well, but I’ll **spell out advanced patterns clearly**

Below is a **concrete, realistic example** of what the **`Calc_Yearly` sheet looks like once filled**, including **values and formulas**, exactly as it would exist in Excel (EU syntax).

---

# Example: `Calc_Yearly` sheet (FILLED)

This sheet is **pure calculation**.
No manual inputs.
Everything flows from `tblTransactions` (+ `tblAmortization`).

---

## 1. Overall layout (row-oriented, year-driven)

Unlike your old `OLD-Stats` (years in columns), this is **normalized**:

```
A        B            C              D                E               F              G               H
----------------------------------------------------------------------------------------------------------
Year | FlightHours | VariableCost | FixedRecurring | FixedNonRec | CapitalAmort | TotalCostHT | CostPerHour
```

Each **row = one year**
Each **column = one metric**

This is the single biggest engineering improvement.

---

## 2. Filled example (with plausible values)

| Year | FlightHours | VariableCost (€) | FixedRecurring (€) | FixedNonRec (€) | CapitalAmort (€) | TotalCostHT (€) | CostPerHour (€/h) |
| ---: | ----------: | ---------------: | -----------------: | --------------: | ---------------: | --------------: | ----------------: |
| 2019 |        82,5 |           12 430 |             18 900 |           3 200 |            4 500 |          39 030 |            472,48 |
| 2020 |        61,2 |            9 850 |             19 300 |               0 |            4 500 |          33 650 |            549,67 |
| 2021 |        97,8 |           14 120 |             19 800 |           6 400 |            4 500 |          44 820 |            458,39 |
| 2022 |       110,4 |           16 980 |             20 100 |           1 200 |            4 500 |          42 780 |            387,39 |
| 2023 |       103,6 |           15 740 |             20 500 |               0 |            4 500 |          40 740 |            393,24 |
| 2024 |       118,9 |           18 320 |             21 000 |           9 800 |            4 500 |          53 620 |            451,11 |

(2025 would appear automatically once data exists)

---

## 3. Exact formulas (EU locale, Excel for Mac)

### Column A — `Year`

**A2 (spills down automatically):**

```excel
=LET(
  yrs; SORT(UNIQUE(tblTransactions[Year]));
  FILTER(yrs; yrs<>"")
)
```

---

### Column B — `FlightHours`

(assuming `tblTransactions[FlightHours]` is already populated)

**B2:**

```excel
=SUMIFS(
  tblTransactions[FlightHours];
  tblTransactions[Year]; $A2
)
```

---

### Column C — `VariableCost`

(categories classified as `Type = "Variable"`)

**C2:**

```excel
=SUMIFS(
  tblTransactions[AmountHT];
  tblTransactions[Year]; $A2;
  tblTransactions[Type]; "Variable"
)
```

---

### Column D — `FixedRecurring`

**D2:**

```excel
=SUMIFS(
  tblTransactions[AmountHT];
  tblTransactions[Year]; $A2;
  tblTransactions[Type]; "FixedRecurring"
)
```

---

### Column E — `FixedNonRec`

**E2:**

```excel
=SUMIFS(
  tblTransactions[AmountHT];
  tblTransactions[Year]; $A2;
  tblTransactions[Type]; "FixedNonRecurring"
)
```

---

### Column F — `CapitalAmort`

(coming from `tblAmortization`, **not** from transactions)

Assume `tblAmortization` has:

* `StartYear`
* `EndYear`
* `Amount`
* `YearsToAmortize`

**F2:**

```excel
=SUMPRODUCT(
  (tblAmortization[StartYear] <= $A2) *
  (tblAmortization[EndYear] >= $A2) *
  (tblAmortization[Amount] / tblAmortization[YearsToAmortize])
)
```

✔ Works perfectly on Excel for Mac
✔ No helper columns needed
✔ Deterministic & auditable

---

### Column G — `TotalCostHT`

**G2:**

```excel
=LET(
  var; C2;
  fix; D2;
  nonrec; E2;
  cap; F2;
  var + fix + nonrec + cap
)
```

Readable, debuggable, self-documenting.

---

### Column H — `CostPerHour`

**H2:**

```excel
=IF(
  B2>0;
  G2 / B2;
  NA()
)
```

Using `NA()` avoids misleading zeroes in charts.

---

## 4. Optional but HIGHLY recommended: named formulas

To make this even clearer, define these **once** in Name Manager:

### `GetYearAmountByType`

```excel
=LAMBDA(y; t;
  SUMIFS(
    tblTransactions[AmountHT];
    tblTransactions[Year]; y;
    tblTransactions[Type]; t
  )
)
```

Then Column C becomes:

```excel
=GetYearAmountByType($A2; "Variable")
```

Column D:

```excel
=GetYearAmountByType($A2; "FixedRecurring")
```

This is **professional-grade Excel** and still 100% Mac-compatible.

---

## 5. Why this is strictly better than `OLD-Stats`

| Old `OLD-Stats`              | New `Calc_Yearly`           |
| ---------------------------- | --------------------------- |
| Years in columns             | Years in rows               |
| Manual extension             | Fully automatic             |
| Layout-dependent             | Layout-independent          |
| Hard to audit                | Transparent formulas        |
| KPI rows tied to row numbers | KPIs refer to named columns |
| Copy-paste errors likely     | Zero copy-paste             |

Adding **2026** now means:

> just add transactions with Year = 2026
> → everything updates automatically

No column copying. No broken KPIs.

---

## 6. Next logical steps (your choice)

We can now:

1. 🔧 Build the **exact Excel file** with this sheet populated from your real data
2. 📊 Add the **Dashboard** (cost/hour trends, fuel vs hours, fixed vs variable)
3. 🧪 Add an **Audit sheet** that proves parity with `OLD-Stats`
4. 🧠 Introduce **provisioning logic** (monthly reserve targets, smoothing)

Tell me what you want next — and good job catching the locale issue, that’s exactly the kind of thing that bites later if ignored.
