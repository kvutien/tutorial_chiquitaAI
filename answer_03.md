# Summary of progress and clarify eventual assumptions

*return to the main article here [(README.md)](./README.md)*

Perfect — thank you, this is *exactly* the level of clarification that turns this into a clean, professional refactor. 👍
I’ve carefully integrated **all** your points and constraints (EU locale, Excel 365 for Mac, no Data Model, SUMIFS-centric design, no amortization yet, multi-pilot balances, VAT reality in Europe, loop avoidance).

Below I’ll do **three things only**, as you requested:

1. **Acknowledge and restate the clarified requirements** (so you can confirm I understood correctly)
2. **Explicitly list the remaining assumptions / open points** that are *not* yet fully specified
3. **Ask you to confirm or correct those assumptions**, then stop — so that in the next prompts you can give me the cost-calculation rules

I will **not** jump ahead into formulas yet.

---

## 1. Acknowledgement & integrated understanding

I confirm and lock in the following design decisions:

### 1.1 Purpose of the refactored workbook

The workbook’s **primary objective** is now:

> **Calculate the financial balance of each pilot**, based on
>
> * shared aircraft ownership costs
> * individual usage (flights, Hobbs, fuel)
> * individual contributions (payments into the fund)

This means:

* Pilot identity is now a **first-class dimension**
* All costs ultimately roll up **per pilot**
* Yearly stats are still important, but **balances are the end goal**

---

### 1.2 Sheet architecture (locked)

#### Data entry sheets (editable)

* **Expenses**

  * General aircraft expenses
  * No Hobbs, no flight, no AVGAS
  * Has helper column `Type` = `Fixed | Variable | Exceptional`
  * Has `FundSource` (who paid)

* **AVGAS**

  * Fuel purchases only
  * Has helper columns for Stats
  * Has `FundSource`

* **Flights**

  * Pilot, date, locations, Hobbs, flight time, etc.
  * Contains **helper columns to calculate per-flight costs**
  * Has `FundSource` if relevant (e.g. who paid landing fee)

* **Contribs**

  * Money injected into the common fund by pilots
  * Has `FundSource` (= pilot)

✔ All data entry sheets include **FundSource** for traceability
✔ No sheet mixes concerns (expenses ≠ flights ≠ fuel)

---

#### Read-only sheets

* **Params**

  * Lookups & constants
  * Some fixed values (e.g. dry hourly cost per pilot)
  * Some *dynamic* parameters (averages since inception)

* **Stats** (formerly `Calc_Yearly`)

  * Core calculation engine
  * Uses `SUMIFS`, no PivotTables
  * Produces yearly and cumulative metrics
  * Metrics reused by Flights helper columns
  * Designed to **avoid circular references**

* **Overview**

  * KPIs + Dashboard merged
  * Read-only surface for humans
  * No calculations beyond light presentation

* **Audits**

  * Data integrity & sanity checks
  * Completeness, consistency, balance validation

* **Misc**

  * Temporary / scratch calculations
  * Not relied upon by core logic

* **README**

  * User guide
  * How to enter data
  * How balances are calculated (high-level)

---

### 1.3 VAT handling (important clarification acknowledged)

* **VAT rates vary by year and country**
* VAT **will not** be auto-derived via `XLOOKUP`
* VAT is:

  * Entered explicitly per row in Expenses / AVGAS
  * Calculated locally at row level
* Stats aggregates VAT, but does not attempt to infer it

✔ This keeps the model **realistic and auditable** in a European context

---

### 1.4 Parameters logic (acknowledged)

* Some Params are **static**:

  * Example: dry hourly cost per pilot
* Some Params are **dynamic**:

  * Fuel price
  * Fuel consumption per hour
  * Fuel cost per hour
* Dynamic Params are defined as:

  > **Averages since the beginning until the moment of the flight**
* These values are computed in **Stats** and *read* by Flights
* Flights does **not** recompute Stats logic → avoids loops

---

### 1.5 Explicitly ignored (for now)

* ❌ Amortization
* ❌ Capital depreciation
* ❌ Power Pivot / Data Model
* ❌ Pivot tables

These are consciously postponed.

---

## 2. Remaining assumptions / open points (need confirmation)

Below are the **only points I believe are still ambiguous**.
Please confirm or correct them — short answers are fine.

---

### A. Definition of “Variable cost” vs “Flight-specific cost”

You stated:

> Expenses need a helper column Type (Fixed, Variable, Exceptional)

Open point:

* Are **Variable expenses**:

  * Allocated to pilots **pro-rata flight hours**?
  * Or **only when flights occurred in that year**?
  * Or **assigned directly to flights when possible**?

➡️ Example:
An oil change — is that:

* spread across all pilots based on hours flown?
* or treated as a general pool cost?

---

### B. FundSource semantics

Each data entry sheet has `FundSource`.

Please confirm:

* `FundSource` is **who paid**, not who benefits
* Balances are computed as:

```
Balance per pilot =
  Contributions
+ Expenses paid by pilot
- Usage costs attributed to pilot
```

Is this correct?

---

### C. Flights helper columns – scope of responsibility

Flights sheet will calculate **per-row costs**.

Please confirm which costs are calculated **at flight level**:

* Dry hourly cost?
* Fuel cost (using Stats averages)?
* Landing fees (if entered elsewhere)?
* Variable maintenance provision?

Or should Flights **only compute usage metrics** (hours, fuel burned) and let Stats allocate money?

---

### D. Time dimension

Stats is still **year-based**.

Confirm:

* Balances are:

  * Calculated cumulatively (since beginning)?
  * And optionally viewable per year?

Or do you want **monthly** capability later?

---

### E. Exceptional expenses

How should `Type = Exceptional` be handled?

Options (you choose later):

* Spread like Fixed
* Charged immediately to fund
* Allocated manually to pilots
* Excluded from balance until decided

For now, confirm:
➡️ *Exceptional costs should be visible and isolated, not silently absorbed.*

---

## 3. Next step (once you confirm)

Once you confirm / correct the points above, in the **next prompts** we will:

1. Define **exact cost allocation rules** (math, not layout)
2. Design the **Stats sheet formulas** to support Flights without loops
3. Define **per-flight helper columns** (with EU-locale formulas)
4. Define **pilot balance computation**
5. Lock the architecture and only then think about visuals

Take your time — your clarifications so far are excellent and make this a solid system rather than “yet another Excel”.

👉 Please confirm or adjust points **A → E**, or tell me if I missed an assumption.
