-	The purpose of re-factored worksheet will be used to calculate the balance of each pilot, based on the general cost of maintaining the plane, the costs of their usage of the plane and their contributions.
-	The re-factored worksheet will be composed of the data entry sheets and read-only sheets. 
-	Data entry sheets are composed of Expenses, AVGAS (fueling), Flights (pilot, time, location, Hobbs, etc.) and Contribs. All data entry sheets have a column FundSource to trace who is paying. 
-	Read-only sheets are composed of Params, Stats, Overview (Dashboard & KPI) and Audits. A last sheet Misc is for temporary calculations and a README sheet is the user guide.
-	Sheet Flights has data entry columns and helper columns to calculate costs related to the trip of the row. These helper columns will be read by sheet Stats to calculate the balances.
-	Sheet Expenses (Data) needs no Subcategories. But it will need a helper column Type (Fixed, Variable, Exceptional).
-	Sheet Expenses (Data) will never contain flight and Hobbs value, nor AVGAS. This information will be in a distinct sheet Flights, respectively AVGAS.
-	Sheet AVGAS has data entry columns and a few helper columns that will be used by Stats to calculate fuel data.
-	Sheet Expenses (Data): In Europe there are a lot of VAT rates depending on the year and on the country. Defining TVARate by XLOOKUP is overly complicated.
-	Sheet Params (Lookups): There are some parameters that can be used with XLOOKUP like the dry hourly cost for pilots. Some others such as fuel price, consumption per hour, fuel per hour, fuel cost per hour, will be the average since the beginning until the moment of the flight.
-	Sheet Params (Lookups): Ignore. There is currently no amortization yet. Only in future versions.
-	Sheet Stats (Calc_Yearly): Very interesting use of SUMIFS which improves on pivot tables. We keep this solution from now on. We will call this sheet Stats.
-	Sheet Stats (Calc_Yearly): Keep in mind that the SUMIFS giving average hourly consumption, fuel price, yearly variable costs, fixed costs fuel costs are reused to calculate the cost of each trip in Flights. We must avoid calculation loops.
-	Sheet Stats (Calc_Yearly): Let's ignore amortization for the time being. It will be introduced later after a few months using this Excel worksheet.
-	Sheets KPIs & Dashboard: Merge these 2 into a sheet named Overview. It is the read-only surface while Stats is the calculation engine.
-	Sheet Audits is a good contribution that ChatGPT "think mode" brought in the last answer.

Acknowledge that you'll take these clarifications into account. Ask me to confirm assumptions if you still have that I have not clarified. Then in the coming prompts I'll give you details on how to calculate costs in the calculation formulas.


- see ChatGPT answer here [(./answer_03.md)](./answer_03.md)
- return to the main article here [(README.md)](./README.md)