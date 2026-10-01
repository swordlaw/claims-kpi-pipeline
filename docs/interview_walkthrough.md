# Interview Walkthrough

A plain-language guide to how this pipeline works and why it's built this way. Each section ends with a short answer you could give out loud.

---

## 1. The 30-second version

> "I built a claims KPI reporting pipeline on public CMS Medicare data, about 860,000 claim rows. A Python loader puts the raw CSVs into SQL Server staging tables. T-SQL then cleans and types the data into a small reporting model of members, coverage and claims, merging claims that arrive split across several rows. I wrote data quality checks and documented every issue I found, with the record count and how I handled it. KPI views calculate monthly spend, PMPM, average length of stay and 30-day readmission rate. Those views feed a spreadsheet report with a Summary dashboard, and a VBA macro refreshes everything and exports a dated PDF in one click."

---

## 2. The data

- **CMS DE-SynPUF Sample 1**: synthetic Medicare data. It's built to look and behave like real CMS claims files, but no row is a real person. That makes it safe to publish and share.
- **Beneficiary files**, one per year: who the members are, plus how many months they were covered that year.
- **Inpatient claims**: hospital stays (admission date, discharge date, payment).
- **Outpatient claims**: hospital outpatient visits.
- **Key IDs:** `DESYNPUF_ID` identifies a member, and `CLM_ID` identifies a claim.

**Something I caught:** on cms.gov, the "Sample 1 2010" download link actually points to a file named Sample 20. Rather than assume, I tested it. Every member ID in it also appears in the 2008 and 2009 Sample 1 files, so it's the right data with the wrong file name. Mixing samples would have broken every join, because Sample 20 holds different synthetic people.

> "Before loading anything I verify the source. The 2010 file had the wrong name on the CMS site, so I proved it was Sample 1 by checking that 100% of its member IDs matched the other years."

---

## 3. Why staging tables?

The CSVs go into **staging tables** first (`stg_beneficiary_2008`, `stg_inpatient`, ...), where **every column is text** (`NVARCHAR(50)`), named exactly as in the file.

Why not load straight into typed tables?

1. **The load never fails on bad data.** If a date column held `"N/A"`, a typed load would crash and you wouldn't know how many rows were bad. Loading as text always succeeds.
2. **Cleaning happens in SQL, where it's visible and countable.** `TRY_CONVERT` turns bad values into NULL, and section 3 counts them. Every decision is in the script, which can be reviewed and re-run.
3. **You can always go back to the raw data.** If a cleaning rule turns out to be wrong, fix the T-SQL and re-run. No need to download or reload the files.
4. **Reconciliation.** Staging row count = file row count is the first check. Then staging vs. final table counts explain exactly what the cleaning did.

**A subtle detail:** empty fields are loaded as `NULL`, not empty strings. In SQL Server, converting an empty string to a date gives **1900-01-01**, not NULL. Every missing date would have quietly become a real-looking 1900 date and slipped past the "missing date" check. I tested this on the server to confirm it.

**Why Python for loading, not BULK INSERT?** BULK INSERT runs as the SQL Server service account, which can't read files in a user's folder without changing permissions. The Python loader reads each CSV's header to build the staging table, so the column names always match the file. It also prints file rows vs. loaded rows for each table.

> "Staging is a raw, all-text copy of the source, so the load never fails and I keep the original data. All type conversion happens afterwards in T-SQL, where I can count exactly what didn't convert."

---

## 4. The reporting model

Three clean tables:

| Table | One row per | What it's for |
|---|---|---|
| `member` | person | demographics (birth date, sex, state), taken from their most recent year |
| `member_coverage` | person × year | months of coverage: the denominator for PMPM |
| `claim` | claim | dates, provider, diagnosis, paid amount; `claim_type` = IP or OP |

Inpatient and outpatient share one `claim` table with a `claim_type` column, because most KPIs (volume, spend, PMPM) work the same for both. The inpatient-only fields (admit and discharge dates) are NULL for outpatient.

The primary keys enforce the grain. If two rows for the same claim ever got through, the insert would fail instead of silently double-counting money.

---

## 5. Why claims are merged across segments

In the source, a long claim can be split across several rows, called **segments** (segment 1, segment 2). I found:
- **11,043 claims** were split (68 inpatient, 10,975 outpatient).
- **Segment-2 rows have no dates**, but they **do carry payments**: 66 of the 68 inpatient multi-segment claims had money on segment 2.

If you treated each row as a claim, you'd:
- **overcount claims** (an extra 11,043 "claims"),
- **break per-claim averages**, since one claim's payment is split in two,
- **break readmissions**, since a segment-2 row with no dates can't be placed in time.

So the T-SQL groups by `CLM_ID` and:
- takes the **earliest from-date and latest thru-date**,
- **sums the paid amount**, so no money is lost,
- takes **diagnosis and provider from segment 1**, the claim header. I changed this from the original `MIN()`, because on 5,543 outpatient claims the segments listed different diagnoses, and `MIN()` just picks the alphabetically lowest code: arbitrary.

Result: 66,773 → 66,705 inpatient claims and 790,790 → 779,815 outpatient claims. The difference (68 and 10,975) is exactly the number of claims that had a second segment.

**Edge case:** 278 outpatient claims had *only* a segment 2, with no header row and therefore no dates. I kept them, so their payments count, and they're automatically excluded from date-based KPIs. The DQ check reports them as `missing_from_date = 278`.

> "One claim can be split over several rows in the source. I merge them to one row per claim and sum the payments, otherwise claim counts are inflated and per-claim averages are wrong. Staging minus final row counts reconcile exactly to the number of extra segment rows."

---

## 6. How readmissions are calculated

**Definition:** a stay counts as readmitted if the same member is admitted again **between the discharge date and 30 days after it**. The rate is the share of stays (the "index stays") followed by such a readmission.

**How it's done in SQL** (`v_inpatient_stays`):

```sql
LEAD(admit_dt) OVER (PARTITION BY member_id ORDER BY admit_dt, claim_id)
```

- `PARTITION BY member_id`: look at each member's stays separately.
- `ORDER BY admit_dt`: put them in time order.
- `LEAD(admit_dt)`: for each stay, get the **next** stay's admission date.
- If that next admission is `BETWEEN discharge_dt AND discharge_dt + 30 days`, set `readmit_30 = 1`, otherwise 0.
- The rate is `AVG(readmit_30)`, the average of 0s and 1s, which equals the percentage.

This is a window function, so there's no self-join. Each stay only needs to be compared with the member's next stay: if the next one isn't within 30 days, no later one can be either.

**Decisions and caveats:**
- **Overlapping stays (1,810):** sometimes the next admission starts *before* the current discharge. That's a data artefact or a transfer, not a return to hospital, so it isn't counted.
- **Same-day readmissions (288)** are counted. In real data many would be hospital transfers. That's a known simplification I documented.
- **End of data (right-censoring):** for 312 stays discharged in the last 30 days of 2010, a readmission would fall in 2011, which isn't in the data. So 2010 Q4 understates the rate.
- **Dates fall back:** if admission or discharge date is missing, the claim's from or thru date is used. In this data no inpatient claim was missing an admission date.

> "I use LEAD over each member's stays ordered by admission date to get the next admission, and flag it if it falls within 30 days of discharge. The rate is the average of that flag. I don't count overlapping stays, and I call out that the last 30 days of data can't show readmissions yet."

---

## 7. The other KPIs

- **PMPM (per member per month)** = total paid ÷ member-months. Member-months add up each member's months of coverage, so someone covered for 6 months counts as 6, not 12. It's the standard benefits cost metric because it normalises spend for membership size, so two plans of different size can be compared. Result: **$223.77 overall**, from $259.49 (2008) to $145.88 (2010).
- **Average length of stay (ALOS)** = average days from admission to discharge for inpatient stays. **5.69 days.**
- **Monthly / weekly volume and spend:** claim count, distinct members with claims, total and average paid. I fixed the weekly grouping. The original used `DATEDIFF(week, ...)`, which counts *Sunday* boundaries, so Sunday claims landed in the following week. The fix counts whole days from a known Monday (1900-01-01) and rounds down to a multiple of 7.
- **Paid by diagnosis:** spend grouped by primary ICD-9 code, for "top N" analysis in the spreadsheet.

---

## 8. Data quality: how I approached it

1. **Reconcile:** file rows = staging rows, and staging rows − final rows = merged segments.
2. **Profile before transforming:** I checked max lengths against column sizes, values that wouldn't convert, duplicate keys, and members on claims missing from the beneficiary files. Doing this *before* running the script meant nothing failed halfway.
3. **Count everything, and decide explicitly:** for each issue, record how many rows, then **keep, exclude, or fix**, and say why. Examples:
   - **Negative payments (2,612):** kept. They're adjustments, and removing them overstates spend.
   - **Zero payments (31,868):** kept. They're real activity with no payment.
   - **Missing diagnosis (215):** kept, but left out of the diagnosis view only.
   - **Claims dated in 2007 (536):** kept in the table, left out of the trend views (they'd look like stray tiny months).
4. **Separate data errors from interpretation caveats:** the big 2010 drop (admissions per 1,000 members halve) is how the synthetic file was generated, not a real-world trend, and the docs say so.

> "I found 10 issues that needed a decision, and I documented the count and the handling for each. Equally important, I recorded the checks that came back clean, so the report's users know what was tested."

---

## 9. The report and the macro

- One sheet per KPI view, plus a **Summary** with headline KPIs and three charts: monthly paid trend, PMPM by year, and readmission rate by quarter.
- The Summary's tables are **formulas over the view sheets** (`SUMIFS`), not pasted numbers. When the data refreshes, the charts update.
- **In Excel**, each view sheet is a live query to SQL Server. The **VBA macro** `RefreshAndExport`:
  1. refreshes all queries (`RefreshAll`),
  2. waits until they finish (`CalculateUntilAsyncQueriesDone`). **Background refresh is turned off** on each query, otherwise Excel would carry on before the data arrived and export stale numbers,
  3. stamps "Last refreshed",
  4. exports the Summary as a dated PDF, ready to email.
- **Google Sheets** can't reach a SQL Server on my PC and doesn't run VBA, so that version is a static snapshot produced by `build_report.py`.

> "The business user clicks one button: the data refreshes from SQL Server, the charts update, and a dated PDF is saved. I turned off background refresh so the export always waits for fresh data."

---

## 10. Likely questions

**Why SQL views instead of doing the maths in Excel?**
The KPI logic lives in one place, the database, so every report or tool that uses it gets the same numbers. Excel just presents them.

**How would you scale this?**
Schedule the load and SQL script (SQL Agent on paid editions, or Task Scheduler with sqlcmd on Express). Load incrementally rather than truncate-and-reload. Add a load log table with row counts per run, and put the views behind a BI tool such as Power BI.

**What would you change for real (non-synthetic) data?**
Use Part B coverage months as the outpatient PMPM denominator. Exclude planned readmissions and transfers, as CMS's official readmission measure does. Add claim adjustment logic, where a later version of a claim replaces the earlier one. Restrict access, because real claims are protected health information (PHI).

**What was the hardest part?**
Segments. Realising that segment-2 rows have no dates but do have money changed how the merge had to work.
