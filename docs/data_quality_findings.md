# Data Quality Findings

Source: CMS DE-SynPUF 2008–2010, Sample 1 (Beneficiary Summary 2008/2009/2010, Inpatient Claims, Outpatient Claims).
Checks come from section 3 of [`sql/claims_kpi_project.sql`](../sql/claims_kpi_project.sql), plus the profiling queries described below. All figures are from the actual load.

## Load reconciliation

| Item | Rows |
|---|---:|
| Beneficiary 2008 / 2009 / 2010 (staging) | 116,352 / 114,538 / 112,754 |
| Inpatient staging rows | 66,773 |
| Inpatient claims after merging segments | 66,705 |
| Outpatient staging rows | 790,790 |
| Outpatient claims after merging segments | 779,815 |
| Members (`dbo.member`) | 116,352 |
| Member-year coverage rows (`dbo.member_coverage`) | 343,644 |

Every staging table's row count matched its CSV exactly (checked by the loader). The difference between staging rows and claims is fully explained by segment merging (finding 2).

## Checks that came back clean

These were tested and found **no** problems, so no handling was needed:

- No unparseable dates or amounts. Every non-empty date converts with style 112 (`yyyymmdd`) and every non-empty `CLM_PMT_AMT` converts to decimal.
- No claim with `thru_date` before `from_date`.
- No claim with a missing paid amount.
- No inpatient claim without an admission date.
- No claim ID shared by two members, and no duplicate (claim ID, segment) rows or fully duplicate rows.
- No member appears twice in the same beneficiary year.
- Every member on a claim exists in the beneficiary files, so no foreign key failures.
- No member is missing birth date, sex or state.
- No value is longer than its target column (the longest IDs are 16 characters for members and 15 for claims).

## Issues found and how each was handled

| # | Issue | Records | Handling |
|---|---|---:|---|
| 1 | **Source file mislabeled.** On cms.gov, the "Sample 1 2010 Beneficiary Summary" link downloads `de1_0_2010_beneficiary_summary_file_sample_20.zip`. | 1 file | Verified before use. The CSV inside is named `..._Sample_1.csv`, and 100% of its 112,754 member IDs also appear in the 2008 and 2009 Sample 1 files, so it is Sample 1 data with a wrong zip name. Used as-is. |
| 2 | **Claims split into segments.** Long claims arrive as several rows (`SEGMENT` 1, 2). | 68 IP + 10,975 OP claims (staging − claims = 68 IP, 10,975 OP rows) | Merged to one row per claim (`GROUP BY CLM_ID`). Dates come from segment 1, because segment-2 rows have no dates. Paid amounts are **summed**, because segment-2 rows carry their own payments (66 of 68 IP and 10,771 of 10,975 OP multi-segment claims have a non-zero segment-2 payment). |
| 3 | **Orphan segment-2 claims.** The claim has a segment-2 row but no segment-1 (header) row, so it has no dates. | 278 OP | Kept in `dbo.claim`, so their payments are not lost and they appear in the diagnosis view. They drop out of all date-based KPIs (weekly, monthly, PMPM), which require a date. Reported as `missing_from_date = 278`. |
| 4 | **Segments disagree.** Segments of the same claim carry different primary diagnoses or providers. | Diagnosis: 5,543 OP. Provider: 48 IP + 7,655 OP | Take the segment-1 value, since segment 1 is the claim header. The original script used `MIN()`, which picked the alphabetically lowest code: an arbitrary choice. |
| 5 | **Negative paid amounts.** | 2,612 claims (55 IP, 2,557 OP), totalling −$113,570 | Kept. Negative payments are adjustments or recoupments and correctly reduce spend. Excluding them would overstate cost. |
| 6 | **Zero paid amounts.** | 31,868 claims (2,155 IP, 29,713 OP) | Kept. They are real utilisation (counted in claim volume) with no Medicare payment, for example when another payer paid or the claim was denied. They lower "average paid per claim", which is noted in the KPI definitions. |
| 7 | **Missing primary diagnosis.** | 215 claims (27 IP, 188 OP) | Kept for volume and spend KPIs, and excluded only from `v_paid_by_diagnosis`, where a NULL diagnosis can't be grouped meaningfully. |
| 8 | **Claims dated outside the 2008–2010 window.** All are November–December 2007 starts. | 536 (224 IP, 312 OP) | Kept in `dbo.claim`. Excluded from the weekly, monthly and quarterly trend views, where they would appear as stray partial months. Also excluded from PMPM automatically, because there is no 2007 member-month denominator. |
| 9 | **Claims for members with zero Part A coverage months that year.** | 5,310 claims (2,780 in 2008, 1,877 in 2009, 653 in 2010) | Kept: the spend is real. Noted as a limitation, because this spend enters the PMPM numerator with no matching member-months in the denominator. Members with zero months: 5,294 (2008), 7,241 (2009), 6,319 (2010). |
| 10 | **Overlapping inpatient stays.** The member's next admission starts before the current stay's discharge. | 1,810 stays | Not counted as readmissions, because the rule requires the next admission to fall between discharge and discharge + 30 days. A new admission during an ongoing stay is a data artefact (or a transfer), not a return to hospital. |

**Total: 10 data quality issues needing a handling decision**, plus the analytic caveats below.

## Analytic caveats (not errors, but they affect interpretation)

- **The synthetic file has much less inpatient activity in 2010.** Admissions per 1,000 member-years: 256 (2008), 236 (2009), 129 (2010), while member-months fell only 3% (1,296,626 → 1,262,399). Inpatient spend and the readmission rate fall with it (quarterly readmission rate goes from 16.5% in 2008 Q1 to 2.6% in 2010 Q4). This is a property of how the DE-SynPUF was generated, not a real Medicare trend. Year-over-year comparisons in this dataset should not be read as clinical findings.
- **Readmissions near the end of the data are under-counted (right-censoring).** 312 stays were discharged in the last 30 days of 2010. A readmission in January 2011 would not be in the data, so 2010 Q4 is biased downward.
- **Same-day readmissions are counted.** 288 stays have a next admission on the discharge date. In real claims many of these would be transfers between hospitals. The rule counts them, which is the simpler, documented choice.
- **The PMPM denominator is Part A months for both claim types.** `BENE_HI_CVRAGE_TOT_MONS` is Part A (hospital) coverage. Outpatient is technically Part B, whose months are in `BENE_SMI_CVRAGE_TOT_MONS`. Using one denominator keeps IP and OP PMPM additive. A production version could use Part B months for OP.
- **Paid amounts are rounded to multiples of $10** in the synthetic data (all 846,520 claims), so averages have that granularity.

## Handled at load time (before SQL)

- **Empty CSV fields are loaded as NULL, not empty strings.** In SQL Server, `TRY_CONVERT(date, '', 112)` returns `1900-01-01` instead of NULL (verified on this server). If empty strings had reached staging, every missing date would have become a valid-looking 1900 date and passed the `missing_from_date` check.
