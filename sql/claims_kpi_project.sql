/* =====================================================================
   Claims KPI Reporting Project  --  SQL Server (Express or Developer)
   Data: CMS DE-SynPUF Sample 1 (synthetic Medicare claims, public)

   Run section by section in SSMS, or end to end with sqlcmd.
   Before running: load the CSVs into staging tables with
   scripts/load_staging.py. It creates the database, names every column
   exactly as the CSV header, makes every column nvarchar(50), and loads
   empty fields as NULL -- this script does the type conversion itself,
   which is part of the cleaning work. (Empty strings must not reach the
   conversions below: '' converts to 1900-01-01 as a date.)

   Staging table names this script expects:
     stg_beneficiary_2008, stg_beneficiary_2009, stg_beneficiary_2010
     stg_inpatient, stg_outpatient

   The script is re-runnable: sections 0-1 drop and recreate everything
   except the staging tables.
   ===================================================================== */


/* ---------- 0. Database ---------- */
IF DB_ID('ClaimsKPI') IS NULL CREATE DATABASE ClaimsKPI;
GO
USE ClaimsKPI;
GO


/* ---------- 1. Reporting schema ---------- */
DROP VIEW  IF EXISTS dbo.v_paid_by_diagnosis, dbo.v_kpi_annual_pmpm, dbo.v_kpi_inpatient_quarterly,
                     dbo.v_inpatient_stays, dbo.v_kpi_monthly, dbo.v_kpi_weekly;
DROP TABLE IF EXISTS dbo.claim, dbo.member_coverage, dbo.member;

CREATE TABLE dbo.member (
    member_id    VARCHAR(16) NOT NULL PRIMARY KEY,
    birth_date   DATE        NULL,
    sex          CHAR(1)     NULL,
    state_code   VARCHAR(2)  NULL
);

CREATE TABLE dbo.member_coverage (
    member_id        VARCHAR(16) NOT NULL REFERENCES dbo.member(member_id),
    coverage_year    INT         NOT NULL,
    coverage_months  INT         NOT NULL,
    PRIMARY KEY (member_id, coverage_year)
);

CREATE TABLE dbo.claim (
    claim_id        VARCHAR(20)   NOT NULL,
    claim_type      CHAR(2)       NOT NULL,   -- 'IP' inpatient, 'OP' outpatient
    member_id       VARCHAR(16)   NOT NULL REFERENCES dbo.member(member_id),
    provider_id     VARCHAR(10)   NULL,
    from_date       DATE          NULL,
    thru_date       DATE          NULL,
    admit_date      DATE          NULL,
    discharge_date  DATE          NULL,
    primary_dx      VARCHAR(10)   NULL,
    paid_amount     DECIMAL(12,2) NULL,
    PRIMARY KEY (claim_id, claim_type)
);
CREATE INDEX ix_claim_member_from ON dbo.claim (member_id, from_date);
CREATE INDEX ix_claim_type_from   ON dbo.claim (claim_type, from_date);
GO


/* ---------- 2. Load + clean ---------- */

-- 2a. Members: one row per person, demographics from their latest year on file
WITH all_years AS (
    SELECT DESYNPUF_ID, BENE_BIRTH_DT, BENE_SEX_IDENT_CD, SP_STATE_CODE, 2010 AS yr FROM dbo.stg_beneficiary_2010
    UNION ALL
    SELECT DESYNPUF_ID, BENE_BIRTH_DT, BENE_SEX_IDENT_CD, SP_STATE_CODE, 2009 FROM dbo.stg_beneficiary_2009
    UNION ALL
    SELECT DESYNPUF_ID, BENE_BIRTH_DT, BENE_SEX_IDENT_CD, SP_STATE_CODE, 2008 FROM dbo.stg_beneficiary_2008
),
ranked AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY DESYNPUF_ID ORDER BY yr DESC) AS rn
    FROM all_years
)
INSERT INTO dbo.member (member_id, birth_date, sex, state_code)
SELECT DESYNPUF_ID,
       TRY_CONVERT(date, CAST(BENE_BIRTH_DT AS varchar(8)), 112),
       CASE CAST(BENE_SEX_IDENT_CD AS varchar(2)) WHEN '1' THEN 'M' WHEN '2' THEN 'F' END,
       CAST(SP_STATE_CODE AS varchar(2))
FROM ranked
WHERE rn = 1;

-- 2b. Coverage months per member per year (Part A months; used for PMPM)
INSERT INTO dbo.member_coverage (member_id, coverage_year, coverage_months)
SELECT DESYNPUF_ID, yr, ISNULL(TRY_CAST(months AS int), 0)
FROM (
    SELECT DESYNPUF_ID, BENE_HI_CVRAGE_TOT_MONS AS months, 2008 AS yr FROM dbo.stg_beneficiary_2008
    UNION ALL
    SELECT DESYNPUF_ID, BENE_HI_CVRAGE_TOT_MONS, 2009 FROM dbo.stg_beneficiary_2009
    UNION ALL
    SELECT DESYNPUF_ID, BENE_HI_CVRAGE_TOT_MONS, 2010 FROM dbo.stg_beneficiary_2010
) b;

-- 2c. Inpatient claims. Long claims are split into multiple SEGMENT rows
--     in the source; collapse them to one row per claim.
--     Dates: segment 2 rows have no dates, so MIN/MAX come from segment 1.
--     Paid: summed, because segment 2 rows carry their own payments.
--     Provider and diagnosis: segments can disagree, so take segment 1's
--     value (the claim header); fall back to MIN only if there is no segment 1.
INSERT INTO dbo.claim (claim_id, claim_type, member_id, provider_id, from_date, thru_date,
                       admit_date, discharge_date, primary_dx, paid_amount)
SELECT CLM_ID, 'IP', DESYNPUF_ID,
       COALESCE(MAX(CASE WHEN SEGMENT = '1' THEN PRVDR_NUM END), MIN(PRVDR_NUM)),
       MIN(TRY_CONVERT(date, CAST(CLM_FROM_DT        AS varchar(8)), 112)),
       MAX(TRY_CONVERT(date, CAST(CLM_THRU_DT        AS varchar(8)), 112)),
       MIN(TRY_CONVERT(date, CAST(CLM_ADMSN_DT       AS varchar(8)), 112)),
       MAX(TRY_CONVERT(date, CAST(NCH_BENE_DSCHRG_DT AS varchar(8)), 112)),
       COALESCE(MAX(CASE WHEN SEGMENT = '1' THEN NULLIF(ICD9_DGNS_CD_1, '') END), MIN(NULLIF(ICD9_DGNS_CD_1, ''))),
       SUM(TRY_CAST(CLM_PMT_AMT AS decimal(12,2)))
FROM dbo.stg_inpatient
GROUP BY CLM_ID, DESYNPUF_ID;

-- 2d. Outpatient claims (same segment rules as 2c)
INSERT INTO dbo.claim (claim_id, claim_type, member_id, provider_id, from_date, thru_date,
                       admit_date, discharge_date, primary_dx, paid_amount)
SELECT CLM_ID, 'OP', DESYNPUF_ID,
       COALESCE(MAX(CASE WHEN SEGMENT = '1' THEN PRVDR_NUM END), MIN(PRVDR_NUM)),
       MIN(TRY_CONVERT(date, CAST(CLM_FROM_DT AS varchar(8)), 112)),
       MAX(TRY_CONVERT(date, CAST(CLM_THRU_DT AS varchar(8)), 112)),
       NULL, NULL,
       COALESCE(MAX(CASE WHEN SEGMENT = '1' THEN NULLIF(ICD9_DGNS_CD_1, '') END), MIN(NULLIF(ICD9_DGNS_CD_1, ''))),
       SUM(TRY_CAST(CLM_PMT_AMT AS decimal(12,2)))
FROM dbo.stg_outpatient
GROUP BY CLM_ID, DESYNPUF_ID;
GO


/* ---------- 3. Data quality checks ----------
   Run these and record the results in your README. Writing up what you
   found (and how you handled it) is the "data quality analyst" story. */

-- Row counts: staging vs. loaded (difference = merged segments)
SELECT 'inpatient staging rows' AS item, COUNT(*) AS n FROM dbo.stg_inpatient
UNION ALL SELECT 'inpatient claims loaded',  COUNT(*) FROM dbo.claim WHERE claim_type = 'IP'
UNION ALL SELECT 'outpatient staging rows',  COUNT(*) FROM dbo.stg_outpatient
UNION ALL SELECT 'outpatient claims loaded', COUNT(*) FROM dbo.claim WHERE claim_type = 'OP';

-- Problem records
SELECT
    SUM(CASE WHEN from_date IS NULL               THEN 1 ELSE 0 END) AS missing_from_date,
    SUM(CASE WHEN thru_date < from_date           THEN 1 ELSE 0 END) AS thru_before_from,
    SUM(CASE WHEN paid_amount IS NULL             THEN 1 ELSE 0 END) AS missing_paid,
    SUM(CASE WHEN paid_amount < 0                 THEN 1 ELSE 0 END) AS negative_paid,
    SUM(CASE WHEN paid_amount = 0                 THEN 1 ELSE 0 END) AS zero_paid,
    SUM(CASE WHEN primary_dx IS NULL              THEN 1 ELSE 0 END) AS missing_dx,
    SUM(CASE WHEN claim_type = 'IP' AND COALESCE(admit_date, from_date) IS NULL THEN 1 ELSE 0 END) AS ip_missing_admit,
    SUM(CASE WHEN from_date < '2008-01-01' OR from_date >= '2011-01-01' THEN 1 ELSE 0 END) AS outside_2008_2010
FROM dbo.claim;

-- Source-structure checks on staging
SELECT 'claims with only a segment 2 row (no header row, so no dates)' AS item, COUNT(*) AS n
FROM (SELECT CLM_ID FROM dbo.stg_inpatient  GROUP BY CLM_ID HAVING MAX(CASE WHEN SEGMENT = '1' THEN 1 ELSE 0 END) = 0
      UNION ALL
      SELECT CLM_ID FROM dbo.stg_outpatient GROUP BY CLM_ID HAVING MAX(CASE WHEN SEGMENT = '1' THEN 1 ELSE 0 END) = 0) x
UNION ALL
SELECT 'multi-segment claims whose segments disagree on primary dx', COUNT(*)
FROM (SELECT CLM_ID FROM dbo.stg_inpatient  GROUP BY CLM_ID HAVING COUNT(DISTINCT ICD9_DGNS_CD_1) > 1
      UNION ALL
      SELECT CLM_ID FROM dbo.stg_outpatient GROUP BY CLM_ID HAVING COUNT(DISTINCT ICD9_DGNS_CD_1) > 1) x;
GO


/* ---------- 4. KPI views (these feed the Excel report) ----------
   Time-trend views are limited to the 2008-2010 coverage window: a few
   claims start in late 2007 and would show up as stray partial months. */

-- Weekly volume and spend (weeks start Monday). Day 0 (1900-01-01) is a
-- Monday, so whole days since day 0, rounded down to a multiple of 7, give
-- the week's Monday. (DATEDIFF(week, ...) counts Sunday boundaries, which
-- would put Sunday claims in the following week.)
CREATE VIEW dbo.v_kpi_weekly AS
SELECT CAST(DATEADD(day, DATEDIFF(day, 0, from_date) / 7 * 7, 0) AS date) AS week_start,
       claim_type,
       COUNT(*)                  AS claim_count,
       COUNT(DISTINCT member_id) AS members_with_claims,
       SUM(paid_amount)          AS total_paid,
       AVG(paid_amount)          AS avg_paid_per_claim
FROM dbo.claim
WHERE from_date >= '2008-01-01' AND from_date < '2011-01-01'
GROUP BY CAST(DATEADD(day, DATEDIFF(day, 0, from_date) / 7 * 7, 0) AS date), claim_type;
GO

-- Monthly volume and spend
CREATE VIEW dbo.v_kpi_monthly AS
SELECT DATEFROMPARTS(YEAR(from_date), MONTH(from_date), 1) AS month_start,
       claim_type,
       COUNT(*)                  AS claim_count,
       COUNT(DISTINCT member_id) AS members_with_claims,
       SUM(paid_amount)          AS total_paid,
       AVG(paid_amount)          AS avg_paid_per_claim
FROM dbo.claim
WHERE from_date >= '2008-01-01' AND from_date < '2011-01-01'
GROUP BY DATEFROMPARTS(YEAR(from_date), MONTH(from_date), 1), claim_type;
GO

-- One row per inpatient stay with length of stay and 30-day readmission flag
CREATE VIEW dbo.v_inpatient_stays AS
WITH stays AS (
    SELECT claim_id, member_id, paid_amount,
           COALESCE(admit_date, from_date)     AS admit_dt,
           COALESCE(discharge_date, thru_date) AS discharge_dt
    FROM dbo.claim
    WHERE claim_type = 'IP'
)
SELECT claim_id, member_id, paid_amount, admit_dt, discharge_dt,
       DATEDIFF(day, admit_dt, discharge_dt) AS los_days,
       CASE WHEN LEAD(admit_dt) OVER (PARTITION BY member_id ORDER BY admit_dt, claim_id)
                 BETWEEN discharge_dt AND DATEADD(day, 30, discharge_dt)
            THEN 1 ELSE 0 END AS readmit_30
FROM stays
WHERE admit_dt IS NOT NULL AND discharge_dt IS NOT NULL;
GO

-- Quarterly inpatient KPIs
CREATE VIEW dbo.v_kpi_inpatient_quarterly AS
SELECT YEAR(admit_dt)              AS yr,
       DATEPART(quarter, admit_dt) AS qtr,
       COUNT(*)                                  AS admissions,
       SUM(paid_amount)                          AS total_paid,
       AVG(CAST(los_days   AS decimal(8,2)))     AS avg_los_days,
       AVG(CAST(readmit_30 AS decimal(6,4)))     AS readmit_30_rate
FROM dbo.v_inpatient_stays
WHERE admit_dt >= '2008-01-01' AND admit_dt < '2011-01-01'
GROUP BY YEAR(admit_dt), DATEPART(quarter, admit_dt);
GO

-- Annual cost per member per month (PMPM) -- the standard benefits-industry cost KPI
CREATE VIEW dbo.v_kpi_annual_pmpm AS
WITH paid AS (
    SELECT YEAR(from_date) AS yr, claim_type, SUM(paid_amount) AS total_paid
    FROM dbo.claim
    WHERE from_date IS NOT NULL
    GROUP BY YEAR(from_date), claim_type
),
mm AS (
    SELECT coverage_year AS yr, SUM(coverage_months) AS member_months
    FROM dbo.member_coverage
    GROUP BY coverage_year
)
SELECT p.yr, p.claim_type, p.total_paid, m.member_months,
       CAST(p.total_paid / NULLIF(m.member_months, 0) AS decimal(12,2)) AS pmpm
FROM paid p
JOIN mm m ON m.yr = p.yr;
GO

-- Spend by primary diagnosis (sort/filter top N in Excel)
CREATE VIEW dbo.v_paid_by_diagnosis AS
SELECT primary_dx, claim_type,
       COUNT(*)         AS claim_count,
       SUM(paid_amount) AS total_paid
FROM dbo.claim
WHERE primary_dx IS NOT NULL
GROUP BY primary_dx, claim_type;
GO


/* ---------- 5. Sanity check the views ---------- */
SELECT TOP 12 * FROM dbo.v_kpi_monthly ORDER BY month_start, claim_type;
SELECT * FROM dbo.v_kpi_inpatient_quarterly ORDER BY yr, qtr;
SELECT * FROM dbo.v_kpi_annual_pmpm ORDER BY yr, claim_type;
