"""
Load the CMS DE-SynPUF Sample 1 CSVs into SQL Server staging tables.

Every staging column is NVARCHAR(50): staging is a raw, untyped copy of the
file. Type conversion and cleaning happen afterwards in T-SQL
(sql/claims_kpi_project.sql, section 2), where problems can be counted.

Empty CSV fields are loaded as NULL, not ''. In SQL Server an empty string
converts to 1900-01-01 as a date and 0 as a number, which would silently
corrupt the cleaned data.

Usage:  py -3.9 scripts/load_staging.py
Needs:  pip install pyodbc ; ODBC Driver 18 for SQL Server
"""
import csv
import pathlib
import time

import pyodbc

SERVER = r"localhost\SQLEXPRESS"
DATABASE = "ClaimsKPI"
RAW = pathlib.Path(__file__).resolve().parent.parent / "data" / "raw"
BATCH = 20_000

FILES = {
    "stg_beneficiary_2008": "DE1_0_2008_Beneficiary_Summary_File_Sample_1.csv",
    "stg_beneficiary_2009": "DE1_0_2009_Beneficiary_Summary_File_Sample_1.csv",
    "stg_beneficiary_2010": "DE1_0_2010_Beneficiary_Summary_File_Sample_1.csv",
    "stg_inpatient":        "DE1_0_2008_to_2010_Inpatient_Claims_Sample_1.csv",
    "stg_outpatient":       "DE1_0_2008_to_2010_Outpatient_Claims_Sample_1.csv",
}


def connect(database):
    return pyodbc.connect(
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={SERVER};DATABASE={database};"
        "Trusted_Connection=yes;TrustServerCertificate=yes",
        autocommit=True,
    )


def main():
    with connect("master") as cn:
        cn.execute(f"IF DB_ID('{DATABASE}') IS NULL CREATE DATABASE {DATABASE};")

    cn = connect(DATABASE)
    cur = cn.cursor()
    cur.fast_executemany = True  # sends rows in bulk instead of one round trip each

    for table, filename in FILES.items():
        start = time.time()
        with open(RAW / filename, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader)
            cols = ", ".join(f"[{c}] NVARCHAR(50) NULL" for c in header)
            cur.execute(f"DROP TABLE IF EXISTS dbo.{table}; CREATE TABLE dbo.{table} ({cols});")

            insert = (f"INSERT INTO dbo.{table} ({', '.join(f'[{c}]' for c in header)}) "
                      f"VALUES ({', '.join('?' * len(header))})")
            file_rows, batch = 0, []
            for row in reader:
                batch.append([v if v != "" else None for v in row])
                if len(batch) == BATCH:
                    cur.executemany(insert, batch)
                    file_rows += len(batch)
                    batch = []
            if batch:
                cur.executemany(insert, batch)
                file_rows += len(batch)

        loaded = cur.execute(f"SELECT COUNT(*) FROM dbo.{table}").fetchval()
        status = "OK" if loaded == file_rows else "MISMATCH"
        print(f"{table:<22} file rows {file_rows:>8,}  loaded {loaded:>8,}  "
              f"{status}  ({time.time() - start:.0f}s)")


if __name__ == "__main__":
    main()
