import pandas as pd
from openpyxl import Workbook
from pathlib import Path
import re
from datetime import date

master_path = Path("/workspaces/DUMP/Training Progress Tracker.xlsx")

# trainings = [pd.read_excel(master_path, sheet_name='Cargo Trainings',),
#              pd.read_excel(master_path, sheet_name='DFW Trainings'),
#              pd.read_excel(master_path, sheet_name='AMH Trainings'),
#              pd.read_excel(master_path, sheet_name='CBF Trainings')]

# sops   = pd.read_excel(master_path, sheet_name="SOPS")
# exams  = pd.read_excel(master_path, sheet_name="EXAMS")   # keep separate variable
trainings = [pd.read_excel(master_path, sheet_name='Cargo Trainings', skiprows=2),
             pd.read_excel(master_path, sheet_name='DFW Trainings', skiprows=2),
             pd.read_excel(master_path, sheet_name='AMH Trainings', skiprows=2),
             pd.read_excel(master_path, sheet_name='CBF Trainings', skiprows=2)]

sops   = pd.read_excel(master_path, sheet_name="SOPS", skiprows=2)
exams  = pd.read_excel(master_path, sheet_name="EXAMS", skiprows=2)
# Combine them all in one list
all_dfs = trainings + [sops, exams]

sheet_names = [
    "Cargo Trainings",
    "DFW Trainings",
    "AMH Trainings",
    "CBF Trainings",
    "SOPS",
    "EXAMS",
]

for name, df in zip(sheet_names, all_dfs):
    # locate a column literally named 'Date' (case‑insensitive)
    # date_col = next((c for c in df.columns if c.lower() == "date"), None)
    # date_col = next((col for col in df.columns if col.lower() == 'date'), None)
    date_col = next(
    (str(c) for c in df.columns if str(c).strip().lower() == "date"),
    None
)

    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.strftime("%Y-%m-%d")
    else:
        print(f"\n No 'Date' column found in sheet: {name}")

    print(f"\n--- First 5 rows of '{name}' ---")
    print(df.head(5))
# employee_dataframes = []
# for name, df in zip(sheet_names, all_dfs):
#     df.columns = [col.strip() for col in df.columns]

#     if "Emp. No." in df.columns and "Employee Name" in df.columns:
#         employee_dataframes.append(df[["Emp. No.", "Employee Name"]])
#     else:
#         print(f" Sheet '{name}' is missing 'Emp. No.' or 'Employee Name' columns. Columns found: {df.columns.tolist()}")
#     print(df.head(3))
# #  Combine list of DataFrames into one master DataFrame
# master = pd.concat(employee_dataframes, ignore_index=True).dropna(subset=["Emp. No."])

#  Remove duplicates and sort
# employees = (
#     master
#     .drop_duplicates(subset=["Emp. No."])
#     .sort_values("Emp. No.")
#     .reset_index(drop=True)
# )

# print(employees.head())
