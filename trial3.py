import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
import re
import qrcode

master_path = Path("/workspaces/DUMP/Training Progress Tracker.xlsx")
output_dir = Path("/workspaces/DUMP/Employee_Reports")
output_dir.mkdir(exist_ok=True)
qr_dir = output_dir / "QR_Codes"
qr_dir.mkdir(exist_ok=True)

BASE_URL = "C://Users//l.karuru//OneDrive - Lödige Industries GmbH//test/"

# --- FUNCTION TO FIND HEADER ROW ---
def find_header_row(sheet_name, file_path):
    """Finds the row number that contains 'Emp. No.' and 'Employee Name'."""
    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
    for i, row in preview.iterrows():
        if "Emp. No." in row.values and "Employee Name" in row.values:
            return i
    raise ValueError(f"Could not find header row in sheet: {sheet_name}")

# --- LOAD ALL SHEETS ---
sheet_names = [
    "Cargo Trainings",
    "DFW Trainings",
    "AMH Trainings",
    "CBF Trainings",
    "SOPS",
    "EXAMS",
]

all_dfs = []
for sheet in sheet_names:
    header_row = find_header_row(sheet, master_path)
    df = pd.read_excel(master_path, sheet_name=sheet, header=header_row)
    df.columns = df.columns.str.strip()  # Clean column names

    # Format all 'date' columns in this sheet
    date_cols = [c for c in df.columns if "date" in str(c).lower()]
    if not date_cols:
        print(f"\nNo 'Date' column found in sheet: {sheet}")
    else:
        for col in date_cols:
            df[col] = pd.to_datetime(df[col], errors="coerce").apply(
                lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else ""
            )

    all_dfs.append(df)

# --- CLEAN FILENAME ---
def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))

# --- GET UNIQUE EMPLOYEES ---
employee_ids = set()
for df in all_dfs:
    if "Emp. No." not in df.columns or "Employee Name" not in df.columns:
        raise ValueError(f"Sheet missing required columns: {df.columns.tolist()}")
    employee_ids.update(zip(df["Emp. No."], df["Employee Name"]))

employee_list = list(employee_ids)[:20]  # limit to first 20 employees for testing

# --- CREATE REPORTS FOR FIRST 20 EMPLOYEES ---
for emp_id, emp_name in employee_list:
    emp_name_clean = clean_filename(emp_name)
    file_name = f"trainings_for_{emp_name_clean}_{emp_id}.xlsx"
    report_path = output_dir / file_name

    wb = Workbook()
    wb.remove(wb.active)

    for df, sheet_name in zip(all_dfs, sheet_names):
        ws = wb.create_sheet(title=sheet_name)
        emp_df = df[df["Emp. No."] == emp_id]
        if emp_df.empty:
            ws.append([f"No records found for {emp_name} in {sheet_name}"])
            continue
        for r in dataframe_to_rows(emp_df, index=False, header=True):
            ws.append(r)

    wb.save(report_path)

# --- GENERATE QR CODES ---
for excel_file in output_dir.glob("*.xlsx"):
    emp_name_id = excel_file.stem.replace("trainings_for_", "")
    qr_file_name = f"Qr_code_for_{emp_name_id}.png"
    qr_file_path = qr_dir / qr_file_name

    file_url = BASE_URL.rstrip("/") + "/" + excel_file.name
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(file_url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    img.save(qr_file_path)

    print(f"QR code created: {qr_file_path}")

print(f"First 20 employee training files created in: {output_dir}")
