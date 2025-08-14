import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
from datetime import datetime, timedelta
import qrcode

# ---------------- CONFIG ----------------
master_path = Path("/workspaces/DUMP/Training Progress Tracker.xlsx")
output_dir = Path("/workspaces/DUMP/Employee_Reports3")
qr_dir = Path("/workspaces/DUMP/Employee_Reports3/QR_Codes")
output_dir.mkdir(exist_ok=True)
qr_dir.mkdir(exist_ok=True)

BASE_URL = "https://lodigeindustries-my.sharepoint.com/:f:/r/personal/l_karuru_lodige_com/Documents/test"

# ------------- HELPER FUNCTIONS -------------
def find_header_row(sheet_name, file_path):
    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
    for i, row in preview.iterrows():
        if "Emp. No." in row.values and "Employee Name" in row.values:
            return i
    raise ValueError(f"Could not find header row in sheet: {sheet_name}")

def deduplicate_columns(cols):
    seen = {}
    new_cols = []
    for col in cols:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    return new_cols

# Sheets to read
sheet_names = [
    "Cargo Trainings",
    "DFW Trainings",
    "AMH Trainings",
    "CBF Trainings",
    "SOPS",
    "EXAMS",
]

# ------------- READ ALL SHEETS -------------
all_dfs = {}
for sheet in sheet_names:
    header_row = find_header_row(sheet, master_path)
    df = pd.read_excel(master_path, sheet_name=sheet, header=header_row)
    df.columns = pd.Series(df.columns).astype(str).str.strip().str.lower().str.replace(r'[^\w\s]', '', regex=True)
    df.columns = deduplicate_columns(df.columns)

    # Format date columns
    date_cols = [c for c in df.columns if str(c).strip().lower() in ("date", ".date") or "exam date" in str(c).strip().lower()]
    seen_cols = set()
    for col in date_cols:
        if col in seen_cols:
            continue
        seen_cols.add(col)
        df[col] = pd.to_datetime(df[col], errors="coerce").apply(
            lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else ""
        )

    all_dfs[sheet] = df

# Get unique employees from first sheet (limit to 25 for testing)
emp_df = all_dfs[sheet_names[0]]
employees = emp_df[['emp no', 'employee name']].drop_duplicates().head(25)

today = datetime.today()

# ------------- LOOP EMPLOYEES -------------
for _, emp in employees.iterrows():
    emp_no = emp['emp no']
    emp_name = emp['employee name']

    # --- Training Dashboard ---
    training_records = []
    for sheet_name, df in all_dfs.items():
        rows = df[df['emp no'] == emp_no]

        for date_col in [c for c in df.columns if c.strip().lower() in ("date", ".date")]:
            date_index = list(df.columns).index(date_col)
            training_name = df.columns[date_index + 1] if date_index + 1 < len(df.columns) else ""

            for _, row in rows.iterrows():
                training_date = pd.to_datetime(row.get(date_col, None), errors='coerce')

                # Calculate expiry date
                if pd.notna(training_date):
                    if sheet_name.lower() == "sops":
                        expiry_date = training_date + timedelta(days=365)
                    else:
                        expiry_date = training_date + timedelta(days=365 * 2)
                else:
                    expiry_date = None

                # Determine status
                if pd.notna(expiry_date):
                    days_left = (expiry_date - today).days
                    status = "VALID" if days_left >= 0 else "NOT VALID"
                else:
                    days_left = None
                    status = "NO EXPIRY DATE"

                training_records.append([
                    None,
                    training_name,
                    training_date.strftime('%d-%b-%Y') if pd.notna(training_date) else '',
                    expiry_date.strftime('%d-%b-%Y') if pd.notna(expiry_date) else '',
                    days_left,
                    today.strftime('%d-%b-%Y'),
                    status
                ])

    training_df = pd.DataFrame(training_records, columns=[
        "SN", "TRAININGS", "TRAINING DATE", "EXPIRY DATE", "PERIOD TO EXPIRE", "CURRENT DATE", "STATUS"
    ])
    training_df["SN"] = range(1, len(training_df) + 1)

    # --- Exam Dashboard ---
    exam_records = []
    for sheet_name, df in all_dfs.items():
        rows = df[df['emp no'] == emp_no]

        for exam_date_col in [c for c in df.columns if "exam date" in c.strip().lower()]:
            exam_index = list(df.columns).index(exam_date_col)
            exam_name = df.columns[exam_index + 1] if exam_index + 1 < len(df.columns) else ""

            for _, row in rows.iterrows():
                exam_date = pd.to_datetime(row.get(exam_date_col, None), errors='coerce')
                marks = row.get('marks attained', None)

                exam_records.append([
                    None,
                    exam_name,
                    exam_date.strftime('%d-%b-%Y') if pd.notna(exam_date) else '',
                    marks
                ])

    exam_df = pd.DataFrame(exam_records, columns=[
        "SN", "EXAM", "EXAM DATE", "MARKS ATTAINED"
    ])
    exam_df["SN"] = range(1, len(exam_df) + 1)
    if not exam_df.empty:
        exam_df.loc["TOTAL"] = ["", "", "TOTAL AVERAGE", exam_df["MARKS ATTAINED"].mean()]

    # --- Write to Excel ---
    wb = Workbook()
    ws = wb.active
    ws.title = f"{emp_name} ({emp_no})"

    # Training Dashboard
    ws.append([f"TRAINING DASHBOARD FOR {emp_name} ({emp_no})"])
    for r in dataframe_to_rows(training_df, index=False, header=True):
        ws.append(r)
    ws.append([])

    # Exam Dashboard
    ws.append([f"EXAM DASHBOARD FOR {emp_name} ({emp_no})"])
    for r in dataframe_to_rows(exam_df, index=False, header=True):
        ws.append(r)

    file_path = output_dir / f"{emp_name} {emp_no}.xlsx"
    wb.save(file_path)

    # --- Create QR Code for this employee file ---
    file_url = BASE_URL + file_path.name
    qr_file_path = qr_dir / f"Qr_code_for_{emp_name}_{emp_no}.png"
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(file_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(qr_file_path)

print("Dashboards & QR codes created successfully!")
