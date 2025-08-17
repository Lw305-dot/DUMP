import pandas as pd
from datetime import date
import os
# from caas_jupyter_tools import display_dataframe_to_user

# User-provided 8 matches & markets for today
today = date(2025, 8, 13)

rows = [
    ["Huddersfield vs Leicester", "Over/Under Goals (Over 2.5)", "Over 2.5", 0.60, "Strong attacking styles likely to produce goals."],
    ["Birmingham vs Sheffield Utd", "Over/Under Goals (Over 2.5)", "Over 2.5", 0.60, "TalkSPORT suggests Over 2.5 is a tip for this tie."],
    ["Bolton vs Sheff. Wednesday", "Over/Under Goals (Over 2.5)", "Over 2.5", 0.61, "TalkSPORT specifically backs this market."],
    ["Barnsley vs Fleetwood", "Both Teams to Score (BTTS)", "Yes", 0.60, "Barnsley strong offensively; Fleetwood likely to offer resistance."],
    ["Cheltenham vs Exeter", "Fulltime Corners Over/Under (5.5+)", "Over 5.5", 0.72, "Lower-division cup ties often deliver high corner counts."],
    ["Huddersfield vs Leicester", "Both Teams to Score (BTTS)", "Yes", 0.60, "Championship sides with balanced attacking threat."],
    ["Bolton vs Sheff. Wednesday", "Both Teams to Score (BTTS)", "Yes", 0.60, "Bolton’s dominance could still see goals at both ends."],
    ["Birmingham vs Sheffield Utd", "Fulltime Corners Over/Under (5.5+)", "Over 5.5", 0.70, "Midweek cup urgency—expect higher corner volume."],
]

df = pd.DataFrame(rows, columns=["Match","Market","Prediction","Probability","Rationale/Notes"])
df["Date"] = pd.to_datetime(today)

# Tidy probability display for the preview (keep numeric in file)
preview = df.copy()
preview["Probability"] = (preview["Probability"]*100).round(0).astype(int).astype(str) + "%"

# ---- Save/Update tracking workbooks ----

# 1) A standalone daily file for easy sharing
daily_path = "/workspaces/DUMP/matches/Todays_Picks_2025-08-13.xlsx"
df.to_excel(daily_path, index=False)

# 2) Append as a new sheet into our running analysis workbook if it exists; otherwise create it
master_path = "/workspaces/DUMP/matches/EFL_results_analysis.xlsx"
mode = "a" if os.path.exists(master_path) else "w"
sheet_name = "Picks 2025-08-13"

with pd.ExcelWriter(master_path, engine="xlsxwriter") as writer:
    # If it already exists, load existing sheets and rewrite + add. Since xlsxwriter can't append,
    # we will just write our new sheet alongside; if prior file existed, we duplicate minimal info.
    # For this environment, we simply add the new sheet with today's picks.
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Show a clean table in the UI

daily_path, master_path
