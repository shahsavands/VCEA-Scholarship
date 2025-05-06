
import streamlit as st
import pandas as pd
import smartsheet
import os

# --- CONFIG ---
SMARTSHEET_TOKEN = st.secrets["SMARTSHEET_TOKEN"]

DEPARTMENTS = [
    "Chemical Engineering and Bioengineering",
    "Civil and Environmental Engineering",
    "Design and Construction",
    "Electrical Engineering and Computer Science",
    "Mechanical and Materials Engineering"
]

ALLOWED_PLANS = [
    "Electrical Engineering", "Computer Science BS", "Computer Science",
    "Software Engineering", "Computer Science - Vancouver",
    "Computer Engineering", "Electrical Power Engineering"
]
ACAD_LEVEL_MAP = {10: "Freshman", 20: "Sophomore", 30: "Junior", 40: "Senior", 50: "Senior"}
AWARD_TIERS = {"high": 2000, "mid": 1000, "low": 500}

# --- FUNCTIONS ---
def process_student_data(df):
    df = df[~df["0290 Career"].str.contains("GRAD", na=False)]
    df = df[df["0360 CampusCd"] == "PULLM"]
    df = df[df["0100 Resident"] != "NON"]
    df = df[df["0205 AcadPlan"].isin(ALLOWED_PLANS)]
    df = df[pd.to_numeric(df["0370 WSUGPA"], errors='coerce') >= 2.5]

    tf_cols = [
        "1470 SingleParent", "1510 WorkForSupport", "1520 Disability",
        "1530 PoliticalAsylum", "1550 CancerTreatment", "1560 Homelessness",
        "1600 FirstGenStudent"
    ]
    df["need factor"] = df[tf_cols].apply(lambda row: sum(str(v).strip().lower() == "true" for v in row), axis=1)
    df["0280 AcadLevel"] = pd.to_numeric(df["0280 AcadLevel"], errors='coerce').map(ACAD_LEVEL_MAP)
    df = df.dropna(subset=["0280 AcadLevel"])

    keep_cols = [
        "Student ID (System Field)", "First Name", "Last Name", "Middle Name",
        "0090 Gender", "0100 Resident", "0110 County", "0130 EthnicDescr",
        "0135 Country", "0136 CitizenshipStatus", "0140 Email", "0225 AcadProg",
        "0280 AcadLevel", "0370 WSUGPA", "need factor"
    ]
    df = df[keep_cols]

    df["Gift ID"] = ""
    df["Award Name"] = ""
    df["Award Amount"] = ""
    df["Letter Received"] = "No"

    df = df.sort_values(by=["0370 WSUGPA", "need factor"], ascending=[False, False])
    return df

def download_scholarship_sheet(ss, sheet_id):
    sheet = ss.Sheets.get_sheet(sheet_id)
    col_map = {col.title: col.id for col in sheet.columns}
    data = []
    for row in sheet.rows:
        row_data = {}
        for cell in row.cells:
            for title, cid in col_map.items():
                if cell.column_id == cid:
                    row_data[title] = cell.display_value
        data.append(row_data)
    return pd.DataFrame(data), col_map, sheet.rows

def match_and_assign_scholarships(df_students, df_scholarships):
    df_scholarships["Remaining to Award"] = pd.to_numeric(df_scholarships["Remaining to Award"], errors='coerce').fillna(0)
    for idx, student in df_students.iterrows():
        gpa = float(student["0370 WSUGPA"])
        need = int(student["need factor"])
        prog = str(student["0225 AcadProg"]).split(",")[0].strip()
        level = student["0280 AcadLevel"]
        gender = str(student["0090 Gender"]).strip()
        county = str(student["0110 County"]).strip()

        if gpa > 3.79:
            award = AWARD_TIERS["high"]
        elif gpa >= 3.60:
            award = AWARD_TIERS["mid"]
        else:
            award = AWARD_TIERS["low"]

        for i, row in df_scholarships.iterrows():
            if row["Remaining to Award"] < award:
                continue

            prog_ok = "Any" in str(row["Program"]) or any(prog == p.strip() for p in str(row["Program"]).split(","))
            level_ok = "Any" in str(row["Level"]) or any(level == l.strip() for l in str(row["Level"]).split(","))
            location_ok = "County" not in str(row.get("Location", "")) or county.lower() == str(row["Location"]).strip().lower()
            gender_ok = not row.get("Gender") or gender == str(row["Gender"]).strip()
            need_ok = "Need-Based" not in str(row.get("Financial Status", "")) or need >= 1

            if prog_ok and level_ok and location_ok and gender_ok and need_ok:
                df_students.at[idx, "Gift ID"] = row["Workday Expendable Account #"]
                df_students.at[idx, "Award Name"] = row["Allocation Long Name"]
                df_students.at[idx, "Award Amount"] = award
                df_scholarships.at[i, "Remaining to Award"] -= award
                break
    return df_students, df_scholarships

def update_remaining_award_in_sheet(ss, sheet_id, df_sch, col_map, original_rows):
    updates = []
    for row in original_rows:
        alloc_name = None
        for cell in row.cells:
            if cell.column_id == col_map["Allocation Long Name"]:
                alloc_name = cell.display_value
                break
        if not alloc_name:
            continue
        match = df_sch[df_sch["Allocation Long Name"] == alloc_name]
        if not match.empty:
            new_val = float(match["Remaining to Award"].values[0])
            update_row = smartsheet.models.Row()
            update_row.id = row.id
            update_row.cells = [{'column_id': col_map["Remaining to Award"], 'value': new_val}]
            updates.append(update_row)
    if updates:
        ss.Sheets.update_rows(sheet_id, updates)

# --- STREAMLIT UI ---
st.set_page_config(page_title="WSU Scholarship Tool", layout="centered")
st.title("🎓 WSU Scholarship Assignment Tool")

with st.sidebar:
    st.image("wsu-shield-mark.png", width=250)
    dept = st.selectbox("Department", DEPARTMENTS)
    sheet_name = st.text_input("New Sheet Name")
    uploaded_file = st.file_uploader("Upload your student Excel file", type="xlsx")

if uploaded_file and sheet_name:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    processed_df = process_student_data(df)
    st.success(f"✅ Processed {len(processed_df)} student records.")
    st.dataframe(processed_df.head(20))

    csv = processed_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download CSV", csv, "processed_students.csv")

    ss = smartsheet.Smartsheet(SMARTSHEET_TOKEN)
    workspaces = {w.name: w.id for w in ss.Workspaces.list_workspaces().data}
    workspace_name = st.selectbox("Select Workspace", list(workspaces.keys()))
    workspace_id = workspaces[workspace_name]

    sheets = {s.name: s.id for s in ss.Workspaces.get_workspace(workspace_id).sheets}
    scholarship_name = st.selectbox("Select Scholarship Sheet", list(sheets.keys()))
    scholarship_id = sheets[scholarship_name]

    if st.button("Match & Upload"):
        df_sch, col_map, original_rows = download_scholarship_sheet(ss, scholarship_id)
        df_matched, df_sch_updated = match_and_assign_scholarships(processed_df, df_sch)
        df_matched.to_excel("final_students.xlsx", index=False)

        result = ss.Sheets.import_xlsx_sheet("final_students.xlsx", header_row_index=0, sheet_name=sheet_name)
        new_sheet_id = result.result.id
        ss.Sheets.move_sheet(new_sheet_id, smartsheet.models.ContainerDestination({
            'destination_type': 'workspace',
            'destination_id': workspace_id
        }))
        update_remaining_award_in_sheet(ss, scholarship_id, df_sch_updated, col_map, original_rows)
        st.success(f"✅ Uploaded and updated Smartsheet (Sheet ID: {new_sheet_id})")
