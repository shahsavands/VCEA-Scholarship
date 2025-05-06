
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import pandas as pd
import smartsheet
import os

# --------------------------- CONFIG ---------------------------
SMARTSHEET_TOKEN = 'r7cpFSozLL2DOJ8Kg4rH0sK5fYFfyNdVCFXgR'
LOGO_PATH = 'wsu-shield-mark.png'

WSU_MAROON = '#981e32'
WSU_GRAY = '#5e6a71'
BG_COLOR = '#f7f7f7'

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

def process_student_data(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')
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

class ScholarshipApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WSU Scholarship Assignment Tool")
        self.root.geometry("700x500")
        self.root.configure(bg=BG_COLOR)

        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.selected_department = tk.StringVar(value=DEPARTMENTS[3])

        self.ss = smartsheet.Smartsheet(SMARTSHEET_TOKEN)
        self.workspaces = self.get_workspaces()
        self.workspace_choice = tk.StringVar()
        self.scholarship_sheets = {}
        self.scholarship_choice = tk.StringVar()

        self.build_ui()

    def get_workspaces(self):
        return {w.name: w.id for w in self.ss.Workspaces.list_workspaces().data}

    def get_scholarship_sheets(self, workspace_id):
        workspace = self.ss.Workspaces.get_workspace(workspace_id)
        return {s.name: s.id for s in workspace.sheets}

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.excel_path.set(path)

    def update_scholarship_dropdown(self, event):
        workspace_id = self.workspaces.get(self.workspace_choice.get())
        if workspace_id:
            self.scholarship_sheets = self.get_scholarship_sheets(workspace_id)
            self.scholarship_menu['values'] = list(self.scholarship_sheets.keys())
            if self.scholarship_sheets:
                self.scholarship_menu.current(0)
                self.scholarship_choice.set(list(self.scholarship_sheets.keys())[0])

    def build_ui(self):
        pad = {'padx': 10, 'pady': 8}
        logo = Image.open(LOGO_PATH)
        logo = logo.resize((90, 90), Image.Resampling.LANCZOS)
        self.logo_img = ImageTk.PhotoImage(logo)
        tk.Label(self.root, image=self.logo_img, bg=BG_COLOR).grid(row=0, column=0, columnspan=3, pady=(15, 5))

        tk.Label(self.root, text="Department:", bg=BG_COLOR, fg=WSU_GRAY, font=("Segoe UI", 10, 'bold')).grid(row=1, column=0, sticky='e', **pad)
        self.department_menu = ttk.Combobox(self.root, textvariable=self.selected_department, width=45)
        self.department_menu['values'] = DEPARTMENTS
        self.department_menu.grid(row=1, column=1, columnspan=2, **pad)

        tk.Label(self.root, text="Excel File:", bg=BG_COLOR, fg=WSU_GRAY).grid(row=2, column=0, sticky='e', **pad)
        tk.Entry(self.root, textvariable=self.excel_path, width=45).grid(row=2, column=1, **pad)
        browse_btn = tk.Button(self.root, text="Browse", command=self.browse_file,
                       font=('Segoe UI', 10, 'bold'), relief='raised', bd=2)
        browse_btn.grid(row=2, column=2, padx=10, pady=8)


        tk.Label(self.root, text="New Sheet Name:", bg=BG_COLOR, fg=WSU_GRAY).grid(row=3, column=0, sticky='e', **pad)
        tk.Entry(self.root, textvariable=self.sheet_name, width=45).grid(row=3, column=1, columnspan=2, **pad)

        tk.Label(self.root, text="Workspace:", bg=BG_COLOR, fg=WSU_GRAY).grid(row=4, column=0, sticky='e', **pad)
        self.workspace_menu = ttk.Combobox(self.root, textvariable=self.workspace_choice, width=45)
        self.workspace_menu['values'] = list(self.workspaces.keys())
        self.workspace_menu.grid(row=4, column=1, columnspan=2, **pad)
        self.workspace_menu.bind("<<ComboboxSelected>>", self.update_scholarship_dropdown)

        tk.Label(self.root, text="Scholarship Sheet:", bg=BG_COLOR, fg=WSU_GRAY).grid(row=5, column=0, sticky='e', **pad)
        self.scholarship_menu = ttk.Combobox(self.root, textvariable=self.scholarship_choice, width=45)
        self.scholarship_menu.grid(row=5, column=1, columnspan=2, **pad)

        upload_btn = tk.Button(self.root, text="Process and Upload", command=self.process_upload,
                       font=('Segoe UI', 10, 'bold'), relief='raised', bd=2, width=25)
        upload_btn.grid(row=6, column=1, columnspan=1, pady=20)

    def process_upload(self):
        try:
            file = self.excel_path.get()
            name = self.sheet_name.get().strip()
            workspace_id = self.workspaces[self.workspace_choice.get()]
            scholarship_id = self.scholarship_sheets[self.scholarship_choice.get()]
            df_students = process_student_data(file)
            df_sch, col_map, original_rows = download_scholarship_sheet(self.ss, scholarship_id)
            df_final, df_sch_updated = match_and_assign_scholarships(df_students, df_sch)
            out_file = os.path.splitext(file)[0] + "_final.xlsx"
            df_final.to_excel(out_file, index=False)
            response = self.ss.Sheets.import_xlsx_sheet(out_file, header_row_index=0, sheet_name=name)
            sheet_id = response.result.id
            self.ss.Sheets.move_sheet(sheet_id, smartsheet.models.ContainerDestination({
                'destination_type': 'workspace',
                'destination_id': workspace_id
            }))
            update_remaining_award_in_sheet(self.ss, scholarship_id, df_sch_updated, col_map, original_rows)
            messagebox.showinfo("Success", f"Uploaded to Smartsheet (Sheet ID: {sheet_id})")
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Start GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = ScholarshipApp(root)
    root.mainloop()

