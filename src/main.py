# -*- coding: utf-8 -*-
"""
Data Cleaning and Excel Export Tool
Created on Fri Mar 14 02:38:28 2025
@author: Cayden
"""

# ======================== Imports ========================
import pandas as pd
import numpy as np
import re
import os
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from xlsxwriter.utility import xl_col_to_name
import traceback
from pathlib import Path
from utils.avery_labels import fill_avery_30up

# ======================== Config ========================

PROJECT_ROOT = Path(__file__).resolve().parents[1]  # src/ -> project_root/
OUTPUT_DIR = PROJECT_ROOT / "output"
TEMPLATE_DIR = PROJECT_ROOT / "templates"
OUTPUT_DIR.mkdir(exist_ok=True)

COLUMNS_TO_DROP = ["Member First Name", "Member Last Name", "Member Username", "Listing Office", "Listing Agent", "MLS Name", "Folder"]
COLUMNS_TO_MOVE = ["Email", "Address1", "Address2", "City", "State", "Zip", "Owner Mailing Address", "Owner Mailing City", "Owner Mailing State", "Owner Mailing Zip", "Tax Owner", "MLS Number"]
ADJUST_COLUMNS_EXCEPT = ["Remarks"]
CONVERT_TO_INT = ["Contact ID", "Zip", "Owner Mailing Zip", "MLS Number", "Year Built", "Days On Market", "Bathrooms", "Bedrooms", "Square Footage"]
DONT_DROP_COLUMNS = ["Total Calls", "Last Call", "First Call Date", "Phone Label", "Phone 2 Label", "Phone 3 Label", "Phone 4 Label", "Phone 5 Label", "Tags"]
APPEND_AT_END = ["Tags", "First Call Date", "Last Call", "Total Calls"]
HIDE_COLUMNS = ["Lead Source", "Phone Label", "Phone 2 Label", "Phone 3 Label", "Phone 4 Label", "Phone 5 Label"]

COLUMN_PADDING = {
    "List Price" : 5,
    "Tags" : 2,
}

US_STATE_ABBREV = {
    "Alabama": "AL",
    "Alaska": "AK",
    "Arizona": "AZ",
    "Arkansas": "AR",
    "California": "CA",
    "Colorado": "CO",
    "Connecticut": "CT",
    "Delaware": "DE",
    "Florida": "FL",
    "Georgia": "GA",
    "Hawaii": "HI",
    "Idaho": "ID",
    "Illinois": "IL",
    "Indiana": "IN",
    "Iowa": "IA",
    "Kansas": "KS",
    "Kentucky": "KY",
    "Louisiana": "LA",
    "Maine": "ME",
    "Maryland": "MD",
    "Massachusetts": "MA",
    "Michigan": "MI",
    "Minnesota": "MN",
    "Mississippi": "MS",
    "Missouri": "MO",
    "Montana": "MT",
    "Nebraska": "NE",
    "Nevada": "NV",
    "New Hampshire": "NH",
    "New Jersey": "NJ",
    "New Mexico": "NM",
    "New York": "NY",
    "North Carolina": "NC",
    "North Dakota": "ND",
    "Ohio": "OH",
    "Oklahoma": "OK",
    "Oregon": "OR",
    "Pennsylvania": "PA",
    "Rhode Island": "RI",
    "South Carolina": "SC",
    "South Dakota": "SD",
    "Tennessee": "TN",
    "Texas": "TX",
    "Utah": "UT",
    "Vermont": "VT",
    "Virginia": "VA",
    "Washington": "WA",
    "West Virginia": "WV",
    "Wisconsin": "WI",
    "Wyoming": "WY",
    "District of Columbia": "DC",
    "American Samoa": "AS",
    "Guam": "GU",
    "Northern Mariana Islands": "MP",
    "Puerto Rico": "PR",
    "United States Minor Outlying Islands": "UM",
    "Virgin Islands, U.S.": "VI",
}

EXCEL_COLUMN_RENAMES = {
    "Square Footage" : "Sqft",
    "Owner Occupied" : "Owner Occ",
    "Hot Prospect Points" : "Hot PP",
    "Days On Market" : "Days OM",
    "Owner Mailing State" : "Owner MS",
    "Owner Mailing Zip" : "Owner MZ",
}

# ======================== Utility Functions ========================
def format_phone_number(phone_number):
    if pd.isna(phone_number): return None
    phone_number_str = str(int(phone_number))
    if not re.fullmatch(r"\d{10}", phone_number_str): return None
    return f"({phone_number_str[:3]}) {phone_number_str[3:6]}-{phone_number_str[6:]}"

def mail_to_format(email):
    if pd.isna(email) or email in ["", "nan", None] or (isinstance(email, float) and np.isnan(email)):
        return None
    return email

def move_column(df, col, pos):
    col_data = df.pop(col)
    df.insert(pos, col, col_data)

def convert_binary_to_yes_no(val):
    return "Yes" if val == 1 else "No"

def column_number_to_letter(col_num):
    letters = ""
    while col_num >= 0:
        letters = chr(col_num % 26 + 65) + letters
        col_num = col_num // 26 - 1
    return letters

def convert_state_abbreviation(name):
    return US_STATE_ABBREV.get(name, name)

def convert_column_to_int(df, column):
    df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0).astype(int)
    return df

def apply_format_to_column(worksheet, df, column, fmt_obj):
    if column not in df.columns:
        return
    col_idx = df.columns.get_loc(column)
    for row_idx in range(len(df)):
        worksheet.write(row_idx + 1, col_idx, df.iloc[row_idx, col_idx], fmt_obj)


def make_avery_label_entry(row : pd.Series) -> str:
    first_name = row["First Name"]
    last_name = row["Last Name"]
    street = row["Address1"]
    city = row["City"]
    state = row["State"]
    zip_code = row["Zip"]

    string = f"{first_name} {last_name}\n{street}\n{city}, {state} {zip_code}"
    return string

def highlight_all_caps(ws, df, column_name, fmt, other_column=None):
    if column_name not in df.columns:
        return

    if other_column and other_column not in df.columns:
        return

    col_idx = df.columns.get_loc(column_name)
    col_letter = column_number_to_letter(col_idx)

    if other_column:
        other_col_idx = df.columns.get_loc(other_column)
        other_col_letter = column_number_to_letter(other_col_idx)

    start_row = 2
    end_row = len(df) + 1

    if other_column:
        formula = (
            f'=AND('
            f'EXACT({col_letter}{start_row}, UPPER({col_letter}{start_row})), '
            f'{col_letter}{start_row}<>"", '
            f'ISNUMBER(SEARCH(".", {other_col_letter}{start_row}))'
            f')'
        )
    else:
        formula = (
            f'=AND('
            f'EXACT({col_letter}{start_row}, UPPER({col_letter}{start_row})), '
            f'{col_letter}{start_row}<>""'
            f')'
        )

    ws.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {
            "type": "formula",
            "criteria": formula,
            "format": fmt
        }
    )

def set_column_date_only(ws, df, column_name, fmt):
    if column_name not in df.columns:
        return

    col_idx = df.columns.get_loc(column_name)
    col_letter = column_number_to_letter(col_idx)

    ws.set_column(f"{col_letter}:{col_letter}", None, fmt)

def is_an_entity(val):
    if pd.isna(val):
        return False
    val = str(val).strip()
    return val != "" and val.isupper()

def make_unique_sheet_name(base, existing):
    name = base
    i = 2
    while name in existing:
        name = f"{base} ({i})"
        i += 1
    return name[:31]

def make_name_keys(row):
    keys = []

    if pd.notna(row["First Name"]) and pd.notna(row["Last Name"]):
        keys.append(
            f"{row['First Name'].strip().upper()}|{row['Last Name'].strip().upper()}"
        )

    if (
        "First Name 2" in row
        and "Last Name 2" in row
        and pd.notna(row["First Name 2"])
        and pd.notna(row["Last Name 2"])
    ):
        keys.append(
            f"{row['First Name 2'].strip().upper()}|{row['Last Name 2'].strip().upper()}"
        )

    return keys

def choose_owner_key(row, duplicate_keys):
    """
    From all duplicate name keys present in this row,
    choose the one with the highest alphabetical FIRST NAME.
    """
    keys = []

    # Primary name
    if pd.notna(row["First Name"]) and pd.notna(row["Last Name"]):
        keys.append(
            f"{row['First Name'].strip().upper()}|{row['Last Name'].strip().upper()}"
        )

    # Secondary name
    if (
        "First Name 2" in row
        and "Last Name 2" in row
        and pd.notna(row["First Name 2"])
        and pd.notna(row["Last Name 2"])
    ):
        keys.append(
            f"{row['First Name 2'].strip().upper()}|{row['Last Name 2'].strip().upper()}"
        )

    # Keep only duplicate-related keys
    keys = [k for k in keys if k in duplicate_keys]

    if not keys:
        return None

    # Sort by FIRST NAME (before "|"), alphabetically
    return max(keys, key=lambda k: k.split("|")[0])



# ======================== Processing Steps ========================
def clean_data(df):
    df.dropna(thresh=5, axis=0, inplace=True)
    df.reset_index(drop=True, inplace=True)
    format_phones(df)
    classify_dnc(df)
    append_phone_types(df)
    restrict_emails(df)
    flag_commercial(df)
    drop_and_reorder_columns(df)
    transform_columns(df)
    drop_empty_columns(df)


    labels = []
    for index, row in df.iterrows():
        if index >= 30:
            break
        labels.append(make_avery_label_entry(row))

    fill_avery_30up(
        template_path=str(TEMPLATE_DIR / "Avery5160AddressLabels.docx"),
        output_path=str(OUTPUT_DIR / "filled_labels.docx"),
        labels = labels
    )

    return df

def format_phones(df):
    for i in range(1, 6):
        col = "Phone" if i == 1 else f"Phone {i}"
        df[col] = df[col].apply(format_phone_number)

def classify_dnc(df):
    global DNC_NUMBERS, NON_DNC_NUMBERS
    DNC_NUMBERS, NON_DNC_NUMBERS = {}, {}
    # In classify_dnc, instead of building DNC_NUMBERS/NON_DNC_NUMBERS dicts:
    for i in range(1, 6):
        col = "Phone" if i == 1 else f"Phone {i}"
        dnc_col = f"{col} DNC Status"
        df[f"_dnc_{col}"] = df[dnc_col] == "DNC"
        df.drop(dnc_col, axis=1, inplace=True)

def append_phone_types(df):
    for i in range(1, 6):
        col = "Phone" if i == 1 else f"Phone {i}"
        type_col, label_col = f"{col} Type", f"{col} Label"
        df[label_col] = df[label_col].astype("string")
        for idx, val in enumerate(df[type_col], start=1):
            if pd.notna(val):
                current_label = df.loc[idx - 1, label_col]
                new_val = f"{current_label} ({val})" if pd.notna(current_label) else f"({val})"
                df.loc[idx - 1, label_col] = new_val
        df.drop(type_col, axis=1, inplace=True)

def restrict_emails(df):
    df["_restricted_email"] = df["Email Status"] == "Restricted"
    df.drop("Email Status", axis=1, inplace=True)

def flag_commercial(df):
    df["_commercial"] = df["Property Type"] == "Commercial Sale"

def drop_and_reorder_columns(df):
    df.drop(COLUMNS_TO_DROP, axis=1, inplace=True)
    idx = df.columns.get_loc("Last Name 2")
    for col in reversed(COLUMNS_TO_MOVE):
        move_column(df, col, idx + 1)
    for col in APPEND_AT_END:
        move_column(df, col, len(df.columns) - 1)

def transform_columns(df):
    df["Owner Occupied"] = df["Owner Occupied"].apply(convert_binary_to_yes_no)
    df["Email"] = df["Email"].apply(mail_to_format)
    df["State"] = df["State"].apply(convert_state_abbreviation)
    df["Owner Mailing State"] = df["Owner Mailing State"].apply(convert_state_abbreviation)
    df["List Price"] = df["List Price"].astype(float)

    for col in CONVERT_TO_INT:
        df = convert_column_to_int(df, col)

def drop_empty_columns(df):
    empty_cols = [col for col in df.columns if df[col].isna().all() and col not in DONT_DROP_COLUMNS]
    df.drop(columns=empty_cols, inplace=True)

# ======================== Excel Output ========================

def export_to_excel(df, output_file):
    # -------- entity split (presentation only) --------
    entity_mask = (
        df["First Name"].apply(is_an_entity)
        & df["Last Name"].apply(lambda x: x == ".")
    )

    df["_name_key"] = (
            df["First Name"].str.strip().str.upper()
            + "|"
            + df["Last Name"].str.strip().str.upper()
    )

    name_map = (
        df.assign(_name_keys=df.apply(make_name_keys, axis=1))
        .explode("_name_keys")
        .dropna(subset=["_name_keys"])
    )

    duplicate_keys = (
        name_map["_name_keys"]
        .value_counts()
        .loc[lambda x: x > 1]
        .index
    )

    df["_owner_key"] = df.apply(
        choose_owner_key,
        axis=1,
        duplicate_keys=duplicate_keys
    )

    duplicate_mask = df["_owner_key"].notna()

    sheets = {
        "Main": df[~entity_mask & ~duplicate_mask].copy(),
        "Companies": df[entity_mask].copy()
    }

    for key in duplicate_keys:
        fn, ln = key.split("|")

        sheet_df = df[df["_owner_key"] == key].copy()
        if sheet_df.empty:
            continue

        sheet_df.drop(columns=["_name_key", "_owner_key"], inplace=True, errors="ignore")
        sheet_df.reset_index(drop=True, inplace=True)

        base_name = f"{fn.title()} {ln.title()}"[:31]
        sheet_name = make_unique_sheet_name(base_name, sheets)

        sheets[sheet_name] = sheet_df

    writer = pd.ExcelWriter(
        output_file,
        engine="xlsxwriter",
        engine_kwargs={"options": {"nan_inf_to_errors": True}}
    )

    workbook = writer.book

    for sheet_name, sheet_df in sheets.items():
        if sheet_df.empty:
            continue

        # ---- ORIGINAL BEHAVIOR STARTS HERE ----

        flag_cols = [c for c in sheet_df.columns if c.startswith("_")]
        display_df = sheet_df.drop(columns=flag_cols)
        display_df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]

        # Format setup (UNCHANGED)
        yellow = workbook.add_format({"bg_color": "#FFFF00"})
        red = workbook.add_format({"bg_color": "#FF9999"})
        currency = workbook.add_format({'num_format': '[$$-409]#,##0.00'})
        comma = workbook.add_format({'num_format': '#,##0'})
        caps_highlight = workbook.add_format({
            "bg_color": "#F0E269",
            "font_color": "#006100"
        })
        date_only = workbook.add_format({
            "num_format": "mm/dd/yyyy"
        })

        # Email hyperlinks (UNCHANGED)
        email_col = ""
        try:
            email_col = column_number_to_letter(display_df.columns.get_loc("Email"))
        except Exception:
            pass

        if email_col:
            for row_num, email in enumerate(display_df["Email"], start=2):
                if email:
                    worksheet.write_url(f"{email_col}{row_num}", f"mailto:{email}")

        color_rows_from_flags(worksheet, display_df, sheet_df, red, yellow, red, yellow)

        apply_format_to_column(worksheet, display_df, "List Price", currency)
        apply_format_to_column(worksheet, display_df, "Square Footage", comma)
        auto_adjust_columns(worksheet, display_df)
        add_excel_table(worksheet, display_df, date_only)
        hide_specified_columns(worksheet, display_df)
        highlight_all_caps(worksheet, display_df, "First Name", caps_highlight, other_column="Last Name")
        set_column_date_only(worksheet, display_df, "Date Added", date_only)
        set_column_date_only(worksheet, display_df, "Status Change Date", date_only)

        # ---- ORIGINAL BEHAVIOR ENDS HERE ----

    writer.close()

def color_rows(ws, df, data_dict, fmt, email=False):
    for col, rows in data_dict.items():
        if col not in df.columns:
            continue

        for row in rows:
            idx = row - 1
            if idx >= len(df):
                continue

            value = (
                f"mailto:{df.iloc[idx][col]}"
                if email else df.iloc[idx][col]
            )

            if pd.notna(value):
                ws.write(row, df.columns.get_loc(col), value, fmt)

def color_rows_from_flags(ws, display_df, flag_df, fmt_dnc, fmt_non_dnc, fmt_restricted, fmt_commercial):
    for i in range(1, 6):
        col = "Phone" if i == 1 else f"Phone {i}"
        flag_col = f"_dnc_{col}"
        if flag_col not in flag_df.columns or col not in display_df.columns:
            continue
        col_idx = display_df.columns.get_loc(col)
        for row_idx in range(len(display_df)):
            val = display_df.iloc[row_idx][col]
            if pd.notna(val):
                fmt = fmt_dnc if flag_df.iloc[row_idx][flag_col] else fmt_non_dnc
                ws.write(row_idx + 1, col_idx, val, fmt)

    if "_restricted_email" in flag_df.columns and "Email" in display_df.columns:
        col_idx = display_df.columns.get_loc("Email")
        for row_idx in range(len(display_df)):
            if flag_df.iloc[row_idx]["_restricted_email"]:
                val = display_df.iloc[row_idx]["Email"]
                if pd.notna(val):
                    ws.write(row_idx + 1, col_idx, f"mailto:{val}", fmt_restricted)

    if "_commercial" in flag_df.columns and "Property Type" in display_df.columns:
        col_idx = display_df.columns.get_loc("Property Type")
        for row_idx in range(len(display_df)):
            if flag_df.iloc[row_idx]["_commercial"]:
                val = display_df.iloc[row_idx]["Property Type"]
                if pd.notna(val):
                    ws.write(row_idx + 1, col_idx, val, fmt_commercial)


def auto_adjust_columns(ws, df):
    for col_num, col_name in enumerate(df.columns):
        if col_name in ADJUST_COLUMNS_EXCEPT:
            continue

        display_name = EXCEL_COLUMN_RENAMES.get(col_name, col_name)
        max_len = max(
            df[col_name].astype(str).map(len).max(),
            len(display_name)
        )

        if col_name in COLUMN_PADDING:
            max_len += COLUMN_PADDING[col_name]

        ws.set_column(col_num, col_num, max_len)

def add_excel_table(ws, df, date_fmt):
    rows, cols = len(df) + 1, len(df.columns)
    col_letter = xl_col_to_name(cols - 1)

    columns = []
    for col in df.columns:
        col_def = {
            "header": EXCEL_COLUMN_RENAMES.get(col, col)
        }

        if col in ("Date Added", "Status Change Date"):
            col_def["format"] = date_fmt

        columns.append(col_def)

    ws.add_table(f"A1:{col_letter}{rows}", {
        "columns": columns,
        "style": "Table Style Medium 9"
    })




def hide_specified_columns(ws, df):
    for col in HIDE_COLUMNS:
        letter = column_number_to_letter(df.columns.get_loc(col))
        ws.set_column(f"{letter}:{letter}", None, None, {"hidden": True})

# ======================== Run Pipeline ========================

def process_csv(csv_path, progress):
    try:
        progress.pack(pady=10)
        progress.start(10)

        df = pd.read_csv(csv_path)
        df = clean_data(df)

        base, _ = os.path.splitext(csv_path)
        output_path = base + ".xlsx"

        export_to_excel(df, output_path)

        progress.stop()
        progress.pack_forget()

        messagebox.showinfo(
            "Success",
            f"File cleaned successfully!\n\nSaved as:\n{output_path}"
        )


    except Exception as e:

        progress.stop()

        progress.pack_forget()

        tb = traceback.format_exc()

        messagebox.showerror(

            "Error",

            f"{type(e).__name__}: {e}\n\nTraceback:\n{tb}"

        )

# def on_drop(event):
#     path = event.data.strip("{}")
#
#     if not path.lower().endswith(".csv"):
#         messagebox.showerror("Invalid File", "Please drop a CSV file.")
#         return
#
#     threading.Thread(
#         target=process_csv,
#         args=(path, progress_bar),
#         daemon=True
#     ).start()

def launch_gui():
    root = TkinterDnD.Tk()
    root.title("CSV Cleaner")
    root.geometry("420x300")
    root.resizable(False, False)

    mode_var = tk.StringVar(value="clean")

    mode_frame = tk.Frame(root)
    mode_frame.pack(pady=(15, 0))

    tk.Label(mode_frame, text="Mode:", font=("Segoe UI", 10)).pack(side="left", padx=(0, 8))
    tk.Radiobutton(mode_frame, text="Clean CSV → Excel", variable=mode_var,
                   value="clean", font=("Segoe UI", 10)).pack(side="left", padx=4)
    tk.Radiobutton(mode_frame, text="Excel → Recipients CSV", variable=mode_var,
                   value="recipients", font=("Segoe UI", 10)).pack(side="left", padx=4)

    hint_var = tk.StringVar(value="Drag & drop your CSV file here")

    def on_mode_change(*_):
        if mode_var.get() == "clean":
            hint_var.set("Drag & drop your CSV file here")
        else:
            hint_var.set("Drag & drop your cleaned .xlsx file here")

    mode_var.trace_add("write", on_mode_change)

    label = tk.Label(
        root,
        textvariable=hint_var,
        font=("Segoe UI", 12),
        relief="ridge",
        borderwidth=2,
        width=40,
        height=6
    )
    label.pack(padx=20, pady=(12, 10))

    def extract_recipients_csv(xlsx_path, progress):
        try:
            progress.pack(pady=10)
            progress.start(10)

            xl = pd.ExcelFile(xlsx_path)
            if "Main" not in xl.sheet_names:
                raise ValueError("No 'Main' sheet found in this workbook.")

            df = pd.read_excel(xlsx_path, sheet_name="Main")

            # Reverse the display renames back to original column names
            reverse_renames = {v: k for k, v in EXCEL_COLUMN_RENAMES.items()}
            df.rename(columns=reverse_renames, inplace=True)

            # Require email
            df = df[df["Email"].notna() & (df["Email"].astype(str).str.strip() != "")]

            out = pd.DataFrame()
            out["email"] = df["Email"].astype(str).str.strip()

            if "First Name" in df.columns and "Last Name" in df.columns:
                out["name"] = (
                        df["First Name"].fillna("").astype(str).str.strip()
                        + " "
                        + df["Last Name"].fillna("").astype(str).str.strip()
                ).str.strip()
            elif "First Name" in df.columns:
                out["name"] = df["First Name"].fillna("").astype(str).str.strip()

            # Pull in other useful columns if they exist
            for col in ["Phone", "Address1", "City", "State", "Zip",
                        "Tax Owner", "List Price", "MLS Number"]:
                if col in df.columns:
                    out[col.lower().replace(" ", "_")] = df[col]

            base = os.path.splitext(xlsx_path)[0]
            output_path = base + "_recipients.csv"
            out.to_csv(output_path, index=False, encoding="utf-8")

            progress.stop()
            progress.pack_forget()
            messagebox.showinfo(
                "Success",
                f"Recipients CSV created!\n{len(out)} rows exported.\n\nSaved as:\n{output_path}"
            )

        except Exception as e:
            progress.stop()
            progress.pack_forget()
            tb = traceback.format_exc()
            messagebox.showerror("Error", f"{type(e).__name__}: {e}\n\nTraceback:\n{tb}")

    def on_drop(event):
        path = event.data.strip("{}")
        if mode_var.get() == "clean":
            if not path.lower().endswith(".csv"):
                messagebox.showerror("Invalid File", "Please drop a CSV file.")
                return
            threading.Thread(target=process_csv, args=(path, progress_bar), daemon=True).start()
        else:
            if not path.lower().endswith(".xlsx"):
                messagebox.showerror("Invalid File", "Please drop an .xlsx file.")
                return
            threading.Thread(target=extract_recipients_csv, args=(path, progress_bar), daemon=True).start()

    label.drop_target_register(DND_FILES)
    label.dnd_bind("<<Drop>>", on_drop)

    global progress_bar
    progress_bar = ttk.Progressbar(root, mode="indeterminate", length=360)
    root.mainloop()

if __name__ == "__main__":
    launch_gui()