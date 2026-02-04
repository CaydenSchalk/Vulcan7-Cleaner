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

# ======================== Config ========================

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
    for idx, _ in df[column].items():
        worksheet.write(idx, df.columns.get_loc(column), df.at[idx, column], fmt_obj)

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
    return df

def format_phones(df):
    for i in range(1, 6):
        col = "Phone" if i == 1 else f"Phone {i}"
        df[col] = df[col].apply(format_phone_number)

def classify_dnc(df):
    global DNC_NUMBERS, NON_DNC_NUMBERS
    DNC_NUMBERS, NON_DNC_NUMBERS = {}, {}
    for i in range(1, 6):
        col = "Phone" if i == 1 else f"Phone {i}"
        dnc_col = f"{col} DNC Status"
        DNC_NUMBERS[col], NON_DNC_NUMBERS[col] = [], []
        for row_num, val in enumerate(df[dnc_col], start=1):
            (DNC_NUMBERS if val == 'DNC' else NON_DNC_NUMBERS)[col].append(row_num)
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
    global RESTRICTED_EMAILS
    RESTRICTED_EMAILS = {"Email": []}
    for idx, val in enumerate(df["Email Status"], start=1):
        if val == 'Restricted':
            RESTRICTED_EMAILS["Email"].append(idx)
    df.drop("Email Status", axis=1, inplace=True)

def flag_commercial(df):
    global COMMERCIAL_SALES
    COMMERCIAL_SALES = {"Property Type": []}
    for idx, val in enumerate(df["Property Type"], start=1):
        if val == 'Commercial Sale':
            COMMERCIAL_SALES["Property Type"].append(idx)

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

    # for col in ["Date Added", "Status Change Date"]:
    #     if col in df.columns:
    #         df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()
    #         # or: .dt.date (see note below)

    for col in CONVERT_TO_INT:
        df = convert_column_to_int(df, col)

def drop_empty_columns(df):
    empty_cols = [col for col in df.columns if df[col].isna().all() and col not in DONT_DROP_COLUMNS]
    df.drop(columns=empty_cols, inplace=True)

# ======================== Excel Output ========================
# def export_to_excel(df, output_file):
#     writer = pd.ExcelWriter(
#         output_file,
#         engine="xlsxwriter",
#         engine_kwargs={"options": {"nan_inf_to_errors": True}}
#     )
#
#     df.to_excel(writer, sheet_name="Sheet1", index=False)
#
#     workbook = writer.book
#     worksheet = writer.sheets["Sheet1"]
#
#     # Format setup
#     yellow = workbook.add_format({"bg_color": "#FFFF00"})
#     red = workbook.add_format({"bg_color": "#FF9999"})
#     currency = workbook.add_format({'num_format': '[$$-409]#,##0.00'})
#     comma = workbook.add_format({'num_format': '#,##0'})
#     caps_highlight = workbook.add_format({
#         "bg_color": "#F0E269",
#         "font_color": "#006100"
#     })
#     date_only = workbook.add_format({
#         "num_format": "mm/dd/yyyy"
#     })
#
#     # Email hyperlinks
#     email_col = ""
#
#     try:
#         email_col = column_number_to_letter(df.columns.get_loc("Email"))
#     except Exception as e:
#         pass
#
#     if email_col:
#         for row_num, email in enumerate(df["Email"], start=2):
#             if email:
#                 worksheet.write_url(f"{email_col}{row_num}", f"mailto:{email}")
#
#     # Color DNC and Commercials
#     color_rows(worksheet, df, DNC_NUMBERS, red)
#     color_rows(worksheet, df, NON_DNC_NUMBERS, yellow)
#     color_rows(worksheet, df, RESTRICTED_EMAILS, red, email=True)
#     color_rows(worksheet, df, COMMERCIAL_SALES, yellow)
#
#     # Apply formatting
#     apply_format_to_column(worksheet, df, "List Price", currency)
#     apply_format_to_column(worksheet, df, "Square Footage", comma)
#     # apply_format_to_column(worksheet, df, "Date Added", date_only)
#     # apply_format_to_column(worksheet, df, "Status Change Date", date_only)
#
#     auto_adjust_columns(worksheet, df)
#     add_excel_table(worksheet, df, date_only)
#     hide_specified_columns(worksheet, df)
#     highlight_all_caps(worksheet, df, "First Name", caps_highlight, other_column="Last Name")
#
#     set_column_date_only(worksheet, df, "Date Added", date_only)
#     set_column_date_only(worksheet, df, "Status Change Date", date_only)
#
#     writer.close()
#

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
    # for key in duplicate_keys:
    #     fn, ln = key.split("|")
    #
    #     sheet_df = df[df["_name_key"] == key].copy()
    #     sheet_df.drop(columns="_name_key", inplace=True)
    #     sheet_df.reset_index(drop=True, inplace=True)
    #
    #     # Excel sheet name rules:
    #     base_name = f"{fn.title()} {ln.title()}"[:31]
    #     sheet_name = make_unique_sheet_name(base_name, sheets)
    #
    #     sheets[sheet_name] = sheet_df

    # for key in duplicate_keys:
    #     fn, ln = key.split("|")
    #
    #     sheet_df = df[
    #         df.apply(
    #             lambda row: key in make_name_keys(row),
    #             axis=1
    #         )
    #     ].copy()
    #
    #     sheet_df.reset_index(drop=True, inplace=True)
    #
    #     base_name = f"{fn.title()} {ln.title()}"[:31]
    #     sheet_name = make_unique_sheet_name(base_name, sheets)
    #
    #     sheets[sheet_name] = sheet_df

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
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
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
            email_col = column_number_to_letter(sheet_df.columns.get_loc("Email"))
        except Exception:
            pass

        if email_col:
            for row_num, email in enumerate(sheet_df["Email"], start=2):
                if email:
                    worksheet.write_url(f"{email_col}{row_num}", f"mailto:{email}")

        # Color DNC and Commercials (UNCHANGED)
        color_rows(worksheet, sheet_df, DNC_NUMBERS, red)
        color_rows(worksheet, sheet_df, NON_DNC_NUMBERS, yellow)
        color_rows(worksheet, sheet_df, RESTRICTED_EMAILS, red, email=True)
        color_rows(worksheet, sheet_df, COMMERCIAL_SALES, yellow)

        # Apply formatting (UNCHANGED)
        apply_format_to_column(worksheet, sheet_df, "List Price", currency)
        apply_format_to_column(worksheet, sheet_df, "Square Footage", comma)

        auto_adjust_columns(worksheet, sheet_df)
        add_excel_table(worksheet, sheet_df, date_only)
        hide_specified_columns(worksheet, sheet_df)

        highlight_all_caps(
            worksheet,
            sheet_df,
            "First Name",
            caps_highlight,
            other_column="Last Name"
        )

        set_column_date_only(worksheet, sheet_df, "Date Added", date_only)
        set_column_date_only(worksheet, sheet_df, "Status Change Date", date_only)
        # ---- ORIGINAL BEHAVIOR ENDS HERE ----

    writer.close()


# def color_rows(ws, df, data_dict, fmt, email=False):
#     for col, rows in data_dict.items():
#         if col not in df.columns:
#             continue
#         for row in rows:
#             value = f"mailto:{df.at[row - 1, col]}" if email else df.at[row - 1, col]
#             if pd.notna(value):
#                 ws.write(row, df.columns.get_loc(col), value, fmt)

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


# def add_excel_table(ws, df):
#     rows, cols = len(df) + 1, len(df.columns)
#     col_letter = xl_col_to_name(cols - 1)
#     ws.add_table(f"A1:{col_letter}{rows}", {
#         "columns": [{"header": col} for col in df.columns],
#         "autofilter": True,
#         "style": "Table Style Medium 9"
#     })

# def add_excel_table(ws, df):
#     rows, cols = len(df) + 1, len(df.columns)
#     col_letter = xl_col_to_name(cols - 1)
#
#     ws.add_table(f"A1:{col_letter}{rows}", {
#         "columns": [
#             {"header": EXCEL_COLUMN_RENAMES.get(col, col)}
#             for col in df.columns
#         ],
#         "autofilter": True,
#         "style": "Table Style Medium 9"
#     })

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
# def main():
#     df = pd.read_csv(INPUT_FILE)
#     df = clean_data(df)
#     print("Final Column Count:", len(df.columns))
#     export_to_excel(df)
#
# if __name__ == "__main__":
#     main()


# def process_csv(csv_path):
#     try:
#         df = pd.read_csv(csv_path)
#         df = clean_data(df)
#
#         base, _ = os.path.splitext(csv_path)
#         output_path = base + ".xlsx"
#
#         export_to_excel(df, output_path)
#
#         messagebox.showinfo(
#             "Success",
#             f"File cleaned successfully!\n\nSaved as:\n{output_path}"
#         )
#
#     except Exception as e:
#         messagebox.showerror("Error", str(e))

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
#     # Windows sometimes wraps paths in {}
#     path = event.data.strip("{}")
#
#     if not path.lower().endswith(".csv"):
#         messagebox.showerror("Invalid File", "Please drop a CSV file.")
#         return
#
#     process_csv(path)

def on_drop(event):
    path = event.data.strip("{}")

    if not path.lower().endswith(".csv"):
        messagebox.showerror("Invalid File", "Please drop a CSV file.")
        return

    threading.Thread(
        target=process_csv,
        args=(path, progress_bar),
        daemon=True
    ).start()



# def launch_gui():
#     root = TkinterDnD.Tk()
#     root.title("CSV Cleaner")
#     root.geometry("420x180")
#     root.resizable(False, False)
#
#     label = tk.Label(
#         root,
#         text="Drag & drop your CSV file here",
#         font=("Segoe UI", 12),
#         relief="ridge",
#         borderwidth=2,
#         width=40,
#         height=6
#     )
#     label.pack(padx=20, pady=30)
#
#     label.drop_target_register(DND_FILES)
#     label.dnd_bind("<<Drop>>", on_drop)
#
#     root.mainloop()

def launch_gui():
    root = TkinterDnD.Tk()
    root.title("CSV Cleaner")
    root.geometry("420x220")
    root.resizable(False, False)

    label = tk.Label(
        root,
        text="Drag & drop your CSV file here",
        font=("Segoe UI", 12),
        relief="ridge",
        borderwidth=2,
        width=40,
        height=6
    )
    label.pack(padx=20, pady=(20, 10))

    label.drop_target_register(DND_FILES)
    label.dnd_bind("<<Drop>>", on_drop)

    global progress_bar
    progress_bar = ttk.Progressbar(
        root,
        mode="indeterminate",
        length=360
    )
    # run loop
    root.mainloop()




if __name__ == "__main__":
    launch_gui()