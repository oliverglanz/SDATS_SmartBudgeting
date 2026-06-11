#!/usr/bin/env python
# coding: utf-8

# # Smart Contract Automation Notebook
# This notebook automates contract extraction and PDF form generation.

# In[1]:


import pandas as pd
import numpy as np
import math
import re
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject


# In[2]:


# Making sure we are working within the same directory as the jupyter notebook (this will allow for poratbility of the project)
import os
## Verify we're in the right directory
print("Current directory:", os.getcwd())


# In[3]:


import pandas as pd

# Show all columns
pd.set_option('display.max_columns', None)

# Show all rows (optional — be careful if your BannerPrevious is huge)
pd.set_option('display.max_rows', None)

# Don’t truncate long strings
pd.set_option('display.max_colwidth', None)

# Make the display wider in the notebook
pd.set_option('display.width', 0)


# In[4]:


from pathlib import Path


def get_real_excels():
    folder = Path("Step3_input_contract-production_ShariSheets_filled")
    excel_files = []
    for f in sorted(folder.glob("*.xlsx")):
        if f.name.startswith("._"):  # skip macOS junk files
            continue
        excel_files.append(f)
    if not excel_files:
        raise FileNotFoundError("No valid Excel files found in Step3_input_contract-production_ShariSheets_filled")
    return excel_files


EXCEL_PATHS = get_real_excels()
EXCEL_PATH = EXCEL_PATHS[0]  # kept for preview/debug cells that expect one workbook



def get_real_pdf(pattern):
    folder = Path("0_source_files/default_PDFcontract")
    for pdf in folder.glob(pattern):
        if pdf.name.startswith("._"):
            continue
        return pdf
    raise FileNotFoundError(f"No valid PDF found for pattern: {pattern}")


print("USING FILTERED GET_REAL_PDF() - dotfiles will be skipped")


PDF_standard_TEMPLATE_PATH  = get_real_pdf("*standard*.pdf")
PDF_intensive_TEMPLATE_PATH = get_real_pdf("*intensive*.pdf")


#PDF_standard_TEMPLATE_PATH  = next(Path("Step3_input_CONTRACTS_empty").glob("*standard*.pdf"))
#PDF_intensive_TEMPLATE_PATH = next(Path("Step3_input_CONTRACTS_empty").glob("*intensive*.pdf"))


OUTPUT_DIR = Path("Step3_output_contract-production_PDF-CONTRACTS_filled")
OUTPUT_DIR.mkdir(exist_ok=True)


MAX_COURSES_PER_CONTRACT = 3
DEFAULT_RATE_PER_CREDIT = 1050


print("Excel files:")
for excel_path in EXCEL_PATHS:
    print(" -", excel_path)
print("Preview Excel:", EXCEL_PATH)
print("Standard:", PDF_standard_TEMPLATE_PATH)
print("Intensive:", PDF_intensive_TEMPLATE_PATH)


# # What are the fillable fields in the PDF?

# In[5]:


from pypdf import PdfReader

reader = PdfReader(PDF_intensive_TEMPLATE_PATH)
fields = reader.get_fields()

print("=== FORM FIELD LIST ===")
for field_name, field_data in fields.items():
    print(f"Field Name : {field_name}")
    print(f"Field Type : {field_data.get('/FT')}")
    print(f"Default    : {field_data.get('/V')}")
    print("-" * 40)


# In[6]:


from PyPDF2 import PdfReader

reader = PdfReader(PDF_intensive_TEMPLATE_PATH)

fields = reader.get_fields()
print(fields)


# In[7]:


from PyPDF2 import PdfReader

def get_pdf_fields(pdf_path):
    """
    Print all PDF form fields and their appearance states.
    """
    reader = PdfReader(str(pdf_path))

    if "/AcroForm" not in reader.trailer["/Root"]:
        print("No AcroForm found.")
        return

    form = reader.trailer["/Root"]["/AcroForm"]
    fields = form.get("/Fields", [])

    print("=== PDF FORM FIELDS ===\n")

    for f in fields:
        field = f.get_object()
        name = field.get("/T")
        ftype = field.get("/FT")
        value = field.get("/V")
        kids = field.get("/Kids")

        print(f"FIELD: {name}")
        print(f"  TYPE: {ftype}")
        print(f"  VALUE: {value}")
        print(f"  KIDS: {kids}\n")

        # Show appearance states if present
        ap = field.get("/AP")
        if ap and "/N" in ap:
            print("  --- Appearance States ---")
            for state_name, appearance in ap["/N"].items():
                print(f"    • {state_name}: {appearance}")
        else:
            print("  --- No Appearance States Found ---")

        print("\n" + "-"*50 + "\n")


# In[8]:


from PyPDF2 import PdfReader

reader = PdfReader(PDF_standard_TEMPLATE_PATH)

fields = reader.get_fields()
for name, field in fields.items():
    print("FIELD:", name)
    print("  TYPE:", field.get('/FT'))
    print("  VALUE:", field.get('/V'))
    print("  OPTIONS:", field.get('/Opt'))
    print("  STATES:", field.get('/AP'))
    print()


# # Preparing the DF

# In[9]:


def split_name(name):
    if not isinstance(name, str) or "," not in name:
        return name, ""
    last, first = name.split(",", 1)
    return last.strip(), first.strip()

def chunk_list(lst, size):
    for i in range(0, len(lst), size):
        yield lst[i:i + size]


# In[10]:


df_raw = pd.read_excel(EXCEL_PATH, sheet_name=0) 
df_raw.head()


# In[11]:


list(df_raw.columns)


# In[12]:


print(df_raw.columns.tolist())


# In[13]:


df_raw.head()


# In[14]:


[col for col in df_raw.columns if "Summer" in col or "Fall" in col or "Spring" in col]


# # Preparing the Spreadsheet-DF

# In[15]:


import pandas as pd
import re

# ---------------------------------------------------------
# Utility: detect columns like "Summer 2026"
# ---------------------------------------------------------
def find_real_year_column(df_raw, semester_name):
    """
    Finds a column like 'Summer 2026', 'Summer 2026 Course', etc.
    Returns: (column_name, '2026') or (None, None).
    """
    pattern = re.compile(rf"^{semester_name}\s+(\d{{4}})\b")
    for col in df_raw.columns:
        m = pattern.match(str(col).strip())
        if m:
            return col, m.group(1)
    return None, None


# ---------------------------------------------------------
# Parse course code like "OTST551 Hebrew I"
# ---------------------------------------------------------
def parse_course(raw):
    if not isinstance(raw, str) or not raw.strip():
        return "", "", ""

    parts = raw.strip().split(maxsplit=1)
    deptnum = parts[0]
    fallback_title = parts[1] if len(parts) > 1 else ""

    m = re.match(r"([A-Za-z]+)(\d+)", deptnum)
    if m:
        dept, num = m.group(1), m.group(2)
    else:
        dept, num = deptnum, ""

    return dept, num, fallback_title


# ---------------------------------------------------------
# Contract row detection
# ---------------------------------------------------------
def is_contract(v):
    """Contract rows must contain ONLY the word 'contract'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    return str(v).strip().lower() == "contract"


# ---------------------------------------------------------
# Course Section getter (supports suffix + fallback)
# ---------------------------------------------------------
def get_course_section(row, suffix: str) -> str:
    """
    Prefer a term-specific 'Course Section{suffix}' if present,
    otherwise fall back to plain 'Course Section'.
    Returns a cleaned string (empty if missing).
    """
    col_suffix = f"Course Section{suffix}"
    if col_suffix in row.index:
        v = row.get(col_suffix, "")
    else:
        v = row.get("Course Section", "")

    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""

    return str(v).strip()


def _safe_source_prefix(path):
    """Build a filename-safe prefix from the source workbook name."""
    stem = Path(path).stem
    stem = re.sub(r"[^\w\-. ]+", "", stem)
    stem = re.sub(r"\s+", "_", stem).strip("_")
    return stem or "ShariSheet"


def extract_contract_records_from_excel(excel_path):
    """Extract all contract course rows from one ShariSheet workbook."""
    df_raw = pd.read_excel(excel_path, sheet_name=0)

    # ---------------------------------------------------------
    # Detect terms automatically per workbook
    # ---------------------------------------------------------
    terms = []
    for sem in ["Summer", "Fall", "Spring"]:
        col, year = find_real_year_column(df_raw, sem)
        if col:
            terms.append((sem, col, year))

    print(f"\nSource workbook: {excel_path}")
    print("Detected term blocks:")
    for t in terms:
        print(t)

    # ---------------------------------------------------------
    # Build suffix map (index-based) per workbook
    # ---------------------------------------------------------
    sem_to_suffix = {
        sem: ("" if idx == 0 else f".{idx}")
        for idx, (sem, col, yr) in enumerate(terms)
    }

    print("Suffix map:")
    for sem, col, yr in terms:
        print(f"  {sem}: column={col}, year={yr}, suffix={sem_to_suffix[sem]}")

    source_name = Path(excel_path).name
    source_prefix = _safe_source_prefix(excel_path)
    workbook_records = []

    # -----------------------------------------------------
    # MAIN EXTRACTION LOOP for this workbook
    # -----------------------------------------------------
    for idx, row in df_raw.iterrows():

        name = row.get("Faculty Name", "")
        if not isinstance(name, str) or "," not in name:
            continue

        instructor = name.strip()

        # Instructor-level NON-term-specific metadata
        base_meta = {
            "SourceFile": source_name,
            "SourcePrefix": source_prefix,
            "Instructor": instructor,
            "ID": row.get("ID#", ""),
            "Email": row.get("email", ""),
            "Telephone": row.get("telephone", ""),
            "Account": row.get("account to be charged", ""),
            "Remote": row.get("remote employee", ""),
            "International": row.get("international", ""),
            "DeptContact": row.get("Dept contact name", ""),
            "DeptContactID": row.get("Dept contact ID#", ""),
            "ChairEmail": row.get("chair email", ""),
            "DeanEmail": row.get("dean email", ""),
            "DeanID": row.get("dean ID#", ""),
            "VPFinanceEmail": row.get("VP finance email", ""),
            "VPFinanceID": row.get("VP finance ID#", ""),
            "HRemail": row.get("HR email", ""),
        }

        # -----------------------------------------------------
        # Process each detected term
        # -----------------------------------------------------
        for (sem, sem_col, year) in terms:

            suffix = sem_to_suffix[sem]
            c_col  = sem_col
            load_c = f"load/contract{suffix}"

            # Contract-only rule
            if not is_contract(row.get(load_c, "")):
                continue

            course_raw = row.get(c_col, "")
            if not isinstance(course_raw, str) or not course_raw.strip():
                continue

            dept, num, fallback_title = parse_course(course_raw)
            course_title = row.get(f"Catalog Title{suffix}", fallback_title)
            course_section = get_course_section(row, suffix)

            # -----------------------------------------------------
            # BUILD TERM-SPECIFIC COMBINED REASON
            # -----------------------------------------------------
            prog = row.get(f"Program{suffix}", "")
            loc  = row.get(f"Location{suffix}", "")
            reason_raw = row.get("Reason for Contract", "")

            reason_parts = []
            for item in (prog, loc, reason_raw):
                if isinstance(item, str) and item.strip():
                    reason_parts.append(item.strip())

            combined_reason = "/".join(reason_parts)

            # -----------------------------------------------------
            # Create the record for this term
            # -----------------------------------------------------
            workbook_records.append({
                **base_meta,
                "Reason": combined_reason,
                "Semester": sem,
                "Year": year,
                "CourseRaw": course_raw,
                "Dept": dept,
                "CourseNum": num,
                "CourseSection": course_section,
                "CourseTitle": course_title,

                "Credits": row.get(f"Cr.{suffix}", None),
                "Rate": row.get(f"rate per credit{suffix}", None),
                "DeptBudget": row.get(f"dept budget{suffix}", ""),

                "BeginDate": row.get(f"Begin Date{suffix}", None),
                "EndDate":   row.get(f"End Date{suffix}", None),

                "prework_start":        row.get(f"pre-work period start{suffix}", None),
                "prework_end":          row.get(f"pre-work period end{suffix}", None),
                "prework_weeks":        row.get(f"pre-work # of weeks{suffix}", None),
                "prework_hours_week":   row.get(f"pre-work hours/week{suffix}", None),
                "prework_hours_period": row.get(f"pre-work hours/period{suffix}", None),

                "intensive_start": row.get(f"intensive period start{suffix}", None),
                "intensive_end":   row.get(f"intensive period end{suffix}", None),
                "intensive_weeks": row.get(f"intensive # of weeks{suffix}", None),
                "intensive_hours_week":   row.get(f"intensive hours/week{suffix}", None),
                "intensive_hours_period": row.get(f"intensive hours/period{suffix}", None),

                "postwork_start": row.get(f"post-work period start{suffix}", None),
                "postwork_end":   row.get(f"post-work period end{suffix}", None),
                "postwork_weeks": row.get(f"post-work # of weeks{suffix}", None),
                "postwork_hours_week":   row.get(f"post-work hours/week{suffix}", None),
                "postwork_hours_period": row.get(f"post-work hours/period{suffix}", None),
            })

    print(f"Extracted contract rows from {source_name}:", len(workbook_records))
    return workbook_records


# ---------------------------------------------------------
# Extract rows from ALL input ShariSheet workbooks
# ---------------------------------------------------------
records = []
for excel_path in EXCEL_PATHS:
    records.extend(extract_contract_records_from_excel(excel_path))


# ---------------------------------------------------------
# Final DF
# ---------------------------------------------------------
df_courses = pd.DataFrame(records)
print("\nTotal extracted contract rows:", len(df_courses))
if not df_courses.empty and "SourceFile" in df_courses.columns:
    print("Rows per source file:")
    print(df_courses["SourceFile"].value_counts())

df_courses.head()


# In[16]:


print(df_raw.filter(like="load/contract").head(20))


# In[17]:


df_courses.columns.tolist()


# In[18]:


df_courses.head()


# # Matching DF fields with PDF fields

# In[19]:


# --- Cell 6: Build PDF field data + PDF filler function ---  
# (UPDATED: ONE COURSE PER PDF + GENERATE ALL CONTRACTS)

import math
import re
from io import BytesIO
from pathlib import Path
import pandas as pd

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject,
    TextStringObject,
    BooleanObject,
    DictionaryObject,
)

from reportlab.pdfgen import canvas
from reportlab.lib.colors import blue


# --------------------------------------------------------
# Utilities
# --------------------------------------------------------

def safe_int_string(x):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    try:
        return str(int(float(str(x).strip())))
    except Exception:
        return ""

def safe_number(x):
    if x is None:
        return 0
    try:
        if pd.isna(x):
            return 0
    except Exception:
        pass
    try:
        return float(str(x).strip())
    except Exception:
        return 0

def normalize_date(raw):
    if raw is None or raw == "" or (isinstance(raw, float) and math.isnan(raw)):
        return ""
    if hasattr(raw, "strftime") and not pd.isna(raw):
        return raw.strftime("%m/%d/%Y")
    s = str(raw).strip()
    parsed = pd.to_datetime(s, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%m/%d/%Y")
    if "-" in s and len(s.split("-")) >= 3:
        try:
            y, m, d = s.split(" ", 1)[0].split("-")[:3]
            return f"{m.zfill(2)}/{d.zfill(2)}/{y}"
        except Exception:
            pass
    if "/" in s:
        return s
    return ""

def split_mmddyyyy(raw):
    if not raw:
        return "", "", ""
    parts = str(raw).strip().split("/")
    if len(parts) != 3:
        return "", "", ""
    mm, dd, yyyy = parts
    return mm.zfill(2), dd.zfill(2), yyyy[-2:]

def format_phone(raw):
    if raw is None:
        return ""
    digits = re.sub(r"\D", "", str(raw))
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    if len(digits) == 11 and digits.startswith("1"):
        d = digits[1:]
        return f"({d[:3]}) {d[3:6]}-{d[6:]}"
    return ""

def format_year_two_digits(raw):
    if raw is None:
        return ""
    s = str(raw).strip()
    return s[-2:] if s[-2:].isdigit() else ""

def format_id(raw):
    if raw is None:
        return ""
    s = str(raw).strip()
    if s.lower() == "nan":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if re.fullmatch(r"\d+\.\d+", s):
        try:
            return str(int(float(s)))
        except Exception:
            return s
    return s


# --------------------------------------------------------
# Course field mapping (UPDATED: ONE COURSE PER PDF)
# --------------------------------------------------------

def map_course_fields(course):
    """
    Convert exactly ONE course into PDF field names.
    Adds CourseSection to course_code when present: e.g., CHIS570-999

    Fills course1_* only, and blanks out course2/course3 fields
    (safe even if those fields don't exist in the PDF).
    """
    c = course or {}
    pdf_fields = {}

    dept = str(c.get("Dept", "") or "").strip()
    num  = str(c.get("CourseNum", "") or "").strip()
    sec  = str(c.get("CourseSection", "") or "").strip()

    if sec.lower() == "nan":
        sec = ""

    course_code = f"{dept}{num}" if (dept or num) else str(c.get("CourseRaw", "") or "").strip()
    if course_code and sec:
        course_code = f"{course_code}-{sec}"

    pdf_fields["course1_code"]       = course_code
    pdf_fields["course1_title"]      = c.get("CourseTitle", "")
    pdf_fields["course1_credits"]    = c.get("Credits")
    pdf_fields["course1_enrollment"] = c.get("Enrollment")

    pdf_fields["course2_code"] = ""
    pdf_fields["course2_title"] = ""
    pdf_fields["course2_credits"] = ""
    pdf_fields["course2_enrollment"] = ""
    pdf_fields["course3_code"] = ""
    pdf_fields["course3_title"] = ""
    pdf_fields["course3_credits"] = ""
    pdf_fields["course3_enrollment"] = ""

    return pdf_fields


def parse_account_number(raw):
    if not isinstance(raw, str):
        return "", "", "", "", ""
    if ":" in raw:
        raw = raw.split(":", 1)[1].strip()

    parts = raw.split("-")
    parts += [""] * (4 - len(parts))

    fund    = parts[0].strip()
    orgn    = parts[1].strip()
    acct    = parts[2].strip()
    program = parts[3].strip()
    activity = parts[4].strip() if len(parts) > 4 else ""

    return fund, orgn, acct, program, activity


# --------------------------------------------------------
# Formatting Tweaks (recursive field tree + widget styling)
# --------------------------------------------------------

def _walk_field_tree(obj):
    """Recursively yield all nodes under a field tree."""
    yield obj
    kids = obj.get("/Kids")
    if kids:
        for kid in kids:
            try:
                kid_obj = kid.get_object()
            except Exception:
                continue
            yield from _walk_field_tree(kid_obj)

def _is_widget(obj):
    return obj.get("/Subtype") == NameObject("/Widget") or obj.get("/Subtype") == "/Widget"

def _resolve_ft(node):
    """Resolve /FT from node or its parent (some widgets inherit it)."""
    ft = node.get("/FT")
    if ft is None and node.get("/Parent"):
        try:
            ft = node["/Parent"].get_object().get("/FT")
        except Exception:
            ft = None
    return ft

def apply_text_style_to_text_fields_only(
    acroform,
    font_name="/Helvetica",
    font_size=8,
    rgb=(0, 0, 0),
    remove_text_ap=True,
):
    """
    Apply fixed font+size+color ONLY to text fields (/FT /Tx) across the entire tree.
    Never touch /Btn fields (checkbox/radio).
    Optionally remove /AP only for text widgets (safe).
    """
    r, g, b = rgb
    da = f"{r} {g} {b} rg {font_name} {font_size} Tf"

    acroform[NameObject("/DA")] = TextStringObject(da)

    for f in acroform.get("/Fields", []):
        root = f.get_object()
        for node in _walk_field_tree(root):
            ft = _resolve_ft(node)
            is_text = (ft == NameObject("/Tx") or ft == "/Tx")

            if is_text:
                node[NameObject("/DA")] = TextStringObject(da)

            if remove_text_ap and is_text and _is_widget(node):
                if "/AP" in node:
                    del node["/AP"]


# --------------------------------------------------------
# Build the PDF data dictionary (UPDATED: ONE COURSE PER PDF)
# --------------------------------------------------------

def make_pdf_data_for_contract(instructor_name, semester, year, courses_chunk):
    """
    This ALWAYS builds data for exactly ONE course per PDF.
    If a list is passed, only the FIRST row is used.
    """
    try:
        last, first = split_name(instructor_name)
    except Exception:
        last, first = instructor_name, ""

    meta = courses_chunk[0] if courses_chunk else {}

    def safe_credit(x):
        try:
            x = float(x)
            return 0 if math.isnan(x) else x
        except Exception:
            return 0

    total_credits = safe_credit(meta.get("Credits"))
    weekly_hours = total_credits * 3

    raw_rate = meta.get("Rate")
    if raw_rate is None or str(raw_rate).strip() in ("", "nan"):
        rate_per_credit = DEFAULT_RATE_PER_CREDIT
    else:
        try:
            rate_per_credit = float(raw_rate)
        except Exception:
            rate_per_credit = raw_rate

    if isinstance(rate_per_credit, (int, float)) and total_credits:
        amount_value = rate_per_credit * total_credits
    else:
        if isinstance(meta.get("Salary"), (int, float)):
            amount_value = meta.get("Salary")
        elif isinstance(meta.get("PhD"), (int, float)):
            amount_value = meta.get("PhD")
        else:
            amount_value = ""

    amount_str = (
        str(int(amount_value))
        if isinstance(amount_value, (int, float)) and not math.isnan(amount_value)
        else ""
    )

    begin_raw = normalize_date(meta.get("BeginDate"))
    end_raw   = normalize_date(meta.get("EndDate"))

    begin_m, begin_d, begin_y = split_mmddyyyy(begin_raw)
    end_m,   end_d,   end_y   = split_mmddyyyy(end_raw)

    acct_raw = meta.get("Account", "")
    fund, orgn, acct_num, program, activity = parse_account_number(acct_raw)

    data = {
        "last_name": last,
        "first_name": first,
        "id": format_id(meta.get("ID")),
        "email": meta.get("Email", ""),
        "telephone": format_phone(meta.get("Telephone")),

        "remote_yes": "Off",
        "payment_yes": "Off",
        "payment_no": "Off",

        "dept_contact_name": meta.get("DeptContact", ""),
        "dept_contact_id": format_id(meta.get("DeptContactID")),

        "reason": meta.get("Reason", ""),

        "sem_spring": "Off",
        "sem_summer": "Off",
        "sem_fall":   "Off",

        "year": format_year_two_digits(year),

        "begin_month": begin_m,
        "begin_day": begin_d,
        "begin_year": begin_y,

        "end_month": end_m,
        "end_day": end_d,
        "end_year": end_y,

        "Begin Date month": begin_m,
        "Begin Date date": begin_d,
        "Begin Date year": begin_y,
        "End Date month": end_m,
        "End Date date": end_d,
        "End Date year": end_y,

        "credits_total": str(int(total_credits)) if total_credits else "",
        "week_srv_hrs": str(int(weekly_hours)) if weekly_hours else "",
        "rate_per_credit": str(rate_per_credit),
        "hourly_rate": "",
        "amount_total": amount_str,

        "fund": fund,
        "orgn": orgn,
        "acct": acct_num,
        "program": program,
        "activity": activity,

        "Signature2ID": format_id(meta.get("DeptContactID")),
        "Signature3ID": format_id(meta.get("DeanID")),
        "Signature4ID": format_id(meta.get("VPFinanceID")),
    }

    remote_raw = str(meta.get("Remote", "")).strip().lower()
    data["remote_yes"] = "/Yes" if remote_raw in ("yes", "y", "true", "1") else "/Off"

    sem_raw = str(meta.get("Semester", "")).strip().lower()
    data["spring_yes"] = NameObject("/Yes") if sem_raw == "spring" else NameObject("/Off")
    data["summer_yes"] = NameObject("/Yes") if sem_raw == "summer" else NameObject("/Off")
    data["fall_yes"]   = NameObject("/Yes") if sem_raw == "fall"   else NameObject("/Off")

    international_raw = str(meta.get("International", "")).strip().lower()
    data["international"] = "/Yes" if international_raw in ("yes", "y", "true", "1") else "/Off"

    dept_budget_raw = str(meta.get("DeptBudget", "")).strip().lower()
    data["dept_budget_yes"] = "/Yes" if dept_budget_raw in ("yes", "y", "true", "1") else "/Off"
    data["dept_budget_no"]  = "/Yes" if dept_budget_raw in ("no", "n", "false", "0") else "/Off"

    data.update(map_course_fields(meta))

    data["prework_start"]        = normalize_date(meta.get("prework_start"))
    data["prework_end"]          = normalize_date(meta.get("prework_end"))
    data["prework_of_weeks"]     = safe_int_string(meta.get("prework_weeks"))
    data["prework_hours_weeks"]  = safe_int_string(meta.get("prework_hours_week"))
    data["prework_hours_period"] = safe_int_string(meta.get("prework_hours_period"))

    data["intensive_start"]        = normalize_date(meta.get("intensive_start"))
    data["intensive_end"]          = normalize_date(meta.get("intensive_end"))
    data["intensive_of_weeks"]     = safe_int_string(meta.get("intensive_weeks"))
    data["intensive_hours_weeks"]  = safe_int_string(meta.get("intensive_hours_week"))
    data["intensive_hours_period"] = safe_int_string(meta.get("intensive_hours_period"))

    data["postwork_start"]        = normalize_date(meta.get("postwork_start"))
    data["postwork_end"]          = normalize_date(meta.get("postwork_end"))
    data["postwork_of_weeks"]     = safe_int_string(meta.get("postwork_weeks"))
    data["postwork_hours_week"]   = safe_int_string(meta.get("postwork_hours_week"))
    data["postwork_hours_period"] = safe_int_string(meta.get("postwork_hours_period"))

    pre_h  = safe_number(meta.get("prework_hours_period"))
    int_h  = safe_number(meta.get("intensive_hours_period"))
    post_h = safe_number(meta.get("postwork_hours_period"))
    data["total_contract_hours"] = safe_int_string(pre_h + int_h + post_h)

    if meta:
        print(
            f"[DEBUG] {instructor_name}, {semester} {year} – "
            f"BeginDate={meta.get('BeginDate')!r} -> {begin_m}/{begin_d}/{begin_y}, "
            f"EndDate={meta.get('EndDate')!r} -> {end_m}/{end_d}/{end_y}, "
            f"course1_code={data.get('course1_code')!r}"
        )

    return data


# --------------------------------------------------------
# PDF Writer
# --------------------------------------------------------

def sanitize_pdf_data(data: dict):
    clean = {}
    for k, v in data.items():
        if isinstance(v, NameObject):
            clean[k] = v
            continue
        if v is None:
            clean[k] = ""
            continue
        try:
            if isinstance(v, float) and math.isnan(v):
                clean[k] = ""
            else:
                clean[k] = str(v)
        except Exception:
            clean[k] = ""
    return clean


def _fit_font_size_for_width(text, max_width, font_name="Helvetica", start_size=10, min_size=5):
    from reportlab.pdfbase.pdfmetrics import stringWidth

    size = start_size
    while size > min_size and stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    return max(size, min_size)


def fill_pdf_contract(template_path, output_path, data):
    clean_data = sanitize_pdf_data(data)

    reader = PdfReader(str(template_path), strict=False)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    # Build field lookup from widget annotations
    page_field_map = {}

    for page_num, page in enumerate(reader.pages):
        annots = page.get("/Annots", [])
        page_field_map[page_num] = []

        for annot_ref in annots:
            try:
                annot = annot_ref.get_object()
            except Exception:
                continue

            if annot.get("/Subtype") != "/Widget":
                continue

            field_name = annot.get("/T")
            parent = None

            if field_name is None and annot.get("/Parent"):
                try:
                    parent = annot["/Parent"].get_object()
                    field_name = parent.get("/T")
                except Exception:
                    parent = None

            if not field_name:
                continue

            field_name = str(field_name)
            rect = annot.get("/Rect")
            if not rect or len(rect) != 4:
                continue

            try:
                x1, y1, x2, y2 = [float(v) for v in rect]
            except Exception:
                continue

            ft = annot.get("/FT")
            if ft is None and parent is not None:
                ft = parent.get("/FT")

            page_field_map[page_num].append({
                "name": field_name,
                "rect": (x1, y1, x2, y2),
                "ft": str(ft) if ft is not None else "",
            })

    # Create overlay per page
    for page_num, page in enumerate(reader.pages):
        packet = BytesIO()
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        c = canvas.Canvas(packet, pagesize=(page_width, page_height))
        c.setFillColor(blue)

        for field in page_field_map.get(page_num, []):
            name = field["name"]
            if name not in clean_data:
                continue

            value = clean_data[name]
            if value in (None, "", "/Off"):
                continue

            x1, y1, x2, y2 = field["rect"]
            width = x2 - x1
            height = y2 - y1
            ft = field["ft"]

            # Buttons / checkboxes
            if ft == "/Btn":
                yes_values = {"/Yes", "Yes", True, "True", "true", "1", 1}
                if value in yes_values:
                    mark_size = max(8, min(height * 0.8, 12))
                    c.setFont("Helvetica", mark_size)
                    c.drawString(x1 + 1, y1 + max(1, (height - mark_size) / 2), "X")
                continue

            # Text / choice fields
            text = str(value)
            max_text_width = max(10, width - 4)
            font_size = _fit_font_size_for_width(
                text,
                max_width=max_text_width,
                font_name="Helvetica",
                start_size=max(6, min(height * 0.7, 10)),
                min_size=5,
            )

            c.setFont("Helvetica", font_size)
            text_y = y1 + max(1, (height - font_size) / 2)
            c.drawString(x1 + 2, text_y, text[:500])

        c.save()
        packet.seek(0)

        overlay_reader = PdfReader(packet)
        overlay_page = overlay_reader.pages[0]
        writer.pages[page_num].merge_page(overlay_page)

    # Remove widget annotations so Acrobat does not show blank form fields on top
    for page in writer.pages:
        annots = page.get("/Annots")
        if not annots:
            continue

        kept = []
        for annot_ref in annots:
            try:
                annot = annot_ref.get_object()
                if annot.get("/Subtype") != "/Widget":
                    kept.append(annot_ref)
            except Exception:
                kept.append(annot_ref)

        if kept:
            page[NameObject("/Annots")] = kept
        elif "/Annots" in page:
            del page["/Annots"]

    # Remove AcroForm if present
    if "/AcroForm" in writer._root_object:
        del writer._root_object["/AcroForm"]

    with open(output_path, "wb") as f:
        writer.write(f)


# --------------------------------------------------------
# NEW: Generate ALL contracts from df_course (ONE PDF PER ROW)
# --------------------------------------------------------

def _build_course_code_from_row(r: dict) -> str:
    dept = str(r.get("Dept", "") or "").strip()
    num  = str(r.get("CourseNum", "") or "").strip()
    sec  = str(r.get("CourseSection", "") or "").strip()
    if sec.lower() == "nan":
        sec = ""
    code = f"{dept}{num}" if (dept or num) else str(r.get("CourseRaw", "") or "").strip()
    if code and sec:
        code = f"{code}-{sec}"
    return code or "COURSE"

def _safe_filename(s: str) -> str:
    s = str(s or "").strip()
    s = re.sub(r"[^\w\-\. ]+", "", s)
    s = re.sub(r"\s+", "_", s).strip("_")
    return s or "UNKNOWN"

def generate_all_contract_pdfs_one_per_course(
    df_course: pd.DataFrame,
    template_path,
    output_dir,
    instructor_col="Instructor",
    semester_col="Semester",
    year_value=None,
    year_col="Year",
    filename_add_row_index=True,
):
    """
    This writes ONE PDF per row in df_course (per course).
    """

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    written = 0

    for idx, row in df_course.iterrows():
        r = row.to_dict()

        instructor = r.get(instructor_col, "") or ""
        semester   = r.get(semester_col, "") or ""
        year       = year_value if year_value is not None else (r.get(year_col, "") or "")

        data = make_pdf_data_for_contract(instructor, semester, year, [r])

        try:
            last, first = split_name(instructor)
        except Exception:
            last, first = instructor, ""

        course_code = _build_course_code_from_row(r)

        base_name = (
            f"{_safe_filename(last)}_{_safe_filename(first)}__"
            f"{_safe_filename(str(semester))}_{_safe_filename(str(year))}__"
            f"{_safe_filename(course_code)}"
        )
        if filename_add_row_index:
            base_name += f"__row{idx}"

        out_path = output_dir / f"{base_name}.pdf"

        fill_pdf_contract(template_path, out_path, data)
        written += 1

    print(f"[DONE] Wrote {written} contract PDFs to: {output_dir}")


# In[20]:


# === Helper for Cell A/B ===

import re

def _clean_str(x):
    if x is None:
        return ""
    s = str(x).strip()
    if s.lower() == "nan":
        return ""
    return s

def _clean_section(sec):
    s = _clean_str(sec)
    # normalize common section formats; keep letters/numbers
    s = s.replace("/", "-")
    s = re.sub(r"\s+", "", s)
    return s

def course_token(course_dict):
    """
    Build token like 'OTST551-01' from a course record dict.
    Expects keys: Dept, CourseNum, CourseSection
    """
    dept = _clean_str(course_dict.get("Dept", ""))
    num  = _clean_str(course_dict.get("CourseNum", ""))
    sec  = _clean_section(course_dict.get("CourseSection", ""))

    base = f"{dept}{num}".strip()
    if not base:
        # fallback: try parsing CourseRaw if needed
        raw = _clean_str(course_dict.get("CourseRaw", ""))
        base = raw.split()[0] if raw else "UNKNOWN"

    return f"{base}-{sec}" if sec else base

def course_tokens_for_chunk(chunk):
    """
    chunk is a list[dict] (from to_dict("records"))
    Returns stable, de-duplicated token string joined by '+'
    """
    toks = [course_token(c) for c in chunk]
    # de-dupe preserving order
    seen = set()
    toks2 = []
    for t in toks:
        if t and t not in seen:
            seen.add(t)
            toks2.append(t)
    return "+".join(toks2) if toks2 else "NOCOURSE"


# In[21]:


# === Cell A: Generate STANDARD contracts from ALL input ShariSheets ===
# UPDATED: ONE COURSE PER PDF (no chunking); generates ALL contracts

semester_code = {"Spring": "01", "Summer": "02", "Fall": "03"}
generated_standard_files = []

grouped = df_courses.groupby(["Instructor", "Semester", "Year"])

for (instructor, semester, year), subdf in grouped:

    # ---- Split intensive vs standard (STANDARD = no intensive_start) ----
    standard_df = subdf[
        subdf["intensive_start"].isna() |
        (subdf["intensive_start"].astype(str).str.strip() == "")
    ]

    # ONE PDF PER ROW/COURSE
    if standard_df.empty:
        continue

    last, first = split_name(instructor)
    safe_name = f"{first}_{last}".replace(" ", "_")
    sem_code = semester_code.get(semester, "00")

    # ----------------------------------------------------
    # PROCESS STANDARD CONTRACTS (ONE COURSE PER PDF)
    # ----------------------------------------------------
    for row_i, row in standard_df.reset_index(drop=False).iterrows():
        c = row.to_dict()

        # Optional: per-course debug (uncomment if needed)
        # print("CourseRaw:", c.get("CourseRaw"))
        # print("Dept:", c.get("Dept"), "CourseNum:", c.get("CourseNum"), "CourseSection:", c.get("CourseSection"))
        # print("----")

        # IMPORTANT: pass ONE course only
        data = make_pdf_data_for_contract(instructor, semester, year, [c])

        # Filename token for this single course
        token_str = course_token(c) if "course_token" in globals() else course_tokens_for_chunk([c])

        source_prefix = _clean_str(c.get("SourcePrefix", ""))
        base = f"{source_prefix}__{year}-{sem_code}_{semester}_standard_{safe_name}_{token_str}" if source_prefix else f"{year}-{sem_code}_{semester}_standard_{safe_name}_{token_str}"

        # Ensure uniqueness (same course token could appear multiple times)
        # Prefer original dataframe index if present; fallback to row number
        unique_suffix = c.get("index", None)
        if unique_suffix is None or str(unique_suffix).strip() == "":
            unique_suffix = row_i
        base = f"{base}_row{unique_suffix}"

        outfile = OUTPUT_DIR / f"{base}.pdf"
        fill_pdf_contract(PDF_standard_TEMPLATE_PATH, outfile, data)
        generated_standard_files.append(outfile)

len(generated_standard_files), generated_standard_files[:10]


# In[22]:


# === Cell B: Generate INTENSIVE contracts from ALL input ShariSheets ===
# UPDATED: ONE COURSE PER PDF (no chunking); generates ALL contracts

semester_code = {"Spring": "01", "Summer": "02", "Fall": "03"}
generated_intensive_files = []

grouped = df_courses.groupby(["Instructor", "Semester", "Year"])

for (instructor, semester, year_full), subdf in grouped:

    # ---- Split intensive vs standard (INTENSIVE = has intensive_start) ----
    intensive_df = subdf[
        subdf["intensive_start"].notna() &
        (subdf["intensive_start"].astype(str).str.strip() != "")
    ]

    # ONE PDF PER ROW/COURSE
    if intensive_df.empty:
        continue

    last, first = split_name(instructor)
    safe_name = f"{first}_{last}".replace(" ", "_")
    sem_code = semester_code.get(semester, "00")

    # ----------------------------------------------------
    # PROCESS INTENSIVE CONTRACTS (ONE COURSE PER PDF)
    # ----------------------------------------------------
    for row_i, row in intensive_df.reset_index(drop=False).iterrows():
        c = row.to_dict()

        # IMPORTANT: pass ONE course only
        data = make_pdf_data_for_contract(instructor, semester, year_full, [c])

        # Filename token for this single course
        token_str = course_token(c) if "course_token" in globals() else course_tokens_for_chunk([c])

        source_prefix = _clean_str(c.get("SourcePrefix", ""))
        base = f"{source_prefix}__{year_full}-{sem_code}_{semester}_intensive_{safe_name}_{token_str}" if source_prefix else f"{year_full}-{sem_code}_{semester}_intensive_{safe_name}_{token_str}"

        # Ensure uniqueness
        unique_suffix = c.get("index", None)
        if unique_suffix is None or str(unique_suffix).strip() == "":
            unique_suffix = row_i
        base = f"{base}_row{unique_suffix}"

        outfile = OUTPUT_DIR / f"{base}.pdf"

        # Optional detailed debug (now single-course)
        if instructor == "Glanz, Oliver" and semester == "Summer":
            print("\n🔍 DEBUG for Summer Intensive — Glanz, Oliver")
            print("course:", c.get("CourseRaw"), "| CourseToken =", token_str, "| postwork_end =", c.get("postwork_end"))
            print("PDF data postwork_end:", data.get("postwork_end"))
            print("PDF data keys with 'postwork':", [k for k in data.keys() if "postwork" in k.lower()])
            print("Account fields:", {
                "fund": data.get("fund"),
                "orgn": data.get("orgn"),
                "acct": data.get("acct"),
                "program": data.get("program"),
                "activity": data.get("activity"),
            })

        fill_pdf_contract(PDF_intensive_TEMPLATE_PATH, outfile, data)
        generated_intensive_files.append(outfile)

len(generated_intensive_files), generated_intensive_files[:10]


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:



