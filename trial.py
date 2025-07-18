import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
from fpdf import FPDF
import os
import re
import random
import io
from PyPDF2 import PdfReader, PdfWriter
from collections import deque, defaultdict

# Set page configuration
st.set_page_config(
    page_title="Exam Timetable Generator",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for consistent dark and light mode styling
st.markdown("""
<style>
    /* Base styles */
    .main-header {
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }

    .main-header h1 {
        color: white;
        text-align: center;
        margin: 0;
        font-size: 2.5rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }

    .main-header p {
        color: #FFF;
        text-align: center;
        margin: 0.5rem 0 0 0;
        font-size: 1.2rem;
        opacity: 0.9;
    }

    .stats-section {
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }

    /* Updated metric card with icons */
    .metric-card {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 1.5rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 0.5rem;
        transition: transform 0.2s;
    }

    .metric-card:hover {
        transform: scale(1.05);
    }

    .metric-card h3 {
        margin: 0;
        font-size: 1.8rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .metric-card p {
        margin: 0.3rem 0 0 0;
        font-size: 1rem;
        opacity: 0.9;
    }

    /* Add gap between difficulty selector and holiday collapsible menu */
    .stCheckbox + .stExpander {
        margin-top: 2rem;
    }

    /* Button hover animations for regular buttons */
    .stButton>button {
        transition: all 0.3s ease;
        border-radius: 5px;
        border: 1px solid transparent;
    }

    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        border: 1px solid #951C1C;
        background-color: #C73E1D;
        color: white;
    }

    /* Download button hover effects (aligned with regular buttons) */
    .stDownloadButton>button {
        transition: all 0.3s ease;
        border-radius: 5px;
        border: 1px solid transparent;
    }

    .stDownloadButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        border: 1px solid #951C1C;
        background-color: #C73E1D;
        color: white;
    }

    /* Light mode styles */
    @media (prefers-color-scheme: light) {
        .main-header {
            background: linear-gradient(90deg, #951C1C, #C73E1D);
        }

        .upload-section {
            background: #f8f9fa;
            padding: 2rem;
            border-radius: 10px;
            border: 2px dashed #951C1C;
            margin: 1rem 0;
        }

        .results-section {
            background: #ffffff;
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin: 1rem 0;
        }

        .stats-section {
            background: #f8f9fa;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }

        .status-success {
            background: #d4edda;
            color: #155724;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #28a745;
        }

        .status-error {
            background: #f8d7da;
            color: #721c24;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #dc3545;
        }

        .status-info {
            background: #d1ecf1;
            color: #0c5460;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #17a2b8;
        }

        .feature-card {
            background: white;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin: 1rem 0;
            border-left: 4px solid #951C1C;
        }

        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }

        .footer {
            text-align: center;
            color: #666;
            padding: 2rem;
        }
    }

    /* Dark mode styles */
    @media (prefers-color-scheme: dark) {
        .main-header {
            background: linear-gradient(90deg, #701515, #A23217);
        }

        .upload-section {
            background: #333;
            padding: 2rem;
            border-radius: 10px;
            border: 2px dashed #A23217;
            margin: 1rem 0;
        }

        .results-section {
            background: #222;
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            margin: 1rem 0;
        }

        .stats-section {
            background: #333;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }

        .status-success {
            background: #2d4b2d;
            color: #e6f4ea;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #4caf50;
        }

        .status-error {
            background: #4b2d2d;
            color: #f8d7da;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #f44336;
        }

        .status-info {
            background: #2d4b4b;
            color: #d1ecf1;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #00bcd4;
        }

        .feature-card {
            background: #333;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.3);
            margin: 1rem 0;
            border-left: 4px solid #A23217;
        }

        .metric-card {
            background: linear-gradient(135deg, #4a5db0 0%, #5a3e8a 100%);
        }

        .footer {
            text-align: center;
            color: #ccc;
            padding: 2rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# Define the mapping of main branch abbreviations to full forms
BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS"
}

# Define logo path (adjust as needed for your environment)
LOGO_PATH = "logo.png"  # Ensure this path is valid in your environment


# Cache for text wrapping results
wrap_text_cache = {}


def safe_split(time_str):
    """Safely splits a time string, handling None, NaN, or non-string inputs."""
    # Handle None, NaN, or empty values
    if pd.isna(time_str) or time_str is None:
        return ["10:00 AM", "1:00 PM"]  # Default fallback
    
    # Convert to string and strip whitespace
    time_str = str(time_str).strip()
    
    # Check if empty string after conversion
    if time_str in ["", "nan"]:
        return ["10:00 AM", "1:00 PM"]  # Default fallback
        
    if " - " in time_str:
        return time_str.split(" - ")
    else:
        # If no delimiter found, return default
        return ["10:00 AM", "1:00 PM"]

# PDF generation utilities
def wrap_text(pdf, text, col_width):
    cache_key = (text, col_width)
    if cache_key in wrap_text_cache:
        return wrap_text_cache[cache_key]
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        test_line = word if not current_line else current_line + " " + word
        if pdf.get_string_width(test_line) <= col_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    wrap_text_cache[cache_key] = lines
    return lines

def print_row_custom(pdf, row_data, col_widths, line_height=5, header=False):
    cell_padding = 2
    header_bg_color = (149, 33, 28)
    header_text_color = (255, 255, 255)
    alt_row_color = (240, 240, 240)

    row_number = getattr(pdf, '_row_counter', 0)
    if header:
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(*header_text_color)
        pdf.set_fill_color(*header_bg_color)
    else:
        pdf.set_font("Arial", size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(*alt_row_color if row_number % 2 == 1 else (255, 255, 255))

    wrapped_cells = []
    max_lines = 0
    for i, cell_text in enumerate(row_data):
        text = str(cell_text) if cell_text is not None else ""
        avail_w = col_widths[i] - 2 * cell_padding
        lines = wrap_text(pdf, text, avail_w)
        wrapped_cells.append(lines)
        max_lines = max(max_lines, len(lines))

    row_h = line_height * max_lines
    x0, y0 = pdf.get_x(), pdf.get_y()
    if header or row_number % 2 == 1:
        pdf.rect(x0, y0, sum(col_widths), row_h, 'F')

    for i, lines in enumerate(wrapped_cells):
        cx = pdf.get_x()
        pad_v = (row_h - len(lines) * line_height) / 2 if len(lines) < max_lines else 0
        for j, ln in enumerate(lines):
            pdf.set_xy(cx + cell_padding, y0 + j * line_height + pad_v)
            pdf.cell(col_widths[i] - 2 * cell_padding, line_height, ln, border=0, align='C')
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, branches=None, time_slot=None):
    if df.empty:
        return
    setattr(pdf, '_row_counter', 0)
    
    # Add footer first
    footer_height = 25
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    pdf.set_font("Arial", size=13)
    pdf.set_xy(10, pdf.h - footer_height + 7)
    pdf.cell(0, 5, "Signature", 0, 1, 'L')
    
    # Add page numbers in bottom right
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))  # Estimate width
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')
    
    # Add header
    header_height = 85  # Increased height to accommodate time slot
    pdf.set_y(0)
    current_date = datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST")
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
    logo_width = 45
    logo_x = (pdf.w - logo_width) / 2
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.rect(10, 30, pdf.w - 20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w - 20, 14,
             "MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING / SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING",
             0, 1, 'C')
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w - 20, 8, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')
    
    # Add time slot if provided
    if time_slot:
        pdf.set_font("Arial", 'B', 14)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, f"Exam Time: {time_slot}", 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    else:
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(71)
    
    # Calculate available space
    available_height = pdf.h - pdf.t_margin - footer_height - header_height
    pdf.set_font("Arial", size=12)
    
    # Print header row
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    
    # Print data rows, checking space
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row):
            continue
            
        # Estimate row height
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 2 * 2
            lines = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines
        
        # Check if row fits
        if pdf.get_y() + row_h > pdf.h - footer_height:
            # Add footer to current page
            add_footer_with_page_number(pdf, footer_height)
            
            # Start new page
            pdf.add_page()
            # Add footer to new page
            add_footer_with_page_number(pdf, footer_height)
            
            # Add header to new page
            add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, branches, time_slot)
            
            # Reprint header row
            pdf.set_font("Arial", size=12)
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)


def add_footer_with_page_number(pdf, footer_height):
    """Add footer with signature and page number"""
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    pdf.set_font("Arial", size=13)
    pdf.set_xy(10, pdf.h - footer_height + 7)
    pdf.cell(0, 5, "Signature", 0, 1, 'L')
    
    # Add page numbers in bottom right
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))  # Estimate width
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')

def add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, branches, time_slot=None):
    """Add header to a new page"""
    pdf.set_y(0)
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.rect(10, 30, pdf.w - 20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w - 20, 14,
             "MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING / SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING",
             0, 1, 'C')
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w - 20, 8, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')
    
    # Add time slot if provided
    if time_slot:
        pdf.set_font("Arial", 'B', 14)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, f"Exam Time: {time_slot}", 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    else:
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 59)
        pdf.cell(pdf.w - 20, 6, "(Check the subject exam time)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(71)

def calculate_end_time(start_time, duration_hours):
    """Calculate the end time given a start time and duration in hours."""
    start = datetime.strptime(start_time, "%I:%M %p")
    duration = timedelta(hours=duration_hours)
    end = start + duration
    return end.strftime("%I:%M %p").replace("AM", "am").replace("PM", "pm")

def read_timetable(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df = df.rename(columns={
            "Program": "Program",
            "Stream": "Stream",
            "Current Session": "Semester",
            "Module Description": "SubjectName",
            "Module Abbreviation": "ModuleCode",
            "Campus Name": "Campus",
            "Difficulty Score": "Difficulty",
            "Exam Duration": "Exam Duration",
            "Student count": "StudentCount"  # Added this line
        })
        
        def convert_sem(sem):
            if pd.isna(sem):
                return 0
            m = {
                "Sem I": 1, "Sem II": 2, "Sem III": 3, "Sem IV": 4,
                "Sem V": 5, "Sem VI": 6, "Sem VII": 7, "Sem VIII": 8,
                "Sem IX": 9, "Sem X": 10, "Sem XI": 11
            }
            return m.get(str(sem).strip(), 0)
        
        df["Semester"] = df["Semester"].apply(convert_sem).astype(int)
        df["Branch"] = df["Program"].astype(str).str.strip() + "-" + df["Stream"].astype(str).str.strip()
        df["Subject"] = df["SubjectName"].astype(str) + " - (" + df["ModuleCode"].astype(str) + ")"
        
        comp_mask = (df["Category"] == "COMP") & df["Difficulty"].notna()
        df["Difficulty"] = pd.to_numeric(df["Difficulty"], errors='coerce')
        
        df["Exam Date"] = ""
        df["Time Slot"] = ""
        df["Exam Duration"] = df["Exam Duration"].fillna(3).astype(float)
        df["StudentCount"] = df["StudentCount"].fillna(0).astype(int)
        
        df_non = df[df["Category"] != "INTD"].copy()
        df_ele = df[df["Category"] == "INTD"].copy()
        
        def split_br(b):
            p = str(b).split("-", 1)
            return pd.Series([p[0].strip(), p[1].strip() if len(p) > 1 else ""])
        
        for d in (df_non, df_ele):
            if not d.empty:
                d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)
        
        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty", "Exam Duration", "StudentCount"]
        
        return df_non[cols], df_ele[cols], df
        
    except Exception as e:
        st.error(f"Error reading the Excel file: {str(e)}")
        return None, None, None

def schedule_semester_non_electives(df_sem, holidays, base_date, schedule_by_difficulty=False):
    df_sem['SubjectCode'] = df_sem['Subject'].str.extract(r'\((.*?)\)', expand=False)

    if isinstance(base_date, date) and not isinstance(base_date, datetime):
        base_date = datetime.combine(base_date, datetime.min.time())

    holidays_dates = {h.date() for h in holidays}
    all_branches = df_sem['Branch'].unique()
    exam_days = {branch: set() for branch in all_branches}

    def find_next_valid_day(start_day, for_branches):
        day = start_day
        while True:
            day_date = day.date() if isinstance(day, datetime) else day
            if day.weekday() == 6 or day_date in holidays_dates:
                day += timedelta(days=1)
                continue
            if all(day_date not in exam_days[branch] for branch in for_branches):
                return day
            day += timedelta(days=1)

    def count_working_days(start_date, end_date, holidays_dates):
        start = start_date.date() if isinstance(start_date, datetime) else start_date
        end = end_date.date() if isinstance(end_date, datetime) else end_date
        working_days = 0
        current_date = start
        while current_date <= end:
            if current_date.weekday() != 6 and current_date not in holidays_dates:
                working_days += 1
            current_date += timedelta(days=1)
        return working_days

    if schedule_by_difficulty:
        subject_info = df_sem.groupby('SubjectCode').agg({
            'Difficulty': 'first',
            'Branch': lambda x: set(x),
            'Subject': 'first'
        }).reset_index()
        subject_to_difficulty = dict(zip(subject_info['SubjectCode'], subject_info['Difficulty']))
        subject_to_branches = dict(zip(subject_info['SubjectCode'], subject_info['Branch']))
        subject_code_to_subject = dict(zip(subject_info['SubjectCode'], subject_info['Subject']))

        common_events = [sc for sc in subject_to_branches if len(subject_to_branches[sc]) > 1]
        individual_events = [(sc, list(subject_to_branches[sc])[0]) for sc in subject_to_branches if len(subject_to_branches[sc]) == 1]

        common_easy = [sc for sc in common_events if subject_to_difficulty.get(sc) == 0]
        common_difficult = [sc for sc in common_events if subject_to_difficulty.get(sc) == 1]
        common_neutral = [sc for sc in common_events if pd.isna(subject_to_difficulty.get(sc))]
        
        sorted_common = []
        while common_easy or common_difficult:
            if common_easy:
                sorted_common.append(common_easy.pop(0))
            if common_difficult:
                sorted_common.append(common_difficult.pop(0))
        sorted_common += common_neutral

        individual_per_branch = defaultdict(list)
        for subj_code, B in individual_events:
            individual_per_branch[B].append((subj_code, subject_to_difficulty.get(subj_code)))

        for B in individual_per_branch:
            subjects = individual_per_branch[B]
            easy = [s[0] for s in subjects if s[1] == 0]
            difficult = [s[0] for s in subjects if s[1] == 1]
            neutral = [s[0] for s in subjects if pd.isna(s[1])]
            sorted_individual_B = []
            last_difficulty = None
            while easy or difficult or neutral:
                if last_difficulty is None or last_difficulty == 1:
                    if easy:
                        s = easy.pop(0)
                        sorted_individual_B.append(s)
                        last_difficulty = 0
                    elif neutral:
                        s = neutral.pop(0)
                        sorted_individual_B.append(s)
                        last_difficulty = None
                    elif difficult:
                        s = difficult.pop(0)
                        sorted_individual_B.append(s)
                        last_difficulty = 1
                else:
                    if difficult:
                        s = difficult.pop(0)
                        sorted_individual_B.append(s)
                        last_difficulty = 1
                    elif neutral:
                        s = neutral.pop(0)
                        sorted_individual_B.append(s)
                        last_difficulty = None
                    elif easy:
                        s = easy.pop(0)
                        sorted_individual_B.append(s)
                        last_difficulty = 0
            individual_per_branch[B] = sorted_individual_B

        assignments = []
        for E_C in sorted_common:
            found = False
            for i in range(len(assignments)):
                day_assignments = assignments[i]
                conflicts = False
                for _, _, branches in day_assignments:
                    if any(branch in branches for branch in subject_to_branches[E_C]):
                        conflicts = True
                        break
                if not conflicts:
                    day_assignments.append((E_C, day_assignments[0][1], subject_to_branches[E_C]))
                    found = True
                    break
            if not found:
                day = find_next_valid_day(base_date if not assignments else assignments[-1][0][1] + timedelta(days=1), subject_to_branches[E_C])
                assignments.append([(E_C, day, subject_to_branches[E_C])])

        for B in all_branches:
            if B not in individual_per_branch:
                continue
            for E_I in individual_per_branch[B]:
                found = False
                for i in range(len(assignments)):
                    day_assignments = assignments[i]
                    conflicts = False
                    for _, _, branches in day_assignments:
                        if B in branches:
                            conflicts = True
                            break
                    if not conflicts:
                        day_assignments.append((E_I, day_assignments[0][1], {B}))
                        found = True
                        break
                if not found:
                    day = find_next_valid_day(base_date if not assignments else assignments[-1][0][1] + timedelta(days=1), {B})
                    assignments.append([(E_I, day, {B})])

        for day_assignments in assignments:
            for E, exam_day, branches in day_assignments:
                mask = (df_sem['SubjectCode'] == E) & (df_sem['Branch'].isin(branches))
                df_sem.loc[mask, 'Exam Date'] = exam_day.strftime("%d-%m-%Y")
                exam_day_date = exam_day.date()
                for branch in branches:
                    exam_days[branch].add(exam_day_date)

        if assignments:
            all_exam_dates = [da[0][1] for da in assignments if da]
            if all_exam_dates:
                first_exam = min(all_exam_dates)
                last_exam_date = max(all_exam_dates)
                total_span = count_working_days(first_exam, last_exam_date, holidays_dates)
                if total_span > 16:
                    print(f"Warning: The timetable spans {total_span} exam days, exceeding the limit of 16 days.")

    else:
        subject_branch_count = df_sem.groupby('SubjectCode')['Branch'].nunique()
        common_subject_codes = subject_branch_count[subject_branch_count > 1].index.tolist()
        common_subject_branches = {
            subj_code: df_sem[df_sem['SubjectCode'] == subj_code]['Branch'].unique()
            for subj_code in common_subject_codes
        }
        common_subject_codes_sorted = sorted(common_subject_codes, key=lambda x: len(common_subject_branches[x]), reverse=True)

        for subj_code in common_subject_codes_sorted:
            branches = common_subject_branches[subj_code]
            exam_day = find_next_valid_day(base_date, branches)
            mask = (df_sem['SubjectCode'] == subj_code) & (df_sem['Branch'].isin(branches))
            df_sem.loc[mask, 'Exam Date'] = exam_day.strftime("%d-%m-%Y")
            exam_day_date = exam_day.date()
            for branch in branches:
                exam_days[branch].add(exam_day_date)

        remaining_subjects = {
            branch: [
                subj for subj in df_sem[df_sem['Branch'] == branch]['Subject'].unique()
                if subj not in df_sem[(df_sem['Branch'] == branch) & (df_sem['Exam Date'] != '')]['Subject'].unique()
            ]
            for branch in all_branches
        }

        current_day = base_date
        while any(remaining_subjects.values()):
            current_day_date = current_day.date()
            if current_day.weekday() == 6 or current_day_date in holidays_dates:
                current_day += timedelta(days=1)
                continue
            
            branch_list = list(all_branches)
            random.shuffle(branch_list)
            for branch in branch_list:
                if remaining_subjects[branch] and current_day_date not in exam_days[branch]:
                    subj = remaining_subjects[branch].pop(0)
                    mask = (df_sem['Branch'] == branch) & (df_sem['Subject'] == subj)
                    df_sem.loc[mask, 'Exam Date'] = current_day.strftime("%d-%m-%Y")
                    exam_days[branch].add(current_day_date)
            current_day += timedelta(days=1)

    sem = df_sem["Semester"].iloc[0]
    if sem % 2 != 0:
        odd_sem_position = (sem + 1) // 2
        slot_str = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:
        even_sem_position = sem // 2
        slot_str = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    df_sem['Time Slot'] = slot_str
    return df_sem

def process_constraints(df, holidays, base_date, schedule_by_difficulty=False):
    final_list = []
    for sem in sorted(df["Semester"].unique()):
        if sem == 0:
            continue
        df_sem = df[df["Semester"] == sem].copy()
        if df_sem.empty:
            continue
        scheduled_sem = schedule_semester_non_electives(df_sem, holidays, base_date, schedule_by_difficulty)
        final_list.append(scheduled_sem)
    if not final_list:
        return {}
    df_combined = pd.concat(final_list, ignore_index=True)
    return {sem: df_combined[df_combined["Semester"] == sem].copy() for sem in sorted(df_combined["Semester"].unique())}

def find_next_valid_global_day(start_day, holidays_dates):
    day = start_day
    while True:
        day_date = day.date()
        if day.weekday() != 6 and day_date not in holidays_dates:
            return day
        day += timedelta(days=1)

def generate_timetable(df_non_elec, df_elec, holidays, base_date, schedule_by_difficulty=False):
    non_elec_sched = process_constraints(df_non_elec, holidays, base_date, schedule_by_difficulty)
    
    all_non_elec_dates = []
    for sem_df in non_elec_sched.values():
        dates = pd.to_datetime(sem_df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        if not dates.empty:
            all_non_elec_dates.append(dates.max())
    
    last_non_elec_date = max(all_non_elec_dates) if all_non_elec_dates else base_date

    holidays_dates = {h.date() for h in holidays}
    current_day = find_next_valid_global_day(last_non_elec_date + timedelta(days=1), holidays_dates)
    oe_dates = {}
    for oe_cat in ['OE1', 'OE2', 'OE5']:
        oe_dates[oe_cat] = current_day
        current_day = find_next_valid_global_day(current_day + timedelta(days=1), holidays_dates)

    dfs_to_concat = list(non_elec_sched.values())
    if not df_elec.empty:
        df_elec['Exam Date'] = df_elec['OE'].map(oe_dates).apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else '')
        df_elec['Time Slot'] = "10:00 AM - 1:00 PM"
        df_elec['SubjectDisplay'] = df_elec.apply(lambda row: f"{row['Subject']} [{row['OE']}]" + (f" [Duration: {row['Exam Duration']} hrs]" if pd.notna(row['Exam Duration']) and row['Exam Duration'] != 3 else ''), axis=1)
        dfs_to_concat.append(df_elec)
    
    return pd.concat(dfs_to_concat, ignore_index=True) if dfs_to_concat else pd.DataFrame()

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    df_dict = pd.read_excel(excel_path, sheet_name=None, dtype={'Time Slot': str})

    def int_to_roman(num):
        num = int(num)
        roman_values = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
        result = ""
        for value, numeral in roman_values:
            while num >= value:
                result += numeral
                num -= value
        return result

    def extract_duration(subject_str):
        match = re.search(r'\[Duration: (\d+\.?\d*) hrs\]', str(subject_str))
        return float(match.group(1)) if match else 3.0

    def get_time_slot_safely(df):
        if 'Time Slot' not in df.columns or df['Time Slot'].dropna().empty:
            return "10:00 AM - 1:00 PM"
        return str(df['Time Slot'].dropna().iloc[0]).strip()

    for sheet_name, pivot_df in df_dict.items():
        if pivot_df.empty:
            continue
        
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0].replace('_Electives', '')
        semester_str = parts[1].replace('_Electives', '') if len(parts) > 1 else ""
        
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester_roman = int_to_roman(semester_str) if semester_str.isdigit() else semester_str
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        if not sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index()
            fixed_cols = ["Exam Date"]
            sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols and "Time Slot" not in c and "index" not in c]
            line_height = 10

            for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
                chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
                cols_to_print = fixed_cols + chunk
                chunk_df = pivot_df[cols_to_print + ['Time Slot']].copy().dropna(subset=chunk, how='all')
                
                if chunk_df.empty:
                    continue

                time_slot = get_time_slot_safely(chunk_df)
                start_time, end_time = safe_split(time_slot)
                
                chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], errors='coerce').dt.strftime("%A, %d %B, %Y")

                for sub_branch in chunk:
                    if sub_branch in chunk_df:
                        chunk_df[sub_branch] = chunk_df[sub_branch].apply(lambda cell: ", ".join([
                            f"{re.sub(r' \\[Duration:.*?\\]', '', s)} ({start_time} to {calculate_end_time(start_time, extract_duration(s))})" if extract_duration(s) != 3 else re.sub(r' \\[Duration:.*?\\]', '', s)
                            for s in str(cell).split(', ')
                        ]) if pd.notna(cell) and str(cell) != '---' else '---')
                
                page_width = pdf.w - 2 * pdf.l_margin
                exam_date_width = 60
                sub_width = (page_width - exam_date_width) / max(len(chunk), 1)
                col_widths = [exam_date_width] + [sub_width] * len(chunk)
                
                pdf.add_page()
                add_footer_with_page_number(pdf, 25)
                print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=line_height, header_content=header_content, branches=chunk, time_slot=f"{start_time} - {end_time}")
        
        else: # Electives Sheet
            time_slot = get_time_slot_safely(pivot_df)
            start_time, _ = safe_split(time_slot)
            
            pivot_df["Exam Date"] = pd.to_datetime(pivot_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

            pivot_df['Subjects'] = pivot_df.apply(lambda row: ", ".join([
                f"{re.sub(r' \\[(OE|INTD).*?\\]', '', s)} ({start_time} to {calculate_end_time(start_time, extract_duration(s))})" if extract_duration(s) != 3 else re.sub(r' \\[(OE|INTD).*?\\]', '', s).strip()
                for s in str(row['SubjectDisplay']).split(', ')
            ]), axis=1)

            elective_data = pivot_df.rename(columns={'OE': 'OE Type'})
            cols_to_print = ['Exam Date', 'OE Type', 'Subjects']
            col_widths = [60, 30, pdf.w - 2 * pdf.l_margin - 90]
            
            pdf.add_page()
            add_footer_with_page_number(pdf, 25)
            print_table_custom(pdf, elective_data, cols_to_print, col_widths, line_height=10, header_content=header_content, branches=['All Streams'], time_slot=f"{start_time} - {calculate_end_time(start_time, 3.0)}")

    pdf.output(pdf_path)


def generate_pdf_timetable(semester_wise_timetable, output_pdf):
    temp_excel = "temp_timetable.xlsx"
    excel_data = save_to_excel(semester_wise_timetable)
    if excel_data:
        with open(temp_excel, "wb") as f:
            f.write(excel_data.getvalue())
        convert_excel_to_pdf(temp_excel, output_pdf)
        if os.path.exists(temp_excel):
            os.remove(temp_excel)
    else:
        st.error("No data to save to Excel.")
        return
    try:
        reader = PdfReader(output_pdf)
        writer = PdfWriter()
        for page in reader.pages:
            if page.extract_text().strip():
                writer.add_page(page)
        if writer.pages:
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
    except Exception as e:
        st.error(f"Error during PDF post-processing: {str(e)}")


def save_to_excel(semester_wise_timetable):
    if not semester_wise_timetable:
        return None

    def int_to_roman(num):
        num = int(num)
        roman_values = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
        result = ""
        for value, numeral in roman_values:
            while num >= value:
                result += numeral
                num -= value
        return result

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        all_electives = []
        for sem, df_sem in semester_wise_timetable.items():
            df_elec = df_sem[df_sem['OE'].notna() & (df_sem['OE'].str.strip() != "")].copy()
            if not df_elec.empty:
                all_electives.append(df_elec)
            
            for main_branch in df_sem["MainBranch"].unique():
                df_mb = df_sem[(df_sem["MainBranch"] == main_branch) & (df_sem['OE'].isna() | (df_sem['OE'].str.strip() == ""))].copy()
                if not df_mb.empty:
                    difficulty_str = df_mb['Difficulty'].map({0.0: 'Easy', 1.0: 'Difficult'}).fillna('')
                    duration_suffix = df_mb.apply(lambda r: f" [Duration: {r['Exam Duration']} hrs]" if pd.notna(r['Exam Duration']) and r['Exam Duration'] != 3 else '', axis=1)
                    df_mb["SubjectDisplay"] = df_mb["Subject"] + difficulty_str.apply(lambda x: f" ({x})" if x else '') + duration_suffix
                    
                    df_mb["Exam Date"] = pd.to_datetime(df_mb["Exam Date"], format="%d-%m-%Y", errors='coerce')
                    
                    pivot_df = df_mb.pivot_table(index=["Exam Date", "Time Slot"], columns="SubBranch", values="SubjectDisplay", aggfunc=lambda x: ", ".join(x)).fillna("---")
                    
                    if not pivot_df.empty:
                        pivot_df = pivot_df.sort_index(level="Exam Date")
                        pivot_df.index = pivot_df.index.set_levels(pivot_df.index.levels[0].strftime('%d-%m-%Y'), level=0)
                        sheet_name = f"{main_branch}_Sem_{int_to_roman(sem)}"
                        pivot_df.to_excel(writer, sheet_name=sheet_name[:31])

        if all_electives:
            electives_df = pd.concat(all_electives, ignore_index=True)
            electives_df['Exam Date'] = pd.to_datetime(electives_df['Exam Date'], format='%d-%m-%Y', errors='coerce').dt.strftime('%d-%m-%Y')
            electives_df.to_excel(writer, sheet_name="All_Electives", index=False)

    output.seek(0)
    return output


def save_verification_excel(original_df, semester_wise_timetable):
    if not semester_wise_timetable:
        return None

    cols_to_retain = ["School Name", "Campus", "Program", "Stream", "Current Academic Year", "Semester", "ModuleCode", "SubjectName", "Difficulty", "Category", "OE", "Exam mode", "Exam Duration"]
    verification_df = original_df[[c for c in cols_to_retain if c in original_df.columns]].copy()
    verification_df.loc[:, ["Exam Date", "Exam Time", "Is Common"]] = ["", "", ""]

    def convert_sem(sem):
        if pd.isna(sem): return 0
        m = {"Sem I": 1, "Sem II": 2, "Sem III": 3, "Sem IV": 4, "Sem V": 5, "Sem VI": 6, "Sem VII": 7, "Sem VIII": 8, "Sem IX": 9, "Sem X": 10, "Sem XI": 11}
        return m.get(str(sem).strip(), 0)
    verification_df['SemesterNum'] = verification_df["Semester"].apply(convert_sem)

    scheduled_data = pd.concat(semester_wise_timetable.values(), ignore_index=True)
    scheduled_data["ModuleCode"] = scheduled_data["Subject"].str.extract(r'\((.*?)\)', expand=False)

    for idx, row in verification_df.iterrows():
        match_mask = (scheduled_data["ModuleCode"] == row["ModuleCode"]) & (scheduled_data["Semester"] == row["SemesterNum"])
        match = scheduled_data[match_mask]
        
        if not match.empty:
            match_row = match.iloc[0]
            start_time, end_time_default = safe_split(match_row["Time Slot"])
            duration = row["Exam Duration"] if pd.notna(row["Exam Duration"]) else 3.0
            final_end_time = calculate_end_time(start_time, duration) if duration != 3.0 else end_time_default
            
            verification_df.at[idx, "Exam Date"] = match_row["Exam Date"]
            verification_df.at[idx, "Exam Time"] = f"{start_time} to {final_end_time}"
            
            branch_count = scheduled_data[match_mask]["Branch"].nunique()
            verification_df.at[idx, "Is Common"] = "YES" if branch_count > 1 else "NO"

    output = io.BytesIO()
    verification_df.drop(columns=['SemesterNum']).to_excel(output, sheet_name="Verification", index=False)
    output.seek(0)
    return output

def main():
    st.markdown("""
    <div class="main-header">
        <h1>📅 Exam Timetable Generator</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    # Initialize session state variables
    for key, default_value in [('num_custom_holidays', 1), ('custom_holidays', [None]), ('timetable_data', {}), ('processing_complete', False), ('excel_data', None), ('pdf_data', None), ('verification_data', None), ('total_exams', 0), ('total_semesters', 0), ('total_branches', 0), ('overall_date_range', 0), ('unique_exam_days', 0), ('non_elective_range', "N/A"), ('elective_dates_str', "N/A"), ('stream_counts', pd.DataFrame())]:
        if key not in st.session_state:
            st.session_state[key] = default_value

    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        st.markdown("#### 📅 Base Date for Scheduling")
        base_date = st.date_input("Start date for exams", value=datetime(2025, 4, 1))
        
        st.markdown("#### 🛠️ Scheduling Mode")
        schedule_by_difficulty = st.checkbox("Schedule by Difficulty (Alternate Easy/Difficult)", value=False)
        st.markdown(f'<div class="status-info">ℹ️ {"Alternate Easy/Difficult scheduling enabled." if schedule_by_difficulty else "Normal scheduling."}</div>', unsafe_allow_html=True)

        st.markdown('<div style="margin-top: 2rem;"></div>', unsafe_allow_html=True)

        with st.expander("Holiday Configuration", expanded=True):
            st.markdown("#### 📅 Select Predefined Holidays")
            holiday_dates = []
            if st.checkbox("April 14, 2025", value=True): holiday_dates.append(datetime(2025, 4, 14))
            if st.checkbox("May 1, 2025", value=True): holiday_dates.append(datetime(2025, 5, 1))
            if st.checkbox("August 15, 2025", value=True): holiday_dates.append(datetime(2025, 8, 15))

            st.markdown("#### 📅 Add Custom Holidays")
            for i in range(st.session_state.num_custom_holidays):
                st.session_state.custom_holidays[i] = st.date_input(f"Custom Holiday {i + 1}", value=st.session_state.custom_holidays[i], key=f"custom_holiday_{i}")
            
            if st.button("➕ Add Another Holiday"):
                st.session_state.num_custom_holidays += 1
                st.session_state.custom_holidays.append(None)
                st.rerun()

            holiday_dates.extend([datetime.combine(h, datetime.min.time()) for h in st.session_state.custom_holidays if h])
            if holiday_dates:
                st.markdown("#### Selected Holidays:")
                for holiday in sorted(holiday_dates):
                    st.write(f"• {holiday.strftime('%B %d, %Y')}")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("""<div class="upload-section"><h3>📁 Upload Excel File</h3><p>Upload your timetable data file (.xlsx format)</p></div>""", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'], help="Upload the Excel file containing your timetable data")
        if uploaded_file:
            st.success(f"✅ File '{uploaded_file.name}' uploaded successfully!")

    with col2:
        st.markdown("""<div class="feature-card"><h4>🚀 Features</h4><ul><li>📊 Excel file processing</li><li>📅 Smart scheduling</li><li>📋 Multiple output formats</li><li>🎯 Conflict resolution</li><li>📱 Mobile-friendly interface</li></ul></div>""", unsafe_allow_html=True)

    if uploaded_file:
        if st.button("🔄 Generate Timetable", type="primary", use_container_width=True):
            with st.spinner("Processing your timetable... Please wait..."):
                try:
                    df_non_elec, df_elec, original_df = read_timetable(uploaded_file)
                    if df_non_elec is not None:
                        final_df = generate_timetable(df_non_elec, df_elec, set(holiday_dates), datetime.combine(base_date, datetime.min.time()), schedule_by_difficulty)
                        if not final_df.empty:
                            final_df["Exam Date"] = pd.to_datetime(final_df["Exam Date"], format="%d-%m-%Y", errors='coerce')
                            final_df.dropna(subset=['Exam Date'], inplace=True)
                            final_df.sort_values(["Exam Date", "Semester", "MainBranch"], inplace=True)
                            
                            sem_dict = {s: final_df[final_df["Semester"] == s].copy() for s in sorted(final_df["Semester"].unique())}
                            st.session_state.timetable_data = sem_dict
                            st.session_state.original_df = original_df
                            st.session_state.processing_complete = True

                            # Compute and store statistics
                            all_dates = final_df['Exam Date'].dropna()
                            st.session_state.total_exams = len(final_df)
                            st.session_state.total_semesters = final_df['Semester'].nunique()
                            st.session_state.total_branches = final_df['MainBranch'].nunique()
                            st.session_state.overall_date_range = (all_dates.max() - all_dates.min()).days if not all_dates.empty else 0
                            st.session_state.unique_exam_days = all_dates.dt.date.nunique()
                            
                            # Generate and cache downloadable files
                            st.session_state.excel_data = save_to_excel(sem_dict)
                            st.session_state.verification_data = save_verification_excel(original_df, sem_dict)
                            
                            pdf_output_path = "timetable_final.pdf"
                            generate_pdf_timetable(sem_dict, pdf_output_path)
                            with open(pdf_output_path, "rb") as f:
                                st.session_state.pdf_data = f.read()
                            if os.path.exists(pdf_output_path):
                                os.remove(pdf_output_path)

                            st.success("🎉 Timetable generated successfully!")
                        else:
                            st.error("❌ No valid data found to process.")
                    else:
                        st.error("❌ Failed to read the Excel file. Please check the format.")
                except Exception as e:
                    st.error(f"An error occurred: {e}", icon="🔥")

    if st.session_state.processing_complete:
        st.markdown("---")
        if st.session_state.unique_exam_days > 16:
            st.warning(f"⚠️ The timetable spans {st.session_state.unique_exam_days} exam days, exceeding the 16-day recommended limit.")
        
        st.markdown("### 📥 Download Options")
        d_cols = st.columns(4)
        if st.session_state.excel_data:
            d_cols[0].download_button("📊 Download Excel", st.session_state.excel_data.getvalue(), f"timetable_{datetime.now():%Y%m%d_%H%M%S}.xlsx", use_container_width=True)
        if st.session_state.pdf_data:
            d_cols[1].download_button("📄 Download PDF", st.session_state.pdf_data, f"timetable_{datetime.now():%Y%m%d_%H%M%S}.pdf", use_container_width=True)
        if st.session_state.verification_data:
            d_cols[2].download_button("📋 Download Verification", st.session_state.verification_data.getvalue(), f"verification_{datetime.now():%Y%m%d_%H%M%S}.xlsx", use_container_width=True)
        
        if d_cols[3].button("🔄 Generate New", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key not in ['num_custom_holidays', 'custom_holidays']: # Persist some settings
                    del st.session_state[key]
            st.rerun()

        st.markdown("---")
        st.markdown("<h2>📊 Timetable Results</h2>", unsafe_allow_html=True)
        for sem, df_sem in st.session_state.timetable_data.items():
            st.markdown(f"### 📚 Semester {sem}")
            for main_branch in df_sem["MainBranch"].unique():
                st.markdown(f"#### {BRANCH_FULL_FORM.get(main_branch, main_branch)}")
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()

                def get_display_suffix(row):
                    suffix = ""
                    if pd.notna(row['Exam Duration']) and row['Exam Duration'] != 3:
                        start_time, _ = safe_split(row['Time Slot'])
                        end_time = calculate_end_time(start_time, row['Exam Duration'])
                        suffix += f" ({start_time} to {end_time})"
                    if pd.notna(row['Difficulty']):
                        suffix += f" ({'Easy' if row['Difficulty'] == 0 else 'Difficult'})"
                    return suffix

                df_mb["SubjectDisplay"] = df_mb["Subject"] + df_mb.apply(get_display_suffix, axis=1)
                
                pivot_df = df_mb.pivot_table(index=["Exam Date", "Time Slot"], columns="SubBranch", values="SubjectDisplay", aggfunc=', '.join).fillna("---")
                if not pivot_df.empty:
                    pivot_df.index = pivot_df.index.set_levels(pivot_df.index.levels[0].strftime('%A, %d-%m-%Y'), level=0)
                    st.dataframe(pivot_df, use_container_width=True)


    st.markdown("---")
    st.markdown("""<div class="footer"><p>🎓 <strong>Exam Timetable Generator</strong></p><p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p></div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
