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
            padding:_legend: 2rem;
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

class RealTimeOptimizer:
    """Handles real-time optimization during scheduling"""
    
    def __init__(self, branches, holidays, time_slots=None):
        self.branches = branches
        self.holidays = holidays
        self.time_slots = time_slots or ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        self.schedule_grid = {}  # date -> time_slot -> branch -> subject/None
        self.optimization_log = []
        self.moves_made = 0
        
    def add_exam_to_grid(self, date_str, time_slot, branch, subject):
        """Add an exam to the schedule grid"""
        if date_str not in self.schedule_grid:
            self.schedule_grid[date_str] = {}
        if time_slot not in self.schedule_grid[date_str]:
            self.schedule_grid[date_str][time_slot] = {}
        self.schedule_grid[date_str][time_slot][branch] = subject
    
    def find_earliest_empty_slot(self, branch, start_date, preferred_time_slot=None):
        """Find the earliest empty slot for a branch - ensuring only one exam per day per branch"""
        sorted_dates = sorted(self.schedule_grid.keys(), 
                            key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        
        for date_str in sorted_dates:
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
            
            if date_obj < start_date:
                continue
            
            if date_obj.weekday() == 6 or date_obj.date() in self.holidays:
                continue
            
            branch_has_exam_today = False
            if date_str in self.schedule_grid:
                for time_slot in self.time_slots:
                    if (time_slot in self.schedule_grid[date_str] and
                        branch in self.schedule_grid[date_str][time_slot] and
                        self.schedule_grid[date_str][time_slot][branch] is not None):
                        branch_has_exam_today = True
                        break
            
            if branch_has_exam_today:
                continue
            
            if preferred_time_slot:
                if (date_str in self.schedule_grid and 
                    preferred_time_slot in self.schedule_grid[date_str]):
                    if self.schedule_grid[date_str][preferred_time_slot].get(branch) is None:
                        return date_str, preferred_time_slot
            
            for time_slot in self.time_slots:
                if date_str not in self.schedule_grid:
                    return date_str, time_slot
                    
                if time_slot not in self.schedule_grid[date_str]:
                    return date_str, time_slot
                    
                if self.schedule_grid[date_str][time_slot].get(branch) is None:
                    return date_str, time_slot
        
        return None, None
    
    def initialize_grid_with_empty_days(self, start_date, num_days=40):
        """Pre-populate grid with empty days"""
        current_date = start_date
        for _ in range(num_days):
            if current_date.weekday() != 6 and current_date.date() not in self.holidays:
                date_str = current_date.strftime("%d-%m-%Y")
                if date_str not in self.schedule_grid:
                    self.schedule_grid[date_str] = {}
                for time_slot in self.time_slots:
                    if time_slot not in self.schedule_grid[date_str]:
                        self.schedule_grid[date_str][time_slot] = {branch: None for branch in self.branches}
            current_date += timedelta(days=1)
    
    def get_schedule_summary(self):
        """Get a summary of the current schedule"""
        total_slots = 0
        filled_slots = 0
        
        for date_str, time_slots in self.schedule_grid.items():
            for time_slot, branches in time_slots.items():
                for branch, subject in branches.items():
                    total_slots += 1
                    if subject is not None:
                        filled_slots += 1
        
        return {
            'total_slots': total_slots,
            'filled_slots': filled_slots,
            'empty_slots': total_slots - filled_slots,
            'utilization': (filled_slots / total_slots * 100) if total_slots > 0 else 0
        }

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
    
    footer_height = 25
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    pdf.set_font("Arial", size=13)
    pdf.set_xy(10, pdf.h - footer_height + 7)
    pdf.cell(0, 5, "Signature", 0, 1, 'L')
    
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')
    
    header_height = 85
    pdf.set_y(0)
    current_date = datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST")
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
    logo_width = 45
    logo_x = (pdf.w - logo_width) / 2
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
    
    available_height = pdf.h - pdf.t_margin - footer_height - header_height
    pdf.set_font("Arial", size=12)
    
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row):
            continue
            
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 2 * 2
            lines = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines
        
        if pdf.get_y() + row_h > pdf.h - footer_height:
            add_footer_with_page_number(pdf, footer_height)
            pdf.add_page()
            add_footer_with_page_number(pdf, footer_height)
            add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, branches, time_slot)
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
    
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
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

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    
    pdf.alias_nb_pages()
    
    df_dict = pd.read_excel(excel_path, sheet_name=None, index_col=[0, 1])

    def int_to_roman(num):
        roman_values = [
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        ]
        result = ""
        for value, numeral in roman_values:
            while num >= value:
                result += numeral
                num -= value
        return result

    def extract_duration(subject_str):
        match = re.search(r'\[Duration: (\d+\.?\d*) hrs\]', subject_str)
        if match:
            return float(match.group(1))
        else:
            return 3.0

    def get_semester_default_time_slot(semester, branch):
        default_timings = {
            ("5", "EXTC"): "10:00 AM - 1:00 PM",
            ("6", "COMP"): "10:00 AM - 1:00 PM",
        }
        return default_timings.get((semester, branch), "10:00 AM - 1:00 PM")

    for sheet_name, pivot_df in df_dict.items():
        if pivot_df.empty:
            continue
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        semester_roman = semester if not semester.isdigit() else int_to_roman(int(semester))
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        if not sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            fixed_cols = ["Exam Date", "Time Slot"]
            sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols]
            exam_date_width = 60
            table_font_size = 12
            line_height = 10

            for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
                chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
                cols_to_print = fixed_cols[:1] + chunk
                chunk_df = pivot_df[fixed_cols + chunk].copy()
                mask = chunk_df[chunk].apply(lambda row: row.astype(str).str.strip() != "").any(axis=1)
                chunk_df = chunk_df[mask].reset_index(drop=True)
                if chunk_df.empty:
                    continue

                time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and not pivot_df['Time Slot'].empty else None

                chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

                default_time_slot = get_semester_default_time_slot(semester, main_branch)
                for sub_branch in chunk:
                    for idx in chunk_df.index:
                        cell_value = chunk_df.at[idx, sub_branch]
                        if pd.isna(cell_value) or cell_value.strip() == "---":
                            continue
                        subjects = cell_value.split(", ")
                        modified_subjects = []
                        for subject in subjects:
                            duration = extract_duration(subject)
                            base_subject = re.sub(r' \[Duration: \d+\.?\d* hrs\]', '', subject)
                            subject_time_slot = chunk_df.at[idx, "Time Slot"] if pd.notna(chunk_df.at[idx, "Time Slot"]) else None
                            if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                                if subject_time_slot != default_time_slot:
                                    modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                                else:
                                    modified_subjects.append(base_subject)
                            else:
                                modified_subjects.append(base_subject)
                            if duration != 3 and subject_time_slot and subject_time_slot.strip():
                                start_time = subject_time_slot.split(" - ")[0]
                                end_time = calculate_end_time(start_time, duration)
                                modified_subjects[-1] = f"{modified_subjects[-1]} ({start_time} to {end_time})"
                        chunk_df.at[idx, sub_branch] = ", ".join(modified_subjects)

                page_width = pdf.w - 2 * pdf.l_margin
                remaining = page_width - exam_date_width
                sub_width = remaining / max(len(chunk), 1)
                col_widths = [exam_date_width] + [sub_width] * len(chunk)
                total_w = sum(col_widths)
                if total_w > page_width:
                    factor = page_width / total_w
                    col_widths = [w * factor for w in col_widths]
                
                pdf.add_page()
                footer_height = 25
                add_footer_with_page_number(pdf, footer_height)
                
                print_table_custom(pdf, chunk_df[cols_to_print], cols_to_print, col_widths, line_height=line_height, 
                                 header_content=header_content, branches=chunk, time_slot=time_slot)

        if sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            
            elective_data = pivot_df.groupby(['OE', 'Exam Date', 'Time Slot']).agg({
                'SubjectDisplay': lambda x: ", ".join(x)
            }).reset_index()

            elective_data["Exam Date"] = pd.to_datetime(elective_data["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

            elective_data['SubjectDisplay'] = elective_data.apply(
                lambda row: ", ".join([s.replace(f" [{row['OE']}]", "") for s in row['SubjectDisplay'].split(", ")]),
                axis=1
            )

            default_time_slot = get_semester_default_time_slot(semester, main_branch)
            time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and not pivot_df['Time Slot'].empty else None
            for idx in elective_data.index:
                cell_value = elective_data.at[idx, 'SubjectDisplay']
                if pd.isna(cell_value) or cell_value.strip() == "---":
                    continue
                subjects = cell_value.split(", ")
                modified_subjects = []
                for subject in subjects:
                    duration = extract_duration(subject)
                    base_subject = re.sub(r' \[Duration: \d+\.?\d* hrs\]', '', subject)
                    subject_time_slot = elective_data.at[idx, "Time Slot"] if pd.notna(elective_data.at[idx, "Time Slot"]) else None
                    if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                        if subject_time_slot != default_time_slot:
                            modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                        else:
                            modified_subjects.append(base_subject)
                    else:
                        modified_subjects.append(base_subject)
                    if duration != 3 and subject_time_slot and subject_time_slot.strip():
                        start_time = subject_time_slot.split(" - ")[0]
                        end_time = calculate_end_time(start_time, duration)
                        modified_subjects[-1] = f"{modified_subjects[-1]} ({start_time} to {end_time})"
                elective_data.at[idx, 'SubjectDisplay'] = ", ".join(modified_subjects)

            elective_data = elective_data.rename(columns={'OE': 'OE Type', 'SubjectDisplay': 'Subjects'})

            exam_date_width = 60
            oe_width = 30
            subject_width = pdf.w - 2 * pdf.l_margin - exam_date_width - oe_width
            col_widths = [exam_date_width, oe_width, subject_width]
            cols_to_print = ['Exam Date', 'OE Type', 'Subjects']
            
            pdf.add_page()
            footer_height = 25
            add_footer_with_page_number(pdf, footer_height)
            
            print_table_custom(pdf, elective_data, cols_to_print, col_widths, line_height=10, 
                             header_content=header_content, branches=['All Streams'], time_slot=time_slot)

    pdf.output(pdf_path)

def generate_pdf_timetable(semester_wise_timetable, output_pdf):
    temp_excel = os.path.join(os.path.dirname(output_pdf), "temp_timetable.xlsx")
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
        page_number_pattern = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*$')
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            try:
                text = page.extract_text() if page else ""
            except:
                text = ""
            cleaned_text = text.strip() if text else ""
            is_blank_or_page_number = (
                    not cleaned_text or
                    page_number_pattern.match(cleaned_text) or
                    len(cleaned_text) <= 10
            )
            if not is_blank_or_page_number:
                writer.add_page(page)
        if len(writer.pages) > 0:
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
        else:
            st.warning("Warning: All pages were filtered out - keeping original PDF")
    except Exception as e:
        st.error(f"Error during PDF post-processing: {str(e)}")

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
            "Student count": "StudentCount",
            "Is Common": "IsCommon"
        })
        
        def convert_sem(sem):
            if pd.isna(sem):
                return 0
            m = {
                "Sem I": 1, "Sem II": 2, "Sem III": 3, "Sem IV": 4,
                "Sem V": 5, "Sem VI": 6, "Sem VII": 7, "Sem VIII": 8,
                "Sem IX": 9, "Sem X": 10, "Sem XI": 11
            }
            return m.get(sem.strip(), 0)
        
        df["Semester"] = df["Semester"].apply(convert_sem).astype(int)
        df["Branch"] = df["Program"].astype(str).str.strip() + "-" + df["Stream"].astype(str).str.strip()
        df["Subject"] = df["SubjectName"].astype(str) + " - (" + df["ModuleCode"].astype(str) + ")"
        
        comp_mask = (df["Category"] == "COMP") & df["Difficulty"].notna()
        df["Difficulty"] = None
        df.loc[comp_mask, "Difficulty"] = df.loc[comp_mask, "Difficulty"]
        
        df["Exam Date"] = ""
        df["Time Slot"] = ""
        df["Exam Duration"] = df["Exam Duration"].fillna(3).astype(float)
        df["StudentCount"] = df["StudentCount"].fillna(0).astype(int)
        df["IsCommon"] = df["IsCommon"].fillna("NO").str.strip().str.upper()
        
        df_non = df[df["Category"] != "INTD"].copy()
        df_ele = df[df["Category"] == "INTD"].copy()
        
        def split_br(b):
            p = b.split("-", 1)
            return pd.Series([p[0].strip(), p[1].strip() if len(p) > 1 else ""])
        
        for d in (df_non, df_ele):
            d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)
        
        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty", "Exam Duration", "StudentCount", "IsCommon", "ModuleCode"]
        
        return df_non[cols], df_ele[cols], df
        
    except Exception as e:
        st.error(f"Error reading the Excel file: {str(e)}")
        return None, None, None

def schedule_semester_non_electives_with_optimization(df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty=False):
    """Enhanced scheduling that fills empty slots in real-time and ensures all subjects are scheduled"""
    
    sem = df_sem["Semester"].iloc[0]
    if sem % 2 != 0:
        odd_sem_position = (sem + 1) // 2
        preferred_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:
        even_sem_position = sem // 2
        preferred_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    
    df_sem = df_sem.drop_duplicates(subset=['Branch', 'Subject', 'Category', 'ModuleCode'])
    
    comp_subjects = df_sem[(df_sem['Category'] == 'COMP') & (df_sem['IsCommon'] == 'NO') & ((df_sem['Exam Date'] == "") | (df_sem['Exam Date'].isna()))]
    
    for idx, row in comp_subjects.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        
        empty_date, empty_slot = optimizer.find_earliest_empty_slot(branch, base_date, preferred_slot)
        
        if empty_date:
            df_sem.at[idx, 'Exam Date'] = empty_date
            df_sem.at[idx, 'Time Slot'] = empty_slot
            optimizer.add_exam_to_grid(empty_date, empty_slot, branch, subject)
            exam_days[branch].add(datetime.strptime(empty_date, "%d-%m-%Y").date())
            optimizer.optimization_log.append(f"✅ Filled empty slot: {subject} on {empty_date} at {empty_slot}")
            optimizer.moves_made += 1
        else:
            current_date = base_date
            scheduled = False
            
            while not scheduled:
                date_str = current_date.strftime("%d-%m-%Y")
                
                if current_date.weekday() != 6 and current_date.date() not in holidays:
                    if current_date.date() not in exam_days[branch]:
                        df_sem.at[idx, 'Exam Date'] = date_str
                        df_sem.at[idx, 'Time Slot'] = preferred_slot
                        optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                        exam_days[branch].add(current_date.date())
                        scheduled = True
                
                current_date += timedelta(days=1)
    
    elec_subjects = df_sem[(df_sem['Category'] == 'ELEC') & (df_sem['IsCommon'] == 'NO') & ((df_sem['Exam Date'] == "") | (df_sem['Exam Date'].isna()))]
    
    for idx, row in elec_subjects.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        
        empty_date, empty_slot = optimizer.find_earliest_empty_slot(branch, base_date, preferred_slot)
        
        if empty_date:
            df_sem.at[idx, 'Exam Date'] = empty_date
            df_sem.at[idx, 'Time Slot'] = empty_slot
            optimizer.add_exam_to_grid(empty_date, empty_slot, branch, subject)
            exam_days[branch].add(datetime.strptime(empty_date, "%d-%m-%Y").date())
            optimizer.optimization_log.append(f"✅ Filled empty slot: ELEC {subject} on {empty_date} at {empty_slot}")
            optimizer.moves_made += 1
        else:
            current_date = base_date
            scheduled = False
            
            while not scheduled:
                date_str = current_date.strftime("%d-%m-%Y")
                
                if current_date.weekday() != 6 and current_date.date() not in holidays:
                    if current_date.date() not in exam_days[branch]:
                        df_sem.at[idx, 'Exam Date'] = date_str
                        df_sem.at[idx, 'Time Slot'] = preferred_slot
                        optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                        exam_days[branch].add(current_date.date())
                        scheduled = True
                        optimizer.optimization_log.append(f"📅 Scheduled ELEC {subject} for {branch} on {date_str}")
                
                current_date += timedelta(days=1)
    
    mask_empty_time = (df_sem['Time Slot'] == "") | (df_sem['Time Slot'].isna())
    df_sem.loc[mask_empty_time, 'Time Slot'] = preferred_slot
    
    unscheduled_mask = (df_sem['Exam Date'] == "") | (df_sem['Exam Date'].isna()) | (df_sem['Exam Date'].str.strip() == "")
    unscheduled = df_sem[unscheduled_mask]
    
    if not unscheduled.empty:
        print(f"DEBUG: Found {len(unscheduled)} truly unscheduled subjects in semester {sem}")
        print("Unscheduled subjects:")
        for idx, row in unscheduled.iterrows():
            print(f"  - {row['Subject']} ({row['Branch']}) - Exam Date: '{row['Exam Date']}'")
        
        for idx, row in unscheduled.iterrows():
            branch = row['Branch']
            subject = row['Subject']
            current_date = base_date
            
            while True:
                date_str = current_date.strftime("%d-%m-%Y")
                if current_date.weekday() != 6 and current_date.date() not in holidays:
                    if current_date.date() not in exam_days[branch]:
                        df_sem.at[idx, 'Exam Date'] = date_str
                        df_sem.at[idx, 'Time Slot'] = preferred_slot
                        optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                        exam_days[branch].add(current_date.date())
                        optimizer.optimization_log.append(f"🔧 Force scheduled {subject} on {date_str}")
                        break
                current_date += timedelta(days=1)
    else:
        print(f"DEBUG: All subjects in semester {sem} are properly scheduled")
    
    return df_sem

def process_constraints_with_real_time_optimization(df, holidays, base_date, schedule_by_difficulty=False):
    """Enhanced process_constraints that fills empty slots during scheduling and removes duplicates"""
    df = df.drop_duplicates(subset=['Branch', 'Subject', 'Category', 'ModuleCode', 'Semester'])
    
    all_branches = df['Branch'].unique()
    exam_days = {branch: set() for branch in all_branches}
    
    optimizer = RealTimeOptimizer(all_branches, holidays)
    optimizer.initialize_grid_with_empty_days(base_date, num_days=40)
    
    st.info(f"🔧 Scheduling {len(df)} subjects across {len(all_branches)} branches...")
    
    def find_earliest_available_slot_with_empty_check(start_day, for_branches, subject, optimizer):
        """Enhanced slot finding that checks empty slots first"""
        if len(for_branches) == 1:
            empty_date, empty_slot = optimizer.find_earliest_empty_slot(for_branches[0], start_day)
            if empty_date:
                return datetime.strptime(empty_date, "%d-%m-%Y"), empty_slot
        else:
            sorted_dates = sorted(optimizer.schedule_grid.keys(), 
                                key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
            
            for date_str in sorted_dates:
                date_obj = datetime.strptime(date_str, "%d-%m-%Y")
                if date_obj < start_day:
                    continue
                    
                for time_slot in optimizer.time_slots:
                    all_empty = True
                    for branch in for_branches:
                        if (date_str in optimizer.schedule_grid and 
                            time_slot in optimizer.schedule_grid[date_str] and
                            optimizer.schedule_grid[date_str][time_slot].get(branch) is not None):
                            all_empty = False
                            break
                    
                    if all_empty:
                        optimizer.optimization_log.append(f"✅ Found common empty slot for {subject} on {date_str}")
                        return date_obj, time_slot
        
        current_date = start_day
        while True:
            current_date_only = current_date.date()
            
            if current_date.weekday() == 6 or current_date_only in holidays:
                current_date += timedelta(days=1)
                continue
            
            if all(current_date_only not in exam_days[branch] for branch in for_branches):
                return current_date, None
            
            current_date += timedelta(days=1)

    comp_common = len(df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'YES')])
    comp_individual = len(df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'NO')])
    elec_common = len(df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'YES')])
    elec_individual = len(df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'NO')])
    
    st.write(f"📊 Subject distribution: COMP (Common: {comp_common}, Individual: {comp_individual}), ELEC (Common: {elec_common}, Individual: {elec_individual})")

    common_comp = df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'YES')]
    for module_code, group in common_comp.groupby('ModuleCode'):
        group = group.drop_duplicates(subset=['Branch', 'Subject'])
        branches = group['Branch'].unique()
        subject = group['Subject'].iloc[0]
        
        exam_day, slot_found = find_earliest_available_slot_with_empty_check(base_date, branches, subject, optimizer)
        
        min_sem = group['Semester'].min()
        if slot_found:
            slot_str = slot_found
        else:
            if min_sem % 2 != 0:
                odd_sem_position = (min_sem + 1) // 2
                slot_str = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
            else:
                even_sem_position = min_sem // 2
                slot_str = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
        
        date_str = exam_day.strftime("%d-%m-%Y")
        df.loc[group.index, 'Exam Date'] = date_str
        df.loc[group.index, 'Time Slot'] = slot_str
        
        for branch in branches:
            exam_days[branch].add(exam_day.date())
            optimizer.add_exam_to_grid(date_str, slot_str, branch, subject)

    common_elec = df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'YES')]
    for module_code, group in common_elec.groupby('ModuleCode'):
        group = group.drop_duplicates(subset=['Branch', 'Subject'])
        branches = group['Branch'].unique()
        subject = group['Subject'].iloc[0]
        
        exam_day, slot_found = find_earliest_available_slot_with_empty_check(base_date, branches, subject, optimizer)
        
        min_sem = group['Semester'].min()
        if slot_found:
            slot_str = slot_found
        else:
            if min_sem % 2 != 0:
                odd_sem_position = (min_sem + 1) // 2
                slot_str = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
            else:
                even_sem_position = min_sem // 2
                slot_str = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
        
        date_str = exam_day.strftime("%d-%m-%Y")
        df.loc[group.index, 'Exam Date'] = date_str
        df.loc[group.index, 'Time Slot'] = slot_str
        
        for branch in branches:
            exam_days[branch].add(exam_day.date())
            optimizer.add_exam_to_grid(date_str, slot_str, branch, subject)

    final_list = []
    for sem in sorted(df["Semester"].unique()):
        if sem == 0:
            continue
        df_sem = df[df["Semester"] == sem].copy()
        if df_sem.empty:
            continue
            
        unscheduled_before_mask = ((df_sem['Exam Date'] == "") | (df_sem['Exam Date'].isna()) | (df_sem['Exam Date'].str.strip() == "")) & (df_sem['IsCommon'] == 'NO')
        unscheduled_before = len(df_sem[unscheduled_before_mask])
        
        print(f"DEBUG: Semester {sem} - Unscheduled before: {unscheduled_before}")
        print(f"DEBUG: Total subjects in semester: {len(df_sem)}")
        print(f"DEBUG: Common subjects already scheduled: {len(df_sem[df_sem['IsCommon'] == 'YES'])}")
        
        scheduled_sem = schedule_semester_non_electives_with_optimization(
            df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty
        )
        
        unscheduled_after_mask = (scheduled_sem['Exam Date'] == "") | (scheduled_sem['Exam Date'].isna()) | (scheduled_sem['Exam Date'].str.strip() == "")
        unscheduled_after = len(scheduled_sem[unscheduled_after_mask])
        
        print(f"DEBUG: Semester {sem} - Unscheduled after: {unscheduled_after}")
        
        if unscheduled_after > 0:
            unscheduled_subjects = scheduled_sem[unscheduled_after_mask]
            print(f"DEBUG: Unscheduled subjects details:")
            for idx, row in unscheduled_subjects.iterrows():
                print(f"  - {row['Subject']} ({row['Branch']}) - Category: {row['Category']} - IsCommon: {row['IsCommon']} - Exam Date: '{row['Exam Date']}'")
        final_list.append(scheduled_sem)

    if not final_list:
        return {}

    df_combined = pd.concat(final_list, ignore_index=True)
    
    df_combined = df_combined.drop_duplicates(subset=['Branch', 'Subject', 'Exam Date', 'Time Slot'])
    
    unscheduled_final_mask = (df_combined['Exam Date'] == "") | (df_combined['Exam Date'].isna()) | (df_combined['Exam Date'].str.strip() == "")
    unscheduled_final = df_combined[unscheduled_final_mask]
    
    if not unscheduled_final.empty:
        st.error(f"❌ {len(unscheduled_final)} subjects remain unscheduled!")
        with st.expander("View unscheduled subjects"):
            st.dataframe(unscheduled_final[['Branch', 'Subject', 'Category', 'IsCommon', 'Exam Date']])
    
    schedule_summary = optimizer.get_schedule_summary()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Slots Filled", optimizer.moves_made)
    with col2:
        st.metric("Empty Slots Used", f"{schedule_summary['utilization']:.1f}%")
    with col3:
        st.metric("Total Subjects Scheduled", len(df_combined[~unscheduled_final_mask]))
    
    if optimizer.moves_made > 0:
        with st.expander("📝 Optimization Log", expanded=False):
            for log in optimizer.optimization_log[-20:]:
                st.write(log)
    
    sem_dict = {}
    for sem in sorted(df_combined["Semester"].unique()):
        sem_dict[sem] = df_combined[df_combined["Semester"] == sem].copy()

    all_dates = pd.to_datetime(df_combined['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    if not all_dates.empty:
        start_date = min(all_dates)
        end_date = max(all_dates)
        total_span = (end_date - start_date).days + 1
        
        if total_span <= 16:
            st.success(f"✅ Timetable optimized successfully! Total span: {total_span} days (within 16-day target)")
        elif total_span <= 20:
            st.info(f"ℹ️ Timetable span: {total_span} days (within 20-day limit but above 16-day target)")
        else:
            st.warning(f"⚠️ The timetable spans {total_span} days, exceeding the limit of 20 days.")

    return sem_dict

def find_next_valid_day_for_electives(start_day, holidays):
    """Find the next valid day for scheduling electives (skip weekends and holidays)"""
    day = start_day
    while True:
        day_date = day.date()
        if day.weekday() == 6 or day_date in holidays:
            day += timedelta(days=1)
            continue
        return day

def optimize_oe_subjects_after_scheduling(sem_dict, holidays, optimizer=None):
    """After main scheduling, check if OE subjects can be moved to earlier empty slots."""
    if not sem_dict:
        return sem_dict
    
    st.info("🎯 Optimizing Open Elective (OE) placement...")
    
    all_data = pd.concat(sem_dict.values(), ignore_index=True)
    
    oe_data = all_data[all_data['OE'].notna() & (all_data['OE'].str.strip() != "")]
    non_oe_data = all_data[all_data['OE'].isna() | (all_data['OE'].str.strip() == "")]
    
    if oe_data.empty:
        st.info("No OE subjects to optimize")
        return sem_dict
    
    schedule_grid = {}
    branches = all_data['Branch'].unique()
    
    for _, row in all_data.iterrows():
        if pd.notna(row['Exam Date']):
            if isinstance(row['Exam Date'], pd.Timestamp):
                date_str = row['Exam Date'].strftime("%d-%m-%Y")
            else:
                date_str = str(row['Exam Date'])
            
            if date_str not in schedule_grid:
                schedule_grid[date_str] = {}
            if row['Time Slot'] not in schedule_grid[date_str]:
                schedule_grid[date_str][row['Time Slot']] = {}
            schedule_grid[date_str][row['Time Slot']][row['Branch']] = row['Subject']
    
    all_dates = sorted(schedule_grid.keys(), 
                      key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    if not all_dates:
        return sem_dict
    
    start_date = datetime.strptime(all_dates[0], "%d-%m-%Y")
    end_date = datetime.strptime(all_dates[-1], "%d-%m-%Y")
    
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() != 6 and current_date.date() not in holidays:
            date_str = current_date.strftime("%d-%m-%Y")
            if date_str not in schedule_grid:
                schedule_grid[date_str] = {}
            for time_slot in ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]:
                if time_slot not in schedule_grid[date_str]:
                    schedule_grid[date_str][time_slot] = {branch: None for branch in branches}
        current_date += timedelta(days=1)
    
    oe_data_copy = oe_data.copy()
    oe_data_copy['Exam Date'] = oe_data_copy['Exam Date'].apply(
        lambda x: x.strftime("%d-%m-%Y") if isinstance(x, pd.Timestamp) else str(x)
    )
    
    oe1_oe5_data = oe_data_copy[oe_data_copy['OE'].isin(['OE1', 'OE5'])]
    oe2_data = oe_data_copy[oe_data_copy['OE'] == 'OE2']
    
    moves_made = 0
    optimization_log = []
    
    if not oe1_oe5_data.empty:
        current_oe1_oe5_date = oe1_oe5_data['Exam Date'].iloc[0]
        current_oe1_oe5_slot = oe1_oe5_data['Time Slot'].iloc[0]
        current_oe1_oe5_date_obj = datetime.strptime(current_oe1_oe5_date, "%d-%m-%Y")
        
        affected_branches = oe1_oe5_data['Branch'].unique()
        
        best_oe1_oe5_date = None
        best_oe1_oe5_slot = None
        
        sorted_dates = sorted(schedule_grid.keys(), 
                            key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        
        for check_date_str in sorted_dates:
            check_date_obj = datetime.strptime(check_date_str, "%d-%m-%Y")
            
            if check_date_obj >= current_oe1_oe5_date_obj:
                break
            
            if check_date_obj.weekday() == 6 or check_date_obj.date() in holidays:
                continue
            
            next_day = find_next_valid_day_for_electives(check_date_obj + timedelta(days=1), holidays)
            next_day_str = next_day.strftime("%d-%m-%Y")
            
            for time_slot in ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]:
                can_move_oe1_oe5 = True
                
                for branch in affected_branches:
                    if (check_date_str in schedule_grid and 
                        time_slot in schedule_grid[check_date_str] and
                        branch in schedule_grid[check_date_str][time_slot] and
                        schedule_grid[check_date_str][time_slot][branch] is not None):
                        can_move_oe1_oe5 = False
                        break
                
                if can_move_oe1_oe5:
                    if not oe2_data.empty:
                        oe2_branches = oe2_data['Branch'].unique()
                        can_move_oe2 = False
                        
                        for oe2_slot in ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]:
                            oe2_can_fit = True
                            for oe2_branch in oe2_branches:
                                if (next_day_str in schedule_grid and 
                                    oe2_slot in schedule_grid[next_day_str] and
                                    oe2_branch in schedule_grid[next_day_str][oe2_slot] and
                                    schedule_grid[next_day_str][oe2_slot][oe2_branch] is not None):
                                    oe2_can_fit = False
                                    break
                            
                            if oe2_can_fit:
                                can_move_oe2 = True
                                best_oe2_slot = oe2_slot
                                break
                        
                        if can_move_oe2:
                            best_oe1_oe5_date = check_date_str
                            best_oe1_oe5_slot = time_slot
                            best_oe2_date = next_day_str
                            break
                    else:
                        best_oe1_oe5_date = check_date_str
                        best_oe1_oe5_slot = time_slot
                        break
            
            if best_oe1_oe5_date:
                break
        
        if best_oe1_oe5_date and best_oe1_oe5_date != current_oe1_oe5_date:
            days_saved = (current_oe1_oe5_date_obj - datetime.strptime(best_oe1_oe5_date, "%d-%m-%Y")).days
            
            for idx in oe1_oe5_data.index:
                sem = all_data.at[idx, 'Semester']
                branch = all_data.at[idx, 'Branch']
                subject = all_data.at[idx, 'Subject']
                
                mask = (sem_dict[sem]['Subject'] == subject) & \
                       (sem_dict[sem]['Branch'] == branch)
                sem_dict[sem].loc[mask, 'Exam Date'] = best_oe1_oe5_date
                sem_dict[sem].loc[mask, 'Time Slot'] = best_oe1_oe5_slot
                
                if (current_oe1_oe5_date in schedule_grid and 
                    current_oe1_oe5_slot in schedule_grid[current_oe1_oe5_date] and
                    branch in schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot]):
                    schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot][branch] = None
                
                if best_oe1_oe5_date not in schedule_grid:
                    schedule_grid[best_oe1_oe5_date] = {}
                if best_oe1_oe5_slot not in schedule_grid[best_oe1_oe5_date]:
                    schedule_grid[best_oe1_oe5_date][best_oe1_oe5_slot] = {}
                schedule_grid[best_oe1_oe5_date][best_oe1_oe5_slot][branch] = subject
            
            if not oe2_data.empty:
                current_oe2_date = oe2_data['Exam Date'].iloc[0]
                current_oe2_slot = oe2_data['Time Slot'].iloc[0]
                
                for idx in oe2_data.index:
                    sem = all_data.at[idx, 'Semester']
                    branch = all_data.at[idx, 'Branch']
                    subject = all_data.at[idx, 'Subject']
                    
                    mask = (sem_dict[sem]['Subject'] == subject) & \
                           (sem_dict[sem]['Branch'] == branch)
                    sem_dict[sem].loc[mask, 'Exam Date'] = best_oe2_date
                    sem_dict[sem].loc[mask, 'Time Slot'] = best_oe2_slot
                    
                    if (current_oe2_date in schedule_grid and 
                        current_oe2_slot in schedule_grid[current_oe2_date] and
                        branch in schedule_grid[current_oe2_date][current_oe2_slot]):
                        schedule_grid[current_oe2_date][current_oe2_slot][branch] = None
                    
                    if best_oe2_date not in schedule_grid:
                        schedule_grid[best_oe2_date] = {}
                    if best_oe2_slot not in schedule_grid[best_oe2_date]:
                        schedule_grid[best_oe2_date][best_oe2_slot] = {}
                    schedule_grid[best_oe2_date][best_oe2_slot][branch] = subject
            
            moves_made += 1
            optimization_log.append(
                f"Moved OE1/OE5 from {current_oe1_oe5_date} to {best_oe1_oe5_date} (saved {days_saved} days)"
            )
            if not oe2_data.empty:
                optimization_log.append(
                    f"Moved OE2 to {best_oe2_date} (day immediately after OE1/OE5)"
                )
    
    if moves_made > 0:
        st.success(f"✅ OE Optimization: Moved {moves_made} OE groups with OE2 immediately after OE1/OE5!")
        with st.expander("📝 OE Optimization Details"):
            for log in optimization_log:
                st.write(f"• {log}")
    else:
        st.info("ℹ️ OE subjects are already optimally placed")
    
    return sem_dict

def reduce_exam_gaps(sem_dict, holidays, max_gap_days=1, time_slots=None):
    """Aggressively reduces gaps between exams by moving exams to fill empty slots."""
    if not sem_dict:
        return sem_dict, {}
    
    if time_slots is None:
        time_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
    
    import pandas as pd
    from datetime import datetime, timedelta
    from collections import defaultdict
    
    stats = {
        'gaps_found': 0,
        'gaps_reduced': 0,
        'exams_moved': 0,
        'days_saved': 0,
        'gap_details': [],
        'avg_gap_days': 0.0
    }
    
    all_data = pd.concat(sem_dict.values(), ignore_index=True)
    if all_data.empty:
        return sem_dict, stats
    
    all_data['Exam Date'] = pd.to_datetime(all_data['Exam Date'], format="%d-%m-%Y", errors='coerce')
    all_data = all_data.dropna(subset=['Exam Date'])
    
    schedule_grid = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: None)))
    branches = all_data['Branch'].unique()
    
    for _, row in all_data.iterrows():
        date_str = row['Exam Date'].strftime("%d-%m-%Y")
        schedule_grid[date_str][row['Time Slot']][row['Branch']] = row['Subject']
    
    all_exam_dates = sorted(all_data['Exam Date'].dt.date.unique())
    
    if len(all_exam_dates) < 2:
        return sem_dict, stats
    
    all_gaps = []
    total_gap_days = 0
    for i in range(len(all_exam_dates) - 1):
        current_date = all_exam_dates[i]
        next_date = all_exam_dates[i + 1]
        
        gap_days = 0
        check_date = current_date + timedelta(days=1)
        while check_date < next_date:
            if check_date.weekday() != 6 and check_date not in holidays:
                gap_days += 1
            check_date += timedelta(days=1)
        
        if gap_days > 0:
            all_gaps.append({
                'start_date': current_date,
                'end_date': next_date,
                'gap_days': gap_days,
                'priority': gap_days
            })
            total_gap_days += gap_days
            if gap_days > max_gap_days:
                stats['gaps_found'] += 1
    
    all_gaps.sort(key=lambda x: x['priority'], reverse=True)
    
    st.info(f"🔍 Found {len(all_gaps)} total gaps, {stats['gaps_found']} exceed {max_gap_days} days")
    
    for gap_info in all_gaps:
        start_date = gap_info['start_date']
        end_date = gap_info['end_date']
        gap_days = gap_info['gap_days']
        
        st.write(f"📅 Processing gap: {start_date} to {end_date} ({gap_days} working days)")
        
        target_date = start_date + timedelta(days=1)
        
        exams_after_gap = all_data[all_data['Exam Date'].dt.date > end_date].copy()
        exams_after_gap = exams_after_gap.sort_values('Exam Date')
        
        moves_in_this_gap = 0
        days_filled = []
        
        while target_date < end_date:
            if target_date.weekday() == 6 or target_date in holidays:
                target_date += timedelta(days=1)
                continue
                
            target_date_str = target_date.strftime("%d-%m-%Y")
            
            for _, exam_row in exams_after_gap.iterrows():
                if exam_row['Exam Date'].date() <= target_date:
                    continue
                    
                branch = exam_row['Branch']
                subject = exam_row['Subject']
                original_date = exam_row['Exam Date'].strftime("%d-%m-%Y")
                
                if pd.notna(exam_row.get('OE', None)) and exam_row['OE'] in ['OE1', 'OE5', 'OE2']:
                    continue
                
                branch_busy = False
                for slot in time_slots:
                    if schedule_grid[target_date_str][slot][branch] is not None:
                        branch_busy = True
                        break
                
                if not branch_busy:
                    available_slot = None
                    for slot in time_slots:
                        slot_free = True
                        module_code = exam_row['Subject'].split('(')[-1].split(')')[0] if '(' in exam_row['Subject'] else ''
                        if module_code:
                            related_exams = all_data[all_data['Subject'].str.contains(f'({module_code})', na=False)]
                            affected_branches = related_exams['Branch'].unique()
                        else:
                            affected_branches = [branch]
                        
                        for affected_branch in affected_branches:
                            if schedule_grid[target_date_str][slot][affected_branch] is not None:
                                slot_free = False
                                break
                        
                        if slot_free:
                            available_slot = slot
                            break
                    
                    if available_slot:
                        for affected_branch in affected_branches:
                            mask = (all_data['Branch'] == affected_branch) & \
                                   (all_data['Subject'] == exam_row['Subject']) & \
                                   (all_data['Exam Date'] == exam_row['Exam Date'])
                            
                            for idx in all_data[mask].index:
                                sem = all_data.at[idx, 'Semester']
                                old_slot = all_data.at[idx, 'Time Slot']
                                schedule_grid[original_date][old_slot][affected_branch] = None
                                schedule_grid[target_date_str][available_slot][affected_branch] = exam_row['Subject']
                                
                                sem_mask = (sem_dict[sem]['Subject'] == exam_row['Subject']) & \
                                          (sem_dict[sem]['Branch'] == affected_branch)
                                sem_dict[sem].loc[sem_mask, 'Exam Date'] = target_date_str
                                sem_dict[sem].loc[sem_mask, 'Time Slot'] = available_slot
                        
                        moves_in_this_gap += 1
                        stats['exams_moved'] += len(affected_branches)
                        days_saved = (exam_row['Exam Date'].date() - target_date).days
                        stats['days_saved'] += days_saved
                        
                        st.write(f"✅ Moved {exam_row['Subject']} from {original_date} to {target_date_str} (saved {days_saved} days)")
                        
                        days_filled.append(target_date)
                        break
            
            target_date += timedelta(days=1)
        
        if moves_in_this_gap > 0:
            stats['gaps_reduced'] += 1
            stats['gap_details'].append({
                'original_gap': gap_days,
                'moves_made': moves_in_this_gap,
                'days_filled': len(days_filled)
            })
    
    st.info("🎯 Compacting schedule further...")
    
    updated_all_data = pd.concat(sem_dict.values(), ignore_index=True)
    updated_all_data['Exam Date'] = pd.to_datetime(updated_all_data['Exam Date'], format="%d-%m-%Y", errors='coerce')
    updated_exam_dates = sorted(updated_all_data['Exam Date'].dt.date.unique())
    
    if len(updated_exam_dates) > 1:
        for late_date in reversed(updated_exam_dates[1:]):
            late_date_str = late_date.strftime("%d-%m-%Y")
            
            exams_on_late_date = updated_all_data[updated_all_data['Exam Date'].dt.date == late_date]
            
            for _, exam in exams_on_late_date.iterrows():
                if pd.notna(exam.get('OE', None)) and exam['OE'] in ['OE1', 'OE5', 'OE2']:
                    continue
                
                for early_date in updated_exam_dates:
                    if early_date >= late_date:
                        break
                    
                    early_date_str = early_date.strftime("%d-%m-%Y")
                    branch = exam['Branch']
                    
                    branch_free = True
                    for slot in time_slots:
                        if schedule_grid[early_date_str][slot][branch] is not None:
                            branch_free = False
                            break
                    
                    if branch_free:
                        for slot in time_slots:
                            if schedule_grid[early_date_str][slot][branch] is None:
                                old_slot = exam['Time Slot']
                                schedule_grid[late_date_str][old_slot][branch] = None
                                schedule_grid[early_date_str][slot][branch] = exam['Subject']
                                
                                sem = exam['Semester']
                                mask = (sem_dict[sem]['Subject'] == exam['Subject']) & \
                                       (sem_dict[sem]['Branch'] == branch)
                                sem_dict[sem].loc[mask, 'Exam Date'] = early_date_str
                                sem_dict[sem].loc[mask, 'Time Slot'] = slot
                                
                                stats['exams_moved'] += 1
                                days_saved = (late_date - early_date).days
                                stats['days_saved'] += days_saved
                                
                                st.write(f"✅ Compacted {exam['Subject']} from {late_date_str} to {early_date_str} (saved {days_saved} days)")
                                break
    
    # Calculate average gap
    if len(all_gaps) > 0:
        stats['avg_gap_days'] = total_gap_days / len(all_gaps)
    
    return sem_dict, stats

def save_to_excel(semester_wise_timetable):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sem, df in semester_wise_timetable.items():
            if df.empty:
                continue
            sheet_name = f"{df['MainBranch'].iloc[0]}_Sem_{sem}"
            df_pivot = df.pivot_table(
                index=["Exam Date", "Time Slot"],
                columns="Branch",
                values="Subject",
                aggfunc='first'
            ).reset_index()
            df_pivot.to_excel(writer, sheet_name=sheet_name, index=False)
            
            if 'OE' in df.columns and df['OE'].notna().any():
                elective_df = df[df['OE'].notna()]
                elective_df['SubjectDisplay'] = elective_df['Subject'] + " [" + elective_df['OE'] + "]"
                elective_pivot = elective_df.pivot_table(
                    index=["Exam Date", "Time Slot"],
                    columns="Branch",
                    values="SubjectDisplay",
                    aggfunc='first'
                ).reset_index()
                elective_sheet = f"{df['MainBranch'].iloc[0]}_Sem_{sem}_Electives"
                elective_pivot.to_excel(writer, sheet_name=elective_sheet, index=False)
    output.seek(0)
    return output

# Streamlit UI
st.markdown("""
<div class="main-header">
    <h1>Exam Timetable Generator</h1>
    <p>Create optimized exam schedules with minimal gaps</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Configuration")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    start_date = st.date_input("Exam Start Date", value=datetime.now().date())
    max_gap_days = st.slider("Maximum Allowed Gap (days)", 0, 3, 1, help="Set the maximum allowable gap between consecutive exam days (0 = no gaps)")
    schedule_by_difficulty = st.checkbox("Schedule by Difficulty", help="Prioritize scheduling based on subject difficulty")

    st.markdown("### Holidays")
    with st.expander("Add Holidays"):
        holidays_input = st.text_area("Enter holiday dates (DD-MM-YYYY, one per line)")
        holidays = set()
        if holidays_input:
            for line in holidays_input.split('\n'):
                try:
                    holiday = datetime.strptime(line.strip(), "%d-%m-%Y").date()
                    holidays.add(holiday)
                except ValueError:
                    st.warning(f"Invalid date format: {line.strip()}")

if uploaded_file:
    df_non, df_ele, df_full = read_timetable(uploaded_file)
    if df_non is not None and df_ele is not None:
        st.success("File uploaded successfully!")
        
        base_date = datetime.combine(start_date, datetime.min.time())
        sem_dict = process_constraints_with_real_time_optimization(df_non, holidays, base_date, schedule_by_difficulty)
        sem_dict = optimize_oe_subjects_after_scheduling(sem_dict, holidays)
        sem_dict, gap_stats = reduce_exam_gaps(sem_dict, holidays, max_gap_days)
        
        if sem_dict:
            output_pdf = "exam_timetable.pdf"
            generate_pdf_timetable(sem_dict, output_pdf)
            
            st.markdown("### Optimization Statistics")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{gap_stats['gaps_found']}</h3>
                    <p>Gaps Found</p>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{gap_stats['gaps_reduced']}</h3>
                    <p>Gaps Reduced</p>
                </div>
                """, unsafe_allow_html=True)
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{gap_stats['exams_moved']}</h3>
                    <p>Exams Moved</p>
                </div>
                """, unsafe_allow_html=True)
            with col4:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{gap_stats['days_saved']}</h3>
                    <p>Days Saved</p>
                </div>
                """, unsafe_allow_html=True)
            
            st.write(f"Average gap between exams: {gap_stats['avg_gap_days']:.2f} days")
            
            with open(output_pdf, "rb") as file:
                st.download_button(
                    label="Download Timetable PDF",
                    data=file,
                    file_name="exam_timetable.pdf",
                    mime="application/pdf"
                )
            
            with st.expander("View Detailed Gap Reduction"):
                for detail in gap_stats['gap_details']:
                    st.write(f"Gap of {detail['original_gap']} days: {detail['moves_made']} exams moved, {detail['days_filled']} days filled")
        else:
            st.error("Failed to generate timetable.")
else:
    st.info("Please upload an Excel file to generate the timetable.")

st.markdown("""
<div class="footer">
    <p>Developed by [Your Name] | © 2023</p>
</div>
""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
