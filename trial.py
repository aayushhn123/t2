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

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    
    # Enable automatic page numbering with alias
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
        # Placeholder function: Replace with actual logic or data source
        # Example: B.Tech Semester 5 EXTC defaults to "10:00 AM - 1:00 PM"
        default_timings = {
            ("5", "EXTC"): "10:00 AM - 1:00 PM",
            ("6", "COMP"): "10:00 AM - 1:00 PM",  # Add for other semesters/branches
            # Add more as needed
        }
        return default_timings.get((semester, branch), "10:00 AM - 1:00 PM")  # Default fallback

    for sheet_name, pivot_df in df_dict.items():
        if pivot_df.empty:
            continue
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        semester_roman = semester if not semester.isdigit() else int_to_roman(int(semester))
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        # Handle normal subjects
        if not sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            fixed_cols = ["Exam Date", "Time Slot"]  # Include Time Slot as a fixed column
            sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols]
            exam_date_width = 60
            table_font_size = 12
            line_height = 10

            for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
                chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
                cols_to_print = fixed_cols[:1] + chunk  # Exclude Time Slot from display but use it
                chunk_df = pivot_df[fixed_cols + chunk].copy()
                mask = chunk_df[chunk].apply(lambda row: row.astype(str).str.strip() != "").any(axis=1)
                chunk_df = chunk_df[mask].reset_index(drop=True)
                if chunk_df.empty:
                    continue

                # Get the time slot for this chunk
                time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and not pivot_df['Time Slot'].empty else None

                # Convert Exam Date to desired format
                chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

                # Modify subjects to include time range if duration != 3 hours or if timing overrides default
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
                            # Get subject-specific time slot
                            subject_time_slot = chunk_df.at[idx, "Time Slot"] if pd.notna(chunk_df.at[idx, "Time Slot"]) else None
                            if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                                if subject_time_slot != default_time_slot:
                                    modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                                else:
                                    modified_subjects.append(base_subject)
                            else:
                                modified_subjects.append(base_subject)
                            # Add duration-based time range if different from 3 hours
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
                
                # Add page before printing the table
                pdf.add_page()
                # Add footer with page number to the new page
                footer_height = 25
                add_footer_with_page_number(pdf, footer_height)
                
                print_table_custom(pdf, chunk_df[cols_to_print], cols_to_print, col_widths, line_height=line_height, 
                                 header_content=header_content, branches=chunk, time_slot=time_slot)

        # Handle electives with updated table structure
        if sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            
            # Group by 'OE' and 'Exam Date' to handle multiple subjects per OE type
            elective_data = pivot_df.groupby(['OE', 'Exam Date', 'Time Slot']).agg({
                'SubjectDisplay': lambda x: ", ".join(x)
            }).reset_index()

            # Convert Exam Date to desired format
            elective_data["Exam Date"] = pd.to_datetime(elective_data["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

            # Clean 'SubjectDisplay' to remove [OE] from each subject
            elective_data['SubjectDisplay'] = elective_data.apply(
                lambda row: ", ".join([s.replace(f" [{row['OE']}]", "") for s in row['SubjectDisplay'].split(", ")]),
                axis=1
            )

            # Modify subjects for timing overrides
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
                    # Get subject-specific time slot
                    subject_time_slot = elective_data.at[idx, "Time Slot"] if pd.notna(elective_data.at[idx, "Time Slot"]) else None
                    if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                        if subject_time_slot != default_time_slot:
                            modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                        else:
                            modified_subjects.append(base_subject)
                    else:
                        modified_subjects.append(base_subject)
                    # Add duration-based time range if different from 3 hours
                    if duration != 3 and subject_time_slot and subject_time_slot.strip():
                        start_time = subject_time_slot.split(" - ")[0]
                        end_time = calculate_end_time(start_time, duration)
                        modified_subjects[-1] = f"{modified_subjects[-1]} ({start_time} to {end_time})"
                elective_data.at[idx, 'SubjectDisplay'] = ", ".join(modified_subjects)

            # Rename columns for clarity in the PDF
            elective_data = elective_data.rename(columns={'OE': 'OE Type', 'SubjectDisplay': 'Subjects'})

            # Set column widths for three columns
            exam_date_width = 60
            oe_width = 30
            subject_width = pdf.w - 2 * pdf.l_margin - exam_date_width - oe_width
            col_widths = [exam_date_width, oe_width, subject_width]
            cols_to_print = ['Exam Date', 'OE Type', 'Subjects']
            
            # Add page before printing the electives table
            pdf.add_page()
            # Add footer with page number to the new page
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
            "Is Common": "IsCommon"  # Added "Is Common" column
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
        df["Exam Duration"] = df["Exam Duration"].fillna(3).astype(float)  # Default to 3 hours if NaN
        df["StudentCount"] = df["StudentCount"].fillna(0).astype(int)
        df["IsCommon"] = df["IsCommon"].fillna("NO").str.strip().str.upper()  # Default to "NO" if NaN
        
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

from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
from collections import defaultdict

# Global constants
MAX_BRANCHES_PER_DAY = 40

# Existing helper functions
def find_gaps_in_schedule(exam_dates, max_gap=2):
    """
    Find gaps larger than max_gap days in a sorted list of dates.
    Returns list of tuples (start_date, end_date, gap_days)
    """
    if len(exam_dates) < 2:
        return []
    
    gaps = []
    sorted_dates = sorted(exam_dates)
    
    for i in range(1, len(sorted_dates)):
        gap_days = (sorted_dates[i] - sorted_dates[i-1]).days
        if gap_days > max_gap:
            gaps.append((sorted_dates[i-1], sorted_dates[i], gap_days))
    
    return gaps

def find_best_day_for_gap_filling(start_date, end_date, branch_exam_days, holidays):
    """
    Find the best day within a gap to schedule an exam.
    Prefers middle of the gap to distribute exams evenly.
    """
    gap_days = (end_date - start_date).days
    middle_offset = gap_days // 2
    candidate_date = start_date + timedelta(days=middle_offset)
    
    for offset in range(-middle_offset + 1, gap_days):
        check_date = start_date + timedelta(days=middle_offset + offset)
        if check_date >= end_date:
            continue
        if check_date.weekday() != 6 and check_date not in holidays and check_date not in branch_exam_days:
            return check_date
    return None

# Enhanced scheduling function with aggressive optimization
def schedule_with_aggressive_gap_optimization(df, holidays, base_date, schedule_by_difficulty=False, target_span=20):
    """
    Aggressive scheduling optimization that ensures:
    1. No gaps > 2 days between exams
    2. Maximum compression to achieve < 20 day span
    3. Common subject consolidation
    """
    # Initialize tracking structures
    all_branches = df['Branch'].unique()
    exam_days = {branch: set() for branch in all_branches}
    daily_schedule = defaultdict(list)  # Single list per day
    
    # Add scheduled flag and initialize columns
    df['Scheduled'] = False
    df['Exam Date'] = ""
    
    def get_next_valid_date(current_date):
        """Get next valid exam date (not Sunday or holiday)"""
        while current_date.weekday() == 6 or current_date.date() in holidays:
            current_date += timedelta(days=1)
        return current_date
    
    def can_schedule_in_day(date, branches_to_check):
        """Check if branches can be scheduled in given day"""
        current_branches = daily_schedule[date]
        if len(current_branches) + len(branches_to_check) > MAX_BRANCHES_PER_DAY:
            return False
        for branch in branches_to_check:
            if branch in current_branches:
                return False
        return True
    
    def schedule_exam(exam_indices, date, branches):
        """Schedule exam for given indices"""
        date_str = date.strftime("%d-%m-%Y")
        
        for idx in exam_indices:
            df.at[idx, 'Exam Date'] = date_str
            df.at[idx, 'Scheduled'] = True
        
        daily_schedule[date].extend(branches)
        for branch in branches:
            exam_days[branch].add(date.date())  # Store as datetime.date
    
    # Phase 1: Common Subject Scheduling
    st.info("Phase 1: Scheduling common subjects...")
    common_subjects = df[df['IsCommon'] == 'YES']
    current_date = get_next_valid_date(base_date)
    
    for module_code, group in common_subjects.groupby('ModuleCode'):
        branches = group['Branch'].unique().tolist()
        scheduled = False
        for date in sorted(daily_schedule.keys()):
            if can_schedule_in_day(date, branches):
                schedule_exam(group.index.tolist(), date, branches)
                scheduled = True
                break
        
        if not scheduled:
            while not can_schedule_in_day(current_date, branches):
                current_date = get_next_valid_date(current_date + timedelta(days=1))
            schedule_exam(group.index.tolist(), current_date, branches)
    
    # Phase 2: Branch-wise Non-Common Subject Scheduling
    st.info("Phase 2: Packing non-common subjects with gap control...")
    non_common_by_branch = defaultdict(list)
    non_common = df[(df['IsCommon'] == 'NO') & (~df['Scheduled'])]
    
    for idx, row in non_common.iterrows():
        non_common_by_branch[row['Branch']].append({
            'idx': idx,
            'semester': row['Semester'],
            'module': row['ModuleCode'],
            'category': row['Category']
        })
    
    sorted_branches = sorted(non_common_by_branch.keys(), 
                           key=lambda x: len(non_common_by_branch[x]), 
                           reverse=True)
    
    for branch in sorted_branches:
        branch_exams = non_common_by_branch[branch]
        branch_exams.sort(key=lambda x: x['semester'])
        
        for exam in branch_exams:
            best_date = None
            min_gap = float('inf')
            
            if exam_days[branch]:
                last_exam_date = max(exam_days[branch])  # datetime.date
                check_start = min(exam_days[branch])     # datetime.date
                check_end = last_exam_date + timedelta(days=3)
                
                current_check = datetime.combine(check_start, datetime.min.time())
                while current_check.date() <= check_end:
                    if (current_check.date() not in exam_days[branch] and 
                        current_check.weekday() != 6 and 
                        current_check.date() not in holidays):
                        gaps_before = [abs((current_check.date() - d).days) for d in exam_days[branch]]
                        max_gap_before = max(gaps_before) if gaps_before else 0
                        if can_schedule_in_day(current_check, [branch]):
                            if max_gap_before < min_gap:
                                min_gap = max_gap_before
                                best_date = current_check
                    
                    current_check += timedelta(days=1)
            
            if best_date is None:
                if exam_days[branch]:
                    last_date = datetime.combine(max(exam_days[branch]), datetime.min.time())
                    search_date = last_date + timedelta(days=1)
                    days_checked = 0
                    while days_checked < 3:
                        search_date = get_next_valid_date(search_date)
                        if can_schedule_in_day(search_date, [branch]):
                            best_date = search_date
                            break
                        search_date += timedelta(days=1)
                        days_checked += 1
                    if best_date is None:
                        best_date = search_date
                        while not can_schedule_in_day(best_date, [branch]):
                            best_date = get_next_valid_date(best_date + timedelta(days=1))
                else:
                    best_date = get_next_valid_date(base_date)
                    while not can_schedule_in_day(best_date, [branch]):
                        best_date = get_next_valid_date(best_date + timedelta(days=1))
            
            schedule_exam([exam['idx']], best_date, [branch])
    
    # Phase 3: Final Compression
    st.info("Phase 3: Final compression pass...")
    dates_to_check = sorted(daily_schedule.keys())
    for i, date in enumerate(dates_to_check):
        current_count = len(daily_schedule[date])
        if current_count < 5 and i > 0:
            for j in range(i-1, max(0, i-3), -1):
                target_date = dates_to_check[j]
                if can_schedule_in_day(target_date, daily_schedule[date]):
                    branches_to_move = daily_schedule[date]
                    mask = (df['Exam Date'] == date.strftime("%d-%m-%Y"))
                    df.loc[mask, 'Exam Date'] = target_date.strftime("%d-%m-%Y")
                    daily_schedule[target_date].extend(branches_to_move)
                    daily_schedule[date] = []
                    for branch in branches_to_move:
                        exam_days[branch].remove(date)
                        exam_days[branch].add(target_date.date())
                    break
    
    # Drop temporary 'Scheduled' column
    df = df.drop('Scheduled', axis=1)
    
    # Final validation and reporting
    all_dates = pd.to_datetime(df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    if not all_dates.empty:
        start_date = min(all_dates)
        end_date = max(all_dates)
        total_span = (end_date - start_date).days + 1
        
        max_gap_found = 0
        gap_violations = []
        for branch in all_branches:
            sorted_dates = sorted(exam_days[branch])
            for i in range(1, len(sorted_dates)):
                gap = (sorted_dates[i] - sorted_dates[i-1]).days
                if gap > 2:
                    gap_violations.append(f"{branch}: {gap}-day gap")
                max_gap_found = max(max_gap_found, gap)
        
        st.write("### Scheduling Results")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Span", f"{total_span} days", 
                     delta=f"{total_span - target_span} days" if total_span > target_span else "✓")
        with col2:
            st.metric("Max Gap", f"{max_gap_found} days",
                     delta="✓" if max_gap_found <= 2 else f"{max_gap_found - 2} over")
        with col3:
            st.metric("Efficiency", f"{len(all_dates)} exam days")
        
        if total_span <= 16:
            st.success(f"✅ Excellent! Timetable optimized to {total_span} days (within 16-day target)")
        elif total_span <= target_span:
            st.success(f"✅ Success! Timetable span: {total_span} days (within {target_span}-day target)")
        else:
            st.warning(f"⚠️ Timetable spans {total_span} days. Further optimization needed.")
        
        if gap_violations:
            st.error(f"❌ Found {len(gap_violations)} gaps exceeding 2 days:")
            for violation in gap_violations[:5]:
                st.write(f"  - {violation}")
        else:
            st.success("✅ All gaps are within the 2-day limit!")
    
    # Organize by semester
    sem_dict = {}
    for sem in sorted(df["Semester"].unique()):
        if sem == 0:
            continue
        sem_dict[sem] = df[df["Semester"] == sem].copy()
    
    return sem_dict

# Updated process_constraints with optimization wrapper
def process_constraints_with_optimization(df, holidays, base_date, schedule_by_difficulty=False, target_span=20):
    """
    Enhanced main function that integrates all optimization strategies.
    """
    st.header("🎯 Exam Scheduling with Optimization")
    
    # Show initial statistics
    st.write("### Input Data Statistics")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Subjects", len(df))
    with col2:
        common_count = len(df[df['IsCommon'] == 'YES']) if 'IsCommon' in df.columns else 0
        st.metric("Common Subjects", common_count)
    with col3:
        branch_count = df['Branch'].nunique()
        st.metric("Branches", branch_count)
    
    # Step 1: Run the aggressive optimization
    st.write("### Step 1: Aggressive Gap-Controlled Scheduling")
    with st.spinner("Scheduling in progress..."):
        sem_dict = schedule_with_aggressive_gap_optimization(df, holidays, base_date, schedule_by_difficulty, target_span)
    
    # Step 2: Validate the initial schedule
    st.write("### Step 2: Initial Validation")
    validation_passed = validate_schedule_constraints(sem_dict)
    
    # Step 3: Post-process if needed
    if not validation_passed or needs_further_optimization(sem_dict, target_span):
        st.write("### Step 3: Post-Processing Optimization")
        with st.spinner("Applying post-processing optimization..."):
            sem_dict = post_process_schedule_optimization(sem_dict, holidays, target_span)
        
        # Re-validate
        st.write("### Step 4: Final Validation")
        validate_schedule_constraints(sem_dict)
    
    # Generate summary report
    generate_optimization_report(sem_dict, target_span)
    
    return sem_dict

# Helper functions from the guide
def needs_further_optimization(sem_dict, target_span):
    """Check if schedule needs further optimization"""
    all_dates = []
    for sem, df in sem_dict.items():
        dates = pd.to_datetime(df['Exam Date'], format="%d-%m-%Y", errors='coerce')
        all_dates.extend(dates.dropna())
    
    if all_dates:
        span = (max(all_dates) - min(all_dates)).days + 1
        return span > target_span
    return False

def generate_optimization_report(sem_dict, target_span):
    """Generate a comprehensive optimization report"""
    st.write("### 📊 Optimization Report")
    
    all_exams = []
    for sem, df in sem_dict.items():
        all_exams.append(df)
    
    if not all_exams:
        st.warning("No exams scheduled")
        return
    
    combined_df = pd.concat(all_exams, ignore_index=True)
    combined_df['ParsedDate'] = pd.to_datetime(combined_df['Exam Date'], format="%d-%m-%Y")
    
    total_span = (combined_df['ParsedDate'].max() - combined_df['ParsedDate'].min()).days + 1
    unique_dates = combined_df['ParsedDate'].nunique()
    total_exams = len(combined_df)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        delta_span = total_span - target_span
        st.metric("Total Span", f"{total_span} days", 
                 delta=f"{delta_span} days" if delta_span > 0 else "✓ On target")
    with col2:
        efficiency = (total_exams / (unique_dates * MAX_BRANCHES_PER_DAY)) * 100
        st.metric("Slot Efficiency", f"{efficiency:.1f}%")
    with col3:
        avg_gap = calculate_average_gap(combined_df)
        st.metric("Avg Gap", f"{avg_gap:.1f} days")
    with col4:
        compression = ((46 - total_span) / 46) * 100
        st.metric("Compression", f"{compression:.1f}%")
    
    if total_span <= target_span:
        st.success(f"🎉 Successfully optimized schedule to {total_span} days!")
        st.balloons()
    else:
        st.warning(f"⚠️ Schedule spans {total_span} days. Manual intervention may be needed.")
        suggestions = generate_optimization_suggestions(combined_df, target_span)
        st.write("### 💡 Suggestions for Further Optimization")
        for i, suggestion in enumerate(suggestions, 1):
            st.write(f"{i}. {suggestion}")
    
    visualize_schedule(combined_df)

def calculate_average_gap(df):
    """Calculate average gap between exams across all branches"""
    total_gaps = 0
    gap_count = 0
    
    for branch in df['Branch'].unique():
        branch_dates = sorted(df[df['Branch'] == branch]['ParsedDate'].unique())
        for i in range(1, len(branch_dates)):
            gap = (branch_dates[i] - branch_dates[i-1]).days
            total_gaps += gap
            gap_count += 1
    
    return total_gaps / gap_count if gap_count > 0 else 0

def generate_optimization_suggestions(df, target_span):
    """Generate specific suggestions for further optimization"""
    suggestions = []
    date_counts = df.groupby('ParsedDate').size()
    sparse_days = date_counts[date_counts < 10]
    if len(sparse_days) > 0:
        suggestions.append(f"Consolidate {len(sparse_days)} sparse days with fewer than 10 exams")
    
    if df['ParsedDate'].nunique() > target_span - 4:
        suggestions.append("Consider using Saturdays for make-up exams if permitted")
    
    last_week = df['ParsedDate'].max() - timedelta(days=7)
    last_week_exams = df[df['ParsedDate'] > last_week]
    if len(last_week_exams) < 30:
        suggestions.append(f"Move {len(last_week_exams)} exams from the final week to earlier gaps")
    
    return suggestions

def visualize_schedule(df):
    """Create a visual representation of the schedule"""
    daily_counts = df.groupby('ParsedDate').size().reset_index(name='Count')
    daily_counts['Day'] = daily_counts['ParsedDate'].dt.strftime('%d-%b')
    st.bar_chart(daily_counts.set_index('Day')['Count'])
    
    st.write("#### Branch Schedule Density")
    branch_date_matrix = df.pivot_table(
        index='Branch', 
        columns='ParsedDate', 
        values='ModuleCode', 
        aggfunc='count',
        fill_value=0
    )
    if len(branch_date_matrix) > 10:
        st.write("(Showing first 10 branches)")
        branch_date_matrix = branch_date_matrix.iloc[:10]
    st.dataframe(branch_date_matrix, use_container_width=True)  # Removed background_gradient

def post_process_schedule_optimization(sem_dict, holidays, max_span_days=20):
    """
    Post-processing function to further optimize an already generated schedule.
    """
    all_exams = []
    for sem, df in sem_dict.items():
        all_exams.append(df)
    
    if not all_exams:
        return sem_dict
    
    combined_df = pd.concat(all_exams, ignore_index=True)
    combined_df['ParsedDate'] = pd.to_datetime(combined_df['Exam Date'], format="%d-%m-%Y")
    current_span = (combined_df['ParsedDate'].max() - combined_df['ParsedDate'].min()).days + 1
    
    if current_span <= max_span_days:
        st.success(f"✅ Schedule already optimized! Current span: {current_span} days")
        return sem_dict
    
    st.warning(f"📊 Current span: {current_span} days. Attempting further optimization...")
    
    date_counts = combined_df.groupby('ParsedDate').size()
    sparse_dates = date_counts[date_counts < 10].index
    
    if len(sparse_dates) > 0:
        st.info(f"Found {len(sparse_dates)} sparse days with < 10 exams. Consolidating...")
        for sparse_date in sparse_dates:
            sparse_exams = combined_df[combined_df['ParsedDate'] == sparse_date]
            for days_offset in [1, -1, 2, -2]:
                target_date = sparse_date + timedelta(days=days_offset)
                if target_date.weekday() == 6 or target_date.date() in holidays:
                    continue
                target_exams = combined_df[combined_df['ParsedDate'] == target_date]
                if len(target_exams) > 0 and len(target_exams) < MAX_BRANCHES_PER_DAY:
                    combined_df.loc[sparse_exams.index, 'Exam Date'] = target_date.strftime("%d-%m-%Y")
                    combined_df.loc[sparse_exams.index, 'ParsedDate'] = target_date
                    break
    
    last_week_start = combined_df['ParsedDate'].max() - timedelta(days=7)
    last_week_exams = combined_df[combined_df['ParsedDate'] >= last_week_start]
    
    if len(last_week_exams) < 50:
        st.info("Compressing sparse final week...")
        all_dates = pd.date_range(combined_df['ParsedDate'].min(), 
                                 combined_df['ParsedDate'].max(), 
                                 freq='D')
        for exam_idx in last_week_exams.index:
            exam = combined_df.loc[exam_idx]
            branch = exam['Branch']
            branch_dates = combined_df[combined_df['Branch'] == branch]['ParsedDate'].unique()
            for i in range(len(all_dates) - 7):
                check_date = all_dates[i]
                if check_date.weekday() == 6 or check_date.date() in holidays:
                    continue
                if check_date not in branch_dates:
                    nearby_dates = [d for d in branch_dates if abs((check_date - d).days) <= 2]
                    if nearby_dates:
                        combined_df.loc[exam_idx, 'Exam Date'] = check_date.strftime("%d-%m-%Y")
                        combined_df.loc[exam_idx, 'ParsedDate'] = check_date
                        break
    
    new_span = (combined_df['ParsedDate'].max() - combined_df['ParsedDate'].min()).days + 1
    if new_span < current_span:
        st.success(f"✅ Optimization successful! Reduced from {current_span} to {new_span} days")
    else:
        st.info("ℹ️ No further optimization possible without violating constraints")
    
    optimized_sem_dict = {}
    for sem in combined_df['Semester'].unique():
        if sem != 0:
            sem_df = combined_df[combined_df['Semester'] == sem].copy()
            sem_df = sem_df.drop('ParsedDate', axis=1)
            optimized_sem_dict[sem] = sem_df
    
    return optimized_sem_dict

def validate_schedule_constraints(sem_dict):
    """
    Validate that the schedule meets all constraints:
    1. No gaps > 2 days for any branch
    2. Common subjects scheduled on same day
    3. No Sunday exams
    """
    all_exams = []
    for sem, df in sem_dict.items():
        all_exams.append(df)
    
    if not all_exams:
        return True
    
    combined_df = pd.concat(all_exams, ignore_index=True)
    combined_df['ParsedDate'] = pd.to_datetime(combined_df['Exam Date'], format="%d-%m-%Y")
    
    validation_passed = True
    
    st.write("### Constraint Validation")
    
    gap_violations = []
    for branch in combined_df['Branch'].unique():
        branch_dates = sorted(combined_df[combined_df['Branch'] == branch]['ParsedDate'].unique())
        for i in range(1, len(branch_dates)):
            gap = (branch_dates[i] - branch_dates[i-1]).days
            if gap > 2:
                gap_violations.append({
                    'branch': branch,
                    'from': branch_dates[i-1].strftime("%d-%m-%Y"),
                    'to': branch_dates[i].strftime("%d-%m-%Y"),
                    'gap': gap
                })
    
    if gap_violations:
        st.error(f"❌ Gap constraint violated: {len(gap_violations)} violations found")
        for v in gap_violations[:5]:
            st.write(f"  - {v['branch']}: {v['gap']}-day gap ({v['from']} to {v['to']})")
        validation_passed = False
    else:
        st.success("✅ Gap constraint: All gaps ≤ 2 days")
    
    common_subjects = combined_df[combined_df['IsCommon'] == 'YES'] if 'IsCommon' in combined_df.columns else pd.DataFrame()
    if not common_subjects.empty:
        module_dates = common_subjects.groupby('ModuleCode')['ParsedDate'].unique()
        multi_date_modules = [mod for mod, dates in module_dates.items() if len(dates) > 1]
        if multi_date_modules:
            st.error(f"❌ Common subject constraint violated: {len(multi_date_modules)} subjects on multiple days")
            validation_passed = False
        else:
            st.success("✅ Common subject constraint: All on single days")
    
    sunday_exams = combined_df[combined_df['ParsedDate'].dt.weekday == 6]
    if not sunday_exams.empty:
        st.error(f"❌ Sunday constraint violated: {len(sunday_exams)} exams on Sundays")
        validation_passed = False
    else:
        st.success("✅ Sunday constraint: No exams on Sundays")
    
    st.write("### Schedule Statistics")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        total_span = (combined_df['ParsedDate'].max() - combined_df['ParsedDate'].min()).days + 1
        st.metric("Total Span", f"{total_span} days")
    with col2:
        unique_dates = combined_df['ParsedDate'].nunique()
        st.metric("Exam Days", unique_dates)
    with col3:
        avg_per_day = len(combined_df) / unique_dates
        st.metric("Avg Exams/Day", f"{avg_per_day:.1f}")
    with col4:
        st.metric("Total Exams", len(combined_df))
    
    return validation_passed

# Backward compatible wrapper
def process_constraints(df, holidays, base_date, schedule_by_difficulty=False):
    """
    Backward compatible wrapper for the optimized scheduling function.
    """
    return process_constraints_with_optimization(
        df, holidays, base_date, 
        schedule_by_difficulty=schedule_by_difficulty,
        target_span=20
    )

# Integration with Streamlit app (maintained as provided)
def test_scheduling_optimization():
    """
    Example function showing how to use the optimized scheduling system
    """
    # [Existing implementation remains unchanged]
    pass

def quick_schedule_test():
    """
    Quick test function for the scheduling algorithm
    """
    # [Existing implementation remains unchanged]
    pass



    
def save_to_excel(semester_wise_timetable):
    if not semester_wise_timetable:
        return None

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

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sem, df_sem in semester_wise_timetable.items():
            for main_branch in df_sem["MainBranch"].unique():
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
                # Separate non-electives and electives
                df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()

                # Process non-electives
                if not df_non_elec.empty:
                    difficulty_str = df_non_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                    difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                    df_non_elec["SubjectDisplay"] = df_non_elec["Subject"]
                    duration_suffix = df_non_elec.apply(
                        lambda row: f" [Duration: {row['Exam Duration']} hrs]" if row['Exam Duration'] != 3 else '', axis=1)
                    df_non_elec["SubjectDisplay"] = df_non_elec["SubjectDisplay"] + difficulty_suffix + duration_suffix
                    df_non_elec["Exam Date"] = pd.to_datetime(
                        df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce'
                    )
                    df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                    pivot_df = df_non_elec.pivot_table(
                        index=["Exam Date", "Time Slot"],
                        columns="SubBranch",
                        values="SubjectDisplay",
                        aggfunc=lambda x: ", ".join(str(i) for i in x)
                    ).fillna("---")
                    pivot_df = pivot_df.sort_index(level="Exam Date", ascending=True)
                    formatted_dates = [d.strftime("%d-%m-%Y") for d in pivot_df.index.levels[0]]
                    pivot_df.index = pivot_df.index.set_levels(formatted_dates, level=0)
                    roman_sem = int_to_roman(sem)
                    sheet_name = f"{main_branch}_Sem_{roman_sem}"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    pivot_df.to_excel(writer, sheet_name=sheet_name)

                # Process electives in a separate sheet
                if not df_elec.empty:
                    difficulty_str = df_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                    difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                    df_elec["SubjectDisplay"] = df_elec["Subject"] + " [" + df_elec["OE"] + "]"
                    duration_suffix = df_elec.apply(
                        lambda row: f" [Duration: {row['Exam Duration']} hrs]" if row['Exam Duration'] != 3 else '', axis=1)
                    df_elec["SubjectDisplay"] = df_elec["SubjectDisplay"] + difficulty_suffix + duration_suffix
                    elec_pivot = df_elec.groupby(['OE', 'Exam Date', 'Time Slot'])['SubjectDisplay'].apply(
                        lambda x: ", ".join(sorted(set(x)))
                    ).reset_index()
                    elec_pivot['Exam Date'] = pd.to_datetime(
                        elec_pivot['Exam Date'], format="%d-%m-%Y", errors='coerce'
                    ).dt.strftime("%d-%m-%Y")
                    elec_pivot = elec_pivot.sort_values(by="Exam Date", ascending=True)
                    roman_sem = int_to_roman(sem)
                    sheet_name = f"{main_branch}_Sem_{roman_sem}_Electives"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    elec_pivot.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output

def save_verification_excel(original_df, semester_wise_timetable):
    if not semester_wise_timetable:
        return None

    columns_to_retain = [
        "School Name", "Campus", "Program", "Stream", "Current Academic Year",
        "Semester", "ModuleCode", "SubjectName", "Difficulty", "Category", "OE", "Exam mode", "Exam Duration"
    ]

    available_columns = [col for col in columns_to_retain if col in original_df.columns]
    verification_df = original_df[available_columns].copy()

    verification_df["Exam Date"] = ""
    verification_df["Exam Time"] = ""
    verification_df["Is Common"] = ""

    scheduled_data = pd.concat(semester_wise_timetable.values(), ignore_index=True)
    scheduled_data["ModuleCode"] = scheduled_data["Subject"].str.extract(r'\((.*?)\)', expand=False)

    for idx, row in verification_df.iterrows():
        module_code = row["ModuleCode"]
        semester = row["Semester"]

        match = scheduled_data[
            (scheduled_data["ModuleCode"] == module_code) &
            (scheduled_data["Semester"] == semester)
        ]

        if not match.empty:
            exam_date = match.iloc[0]["Exam Date"]
            time_slot = match.iloc[0]["Time Slot"]
            duration = row["Exam Duration"]
            start_time = time_slot.split(" - ")[0]
            end_time = calculate_end_time(start_time, duration)
            exam_time = f"{start_time} to {end_time}"
            verification_df.at[idx, "Exam Date"] = exam_date
            verification_df.at[idx, "Exam Time"] = exam_time
            branch_count = len(scheduled_data[scheduled_data["ModuleCode"] == module_code]["Branch"].unique())
            verification_df.at[idx, "Is Common"] = "YES" if branch_count > 1 else "NO"

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        verification_df.to_excel(writer, sheet_name="Verification", index=False)

    output.seek(0)
    return output

def main():
    st.markdown("""
    <div class="main-header">
        <h1>📅 Exam Timetable Generator</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    # Initialize session state variables if not present
    if 'num_custom_holidays' not in st.session_state:
        st.session_state.num_custom_holidays = 1
    if 'custom_holidays' not in st.session_state:
        st.session_state.custom_holidays = [None]
    if 'timetable_data' not in st.session_state:
        st.session_state.timetable_data = {}
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    if 'pdf_data' not in st.session_state:
        st.session_state.pdf_data = None
    if 'verification_data' not in st.session_state:
        st.session_state.verification_data = None
    if 'total_exams' not in st.session_state:
        st.session_state.total_exams = 0
    if 'total_semesters' not in st.session_state:
        st.session_state.total_semesters = 0
    if 'total_branches' not in st.session_state:
        st.session_state.total_branches = 0
    if 'overall_date_range' not in st.session_state:
        st.session_state.overall_date_range = 0
    if 'unique_exam_days' not in st.session_state:
        st.session_state.unique_exam_days = 0
    if 'non_elective_range' not in st.session_state:
        st.session_state.non_elective_range = "N/A"
    if 'elective_dates_str' not in st.session_state:
        st.session_state.elective_dates_str = "N/A"
    if 'stream_counts' not in st.session_state:
        st.session_state.stream_counts = pd.DataFrame()

    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        st.markdown("#### 📅 Base Date for Scheduling")
        base_date = st.date_input("Start date for exams", value=datetime(2025, 4, 1))
        base_date = datetime.combine(base_date, datetime.min.time())

        st.markdown("#### 🛠️ Scheduling Mode")
        schedule_by_difficulty = st.checkbox("Schedule by Difficulty (Alternate Easy/Difficult)", value=False)
        if schedule_by_difficulty:
            st.markdown('<div class="status-info">ℹ️ Exams will alternate between Easy and Difficult subjects.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-info">ℹ️ Normal scheduling without considering difficulty.</div>',
                        unsafe_allow_html=True)

        st.markdown('<div style="margin-top: 2rem;"></div>', unsafe_allow_html=True)

        with st.expander("Holiday Configuration", expanded=True):
            st.markdown("#### 📅 Select Predefined Holidays")
            holiday_dates = []

            col1, col2 = st.columns(2)
            with col1:
                if st.checkbox("April 14, 2025", value=True):
                    holiday_dates.append(datetime(2025, 4, 14))
            with col2:
                if st.checkbox("May 1, 2025", value=True):
                    holiday_dates.append(datetime(2025, 5, 1))

            if st.checkbox("August 15, 2025", value=True):
                holiday_dates.append(datetime(2025, 8, 15))

            st.markdown("#### 📅 Add Custom Holidays")
            if len(st.session_state.custom_holidays) < st.session_state.num_custom_holidays:
                st.session_state.custom_holidays.extend(
                    [None] * (st.session_state.num_custom_holidays - len(st.session_state.custom_holidays))
                )

            for i in range(st.session_state.num_custom_holidays):
                st.session_state.custom_holidays[i] = st.date_input(
                    f"Custom Holiday {i + 1}",
                    value=st.session_state.custom_holidays[i],
                    key=f"custom_holiday_{i}"
                )

            if st.button("➕ Add Another Holiday"):
                st.session_state.num_custom_holidays += 1
                st.session_state.custom_holidays.append(None)
                st.rerun()

            custom_holidays = [h for h in st.session_state.custom_holidays if h is not None]
            for custom_holiday in custom_holidays:
                holiday_dates.append(datetime.combine(custom_holiday, datetime.min.time()))

            if holiday_dates:
                st.markdown("#### Selected Holidays:")
                for holiday in sorted(holiday_dates):
                    st.write(f"• {holiday.strftime('%B %d, %Y')}")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("""
        <div class="upload-section">
            <h3>📁 Upload Excel File</h3>
            <p>Upload your timetable data file (.xlsx format)</p>
        </div>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload the Excel file containing your timetable data"
        )

        if uploaded_file is not None:
            st.markdown('<div class="status-success">✅ File uploaded successfully!</div>', unsafe_allow_html=True)

            file_details = {
                "Filename": uploaded_file.name,
                "File size": f"{uploaded_file.size / 1024:.2f} KB",
                "File type": uploaded_file.type
            }

            st.markdown("#### File Details:")
            for key, value in file_details.items():
                st.write(f"**{key}:** {value}")

    with col2:
        st.markdown("""
        <div class="feature-card">
            <h4>🚀 Features</h4>
            <ul>
                <li>📊 Excel file processing</li>
                <li>📅 Smart scheduling</li>
                <li>📋 Multiple output formats</li>
                <li>🎯 Conflict resolution</li>
                <li>📱 Mobile-friendly interface</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    if uploaded_file is not None:
        if st.button("🔄 Generate Timetable", type="primary", use_container_width=True):
            with st.spinner("Processing your timetable... Please wait..."):
                try:
                    holidays_set = set(holiday_dates)
                    st.write("Reading timetable...")
                    df_non_elec, df_ele, original_df = read_timetable(uploaded_file)
                    st.write(f"df_non_elec shape: {df_non_elec.shape if df_non_elec is not None else 'None'}")
                    st.write(f"df_ele shape: {df_ele.shape if df_ele is not None else 'None'}")

                    if df_non_elec is not None and df_ele is not None:
                        st.write("Processing constraints...")
                        non_elec_sched = process_constraints(df_non_elec, holidays_set, base_date, schedule_by_difficulty)
                        st.write(f"non_elec_sched keys: {list(non_elec_sched.keys())}")

                        # Find the maximum date from non-elective exams
                        non_elec_df = pd.concat(non_elec_sched.values(), ignore_index=True) if non_elec_sched else pd.DataFrame()
                        non_elec_dates = pd.to_datetime(non_elec_df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        max_non_elec_date = max(non_elec_dates).date() if not non_elec_dates.empty else base_date.date()
                        st.write(f"Max non-elective date: {max_non_elec_date}")

                        # Define function to find next valid day for electives
                        def find_next_valid_day(start_day):
                            day = start_day
                            while True:
                                day_date = day.date()
                                if day.weekday() == 6 or day_date in holidays_set:
                                    day += timedelta(days=1)
                                    continue
                                return day

                        # Schedule electives globally only if df_ele is not None
                        if df_ele is not None and not df_ele.empty:
                            st.write("Scheduling electives...")
                            elective_day1 = find_next_valid_day(datetime.combine(max_non_elec_date, datetime.min.time()) + timedelta(days=1))
                            elective_day2 = find_next_valid_day(elective_day1 + timedelta(days=1))
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Exam Date'] = elective_day1.strftime("%d-%m-%Y")
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Time Slot'] = "10:00 AM - 1:00 PM"
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Exam Date'] = elective_day2.strftime("%d-%m-%Y")
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Time Slot'] = "2:00 PM - 5:00 PM"

                            # Combine non-electives and electives
                            final_df = pd.concat([non_elec_df, df_ele], ignore_index=True)
                        else:
                            final_df = non_elec_df
                            st.write("No electives to schedule.")

                        final_df["Exam Date"] = pd.to_datetime(final_df["Exam Date"], format="%d-%m-%Y", errors='coerce')
                        final_df = final_df.sort_values(["Exam Date", "Semester", "MainBranch"], ascending=True, na_position='last')
                        sem_dict = {s: final_df[final_df["Semester"] == s].copy() for s in sorted(final_df["Semester"].unique())}
                        st.write(f"Semesters in sem_dict: {list(sem_dict.keys())}")

                        st.session_state.timetable_data = sem_dict
                        st.session_state.original_df = original_df
                        st.session_state.processing_complete = True

                        # Compute statistics
                        total_exams = sum(len(df) for df in sem_dict.values())
                        total_semesters = len(sem_dict)
                        total_branches = len(set(branch for df in sem_dict.values() for branch in df['MainBranch'].unique()))

                        all_data = pd.concat(sem_dict.values(), ignore_index=True)
                        all_dates = pd.to_datetime(all_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        overall_date_range = (max(all_dates) - min(all_dates)).days + 1 if all_dates.size > 0 else 0
                        unique_exam_days = len(all_dates.dt.date.unique())

                        non_elective_data = all_data[all_data['OE'].isna() | (all_data['OE'].str.strip() == "")]
                        non_elective_dates = pd.to_datetime(non_elective_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        non_elective_range = f"{min(non_elective_dates).strftime('%d %b %Y')} to {max(non_elective_dates).strftime('%d %b %Y')}" if non_elective_dates.size > 0 else "N/A"

                        elective_data = all_data[all_data['OE'].notna() & (all_data['OE'].str.strip() != "")]
                        elective_dates = pd.to_datetime(elective_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        elective_dates_str = ", ".join(sorted(set(elective_dates.dt.strftime("%d %b %Y")))) if elective_dates.size > 0 else "N/A"

                        non_oe_data = all_data[all_data['OE'].isna() | (all_data['OE'].str.strip() == "")]
                        stream_counts = non_oe_data.groupby(['MainBranch', 'SubBranch'])['Subject'].count().reset_index()
                        stream_counts['Stream'] = stream_counts['MainBranch'] + " " + stream_counts['SubBranch']
                        stream_counts = stream_counts[['Stream', 'Subject']].rename(columns={'Subject': 'Subject Count'}).sort_values('Stream')

                        # Store statistics in session state
                        st.session_state.total_exams = total_exams
                        st.session_state.total_semesters = total_semesters
                        st.session_state.total_branches = total_branches
                        st.session_state.overall_date_range = overall_date_range
                        st.session_state.unique_exam_days = unique_exam_days
                        st.session_state.non_elective_range = non_elective_range
                        st.session_state.elective_dates_str = elective_dates_str
                        st.session_state.stream_counts = stream_counts

                        # Generate and store downloadable files
                        st.write("Generating Excel...")
                        excel_data = save_to_excel(sem_dict)
                        if excel_data:
                            st.session_state.excel_data = excel_data.getvalue()
                        else:
                            st.write("Excel generation failed.")

                        st.write("Generating PDF...")
                        if sem_dict:
                            pdf_output = io.BytesIO()
                            temp_pdf_path = "temp_timetable.pdf"
                            generate_pdf_timetable(sem_dict, temp_pdf_path)
                            with open(temp_pdf_path, "rb") as f:
                                pdf_output.write(f.read())
                            pdf_output.seek(0)
                            if os.path.exists(temp_pdf_path):
                                os.remove(temp_pdf_path)
                            st.session_state.pdf_data = pdf_output.getvalue()
                        else:
                            st.write("PDF generation skipped due to empty sem_dict.")

                        st.write("Generating verification...")
                        verification_data = save_verification_excel(st.session_state.original_df, sem_dict)
                        if verification_data:
                            st.session_state.verification_data = verification_data.getvalue()
                        else:
                            st.write("Verification generation failed.")

                        st.markdown('<div class="status-success">🎉 Timetable generated successfully!</div>',
                                    unsafe_allow_html=True)

                    else:
                        st.markdown(
                            '<div class="status-error">❌ Failed to read the Excel file. Please check the format.</div>',
                            unsafe_allow_html=True)

                except Exception as e:
                    st.markdown(f'<div class="status-error">❌ An error occurred: {str(e)}</div>',
                                unsafe_allow_html=True)

    # Display timetable results if processing is complete
    if st.session_state.processing_complete:
        st.markdown("---")

        # Warn if exam days exceed limit
        if st.session_state.unique_exam_days > 20:
            st.warning(f"⚠️ The timetable spans {st.session_state.unique_exam_days} exam days, exceeding the limit of 20 days.")

        # Download options
        st.markdown("---")
        st.markdown("### 📥 Download Options")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            if st.session_state.excel_data:
                st.download_button(
                    label="📊 Download Excel File",
                    data=st.session_state.excel_data,
                    file_name=f"timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_excel"
                )

        with col2:
            if st.session_state.pdf_data:
                st.download_button(
                    label="📄 Download PDF File",
                    data=st.session_state.pdf_data,
                    file_name=f"timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="download_pdf"
                )

        with col3:
            if st.session_state.verification_data:
                st.download_button(
                    label="📋 Download Verification File",
                    data=st.session_state.verification_data,
                    file_name=f"verification_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_verification"
                )

        with col4:
            if st.button("🔄 Generate New Timetable", use_container_width=True):
                # Clear session state and rerun
                st.session_state.processing_complete = False
                st.session_state.timetable_data = {}
                st.session_state.original_df = None
                st.session_state.excel_data = None
                st.session_state.pdf_data = None
                st.session_state.verification_data = None
                st.session_state.total_exams = 0
                st.session_state.total_semesters = 0
                st.session_state.total_branches = 0
                st.session_state.overall_date_range = 0
                st.session_state.unique_exam_days = 0
                st.session_state.non_elective_range = "N/A"
                st.session_state.elective_dates_str = "N/A"
                st.session_state.stream_counts = pd.DataFrame()
                st.rerun()

        # Statistics Overview
        st.markdown("""
        <div class="stats-section">
            <h2>📈 Statistics Overview</h2>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="metric-card"><h3>📝 {st.session_state.total_exams}</h3><p>Total Exams</p></div>',
                        unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><h3>🎓 {st.session_state.total_semesters}</h3><p>Semesters</p></div>',
                        unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-card"><h3>🏫 {st.session_state.total_branches}</h3><p>Branches</p></div>',
                        unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-card"><h3>📅 {st.session_state.overall_date_range}</h3><p>Days Span</p></div>',
                        unsafe_allow_html=True)

        st.markdown("""
        <div class="metric-card">
            <h3>📆 Exam Dates Overview</h3>
            <table style="width: 100%; border-collapse: collapse; margin-top: 0.5rem;">
                <tr style="background: rgba(255, 255, 255, 0.1);">
                    <th style="padding: 0.5rem; text-align: left; border-bottom: 1px solid #ddd;">Type</th>
                    <th style="padding: 0.5rem; text-align: left; border-bottom: 1px solid #ddd;">Dates</th>
                </tr>
                <tr>
                    <td style="padding: 0.5rem; border-bottom: 1px solid #ddd;">Non-Elective Range</td>
                    <td style="padding: 0.5rem; border-bottom: 1px solid #ddd;">{non_elective_range}</td>
                </tr>
                <tr>
                    <td style="padding: 0.5rem;">Elective Dates</td>
                    <td style="padding: 0.5rem;">{elective_dates_str}</td>
                </tr>
            </table>
        </div>
        """.format(non_elective_range=st.session_state.non_elective_range, elective_dates_str=st.session_state.elective_dates_str),
                    unsafe_allow_html=True)

        st.markdown("#### Subjects Per Stream")
        if not st.session_state.stream_counts.empty:
            st.dataframe(st.session_state.stream_counts, hide_index=True, use_container_width=True)
        else:
            st.markdown('<div class="status-info">ℹ️ No stream data available.</div>', unsafe_allow_html=True)

        # Timetable Results
        st.markdown("---")
        st.markdown("""
        <div class="results-section">
            <h2>📊 Timetable Results</h2>
        </div>
        """, unsafe_allow_html=True)

        for sem, df_sem in st.session_state.timetable_data.items():
            st.markdown(f"### 📚 Semester {sem}")

            for main_branch in df_sem["MainBranch"].unique():
                main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()

                # Separate non-electives and electives for display
                df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()

                # Display non-electives
                if not df_non_elec.empty:
                    difficulty_str = df_non_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                    difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                    time_range_suffix = df_non_elec.apply(
                        lambda row: f" ({row['Time Slot'].split(' - ')[0]} to {calculate_end_time(row['Time Slot'].split(' - ')[0], row['Exam Duration'])}"
                        if row['Exam Duration'] != 3 else '', axis=1
                    )
                    df_non_elec["SubjectDisplay"] = df_non_elec["Subject"] + time_range_suffix + difficulty_suffix
                    df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                    df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                    df_non_elec = df_non_elec.drop_duplicates(subset=["Exam Date", "Time Slot", "SubBranch", "SubjectDisplay"])
                    pivot_df = df_non_elec.pivot_table(
                        index=["Exam Date", "Time Slot"],
                        columns="SubBranch",
                        values="SubjectDisplay",
                        aggfunc=lambda x: ", ".join(x)
                    ).fillna("---")
                    if not pivot_df.empty:
                        st.markdown(f"#### {main_branch_full} - Core Subjects")
                        formatted_pivot = pivot_df.copy()
                        if len(formatted_pivot.index.levels) > 0:
                            formatted_dates = [d.strftime("%d-%m-%Y") if pd.notna(d) else "" for d in
                                               formatted_pivot.index.levels[0]]
                            formatted_pivot.index = formatted_pivot.index.set_levels(formatted_dates, level=0)
                        st.dataframe(formatted_pivot, use_container_width=True)

                # Display electives
                if not df_elec.empty:
                    difficulty_str = df_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                    difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                    time_range_suffix = df_elec.apply(
                        lambda row: f" ({row['Time Slot'].split(' - ')[0]} to {calculate_end_time(row['Time Slot'].split(' - ')[0], row['Exam Duration'])}"
                        if row['Exam Duration'] != 3 else '', axis=1
                    )
                    df_elec["SubjectDisplay"] = df_elec["Subject"] + " [" + df_elec["OE"] + "]" + time_range_suffix + difficulty_suffix
                    df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                    df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                    elec_pivot = df_elec.groupby(['OE', 'Exam Date', 'Time Slot'])['SubjectDisplay'].apply(
                        lambda x: ", ".join(x)
                    ).reset_index()
                    if not elec_pivot.empty:
                        st.markdown(f"#### {main_branch_full} - Open Electives")
                        st.dataframe(elec_pivot, use_container_width=True)

    # Display footer
    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <p>🎓 <strong>Exam Timetable Generator</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
        <p style="font-size: 0.9em;">Streamlined scheduling • Conflict-free timetables • Multiple export formats</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
