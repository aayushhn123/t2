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
        if pdf.get_string_width(test_line)  pdf.h - footer_height:
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
            fixed_cols = ["Exam Date"]
            sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols and c != "Time Slot"]
            exam_date_width = 60
            table_font_size = 12
            line_height = 10

            for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
                chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
                cols_to_print = fixed_cols + chunk
                chunk_df = pivot_df[cols_to_print].copy()
                mask = chunk_df[chunk].apply(lambda row: row.astype(str).str.strip() != "").any(axis=1)
                chunk_df = chunk_df[mask].reset_index(drop=True)
                if chunk_df.empty:
                    continue

                # Get the time slot for this chunk
                time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and not pivot_df['Time Slot'].empty else None

                # Convert Exam Date to desired format
                chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

                # Modify subjects to include time range if duration != 3 hours
                for sub_branch in chunk:
                    for idx in chunk_df.index:
                        cell_value = chunk_df.at[idx, sub_branch]
                        if cell_value == "---":
                            continue
                        subjects = cell_value.split(", ")
                        modified_subjects = []
                        for subject in subjects:
                            duration = extract_duration(subject)
                            base_subject = re.sub(r' \[Duration: \d+\.?\d* hrs\]', '', subject)
                            if duration != 3 and time_slot:
                                start_time = time_slot.split(" - ")[0]
                                end_time = calculate_end_time(start_time, duration)
                                modified_subjects.append(f"{base_subject} ({start_time} to {end_time})")
                            else:
                                modified_subjects.append(base_subject)
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
                
                print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=line_height, 
                                 header_content=header_content, branches=chunk, time_slot=time_slot)

        # Handle electives
        if sheet_name.endswith('_Electives'):
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and not pivot_df['Time Slot'].empty else None
            
            elective_data = pivot_df.groupby('OE').agg({
                'Exam Date': 'first',
                'SubjectDisplay': lambda x: ", ".join(x)
            }).reset_index()

            # Convert Exam Date to desired format
            elective_data["Exam Date"] = pd.to_datetime(elective_data["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

            exam_date_width = 60
            subject_width = pdf.w - 2 * pdf.l_margin - exam_date_width
            col_widths = [exam_date_width, subject_width]
            cols_to_print = ['Exam Date', 'SubjectDisplay']
            
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
                    len(cleaned_text)  0:
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
            return m.get(sem.strip(), 0)
        
        df["Semester"] = df["Semester"].apply(convert_sem).astype(int)
        df["Branch"] = df["Program"].astype(str).str.strip() + "-" + df["Stream"].astype(str).str.strip()
        df["Subject"] = df["SubjectName"].astype(str) + " - (" + df["ModuleCode"].astype(str) + ")"
        #df = df[df["Campus"].str.strip() == "Mumbai"] # Campus filter
        
        comp_mask = (df["Category"] == "COMP") & df["Difficulty"].notna()
        df["Difficulty"] = None
        df.loc[comp_mask, "Difficulty"] = df.loc[comp_mask, "Difficulty"]
        
        df["Exam Date"] = ""
        df["Time Slot"] = ""
        df["Exam Duration"] = df["Exam Duration"].fillna(3).astype(float)  # Default to 3 hours if NaN
        
        # Handle StudentCount - fill NaN with 0 and convert to int
        df["StudentCount"] = df["StudentCount"].fillna(0).astype(int)
        
        df_non = df[df["Category"] != "INTD"].copy()
        df_ele = df[df["Category"] == "INTD"].copy()
        
        def split_br(b):
            p = b.split("-", 1)
            return pd.Series([p[0].strip(), p[1].strip() if len(p) > 1 else ""])
        
        for d in (df_non, df_ele):
            d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)
        
        # Updated columns list to include StudentCount
        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty", "Exam Duration", "StudentCount"]
        
        return df_non[cols], df_ele[cols], df
        
    except Exception as e:
        st.error(f"Error reading the Excel file: {str(e)}")
        return None, None, None

def schedule_semester_non_electives(df_sem, holidays, base_date, schedule_by_difficulty=False, global_module_schedule=None, target_span_days=20):
    df_sem['SubjectCode'] = df_sem['Subject'].str.extract(r'\((.*?)\)', expand=False)

    if isinstance(base_date, date) and not isinstance(base_date, datetime):
        base_date = datetime.combine(base_date, datetime.min.time())

    holidays_dates = {h.date() for h in holidays}
    all_branches = df_sem['Branch'].unique()
    
    # Enhanced tracking with better time slot utilization
    exam_schedule = {branch: {} for branch in all_branches}
    time_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
    
    # Track daily capacity utilization
    daily_capacity = {}  # {date: {time_slot: branches_count}}
    max_branches_per_slot = len(all_branches)

    def find_next_valid_day(start_day):
        day = start_day
        while True:
            day_date = day.date() if isinstance(day, datetime) else day
            if day.weekday() != 6 and day_date not in holidays_dates:
                return day
            day = day + timedelta(days=1)

    def get_daily_utilization(target_date):
        """Calculate how many branch-slots are used on a given date"""
        target_date_key = target_date.strftime("%Y-%m-%d")
        used_slots = 0
        total_slots = len(all_branches) * len(time_slots)
        
        for branch in all_branches:
            if target_date_key in exam_schedule[branch]:
                used_slots += len(exam_schedule[branch][target_date_key])
        
        return used_slots / total_slots if total_slots > 0 else 0

    def can_schedule_subject_optimized(subject_code, branches, target_date, target_time_slot):
        """Enhanced scheduling check with capacity awareness"""
        target_date_key = target_date.strftime("%Y-%m-%d")
        
        # Check basic conflicts
        for branch in branches:
            if target_date_key in exam_schedule[branch]:
                if target_time_slot in exam_schedule[branch][target_date_key]:
                    return False
        
        # Check if we're within target span and prefer fuller days
        utilization = get_daily_utilization(target_date)
        return True

    def find_optimal_slot(subject_code, branches, start_date, max_search_days=None):
        """Find the optimal slot considering span constraints and utilization"""
        if max_search_days is None:
            max_search_days = target_span_days
        
        best_slot = None
        best_score = -1
        
        for day_offset in range(max_search_days):
            candidate_date = start_date + timedelta(days=day_offset)
            candidate_date = find_next_valid_day(candidate_date)
            
            current_utilization = get_daily_utilization(candidate_date)
            
            for time_slot in time_slots:
                if can_schedule_subject_optimized(subject_code, branches, candidate_date, time_slot):
                    # Score based on: utilization (prefer fuller days), branch coverage, date proximity
                    utilization_score = current_utilization * 2  # Prefer days with existing exams
                    branch_coverage_score = len(branches) / len(all_branches)
                    date_proximity_score = 1 - (day_offset / max_search_days)
                    
                    total_score = utilization_score + branch_coverage_score + date_proximity_score
                    
                    if total_score > best_score:
                        best_score = total_score
                        best_slot = (candidate_date, time_slot)
        
        return best_slot

    def schedule_subject(subject_code, branches, target_date, target_time_slot):
        """Schedule a subject with enhanced tracking"""
        target_date_key = target_date.strftime("%Y-%m-%d")
        
        for branch in branches:
            if target_date_key not in exam_schedule[branch]:
                exam_schedule[branch][target_date_key] = {}
            exam_schedule[branch][target_date_key][target_time_slot] = subject_code

    # Apply global module schedule first
    if global_module_schedule:
        for idx, row in df_sem.iterrows():
            module_code = row['SubjectCode']
            if module_code in global_module_schedule:
                schedule_info = global_module_schedule[module_code]
                primary_time = schedule_info['primary_time']
                primary_date = schedule_info['primary_date']
                
                sem_branch_key = f"{row['Semester']}_{row['Branch']}"
                if sem_branch_key == schedule_info['primary_sem_branch']:
                    df_sem.at[idx, 'Exam Date'] = primary_date
                    df_sem.at[idx, 'Time Slot'] = primary_time
                else:
                    original_subject = row['Subject']
                    df_sem.at[idx, 'Subject'] = f"{original_subject} ({primary_time})"
                    df_sem.at[idx, 'Exam Date'] = primary_date
                    df_sem.at[idx, 'Time Slot'] = primary_time
                
                exam_date = datetime.strptime(primary_date, "%d-%m-%Y")
                schedule_subject(module_code, [row['Branch']], exam_date, primary_time)

    # Get unscheduled subjects and optimize their arrangement
    unscheduled_subjects = df_sem[df_sem['Exam Date'].isna() | (df_sem['Exam Date'] == '')]
    
    if unscheduled_subjects.empty:
        return df_sem

    # Enhanced subject grouping and prioritization
    subject_info = unscheduled_subjects.groupby('SubjectCode').agg({
        'Branch': lambda x: set(x),
        'Subject': 'first',
        'Difficulty': 'first' if 'Difficulty' in unscheduled_subjects.columns else lambda x: None,
        'StudentCount': 'sum' if 'StudentCount' in unscheduled_subjects.columns else lambda x: 0
    }).reset_index()
    
    # Multi-criteria sorting for optimal scheduling
    subject_info['branch_count'] = subject_info['Branch'].apply(len)
    subject_info['student_count'] = subject_info.get('StudentCount', 0)
    subject_info['difficulty_sort'] = subject_info['Difficulty'].fillna(0.5)
    
    # Sort by: 1) Branch coverage (desc), 2) Student count (desc), 3) Difficulty balance
    subject_info = subject_info.sort_values(
        ['branch_count', 'student_count', 'difficulty_sort'], 
        ascending=[False, False, True]
    )

    # Optimized scheduling with span awareness
    current_date = find_next_valid_day(base_date)
    
    for _, subject_row in subject_info.iterrows():
        subject_code = subject_row['SubjectCode']
        branches = subject_row['Branch']
        
        # Find optimal slot within target span
        optimal_slot = find_optimal_slot(subject_code, branches, current_date, target_span_days)
        
        if optimal_slot:
            target_date, target_time_slot = optimal_slot
            schedule_subject(subject_code, branches, target_date, target_time_slot)
            
            # Update dataframe
            mask = (df_sem['SubjectCode'] == subject_code) & df_sem['Branch'].isin(branches)
            df_sem.loc[mask, 'Exam Date'] = target_date.strftime("%d-%m-%Y")
            df_sem.loc[mask, 'Time Slot'] = target_time_slot
        else:
            # Fallback: extend beyond target span if absolutely necessary
            extend_date = current_date + timedelta(days=target_span_days)
            extend_date = find_next_valid_day(extend_date)
            
            semester = df_sem["Semester"].iloc[0]
            default_time_slot = get_subject_time_slot(semester)
            
            while not can_schedule_subject_optimized(subject_code, branches, extend_date, default_time_slot):
                extend_date = find_next_valid_day(extend_date + timedelta(days=1))
            
            schedule_subject(subject_code, branches, extend_date, default_time_slot)
            mask = (df_sem['SubjectCode'] == subject_code) & df_sem['Branch'].isin(branches)
            df_sem.loc[mask, 'Exam Date'] = extend_date.strftime("%d-%m-%Y")
            df_sem.loc[mask, 'Time Slot'] = default_time_slot

    return df_sem

def get_subject_time_slot(semester):
    """Optimized time slot assignment for better distribution"""
    if semester % 2 != 0:  # Odd semesters
        odd_sem_position = (semester + 1) // 2
        return "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:  # Even semesters
        even_sem_position = semester // 2
        return "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"

def schedule_electives_mainbranch(df_elec, elective_base_date, holidays, max_days=20, global_module_schedule=None):
    """
    Optimized elective scheduling with time slot utilization
    """
    df_elec['SubjectCode'] = df_elec['Subject'].str.extract(r'\((.*?)\)', expand=False)
    
    # Apply global module schedule first if available
    if global_module_schedule:
        for idx, row in df_elec.iterrows():
            module_code = row['SubjectCode']
            if module_code in global_module_schedule:
                schedule_info = global_module_schedule[module_code]
                primary_time = schedule_info['primary_time']
                primary_date = schedule_info['primary_date']
                
                sem_branch_key = f"{row['Semester']}_{row.get('Branch', row.get('SubBranch', ''))}"
                if sem_branch_key == schedule_info['primary_sem_branch']:
                    df_elec.at[idx, 'Exam Date'] = primary_date
                    df_elec.at[idx, 'Time Slot'] = primary_time
                else:
                    original_subject = row['Subject']
                    df_elec.at[idx, 'Subject'] = f"{original_subject} ({primary_time})"
                    df_elec.at[idx, 'Exam Date'] = primary_date
                    df_elec.at[idx, 'Time Slot'] = primary_time
                continue
    
    # Filter unscheduled electives
    unscheduled_electives = df_elec[df_elec['Exam Date'].isna() | (df_elec['Exam Date'] == '')]
    if unscheduled_electives.empty:
        return df_elec
    
    # Group by OE (Open Elective) and SubBranch to optimize scheduling
    oe_info = unscheduled_electives.groupby('OE').agg({
        'SubBranch': lambda x: set(x),
        'Subject': list
    }).reset_index()
    
    # Sort by number of sub-branches (most constrained first)
    oe_info['subbranch_count'] = oe_info['SubBranch'].apply(len)
    oe_info = oe_info.sort_values('subbranch_count', ascending=False)
    
    # Track scheduling per sub-branch per day per time slot
    subbranch_schedule = {}
    for sb in df_elec['SubBranch'].unique():
        subbranch_schedule[sb] = {}
    
    if isinstance(elective_base_date, date) and not isinstance(elective_base_date, datetime):
        elective_base_date = datetime.combine(elective_base_date, datetime.min.time())
    
    holidays_dates = {h.date() for h in holidays}
    time_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
    
    def find_next_valid_day(start_day):
        day = start_day
        while True:
            day_date = day.date() if isinstance(day, datetime) else day
            if day.weekday() != 6 and day_date not in holidays_dates:
                return day
            day = day + timedelta(days=1)
    
    def can_schedule_oe(subbranches, target_date, target_time_slot):
        target_date_key = target_date.strftime("%Y-%m-%d")
        for sb in subbranches:
            if target_date_key in subbranch_schedule[sb]:
                if target_time_slot in subbranch_schedule[sb][target_date_key]:
                    return False
        return True
    
    def schedule_oe(oe_code, subbranches, target_date, target_time_slot):
        target_date_key = target_date.strftime("%Y-%m-%d")
        for sb in subbranches:
            if target_date_key not in subbranch_schedule[sb]:
                subbranch_schedule[sb][target_date_key] = {}
            subbranch_schedule[sb][target_date_key][target_time_slot] = oe_code
    
    # Get default time slot
    semester = df_elec["Semester"].iloc[0]
    if semester % 2 != 0:  # Odd semesters
        odd_sem_position = (semester + 1) // 2
        default_time_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:  # Even semesters
        even_sem_position = semester // 2
        default_time_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    
    # Schedule OEs efficiently
    current_date = find_next_valid_day(elective_base_date)
    
    for _, oe_row in oe_info.iterrows():
        oe_code = oe_row['OE']
        subbranches = oe_row['SubBranch']
        
        scheduled = False
        
        # Try to schedule within max_days range
        for day_offset in range(max_days):
            candidate_date = current_date + timedelta(days=day_offset)
            candidate_date = find_next_valid_day(candidate_date)
            
            # Try both time slots
            for time_slot in time_slots:
                if can_schedule_oe(subbranches, candidate_date, time_slot):
                    schedule_oe(oe_code, subbranches, candidate_date, time_slot)
                    
                    # Update dataframe
                    mask = df_elec['OE'] == oe_code
                    df_elec.loc[mask, 'Exam Date'] = candidate_date.strftime("%d-%m-%Y")
                    df_elec.loc[mask, 'Time Slot'] = time_slot
                    
                    scheduled = True
                    break
            
            if scheduled:
                break
        
        if not scheduled:
            # Extend beyond max_days if necessary
            extend_date = current_date + timedelta(days=max_days)
            extend_date = find_next_valid_day(extend_date)
            
            while not can_schedule_oe(subbranches, extend_date, default_time_slot):
                extend_date = find_next_valid_day(extend_date + timedelta(days=1))
            
            schedule_oe(oe_code, subbranches, extend_date, default_time_slot)
            mask = df_elec['OE'] == oe_code
            df_elec.loc[mask, 'Exam Date'] = extend_date.strftime("%d-%m-%Y")
            df_elec.loc[mask, 'Time Slot'] = default_time_slot
    
    return df_elec

def create_global_module_schedule(df_all_semesters, holidays, base_date, target_span_days=20):
    """Optimized global scheduling with span constraints"""
    global_schedule = {}
    
    df_all_semesters['SubjectCode'] = df_all_semesters['Subject'].str.extract(r'\((.*?)\)', expand=False)
    
    module_info = df_all_semesters.groupby('SubjectCode').agg({
        'Semester': list,
        'Branch': list,
        'StudentCount': list,
        'Subject': 'first'
    }).reset_index()
    
    # Enhanced filtering and sorting for common modules
    common_modules = []
    for idx, row in module_info.iterrows():
        if pd.isna(row['SubjectCode']) or row['SubjectCode'] == '':
            continue
            
        sem_branch_combinations = list(zip(row['Semester'], row['Branch']))
        unique_combinations = set(sem_branch_combinations)
        
        if len(unique_combinations) > 1:
            total_students = sum(row['StudentCount'])
            complexity_score = len(unique_combinations) * total_students
            
            common_modules.append({
                'module_code': row['SubjectCode'],
                'combinations': sem_branch_combinations,
                'student_counts': row['StudentCount'],
                'subject_name': row['Subject'],
                'complexity_score': complexity_score,
                'unique_combinations': len(unique_combinations)
            })
    
    # Sort by complexity and student impact
    common_modules.sort(key=lambda x: (x['complexity_score'], x['unique_combinations']), reverse=True)
    
    # Optimized scheduling with better time slot distribution
    current_date = base_date
    if isinstance(current_date, date) and not isinstance(current_date, datetime):
        current_date = datetime.combine(current_date, datetime.min.time())
    
    holidays_dates = {h.date() for h in holidays}
    time_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
    
    # Track utilization more efficiently
    daily_slot_usage = {}  # {date_str: [used_slots]}
    max_slots_per_day = 2
    
    def get_next_optimal_slot(start_date, prefer_fuller_days=True):
        """Find next slot optimizing for span and utilization"""
        date_to_check = start_date
        
        for day_offset in range(target_span_days):
            current_check_date = date_to_check + timedelta(days=day_offset)
            date_only = current_check_date.date()
            date_str = current_check_date.strftime("%Y-%m-%d")
            
            if (current_check_date.weekday() != 6 and 
                date_only not in holidays_dates):
                
                used_slots = daily_slot_usage.get(date_str, [])
                available_slots = [slot for slot in time_slots if slot not in used_slots]
                
                if available_slots:
                    # Prefer days with some usage if prefer_fuller_days is True
                    if prefer_fuller_days and len(used_slots) > 0:
                        return current_check_date, available_slots[0]
                    elif not prefer_fuller_days and len(used_slots) == 0:
                        return current_check_date, available_slots[0]
                    elif len(available_slots) > 0:
                        return current_check_date, available_slots[0]
        
        # If no slot found in target span, extend
        extend_date = start_date + timedelta(days=target_span_days)
        while True:
            date_only = extend_date.date()
            if (extend_date.weekday() != 6 and date_only not in holidays_dates):
                return extend_date, time_slots[0]
            extend_date = extend_date + timedelta(days=1)
    
    # Schedule common modules with span optimization
    for module in common_modules:
        module_code = module['module_code']
        combinations = module['combinations']
        student_counts = module['student_counts']
        
        # Find primary combination (highest student count)
        max_count_idx = student_counts.index(max(student_counts))
        primary_sem, primary_branch = combinations[max_count_idx]
        primary_sem_branch = f"{primary_sem}_{primary_branch}"
        
        # Get optimal slot
        exam_date, time_slot = get_next_optimal_slot(current_date)
        
        # Track usage
        date_str = exam_date.strftime("%Y-%m-%d")
        if date_str not in daily_slot_usage:
            daily_slot_usage[date_str] = []
        daily_slot_usage[date_str].append(time_slot)
        
        global_schedule[module_code] = {
            'primary_date': exam_date.strftime("%d-%m-%Y"),
            'primary_time': time_slot,
            'primary_sem_branch': primary_sem_branch,
            'all_combinations': combinations
        }
        
        # Advance current_date strategically
        if len(daily_slot_usage[date_str]) >= max_slots_per_day:
            current_date = exam_date + timedelta(days=1)
    
    return global_schedule

def process_constraints(df, holidays, base_date, schedule_by_difficulty=False, target_span_days=20):
    """Enhanced constraint processing with span optimization"""
    
    # Create optimized global schedule
    global_module_schedule = create_global_module_schedule(df, holidays, base_date, target_span_days)
    
    final_list = []
    for sem in sorted(df["Semester"].unique()):
        if sem == 0:
            continue
        df_sem = df[df["Semester"] == sem].copy()
        if df_sem.empty:
            continue
        
        # Pass target span to semester scheduling
        scheduled_sem = schedule_semester_non_electives(
            df_sem, holidays, base_date, schedule_by_difficulty, 
            global_module_schedule, target_span_days
        )
        final_list.append(scheduled_sem)
    
    if not final_list:
        return {}
    
    df_combined = pd.concat(final_list, ignore_index=True)
    
    # Validate and report span
    all_exam_dates = df_combined['Exam Date'].dropna()
    if not all_exam_dates.empty:
        exam_dates = [datetime.strptime(date_str, "%d-%m-%Y") for date_str in all_exam_dates if date_str != '']
        if exam_dates:
            first_exam = min(exam_dates)
            last_exam = max(exam_dates)
            actual_span = (last_exam - first_exam).days + 1
            
            # Count working days
            working_days = 0
            current = first_exam
            holidays_dates = {h.date() for h in holidays}
            while current  target_span_days:
                print(f"⚠️ Target span of {target_span_days} days exceeded by {working_days - target_span_days} days")
    
    sem_dict = {}
    for sem in sorted(df_combined["Semester"].unique()):
        sem_dict[sem] = df_combined[df_combined["Semester"] == sem].copy()
    
    return sem_dict

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
                    df_non_elec["SubjectDisplay"] = df_non_elec["Subject"] + difficulty_suffix
                    
                    # Add duration info for non-3-hour exams
                    duration_suffix = df_non_elec['Exam Duration'].apply(
                        lambda x: f" [Duration: {x} hrs]" if pd.notna(x) and x != 3 else ""
                    )
                    df_non_elec["SubjectDisplay"] = df_non_elec["SubjectDisplay"] + duration_suffix
                    
                    # Create pivot table for non-electives
                    pivot_df = df_non_elec.pivot_table(
                        index=['Exam Date', 'Time Slot'], 
                        columns='SubBranch', 
                        values='SubjectDisplay', 
                        aggfunc=lambda x: ', '.join(x), 
                        fill_value='---'
                    )
                    
                    # Reset index to make Exam Date and Time Slot regular columns
                    pivot_df = pivot_df.reset_index()
                    
                    # Sort by exam date
                    pivot_df['Exam Date'] = pd.to_datetime(pivot_df['Exam Date'], format='%d-%m-%Y')
                    pivot_df = pivot_df.sort_values('Exam Date')
                    pivot_df['Exam Date'] = pivot_df['Exam Date'].dt.strftime('%d-%m-%Y')
                    
                    # Save to Excel
                    sheet_name = f"{main_branch}_Sem_{int_to_roman(sem)}"
                    pivot_df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Process electives
                if not df_elec.empty:
                    difficulty_str = df_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                    difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                    df_elec["SubjectDisplay"] = df_elec["Subject"] + difficulty_suffix
                    
                    # Add duration info for non-3-hour exams
                    duration_suffix = df_elec['Exam Duration'].apply(
                        lambda x: f" [Duration: {x} hrs]" if pd.notna(x) and x != 3 else ""
                    )
                    df_elec["SubjectDisplay"] = df_elec["SubjectDisplay"] + duration_suffix
                    
                    # Create pivot table for electives
                    elective_pivot = df_elec.pivot_table(
                        index=['Exam Date', 'Time Slot', 'OE'], 
                        columns='SubBranch', 
                        values='SubjectDisplay', 
                        aggfunc=lambda x: ', '.join(x), 
                        fill_value='---'
                    )
                    
                    # Reset index
                    elective_pivot = elective_pivot.reset_index()
                    
                    # Sort by exam date
                    elective_pivot['Exam Date'] = pd.to_datetime(elective_pivot['Exam Date'], format='%d-%m-%Y')
                    elective_pivot = elective_pivot.sort_values('Exam Date')
                    elective_pivot['Exam Date'] = elective_pivot['Exam Date'].dt.strftime('%d-%m-%Y')
                    
                    # Save to Excel
                    sheet_name = f"{main_branch}_Sem_{int_to_roman(sem)}_Electives"
                    elective_pivot.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output

def process_electives(df_electives, holidays, base_date, schedule_by_difficulty=False, target_span_days=20):
    """Process electives with optimized scheduling"""
    if df_electives.empty:
        return {}
    
    # Create optimized global schedule for electives
    global_module_schedule = create_global_module_schedule(df_electives, holidays, base_date, target_span_days)
    
    final_list = []
    for sem in sorted(df_electives["Semester"].unique()):
        if sem == 0:
            continue
        df_sem = df_electives[df_electives["Semester"] == sem].copy()
        if df_sem.empty:
            continue
        
        for main_branch in df_sem["MainBranch"].unique():
            df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
            if df_mb.empty:
                continue
            
            # Calculate elective base date (after non-electives)
            elective_base_date = base_date + timedelta(days=target_span_days)
            
            # Schedule electives for this main branch
            scheduled_electives = schedule_electives_mainbranch(
                df_mb, elective_base_date, holidays, target_span_days, global_module_schedule
            )
            final_list.append(scheduled_electives)
    
    if not final_list:
        return {}
    
    df_combined = pd.concat(final_list, ignore_index=True)
    
    # Validate and report elective span
    all_exam_dates = df_combined['Exam Date'].dropna()
    if not all_exam_dates.empty:
        exam_dates = [datetime.strptime(date_str, "%d-%m-%Y") for date_str in all_exam_dates if date_str != '']
        if exam_dates:
            first_exam = min(exam_dates)
            last_exam = max(exam_dates)
            actual_span = (last_exam - first_exam).days + 1
            
            # Count working days
            working_days = 0
            current = first_exam
            holidays_dates = {h.date() for h in holidays}
            while current <= last_exam:
                if current.weekday() != 6 and current.date() not in holidays_dates:
                    working_days += 1
                current += timedelta(days=1)
            
            print(f"Electives span: {actual_span} calendar days ({working_days} working days)")
    
    sem_dict = {}
    for sem in sorted(df_combined["Semester"].unique()):
        sem_dict[sem] = df_combined[df_combined["Semester"] == sem].copy()
    
    return sem_dict

def display_statistics(semester_wise_timetable, elective_timetable=None):
    """Display enhanced statistics with span information"""
    if not semester_wise_timetable:
        return
    
    # Combine all data
    all_data = []
    for sem_data in semester_wise_timetable.values():
        all_data.append(sem_data)
    
    if elective_timetable:
        for sem_data in elective_timetable.values():
            all_data.append(sem_data)
    
    if not all_data:
        return
    
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Calculate statistics
    total_subjects = len(combined_df)
    total_branches = len(combined_df['Branch'].unique())
    total_semesters = len(combined_df['Semester'].unique())
    
    # Calculate exam span
    all_exam_dates = combined_df['Exam Date'].dropna()
    span_info = "Not calculated"
    working_days = 0
    
    if not all_exam_dates.empty:
        exam_dates = [datetime.strptime(date_str, "%d-%m-%Y") for date_str in all_exam_dates if date_str != '']
        if exam_dates:
            first_exam = min(exam_dates)
            last_exam = max(exam_dates)
            total_span = (last_exam - first_exam).days + 1
            
            # Count working days (excluding Sundays and holidays)
            current = first_exam
            while current <= last_exam:
                if current.weekday() != 6:  # Not Sunday
                    working_days += 1
                current += timedelta(days=1)
            
            span_info = f"{total_span} calendar days ({working_days} working days)"
    
    # Display metrics in columns
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <h3>📚 {total_subjects}</h3>
            <p>Total Subjects</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <h3>🎓 {total_branches}</h3>
            <p>Total Branches</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <h3>📖 {total_semesters}</h3>
            <p>Semesters</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <h3>📅 {working_days}</h3>
            <p>Working Days</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Additional statistics
    st.markdown(f"""
    <div class="stats-section">
        <h3>📊 Detailed Statistics</h3>
        <p><strong>Exam Span:</strong> {span_info}</p>
        <p><strong>First Exam:</strong> {first_exam.strftime('%A, %d %B %Y') if exam_dates else 'N/A'}</p>
        <p><strong>Last Exam:</strong> {last_exam.strftime('%A, %d %B %Y') if exam_dates else 'N/A'}</p>
        <p><strong>Average Subjects per Day:</strong> {total_subjects / working_days:.1f} subjects/day</p>
    </div>
    """, unsafe_allow_html=True)

def main():
    """Main application function"""
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📅 Exam Timetable Generator</h1>
        <p>Generate optimized exam schedules with intelligent conflict resolution</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # File upload
        uploaded_file = st.file_uploader(
            "📁 Upload Excel File", 
            type=['xlsx', 'xls'],
            help="Upload your exam data in Excel format"
        )
        
        # Date selection
        base_date = st.date_input(
            "📅 Exam Start Date",
            value=datetime.now().date(),
            help="Select the starting date for exams"
        )
        
        # Target span configuration
        target_span = st.slider(
            "🎯 Target Exam Span (Working Days)",
            min_value=15,
            max_value=30,
            value=20,
            help="Target number of working days for the exam period"
        )
        
        # Difficulty-based scheduling
        schedule_by_difficulty = st.checkbox(
            "🧠 Schedule by Difficulty",
            value=False,
            help="Prioritize difficult subjects for earlier dates"
        )
        
        # Holiday configuration
        with st.expander("🏖️ Holiday Configuration"):
            st.write("Configure holidays that should be excluded from exam scheduling")
            
            # Predefined holidays
            predefined_holidays = [
                datetime(2025, 1, 26),  # Republic Day
                datetime(2025, 8, 15),  # Independence Day
                datetime(2025, 10, 2),  # Gandhi Jayanti
            ]
            
            # Custom holiday input
            custom_holiday = st.date_input(
                "Add Custom Holiday",
                value=None,
                help="Add additional holidays"
            )
            
            if custom_holiday:
                predefined_holidays.append(datetime.combine(custom_holiday, datetime.min.time()))
            
            holidays = predefined_holidays
            
            # Display selected holidays
            if holidays:
                st.write("**Selected Holidays:**")
                for holiday in holidays:
                    st.write(f"• {holiday.strftime('%d %B %Y')}")
    
    # Main content area
    if uploaded_file is not None:
        try:
            # Read and process the uploaded file
            with st.spinner("📖 Reading and processing file..."):
                df_non_electives, df_electives, df_original = read_timetable(uploaded_file)
            
            if df_non_electives is not None and df_electives is not None:
                st.success("✅ File uploaded and processed successfully!")
                
                # Display file statistics
                st.markdown("""
                <div class="stats-section">
                    <h3>📋 File Statistics</h3>
                </div>
                """, unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Non-Elective Subjects", len(df_non_electives))
                with col2:
                    st.metric("Elective Subjects", len(df_electives))
                with col3:
                    st.metric("Total Records", len(df_original))
                
                # Generate timetable button
                if st.button("🚀 Generate Timetable", type="primary"):
                    with st.spinner("⚡ Generating optimized timetable..."):
                        # Process non-electives
                        semester_wise_timetable = process_constraints(
                            df_non_electives, 
                            holidays, 
                            base_date, 
                            schedule_by_difficulty,
                            target_span
                        )
                        
                        # Process electives
                        elective_timetable = process_electives(
                            df_electives, 
                            holidays, 
                            base_date, 
                            schedule_by_difficulty,
                            target_span
                        )
                    
                    if semester_wise_timetable or elective_timetable:
                        st.success("🎉 Timetable generated successfully!")
                        
                        # Display statistics
                        display_statistics(semester_wise_timetable, elective_timetable)
                        
                        # Results section
                        st.markdown("""
                        <div class="results-section">
                            <h2>📊 Generated Timetables</h2>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Display timetables
                        if semester_wise_timetable:
                            st.subheader("📚 Non-Elective Subjects")
                            for sem, df_sem in semester_wise_timetable.items():
                                with st.expander(f"Semester {sem}", expanded=False):
                                    st.dataframe(df_sem, use_container_width=True)
                        
                        if elective_timetable:
                            st.subheader("🎯 Elective Subjects")
                            for sem, df_sem in elective_timetable.items():
                                with st.expander(f"Semester {sem} - Electives", expanded=False):
                                    st.dataframe(df_sem, use_container_width=True)
                        
                        # Download options
                        st.markdown("### 📥 Download Options")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            # Excel download
                            combined_timetable = {**semester_wise_timetable}
                            if elective_timetable:
                                for sem, df_sem in elective_timetable.items():
                                    if sem in combined_timetable:
                                        combined_timetable[sem] = pd.concat([combined_timetable[sem], df_sem], ignore_index=True)
                                    else:
                                        combined_timetable[sem] = df_sem
                            
                            excel_data = save_to_excel(combined_timetable)
                            if excel_data:
                                st.download_button(
                                    label="📊 Download Excel",
                                    data=excel_data.getvalue(),
                                    file_name=f"exam_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        with col2:
                            # PDF download
                            if st.button("📄 Generate PDF"):
                                with st.spinner("📄 Generating PDF..."):
                                    try:
                                        pdf_path = f"exam_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                                        generate_pdf_timetable(combined_timetable, pdf_path)
                                        
                                        with open(pdf_path, "rb") as pdf_file:
                                            st.download_button(
                                                label="📄 Download PDF",
                                                data=pdf_file.read(),
                                                file_name=pdf_path,
                                                mime="application/pdf"
                                            )
                                        
                                        # Clean up
                                        if os.path.exists(pdf_path):
                                            os.remove(pdf_path)
                                            
                                    except Exception as e:
                                        st.error(f"Error generating PDF: {str(e)}")
                    else:
                        st.error("❌ Failed to generate timetable. Please check your data.")
            else:
                st.error("❌ Error processing the uploaded file. Please check the file format.")
                
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
    else:
        # Welcome message and instructions
        st.markdown("""
        <div class="upload-section">
            <h2>👋 Welcome to the Exam Timetable Generator</h2>
            <p>This application helps you generate optimized exam schedules with the following features:</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Feature cards
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            <div class="feature-card">
                <h3>🎯 Smart Scheduling</h3>
                <p>Automatically resolves conflicts and optimizes exam distribution across the target span of days.</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("""
            <div class="feature-card">
                <h3>📊 Multiple Formats</h3>
                <p>Export your timetables in both Excel and PDF formats for easy sharing and printing.</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="feature-card">
                <h3>🏖️ Holiday Management</h3>
                <p>Automatically excludes weekends and holidays from exam scheduling.</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("""
            <div class="feature-card">
                <h3>⚡ Optimized Performance</h3>
                <p>Advanced algorithms ensure efficient scheduling within your target timeframe.</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="status-info">
            <h3>📋 Getting Started</h3>
            <ol>
                <li>Upload your Excel file containing exam data</li>
                <li>Configure your preferences in the sidebar</li>
                <li>Set your target exam span (recommended: 20 working days)</li>
                <li>Click "Generate Timetable" to create your schedule</li>
                <li>Download in your preferred format</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    # Footer
    st.markdown("""
    <div class="footer">
        <p>🎓 Exam Timetable Generator | Optimized for Academic Excellence</p>
        <p>Built with ❤️ using Streamlit</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()


