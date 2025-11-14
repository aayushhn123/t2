import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import os
import io
from PyPDF2 import PdfReader, PdfWriter
import re

# Set page configuration
st.set_page_config(
    page_title="Excel to PDF Timetable Converter",
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

        .feature-card {
            background: white;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin: 1rem 0;
            border-left: 4px solid #951C1C;
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

        .feature-card {
            background: #333;
            padding: 1.5rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.3);
            margin: 1rem 0;
            border-left: 4px solid #A23217;
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
    "MCA": "MASTER OF COMPUTER APPLICATIONS",
    "DIPLOMA": "DIPLOMA IN ENGINEERING"
}

# Define logo path
LOGO_PATH = "logo.png"

# Cache for text wrapping results
wrap_text_cache = {}

def wrap_text(pdf, text, col_width):
    cache_key = (text, col_width)
    if cache_key in wrap_text_cache:
        return wrap_text_cache[cache_key]
    
    lines = []
    current_line = ""
    words = re.split(r'(\s+)', text)
    for word in words:
        test_line = current_line + word
        if pdf.get_string_width(test_line) <= col_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line.strip())
            current_line = word
    if current_line:
        lines.append(current_line.strip())
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
            pdf.cell(col_widths[i] - 2 * cell_padding, line_height, ln, border=0, align='L' if i == 0 else 'C')
        pdf.rect(cx, y0, col_widths[i], row_h)
        pdf.set_xy(cx + col_widths[i], y0)

    setattr(pdf, '_row_counter', row_number + 1)
    pdf.set_xy(x0, y0 + row_h)

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, branches=None, time_slot=None, actual_time_slots=None):
    if df.empty:
        return
    setattr(pdf, '_row_counter', 0)
    
    # Add footer first
    footer_height = 25
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    
    # Add page numbers in bottom right
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')
    
    # Add header
    header_height = 95
    pdf.set_y(0)
    
    logo_width = 45
    logo_x = (pdf.w - logo_width) / 2
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    
    # Get selected college name from session state
    college_name = st.session_state.get('selected_college', 'SVKM\'s NMIMS University')
    
    # Adjust font size based on college name length
    if len(college_name) > 80:
        pdf.set_font("Arial", 'B', 12)
    elif len(college_name) > 60:
        pdf.set_font("Arial", 'B', 14)
    else:
        pdf.set_font("Arial", 'B', 16)
    
    pdf.rect(10, 30, pdf.w - 20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w - 20, 14, college_name, 0, 1, 'C')
    
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w - 20, 8, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')
    
    # Display actual time slots used
    if actual_time_slots and len(actual_time_slots) > 0:
        pdf.set_font("Arial", 'B', 13)
        pdf.set_xy(10, 59)
        
        if len(actual_time_slots) == 1:
            slot_text = f"Exam Time: {list(actual_time_slots)[0]}"
        else:
            sorted_slots = sorted(actual_time_slots)
            slot_text = f"Exam Times: {' | '.join(sorted_slots)}"
        
        pdf.cell(pdf.w - 20, 6, slot_text, 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check individual subject for specific exam time if multiple slots shown)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    elif time_slot:
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
    
    # Print data rows
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row):
            continue
            
        # Estimate row height
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 4
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
            add_footer_with_page_number(pdf, footer_height)
            
            # Add header to new page
            add_header_to_page(pdf, logo_x, logo_width, header_content, branches, time_slot, actual_time_slots)
            
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
    
    # Add page numbers in bottom right
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')

def add_header_to_page(pdf, logo_x, logo_width, header_content, branches, time_slot=None, actual_time_slots=None):
    """Add header to a new page"""
    pdf.set_y(0)
    
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)
    
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    
    # Get selected college name
    college_name = st.session_state.get('selected_college', 'SVKM\'s NMIMS University')
    
    # Adjust font size
    if len(college_name) > 80:
        pdf.set_font("Arial", 'B', 12)
    elif len(college_name) > 60:
        pdf.set_font("Arial", 'B', 14)
    else:
        pdf.set_font("Arial", 'B', 16)
    
    pdf.rect(10, 30, pdf.w - 20, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(pdf.w - 20, 14, college_name, 0, 1, 'C')
    
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(10, 51)
    pdf.cell(pdf.w - 20, 8, f"{header_content['main_branch_full']} - Semester {header_content['semester_roman']}", 0, 1, 'C')
    
    # Display time slots
    if actual_time_slots and len(actual_time_slots) > 0:
        pdf.set_font("Arial", 'B', 13)
        pdf.set_xy(10, 59)
        
        if len(actual_time_slots) == 1:
            slot_text = f"Exam Time: {list(actual_time_slots)[0]}"
        else:
            sorted_slots = sorted(actual_time_slots)
            slot_text = f"Exam Times: {' | '.join(sorted_slots)}"
        
        pdf.cell(pdf.w - 20, 6, slot_text, 0, 1, 'C')
        pdf.set_font("Arial", 'I', 10)
        pdf.set_xy(10, 65)
        pdf.cell(pdf.w - 20, 6, "(Check individual subject for specific exam time if multiple slots shown)", 0, 1, 'C')
        pdf.set_font("Arial", '', 12)
        pdf.set_xy(10, 71)
        pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(branches)}", 0, 1, 'C')
        pdf.set_y(85)
    elif time_slot:
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

def int_to_roman(num):
    """Convert integer to Roman numeral"""
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

def read_verification_excel(uploaded_file):
    """Read the NEW verification Excel file format"""
    try:
        # Try to read all sheets
        excel_file = pd.ExcelFile(uploaded_file)
        st.write(f"📋 Found {len(excel_file.sheet_names)} sheets in the file")
        
        # Look for the Verification sheet
        if 'Verification' in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name='Verification', engine='openpyxl')
            st.success("✅ Found 'Verification' sheet")
        else:
            # Use the first sheet if Verification not found
            df = pd.read_excel(uploaded_file, sheet_name=0, engine='openpyxl')
            st.warning(f"⚠️ 'Verification' sheet not found, using sheet: {excel_file.sheet_names[0]}")
        
        st.write("📊 Columns found in the verification file:")
        st.write(list(df.columns))
        
        # NEW column names from your updated code
        required_columns = ['Program', 'Stream', 'Current Session', 'Module Description', 'Exam Date', 'Exam Time']
        optional_columns = ['Module Abbreviation', 'CM Group', 'Exam Slot Number', 'Student count', 'Campus']
        
        # Check for required columns
        missing_required = [col for col in required_columns if col not in df.columns]
        if missing_required:
            st.error(f"❌ Missing required columns: {missing_required}")
            st.info("💡 Required columns: " + ", ".join(required_columns))
            return None
        
        # Filter out rows with no exam date or "Not Scheduled"
        df = df[
            (df['Exam Date'].notna()) & 
            (df['Exam Date'] != "") & 
            (df['Exam Date'] != 'Not Scheduled') &
            (df['Exam Date'].astype(str).str.strip() != "")
        ].copy()
        
        if df.empty:
            st.error("❌ No valid scheduled exams found in the verification file")
            return None
        
        # Clean data
        df['Program'] = df['Program'].astype(str).str.strip().str.upper()
        df['Stream'] = df['Stream'].astype(str).str.strip()
        df['Current Session'] = df['Current Session'].astype(str).str.strip()
        df['Module Description'] = df['Module Description'].astype(str).str.strip()
        df['Exam Date'] = df['Exam Date'].astype(str).str.strip()
        df['Exam Time'] = df['Exam Time'].astype(str).str.strip()
        
        # Handle optional columns
        if 'Module Abbreviation' in df.columns:
            df['Module Abbreviation'] = df['Module Abbreviation'].astype(str).str.strip()
        
        if 'CM Group' in df.columns:
            df['CM Group'] = df['CM Group'].fillna("").astype(str).str.strip()
        
        if 'Exam Slot Number' in df.columns:
            df['Exam Slot Number'] = pd.to_numeric(df['Exam Slot Number'], errors='coerce').fillna(0).astype(int)
        
        st.success(f"✅ Successfully loaded {len(df)} scheduled exam records")
        
        # Show sample
        st.write("📋 Sample of loaded data:")
        display_cols = ['Program', 'Stream', 'Current Session', 'Module Description', 'Exam Date', 'Exam Time']
        available_display_cols = [col for col in display_cols if col in df.columns]
        st.dataframe(df[available_display_cols].head(3))
        
        return df
        
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")
        import traceback
        st.error(f"Full error details: {traceback.format_exc()}")
        return None

def create_excel_sheets_for_pdf(df):
    """Convert verification data to Excel sheet format for PDF generation"""
    
    # Parse Program-Stream combination
    def parse_program_stream(row):
        program = str(row['Program']).strip().upper()
        stream = str(row['Stream']).strip()
        
        # Create combined identifier
        if stream and stream != 'nan' and stream != program:
            return f"{program}-{stream}"
        else:
            return program
    
    df['Branch'] = df.apply(parse_program_stream, axis=1)
    
    # Parse semester to get number
    def parse_semester(session_str):
        if pd.isna(session_str):
            return 1
        
        session_str = str(session_str).strip()
        
        # Extract number from various formats
        import re
        num_match = re.search(r'(\d+)', session_str)
        if num_match:
            return int(num_match.group(1))
        
        # Try Roman numerals
        roman_to_num = {
            'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6,
            'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10, 'XI': 11, 'XII': 12
        }
        for roman, num in roman_to_num.items():
            if roman in session_str.upper():
                return num
        
        return 1
    
    df['Semester'] = df['Current Session'].apply(parse_semester)
    
    # Group by main program and semester
    def get_main_program(branch):
        # Extract main program part before dash
        if '-' in branch:
            return branch.split('-')[0].strip()
        return branch
    
    df['MainBranch'] = df['Branch'].apply(get_main_program)
    
    # Extract SubBranch (stream)
    def get_sub_branch(branch):
        if '-' in branch:
            parts = branch.split('-', 1)
            return parts[1].strip() if len(parts) > 1 else ""
        return ""
    
    df['SubBranch'] = df['Branch'].apply(get_sub_branch)
    
    # If no subbranch, use main branch
    df.loc[df['SubBranch'] == "", 'SubBranch'] = df.loc[df['SubBranch'] == "", 'MainBranch']
    
    excel_data = {}
    
    # Group by MainBranch and Semester
    for (main_branch, semester), group_df in df.groupby(['MainBranch', 'Semester']):
        
        # Get all unique sub-branches (streams)
        sub_branches = sorted(group_df['SubBranch'].unique())
        
        # Create sheet name
        roman_sem = int_to_roman(semester)
        sheet_name = f"{main_branch}_Sem_{roman_sem}"
        
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
        
        # Get all unique exam dates
        all_dates = sorted(group_df['Exam Date'].unique(), key=lambda x: pd.to_datetime(x, format='%d-%m-%Y', errors='coerce'))
        
        processed_data = []
        
        for exam_date in all_dates:
            # Format date
            try:
                parsed_date = pd.to_datetime(exam_date, format='%d-%m-%Y', errors='coerce')
                if pd.notna(parsed_date):
                    formatted_date = parsed_date.strftime("%A, %d %B, %Y")
                else:
                    formatted_date = str(exam_date)
            except:
                formatted_date = str(exam_date)
            
            row_data = {'Exam Date': formatted_date}
            
            # For each sub-branch
            for sub_branch in sub_branches:
                subjects_on_date = group_df[
                    (group_df['Exam Date'] == exam_date) & 
                    (group_df['SubBranch'] == sub_branch)
                ]
                
                if not subjects_on_date.empty:
                    subjects = []
                    for _, row in subjects_on_date.iterrows():
                        subject_name = str(row['Module Description'])
                        module_code = str(row.get('Module Abbreviation', ''))
                        exam_time = str(row.get('Exam Time', ''))
                        cm_group = str(row.get('CM Group', '')).strip()
                        exam_slot = row.get('Exam Slot Number', 0)
                        
                        # Build subject display
                        subject_display = subject_name
                        
                        # Add module code if present
                        if module_code and module_code != 'nan':
                            subject_display = f"{subject_display} - ({module_code})"
                        
                        # Add CM Group prefix if present
                        if cm_group and cm_group != 'nan' and cm_group != '':
                            try:
                                cm_num = int(float(cm_group))
                                subject_display = f"[CM:{cm_num}] {subject_display}"
                            except:
                                subject_display = f"[CM:{cm_group}] {subject_display}"
                        
                        # Add exam time
                        if exam_time and exam_time != 'nan' and exam_time.strip():
                            subject_display = f"{subject_display} ({exam_time})"
                        
                        # Add slot number if present
                        if exam_slot and exam_slot != 0:
                            subject_display = f"{subject_display} [Slot {exam_slot}]"
                        
                        subjects.append(subject_display)
                    
                    # Join multiple subjects with line breaks
                    row_data[sub_branch] = "\n".join(subjects) if len(subjects) > 1 else subjects[0]
                else:
                    # No subjects for this stream on this date
                    row_data[sub_branch] = "---"
            
            processed_data.append(row_data)
        
        # Convert to DataFrame
        if processed_data:
            sheet_df = pd.DataFrame(processed_data)
            
            # Reorder columns to have Exam Date first, then the streams in order
            column_order = ['Exam Date'] + sub_branches
            sheet_df = sheet_df[column_order]
            
            # Fill any missing cells with "---"
            sheet_df = sheet_df.fillna("---")
            
            excel_data[sheet_name] = sheet_df
    
    return excel_data

def generate_pdf_from_excel_data(excel_data, output_pdf):
    """Generate PDF from Excel data dictionary"""
    pdf = FPDF(orientation='L', unit='mm', format='A3')
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    sheets_processed = 0
    
    for sheet_name, sheet_df in excel_data.items():
        if sheet_df.empty:
            continue
        
        # Parse sheet name to get program and semester
        try:
            name_parts = sheet_name.split('_')
            sem_index = name_parts.index('Sem')
            program_parts = name_parts[:sem_index]
            program = '_'.join(program_parts) if program_parts else ''
            semester_roman = name_parts[sem_index + 1]
            
            program_norm = re.sub(r'[.\s]+', ' ', program).strip().upper()
            main_branch_full = BRANCH_FULL_FORM.get(program_norm, program_norm)
            header_content = {
                'main_branch_full': main_branch_full, 
                'semester_roman': semester_roman
            }
            
            # Get column structure
            fixed_cols = ["Exam Date"]
            stream_cols = [c for c in sheet_df.columns if c not in fixed_cols]
            
            if not stream_cols:
                continue
            
            # Use actual number of streams
            actual_stream_count = len(stream_cols)
            cols_to_print = fixed_cols + stream_cols
            
            # Collect all unique time slots for this sheet
            actual_time_slots = set()
            for col in stream_cols:
                for cell_value in sheet_df[col]:
                    if pd.notna(cell_value) and str(cell_value) != "---":
                        # Extract time from cell value (format: "Subject (time)")
                        time_match = re.findall(r'\(([^)]+)\)', str(cell_value))
                        for match in time_match:
                            # Check if it looks like a time
                            if any(time_str in match for time_str in ['AM', 'PM', 'am', 'pm', ':']):
                                actual_time_slots.add(match)
            
            # Set column widths based on A3 size
            exam_date_width = 60
            remaining_width = pdf.w - 20 - exam_date_width
            stream_width = remaining_width / actual_stream_count if actual_stream_count > 0 else remaining_width
            col_widths = [exam_date_width] + [stream_width] * actual_stream_count
            
            # Add page and print table with actual streams
            pdf.add_page()
            print_table_custom(
                pdf, sheet_df, cols_to_print, col_widths, 
                line_height=10, header_content=header_content, 
                branches=stream_cols, actual_time_slots=actual_time_slots
            )
            
            sheets_processed += 1
            
        except Exception as e:
            st.warning(f"Error processing sheet {sheet_name}: {e}")
            continue
    
    if sheets_processed == 0:
        st.error("No sheets were processed for PDF generation!")
        return False
    
    try:
        pdf.output(output_pdf)
        st.success(f"PDF generated successfully with {sheets_processed} pages")
        return True
    except Exception as e:
        st.error(f"Error saving PDF: {e}")
        return False

def main():
    st.markdown("""
    <div class="main-header">
        <h1>📅 Excel to PDF Timetable Converter</h1>
        <p>Convert Excel verification files to formatted PDF timetables</p>
    </div>
    """, unsafe_allow_html=True)

    # Initialize session state
    if 'pdf_data' not in st.session_state:
        st.session_state.pdf_data = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'selected_college' not in st.session_state:
        st.session_state.selected_college = 'SVKM\'s NMIMS University'

    # Sidebar for college selection
    with st.sidebar:
        st.header("⚙️ Settings")
        
        college_options = [
            'SVKM\'s NMIMS University',
            'MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING',
            'SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING',
            'Custom College Name'
        ]
        
        selected_college = st.selectbox(
            "Select College Name",
            college_options,
            index=0
        )
        
        if selected_college == 'Custom College Name':
            custom_college = st.text_input("Enter Custom College Name", "")
            if custom_college:
                st.session_state.selected_college = custom_college
            else:
                st.session_state.selected_college = 'SVKM\'s NMIMS University'
        else:
            st.session_state.selected_college = selected_college

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("""
        <div class="upload-section">
            <h3>📁 Upload Excel Verification File</h3>
            <p>Upload your verification Excel file with Exam Date and Exam Time columns</p>
        </div>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload the Excel verification file containing exam dates and times"
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
                <li>📊 Direct Excel to PDF conversion</li>
                <li>📅 Preserves exam dates and times</li>
                <li>🎯 Program and stream grouping</li>
                <li>📝 Professional PDF formatting</li>
                <li>🏫 Customizable college header</li>
                <li>📋 Automatic page management</li>
                <li>⚡ Supports CM Groups & Slots</li>
                <li>📱 Mobile-friendly interface</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    if uploaded_file is not None:
        if st.button("🔄 Convert to PDF", type="primary", use_container_width=True):
            with st.spinner("Converting your Excel file to PDF... Please wait..."):
                try:
                    # Read the Excel file
                    df = read_verification_excel(uploaded_file)
                    
                    if df is not None:
                        st.write(f"📊 Processing {len(df)} records from Excel file...")
                        
                        # Show preview of data
                        with st.expander("📋 Data Preview (First 5 rows)"):
                            st.dataframe(df.head())
                        
                        # Create Excel sheets format for PDF
                        excel_data = create_excel_sheets_for_pdf(df)
                        
                        if excel_data:
                            st.write(f"📋 Created {len(excel_data)} sheets for PDF generation")
                            
                            # Generate PDF
                            temp_pdf_path = "temp_timetable_conversion.pdf"
                            
                            if generate_pdf_from_excel_data(excel_data, temp_pdf_path):
                                # Read the generated PDF
                                if os.path.exists(temp_pdf_path):
                                    with open(temp_pdf_path, "rb") as f:
                                        st.session_state.pdf_data = f.read()
                                    os.remove(temp_pdf_path)
                                    
                                    st.session_state.processing_complete = True
                                    
                                    st.markdown('<div class="status-success">🎉 PDF conversion completed successfully!</div>',
                                              unsafe_allow_html=True)
                                    
                                    # Show statistics
                                    total_records = len(df)
                                    unique_programs = df['Program'].nunique()
                                    unique_streams = df['Stream'].nunique() 
                                    unique_sessions = df['Current Session'].nunique()
                                    unique_dates = df['Exam Date'].nunique()
                                    
                                    st.success(f"📊 Conversion Summary:")
                                    st.info(f"• Total Records: {total_records}")
                                    st.info(f"• Programs: {unique_programs}")
                                    st.info(f"• Streams: {unique_streams}")
                                    st.info(f"• Sessions: {unique_sessions}")
                                    st.info(f"• Unique Exam Dates: {unique_dates}")
                                    st.info(f"• PDF Sheets Generated: {len(excel_data)}")
                                else:
                                    st.error("PDF file was not created successfully")
                            else:
                                st.error("Failed to generate PDF")
                        else:
                            st.error("No valid data found for PDF generation")
                    else:
                        st.error("Failed to read Excel file. Please check the format and required columns.")
                        
                except Exception as e:
                    st.error(f"An error occurred during conversion: {str(e)}")
                    import traceback
                    st.error(f"Full error details: {traceback.format_exc()}")

    # Display results and download option
    if st.session_state.processing_complete and st.session_state.pdf_data:
        st.markdown("---")
        
        st.markdown("### 📥 Download PDF")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.download_button(
                label="📄 Download PDF Timetable",
                data=st.session_state.pdf_data,
                file_name=f"timetable_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        
        with col2:
            if st.button("🔄 Convert Another File", use_container_width=True):
                st.session_state.processing_complete = False
                st.session_state.pdf_data = None
                st.rerun()
        
        with col3:
            if st.session_state.pdf_data:
                pdf_size = len(st.session_state.pdf_data) / 1024
                st.metric("PDF Size", f"{pdf_size:.1f} KB")
        
        st.markdown("### 👁️ Conversion Details")
        
        if uploaded_file is not None:
            try:
                preview_df = read_verification_excel(uploaded_file)
                if preview_df is not None:
                    st.markdown("#### 📊 Data Organization")
                    
                    grouping_info = preview_df.groupby(['Program', 'Stream', 'Current Session']).size().reset_index(name='Count')
                    st.dataframe(grouping_info, use_container_width=True)
                    
                    st.markdown("#### 📅 Exam Date Distribution")
                    
                    date_dist = preview_df['Exam Date'].value_counts().sort_index()
                    st.bar_chart(date_dist)
                    
            except Exception as e:
                st.warning(f"Could not generate preview: {str(e)}")
    
    # Instructions and help
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        #### 📌 Required Excel Columns:
        - **Program**: B TECH, M TECH, etc.
        - **Stream**: IT, COMPUTER, etc. 
        - **Current Session**: Sem 1, Sem 2, etc.
        - **Module Description**: Subject name
        - **Exam Date**: Date format (DD-MM-YYYY)
        - **Exam Time**: Time of examination
        """)
    
    with col2:
        st.markdown("""
        #### ✨ Optional Columns:
        - **Module Abbreviation**: Subject code
        - **CM Group**: Common module group number
        - **Exam Slot Number**: Slot identifier
        - **Student count**: Number of students
        - **Campus**: Campus name
        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; padding: 2rem; color: #666;">
        <p><strong>📄 Excel to PDF Timetable Converter</strong></p>
        <p>Convert verification Excel files to professionally formatted PDF timetables</p>
        <p style="font-size: 0.9em;">Direct conversion • Professional formatting • Custom branding • Automatic organization</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
