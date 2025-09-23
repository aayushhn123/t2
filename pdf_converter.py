import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import os
import io
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
    "B.TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS",
    "DIPLOMA": "DIPLOMA IN ENGINEERING"
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
    header_bg_color = (149, 28, 33)
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

def add_header_to_page(pdf, current_date, header_content, branches, time_slot=None):
    """Add header to a new page"""
    pdf.set_y(0)
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')

    if os.path.exists(LOGO_PATH):
        logo_width = 45
        logo_x = (pdf.w - logo_width) / 2
        pdf.image(LOGO_PATH, x=logo_x, y=10, w=logo_width)

    pdf.set_fill_color(149, 28, 33)
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

    # Add page numbers
    pdf.set_font("Arial", size=14)
    pdf.set_text_color(0, 0, 0)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')

    # Add header
    add_header_to_page(pdf, datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST"), header_content, branches, time_slot)

    # Print header row
    pdf.set_font("Arial", size=12)
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)

    # Print data rows
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
            pdf.add_page()
            add_header_to_page(pdf, datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST"), header_content, branches, time_slot)
            pdf.set_font("Arial", size=12)
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)

        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def int_to_roman(num):
    """Convert integer to Roman numeral"""
    if not isinstance(num, int) or num < 1:
        return "I"
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

def convert_semester_string_to_roman(semester_str):
    """Convert semester string to Roman numeral"""
    if pd.isna(semester_str):
        return "I"
    semester_str = str(semester_str).strip()
    semester_mappings = {
        "Sem I": "I", "Sem II": "II", "Sem III": "III", "Sem IV": "IV",
        "Sem V": "V", "Sem VI": "VI", "Sem VII": "VII", "Sem VIII": "VIII",
        "Sem IX": "IX", "Sem X": "X", "Sem XI": "XI", "Sem XII": "XII",
        "1": "I", "2": "II", "3": "III", "4": "IV", "5": "V", "6": "VI",
        "7": "VII", "8": "VIII", "9": "IX", "10": "X", "11": "XI", "12": "XII"
    }
    match = re.search(r'\d+', semester_str)
    if match:
        num = int(match.group())
        return int_to_roman(num)
    return semester_mappings.get(semester_str, "I")

def read_verification_excel(uploaded_file):
    """Read the verification Excel file and process it"""
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Column mapping for different possible names
        column_mapping = {
            "School Name": ["School Name", "School"],
            "Program": ["Program", "Programme"],
            "Stream": ["Stream", "Specialization", "SubBranch"],
            "Current Session": ["Current Session", "Session", "Semester"],
            "Module Description": ["Module Description", "Subject", "Subject Name", "Description"],
            "Module Abbreviation": ["Module Abbreviation", "Module Code", "Code"],
            "Exam Date": ["Exam Date", "Date"],
            "Exam Time": ["Exam Time", "Time", "Time Slot"],
            "Category": ["Category"],
            "OE": ["OE", "Open Elective"],
            "Common across sems": ["Common across sems", "Common Across Sems"],
            "Is Common": ["Is Common", "IsCommon"],
            "Exam Mode": ["Exam Mode", "Exam mode"],
            "Exam Duration": ["Exam Duration", "Duration"],
            "Student Count": ["Student Count", "Student count", "Count"]
        }
        
        # Apply column mapping
        actual_columns = {}
        for standard_name, possible_names in column_mapping.items():
            for possible_name in possible_names:
                if possible_name in df.columns:
                    actual_columns[standard_name] = possible_name
                    break
                for col in df.columns:
                    if col.lower() == possible_name.lower():
                        actual_columns[standard_name] = col
                        break
        
        # Rename columns to standard names
        rename_dict = {v: k for k, v in actual_columns.items()}
        df = df.rename(columns=rename_dict)
        
        # Ensure required columns exist
        required_columns = ["Program", "Stream", "Current Session", "Module Description", "Exam Date", "Exam Time"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Missing required columns: {', '.join(missing_columns)}")
            return None
        
        # Parse Exam Date
        def parse_date(date_str):
            if pd.isna(date_str):
                return pd.NaT
            try:
                return pd.to_datetime(date_str, format="%A, %d %B, %Y", errors='coerce')
            except:
                return pd.to_datetime(date_str, errors='coerce')
        
        df['Exam Date'] = df['Exam Date'].apply(parse_date)
        
        # Normalize Exam Time
        def normalize_time(time_str):
            if pd.isna(time_str):
                return ""
            time_str = str(time_str).strip().replace('.', ':').replace('pm', ' PM').replace('am', ' AM')
            time_str = re.sub(r'noon', 'PM', time_str, flags=re.IGNORECASE)
            return time_str
        
        df['Exam Time'] = df['Exam Time'].apply(normalize_time)
        
        # Add derived columns
        if 'School Name' in df.columns:
            df['Branch'] = df['School Name'] + " " + df['Program']
        else:
            df['Branch'] = df['Program']
        df['MainBranch'] = df['Program'].str.split(',', expand=True)[0].str.strip()
        df['SubBranch'] = df['Stream']
        df['Semester'] = df['Current Session'].str.extract('(\d+)', expand=False).astype(float).astype(int)
        df['Subject'] = df['Module Description']
        df['Time Slot'] = df['Exam Time']
        df['OE'] = df.get('OE', '').fillna('')
        df['Category'] = df.get('Category', '').fillna('')
        if 'Common across sems' not in df.columns:
            df['Common across sems'] = False
        
        return df
    
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None

def create_excel_sheets_for_pdf(df):
    """Convert the verification data into grouped sheets with maximum 4 branches per page"""
    # Separate electives and non-electives
    df_non_elec = df[df['Category'] != 'INTD'].copy()
    df_elec = df[df['Category'] == 'INTD'].copy()
    
    excel_data = {}
    
    # Process non-electives
    if not df_non_elec.empty:
        grouped = df_non_elec.groupby(['MainBranch', 'Semester'])
        for (main_branch, sem), group_df in grouped:
            all_streams = sorted(group_df['SubBranch'].unique())
            max_streams_per_page = 4
            stream_chunks = [all_streams[i:i + max_streams_per_page] for i in range(0, len(all_streams), max_streams_per_page)]
            
            for chunk_idx, stream_chunk in enumerate(stream_chunks):
                sheet_name = f"{main_branch}_Sem_{int_to_roman(sem)}"
                if len(stream_chunks) > 1:
                    sheet_name += f"_Part{chunk_idx + 1}"
                
                processed_data = []
                all_exam_dates = sorted(group_df['Exam Date'].unique(), na_position='last')
                
                for exam_date in all_exam_dates:
                    date_str = exam_date.strftime("%A, %d %B, %Y") if pd.notna(exam_date) else "Unknown Date"
                    row_data = {'Exam Date': date_str}
                    
                    for stream in stream_chunk:
                        stream_subjects = group_df[
                            (group_df['Exam Date'] == exam_date) & 
                            (group_df['SubBranch'] == stream)
                        ]
                        if not stream_subjects.empty:
                            subjects = []
                            for _, row in stream_subjects.iterrows():
                                subject_display = row['Subject']
                                if pd.notna(row.get('Module Abbreviation', '')):
                                    subject_display += f" - ({row['Module Abbreviation']})"
                                if pd.notna(row['Time Slot']) and row['Time Slot'].strip():
                                    subject_display += f" [{row['Time Slot']}]"
                                subjects.append(subject_display)
                            row_data[stream] = ", ".join(subjects)
                        else:
                            row_data[stream] = "---"
                    
                    processed_data.append(row_data)
                
                if processed_data:
                    sheet_df = pd.DataFrame(processed_data)
                    for stream in stream_chunk:
                        if stream not in sheet_df.columns:
                            sheet_df[stream] = "---"
                    column_order = ['Exam Date'] + sorted(stream_chunk)
                    sheet_df = sheet_df[column_order]
                    excel_data[sheet_name] = {'df': sheet_df, 'type': 'core'}
    
    # Process electives
    if not df_elec.empty:
        grouped = df_elec.groupby(['MainBranch', 'Semester'])
        for (main_branch, sem), group_df in grouped:
            sheet_name = f"{main_branch}_Sem_{int_to_roman(sem)}_Electives"
            processed_data = []
            all_exam_dates = sorted(group_df['Exam Date'].unique(), na_position='last')
            
            for exam_date in all_exam_dates:
                date_str = exam_date.strftime("%A, %d %B, %Y") if pd.notna(exam_date) else "Unknown Date"
                for oe_type in group_df['OE'].unique():
                    if pd.isna(oe_type) or not oe_type.strip():
                        continue
                    oe_subjects = group_df[
                        (group_df['Exam Date'] == exam_date) & 
                        (group_df['OE'] == oe_type)
                    ]
                    if not oe_subjects.empty:
                        subjects = []
                        for _, row in oe_subjects.iterrows():
                            subject_display = row['Subject']
                            if pd.notna(row.get('Module Abbreviation', '')):
                                subject_display += f" - ({row['Module Abbreviation']})"
                            if pd.notna(row['Time Slot']) and row['Time Slot'].strip():
                                subject_display += f" [{row['Time Slot']}]"
                            subjects.append(subject_display)
                        processed_data.append({
                            'Exam Date': date_str,
                            'OE Type': oe_type,
                            'Subjects': ", ".join(subjects)
                        })
            
            if processed_data:
                sheet_df = pd.DataFrame(processed_data)
                column_order = ['Exam Date', 'OE Type', 'Subjects']
                sheet_df = sheet_df[column_order]
                excel_data[sheet_name] = {'df': sheet_df, 'type': 'elective'}
    
    return excel_data

def generate_pdf_from_excel_data(excel_data, output_pdf):
    """Generate PDF from Excel data dictionary"""
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    sheets_processed = 0
    
    for sheet_name, sheet_info in sorted(excel_data.items()):
        sheet_df = sheet_info['df']
        table_type = sheet_info['type']
        if sheet_df.empty:
            continue
        
        try:
            name_parts = sheet_name.split('_')
            main_branch = "_".join(name_parts[:-2]) if 'Electives' in name_parts else "_".join(name_parts[:-1])
            semester_roman = name_parts[-2] if 'Electives' in name_parts else name_parts[-1]
            if 'Part' in semester_roman:
                semester_roman = name_parts[-3]
            main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
            header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}
            
            if table_type == 'core':
                fixed_cols = ["Exam Date"]
                stream_cols = [c for c in sheet_df.columns if c not in fixed_cols]
                cols_to_print = fixed_cols + stream_cols
                exam_date_width = 60
                remaining_width = pdf.w - 2 * pdf.l_margin - exam_date_width
                stream_width = remaining_width / len(stream_cols) if stream_cols else 0
                col_widths = [exam_date_width] + [stream_width] * len(stream_cols)
                title = f"{main_branch_full} - Core Subjects"
            else:
                cols_to_print = ['Exam Date', 'OE Type', 'Subjects']
                col_widths = [(pdf.w - 2 * pdf.l_margin) / 3] * 3
                stream_cols = []
                title = f"{main_branch_full} - Open Electives"
            
            pdf.add_page()
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, txt=title, ln=1, align='C')
            print_table_custom(
                pdf, sheet_df, cols_to_print, col_widths,
                line_height=10, header_content=header_content,
                branches=stream_cols
            )
            sheets_processed += 1
        
        except Exception as e:
            st.warning(f"Error processing sheet {sheet_name}: {str(e)}")
            continue
    
    if sheets_processed == 0:
        st.error("No sheets were processed for PDF generation!")
        return False
    
    try:
        pdf.output(output_pdf)
        return True
    except Exception as e:
        st.error(f"Error saving PDF: {str(e)}")
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
                <li>🎯 Up to 4 branches per page</li>
                <li>📝 Professional PDF formatting</li>
                <li>🏫 School header and branding</li>
                <li>📋 Automatic page management</li>
                <li>⚡ No scheduling logic applied</li>
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
                    if df is None:
                        st.error("Failed to process the Excel file. Please check the file format and required columns.")
                        return

                    # Create sheets with max 4 branches
                    excel_data = create_excel_sheets_for_pdf(df)
                    if not excel_data:
                        st.error("No valid data to process for PDF generation.")
                        return

                    # Generate PDF
                    pdf_output = io.BytesIO()
                    temp_pdf_path = "temp_timetable.pdf"
                    success = generate_pdf_from_excel_data(excel_data, temp_pdf_path)
                    if success:
                        with open(temp_pdf_path, "rb") as f:
                            pdf_output.write(f.read())
                        pdf_output.seek(0)
                        if os.path.exists(temp_pdf_path):
                            os.remove(temp_pdf_path)

                        # Store PDF data in session state
                        st.session_state.pdf_data = pdf_output.getvalue()
                        st.session_state.processing_complete = True

                        st.markdown('<div class="status-success">🎉 PDF generated successfully!</div>', unsafe_allow_html=True)

                        # Show statistics
                        total_records = len(df)
                        unique_programs = df['Program'].nunique()
                        unique_streams = df['Stream'].nunique()
                        unique_sessions = df['Current Session'].nunique()
                        unique_dates = df['Exam Date'].nunique()
                        
                        st.success("📊 Conversion Summary:")
                        st.info(f"• Total Records: {total_records}")
                        st.info(f"• Programs: {unique_programs}")
                        st.info(f"• Streams: {unique_streams}")
                        st.info(f"• Sessions: {unique_sessions}")
                        st.info(f"• Unique Exam Dates: {unique_dates}")

                except Exception as e:
                    st.markdown(f'<div class="status-error">❌ An error occurred: {str(e)}</div>', unsafe_allow_html=True)

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
            pdf_size = len(st.session_state.pdf_data) / 1024
            st.metric("PDF Size", f"{pdf_size:.1f} KB")

    # Instructions and help
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        #### 📁 Required Excel Columns:
        - **Program**: B.TECH, M TECH, etc.
        - **Stream**: IT, COMPUTER, etc.
        - **Current Session**: Sem I, Sem II, etc.
        - **Module Description**: Subject name
        - **Exam Date**: Date of examination (e.g., Tuesday, 25 November, 2025)
        - **Exam Time**: Time of examination (e.g., 2:00 PM - 5:00 PM)
        """)
    
    with col2:
        st.markdown("""
        #### ✨ Optional Columns:
        - **School Name**: School name
        - **Module Abbreviation**: Subject code
        - **OE**: Open elective type (OE1, OE2, etc.)
        - **Category**: Subject category (e.g., INTD for electives)
        - **Student Count**: Number of students
        """)

    # Display footer
    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <p>🎓 <strong>Excel to PDF Converter</strong></p>
        <p>Developed for Timetable Conversion</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
