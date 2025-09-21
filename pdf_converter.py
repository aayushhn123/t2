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
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')
    
    # Add header
    header_height = 85
    pdf.set_y(0)
    current_date = datetime.now().strftime("%A, %B %d, %Y, %I:%M %p IST")
    pdf.set_font("Arial", size=14)
    text_width = pdf.get_string_width(current_date)
    x = pdf.w - 10 - text_width
    pdf.set_xy(x, 5)
    pdf.cell(text_width, 10, f"Generated on: {current_date}", 0, 0, 'R')
    
    # Add logo if exists
    if os.path.exists(LOGO_PATH):
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
    
    # Print header row
    pdf.set_font("Arial", size=12)
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    
    # Print data rows
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if not any(cell.strip() for cell in row):
            continue
            
        # Check if new page is needed
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
            # Re-add header to new page
            add_header_to_page(pdf, current_date, header_content, branches, time_slot)
            pdf.set_font("Arial", size=12)
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
        
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

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

def convert_semester_string_to_roman(semester_str):
    """Convert semester string to Roman numeral"""
    # Extract number from semester string
    if pd.isna(semester_str):
        return "I"
    
    semester_str = str(semester_str).strip()
    
    # Handle different semester formats
    semester_mappings = {
        "Sem I": "I", "Sem II": "II", "Sem III": "III", "Sem IV": "IV",
        "Sem V": "V", "Sem VI": "VI", "Sem VII": "VII", "Sem VIII": "VIII",
        "1": "I", "2": "II", "3": "III", "4": "IV", "5": "V", "6": "VI", 
        "7": "VII", "8": "VIII"
    }
    
    return semester_mappings.get(semester_str, "I")

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
    """Read the verification Excel file and process it"""
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Display actual columns for debugging
        st.write("Columns found in Excel file:")
        st.write(list(df.columns))
        
        # Column mapping for different possible names
        column_mapping = {
            "School": ["School", "School Name"],
            "Campus": ["Campus", "Campus Name"],
            "Program": ["Program", "Programme"],
            "Stream": ["Stream", "Specialization"],
            "CurrentAcademicYear": ["Current Academic Year", "Academic Year"],
            "CurrentSession": ["Current Session", "Session", "Semester"],
            "ModuleAbbreviation": ["Module Abbreviation", "Module Code", "Code"],
            "ModuleDescription": ["Module Description", "Subject Name", "Description"],
            "CommonAcrossSems": ["Common across sems", "Common Across Sems"],
            "DifficultyScore": ["Difficulty Score", "Difficulty"],
            "IsCommon": ["Is Common", "IsCommon"],
            "Category": ["Category"],
            "OE": ["OE"],
            "ExamMode": ["Exam mode", "Exam Mode"],
            "ExamDuration": ["Exam Duration", "Duration"],
            "StudentCount": ["Student count", "StudentCount", "Count"],
            "ExamDate": ["Exam Date"],
            "ExamTime": ["Exam Time"]
        }
        
        # Apply column mapping
        actual_columns = {}
        for standard_name, possible_names in column_mapping.items():
            for possible_name in possible_names:
                if possible_name in df.columns:
                    actual_columns[standard_name] = possible_name
                    break
        
        # Rename columns to standard names
        rename_dict = {v: k for k, v in actual_columns.items()}
        df = df.rename(columns=rename_dict)
        
        # Ensure required columns exist
        required_columns = ["Program", "Stream", "CurrentSession", "ModuleDescription", "ExamDate", "ExamTime"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Missing required columns: {missing_columns}")
            return None
        
        # Filter out rows with no exam date
        df = df[df['ExamDate'].notna() & (df['ExamDate'].str.strip() != "")]
        
        if df.empty:
            st.error("No valid exam data found in the file")
            return None
        
        return df
        
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None

def create_excel_sheets_for_pdf(df):
    """Convert the verification data into the Excel sheet format expected by PDF generator"""
    
    # Group by Program, Stream, and Session
    grouped = df.groupby(['Program', 'Stream', 'CurrentSession'])
    
    excel_data = {}
    
    for (program, stream, session), group_df in grouped:
        # Create sheet name
        sheet_name = f"{program}_{stream}_Sem_{convert_semester_string_to_roman(session)}"
        
        # Process the data for this sheet
        processed_data = []
        
        # Group by exam date
        date_groups = group_df.groupby('ExamDate')
        
        for exam_date, date_group in date_groups:
            # Format the exam date
            try:
                # Try to parse the date and format it consistently
                parsed_date = pd.to_datetime(exam_date, errors='coerce')
                if pd.notna(parsed_date):
                    formatted_date = parsed_date.strftime("%d-%m-%Y")
                else:
                    formatted_date = str(exam_date)
            except:
                formatted_date = str(exam_date)
            
            # Create subject display with exam time
            subjects = []
            for _, row in date_group.iterrows():
                subject_name = str(row.get('ModuleDescription', ''))
                module_code = str(row.get('ModuleAbbreviation', ''))
                exam_time = str(row.get('ExamTime', ''))
                oe_type = str(row.get('OE', '')) if pd.notna(row.get('OE', '')) else ""
                
                # Create subject display
                if module_code and module_code != 'nan':
                    subject_display = f"{subject_name} - ({module_code})"
                else:
                    subject_display = subject_name
                
                # Add OE type if present
                if oe_type and oe_type != 'nan':
                    subject_display = f"{subject_display} [{oe_type}]"
                
                # Add exam time if present and different from default
                if exam_time and exam_time != 'nan' and exam_time.strip():
                    subject_display = f"{subject_display} [{exam_time}]"
                
                subjects.append(subject_display)
            
            # Create row data
            row_data = {
                'Exam Date': formatted_date,
                stream: ", ".join(subjects)  # Use stream as column name
            }
            processed_data.append(row_data)
        
        # Convert to DataFrame
        if processed_data:
            sheet_df = pd.DataFrame(processed_data)
            # Fill missing values with "---"
            sheet_df = sheet_df.fillna("---")
            excel_data[sheet_name] = sheet_df
    
    return excel_data

def generate_pdf_from_excel_data(excel_data, output_pdf):
    """Generate PDF from Excel data dictionary"""
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    pdf.alias_nb_pages()
    
    sheets_processed = 0
    
    for sheet_name, sheet_df in excel_data.items():
        if sheet_df.empty:
            continue
        
        # Parse sheet name to get program, stream, and semester
        try:
            name_parts = sheet_name.split('_')
            if len(name_parts) >= 4 and name_parts[-2] == "Sem":
                program = name_parts[0]
                stream = "_".join(name_parts[1:-2])  # Handle multi-word streams
                semester_roman = name_parts[-1]
                
                main_branch_full = BRANCH_FULL_FORM.get(program, program)
                header_content = {
                    'main_branch_full': main_branch_full, 
                    'semester_roman': semester_roman
                }
                
                # Get column structure
                fixed_cols = ["Exam Date"]
                stream_cols = [c for c in sheet_df.columns if c not in fixed_cols]
                
                if not stream_cols:
                    continue
                
                cols_to_print = fixed_cols + stream_cols
                
                # Set column widths
                exam_date_width = 60
                remaining_width = pdf.w - 2 * pdf.l_margin - exam_date_width
                stream_width = remaining_width / len(stream_cols) if stream_cols else 0
                col_widths = [exam_date_width] + [stream_width] * len(stream_cols)
                
                # Convert exam dates to proper format for display
                try:
                    sheet_df["Exam Date"] = pd.to_datetime(
                        sheet_df["Exam Date"], format="%d-%m-%Y", errors='coerce'
                    ).dt.strftime("%A, %d %B, %Y")
                except:
                    pass  # Keep original format if conversion fails
                
                # Add page and print table
                pdf.add_page()
                print_table_custom(
                    pdf, sheet_df, cols_to_print, col_widths, 
                    line_height=10, header_content=header_content, 
                    branches=stream_cols
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
                                    os.remove(temp_pdf_path)  # Clean up temp file
                                    
                                    st.session_state.processing_complete = True
                                    
                                    st.markdown('<div class="status-success">🎉 PDF conversion completed successfully!</div>',
                                              unsafe_allow_html=True)
                                    
                                    # Show statistics
                                    total_records = len(df)
                                    unique_programs = df['Program'].nunique()
                                    unique_streams = df['Stream'].nunique() 
                                    unique_sessions = df['CurrentSession'].nunique()
                                    unique_dates = df['ExamDate'].nunique()
                                    
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
                # Clear session state and rerun
                st.session_state.processing_complete = False
                st.session_state.pdf_data = None
                st.rerun()
        
        with col3:
            # Show file size
            if st.session_state.pdf_data:
                pdf_size = len(st.session_state.pdf_data) / 1024  # Size in KB
                st.metric("PDF Size", f"{pdf_size:.1f} KB")
        
        # Preview section
        st.markdown("### 👁️ Conversion Details")
        
        if uploaded_file is not None:
            try:
                # Re-read the file for display purposes
                preview_df = read_verification_excel(uploaded_file)
                if preview_df is not None:
                    # Show data grouping information
                    st.markdown("#### 📊 Data Organization")
                    
                    grouping_info = preview_df.groupby(['Program', 'Stream', 'CurrentSession']).size().reset_index(name='Count')
                    st.dataframe(grouping_info, use_container_width=True)
                    
                    # Show exam date distribution
                    st.markdown("#### 📅 Exam Date Distribution")
                    
                    date_dist = preview_df['ExamDate'].value_counts().sort_index()
                    st.bar_chart(date_dist)
                    
            except Exception as e:
                st.warning(f"Could not generate preview: {str(e)}")
    
    # Instructions and help
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        #### 📁 Required Excel Columns:
        - **Program**: B TECH, M TECH, etc.
        - **Stream**: IT, COMPUTER, etc. 
        - **Current Session**: Sem I, Sem II, etc.
        - **Module Description**: Subject name
        - **Exam Date**: Date of examination
        - **Exam Time**: Time of examination
        """)
    
    with col2:
        st.markdown("""
        #### ✨ Optional Columns:
        - **Module Abbreviation**: Subject code
        - **OE**: Open elective type (OE1, OE2, etc.)
        - **School**: School name
        - **Campus**: Campus name
        - **Category**: Subject category
        - **Student count**: Number of students
        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; padding: 2rem; color: #666;">
        <p><strong>📄 Excel to PDF Timetable Converter</strong></p>
        <p>Convert verification Excel files to professionally formatted PDF timetables</p>
        <p style="font-size: 0.9em;">Direct conversion • Professional formatting • School branding • Automatic organization</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
