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
    "MCA": "MASTER OF COMPUTER APPLICATIONS"
}

# Define logo path (adjust as needed for your environment)
LOGO_PATH = "logo.png"  # Ensure this path is valid in your environment

def read_timetable(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        
        # Standardize column names
        column_mapping = {
            'School Name': ['School Name', 'School'],
            'Program': ['Program', 'Programme'],
            'Stream': ['Stream', 'Branch', 'SubBranch'],
            'Current Session': ['Current Session', 'Session', 'Semester'],
            'Module Description': ['Module Description', 'Subject', 'Course'],
            'Exam Date': ['Exam Date', 'Date'],
            'Exam Time': ['Exam Time', 'Time Slot', 'Time'],
            'Category': ['Category'],
            'OE': ['OE', 'Open Elective']
        }
        
        # Map actual column names to standard names
        standard_columns = {}
        for standard_name, possible_names in column_mapping.items():
            for name in possible_names:
                if name in df.columns:
                    standard_columns[standard_name] = name
                    break
                # Also check for case-insensitive matches
                for col in df.columns:
                    if col.lower() == name.lower():
                        standard_columns[standard_name] = col
                        break
        
        # Rename columns to standard names
        df = df.rename(columns={v: k for k, v in standard_columns.items()})
        
        # Check for required columns
        required_columns = ['Program', 'Stream', 'Current Session', 'Module Description', 'Exam Date', 'Exam Time']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Missing required columns: {', '.join(missing_columns)}")
            return None
        
        # Normalize Exam Date
        def parse_date(date_str):
            if pd.isna(date_str):
                return pd.NaT
            try:
                # Handle formats like "Tuesday, 25 November, 2025"
                return pd.to_datetime(date_str, format="%A, %d %B, %Y", errors='coerce')
            except:
                # Fallback to general parsing
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
        
        # Separate electives (INTD category or OE not empty)
        df['Category'] = df.get('Category', '').fillna('')
        df['OE'] = df.get('OE', '').fillna('')
        df_ele = df[df['Category'] == 'INTD'].copy()
        df_non_elec = df[df['Category'] != 'INTD'].copy()
        
        # Add Branch as School Name + Program (if School Name exists)
        if 'School Name' in df.columns:
            df_non_elec['Branch'] = df_non_elec['School Name'] + " " + df_non_elec['Program']
            df_ele['Branch'] = df_ele['School Name'] + " " + df_ele['Program']
        else:
            df_non_elec['Branch'] = df_non_elec['Program']
            df_ele['Branch'] = df_ele['Program']
        
        # Extract MainBranch and SubBranch
        df_non_elec['MainBranch'] = df_non_elec['Program'].str.split(',', expand=True)[0].str.strip()
        df_ele['MainBranch'] = df_ele['Program'].str.split(',', expand=True)[0].str.strip()
        df_non_elec['SubBranch'] = df_non_elec['Stream']
        df_ele['SubBranch'] = df_ele['Stream']
        
        # Extract Semester number
        df_non_elec['Semester'] = df_non_elec['Current Session'].str.extract('(\d+)', expand=False).astype(float).astype(int)
        df_ele['Semester'] = df_ele['Current Session'].str.extract('(\d+)', expand=False).astype(float).astype(int)
        
        # Set 'Subject' to 'Module Description'
        df_non_elec['Subject'] = df_non_elec['Module Description']
        df_ele['Subject'] = df_ele['Module Description']
        
        # Use 'Exam Time' as 'Time Slot'
        df_non_elec['Time Slot'] = df_non_elec['Exam Time']
        df_ele['Time Slot'] = df_ele['Exam Time']
        
        # Add 'Common across sems' if missing
        if 'Common across sems' not in df_non_elec.columns:
            df_non_elec['Common across sems'] = False
        if 'Common across sems' not in df_ele.columns:
            df_ele['Common across sems'] = False
        
        return pd.concat([df_non_elec, df_ele], ignore_index=True)
    
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None

class PDF(FPDF):
    def header(self):
        if os.path.exists(LOGO_PATH):
            self.image(LOGO_PATH, 10, 8, 33)
        self.set_font('Arial', 'B', 15)
        self.cell(80)
        self.cell(30, 10, 'Exam Timetable', 0, 0, 'C')
        self.ln(20)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')

def generate_pdf_timetable(sem_dict, pdf_path):
    pdf = PDF()
    
    for sem in sorted(sem_dict.keys()):
        df_sem = sem_dict[sem]
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt=f"Semester {sem}", ln=1, align='C')
        
        for main_branch in df_sem["MainBranch"].unique():
            main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
            df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
            
            if not df_mb.empty:
                # Separate non-electives and electives
                df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()
                
                # Non-electives
                if not df_non_elec.empty:
                    pdf.set_font("Arial", 'B', 12)
                    pdf.cell(200, 10, txt=f"{main_branch_full} - Core Subjects", ln=1)
                    
                    df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], errors='coerce')
                    df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                    
                    pivot_df = df_non_elec.pivot_table(
                        index=["Exam Date", "Time Slot"],
                        columns="SubBranch",
                        values="Subject",
                        aggfunc=lambda x: ", ".join(x)
                    ).fillna("---")
                    
                    col_width = 190 / (len(pivot_df.columns) + 2)
                    pdf.set_font("Arial", size=10)
                    
                    # Headers
                    pdf.cell(col_width, 10, "Date", 1)
                    pdf.cell(col_width, 10, "Exam Time", 1)
                    for col in pivot_df.columns:
                        pdf.cell(col_width, 10, col, 1)
                    pdf.ln()
                    
                    for (date, slot), row in pivot_df.iterrows():
                        if pd.notna(date):
                            pdf.cell(col_width, 10, date.strftime("%d-%m-%Y"), 1)
                        else:
                            pdf.cell(col_width, 10, "Unknown Date", 1)
                        pdf.cell(col_width, 10, slot, 1)
                        for val in row:
                            pdf.cell(col_width, 10, val, 1)
                        pdf.ln()
                
                # Electives
                if not df_elec.empty:
                    pdf.set_font("Arial", 'B', 12)
                    pdf.cell(200, 10, txt=f"{main_branch_full} - Open Electives", ln=1)
                    
                    df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], errors='coerce')
                    df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                    
                    elec_pivot = df_elec.groupby(['OE', 'Exam Date', 'Time Slot'])['Subject'].apply(
                        lambda x: ", ".join(x)
                    ).reset_index()
                    
                    col_width = 190 / 4
                    pdf.set_font("Arial", size=10)
                    
                    # Headers
                    pdf.cell(col_width, 10, "OE Type", 1)
                    pdf.cell(col_width, 10, "Date", 1)
                    pdf.cell(col_width, 10, "Exam Time", 1)
                    pdf.cell(col_width, 10, "Subjects", 1)
                    pdf.ln()
                    
                    for _, row in elec_pivot.iterrows():
                        pdf.cell(col_width, 10, row['OE'], 1)
                        if pd.notna(row['Exam Date']):
                            pdf.cell(col_width, 10, row['Exam Date'].strftime("%d-%m-%Y"), 1)
                        else:
                            pdf.cell(col_width, 10, "Unknown Date", 1)
                        pdf.cell(col_width, 10, row['Time Slot'], 1)
                        pdf.cell(col_width, 10, row['Subject'], 1)
                        pdf.ln()
    
    pdf.output(pdf_path)

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
                    all_data = read_timetable(uploaded_file)
                    if all_data is None:
                        st.error("Failed to process the Excel file. Please check the file format and required columns.")
                        return

                    # Create semester dictionary
                    sem_dict = {}
                    for sem in sorted(all_data["Semester"].unique()):
                        sem_data = all_data[all_data["Semester"] == sem].copy()
                        sem_dict[sem] = sem_data

                    # Generate PDF
                    pdf_output = io.BytesIO()
                    temp_pdf_path = "temp_timetable.pdf"
                    generate_pdf_timetable(sem_dict, temp_pdf_path)
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
                    total_records = len(all_data)
                    unique_programs = all_data['Program'].nunique()
                    unique_streams = all_data['Stream'].nunique()
                    unique_sessions = all_data['Current Session'].nunique()
                    unique_dates = all_data['Exam Date'].nunique()
                    
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
        - **Program**: B TECH, M TECH, etc.
        - **Stream**: IT, COMPUTER, etc.
        - **Current Session**: Sem I, Sem II, etc.
        - **Module Description**: Subject name
        - **Exam Date**: Date of examination (e.g., Tuesday, 25 November, 2025)
        - **Exam Time**: Time of examination (e.g., 2:00 PM - 5:00 PM)
        """)
    
    with col2:
        st.markdown("""
        #### ✨ Optional Columns:
        - **Module Abbreviation**: Subject code
        - **OE**: Open elective type (OE1, OE2, etc.)
        - **School Name**: School name
        - **Category**: Subject category (e.g., INTD for electives)
        - **Student count**: Number of students
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
