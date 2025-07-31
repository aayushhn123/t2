import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import io
from fpdf import FPDF
import os
from PyPDF2 import PdfReader, PdfWriter

# Set page configuration
st.set_page_config(
    page_title="Re-exam Timetable Maker",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS from LATEST-FINAL EXAM - 35 days.py
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
    .metric-card {
        display: flex;
        flex-direction: column;
        align-items: center;
        padding: 1.5rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 0.5rem;
        transition: transform 0.2s;
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
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
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
        .metric-card {
            background: linear-gradient(135deg, #4a5db0 0%, #5a3e8a 100%);
        }
    }
</style>
""", unsafe_allow_html=True)

# Define RealTimeOptimizer class (adapted from provided code)
class RealTimeOptimizer:
    def __init__(self, branches, holidays):
        self.branches = branches
        self.holidays = holidays
        self.time_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        self.schedule_grid = {}
        self.log = []

    def add_exam_to_grid(self, date_str, time_slot, branch, subject):
        if date_str not in self.schedule_grid:
            self.schedule_grid[date_str] = {}
        if time_slot not in self.schedule_grid[date_str]:
            self.schedule_grid[date_str][time_slot] = {}
        self.schedule_grid[date_str][time_slot][branch] = subject
        self.log.append(f"Scheduled {subject} for {branch} on {date_str} at {time_slot}")

    def find_earliest_empty_slot(self, branch, start_date, preferred_time_slot):
        current_date = start_date
        max_attempts = 200
        attempts = 0
        while attempts < max_attempts:
            date_str = current_date.strftime("%d-%m-%Y")
            if current_date.weekday() != 6 and current_date.date() not in self.holidays:
                branch_has_exam = False
                if date_str in self.schedule_grid:
                    for slot in self.time_slots:
                        if (slot in self.schedule_grid[date_str] and 
                            branch in self.schedule_grid[date_str][slot] and 
                            self.schedule_grid[date_str][slot][branch] is not None):
                            branch_has_exam = True
                            break
                if not branch_has_exam:
                    return date_str, preferred_time_slot
            current_date += timedelta(days=1)
            attempts += 1
        return None, None

    def initialize_grid_with_empty_days(self, start_date, num_days=40):
        current_date = start_date
        for _ in range(num_days):
            if current_date.weekday() != 6 and current_date.date() not in self.holidays:
                date_str = current_date.strftime("%d-%m-%Y")
                if date_str not in self.schedule_grid:
                    self.schedule_grid[date_str] = {}
                    for slot in self.time_slots:
                        self.schedule_grid[date_str][slot] = {branch: None for branch in self.branches}
            current_date += timedelta(days=1)

# Helper functions
def roman_to_int(roman):
    roman_dict = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10, 'XI':11}
    return roman_dict.get(roman.upper(), 0)

def extract_roman(semester_str):
    match = re.search(r'Semester (I|II|III|IV|V|VI|VII|VIII|IX|X|XI)', semester_str, re.IGNORECASE)
    return match.group(1).upper() if match else None

def normalize_time_slot(time_str):
    return time_str.replace("to", "-").replace("pm", "PM").replace("am", "AM")

def int_to_roman(num):
    roman_values = [
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
    ]
    result = ""
    for value, numeral in roman_values:
        while num >= value:
            result += numeral
            num -= value
    return result

# Functions for processing
def read_final_exam_schedule(final_exam_file):
    required_cols = ['Program', 'Stream', 'Semester', 'SubjectName', 'Exam Date', 'Exam Time', 'Is Common']
    final_df = pd.read_excel(final_exam_file)
    missing_cols = [col for col in required_cols if col not in final_df.columns]
    if missing_cols:
        st.error(f"Missing columns in Final Exam Schedule: {', '.join(missing_cols)}")
        return None, None
    final_df['Branch'] = final_df['Program'] + '-' + final_df['Stream']
    common_subjects = {}
    non_common_subjects = {}
    for (subject, semester), group in final_df.groupby(['SubjectName', 'Semester']):
        if all(group['Is Common'] == "YES") and group['Exam Date'].nunique() == 1 and group['Exam Time'].nunique() == 1:
            common_subjects[(subject, semester)] = (group['Exam Date'].iloc[0], group['Exam Time'].iloc[0])
        else:
            for _, row in group.iterrows():
                branch = row['Branch']
                non_common_subjects[(subject, semester, branch)] = (row['Exam Date'], row['Exam Time'])
    return common_subjects, non_common_subjects

def read_re_exam_data(re_exam_file):
    required_cols = ['Program', 'Stream', 'Semester', 'Subject name']
    re_exam_df = pd.read_excel(re_exam_file)
    missing_cols = [col for col in required_cols if col not in re_exam_df.columns]
    if missing_cols:
        st.error(f"Missing columns in Re-exam Data: {', '.join(missing_cols)}")
        return None
    re_exam_df['Branch'] = re_exam_df['Program'] + '-' + re_exam_df['Stream']
    re_exam_df['Subject'] = re_exam_df['Subject name']
    re_exam_df['SemesterInt'] = re_exam_df['Semester'].apply(lambda x: roman_to_int(extract_roman(x)))
    return re_exam_df

def assign_matched_dates(re_exam_df, common_subjects, non_common_subjects, optimizer):
    for idx, row in re_exam_df.iterrows():
        subject = row['Subject']
        semester = row['SemesterInt']
        branch = row['Branch']
        if (subject, semester) in common_subjects:
            exam_date, exam_time = common_subjects[(subject, semester)]
            re_exam_df.at[idx, 'Exam Date'] = exam_date
            re_exam_df.at[idx, 'Exam Time'] = normalize_time_slot(exam_time)
            optimizer.add_exam_to_grid(exam_date, re_exam_df.at[idx, 'Exam Time'], branch, subject)
            optimizer.log.append(f"Matched {subject} for {branch} on {exam_date} from final schedule: {exam_time}")
        elif (subject, semester, branch) in non_common_subjects:
            exam_date, exam_time = non_common_subjects[(subject, semester, branch)]
            re_exam_df.at[idx, 'Exam Date'] = exam_date
            re_exam_df.at[idx, 'Exam Time'] = normalize_time_slot(exam_time)
            optimizer.add_exam_to_grid(exam_date, re_exam_df.at[idx, 'Exam Time'], branch, subject)
            optimizer.log.append(f"Matched {subject} for {branch} on {exam_date} from final schedule: {exam_time}")
        else:
            re_exam_df.at[idx, 'Exam Date'] = ""
            re_exam_df.at[idx, 'Exam Time'] = ""

def schedule_unmatched_subjects(re_exam_df, optimizer, base_date):
    unmatched = re_exam_df[re_exam_df['Exam Date'] == ""]
    for idx, row in unmatched.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        semester = row['SemesterInt']
        if semester % 2 != 0:
            preferred_slot = "10:00 AM - 1:00 PM"
        else:
            preferred_slot = "2:00 PM - 5:00 PM"
        date_str, time_slot = optimizer.find_earliest_empty_slot(branch, base_date, preferred_slot)
        if date_str:
            re_exam_df.at[idx, 'Exam Date'] = date_str
            re_exam_df.at[idx, 'Exam Time'] = time_slot
            optimizer.add_exam_to_grid(date_str, time_slot, branch, subject)
        else:
            st.error(f"Could not schedule {subject} for {branch} after 200 attempts")
            optimizer.log.append(f"Failed to schedule {subject} for {branch}")

def save_to_excel(re_exam_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for (program, semester), group in re_exam_df.groupby(['Program', 'SemesterInt']):
            sheet_df = group[['Subject', 'Exam Date', 'Exam Time']].copy()
            sheet_name = f"{program}_Sem_{int_to_roman(semester)}"
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None):
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 10, header_content, 0, 1, 'C')
    pdf.ln(5)
    for col, width in zip(columns, col_widths):
        pdf.cell(width, 10, col, 1, 0, 'C')
    pdf.ln()
    pdf.set_font("Arial", size=10)
    for _, row in df.iterrows():
        for col, width in zip(columns, col_widths):
            pdf.cell(width, 10, str(row[col]), 1, 0, 'C')
        pdf.ln()

def generate_pdf_timetable(re_exam_df):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=False, margin=15)
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=130, y=10, w=40)
    pdf.set_font("Arial", 'B', 16)
    pdf.set_fill_color(149, 33, 28)
    pdf.set_text_color(255, 255, 255)
    pdf.rect(10, 30, 277, 14, 'F')
    pdf.set_xy(10, 30)
    pdf.cell(277, 14, "MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING", 0, 1, 'C')
    pdf.set_text_color(0, 0, 0)
    
    for (program, semester), group in re_exam_df.groupby(['Program', 'SemesterInt']):
        pdf.add_page()
        header = f"{program} - Semester {int_to_roman(semester)} Re-exam Timetable"
        df = group[['Exam Date', 'Subject', 'Exam Time']]
        col_widths = [60, 157, 60]
        print_table_custom(pdf, df, ['Exam Date', 'Subject', 'Exam Time'], col_widths, header_content=header)
    
    pdf_output = io.BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)
    return pdf_output.getvalue()

def save_verification_excel(re_exam_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        re_exam_df[['Program', 'Stream', 'Semester', 'Subject name', 'Academic Year', 'Exam Date', 'Exam Time']].to_excel(writer, sheet_name="Verification", index=False)
    output.seek(0)
    return output

# Main function
def main():
    st.markdown("""
    <div class="main-header">
        <h1>📅 Re-exam Timetable Maker</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        base_date = st.date_input("Base Date for Re-exams", value=datetime(2025, 4, 1))
        base_date = datetime.combine(base_date, datetime.min.time())
        st.markdown("#### 📅 Holiday Configuration")
        holidays = [
            datetime(2025, 4, 14),
            datetime(2025, 5, 1),
            datetime(2025, 8, 15)
        ]
        custom_holidays = st.text_area("Custom Holidays (DD-MM-YYYY, one per line)")
        if custom_holidays:
            for line in custom_holidays.split('\n'):
                try:
                    holidays.append(datetime.strptime(line.strip(), "%d-%m-%Y"))
                except:
                    st.warning(f"Invalid date format: {line}")

    # Main interface
    st.markdown('<div class="upload-section"><h3>📁 Upload Files</h3></div>', unsafe_allow_html=True)
    re_exam_file = st.file_uploader("Re-exam Data Excel", type=['xlsx'])
    final_exam_file = st.file_uploader("Final Exam Schedule Excel", type=['xlsx'])

    if re_exam_file and final_exam_file:
        if st.button("Generate Re-exam Timetable"):
            with st.spinner("Processing..."):
                common_subjects, non_common_subjects = read_final_exam_schedule(final_exam_file)
                if common_subjects is None:
                    return
                re_exam_df = read_re_exam_data(re_exam_file)
                if re_exam_df is None:
                    return
                
                optimizer = RealTimeOptimizer(re_exam_df['Branch'].unique(), set(h.date() for h in holidays))
                optimizer.initialize_grid_with_empty_days(base_date)
                assign_matched_dates(re_exam_df, common_subjects, non_common_subjects, optimizer)
                schedule_unmatched_subjects(re_exam_df, optimizer, base_date)
                
                # Generate outputs
                excel_data = save_to_excel(re_exam_df)
                pdf_data = generate_pdf_timetable(re_exam_df)
                verification_data = save_verification_excel(re_exam_df)
                
                # Statistics
                total_exams = len(re_exam_df)
                unique_branches = re_exam_df['Branch'].nunique()
                all_dates = pd.to_datetime(re_exam_df['Exam Date'], format="%d-%m-%Y").dropna()
                overall_date_range = (all_dates.max() - all_dates.min()).days + 1 if not all_dates.empty else 0
                unique_exam_days = len(all_dates.dt.date.unique())
                
                st.markdown("### 📈 Statistics")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f'<div class="metric-card"><h3>📝 {total_exams}</h3><p>Total Exams</p></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="metric-card"><h3>🏫 {unique_branches}</h3><p>Unique Branches</p></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="metric-card"><h3>📅 {overall_date_range}</h3><p>Days Span</p></div>', unsafe_allow_html=True)
                with col4:
                    st.markdown(f'<div class="metric-card"><h3>🗓️ {unique_exam_days}</h3><p>Exam Days</p></div>', unsafe_allow_html=True)
                
                # Display timetable
                st.markdown('<div class="results-section"><h2>📊 Re-exam Timetable</h2></div>', unsafe_allow_html=True)
                for (program, semester), group in re_exam_df.groupby(['Program', 'SemesterInt']):
                    st.markdown(f"#### {program} - Semester {int_to_roman(semester)}")
                    st.dataframe(group[['Subject', 'Exam Date', 'Exam Time']])
                
                # Download options
                st.markdown("### 📥 Download Outputs")
                col1, col2, col3 = st.columns(3)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                with col1:
                    st.download_button("Download Excel", excel_data.getvalue(), f"re_exam_timetable_{timestamp}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col2:
                    st.download_button("Download PDF", pdf_data, f"re_exam_timetable_{timestamp}.pdf", "application/pdf")
                with col3:
                    st.download_button("Download Verification", verification_data.getvalue(), f"re_exam_verification_{timestamp}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                # Debugging log
                with st.expander("Scheduling Log"):
                    for log in optimizer.log:
                        st.write(log)

if __name__ == "__main__":
    main()
