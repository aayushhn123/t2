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

# Custom CSS (omitted for brevity, assumed unchanged from provided code)
st.markdown("""
<style>
    /* CSS styles as provided earlier */
</style>
""", unsafe_allow_html=True)

# Define branch mappings and logo path (assumed unchanged)
BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS"
}
LOGO_PATH = "logo.png"

# Cache and class definitions (assumed unchanged)
wrap_text_cache = {}

class RealTimeOptimizer:
    # (Class definition remains unchanged from provided code)
    def __init__(self, branches, holidays, time_slots=None):
        self.branches = branches
        self.holidays = holidays
        self.time_slots = time_slots or ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        self.schedule_grid = {}
        self.optimization_log = []
        self.moves_made = 0
        
    def add_exam_to_grid(self, date_str, time_slot, branch, subject):
        if date_str not in self.schedule_grid:
            self.schedule_grid[date_str] = {}
        if time_slot not in self.schedule_grid[date_str]:
            self.schedule_grid[date_str][time_slot] = {}
        self.schedule_grid[date_str][time_slot][branch] = subject
    
    def find_earliest_empty_slot(self, branch, start_date, preferred_time_slot=None):
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

# PDF generation functions (assumed unchanged)
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
    # (Function remains unchanged)
    pass

def print_table_custom(pdf, df, columns, col_widths, line_height=5, header_content=None, branches=None, time_slot=None):
    # (Function remains unchanged)
    pass

def add_footer_with_page_number(pdf, footer_height):
    # (Function remains unchanged)
    pass

def add_header_to_page(pdf, current_date, logo_x, logo_width, header_content, branches, time_slot=None):
    # (Function remains unchanged)
    pass

def calculate_end_time(start_time, duration_hours):
    start = datetime.strptime(start_time, "%I:%M %p")
    duration = timedelta(hours=duration_hours)
    end = start + duration
    return end.strftime("%I:%M %p").replace("AM", "am").replace("PM", "pm")

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    # (Function remains unchanged)
    pass

def generate_pdf_timetable(semester_wise_timetable, output_pdf):
    # (Function remains unchanged)
    pass

# Data processing functions (assumed unchanged)
def read_timetable(uploaded_file):
    # (Function remains unchanged)
    pass

def schedule_semester_non_electives_with_optimization(df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty=False):
    # (Function remains unchanged)
    pass

def process_constraints_with_real_time_optimization(df, holidays, base_date, schedule_by_difficulty=False):
    # (Function remains unchanged)
    pass

def find_next_valid_day_for_electives(start_day, holidays):
    # (Function remains unchanged)
    pass

def optimize_oe_subjects_after_scheduling(sem_dict, holidays, optimizer=None):
    # (Function remains unchanged)
    pass

def reduce_exam_gaps(sem_dict, holidays, max_gap_days=1, time_slots=None):
    """
    Aggressively reduces gaps between exams by moving exams to fill empty slots.
    """
    if not sem_dict:
        return sem_dict, {}
    
    if time_slots is None:
        time_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
    
    # Statistics tracking
    stats = {
        'gaps_found': 0,
        'gaps_reduced': 0,
        'exams_moved': 0,
        'days_saved': 0,
        'gap_details': []
    }
    
    # Combine all data for analysis
    all_data = pd.concat(sem_dict.values(), ignore_index=True)
    if all_data.empty:
        return sem_dict, stats
    
    # Convert exam dates to datetime
    all_data['Exam Date'] = pd.to_datetime(all_data['Exam Date'], format="%d-%m-%Y", errors='coerce')
    all_data = all_data.dropna(subset=['Exam Date'])
    
    # Build schedule grid
    schedule_grid = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: None)))
    branches = all_data['Branch'].unique()
    
    for _, row in all_data.iterrows():
        date_str = row['Exam Date'].strftime("%d-%m-%Y")
        schedule_grid[date_str][row['Time Slot']][row['Branch']] = row['Subject']
    
    # Get all exam dates
    all_exam_dates = sorted(all_data['Exam Date'].dt.date.unique())
    
    if len(all_exam_dates) < 2:
        return sem_dict, stats
    
    # Find all gaps
    all_gaps = []
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
            if gap_days > max_gap_days:
                stats['gaps_found'] += 1
    
    all_gaps.sort(key=lambda x: x['priority'], reverse=True)
    print(f"🔍 Found {len(all_gaps)} total gaps, {stats['gaps_found']} exceed {max_gap_days} days")
    
    # Phase 1: Fill gaps
    for gap_info in all_gaps:
        start_date = gap_info['start_date']
        end_date = gap_infoHEAD['end_date']
        gap_days = gap_info['gap_days']
        
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
                            related_exams = all_data[all_data['Subject'].str.contains(f'({module_code})', regex=False, na=False)]
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
                        print(f"✅ Moved {exam_row['Subject']} from {original_date} to {target_date_str} (saved {days_saved} days)")
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
    
    # Phase 2: Compact schedule
    print("🎯 Compacting schedule further...")
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
                                print(f"✅ Moved {exam['Subject']} from {late_date_str} to {early_date_str} (saved {days_saved} days)")
                                break
    
    # Calculate average gap
    branch_gaps = []
    for branch in branches:
        branch_exams = updated_all_data[updated_all_data['Branch'] == branch]
        if len(branch_exams) > 1:
            branch_dates = sorted(branch_exams['Exam Date'].dt.date.unique())
            gaps = [(branch_dates[i+1] - branch_dates[i]).days - 1 for i in range(len(branch_dates)-1)]
            if gaps:
                branch_gaps.append(sum(gaps) / len(gaps))
    
    stats['average_gap'] = sum(branch_gaps) / len(branch_gaps) if branch_gaps else 0
    
    return sem_dict, stats

def save_to_excel(semester_wise_timetable):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sem, df_sem in semester_wise_timetable.items():
            df_mb = df_sem.copy()
            df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")]
            df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")]
            
            if not df_non_elec.empty:
                df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                df_non_elec = df_non_elec.drop_duplicates(subset=["Exam Date", "Time Slot", "SubBranch", "Subject"])
                pivot_df = df_non_elec.pivot_table(
                    index=["Exam Date", "Time Slot"],
                    columns="SubBranch",
                    values="Subject",
                    aggfunc=lambda x: ", ".join(x)
                ).fillna("---")
                if not pivot_df.empty:
                    sheet_name = f"{df_non_elec['MainBranch'].iloc[0]}_Sem_{sem}"
                    pivot_df.to_excel(writer, sheet_name=sheet_name)
            
            if not df_elec.empty:
                df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                elec_pivot = df_elec.groupby(['OE', 'Exam Date', 'Time Slot'])['Subject'].apply(
                    lambda x: ", ".join(x)
                ).reset_index()
                if not elec_pivot.empty:
                    sheet_name = f"{df_elec['MainBranch'].iloc[0]}_Sem_{sem}_Electives"
                    elec_pivot.to_excel(writer, sheet_name=sheet_name)
    output.seek(0)
    return output

def save_verification_excel(original_df, semester_wise_timetable):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        all_scheduled = pd.concat(semester_wise_timetable.values(), ignore_index=True)
        all_scheduled['Subject'] = all_scheduled['Subject'].str.strip()
        original_df['SubjectName'] = original_df['SubjectName'].str.strip()
        merged_df = original_df.merge(
            all_scheduled[['Subject', 'Exam Date', 'Time Slot']],
            left_on='SubjectName',
            right_on='Subject',
            how='left'
        )
        merged_df['Scheduled'] = merged_df['Exam Date'].notna()
        merged_df.to_excel(writer, sheet_name='Verification', index=False)
    output.seek(0)
    return output

def main():
    st.markdown("""
    <div class="main-header">
        <h1>📅 Exam Timetable Generator</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    # Initialize session state
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
    if 'gap_reduction_stats' not in st.session_state:
        st.session_state.gap_reduction_stats = {}

    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        st.markdown("#### 📅 Base Date for Scheduling")
        base_date = st.date_input("Start date for exams", value=datetime(2025, 4, 1))
        base_date = datetime.combine(base_date, datetime.min.time())

        st.markdown("#### 🛠️ Scheduling Mode")
        schedule_by_difficulty = st.checkbox("Schedule by Difficulty", value=False)
        if schedule_by_difficulty:
            st.markdown('<div class="status-info">ℹ️ Exams will alternate between Easy and Difficult subjects.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-info">ℹ️ Normal scheduling without considering difficulty.</div>',
                        unsafe_allow_html=True)

        st.markdown("#### 📅 Maximum Allowed Gap")
        max_gap_days = st.slider("Maximum days between consecutive exams", 0, 3, 1)
        st.markdown('<div class="status-info">ℹ️ Set to 0 for no gaps (exams on consecutive working days).</div>',
                    unsafe_allow_html=True)

        with st.expander("Holiday Configuration", expanded=True):
            st.markdown("#### 📅 Select Predefined Holidays")
            holiday_dates = []
            col1, col2 = st.columns(2)
            with col1:
                if st.checkbox("April 14, 2025", value=True):
                    holiday_dates.append(date(2025, 4, 14))
            with col2:
                if st.checkbox("May 1, 2025", value=True):
                    holiday_dates.append(date(2025, 5, 1))
            if st.checkbox("August 15, 2025", value=True):
                holiday_dates.append(date(2025, 8, 15))

            st.markdown("#### 📅 Add Custom Holidays")
            for i in range(st.session_state.num_custom_holidays):
                custom_holiday = st.date_input(f"Custom Holiday {i + 1}", value=st.session_state.custom_holidays[i], key=f"custom_holiday_{i}")
                if custom_holiday:
                    holiday_dates.append(custom_holiday)
            if st.button("➕ Add Another Holiday"):
                st.session_state.num_custom_holidays += 1
                st.session_state.custom_holidays.append(None)
                st.rerun()

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
        uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

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
                    df_non_elec, df_ele, original_df = read_timetable(uploaded_file)
                    if df_non_elec is not None and df_ele is not None:
                        non_elec_sched = process_constraints_with_real_time_optimization(df_non_elec, holidays_set, base_date, schedule_by_difficulty)
                        non_elec_df = pd.concat(non_elec_sched.values(), ignore_index=True) if non_elec_sched else pd.DataFrame()
                        non_elec_dates = pd.to_datetime(non_elec_df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        max_non_elec_date = max(non_elec_dates).date() if not non_elec_dates.empty else base_date.date()

                        if df_ele is not None and not df_ele.empty:
                            elective_day1 = find_next_valid_day_for_electives(datetime.combine(max_non_elec_date, datetime.min.time()) + timedelta(days=1), holidays_set)
                            elective_day2 = find_next_valid_day_for_electives(elective_day1 + timedelta(days=1), holidays_set)
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Exam Date'] = elective_day1.strftime("%d-%m-%Y")
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Time Slot'] = "10:00 AM - 1:00 PM"
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Exam Date'] = elective_day2.strftime("%d-%m-%Y")
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Time Slot'] = "2:00 PM - 5:00 PM"
                            final_df = pd.concat([non_elec_df, df_ele], ignore_index=True)
                        else:
                            final_df = non_elec_df

                        final_df["Exam Date"] = pd.to_datetime(final_df["Exam Date"], format="%d-%m-%Y", errors='coerce')
                        final_df = final_df.sort_values(["Exam Date", "Semester", "MainBranch"], ascending=True, na_position='last')
                        sem_dict = {s: final_df[final_df["Semester"] == s].copy() for s in sorted(final_df["Semester"].unique())}

                        sem_dict = optimize_oe_subjects_after_scheduling(sem_dict, holidays_set)
                        sem_dict, gap_stats = reduce_exam_gaps(sem_dict, holidays_set, max_gap_days=max_gap_days)

                        st.session_state.timetable_data = sem_dict
                        st.session_state.original_df = original_df
                        st.session_state.processing_complete = True
                        st.session_state.gap_reduction_stats = gap_stats

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

                        st.session_state.total_exams = total_exams
                        st.session_state.total_semesters = total_semesters
                        st.session_state.total_branches = total_branches
                        st.session_state.overall_date_range = overall_date_range
                        st.session_state.unique_exam_days = unique_exam_days
                        st.session_state.non_elective_range = non_elective_range
                        st.session_state.elective_dates_str = elective_dates_str
                        st.session_state.stream_counts = stream_counts

                        excel_data = save_to_excel(sem_dict)
                        if excel_data:
                            st.session_state.excel_data = excel_data.getvalue()
                        pdf_output = io.BytesIO()
                        temp_pdf_path = "temp_timetable.pdf"
                        generate_pdf_timetable(sem_dict, temp_pdf_path)
                        with open(temp_pdf_path, "rb") as f:
                            pdf_output.write(f.read())
                        pdf_output.seek(0)
                        if os.path.exists(temp_pdf_path):
                            os.remove(temp_pdf_path)
                        st.session_state.pdf_data = pdf_output.getvalue()
                        verification_data = save_verification_excel(original_df, sem_dict)
                        if verification_data:
                            st.session_state.verification_data = verification_data.getvalue()

                        st.markdown('<div class="status-success">🎉 Timetable generated successfully!</div>',
                                    unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="status-error">❌ Failed to read the Excel file.</div>',
                                    unsafe_allow_html=True)
                except Exception as e:
                    st.markdown(f'<div class="status-error">❌ An error occurred: {str(e)}</div>',
                                unsafe_allow_html=True)

    if st.session_state.processing_complete:
        st.markdown("---")
        if st.session_state.unique_exam_days > 20:
            st.warning(f"⚠️ The timetable spans {st.session_state.unique_exam_days} exam days, exceeding the limit of 20 days.")
        
        if 'gap_reduction_stats' in st.session_state:
            stats = st.session_state.gap_reduction_stats
            st.markdown("### 📊 Gap Reduction Results")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Gaps Reduced", stats.get('gaps_reduced', 0))
            with col2:
                st.metric("Exams Moved", stats.get('exams_moved', 0))
            with col3:
                st.metric("Average Gap (days)", f"{stats.get('average_gap', 0):.1f}")

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
                    use_container_width=True
                )
        with col2:
            if st.session_state.pdf_data:
                st.download_button(
                    label="📄 Download PDF File",
                    data=st.session_state.pdf_data,
                    file_name=f"timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        with col3:
            if st.session_state.verification_data:
                st.download_button(
                    label="📋 Download Verification File",
                    data=st.session_state.verification_data,
                    file_name=f"verification_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        with col4:
            if st.button("🔄 Generate New Timetable", use_container_width=True):
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
                st.session_state.gap_reduction_stats = {}
                st.rerun()

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
                df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")]
                df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")]
                
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
                            formatted_dates = [d.strftime("%d-%m-%Y") if pd.notna(d) else "" for d in formatted_pivot.index.levels[0]]
                            formatted_pivot.index = formatted_pivot.index.set_levels(formatted_dates, level=0)
                        st.dataframe(formatted_pivot, use_container_width=True)

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
