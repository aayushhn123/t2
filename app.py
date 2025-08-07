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

class StreamwiseScheduler:
    """New scheduler that handles each stream independently with 15-day limit and no gaps"""
    
    def __init__(self, holidays, base_date):
        self.holidays = holidays
        self.base_date = base_date
        self.time_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        self.scheduling_log = []
        
    def get_valid_exam_days(self, start_date, max_days=50):
        """Get list of valid exam days (excluding Sundays and holidays)"""
        valid_days = []
        current_date = start_date
        days_checked = 0
        
        while len(valid_days) < max_days and days_checked < 100:
            # Skip Sundays (weekday 6) and holidays
            if current_date.weekday() != 6 and current_date.date() not in self.holidays:
                valid_days.append(current_date)
            current_date += timedelta(days=1)
            days_checked += 1
            
        return valid_days
    
    def get_semester_time_slot(self, semester):
        """Get preferred time slot based on semester"""
        if semester % 2 != 0:  # Odd semester
            odd_sem_position = (semester + 1) // 2
            return "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
        else:  # Even semester
            even_sem_position = semester // 2
            return "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    
    def schedule_stream_independently(self, stream_data, stream_name):
        """Schedule all subjects for a specific stream independently"""
        st.write(f"🔧 Scheduling stream: {stream_name}")
        
        # Group by semester
        semesters = stream_data['Semester'].unique()
        stream_schedule = {}
        overall_start_date = self.base_date
        
        for semester in sorted(semesters):
            if semester == 0:
                continue
                
            sem_data = stream_data[stream_data['Semester'] == semester].copy()
            if sem_data.empty:
                continue
            
            st.write(f"  📚 Processing Semester {semester}")
            
            # Get all subjects for this semester and stream
            # Separate common and individual subjects
            common_subjects = sem_data[sem_data['IsCommon'] == 'YES'].copy()
            individual_subjects = sem_data[sem_data['IsCommon'] == 'NO'].copy()
            
            # Count subjects to schedule
            total_subjects = len(common_subjects.drop_duplicates(['ModuleCode'])) + len(individual_subjects)
            st.write(f"    📝 Total subjects to schedule: {total_subjects}")
            
            if total_subjects == 0:
                continue
            
            # Check if we can fit within 15 days
            if total_subjects > 15:
                st.warning(f"⚠️ Stream {stream_name}, Semester {semester} has {total_subjects} subjects, exceeding 15-day limit!")
            
            # Get valid exam days starting from overall_start_date
            valid_days = self.get_valid_exam_days(overall_start_date, max_days=30)
            
            if len(valid_days) < total_subjects:
                st.error(f"❌ Not enough valid days to schedule {total_subjects} subjects for {stream_name} Semester {semester}")
                continue
            
            # Get preferred time slot for this semester
            preferred_slot = self.get_semester_time_slot(semester)
            
            # Schedule subjects one per day with no gaps
            day_index = 0
            
            # First schedule common subjects (avoid duplicates)
            scheduled_modules = set()
            
            for _, row in common_subjects.iterrows():
                module_code = row['ModuleCode']
                if module_code in scheduled_modules:
                    # Skip duplicate common subjects
                    continue
                
                if day_index >= len(valid_days):
                    st.error(f"❌ Ran out of valid days while scheduling common subjects for {stream_name} Semester {semester}")
                    break
                
                exam_date = valid_days[day_index]
                date_str = exam_date.strftime("%d-%m-%Y")
                
                # Update all rows with this module code
                mask = (sem_data['ModuleCode'] == module_code)
                sem_data.loc[mask, 'Exam Date'] = date_str
                sem_data.loc[mask, 'Time Slot'] = preferred_slot
                
                scheduled_modules.add(module_code)
                day_index += 1
                
                st.write(f"    ✅ Scheduled common subject {row['Subject']} on {date_str}")
                self.scheduling_log.append(f"Scheduled common {row['Subject']} for {stream_name} Sem {semester} on {date_str}")
            
            # Then schedule individual subjects
            for _, row in individual_subjects.iterrows():
                if day_index >= len(valid_days):
                    st.error(f"❌ Ran out of valid days while scheduling individual subjects for {stream_name} Semester {semester}")
                    break
                
                exam_date = valid_days[day_index]
                date_str = exam_date.strftime("%d-%m-%Y")
                
                sem_data.loc[sem_data.index == row.name, 'Exam Date'] = date_str
                sem_data.loc[sem_data.index == row.name, 'Time Slot'] = preferred_slot
                
                day_index += 1
                
                st.write(f"    ✅ Scheduled individual subject {row['Subject']} on {date_str}")
                self.scheduling_log.append(f"Scheduled individual {row['Subject']} for {stream_name} Sem {semester} on {date_str}")
            
            # Calculate actual days used
            actual_days_used = day_index
            
            if actual_days_used <= 15:
                st.success(f"✅ Stream {stream_name} Semester {semester}: {actual_days_used} days used (within 15-day limit)")
            else:
                st.warning(f"⚠️ Stream {stream_name} Semester {semester}: {actual_days_used} days used (exceeds 15-day limit)")
            
            # Update overall start date for next semester (start after current semester ends + 1 day gap)
            if day_index > 0:
                last_exam_date = valid_days[day_index - 1]
                overall_start_date = last_exam_date + timedelta(days=2)  # 1 day gap between semesters
            
            stream_schedule[semester] = sem_data
        
        return stream_schedule

def parse_date_safely(date_input, input_format="%d-%m-%Y"):
    """Safely parse date input ensuring DD-MM-YYYY format interpretation"""
    if pd.isna(date_input):
        return None
        
    if isinstance(date_input, pd.Timestamp):
        return date_input
    
    if isinstance(date_input, str):
        try:
            return pd.to_datetime(date_input, format=input_format, errors='raise')
        except:
            try:
                return pd.to_datetime(date_input, dayfirst=True, errors='raise')
            except:
                return None
    
    return pd.to_datetime(date_input, errors='coerce')

def process_new_scheduling_logic(df, holidays, base_date):
    """New main scheduling function that processes each stream independently"""
    st.info("🚀 Starting new streamwise independent scheduling...")
    
    # Initialize scheduler
    scheduler = StreamwiseScheduler(holidays, base_date)
    
    # Get all unique streams (MainBranch-SubBranch combinations)
    df['Stream'] = df['MainBranch'] + '-' + df['SubBranch']
    streams = df['Stream'].unique()
    
    st.write(f"📊 Found {len(streams)} streams to schedule: {', '.join(streams)}")
    
    all_stream_schedules = {}
    
    # Process each stream independently
    for stream in sorted(streams):
        st.write(f"\n🔄 Processing stream: {stream}")
        
        # Get all data for this stream
        stream_data = df[df['Stream'] == stream].copy()
        
        # Schedule this stream independently
        stream_schedule = scheduler.schedule_stream_independently(stream_data, stream)
        
        # Merge into overall schedule
        for semester, sem_data in stream_schedule.items():
            if semester not in all_stream_schedules:
                all_stream_schedules[semester] = []
            all_stream_schedules[semester].append(sem_data)
    
    # Combine all streams for each semester
    final_schedule = {}
    for semester, sem_data_list in all_stream_schedules.items():
        if sem_data_list:
            final_schedule[semester] = pd.concat(sem_data_list, ignore_index=True)
    
    # Validation: Check that each stream has max 15 exam days per semester
    st.write("\n📋 Validation Summary:")
    
    validation_results = []
    
    for semester, sem_data in final_schedule.items():
        for stream in sem_data['Stream'].unique():
            stream_sem_data = sem_data[sem_data['Stream'] == stream]
            unique_dates = stream_sem_data['Exam Date'].dropna().unique()
            days_used = len(unique_dates)
            
            validation_results.append({
                'Stream': stream,
                'Semester': semester,
                'Days Used': days_used,
                'Within Limit': days_used <= 15
            })
            
            if days_used <= 15:
                st.write(f"✅ {stream} Semester {semester}: {days_used} days (✓)")
            else:
                st.write(f"❌ {stream} Semester {semester}: {days_used} days (exceeds limit)")
    
    # Show validation summary
    validation_df = pd.DataFrame(validation_results)
    if not validation_df.empty:
        with st.expander("📊 Detailed Validation Results"):
            st.dataframe(validation_df, use_container_width=True)
    
    # Show scheduling log
    if scheduler.scheduling_log:
        with st.expander("📝 Detailed Scheduling Log"):
            for log_entry in scheduler.scheduling_log[-50:]:  # Show last 50 entries
                st.write(f"• {log_entry}")
    
    return final_schedule

def schedule_electives_after_main(sem_dict, holidays):
    """Schedule electives after main scheduling is complete"""
    st.info("🎯 Scheduling Open Electives (OE) after main subjects...")
    
    # Find the latest date from all scheduled subjects
    all_dates = []
    for sem_data in sem_dict.values():
        dates = pd.to_datetime(sem_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        all_dates.extend(dates.tolist())
    
    if not all_dates:
        start_date = datetime(2025, 4, 1)
    else:
        latest_date = max(all_dates)
        start_date = latest_date + timedelta(days=1)
    
    # Find next valid days for electives
    def find_next_valid_day(start_day):
        day = start_day
        while True:
            if day.weekday() != 6 and day.date() not in holidays:
                return day
            day += timedelta(days=1)
    
    # Schedule OE1 and OE5 together, then OE2 the next day
    oe1_oe5_date = find_next_valid_day(start_date)
    oe2_date = find_next_valid_day(oe1_oe5_date + timedelta(days=1))
    
    oe1_oe5_date_str = oe1_oe5_date.strftime("%d-%m-%Y")
    oe2_date_str = oe2_date.strftime("%d-%m-%Y")
    
    # Update electives in all semesters
    electives_scheduled = 0
    for semester, sem_data in sem_dict.items():
        # Schedule OE1 and OE5
        oe1_oe5_mask = sem_data['OE'].isin(['OE1', 'OE5'])
        if oe1_oe5_mask.any():
            sem_dict[semester].loc[oe1_oe5_mask, 'Exam Date'] = oe1_oe5_date_str
            sem_dict[semester].loc[oe1_oe5_mask, 'Time Slot'] = "10:00 AM - 1:00 PM"
            electives_scheduled += oe1_oe5_mask.sum()
        
        # Schedule OE2
        oe2_mask = sem_data['OE'] == 'OE2'
        if oe2_mask.any():
            sem_dict[semester].loc[oe2_mask, 'Exam Date'] = oe2_date_str
            sem_dict[semester].loc[oe2_mask, 'Time Slot'] = "2:00 PM - 5:00 PM"
            electives_scheduled += oe2_mask.sum()
    
    if electives_scheduled > 0:
        st.success(f"✅ Scheduled {electives_scheduled} elective subjects")
        st.write(f"📅 OE1/OE5: {oe1_oe5_date_str}")
        st.write(f"📅 OE2: {oe2_date_str}")
    
    return sem_dict

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

        # Handle normal subjects
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
                            
                            if duration != 3 and subject_time_slot and subject_time_slot.strip():
                                start_time = subject_time_slot.split(" - ")[0]
                                end_time = calculate_end_time(start_time, duration)
                                modified_subjects.append(f"{base_subject} ({start_time} to {end_time})")
                            else:
                                if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                                    if subject_time_slot != default_time_slot:
                                        modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                                    else:
                                        modified_subjects.append(base_subject)
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
                
                pdf.add_page()
                footer_height = 25
                add_footer_with_page_number(pdf, footer_height)
                
                print_table_custom(pdf, chunk_df[cols_to_print], cols_to_print, col_widths, line_height=line_height, 
                                 header_content=header_content, branches=chunk, time_slot=time_slot)

        # Handle electives with updated table structure
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
                    
                    if duration != 3 and subject_time_slot and subject_time_slot.strip():
                        start_time = subject_time_slot.split(" - ")[0]
                        end_time = calculate_end_time(start_time, duration)
                        modified_subjects.append(f"{base_subject} ({start_time} to {end_time})")
                    else:
                        if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                            if subject_time_slot != default_time_slot:
                                modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                            else:
                                modified_subjects.append(base_subject)
                        else:
                            modified_subjects.append(base_subject)
                
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
                df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()

                if not df_non_elec.empty:
                    difficulty_str = df_non_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                    difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                    
                    df_non_elec["SubjectDisplay"] = df_non_elec["Subject"]
                    duration_suffix = df_non_elec.apply(
                        lambda row: f" [Duration: {row['Exam Duration']} hrs]" if row['Exam Duration'] != 3 else '', axis=1)
                    df_non_elec["SubjectDisplay"] = df_non_elec["SubjectDisplay"] + difficulty_suffix + duration_suffix
                    df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
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
        
        # File upload section
        st.markdown("#### 📂 Upload Timetable")
        uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'], key="timetable_uploader")
        
        # Starting date selection
        st.markdown("#### 📅 Starting Date")
        start_date = st.date_input("Select starting date for scheduling", 
                                 value=date(2025, 4, 1),
                                 min_value=date(2025, 1, 1),
                                 max_value=date(2025, 12, 31),
                                 key="start_date")
        
        # Holiday configuration
        st.markdown("#### 🎉 Holidays")
        with st.expander("Configure Holidays"):
            st.session_state.num_custom_holidays = st.number_input(
                "Number of custom holidays",
                min_value=0,
                max_value=20,
                value=st.session_state.num_custom_holidays,
                step=1,
                key="num_holidays"
            )
            
            # Update holiday inputs based on number selected
            if len(st.session_state.custom_holidays) < st.session_state.num_custom_holidays:
                st.session_state.custom_holidays.extend(
                    [None] * (st.session_state.num_custom_holidays - len(st.session_state.custom_holidays))
                )
            elif len(st.session_state.custom_holidays) > st.session_state.num_custom_holidays:
                st.session_state.custom_holidays = st.session_state.custom_holidays[:st.session_state.num_custom_holidays]
            
            for i in range(st.session_state.num_custom_holidays):
                st.session_state.custom_holidays[i] = st.date_input(
                    f"Holiday {i + 1}",
                    value=st.session_state.custom_holidays[i],
                    min_value=date(2025, 1, 1),
                    max_value=date(2025, 12, 31),
                    key=f"holiday_{i}"
                )
        
        # Generate timetable button
        generate_button = st.button("Generate Timetable", key="generate_button", use_container_width=True)

    # Main content area
    if uploaded_file is not None and generate_button:
        try:
            st.markdown('<div class="status-info">⏳ Processing timetable...</div>', unsafe_allow_html=True)
            
            # Read and process input file
            df_non, df_ele, original_df = read_timetable(uploaded_file)
            if df_non is None or original_df is None:
                st.markdown('<div class="status-error">❌ Failed to read timetable data.</div>', 
                           unsafe_allow_html=True)
                st.session_state.processing_complete = False
                return

            # Process holidays
            holidays_set = set(h.date() for h in st.session_state.custom_holidays if h is not None)
            
            # Schedule non-elective subjects using the new scheduling logic
            semester_wise_timetable = process_new_scheduling_logic(
                df_non,
                holidays_set,
                datetime.combine(start_date, datetime.min.time())
            )
            
            # Handle electives
            if df_ele is not None and not df_ele.empty:
                semester_wise_timetable = schedule_electives_after_main(semester_wise_timetable, holidays_set)
                final_df = pd.concat([df_non, df_ele], ignore_index=True)
            else:
                final_df = df_non
                st.write("No electives to schedule.")

            # Save to session state
            st.session_state.timetable_data = semester_wise_timetable

            # Generate Excel and PDF
            st.write("Generating Excel file...")
            st.session_state.excel_data = save_to_excel(semester_wise_timetable)
            st.session_state.verification_data = save_verification_excel(original_df, semester_wise_timetable)

            output_pdf = "timetable_output.pdf"
            st.write("Generating PDF...")
            generate_pdf_timetable(semester_wise_timetable, output_pdf)
            with open(output_pdf, "rb") as f:
                st.session_state.pdf_data = f.read()
            if os.path.exists(output_pdf):
                os.remove(output_pdf)

            # Calculate statistics
            all_data = pd.concat(semester_wise_timetable.values(), ignore_index=True)
            st.session_state.total_exams = len(all_data)
            st.session_state.total_semesters = len(semester_wise_timetable)
            st.session_state.total_branches = len(all_data['Branch'].unique())
            
            # Calculate date range
            valid_dates = pd.to_datetime(
                all_data['Exam Date'], format="%d-%m-%Y", errors='coerce'
            ).dropna()
            if not valid_dates.empty:
                min_date = valid_dates.min()
                max_date = valid_dates.max()
                st.session_state.overall_date_range = (max_date - min_date).days + 1
                st.session_state.unique_exam_days = len(valid_dates.dt.date.unique())
                st.session_state.non_elective_range = (
                    f"{min_date.strftime('%d-%m-%Y')} to {max_date.strftime('%d-%m-%Y')}"
                    if not valid_dates.empty else "N/A"
                )
                elective_dates = pd.to_datetime(
                    df_ele['Exam Date'], format="%d-%m-%Y", errors='coerce'
                ).dropna()
                st.session_state.elective_dates_str = (
                    ", ".join(sorted(set(d.strftime("%d-%m-%Y") for d in elective_dates)))
                    if not elective_dates.empty else "N/A"
                )
            else:
                st.session_state.overall_date_range = 0
                st.session_state.unique_exam_days = 0
                st.session_state.non_elective_range = "N/A"
                st.session_state.elective_dates_str = "N/A"

            # Stream-wise exam counts
            stream_counts = all_data.groupby('Branch').size().reset_index(name='Exam Count')
            st.session_state.stream_counts = stream_counts

            st.session_state.processing_complete = True
            st.markdown('<div class="status-success">✅ Timetable generated successfully!</div>',
                        unsafe_allow_html=True)

        except Exception as e:
            st.markdown(f'<div class="status-error">❌ Error: {str(e)}</div>', unsafe_allow_html=True)
            st.session_state.processing_complete = False

    if st.session_state.processing_complete:
        st.markdown("""
        <div class="results-section">
            <h2>📊 Results</h2>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("""
            <div class="metric-card">
                <h3>📚 <span>{}</span></h3>
                <p>Total Exams</p>
            </div>
            """.format(st.session_state.total_exams), unsafe_allow_html=True)
        with col2:
            st.markdown("""
            <div class="metric-card">
                <h3>🏫 <span>{}</span></h3>
                <p>Total Semesters</p>
            </div>
            """.format(st.session_state.total_semesters), unsafe_allow_html=True)
        with col3:
            st.markdown("""
            <div class="metric-card">
                <h3>🌿 <span>{}</span></h3>
                <p>Total Streams</p>
            </div>
            """.format(st.session_state.total_branches), unsafe_allow_html=True)
        with col4:
            st.markdown("""
            <div class="metric-card">
                <h3>📅 <span>{}</span></h3>
                <p>Exam Days</p>
            </div>
            """.format(st.session_state.unique_exam_days), unsafe_allow_html=True)

        st.markdown("""
        <div class="stats-section">
            <h3>📈 Schedule Statistics</h3>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"**Overall Date Range:** {st.session_state.overall_date_range} days")
        st.markdown(f"**Non-Elective Exam Range:** {st.session_state.non_elective_range}")
        st.markdown(f"**Elective Exam Dates:** {st.session_state.elective_dates_str}")

        st.markdown("#### Stream-wise Exam Counts")
        st.dataframe(st.session_state.stream_counts, use_container_width=True)

        if st.session_state.excel_data:
            st.download_button(
                label="📥 Download Excel Timetable",
                data=st.session_state.excel_data,
                file_name="timetable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        if st.session_state.pdf_data:
            st.download_button(
                label="📥 Download PDF Timetable",
                data=st.session_state.pdf_data,
                file_name="timetable.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        if st.session_state.verification_data:
            st.download_button(
                label="📥 Download Verification Excel",
                data=st.session_state.verification_data,
                file_name="verification_timetable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown("""
    <div class="footer">
        <p>Developed by MPSTME Exam Scheduling Team | © 2025 All Rights Reserved</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
