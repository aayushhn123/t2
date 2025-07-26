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

class EnhancedRealTimeOptimizer:
    """Enhanced optimizer that specifically targets gap reduction"""
    
    def __init__(self, branches, holidays, time_slots=None):
        self.branches = branches
        self.holidays = holidays
        self.time_slots = time_slots or ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        self.schedule_grid = {}  # date -> time_slot -> branch -> subject/None
        self.optimization_log = []
        self.moves_made = 0
        self.gap_reductions = 0
        
    def add_exam_to_grid(self, date_str, time_slot, branch, subject):
        """Add an exam to the schedule grid"""
        if date_str not in self.schedule_grid:
            self.schedule_grid[date_str] = {}
        if time_slot not in self.schedule_grid[date_str]:
            self.schedule_grid[date_str][time_slot] = {}
        self.schedule_grid[date_str][time_slot][branch] = subject
    
    def remove_exam_from_grid(self, date_str, time_slot, branch):
        """Remove an exam from the schedule grid"""
        if (date_str in self.schedule_grid and 
            time_slot in self.schedule_grid[date_str] and
            branch in self.schedule_grid[date_str][time_slot]):
            self.schedule_grid[date_str][time_slot][branch] = None
    
    def find_gaps_in_schedule(self):
        """Identify gaps of more than 2 days in the exam schedule"""
        # Get all dates with exams
        exam_dates = []
        for date_str in self.schedule_grid:
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
            for time_slot in self.schedule_grid[date_str]:
                for branch in self.schedule_grid[date_str][time_slot]:
                    if self.schedule_grid[date_str][time_slot][branch] is not None:
                        exam_dates.append(date_obj.date())
                        break
        
        # Remove duplicates and sort
        unique_exam_dates = sorted(set(exam_dates))
        
        if len(unique_exam_dates) < 2:
            return []
        
        gaps = []
        for i in range(len(unique_exam_dates) - 1):
            current_date = unique_exam_dates[i]
            next_date = unique_exam_dates[i + 1]
            
            # Calculate business days between dates (excluding weekends and holidays)
            gap_days = 0
            check_date = current_date + timedelta(days=1)
            while check_date < next_date:
                if check_date.weekday() != 6 and check_date not in self.holidays:
                    gap_days += 1
                check_date += timedelta(days=1)
            
            if gap_days > 2:
                gaps.append({
                    'start_date': current_date,
                    'end_date': next_date,
                    'gap_days': gap_days,
                    'start_str': current_date.strftime("%d-%m-%Y"),
                    'end_str': next_date.strftime("%d-%m-%Y")
                })
        
        return gaps
    
    def find_moveable_exams_after_gap(self, gap_start_date):
        """Find exams that can be moved to fill a gap"""
        moveable_exams = []
        
        # Look for exams scheduled after the gap that could be moved earlier
        for date_str in self.schedule_grid:
            date_obj = datetime.strptime(date_str, "%d-%m-%Y").date()
            
            if date_obj > gap_start_date:
                for time_slot in self.schedule_grid[date_str]:
                    for branch in self.schedule_grid[date_str][time_slot]:
                        subject = self.schedule_grid[date_str][time_slot][branch]
                        if subject is not None:
                            moveable_exams.append({
                                'date_str': date_str,
                                'date_obj': date_obj,
                                'time_slot': time_slot,
                                'branch': branch,
                                'subject': subject
                            })
        
        # Sort by date to prioritize moving exams that are furthest from the gap
        return sorted(moveable_exams, key=lambda x: x['date_obj'], reverse=True)
    
    def find_available_slot_in_gap(self, gap_start_date, gap_end_date, branch):
        """Find an available slot within a gap for a specific branch"""
        current_date = gap_start_date + timedelta(days=1)
        
        while current_date < gap_end_date:
            # Skip weekends and holidays
            if current_date.weekday() == 6 or current_date in self.holidays:
                current_date += timedelta(days=1)
                continue
            
            date_str = current_date.strftime("%d-%m-%Y")
            
            # Check if this branch already has an exam on this date
            branch_has_exam = False
            if date_str in self.schedule_grid:
                for time_slot in self.time_slots:
                    if (time_slot in self.schedule_grid[date_str] and
                        branch in self.schedule_grid[date_str][time_slot] and
                        self.schedule_grid[date_str][time_slot][branch] is not None):
                        branch_has_exam = True
                        break
            
            # If branch has no exam on this date, find an available time slot
            if not branch_has_exam:
                for time_slot in self.time_slots:
                    if date_str not in self.schedule_grid:
                        return date_str, time_slot
                    if time_slot not in self.schedule_grid[date_str]:
                        return date_str, time_slot
                    if self.schedule_grid[date_str][time_slot].get(branch) is None:
                        return date_str, time_slot
            
            current_date += timedelta(days=1)
        
        return None, None
    
    def optimize_gaps(self):
        """Main function to optimize gaps in the schedule"""
        st.info("🎯 Analyzing schedule for gaps larger than 2 days...")
        
        gaps = self.find_gaps_in_schedule()
        
        if not gaps:
            st.success("✅ No significant gaps found in the schedule!")
            return
        
        st.info(f"📊 Found {len(gaps)} gaps larger than 2 days")
        
        for gap in gaps:
            st.write(f"🔍 Gap detected: {gap['gap_days']} days between {gap['start_str']} and {gap['end_str']}")
            
            # Find exams that can be moved to fill this gap
            moveable_exams = self.find_moveable_exams_after_gap(gap['start_date'])
            
            moved_count = 0
            for exam in moveable_exams:
                # Try to find a slot in the gap for this exam
                new_date_str, new_time_slot = self.find_available_slot_in_gap(
                    gap['start_date'], gap['end_date'], exam['branch']
                )
                
                if new_date_str and new_time_slot:
                    # Move the exam
                    self.remove_exam_from_grid(exam['date_str'], exam['time_slot'], exam['branch'])
                    self.add_exam_to_grid(new_date_str, new_time_slot, exam['branch'], exam['subject'])
                    
                    moved_count += 1
                    self.moves_made += 1
                    
                    days_moved = (datetime.strptime(exam['date_str'], "%d-%m-%Y") - 
                                datetime.strptime(new_date_str, "%d-%m-%Y")).days
                    
                    self.optimization_log.append(
                        f"📅 Moved {exam['subject']} ({exam['branch']}) from {exam['date_str']} to {new_date_str} (saved {days_moved} days)"
                    )
                    
                    # Stop if we've filled enough slots in this gap
                    if moved_count >= min(gap['gap_days'] - 2, 3):  # Don't overfill
                        break
            
            if moved_count > 0:
                self.gap_reductions += 1
                st.success(f"✅ Reduced gap by moving {moved_count} exams")
    
    def update_dataframe_from_grid(self, df_combined):
        """Update the DataFrame with optimized schedule from grid"""
        for idx, row in df_combined.iterrows():
            branch = row['Branch']
            subject = row['Subject']
            
            # Find this exam in the grid
            for date_str in self.schedule_grid:
                for time_slot in self.schedule_grid[date_str]:
                    if (branch in self.schedule_grid[date_str][time_slot] and
                        self.schedule_grid[date_str][time_slot][branch] == subject):
                        df_combined.at[idx, 'Exam Date'] = date_str
                        df_combined.at[idx, 'Time Slot'] = time_slot
                        break
        
        return df_combined
    
    def find_earliest_empty_slot(self, branch, start_date, preferred_time_slot=None):
        """Find the earliest empty slot for a branch - ensuring only one exam per day per branch"""
        # Sort dates chronologically
        sorted_dates = sorted(self.schedule_grid.keys(), 
                            key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        
        for date_str in sorted_dates:
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
            
            # Skip dates before start_date
            if date_obj < start_date:
                continue
            
            # Skip weekends and holidays
            if date_obj.weekday() == 6 or date_obj.date() in self.holidays:
                continue
            
            # Check if this branch already has ANY exam on this date
            branch_has_exam_today = False
            if date_str in self.schedule_grid:
                for time_slot in self.time_slots:
                    if (time_slot in self.schedule_grid[date_str] and
                        branch in self.schedule_grid[date_str][time_slot] and
                        self.schedule_grid[date_str][time_slot][branch] is not None):
                        branch_has_exam_today = True
                        break
            
            # If branch already has an exam today, skip this date entirely
            if branch_has_exam_today:
                continue
            
            # Check preferred time slot first
            if preferred_time_slot:
                if (date_str in self.schedule_grid and 
                    preferred_time_slot in self.schedule_grid[date_str]):
                    if self.schedule_grid[date_str][preferred_time_slot].get(branch) is None:
                        return date_str, preferred_time_slot
            
            # Check all time slots if preferred slot wasn't available
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
        
        # Calculate new span after optimization
        exam_dates = []
        for date_str in self.schedule_grid:
            for time_slot in self.schedule_grid[date_str]:
                for branch in self.schedule_grid[date_str][time_slot]:
                    if self.schedule_grid[date_str][time_slot][branch] is not None:
                        exam_dates.append(datetime.strptime(date_str, "%d-%m-%Y").date())
                        break
        
        unique_exam_dates = sorted(set(exam_dates))
        optimized_span = 0
        if len(unique_exam_dates) >= 2:
            optimized_span = (max(unique_exam_dates) - min(unique_exam_dates)).days + 1
        
        return {
            'total_slots': total_slots,
            'filled_slots': filled_slots,
            'empty_slots': total_slots - filled_slots,
            'utilization': (filled_slots / total_slots * 100) if total_slots > 0 else 0,
            'optimized_span': optimized_span,
            'gap_reductions': self.gap_reductions
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
        page_number_pattern = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*
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

def schedule_semester_non_electives_with_optimization(df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty=False):
    """Enhanced scheduling that fills empty slots in real-time and ensures all subjects are scheduled"""
    
    # Get semester time slot
    sem = df_sem["Semester"].iloc[0]
    if sem % 2 != 0:
        odd_sem_position = (sem + 1) // 2
        preferred_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:
        even_sem_position = sem // 2
        preferred_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    
    # Schedule COMP subjects
    comp_subjects = df_sem[(df_sem['Category'] == 'COMP') & (df_sem['IsCommon'] == 'NO') & (df_sem['Exam Date'] == "")]
    
    for idx, row in comp_subjects.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        
        # First, try to find an empty slot
        empty_date, empty_slot = optimizer.find_earliest_empty_slot(branch, base_date, preferred_slot)
        
        if empty_date:
            # Found an empty slot - use it
            df_sem.at[idx, 'Exam Date'] = empty_date
            df_sem.at[idx, 'Time Slot'] = empty_slot
            optimizer.add_exam_to_grid(empty_date, empty_slot, branch, subject)
            exam_days[branch].add(datetime.strptime(empty_date, "%d-%m-%Y").date())
            optimizer.optimization_log.append(f"✅ Filled empty slot: {subject} on {empty_date} at {empty_slot}")
            optimizer.moves_made += 1
        else:
            # No empty slot found - schedule normally
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
    
    # Schedule ELEC subjects
    elec_subjects = df_sem[(df_sem['Category'] == 'ELEC') & (df_sem['IsCommon'] == 'NO') & (df_sem['Exam Date'] == "")]
    
    for idx, row in elec_subjects.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        
        # Try to find empty slot first
        empty_date, empty_slot = optimizer.find_earliest_empty_slot(branch, base_date, preferred_slot)
        
        if empty_date:
            # Found an empty slot - use it
            df_sem.at[idx, 'Exam Date'] = empty_date
            df_sem.at[idx, 'Time Slot'] = empty_slot
            optimizer.add_exam_to_grid(empty_date, empty_slot, branch, subject)
            exam_days[branch].add(datetime.strptime(empty_date, "%d-%m-%Y").date())
            optimizer.optimization_log.append(f"✅ Filled empty slot: ELEC {subject} on {empty_date} at {empty_slot}")
            optimizer.moves_made += 1
        else:
            # No empty slot found - find next available day with NO EXAM for this branch
            current_date = base_date
            scheduled = False
            
            while not scheduled:
                date_str = current_date.strftime("%d-%m-%Y")
                
                if current_date.weekday() != 6 and current_date.date() not in holidays:
                    # Only schedule if this branch has NO exam on this date
                    if current_date.date() not in exam_days[branch]:
                        df_sem.at[idx, 'Exam Date'] = date_str
                        df_sem.at[idx, 'Time Slot'] = preferred_slot
                        optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                        exam_days[branch].add(current_date.date())
                        scheduled = True
                        optimizer.optimization_log.append(f"📅 Scheduled ELEC {subject} for {branch} on {date_str}")
                
                current_date += timedelta(days=1)
    
    # Assign time slot to any remaining exams (safety net)
    df_sem.loc[df_sem['Time Slot'] == "", 'Time Slot'] = preferred_slot
    
    # Final check - ensure no exam is left unscheduled
    unscheduled = df_sem[df_sem['Exam Date'] == ""]
    if not unscheduled.empty:
        st.warning(f"⚠️ {len(unscheduled)} subjects remain unscheduled in semester {sem}")
        # Force schedule them on empty days
        for idx, row in unscheduled.iterrows():
            branch = row['Branch']
            subject = row['Subject']
            current_date = base_date
            
            while True:
                date_str = current_date.strftime("%d-%m-%Y")
                if current_date.weekday() != 6 and current_date.date() not in holidays:
                    # Force schedule only on days with no exams for this branch
                    if current_date.date() not in exam_days[branch]:
                        df_sem.at[idx, 'Exam Date'] = date_str
                        df_sem.at[idx, 'Time Slot'] = preferred_slot
                        optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                        exam_days[branch].add(current_date.date())
                        optimizer.optimization_log.append(f"🔧 Force scheduled {subject} on {date_str}")
                        break
                current_date += timedelta(days=1)
    
    return df_sem

def enhanced_process_constraints_with_gap_optimization(df, holidays, base_date, schedule_by_difficulty=False):
    """
    Enhanced process_constraints that reduces gaps between exams
    """
    # Initialize exam_days for all branches
    all_branches = df['Branch'].unique()
    exam_days = {branch: set() for branch in all_branches}
    
    # Initialize the enhanced optimizer
    optimizer = EnhancedRealTimeOptimizer(all_branches, holidays)
    optimizer.initialize_grid_with_empty_days(base_date, num_days=40)
    
    st.info(f"🔧 Scheduling {len(df)} subjects across {len(all_branches)} branches...")
    
    # Helper function for finding earliest slot
    def find_earliest_available_slot_with_empty_check(start_day, for_branches, subject, optimizer):
        """Enhanced slot finding that checks empty slots first"""
        
        # First, check if there's a common empty slot for all branches
        if len(for_branches) == 1:
            # Single branch - use optimizer
            empty_date, empty_slot = optimizer.find_earliest_empty_slot(for_branches[0], start_day)
            if empty_date:
                return datetime.strptime(empty_date, "%d-%m-%Y"), empty_slot
        else:
            # Multiple branches - find common empty slot
            sorted_dates = sorted(optimizer.schedule_grid.keys(), 
                                key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
            
            for date_str in sorted_dates:
                date_obj = datetime.strptime(date_str, "%d-%m-%Y")
                if date_obj < start_day:
                    continue
                    
                for time_slot in optimizer.time_slots:
                    # Check if slot is empty for all branches
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
        
        # No empty slot found - use normal scheduling
        current_date = start_day
        while True:
            current_date_only = current_date.date()
            
            if current_date.weekday() == 6 or current_date_only in holidays:
                current_date += timedelta(days=1)
                continue
            
            if all(current_date_only not in exam_days[branch] for branch in for_branches):
                return current_date, None
            
            current_date += timedelta(days=1)

    # Count subjects by category
    comp_common = len(df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'YES')])
    comp_individual = len(df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'NO')])
    elec_common = len(df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'YES')])
    elec_individual = len(df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'NO')])
    
    st.write(f"📊 Subject distribution: COMP (Common: {comp_common}, Individual: {comp_individual}), ELEC (Common: {elec_common}, Individual: {elec_individual})")

    # Schedule common COMP subjects with optimization
    common_comp = df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'YES')]
    for module_code, group in common_comp.groupby('ModuleCode'):
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

    # Schedule common ELEC subjects with optimization
    common_elec = df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'YES')]
    for module_code, group in common_elec.groupby('ModuleCode'):
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

    # Schedule remaining subjects per semester with optimization
    final_list = []
    for sem in sorted(df["Semester"].unique()):
        if sem == 0:
            continue
        df_sem = df[df["Semester"] == sem].copy()
        if df_sem.empty:
            continue
            
        # Count unscheduled subjects before processing
        unscheduled_before = len(df_sem[df_sem['Exam Date'] == ""])
        
        scheduled_sem = schedule_semester_non_electives_with_optimization(
            df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty
        )
        
        # Count unscheduled subjects after processing
        unscheduled_after = len(scheduled_sem[scheduled_sem['Exam Date'] == ""])
        
        if unscheduled_after > 0:
            st.warning(f"⚠️ Semester {sem}: {unscheduled_after} subjects remain unscheduled out of {unscheduled_before}")
        else:
            st.success(f"✅ Semester {sem}: All {unscheduled_before} subjects scheduled successfully")
        
        final_list.append(scheduled_sem)

    if not final_list:
        return {}

    df_combined = pd.concat(final_list, ignore_index=True)
    
    # NEW: Perform gap optimization after initial scheduling
    st.markdown("---")
    st.info("🎯 Starting gap optimization phase...")
    
    # First, populate the optimizer grid with current schedule
    for idx, row in df_combined.iterrows():
        if pd.notna(row['Exam Date']) and row['Exam Date'] != "":
            date_str = row['Exam Date']
            time_slot = row['Time Slot']
            branch = row['Branch']
            subject = row['Subject']
            optimizer.add_exam_to_grid(date_str, time_slot, branch, subject)
    
    # Calculate initial span
    initial_dates = pd.to_datetime(df_combined['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    initial_span = (max(initial_dates) - min(initial_dates)).days + 1 if not initial_dates.empty else 0
    st.info(f"📊 Initial exam span: {initial_span} days")
    
    # Perform gap optimization
    optimizer.optimize_gaps()
    
    # Update DataFrame with optimized schedule
    df_combined = optimizer.update_dataframe_from_grid(df_combined)
    
    # Check for any unscheduled subjects
    unscheduled_final = df_combined[df_combined['Exam Date'] == ""]
    if not unscheduled_final.empty:
        st.error(f"❌ {len(unscheduled_final)} subjects remain unscheduled!")
        with st.expander("View unscheduled subjects"):
            st.dataframe(unscheduled_final[['Branch', 'Subject', 'Category', 'IsCommon']])
    
    # Display optimization summary
    schedule_summary = optimizer.get_schedule_summary()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Optimization Moves", optimizer.moves_made)
    with col2:
        st.metric("Gaps Reduced", schedule_summary['gap_reductions'])
    with col3:
        st.metric("Final Exam Span", f"{schedule_summary['optimized_span']} days")
    with col4:
        span_reduction = initial_span - schedule_summary['optimized_span']
        st.metric("Days Saved", span_reduction, delta=f"-{span_reduction}")
    
    if optimizer.moves_made > 0:
        with st.expander("📝 Gap Optimization Log", expanded=False):
            for log in optimizer.optimization_log:
                st.write(log)
    
    # Rest of the function remains the same
    sem_dict = {}
    for sem in sorted(df_combined["Semester"].unique()):
        sem_dict[sem] = df_combined[df_combined["Semester"] == sem].copy()

    # Calculate final span and provide feedback
    final_dates = pd.to_datetime(df_combined['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    if not final_dates.empty:
        start_date = min(final_dates)
        end_date = max(final_dates)
        final_span = (end_date - start_date).days + 1
        
        if final_span <= 14:
            st.success(f"🎉 Excellent optimization! Final span: {final_span} days (target achieved)")
        elif final_span <= 16:
            st.success(f"✅ Great optimization! Final span: {final_span} days (within 16-day target)")
        elif final_span <= 20:
            st.info(f"✅ Good optimization! Final span: {final_span} days (within 20-day limit)")
        else:
            st.warning(f"⚠️ The timetable spans {final_span} days, still exceeding the 20-day limit.")

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

def enhanced_optimize_oe_subjects_after_scheduling(sem_dict, holidays, optimizer=None):
    """
    Enhanced OE optimization that considers gap reduction
    """
    if not sem_dict:
        return sem_dict
    
    st.info("🎯 Optimizing Open Elective (OE) placement with gap reduction...")
    
    # Combine all data to analyze the schedule
    all_data = pd.concat(sem_dict.values(), ignore_index=True)
    
    # Separate OE and non-OE data
    oe_data = all_data[all_data['OE'].notna() & (all_data['OE'].str.strip() != "")]
    non_oe_data = all_data[all_data['OE'].isna() | (all_data['OE'].str.strip() == "")]
    
    if oe_data.empty:
        st.info("No OE subjects to optimize")
        return sem_dict
    
    # Build complete schedule grid from current state
    schedule_grid = {}
    branches = all_data['Branch'].unique()
    
    # First, populate with all scheduled exams
    for _, row in all_data.iterrows():
        if pd.notna(row['Exam Date']):
            # Handle both string and Timestamp types for Exam Date
            if isinstance(row['Exam Date'], pd.Timestamp):
                date_str = row['Exam Date'].strftime("%d-%m-%Y")
            else:
                date_str = str(row['Exam Date'])
            
            if date_str not in schedule_grid:
                schedule_grid[date_str] = {}
            if row['Time Slot'] not in schedule_grid[date_str]:
                schedule_grid[date_str][row['Time Slot']] = {}
            schedule_grid[date_str][row['Time Slot']][row['Branch']] = row['Subject']
    
    # Find all dates in the schedule
    all_dates = sorted(schedule_grid.keys(), 
                      key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    if not all_dates:
        return sem_dict
    
    # Get date range
    start_date = datetime.strptime(all_dates[0], "%d-%m-%Y")
    end_date = datetime.strptime(all_dates[-1], "%d-%m-%Y")
    
    # Fill in empty days in the grid
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
    
    # Group OE subjects by type and date
    oe_data_copy = oe_data.copy()
    oe_data_copy['Exam Date'] = oe_data_copy['Exam Date'].apply(
        lambda x: x.strftime("%d-%m-%Y") if isinstance(x, pd.Timestamp) else str(x)
    )
    
    # Get current OE1/OE5 and OE2 dates
    oe1_oe5_data = oe_data_copy[oe_data_copy['OE'].isin(['OE1', 'OE5'])]
    oe2_data = oe_data_copy[oe_data_copy['OE'] == 'OE2']
    
    moves_made = 0
    optimization_log = []
    
    # NEW: Look for gaps that can be filled by moving OE subjects
    def find_gaps_before_oe():
        """Find gaps in the schedule before OE subjects"""
        # Get all non-OE exam dates
        non_oe_dates = []
        for date_str in schedule_grid:
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
            for time_slot in schedule_grid[date_str]:
                for branch in schedule_grid[date_str][time_slot]:
                    subject = schedule_grid[date_str][time_slot][branch]
                    if subject and not any(oe in subject for oe in ['OE1', 'OE2', 'OE5']):
                        non_oe_dates.append(date_obj.date())
                        break
        
        if not non_oe_dates:
            return []
        
        unique_non_oe_dates = sorted(set(non_oe_dates))
        last_non_oe_date = max(unique_non_oe_dates)
        
        # Get current OE dates
        oe_dates = []
        if not oe1_oe5_data.empty:
            oe_date = datetime.strptime(oe1_oe5_data['Exam Date'].iloc[0], "%d-%m-%Y").date()
            oe_dates.append(oe_date)
        if not oe2_data.empty:
            oe_date = datetime.strptime(oe2_data['Exam Date'].iloc[0], "%d-%m-%Y").date()
            oe_dates.append(oe_date)
        
        if not oe_dates:
            return []
        
        first_oe_date = min(oe_dates)
        
        # Count gap days between last non-OE and first OE
        gap_days = 0
        check_date = last_non_oe_date + timedelta(days=1)
        available_dates = []
        
        while check_date < first_oe_date:
            if check_date.weekday() != 6 and check_date not in holidays:
                gap_days += 1
                available_dates.append(check_date)
            check_date += timedelta(days=1)
        
        return available_dates if gap_days > 2 else []
    
    # Check for gaps that can be filled
    available_gap_dates = find_gaps_before_oe()
    
    # Process OE1/OE5 together with gap consideration
    if not oe1_oe5_data.empty and available_gap_dates:
        # Current OE1/OE5 date
        current_oe1_oe5_date = oe1_oe5_data['Exam Date'].iloc[0]
        current_oe1_oe5_slot = oe1_oe5_data['Time Slot'].iloc[0]
        current_oe1_oe5_date_obj = datetime.strptime(current_oe1_oe5_date, "%d-%m-%Y")
        
        affected_branches = oe1_oe5_data['Branch'].unique()
        
        # Try to move OE1/OE5 to fill a gap
        for gap_date in available_gap_dates:
            gap_date_str = gap_date.strftime("%d-%m-%Y")
            
            # Check if this date can accommodate all OE1/OE5 branches
            for time_slot in ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]:
                can_move_oe1_oe5 = True
                
                # Check if this slot is empty for all OE1/OE5 branches
                for branch in affected_branches:
                    if (gap_date_str in schedule_grid and 
                        time_slot in schedule_grid[gap_date_str] and
                        branch in schedule_grid[gap_date_str][time_slot] and
                        schedule_grid[gap_date_str][time_slot][branch] is not None):
                        can_move_oe1_oe5 = False
                        break
                
                if can_move_oe1_oe5:
                    # Check if OE2 can be scheduled on the next day after this gap date
                    next_day = gap_date + timedelta(days=1)
                    while next_day.weekday() == 6 or next_day in holidays:
                        next_day += timedelta(days=1)
                    
                    next_day_str = next_day.strftime("%d-%m-%Y")
                    
                    if not oe2_data.empty:
                        oe2_branches = oe2_data['Branch'].unique()
                        can_move_oe2 = False
                        
                        # Check both time slots for OE2 on the next day
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
                            # Move both OE1/OE5 and OE2
                            days_saved = (current_oe1_oe5_date_obj - datetime.strptime(gap_date_str, "%d-%m-%Y")).days
                            
                            # Update all OE1/OE5 exams
                            for idx in oe1_oe5_data.index:
                                sem = all_data.at[idx, 'Semester']
                                branch = all_data.at[idx, 'Branch']
                                subject = all_data.at[idx, 'Subject']
                                
                                # Update in the semester dictionary
                                mask = (sem_dict[sem]['Subject'] == subject) & \
                                       (sem_dict[sem]['Branch'] == branch)
                                sem_dict[sem].loc[mask, 'Exam Date'] = gap_date_str
                                sem_dict[sem].loc[mask, 'Time Slot'] = time_slot
                                
                                # Update schedule grid
                                if (current_oe1_oe5_date in schedule_grid and 
                                    current_oe1_oe5_slot in schedule_grid[current_oe1_oe5_date] and
                                    branch in schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot]):
                                    schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot][branch] = None
                                
                                if gap_date_str not in schedule_grid:
                                    schedule_grid[gap_date_str] = {}
                                if time_slot not in schedule_grid[gap_date_str]:
                                    schedule_grid[gap_date_str][time_slot] = {}
                                schedule_grid[gap_date_str][time_slot][branch] = subject
                            
                            # Update all OE2 exams
                            current_oe2_date = oe2_data['Exam Date'].iloc[0]
                            current_oe2_slot = oe2_data['Time Slot'].iloc[0]
                            
                            for idx in oe2_data.index:
                                sem = all_data.at[idx, 'Semester']
                                branch = all_data.at[idx, 'Branch']
                                subject = all_data.at[idx, 'Subject']
                                
                                # Update in the semester dictionary
                                mask = (sem_dict[sem]['Subject'] == subject) & \
                                       (sem_dict[sem]['Branch'] == branch)
                                sem_dict[sem].loc[mask, 'Exam Date'] = next_day_str
                                sem_dict[sem].loc[mask, 'Time Slot'] = best_oe2_slot
                                
                                # Update schedule grid
                                if (current_oe2_date in schedule_grid and 
                                    current_oe2_slot in schedule_grid[current_oe2_date] and
                                    branch in schedule_grid[current_oe2_date][current_oe2_slot]):
                                    schedule_grid[current_oe2_date][current_oe2_slot][branch] = None
                                
                                if next_day_str not in schedule_grid:
                                    schedule_grid[next_day_str] = {}
                                if best_oe2_slot not in schedule_grid[next_day_str]:
                                    schedule_grid[next_day_str][best_oe2_slot] = {}
                                schedule_grid[next_day_str][best_oe2_slot][branch] = subject
                            
                            moves_made += 1
                            optimization_log.append(
                                f"🎯 Moved OE1/OE5 to fill gap: {current_oe1_oe5_date} → {gap_date_str} (saved {days_saved} days)"
                            )
                            optimization_log.append(
                                f"🎯 Moved OE2 to maintain sequence: {current_oe2_date} → {next_day_str}"
                            )
                            
                            # Only move once
                            break
                    else:
                        # No OE2 to worry about, just move OE1/OE5
                        days_saved = (current_oe1_oe5_date_obj - datetime.strptime(gap_date_str, "%d-%m-%Y")).days
                        
                        # Update all OE1/OE5 exams
                        for idx in oe1_oe5_data.index:
                            sem = all_data.at[idx, 'Semester']
                            branch = all_data.at[idx, 'Branch']
                            subject = all_data.at[idx, 'Subject']
                            
                            # Update in the semester dictionary
                            mask = (sem_dict[sem]['Subject'] == subject) & \
                                   (sem_dict[sem]['Branch'] == branch)
                            sem_dict[sem].loc[mask, 'Exam Date'] = gap_date_str
                            sem_dict[sem].loc[mask, 'Time Slot'] = time_slot
                            
                            # Update schedule grid
                            if (current_oe1_oe5_date in schedule_grid and 
                                current_oe1_oe5_slot in schedule_grid[current_oe1_oe5_date] and
                                branch in schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot]):
                                schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot][branch] = None
                            
                            if gap_date_str not in schedule_grid:
                                schedule_grid[gap_date_str] = {}
                            if time_slot not in schedule_grid[gap_date_str]:
                                schedule_grid[gap_date_str][time_slot] = {}
                            schedule_grid[gap_date_str][time_slot][branch] = subject
                        
                        moves_made += 1
                        optimization_log.append(
                            f"🎯 Moved OE1/OE5 to fill gap: {current_oe1_oe5_date} → {gap_date_str} (saved {days_saved} days)"
                        )
                        break
                
                if moves_made > 0:
                    break
            
            if moves_made > 0:
                break
    
    if moves_made > 0:
        st.success(f"✅ OE Gap Optimization: Filled {len(available_gap_dates)} day gap by moving {moves_made} OE groups!")
        with st.expander("📝 OE Gap Optimization Details"):
            for log in optimization_log:
                st.write(f"• {log}")
    else:
        st.info("ℹ️ OE subjects are optimally placed - no beneficial gaps found")
    
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
                    df_non_elec["SubjectDisplay"] = df_non_elec["Subject"]
                    duration_suffix = df_non_elec.apply(
                        lambda row: f" [Duration: {row['Exam Duration']} hrs]" if row['Exam Duration'] != 3 else '', axis=1)
                    df_non_elec["SubjectDisplay"] = df_non_elec["SubjectDisplay"] + difficulty_suffix + duration_suffix
                    df_non_elec["Exam Date"] = pd.to_datetime(
                        df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce'
                    )
                    df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                    
                    # FIXED: First group by Exam Date and SubBranch to combine all subjects for that date/branch
                    # This handles cases where a branch has subjects in both morning and afternoon slots
                    consolidated_df = df_non_elec.groupby(['Exam Date', 'SubBranch']).agg({
                        'SubjectDisplay': lambda x: ", ".join(str(i) for i in x),
                        'Time Slot': lambda x: list(set(x))[0]  # Take any time slot since we're consolidating
                    }).reset_index()
                    
                    # Now group by just Exam Date to ensure only one row per date
                    # Keep all SubBranch data but ensure single time slot per date
                    date_consolidated = consolidated_df.groupby('Exam Date').agg({
                        'Time Slot': 'first'  # Use first time slot found for this date
                    }).reset_index()
                    
                    # Merge back to get the consolidated subject data
                    final_df = consolidated_df.merge(date_consolidated[['Exam Date', 'Time Slot']], 
                                                   on=['Exam Date', 'Time Slot'], how='inner')
                    
                    # If there are still subjects on the same date with different time slots,
                    # we need to group them under the primary time slot
                    remaining_subjects = consolidated_df[~consolidated_df.index.isin(final_df.index)]
                    if not remaining_subjects.empty:
                        # For remaining subjects, group them by date and subbranch and add to final_df
                        for date in remaining_subjects['Exam Date'].unique():
                            date_subjects = remaining_subjects[remaining_subjects['Exam Date'] == date]
                            primary_time_slot = date_consolidated[date_consolidated['Exam Date'] == date]['Time Slot'].iloc[0]
                            
                            for _, row in date_subjects.iterrows():
                                # Check if this SubBranch already exists for this date in final_df
                                existing_mask = (final_df['Exam Date'] == date) & (final_df['SubBranch'] == row['SubBranch'])
                                if existing_mask.any():
                                    # Combine with existing subjects
                                    existing_idx = final_df[existing_mask].index[0]
                                    final_df.at[existing_idx, 'SubjectDisplay'] = (
                                        final_df.at[existing_idx, 'SubjectDisplay'] + ", " + row['SubjectDisplay']
                                    )
                                else:
                                    # Add new row with primary time slot
                                    new_row = row.copy()
                                    new_row['Time Slot'] = primary_time_slot
                                    final_df = pd.concat([final_df, new_row.to_frame().T], ignore_index=True)
                    
                    # Create pivot table with consolidated data
                    pivot_df = final_df.pivot_table(
                        index=["Exam Date", "Time Slot"],
                        columns="SubBranch",
                        values="SubjectDisplay",
                        aggfunc='first'
                    ).fillna("---")
                    
                    pivot_df = pivot_df.sort_index(ascending=True)
                    # Format dates in the multi-level index
                    if len(pivot_df.index.levels) > 0:
                        formatted_dates = [d.strftime("%d-%m-%Y") for d in pivot_df.index.levels[0]]
                        pivot_df.index = pivot_df.index.set_levels(formatted_dates, level=0)
                    roman_sem = int_to_roman(sem)
                    sheet_name = f"{main_branch}_Sem_{roman_sem}"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    pivot_df.to_excel(writer, sheet_name=sheet_name)

                # Process electives in a separate sheet (unchanged)
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
                <li>📅 Smart scheduling with gap optimization</li>
                <li>📋 Multiple output formats</li>
                <li>🎯 Conflict resolution</li>
                <li>📱 Mobile-friendly interface</li>
                <li>🔧 Real-time optimization</li>
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
                        st.write("Processing constraints with enhanced gap optimization...")
                        non_elec_sched = enhanced_process_constraints_with_gap_optimization(df_non_elec, holidays_set, base_date, schedule_by_difficulty)
                        st.write(f"non_elec_sched keys: {list(non_elec_sched.keys())}")

                        # Find the maximum date from non-elective exams
                        non_elec_df = pd.concat(non_elec_sched.values(), ignore_index=True) if non_elec_sched else pd.DataFrame()
                        non_elec_dates = pd.to_datetime(non_elec_df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        max_non_elec_date = max(non_elec_dates).date() if not non_elec_dates.empty else base_date.date()
                        st.write(f"Max non-elective date: {max_non_elec_date}")

                        # Schedule electives globally only if df_ele is not None
                        if df_ele is not None and not df_ele.empty:
                            st.write("Scheduling electives...")
                            elective_day1 = find_next_valid_day_for_electives(datetime.combine(max_non_elec_date, datetime.min.time()) + timedelta(days=1), holidays_set)
                            elective_day2 = find_next_valid_day_for_electives(elective_day1 + timedelta(days=1), holidays_set)
    
                            # CRITICAL FIX: OE1 and OE5 must be scheduled together on the same date/time
                            # Schedule OE1 and OE5 together on the first elective day
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Exam Date'] = elective_day1.strftime("%d-%m-%Y")
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Time Slot'] = "10:00 AM - 1:00 PM"
    
                            # Schedule OE2 on the second elective day (afternoon slot)
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Exam Date'] = elective_day2.strftime("%d-%m-%Y")
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Time Slot'] = "2:00 PM - 5:00 PM"
    
                            st.write(f"✅ OE1 and OE5 scheduled together on {elective_day1.strftime('%d-%m-%Y')} at 10:00 AM - 1:00 PM")
                            st.write(f"✅ OE2 scheduled on {elective_day2.strftime('%d-%m-%Y')} at 2:00 PM - 5:00 PM")

                            # Combine non-electives and electives
                            final_df = pd.concat([non_elec_df, df_ele], ignore_index=True)
                        else:
                            final_df = non_elec_df
                            st.write("No electives to schedule.")

                        final_df["Exam Date"] = pd.to_datetime(final_df["Exam Date"], format="%d-%m-%Y", errors='coerce')
                        final_df = final_df.sort_values(["Exam Date", "Semester", "MainBranch"], ascending=True, na_position='last')
                        sem_dict = {s: final_df[final_df["Semester"] == s].copy() for s in sorted(final_df["Semester"].unique())}
                        
                        # Enhanced OE optimization
                        sem_dict = enhanced_optimize_oe_subjects_after_scheduling(sem_dict, holidays_set)
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

                        st.markdown('<div class="status-success">🎉 Timetable generated successfully with enhanced gap optimization!</div>',
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
        <p>🎓 <strong>Enhanced Exam Timetable Generator with Gap Optimization</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
        <p style="font-size: 0.9em;">Advanced scheduling • Gap reduction • Conflict-free timetables • Multiple export formats</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main())
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

def schedule_semester_non_electives_with_optimization(df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty=False):
    """Enhanced scheduling that fills empty slots in real-time and ensures all subjects are scheduled"""
    
    # Get semester time slot
    sem = df_sem["Semester"].iloc[0]
    if sem % 2 != 0:
        odd_sem_position = (sem + 1) // 2
        preferred_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:
        even_sem_position = sem // 2
        preferred_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    
    # Schedule COMP subjects
    comp_subjects = df_sem[(df_sem['Category'] == 'COMP') & (df_sem['IsCommon'] == 'NO') & (df_sem['Exam Date'] == "")]
    
    for idx, row in comp_subjects.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        
        # First, try to find an empty slot
        empty_date, empty_slot = optimizer.find_earliest_empty_slot(branch, base_date, preferred_slot)
        
        if empty_date:
            # Found an empty slot - use it
            df_sem.at[idx, 'Exam Date'] = empty_date
            df_sem.at[idx, 'Time Slot'] = empty_slot
            optimizer.add_exam_to_grid(empty_date, empty_slot, branch, subject)
            exam_days[branch].add(datetime.strptime(empty_date, "%d-%m-%Y").date())
            optimizer.optimization_log.append(f"✅ Filled empty slot: {subject} on {empty_date} at {empty_slot}")
            optimizer.moves_made += 1
        else:
            # No empty slot found - schedule normally
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
    
    # Schedule ELEC subjects
    elec_subjects = df_sem[(df_sem['Category'] == 'ELEC') & (df_sem['IsCommon'] == 'NO') & (df_sem['Exam Date'] == "")]
    
    for idx, row in elec_subjects.iterrows():
        branch = row['Branch']
        subject = row['Subject']
        
        # Try to find empty slot first
        empty_date, empty_slot = optimizer.find_earliest_empty_slot(branch, base_date, preferred_slot)
        
        if empty_date:
            # Found an empty slot - use it
            df_sem.at[idx, 'Exam Date'] = empty_date
            df_sem.at[idx, 'Time Slot'] = empty_slot
            optimizer.add_exam_to_grid(empty_date, empty_slot, branch, subject)
            exam_days[branch].add(datetime.strptime(empty_date, "%d-%m-%Y").date())
            optimizer.optimization_log.append(f"✅ Filled empty slot: ELEC {subject} on {empty_date} at {empty_slot}")
            optimizer.moves_made += 1
        else:
            # No empty slot found - find next available day with NO EXAM for this branch
            current_date = base_date
            scheduled = False
            
            while not scheduled:
                date_str = current_date.strftime("%d-%m-%Y")
                
                if current_date.weekday() != 6 and current_date.date() not in holidays:
                    # Only schedule if this branch has NO exam on this date
                    if current_date.date() not in exam_days[branch]:
                        df_sem.at[idx, 'Exam Date'] = date_str
                        df_sem.at[idx, 'Time Slot'] = preferred_slot
                        optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                        exam_days[branch].add(current_date.date())
                        scheduled = True
                        optimizer.optimization_log.append(f"📅 Scheduled ELEC {subject} for {branch} on {date_str}")
                
                current_date += timedelta(days=1)
    
    # Assign time slot to any remaining exams (safety net)
    df_sem.loc[df_sem['Time Slot'] == "", 'Time Slot'] = preferred_slot
    
    # Final check - ensure no exam is left unscheduled
    unscheduled = df_sem[df_sem['Exam Date'] == ""]
    if not unscheduled.empty:
        st.warning(f"⚠️ {len(unscheduled)} subjects remain unscheduled in semester {sem}")
        # Force schedule them on empty days
        for idx, row in unscheduled.iterrows():
            branch = row['Branch']
            subject = row['Subject']
            current_date = base_date
            
            while True:
                date_str = current_date.strftime("%d-%m-%Y")
                if current_date.weekday() != 6 and current_date.date() not in holidays:
                    # Force schedule only on days with no exams for this branch
                    if current_date.date() not in exam_days[branch]:
                        df_sem.at[idx, 'Exam Date'] = date_str
                        df_sem.at[idx, 'Time Slot'] = preferred_slot
                        optimizer.add_exam_to_grid(date_str, preferred_slot, branch, subject)
                        exam_days[branch].add(current_date.date())
                        optimizer.optimization_log.append(f"🔧 Force scheduled {subject} on {date_str}")
                        break
                current_date += timedelta(days=1)
    
    return df_sem

def enhanced_process_constraints_with_gap_optimization(df, holidays, base_date, schedule_by_difficulty=False):
    """
    Enhanced process_constraints that reduces gaps between exams
    """
    # Initialize exam_days for all branches
    all_branches = df['Branch'].unique()
    exam_days = {branch: set() for branch in all_branches}
    
    # Initialize the enhanced optimizer
    optimizer = EnhancedRealTimeOptimizer(all_branches, holidays)
    optimizer.initialize_grid_with_empty_days(base_date, num_days=40)
    
    st.info(f"🔧 Scheduling {len(df)} subjects across {len(all_branches)} branches...")
    
    # Helper function for finding earliest slot
    def find_earliest_available_slot_with_empty_check(start_day, for_branches, subject, optimizer):
        """Enhanced slot finding that checks empty slots first"""
        
        # First, check if there's a common empty slot for all branches
        if len(for_branches) == 1:
            # Single branch - use optimizer
            empty_date, empty_slot = optimizer.find_earliest_empty_slot(for_branches[0], start_day)
            if empty_date:
                return datetime.strptime(empty_date, "%d-%m-%Y"), empty_slot
        else:
            # Multiple branches - find common empty slot
            sorted_dates = sorted(optimizer.schedule_grid.keys(), 
                                key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
            
            for date_str in sorted_dates:
                date_obj = datetime.strptime(date_str, "%d-%m-%Y")
                if date_obj < start_day:
                    continue
                    
                for time_slot in optimizer.time_slots:
                    # Check if slot is empty for all branches
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
        
        # No empty slot found - use normal scheduling
        current_date = start_day
        while True:
            current_date_only = current_date.date()
            
            if current_date.weekday() == 6 or current_date_only in holidays:
                current_date += timedelta(days=1)
                continue
            
            if all(current_date_only not in exam_days[branch] for branch in for_branches):
                return current_date, None
            
            current_date += timedelta(days=1)

    # Count subjects by category
    comp_common = len(df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'YES')])
    comp_individual = len(df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'NO')])
    elec_common = len(df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'YES')])
    elec_individual = len(df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'NO')])
    
    st.write(f"📊 Subject distribution: COMP (Common: {comp_common}, Individual: {comp_individual}), ELEC (Common: {elec_common}, Individual: {elec_individual})")

    # Schedule common COMP subjects with optimization
    common_comp = df[(df['Category'] == 'COMP') & (df['IsCommon'] == 'YES')]
    for module_code, group in common_comp.groupby('ModuleCode'):
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

    # Schedule common ELEC subjects with optimization
    common_elec = df[(df['Category'] == 'ELEC') & (df['IsCommon'] == 'YES')]
    for module_code, group in common_elec.groupby('ModuleCode'):
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

    # Schedule remaining subjects per semester with optimization
    final_list = []
    for sem in sorted(df["Semester"].unique()):
        if sem == 0:
            continue
        df_sem = df[df["Semester"] == sem].copy()
        if df_sem.empty:
            continue
            
        # Count unscheduled subjects before processing
        unscheduled_before = len(df_sem[df_sem['Exam Date'] == ""])
        
        scheduled_sem = schedule_semester_non_electives_with_optimization(
            df_sem, holidays, base_date, exam_days, optimizer, schedule_by_difficulty
        )
        
        # Count unscheduled subjects after processing
        unscheduled_after = len(scheduled_sem[scheduled_sem['Exam Date'] == ""])
        
        if unscheduled_after > 0:
            st.warning(f"⚠️ Semester {sem}: {unscheduled_after} subjects remain unscheduled out of {unscheduled_before}")
        else:
            st.success(f"✅ Semester {sem}: All {unscheduled_before} subjects scheduled successfully")
        
        final_list.append(scheduled_sem)

    if not final_list:
        return {}

    df_combined = pd.concat(final_list, ignore_index=True)
    
    # NEW: Perform gap optimization after initial scheduling
    st.markdown("---")
    st.info("🎯 Starting gap optimization phase...")
    
    # First, populate the optimizer grid with current schedule
    for idx, row in df_combined.iterrows():
        if pd.notna(row['Exam Date']) and row['Exam Date'] != "":
            date_str = row['Exam Date']
            time_slot = row['Time Slot']
            branch = row['Branch']
            subject = row['Subject']
            optimizer.add_exam_to_grid(date_str, time_slot, branch, subject)
    
    # Calculate initial span
    initial_dates = pd.to_datetime(df_combined['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    initial_span = (max(initial_dates) - min(initial_dates)).days + 1 if not initial_dates.empty else 0
    st.info(f"📊 Initial exam span: {initial_span} days")
    
    # Perform gap optimization
    optimizer.optimize_gaps()
    
    # Update DataFrame with optimized schedule
    df_combined = optimizer.update_dataframe_from_grid(df_combined)
    
    # Check for any unscheduled subjects
    unscheduled_final = df_combined[df_combined['Exam Date'] == ""]
    if not unscheduled_final.empty:
        st.error(f"❌ {len(unscheduled_final)} subjects remain unscheduled!")
        with st.expander("View unscheduled subjects"):
            st.dataframe(unscheduled_final[['Branch', 'Subject', 'Category', 'IsCommon']])
    
    # Display optimization summary
    schedule_summary = optimizer.get_schedule_summary()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Optimization Moves", optimizer.moves_made)
    with col2:
        st.metric("Gaps Reduced", schedule_summary['gap_reductions'])
    with col3:
        st.metric("Final Exam Span", f"{schedule_summary['optimized_span']} days")
    with col4:
        span_reduction = initial_span - schedule_summary['optimized_span']
        st.metric("Days Saved", span_reduction, delta=f"-{span_reduction}")
    
    if optimizer.moves_made > 0:
        with st.expander("📝 Gap Optimization Log", expanded=False):
            for log in optimizer.optimization_log:
                st.write(log)
    
    # Rest of the function remains the same
    sem_dict = {}
    for sem in sorted(df_combined["Semester"].unique()):
        sem_dict[sem] = df_combined[df_combined["Semester"] == sem].copy()

    # Calculate final span and provide feedback
    final_dates = pd.to_datetime(df_combined['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
    if not final_dates.empty:
        start_date = min(final_dates)
        end_date = max(final_dates)
        final_span = (end_date - start_date).days + 1
        
        if final_span <= 14:
            st.success(f"🎉 Excellent optimization! Final span: {final_span} days (target achieved)")
        elif final_span <= 16:
            st.success(f"✅ Great optimization! Final span: {final_span} days (within 16-day target)")
        elif final_span <= 20:
            st.info(f"✅ Good optimization! Final span: {final_span} days (within 20-day limit)")
        else:
            st.warning(f"⚠️ The timetable spans {final_span} days, still exceeding the 20-day limit.")

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

def enhanced_optimize_oe_subjects_after_scheduling(sem_dict, holidays, optimizer=None):
    """
    Enhanced OE optimization that considers gap reduction
    """
    if not sem_dict:
        return sem_dict
    
    st.info("🎯 Optimizing Open Elective (OE) placement with gap reduction...")
    
    # Combine all data to analyze the schedule
    all_data = pd.concat(sem_dict.values(), ignore_index=True)
    
    # Separate OE and non-OE data
    oe_data = all_data[all_data['OE'].notna() & (all_data['OE'].str.strip() != "")]
    non_oe_data = all_data[all_data['OE'].isna() | (all_data['OE'].str.strip() == "")]
    
    if oe_data.empty:
        st.info("No OE subjects to optimize")
        return sem_dict
    
    # Build complete schedule grid from current state
    schedule_grid = {}
    branches = all_data['Branch'].unique()
    
    # First, populate with all scheduled exams
    for _, row in all_data.iterrows():
        if pd.notna(row['Exam Date']):
            # Handle both string and Timestamp types for Exam Date
            if isinstance(row['Exam Date'], pd.Timestamp):
                date_str = row['Exam Date'].strftime("%d-%m-%Y")
            else:
                date_str = str(row['Exam Date'])
            
            if date_str not in schedule_grid:
                schedule_grid[date_str] = {}
            if row['Time Slot'] not in schedule_grid[date_str]:
                schedule_grid[date_str][row['Time Slot']] = {}
            schedule_grid[date_str][row['Time Slot']][row['Branch']] = row['Subject']
    
    # Find all dates in the schedule
    all_dates = sorted(schedule_grid.keys(), 
                      key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    if not all_dates:
        return sem_dict
    
    # Get date range
    start_date = datetime.strptime(all_dates[0], "%d-%m-%Y")
    end_date = datetime.strptime(all_dates[-1], "%d-%m-%Y")
    
    # Fill in empty days in the grid
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
    
    # Group OE subjects by type and date
    oe_data_copy = oe_data.copy()
    oe_data_copy['Exam Date'] = oe_data_copy['Exam Date'].apply(
        lambda x: x.strftime("%d-%m-%Y") if isinstance(x, pd.Timestamp) else str(x)
    )
    
    # Get current OE1/OE5 and OE2 dates
    oe1_oe5_data = oe_data_copy[oe_data_copy['OE'].isin(['OE1', 'OE5'])]
    oe2_data = oe_data_copy[oe_data_copy['OE'] == 'OE2']
    
    moves_made = 0
    optimization_log = []
    
    # NEW: Look for gaps that can be filled by moving OE subjects
    def find_gaps_before_oe():
        """Find gaps in the schedule before OE subjects"""
        # Get all non-OE exam dates
        non_oe_dates = []
        for date_str in schedule_grid:
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
            for time_slot in schedule_grid[date_str]:
                for branch in schedule_grid[date_str][time_slot]:
                    subject = schedule_grid[date_str][time_slot][branch]
                    if subject and not any(oe in subject for oe in ['OE1', 'OE2', 'OE5']):
                        non_oe_dates.append(date_obj.date())
                        break
        
        if not non_oe_dates:
            return []
        
        unique_non_oe_dates = sorted(set(non_oe_dates))
        last_non_oe_date = max(unique_non_oe_dates)
        
        # Get current OE dates
        oe_dates = []
        if not oe1_oe5_data.empty:
            oe_date = datetime.strptime(oe1_oe5_data['Exam Date'].iloc[0], "%d-%m-%Y").date()
            oe_dates.append(oe_date)
        if not oe2_data.empty:
            oe_date = datetime.strptime(oe2_data['Exam Date'].iloc[0], "%d-%m-%Y").date()
            oe_dates.append(oe_date)
        
        if not oe_dates:
            return []
        
        first_oe_date = min(oe_dates)
        
        # Count gap days between last non-OE and first OE
        gap_days = 0
        check_date = last_non_oe_date + timedelta(days=1)
        available_dates = []
        
        while check_date < first_oe_date:
            if check_date.weekday() != 6 and check_date not in holidays:
                gap_days += 1
                available_dates.append(check_date)
            check_date += timedelta(days=1)
        
        return available_dates if gap_days > 2 else []
    
    # Check for gaps that can be filled
    available_gap_dates = find_gaps_before_oe()
    
    # Process OE1/OE5 together with gap consideration
    if not oe1_oe5_data.empty and available_gap_dates:
        # Current OE1/OE5 date
        current_oe1_oe5_date = oe1_oe5_data['Exam Date'].iloc[0]
        current_oe1_oe5_slot = oe1_oe5_data['Time Slot'].iloc[0]
        current_oe1_oe5_date_obj = datetime.strptime(current_oe1_oe5_date, "%d-%m-%Y")
        
        affected_branches = oe1_oe5_data['Branch'].unique()
        
        # Try to move OE1/OE5 to fill a gap
        for gap_date in available_gap_dates:
            gap_date_str = gap_date.strftime("%d-%m-%Y")
            
            # Check if this date can accommodate all OE1/OE5 branches
            for time_slot in ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]:
                can_move_oe1_oe5 = True
                
                # Check if this slot is empty for all OE1/OE5 branches
                for branch in affected_branches:
                    if (gap_date_str in schedule_grid and 
                        time_slot in schedule_grid[gap_date_str] and
                        branch in schedule_grid[gap_date_str][time_slot] and
                        schedule_grid[gap_date_str][time_slot][branch] is not None):
                        can_move_oe1_oe5 = False
                        break
                
                if can_move_oe1_oe5:
                    # Check if OE2 can be scheduled on the next day after this gap date
                    next_day = gap_date + timedelta(days=1)
                    while next_day.weekday() == 6 or next_day in holidays:
                        next_day += timedelta(days=1)
                    
                    next_day_str = next_day.strftime("%d-%m-%Y")
                    
                    if not oe2_data.empty:
                        oe2_branches = oe2_data['Branch'].unique()
                        can_move_oe2 = False
                        
                        # Check both time slots for OE2 on the next day
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
                            # Move both OE1/OE5 and OE2
                            days_saved = (current_oe1_oe5_date_obj - datetime.strptime(gap_date_str, "%d-%m-%Y")).days
                            
                            # Update all OE1/OE5 exams
                            for idx in oe1_oe5_data.index:
                                sem = all_data.at[idx, 'Semester']
                                branch = all_data.at[idx, 'Branch']
                                subject = all_data.at[idx, 'Subject']
                                
                                # Update in the semester dictionary
                                mask = (sem_dict[sem]['Subject'] == subject) & \
                                       (sem_dict[sem]['Branch'] == branch)
                                sem_dict[sem].loc[mask, 'Exam Date'] = gap_date_str
                                sem_dict[sem].loc[mask, 'Time Slot'] = time_slot
                                
                                # Update schedule grid
                                if (current_oe1_oe5_date in schedule_grid and 
                                    current_oe1_oe5_slot in schedule_grid[current_oe1_oe5_date] and
                                    branch in schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot]):
                                    schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot][branch] = None
                                
                                if gap_date_str not in schedule_grid:
                                    schedule_grid[gap_date_str] = {}
                                if time_slot not in schedule_grid[gap_date_str]:
                                    schedule_grid[gap_date_str][time_slot] = {}
                                schedule_grid[gap_date_str][time_slot][branch] = subject
                            
                            # Update all OE2 exams
                            current_oe2_date = oe2_data['Exam Date'].iloc[0]
                            current_oe2_slot = oe2_data['Time Slot'].iloc[0]
                            
                            for idx in oe2_data.index:
                                sem = all_data.at[idx, 'Semester']
                                branch = all_data.at[idx, 'Branch']
                                subject = all_data.at[idx, 'Subject']
                                
                                # Update in the semester dictionary
                                mask = (sem_dict[sem]['Subject'] == subject) & \
                                       (sem_dict[sem]['Branch'] == branch)
                                sem_dict[sem].loc[mask, 'Exam Date'] = next_day_str
                                sem_dict[sem].loc[mask, 'Time Slot'] = best_oe2_slot
                                
                                # Update schedule grid
                                if (current_oe2_date in schedule_grid and 
                                    current_oe2_slot in schedule_grid[current_oe2_date] and
                                    branch in schedule_grid[current_oe2_date][current_oe2_slot]):
                                    schedule_grid[current_oe2_date][current_oe2_slot][branch] = None
                                
                                if next_day_str not in schedule_grid:
                                    schedule_grid[next_day_str] = {}
                                if best_oe2_slot not in schedule_grid[next_day_str]:
                                    schedule_grid[next_day_str][best_oe2_slot] = {}
                                schedule_grid[next_day_str][best_oe2_slot][branch] = subject
                            
                            moves_made += 1
                            optimization_log.append(
                                f"🎯 Moved OE1/OE5 to fill gap: {current_oe1_oe5_date} → {gap_date_str} (saved {days_saved} days)"
                            )
                            optimization_log.append(
                                f"🎯 Moved OE2 to maintain sequence: {current_oe2_date} → {next_day_str}"
                            )
                            
                            # Only move once
                            break
                    else:
                        # No OE2 to worry about, just move OE1/OE5
                        days_saved = (current_oe1_oe5_date_obj - datetime.strptime(gap_date_str, "%d-%m-%Y")).days
                        
                        # Update all OE1/OE5 exams
                        for idx in oe1_oe5_data.index:
                            sem = all_data.at[idx, 'Semester']
                            branch = all_data.at[idx, 'Branch']
                            subject = all_data.at[idx, 'Subject']
                            
                            # Update in the semester dictionary
                            mask = (sem_dict[sem]['Subject'] == subject) & \
                                   (sem_dict[sem]['Branch'] == branch)
                            sem_dict[sem].loc[mask, 'Exam Date'] = gap_date_str
                            sem_dict[sem].loc[mask, 'Time Slot'] = time_slot
                            
                            # Update schedule grid
                            if (current_oe1_oe5_date in schedule_grid and 
                                current_oe1_oe5_slot in schedule_grid[current_oe1_oe5_date] and
                                branch in schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot]):
                                schedule_grid[current_oe1_oe5_date][current_oe1_oe5_slot][branch] = None
                            
                            if gap_date_str not in schedule_grid:
                                schedule_grid[gap_date_str] = {}
                            if time_slot not in schedule_grid[gap_date_str]:
                                schedule_grid[gap_date_str][time_slot] = {}
                            schedule_grid[gap_date_str][time_slot][branch] = subject
                        
                        moves_made += 1
                        optimization_log.append(
                            f"🎯 Moved OE1/OE5 to fill gap: {current_oe1_oe5_date} → {gap_date_str} (saved {days_saved} days)"
                        )
                        break
                
                if moves_made > 0:
                    break
            
            if moves_made > 0:
                break
    
    if moves_made > 0:
        st.success(f"✅ OE Gap Optimization: Filled {len(available_gap_dates)} day gap by moving {moves_made} OE groups!")
        with st.expander("📝 OE Gap Optimization Details"):
            for log in optimization_log:
                st.write(f"• {log}")
    else:
        st.info("ℹ️ OE subjects are optimally placed - no beneficial gaps found")
    
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
                    df_non_elec["SubjectDisplay"] = df_non_elec["Subject"]
                    duration_suffix = df_non_elec.apply(
                        lambda row: f" [Duration: {row['Exam Duration']} hrs]" if row['Exam Duration'] != 3 else '', axis=1)
                    df_non_elec["SubjectDisplay"] = df_non_elec["SubjectDisplay"] + difficulty_suffix + duration_suffix
                    df_non_elec["Exam Date"] = pd.to_datetime(
                        df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce'
                    )
                    df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                    
                    # FIXED: First group by Exam Date and SubBranch to combine all subjects for that date/branch
                    # This handles cases where a branch has subjects in both morning and afternoon slots
                    consolidated_df = df_non_elec.groupby(['Exam Date', 'SubBranch']).agg({
                        'SubjectDisplay': lambda x: ", ".join(str(i) for i in x),
                        'Time Slot': lambda x: list(set(x))[0]  # Take any time slot since we're consolidating
                    }).reset_index()
                    
                    # Now group by just Exam Date to ensure only one row per date
                    # Keep all SubBranch data but ensure single time slot per date
                    date_consolidated = consolidated_df.groupby('Exam Date').agg({
                        'Time Slot': 'first'  # Use first time slot found for this date
                    }).reset_index()
                    
                    # Merge back to get the consolidated subject data
                    final_df = consolidated_df.merge(date_consolidated[['Exam Date', 'Time Slot']], 
                                                   on=['Exam Date', 'Time Slot'], how='inner')
                    
                    # If there are still subjects on the same date with different time slots,
                    # we need to group them under the primary time slot
                    remaining_subjects = consolidated_df[~consolidated_df.index.isin(final_df.index)]
                    if not remaining_subjects.empty:
                        # For remaining subjects, group them by date and subbranch and add to final_df
                        for date in remaining_subjects['Exam Date'].unique():
                            date_subjects = remaining_subjects[remaining_subjects['Exam Date'] == date]
                            primary_time_slot = date_consolidated[date_consolidated['Exam Date'] == date]['Time Slot'].iloc[0]
                            
                            for _, row in date_subjects.iterrows():
                                # Check if this SubBranch already exists for this date in final_df
                                existing_mask = (final_df['Exam Date'] == date) & (final_df['SubBranch'] == row['SubBranch'])
                                if existing_mask.any():
                                    # Combine with existing subjects
                                    existing_idx = final_df[existing_mask].index[0]
                                    final_df.at[existing_idx, 'SubjectDisplay'] = (
                                        final_df.at[existing_idx, 'SubjectDisplay'] + ", " + row['SubjectDisplay']
                                    )
                                else:
                                    # Add new row with primary time slot
                                    new_row = row.copy()
                                    new_row['Time Slot'] = primary_time_slot
                                    final_df = pd.concat([final_df, new_row.to_frame().T], ignore_index=True)
                    
                    # Create pivot table with consolidated data
                    pivot_df = final_df.pivot_table(
                        index=["Exam Date", "Time Slot"],
                        columns="SubBranch",
                        values="SubjectDisplay",
                        aggfunc='first'
                    ).fillna("---")
                    
                    pivot_df = pivot_df.sort_index(ascending=True)
                    # Format dates in the multi-level index
                    if len(pivot_df.index.levels) > 0:
                        formatted_dates = [d.strftime("%d-%m-%Y") for d in pivot_df.index.levels[0]]
                        pivot_df.index = pivot_df.index.set_levels(formatted_dates, level=0)
                    roman_sem = int_to_roman(sem)
                    sheet_name = f"{main_branch}_Sem_{roman_sem}"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    pivot_df.to_excel(writer, sheet_name=sheet_name)

                # Process electives in a separate sheet (unchanged)
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
                <li>📅 Smart scheduling with gap optimization</li>
                <li>📋 Multiple output formats</li>
                <li>🎯 Conflict resolution</li>
                <li>📱 Mobile-friendly interface</li>
                <li>🔧 Real-time optimization</li>
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
                        st.write("Processing constraints with enhanced gap optimization...")
                        non_elec_sched = enhanced_process_constraints_with_gap_optimization(df_non_elec, holidays_set, base_date, schedule_by_difficulty)
                        st.write(f"non_elec_sched keys: {list(non_elec_sched.keys())}")

                        # Find the maximum date from non-elective exams
                        non_elec_df = pd.concat(non_elec_sched.values(), ignore_index=True) if non_elec_sched else pd.DataFrame()
                        non_elec_dates = pd.to_datetime(non_elec_df['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                        max_non_elec_date = max(non_elec_dates).date() if not non_elec_dates.empty else base_date.date()
                        st.write(f"Max non-elective date: {max_non_elec_date}")

                        # Schedule electives globally only if df_ele is not None
                        if df_ele is not None and not df_ele.empty:
                            st.write("Scheduling electives...")
                            elective_day1 = find_next_valid_day_for_electives(datetime.combine(max_non_elec_date, datetime.min.time()) + timedelta(days=1), holidays_set)
                            elective_day2 = find_next_valid_day_for_electives(elective_day1 + timedelta(days=1), holidays_set)
    
                            # CRITICAL FIX: OE1 and OE5 must be scheduled together on the same date/time
                            # Schedule OE1 and OE5 together on the first elective day
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Exam Date'] = elective_day1.strftime("%d-%m-%Y")
                            df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Time Slot'] = "10:00 AM - 1:00 PM"
    
                            # Schedule OE2 on the second elective day (afternoon slot)
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Exam Date'] = elective_day2.strftime("%d-%m-%Y")
                            df_ele.loc[df_ele['OE'] == 'OE2', 'Time Slot'] = "2:00 PM - 5:00 PM"
    
                            st.write(f"✅ OE1 and OE5 scheduled together on {elective_day1.strftime('%d-%m-%Y')} at 10:00 AM - 1:00 PM")
                            st.write(f"✅ OE2 scheduled on {elective_day2.strftime('%d-%m-%Y')} at 2:00 PM - 5:00 PM")

                            # Combine non-electives and electives
                            final_df = pd.concat([non_elec_df, df_ele], ignore_index=True)
                        else:
                            final_df = non_elec_df
                            st.write("No electives to schedule.")

                        final_df["Exam Date"] = pd.to_datetime(final_df["Exam Date"], format="%d-%m-%Y", errors='coerce')
                        final_df = final_df.sort_values(["Exam Date", "Semester", "MainBranch"], ascending=True, na_position='last')
                        sem_dict = {s: final_df[final_df["Semester"] == s].copy() for s in sorted(final_df["Semester"].unique())}
                        
                        # Enhanced OE optimization
                        sem_dict = enhanced_optimize_oe_subjects_after_scheduling(sem_dict, holidays_set)
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

                        st.markdown('<div class="status-success">🎉 Timetable generated successfully with enhanced gap optimization!</div>',
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
        <p>🎓 <strong>Enhanced Exam Timetable Generator with Gap Optimization</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
        <p style="font-size: 0.9em;">Advanced scheduling • Gap reduction • Conflict-free timetables • Multiple export formats</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
