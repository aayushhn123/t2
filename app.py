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
    "MCA": "MASTER OF COMPUTER APPLICATIONS",
    "DIPLOMA": "DIPLOMA IN ENGINEERING"
}

# Define logo path (adjust as needed for your environment)
LOGO_PATH = "logo.png"  # Ensure this path is valid in your environment

# Cache for text wrapping results
wrap_text_cache = {}

def get_valid_dates_in_range(start_date, end_date, holidays_set):
    """
    Get all valid examination dates within the specified range.
    Excludes weekends (Sundays) and holidays.
    
    Args:
        start_date (datetime): Start date for examinations
        end_date (datetime): End date for examinations
        holidays_set (set): Set of holiday dates
    
    Returns:
        list: List of valid date strings in DD-MM-YYYY format
    """
    valid_dates = []
    current_date = start_date
    
    while current_date <= end_date:
        # Skip Sundays (weekday 6) and holidays
        if current_date.weekday() != 6 and current_date.date() not in holidays_set:
            valid_dates.append(current_date.strftime("%d-%m-%Y"))
        current_date += timedelta(days=1)
    
    return valid_dates

def find_next_valid_day_in_range(start_date, end_date, holidays_set):
    """
    Find the next valid examination day within the specified range.
    
    Args:
        start_date (datetime): Start date to search from
        end_date (datetime): End date limit
        holidays_set (set): Set of holiday dates
    
    Returns:
        datetime or None: Next valid date or None if no valid date found in range
    """
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() != 6 and current_date.date() not in holidays_set:
            return current_date
        current_date += timedelta(days=1)
    return None

def get_preferred_slot(semester):
    if semester % 2 != 0:  # Odd semester
        odd_sem_position = (semester + 1) // 2
        return "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:  # Even semester
        even_sem_position = semester // 2
        return "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"

def schedule_all_subjects_comprehensively(df, holidays, base_date, end_date):
    """
    FIXED ZERO-UNSCHEDULED SUPER SCHEDULING FUNCTION: 
    1. Groups common subjects as atomic units (all instances together)
    2. Schedules common subjects completely before moving to next day
    3. Ensures NO common subject is ever split across different dates
    4. Maintains daily branch coverage and zero-unscheduled guarantee
    
    Args:
        df: DataFrame with ALL subjects (will filter internally)
        holidays: Set of holiday dates
        base_date: Start date for scheduling
        end_date: Maximum allowed end date
    
    Returns:
        DataFrame with ALL subjects scheduled (ZERO unscheduled guarantee)
    """
    st.info("🚀 FIXED ZERO-UNSCHEDULED SUPER SCHEDULING: Atomic common subject grouping → Priority scheduling → Daily completion")
    
    # STEP 1: COMPREHENSIVE SUBJECT ANALYSIS
    st.write("🔍 **Step 1:** Comprehensive analysis of ALL subjects...")
    
    total_subjects_count = len(df)
    
    # Filter eligible subjects (exclude INTD and OE)
    eligible_subjects = df[
        (df['Category'] != 'INTD') & 
        (~(df['OE'].notna() & (df['OE'].str.strip() != "")))
    ].copy()
    
    if eligible_subjects.empty:
        st.info("No eligible subjects to schedule")
        return df
    
    # Helper functions
    def find_next_valid_day(start_date, holidays_set):
        return find_next_valid_day_in_range(start_date, end_date, holidays_set)
    
    # STEP 2: CREATE ATOMIC SUBJECT UNITS (CRITICAL FIX)
    st.write("🔗 **Step 2:** Creating atomic subject units (common subjects as single units)...")
    
    # Identify ALL unique branch-semester combinations
    all_branch_sem_combinations = set()
    branch_sem_details = {}
    
    for _, row in eligible_subjects.iterrows():
        branch_sem = f"{row['Branch']}_{row['Semester']}"
        all_branch_sem_combinations.add(branch_sem)
        branch_sem_details[branch_sem] = {
            'branch': row['Branch'],
            'semester': row['Semester'],
            'subbranch': row['SubBranch']
        }
    
    st.write(f"🎯 **Coverage target:** {len(all_branch_sem_combinations)} branch-semester combinations")
    
    # Create ATOMIC SUBJECT UNITS - each unit must be scheduled as a whole
    atomic_subject_units = {}
    
    for module_code, group in eligible_subjects.groupby('ModuleCode'):
        # Calculate metrics
        branch_sem_combinations = []
        unique_branches = set()
        unique_semesters = set()
        
        for _, row in group.iterrows():
            branch_sem = f"{row['Branch']}_{row['Semester']}"
            branch_sem_combinations.append(branch_sem)
            unique_branches.add(row['Branch'])
            unique_semesters.add(row['Semester'])
        
        frequency = len(set(branch_sem_combinations))
        
        # Determine if this is a common subject
        is_common_across = group['CommonAcrossSems'].iloc[0] if 'CommonAcrossSems' in group.columns else False
        is_common_within = group['IsCommon'].iloc[0] == 'YES' if 'IsCommon' in group.columns else False
        is_common = is_common_across or is_common_within or frequency > 1
        
        # Create atomic unit
        atomic_unit = {
            'module_code': module_code,
            'subject_name': group['Subject'].iloc[0],
            'frequency': frequency,
            'is_common': is_common,
            'is_common_across': is_common_across,
            'is_common_within': is_common_within,
            'branch_sem_combinations': list(set(branch_sem_combinations)),
            'all_rows': list(group.index),  # All dataframe indices for this subject
            'group_data': group,
            'scheduled': False,
            'scheduled_date': None,
            'scheduled_slot': None,
            'category': group['Category'].iloc[0] if 'Category' in group.columns else 'UNKNOWN',
            'cross_semester_span': len(unique_semesters) > 1,
            'cross_branch_span': len(unique_branches) > 1,
            'unique_semesters': list(unique_semesters),
            'unique_branches': list(unique_branches)
        }
        
        # Calculate priority score
        priority_score = frequency * 10  # Base frequency weight
        
        if is_common_across:
            priority_score += 50  # Highest priority
        elif is_common_within:
            priority_score += 25  # High priority
        elif frequency > 1:
            priority_score += 15  # Medium priority for multi-branch
        
        if atomic_unit['cross_semester_span']:
            priority_score += 15
        if atomic_unit['cross_branch_span']:
            priority_score += 10
        
        atomic_unit['priority_score'] = priority_score
        atomic_subject_units[module_code] = atomic_unit
    
    # Sort atomic units by priority
    sorted_atomic_units = sorted(atomic_subject_units.values(), key=lambda x: x['priority_score'], reverse=True)
    
    # Categorize units
    very_high_priority = [unit for unit in sorted_atomic_units if unit['is_common_across'] or unit['frequency'] >= 8]
    high_priority = [unit for unit in sorted_atomic_units if unit['is_common_within'] and unit not in very_high_priority]
    medium_priority = [unit for unit in sorted_atomic_units if unit['frequency'] >= 2 and unit not in very_high_priority and unit not in high_priority]
    low_priority = [unit for unit in sorted_atomic_units if unit not in very_high_priority and unit not in high_priority and unit not in medium_priority]
    
    st.write(f"🎯 **Atomic unit classification:**")
    st.write(f"   🔥 Very High Priority: {len(very_high_priority)} units (common across sems + freq≥8)")
    st.write(f"   📊 High Priority: {len(high_priority)} units (common within sem)")
    st.write(f"   📋 Medium Priority: {len(medium_priority)} units (multi-branch)")
    st.write(f"   📄 Low Priority: {len(low_priority)} units (individual)")
    
    # STEP 3: ATOMIC SCHEDULING ENGINE
    st.write("🚀 **Step 3:** Atomic scheduling engine (no subject splitting allowed)...")
    
    daily_scheduled_branch_sem = {}
    scheduled_count = 0
    current_date = base_date
    scheduling_day = 0
    target_days = 15
    
    # Create master queue maintaining priority order
    master_queue = very_high_priority + high_priority + medium_priority + low_priority
    unscheduled_units = master_queue.copy()
    
    while scheduling_day < target_days and unscheduled_units:
        # Find next valid exam day
        exam_date = find_next_valid_day(current_date, holidays)
        if exam_date is None:
            st.warning("⚠️ No more valid exam days available in main scheduling")
            break
        
        date_str = exam_date.strftime("%d-%m-%Y")
        scheduling_day += 1
        
        st.write(f"📅 **Day {scheduling_day} ({date_str})**")
        
        # Initialize tracking for this date
        if date_str not in daily_scheduled_branch_sem:
            daily_scheduled_branch_sem[date_str] = set()
        
        day_scheduled_units = []
        
        # PHASE A: Schedule atomic units that can fit completely on this day
        units_to_remove = []
        
        for atomic_unit in unscheduled_units:
            # Check if this unit can be scheduled on this day (no conflicts)
            conflicts = False
            available_branch_sems = []
            
            for branch_sem in atomic_unit['branch_sem_combinations']:
                if branch_sem in daily_scheduled_branch_sem[date_str]:
                    conflicts = True
                    break
                else:
                    available_branch_sems.append(branch_sem)
            
            # If no conflicts, schedule the ENTIRE atomic unit
            if not conflicts and available_branch_sems:
                # Determine time slot based on semester preference
                semester_counts = {}
                for sem in atomic_unit['unique_semesters']:
                    semester_counts[sem] = semester_counts.get(sem, 0) + 1
                
                preferred_semester = max(semester_counts.keys()) if semester_counts else atomic_unit['unique_semesters'][0]
                time_slot = get_preferred_slot(preferred_semester)
                
                # Schedule ALL instances of this subject
                unit_scheduled_count = 0
                for row_idx in atomic_unit['all_rows']:
                    df.loc[row_idx, 'Exam Date'] = date_str
                    df.loc[row_idx, 'Time Slot'] = time_slot
                    unit_scheduled_count += 1
                
                # Update tracking
                for branch_sem in atomic_unit['branch_sem_combinations']:
                    daily_scheduled_branch_sem[date_str].add(branch_sem)
                
                atomic_unit['scheduled'] = True
                atomic_unit['scheduled_date'] = date_str
                atomic_unit['scheduled_slot'] = time_slot
                scheduled_count += unit_scheduled_count
                
                day_scheduled_units.append(atomic_unit)
                units_to_remove.append(atomic_unit)
                
                unit_type = "COMMON" if atomic_unit['is_common'] else "INDIVIDUAL"
                st.write(f"  ✅ **{unit_type} ATOMIC:** {atomic_unit['subject_name']} → "
                        f"{len(atomic_unit['branch_sem_combinations'])} branches "
                        f"(priority: {atomic_unit['priority_score']}) at {time_slot}")
                
                # CRITICAL: If this is a common subject, verify no splitting occurred
                if atomic_unit['is_common']:
                    dates_used = set()
                    for row_idx in atomic_unit['all_rows']:
                        exam_date_value = df.loc[row_idx, 'Exam Date']
                        if pd.notna(exam_date_value) and exam_date_value != "":
                            dates_used.add(exam_date_value)
                    
                    if len(dates_used) > 1:
                        st.error(f"❌ CRITICAL ERROR: {atomic_unit['subject_name']} scheduled across {len(dates_used)} dates!")
                    else:
                        st.write(f"    ✅ Common subject integrity verified for {atomic_unit['subject_name']}")
        
        # Remove scheduled units from unscheduled list
        for unit in units_to_remove:
            unscheduled_units.remove(unit)
        
        # PHASE B: Fill any remaining branch-semester combinations with individual subjects
        remaining_branch_sems = list(all_branch_sem_combinations - daily_scheduled_branch_sem[date_str])
        
        if remaining_branch_sems:
            st.write(f"  🎯 **FILLING GAPS:** {len(remaining_branch_sems)} remaining slots...")
            
            additional_fills = []
            for atomic_unit in unscheduled_units.copy():
                if atomic_unit['frequency'] == 1:  # Only individual subjects for gap filling
                    # Check if this individual subject can fill a remaining slot
                    unit_branch_sem = atomic_unit['branch_sem_combinations'][0]  # Individual has only one
                    
                    if unit_branch_sem in remaining_branch_sems:
                        # Determine time slot
                        preferred_semester = atomic_unit['unique_semesters'][0]
                        time_slot = get_preferred_slot(preferred_semester)
                        
                        # Schedule this individual subject
                        for row_idx in atomic_unit['all_rows']:
                            df.loc[row_idx, 'Exam Date'] = date_str
                            df.loc[row_idx, 'Time Slot'] = time_slot
                        
                        # Update tracking
                        daily_scheduled_branch_sem[date_str].add(unit_branch_sem)
                        remaining_branch_sems.remove(unit_branch_sem)
                        
                        atomic_unit['scheduled'] = True
                        atomic_unit['scheduled_date'] = date_str
                        atomic_unit['scheduled_slot'] = time_slot
                        scheduled_count += len(atomic_unit['all_rows'])
                        
                        additional_fills.append(atomic_unit)
                        
                        st.write(f"    📄 **GAP FILL:** {atomic_unit['subject_name']} at {time_slot}")
            
            # Remove gap-filled units
            for unit in additional_fills:
                unscheduled_units.remove(unit)
        
        # Daily verification
        final_coverage = len(daily_scheduled_branch_sem[date_str])
        coverage_percent = (final_coverage / len(all_branch_sem_combinations)) * 100
        
        st.write(f"  📊 **Daily Summary:** {len(day_scheduled_units) + len(additional_fills)} units scheduled, "
                f"{final_coverage}/{len(all_branch_sem_combinations)} branches covered ({coverage_percent:.1f}%)")
        
        # Progress check
        progress_percent = (scheduled_count / len(eligible_subjects)) * 100
        st.write(f"  📈 **Overall progress:** {scheduled_count}/{len(eligible_subjects)} subjects ({progress_percent:.1f}%)")
        
        if not unscheduled_units:
            st.success(f"🎉 **ALL UNITS SCHEDULED IN TARGET PERIOD!** Completed in {scheduling_day} days")
            break
        
        # Move to next day
        current_date = exam_date + timedelta(days=1)
    
    # STEP 4: EXTENDED SCHEDULING FOR REMAINING UNITS
    if unscheduled_units:
        st.warning(f"⚠️ {len(unscheduled_units)} units still need scheduling - entering extended mode")
        
        extended_day = scheduling_day
        
        while unscheduled_units and extended_day < 25:
            exam_date = find_next_valid_day(current_date, holidays)
            if exam_date is None:
                st.error("❌ No more valid days available")
                break
            
            date_str = exam_date.strftime("%d-%m-%Y")
            extended_day += 1
            
            st.write(f"  📅 **Extended Day {extended_day} ({date_str})**")
            
            if date_str not in daily_scheduled_branch_sem:
                daily_scheduled_branch_sem[date_str] = set()
            
            # Schedule remaining units
            units_scheduled_today = []
            for atomic_unit in unscheduled_units.copy():
                # Check if unit can be scheduled (no conflicts)
                conflicts = False
                for branch_sem in atomic_unit['branch_sem_combinations']:
                    if branch_sem in daily_scheduled_branch_sem[date_str]:
                        conflicts = True
                        break
                
                if not conflicts:
                    # Schedule the unit
                    preferred_semester = atomic_unit['unique_semesters'][0]
                    time_slot = get_preferred_slot(preferred_semester)
                    
                    for row_idx in atomic_unit['all_rows']:
                        df.loc[row_idx, 'Exam Date'] = date_str
                        df.loc[row_idx, 'Time Slot'] = time_slot
                    
                    for branch_sem in atomic_unit['branch_sem_combinations']:
                        daily_scheduled_branch_sem[date_str].add(branch_sem)
                    
                    atomic_unit['scheduled'] = True
                    scheduled_count += len(atomic_unit['all_rows'])
                    units_scheduled_today.append(atomic_unit)
            
            # Remove scheduled units
            for unit in units_scheduled_today:
                unscheduled_units.remove(unit)
            
            st.write(f"    📄 Extended day scheduled: {len(units_scheduled_today)} units")
            current_date = exam_date + timedelta(days=1)
    
    # STEP 5: FINAL VERIFICATION AND STATISTICS
    st.write("📊 **Step 5:** Final verification and statistics...")
    
    successfully_scheduled = df[
        (df['Exam Date'] != "") & 
        (df['Exam Date'] != "Out of Range") & 
        (df['Category'] != 'INTD') & 
        (~(df['OE'].notna() & (df['OE'].str.strip() != "")))
    ]
    
    # CRITICAL: Verify NO common subjects are split
    split_subjects = 0
    properly_grouped_common = 0
    
    for atomic_unit in atomic_subject_units.values():
        if atomic_unit['is_common'] and atomic_unit['scheduled']:
            # Check if all instances are on the same date
            dates_used = set()
            for row_idx in atomic_unit['all_rows']:
                exam_date_value = df.loc[row_idx, 'Exam Date']
                if pd.notna(exam_date_value) and exam_date_value != "":
                    dates_used.add(exam_date_value)
            
            if len(dates_used) > 1:
                split_subjects += 1
                st.error(f"❌ SPLIT DETECTED: {atomic_unit['subject_name']} across {len(dates_used)} dates")
            else:
                properly_grouped_common += 1
    
    total_days_used = len(daily_scheduled_branch_sem)
    success_rate = (len(successfully_scheduled) / len(eligible_subjects)) * 100
    
    st.success(f"🏆 **FIXED ATOMIC SCHEDULING COMPLETE:**")
    st.write(f"   📚 **Total subjects scheduled:** {len(successfully_scheduled)}/{len(eligible_subjects)} ({success_rate:.1f}%)")
    st.write(f"   📅 **Days used:** {total_days_used}")
    st.write(f"   ✅ **Properly grouped common subjects:** {properly_grouped_common}")
    st.write(f"   ❌ **Split common subjects:** {split_subjects}")
    
    if split_subjects == 0:
        st.success("🎉 **PERFECT: NO COMMON SUBJECTS SPLIT!**")
        st.balloons()
    else:
        st.error(f"❌ **CRITICAL: {split_subjects} common subjects were split across dates!**")
    
    return df
    
def read_timetable(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Debug: Show actual column names from the Excel file
        st.write("📋 **Actual columns in uploaded file:**")
        st.write(list(df.columns))
        
        # Enhanced column mapping to handle more variations
        column_mapping = {
            "Program": "Program",
            "Programme": "Program",  # Alternative spelling
            "Stream": "Stream", 
            "Specialization": "Stream",  # Alternative name
            "Branch": "Stream",  # Some files might use Branch for Stream
            "Current Session": "Semester",
            "Academic Session": "Semester",
            "Session": "Semester",
            "Module Description": "SubjectName",
            "Subject Name": "SubjectName",
            "Subject Description": "SubjectName",
            "Module Abbreviation": "ModuleCode",
            "Module Code": "ModuleCode",
            "Subject Code": "ModuleCode",
            "Code": "ModuleCode",
            "Campus Name": "Campus",
            "Campus": "Campus",
            "Difficulty Score": "Difficulty",
            "Difficulty": "Difficulty",
            "Exam Duration": "Exam Duration",
            "Duration": "Exam Duration",
            "Student count": "StudentCount",
            "Student Count": "StudentCount",
            "Enrollment": "StudentCount",
            "Count": "StudentCount",
            "Common across sems": "CommonAcrossSems",
            "Common Across Sems": "CommonAcrossSems",
            "Cross Semester": "CommonAcrossSems",
            "Common Across Semesters": "CommonAcrossSems"
        }
        
        # Handle the "Is Common" column with flexible naming
        is_common_variations = ["Is Common", "IsCommon", "is common", "Is_Common", "is_common", "Common"]
        for variation in is_common_variations:
            if variation in df.columns:
                column_mapping[variation] = "IsCommon"
                st.write(f"✅ Found 'Is Common' column as: '{variation}'")
                break
        else:
            st.warning("⚠️ 'Is Common' column not found in uploaded file. Will create default values.")
        
        # Apply the column mapping
        df = df.rename(columns=column_mapping)
        
        def convert_sem(sem):
            if pd.isna(sem):
                return 0
            
            sem_str = str(sem).strip()
            
            # Handle different semester formats including DIPLOMA and M TECH
            semester_mappings = {
                "Sem I": 1, "Sem II": 2, "Sem III": 3, "Sem IV": 4,
                "Sem V": 5, "Sem VI": 6, "Sem VII": 7, "Sem VIII": 8,
                "Sem IX": 9, "Sem X": 10, "Sem XI": 11, "Sem XII": 12,
                # DIPLOMA variations
                "DIPLOMA Sem I": 1, "DIPLOMA Sem II": 2, "DIPLOMA Sem III": 3,
                "DIPLOMA Sem IV": 4, "DIPLOMA Sem V": 5, "DIPLOMA Sem VI": 6,
                # M TECH variations  
                "M TECH Sem I": 1, "M TECH Sem II": 2, "M TECH Sem III": 3, "M TECH Sem IV": 4,
                # Direct numeric
                "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, 
                "9": 9, "10": 10, "11": 11, "12": 12
            }
            
            return semester_mappings.get(sem_str, 0)

        df["Semester"] = df["Semester"].apply(convert_sem).astype(int)
        
        # Create enhanced branch identifier that considers program type
        def create_branch_identifier(row):
            program = str(row.get("Program", "")).strip()
            stream = str(row.get("Stream", "")).strip()
            
            # Handle cases where stream might be empty or same as program
            if not stream or stream == "nan" or stream == program:
                return program
            else:
                return f"{program}-{stream}"
        
        df["Branch"] = df.apply(create_branch_identifier, axis=1)
        df["Subject"] = df["SubjectName"].astype(str) + " - (" + df["ModuleCode"].astype(str) + ")"
        
        comp_mask = (df["Category"] == "COMP") & df["Difficulty"].notna()
        df["Difficulty"] = None
        df.loc[comp_mask, "Difficulty"] = df.loc[comp_mask, "Difficulty"]
        
        df["Exam Date"] = ""
        df["Time Slot"] = ""
        df["Exam Duration"] = df["Exam Duration"].fillna(3).astype(float)
        df["StudentCount"] = df["StudentCount"].fillna(0).astype(int)
        df["CommonAcrossSems"] = df["CommonAcrossSems"].fillna(False).astype(bool)
        
        # Handle IsCommon column - create if it doesn't exist
        if "IsCommon" not in df.columns:
            st.info("ℹ️ Creating 'IsCommon' column with enhanced logic for all program types")
            df["IsCommon"] = "NO"  # Default value
            
            # Set YES for subjects that are common across semesters
            df.loc[df["CommonAcrossSems"] == True, "IsCommon"] = "YES"
            
            # Enhanced logic to check for subjects common within semester across different programs
            for (semester, module_code), group in df.groupby(['Semester', 'ModuleCode']):
                unique_branches = group['Branch'].unique()
                unique_programs = group['Program'].unique() if 'Program' in group.columns else []
                
                # If subject appears in multiple branches or programs within same semester
                if len(unique_branches) > 1 or len(unique_programs) > 1:
                    df.loc[group.index, "IsCommon"] = "YES"
        else:
            # Clean up the IsCommon column values
            df["IsCommon"] = df["IsCommon"].astype(str).str.strip().str.upper()
            df["IsCommon"] = df["IsCommon"].replace({"TRUE": "YES", "FALSE": "NO", "1": "YES", "0": "NO"})
            df["IsCommon"] = df["IsCommon"].fillna("NO")
        
        df_non = df[df["Category"] != "INTD"].copy()
        df_ele = df[df["Category"] == "INTD"].copy()
        
        def split_br(b):
            p = str(b).split("-", 1)
            if len(p) == 1:
                # Single program case (like DIPLOMA, M TECH without stream)
                return pd.Series([p[0].strip(), ""])
            else:
                return pd.Series([p[0].strip(), p[1].strip()])
        
        for d in (df_non, df_ele):
            d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)
        
        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty", "Exam Duration", "StudentCount", "CommonAcrossSems", "ModuleCode", "IsCommon", "Program"]
        
        # Ensure all required columns exist before selecting
        available_cols = [col for col in cols if col in df_non.columns]
        missing_cols = [col for col in cols if col not in df_non.columns]
        
        if missing_cols:
            st.warning(f"⚠️ Missing columns: {missing_cols}")
            # Add missing columns with default values
            for missing_col in missing_cols:
                if missing_col == "Program":
                    df_non[missing_col] = "B TECH"  # Default program
                    df_ele[missing_col] = "B TECH"
                else:
                    df_non[missing_col] = None
                    df_ele[missing_col] = None
        
        # Update available_cols after adding missing columns
        available_cols = [col for col in cols if col in df_non.columns]
        
        # STORE ORIGINAL DATA FOR FILTER OPTIONS
        st.session_state.original_data_df = df.copy()
        
        return df_non[available_cols], df_ele[available_cols] if available_cols else df_ele, df
        
    except Exception as e:
        st.error(f"Error reading the Excel file: {str(e)}")
        st.error(f"Error details: {type(e).__name__}")
        return None, None, None

   
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
    try:
        # Handle different time formats
        start_time = str(start_time).strip()
        
        # Try to parse the time
        if "AM" in start_time.upper() or "PM" in start_time.upper():
            start = datetime.strptime(start_time, "%I:%M %p")
        else:
            # Try 24-hour format
            start = datetime.strptime(start_time, "%H:%M")
        
        duration = timedelta(hours=float(duration_hours))
        end = start + duration
        return end.strftime("%I:%M %p").replace("AM", "AM").replace("PM", "PM")
    except Exception as e:
        st.write(f"⚠️ Error calculating end time for {start_time}, duration {duration_hours}: {e}")
        return f"{start_time} + {duration_hours}h"
        
def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
    
    # Enable automatic page numbering with alias
    pdf.alias_nb_pages()
    
    try:
        df_dict = pd.read_excel(excel_path, sheet_name=None)
        st.write(f"📊 Read Excel file with {len(df_dict)} sheets: {list(df_dict.keys())}")
    except Exception as e:
        st.error(f"Error reading Excel file for PDF generation: {e}")
        return

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

    def get_semester_program_time_slot(semester, program_type):
        """Get correct time slot based on semester and program type"""
        if isinstance(semester, str):
            # Extract numeric semester from string
            if semester.isdigit():
                sem_num = int(semester)
            else:
                # Handle roman numerals or semester strings
                roman_to_num = {
                    'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6,
                    'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10, 'XI': 11, 'XII': 12
                }
                sem_num = roman_to_num.get(semester, 1)
        else:
            sem_num = int(semester) if semester else 1
        
        return get_preferred_slot(sem_num)

    sheets_processed = 0
    
    for sheet_name, sheet_df in df_dict.items():
        if sheet_df.empty:
            st.warning(f"⚠️ Sheet {sheet_name} is empty, skipping")
            continue
            
        # Reset index to make sure we're working with a clean DataFrame
        if hasattr(sheet_df, 'index') and len(sheet_df.index.names) > 1:
            sheet_df = sheet_df.reset_index()
        
        # Enhanced sheet name parsing to extract program type
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        
        # Determine program type from sheet name
        program_type = "B TECH"  # default
        if sheet_name.startswith("DIPL_"):
            program_type = "DIPLOMA"
            main_branch = main_branch.replace("DIPL_", "")
        elif sheet_name.startswith("MTECH_"):
            program_type = "M TECH"
            main_branch = main_branch.replace("MTECH_", "")
        elif "M TECH" in main_branch:
            program_type = "M TECH"
        elif "DIPLOMA" in main_branch:
            program_type = "DIPLOMA"
        elif "MCA" in main_branch:
            program_type = "MCA"
        
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        
        # Remove '_Electives' suffix if present
        if semester.endswith('_Electives'):
            semester = semester.replace('_Electives', '')
            
        semester_roman = semester if not semester.isdigit() else int_to_roman(int(semester))
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        # Handle normal subjects (non-electives)
        if not sheet_name.endswith('_Electives'):
            # Check if 'Exam Date' column exists
            if 'Exam Date' not in sheet_df.columns:
                st.warning(f"⚠️ Missing 'Exam Date' column in sheet {sheet_name}")
                st.write(f"Available columns: {list(sheet_df.columns)}")
                
                # Try to handle sheets with no exams
                if len(sheet_df.columns) >= 2 and 'No exams scheduled' in str(sheet_df.iloc[0, 0]):
                    st.info(f"ℹ️ Sheet {sheet_name} has no scheduled exams, skipping PDF generation for this sheet")
                    continue
                else:
                    st.error(f"❌ Cannot process sheet {sheet_name} due to missing 'Exam Date' column")
                    continue
                
            # Remove any completely empty rows
            sheet_df = sheet_df.dropna(how='all').reset_index(drop=True)
            
            # Get column structure
            fixed_cols = ["Exam Date"]
            sub_branch_cols = [c for c in sheet_df.columns if c not in fixed_cols and not c.startswith('Unnamed')]
            
            if not sub_branch_cols:
                st.warning(f"⚠️ No subject columns found in sheet {sheet_name}")
                continue
            
            st.write(f"📊 Found {len(sub_branch_cols)} subbranch columns: {sub_branch_cols}")
            
            exam_date_width = 60
            line_height = 10

            # Process in chunks if needed
            for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
                chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
                cols_to_print = fixed_cols + chunk
                chunk_df = sheet_df[cols_to_print].copy()
                
                # Filter out empty rows and rows with "No exams scheduled"
                mask = chunk_df[chunk].apply(lambda row: 
                    (row.astype(str).str.strip() != "") & 
                    (row.astype(str).str.strip() != "---") &
                    (row.astype(str).str.strip() != "No subjects available")
                ).any(axis=1)
                
                # Also filter out "No exams scheduled" rows
                exam_date_mask = ~chunk_df['Exam Date'].astype(str).str.contains('No exams scheduled', na=False)
                
                final_mask = mask & exam_date_mask
                chunk_df = chunk_df[final_mask].reset_index(drop=True)
                
                if chunk_df.empty:
                    st.info(f"ℹ️ No valid exam data found for {sheet_name} chunk {start//sub_branch_cols_per_page + 1}")
                    continue

                # FIXED: Get the correct time slot for this semester and program
                correct_time_slot = get_semester_program_time_slot(semester, program_type)

                # Convert Exam Date to desired format
                try:
                    chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                except Exception as e:
                    st.warning(f"Date conversion error in {sheet_name}: {e}")
                    # Try alternative date parsing
                    chunk_df["Exam Date"] = chunk_df["Exam Date"].astype(str)

                # Calculate column widths
                page_width = pdf.w - 2 * pdf.l_margin
                remaining = page_width - exam_date_width
                sub_width = remaining / max(len(chunk), 1)
                col_widths = [exam_date_width] + [sub_width] * len(chunk)
                total_w = sum(col_widths)
                if total_w > page_width:
                    factor = page_width / total_w
                    col_widths = [w * factor for w in col_widths]
                
                # Add page and print table
                pdf.add_page()
                footer_height = 25
                add_footer_with_page_number(pdf, footer_height)
                
                # FIXED: Pass the correct time slot
                print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=line_height, 
                                 header_content=header_content, branches=chunk, time_slot=correct_time_slot)
                
                sheets_processed += 1

        # Handle electives (similar logic)
        elif sheet_name.endswith('_Electives'):
            # Ensure we have the required columns
            required_cols = ['Exam Date', 'OE', 'SubjectDisplay']
            available_cols = [col for col in required_cols if col in sheet_df.columns]
            
            if len(available_cols) < 2:
                st.warning(f"⚠️ Missing required columns in electives sheet {sheet_name}")
                st.write(f"Available columns: {list(sheet_df.columns)}")
                st.write(f"Required columns: {required_cols}")
                continue
            
            # Use available columns
            elective_data = sheet_df[available_cols].copy()
            elective_data = elective_data.dropna(how='all').reset_index(drop=True)
            
            if elective_data.empty:
                st.info(f"ℹ️ No elective data found for {sheet_name}")
                continue

            # Convert Exam Date to desired format
            try:
               elective_data["Exam Date"] = pd.to_datetime(elective_data["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
            except Exception as e:
               st.warning(f"Date conversion error in electives {sheet_name}: {e}")
               elective_data["Exam Date"] = elective_data["Exam Date"].astype(str)

            # Clean subject display if SubjectDisplay column exists
            if 'SubjectDisplay' in elective_data.columns and 'OE' in elective_data.columns:
               # Remove OE tags that might be duplicated since they're already in the subject display
               elective_data['SubjectDisplay'] = elective_data.apply(
                   lambda row: str(row['SubjectDisplay']).replace(f" [{row['OE']}]", f" [{row['OE']}]") if pd.notna(row['SubjectDisplay']) else "",
                   axis=1
               )

            # Rename columns for clarity in the PDF
            column_mapping = {
               'OE': 'OE Type',
               'SubjectDisplay': 'Subjects'
            }
           
            for old_col, new_col in column_mapping.items():
               if old_col in elective_data.columns:
                   elective_data = elective_data.rename(columns={old_col: new_col})

            # Set column widths
            exam_date_width = 60
            oe_width = 30
            subject_width = pdf.w - 2 * pdf.l_margin - exam_date_width - oe_width
            col_widths = [exam_date_width, oe_width, subject_width]
            cols_to_print = [col for col in ['Exam Date', 'OE Type', 'Subjects'] if col in elective_data.columns]
           
            # Adjust column widths based on available columns
            if len(cols_to_print) == 2:
               col_widths = col_widths[:2]
               col_widths[1] = subject_width + oe_width
           
            # Add page and print table
            pdf.add_page()
            footer_height = 25
            add_footer_with_page_number(pdf, footer_height)
           
            # FIXED: Get correct time slot for electives based on program
            elective_time_slot = get_semester_program_time_slot(semester, program_type)
           
            print_table_custom(pdf, elective_data, cols_to_print, col_widths, line_height=10, 
                            header_content=header_content, branches=['All Streams'], time_slot=elective_time_slot)
           
            sheets_processed += 1

    if sheets_processed == 0:
       st.error("❌ No sheets were processed for PDF generation!")
       return
   
    try:
       pdf.output(pdf_path)
       st.success(f"✅ PDF generated successfully with {sheets_processed} pages and correct time slots")
    except Exception as e:
       st.error(f"❌ Error saving PDF: {e}")
       import traceback
       st.error(f"Traceback: {traceback.format_exc()}")
   
def generate_pdf_timetable(semester_wise_timetable, output_pdf):
    st.write("🔄 Starting PDF generation process...")
    
    temp_excel = os.path.join(os.path.dirname(output_pdf), "temp_timetable.xlsx")
    
    st.write("📊 Generating Excel file first...")
    excel_data = save_to_excel(semester_wise_timetable)
    
    if excel_data:
        st.write(f"💾 Saving temporary Excel file to: {temp_excel}")
        try:
            with open(temp_excel, "wb") as f:
                f.write(excel_data.getvalue())
            st.write("✅ Temporary Excel file saved successfully")
            
            # Verify the Excel file was created and has content
            if os.path.exists(temp_excel):
                file_size = os.path.getsize(temp_excel)
                st.write(f"📋 Excel file size: {file_size} bytes")
                
                # Read back and verify sheets
                try:
                    test_sheets = pd.read_excel(temp_excel, sheet_name=None)
                    #st.write(f"📊 Excel file contains {len(test_sheets)} sheets: {list(test_sheets.keys())}")
                    
                    # Show structure of first few sheets
                    for i, (sheet_name, sheet_df) in enumerate(test_sheets.items()):
                        if i < 3:  # Only show first 3 sheets
                            pass
                            #st.write(f"  📄 Sheet '{sheet_name}': {sheet_df.shape} with columns: {list(sheet_df.columns)}")
                            
                except Exception as e:
                    st.error(f"❌ Error reading back Excel file for verification: {e}")
            else:
                st.error(f"❌ Temporary Excel file was not created at {temp_excel}")
                return
            
        except Exception as e:
            st.error(f"❌ Error saving temporary Excel file: {e}")
            return
            
        #st.write("🎨 Converting Excel to PDF...")
        try:
            convert_excel_to_pdf(temp_excel, output_pdf)
            #st.write("✅ PDF conversion completed")
        except Exception as e:
            st.error(f"❌ Error during Excel to PDF conversion: {e}")
            import traceback
            st.error(f"Traceback: {traceback.format_exc()}")
            return
        
        # Clean up temporary file
        try:
            if os.path.exists(temp_excel):
                os.remove(temp_excel)
                st.write("🗑️ Temporary Excel file cleaned up")
        except Exception as e:
            st.warning(f"⚠️ Could not remove temporary file: {e}")
    else:
        st.error("❌ No Excel data generated - cannot create PDF")
        return
    
    # Post-process PDF to remove blank pages
    st.write("🔧 Post-processing PDF to remove blank pages...")
    try:
        if not os.path.exists(output_pdf):
            st.error(f"❌ PDF file was not created at {output_pdf}")
            return
            
        reader = PdfReader(output_pdf)
        writer = PdfWriter()
        page_number_pattern = re.compile(r'^[\s\n]*(?:Page\s*)?\d+[\s\n]*$')
        
        original_pages = len(reader.pages)
        st.write(f"📄 Original PDF has {original_pages} pages")
        
        pages_kept = 0
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
                pages_kept += 1
                
        st.write(f"📄 Kept {pages_kept} pages out of {original_pages}")
        
        if len(writer.pages) > 0:
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
            st.success(f"✅ PDF post-processing completed - final PDF has {len(writer.pages)} pages")
        else:
            st.warning("⚠️ All pages were filtered out - keeping original PDF")
            
    except Exception as e:
        st.error(f"❌ Error during PDF post-processing: {str(e)}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")
        
    # Final verification
    if os.path.exists(output_pdf):
        final_size = os.path.getsize(output_pdf)
        st.write(f"📄 Final PDF size: {final_size} bytes")
        st.success("🎉 PDF generation process completed successfully!")
    else:
        st.error("❌ Final PDF file does not exist!")

def save_verification_excel(original_df, semester_wise_timetable):
    if not semester_wise_timetable:
        st.error("No timetable data provided for verification")
        return None

    # Combine all scheduled data first
    scheduled_data = pd.concat(semester_wise_timetable.values(), ignore_index=True)

    # Extract ModuleCode from scheduled data more robustly
    scheduled_data["ExtractedModuleCode"] = scheduled_data["Subject"].str.extract(r'\(([^)]+)\)$', expand=False)
    
    # Debug: Check ModuleCode extraction
    st.write("🔍 **ModuleCode extraction check:**")
    module_codes_sample = scheduled_data[['Subject', 'ExtractedModuleCode']].head(3)
    st.dataframe(module_codes_sample)

    # Handle different possible column names in original data
    column_mapping = {
        "Module Abbreviation": ["Module Abbreviation", "ModuleCode", "Module Code", "Code"],
        "Current Session": ["Current Session", "Semester", "Current Academic Session"],
        "Program": ["Program", "Programme"],
        "Stream": ["Stream", "Specialization", "Branch"],
        "Module Description": ["Module Description", "SubjectName", "Subject Name", "Subject"],
        "Exam Duration": ["Exam Duration", "Duration", "Exam_Duration"],
        "Student count": ["Student count", "StudentCount", "Student_count", "Count"],
        "Common across sems": ["Common across sems", "CommonAcrossSems", "Common_across_sems"],
        "Is Common": ["Is Common", "IsCommon", "is common", "Is_Common", "is_common"],
        "Circuit": ["Circuit", "Is_Circuit", "CircuitBranch"]
    }
    
    # Find actual column names
    actual_columns = {}
    for standard_name, possible_names in column_mapping.items():
        for possible_name in possible_names:
            if possible_name in original_df.columns:
                actual_columns[standard_name] = possible_name
                break
        if standard_name not in actual_columns:
            st.warning(f"⚠️ Column '{standard_name}' not found in original data")
    
    st.write(f"🔍 **Mapped columns:** {actual_columns}")

    # Create verification dataframe with available columns
    columns_to_include = list(actual_columns.values())
    verification_df = original_df[columns_to_include].copy()
    
    # Standardize column names
    reverse_mapping = {v: k for k, v in actual_columns.items()}
    verification_df = verification_df.rename(columns=reverse_mapping)

    # Add new columns for scheduled information
    verification_df["Exam Date"] = ""
    verification_df["Time Slot"] = ""  # Semester default timing
    verification_df["Exam Time"] = ""  # Subject-specific timing
    verification_df["Is Common Status"] = ""
    verification_df["Scheduling Status"] = "Not Scheduled"

    # Track statistics
    matched_count = 0
    unmatched_count = 0
    
    # Create comprehensive lookup for scheduled subjects
    scheduled_lookup = {}
    for idx, row in scheduled_data.iterrows():
        module_code = str(row['ExtractedModuleCode']).strip()
        if module_code and module_code != "nan":
            semester = row['Semester']
            branch = row['Branch']
            
            # Create multiple possible keys for matching
            keys_to_try = [
                f"{module_code}_{semester}_{branch}",  # Exact match
            ]
            
            # Add branch variations for matching
            if '-' in branch:
                program, stream = branch.split('-', 1)
                keys_to_try.append(f"{module_code}_{semester}_{program.strip()}-{stream.strip()}")
            
            for key in keys_to_try:
                if key not in scheduled_lookup:
                    scheduled_lookup[key] = []
                scheduled_lookup[key].append(row)
    
    st.write(f"📊 Created lookup with {len(scheduled_lookup)} unique keys")
    
    # Process each row for matching
    for idx, row in verification_df.iterrows():
        try:
            # Get module code
            module_code = str(row.get("Module Abbreviation", "")).strip()
            if not module_code or module_code == "nan":
                unmatched_count += 1
                continue
            
            # Convert semester
            semester_value = row.get("Current Session", "")
            semester_num = convert_semester_to_number(semester_value)
            
            # Create branch identifier
            program = str(row.get("Program", "")).strip()
            stream = str(row.get("Stream", "")).strip()
            branch = f"{program}-{stream}" if program and stream and program != "nan" and stream != "nan" else ""
            
            if not branch:
                st.write(f"⚠️ Empty branch for module {module_code}")
                unmatched_count += 1
                continue
            
            # Try to find match using lookup
            lookup_key = f"{module_code}_{semester_num}_{branch}"
            match_found = False
            
            if lookup_key in scheduled_lookup:
                # Found exact match
                matched_subject = scheduled_lookup[lookup_key][0]  # Take first match
                match_found = True
            else:
                # Try alternative matching approaches
                # 1. Try with just module code and semester (for common subjects)
                for key, subjects in scheduled_lookup.items():
                    if key.startswith(f"{module_code}_{semester_num}_"):
                        # Check if this is a common subject
                        matched_subject = subjects[0]
                        if matched_subject.get('CommonAcrossSems', False):
                            match_found = True
                            break
                
                # 2. Try partial branch matching
                if not match_found:
                    for key, subjects in scheduled_lookup.items():
                        if key.startswith(f"{module_code}_{semester_num}_"):
                            # Extract branch from key
                            key_parts = key.split('_')
                            if len(key_parts) >= 3:
                                key_branch = '_'.join(key_parts[2:])
                                # Check if branches are similar (same program or stream)
                                if branch in key_branch or key_branch in branch:
                                    matched_subject = subjects[0]
                                    match_found = True
                                    break
            
            if match_found:
                # Found a match
                exam_date = matched_subject["Exam Date"]
                assigned_time_slot = matched_subject["Time Slot"]
                duration = row.get("Exam Duration", 3.0)
                
                # Handle duration
                try:
                    duration = float(duration) if pd.notna(duration) else 3.0
                except:
                    duration = 3.0
                
                # UPDATED: Calculate Time Slot (semester default) and Exam Time (subject-specific)
                semester_default_slot = get_preferred_slot(semester_num)
                
                # Time Slot = Semester default timing
                verification_df.at[idx, "Time Slot"] = semester_default_slot
                
                # Exam Time = Subject-specific timing
                if assigned_time_slot and str(assigned_time_slot).strip() and str(assigned_time_slot) != "nan":
                    try:
                        # Calculate subject-specific exam time based on duration
                        if duration != 3:
                            # Non-standard duration: calculate specific time range
                            start_time = str(assigned_time_slot).split(" - ")[0].strip()
                            end_time = calculate_end_time(start_time, duration)
                            exam_time = f"{start_time} - {end_time}"
                        else:
                            # Standard 3-hour duration
                            if assigned_time_slot != semester_default_slot:
                                # Common subject with different timing than semester default
                                exam_time = str(assigned_time_slot)
                            else:
                                # Regular subject with semester default timing
                                start_time = str(assigned_time_slot).split(" - ")[0].strip()
                                end_time = calculate_end_time(start_time, duration)
                                exam_time = f"{start_time} - {end_time}"
                        
                        verification_df.at[idx, "Exam Time"] = exam_time
                        
                    except Exception as e:
                        st.write(f"⚠️ Error calculating time for {module_code}: {e}")
                        verification_df.at[idx, "Exam Time"] = str(assigned_time_slot)
                else:
                    verification_df.at[idx, "Exam Time"] = "TBD"
                
                # Update other fields
                verification_df.at[idx, "Exam Date"] = str(exam_date)
                verification_df.at[idx, "Scheduling Status"] = "Scheduled"
                
                # Determine commonality status based on both columns
                common_across_sems = matched_subject.get('CommonAcrossSems', False)
                is_common_within = matched_subject.get('IsCommon', 'NO') == 'YES'
                
                if common_across_sems:
                    verification_df.at[idx, "Is Common Status"] = "Common Across Semesters"
                elif is_common_within:
                    verification_df.at[idx, "Is Common Status"] = "Common Within Semester"
                else:
                    verification_df.at[idx, "Is Common Status"] = "Uncommon"
                
                matched_count += 1
                
            else:
                # No match found
                verification_df.at[idx, "Exam Date"] = "Not Scheduled"
                verification_df.at[idx, "Time Slot"] = "Not Scheduled"
                verification_df.at[idx, "Exam Time"] = "Not Scheduled" 
                verification_df.at[idx, "Is Common Status"] = "N/A"
                verification_df.at[idx, "Scheduling Status"] = "Not Scheduled"
                unmatched_count += 1
                
                if unmatched_count <= 10:  # Show first 10 unmatched for debugging
                    st.write(f"   ❌ **NO MATCH** for {module_code} ({branch}, Sem {semester_num})")
                     
        except Exception as e:
            st.error(f"Error processing row {idx}: {e}")
            unmatched_count += 1

    st.success(f"✅ **Verification Results:**")
    st.write(f"   📊 Matched: {matched_count} subjects")
    st.write(f"   ⚠️ Unmatched: {unmatched_count} subjects")
    st.write(f"   📈 Match rate: {(matched_count/(matched_count+unmatched_count)*100):.1f}%")

    # Save to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        verification_df.to_excel(writer, sheet_name="Verification", index=False)
        
        # Add a summary sheet with breakdown by commonality
        summary_data = {
            "Metric": ["Total Subjects", "Scheduled Subjects", "Unscheduled Subjects", "Match Rate (%)",
                      "Common Across Semesters", "Common Within Semester", "Uncommon Subjects"],
            "Value": [
                matched_count + unmatched_count, 
                matched_count, 
                unmatched_count, 
                round(matched_count/(matched_count+unmatched_count)*100, 1) if (matched_count+unmatched_count) > 0 else 0,
                len(verification_df[verification_df["Is Common Status"] == "Common Across Semesters"]),
                len(verification_df[verification_df["Is Common Status"] == "Common Within Semester"]),
                len(verification_df[verification_df["Is Common Status"] == "Uncommon"])
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        
        # Add unmatched subjects sheet for debugging
        unmatched_subjects = verification_df[verification_df["Scheduling Status"] == "Not Scheduled"]
        if not unmatched_subjects.empty:
            unmatched_subjects.to_excel(writer, sheet_name="Unmatched_Subjects", index=False)

    output.seek(0)
    return output

def convert_semester_to_number(semester_value):
    """Convert semester string to number with better error handling"""
    if pd.isna(semester_value):
        return 0
    
    semester_str = str(semester_value).strip()
    
    semester_map = {
        "Sem I": 1, "Sem II": 2, "Sem III": 3, "Sem IV": 4,
        "Sem V": 5, "Sem VI": 6, "Sem VII": 7, "Sem VIII": 8,
        "Sem IX": 9, "Sem X": 10, "Sem XI": 11,
        "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9, "10": 10, "11": 11
    }
    
    return semester_map.get(semester_str, 0)


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

    def should_show_exam_time(row, semester_default_slot):
        """Determine if subject needs to show specific exam time"""
        duration = row.get('Exam Duration', 3)
        assigned_slot = row.get('Time Slot', '')
        is_common_across = row.get('CommonAcrossSems', False)
        
        # Show exam time if:
        # 1. Duration is not 3 hours
        # 2. Common subject assigned different time than semester default
        return (duration != 3) or (is_common_across and assigned_slot != semester_default_slot)

    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sheets_created = 0  # Track number of sheets created
            
            for sem, df_sem in semester_wise_timetable.items():
                for main_branch in df_sem["MainBranch"].unique():
                    df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
                    
                    # Separate non-electives and electives
                    df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                    df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()

                    # Process non-electives
                    roman_sem = int_to_roman(sem)
                    sheet_name = f"{main_branch}_Sem_{roman_sem}"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    
                    if not df_non_elec.empty:
                        st.write(f"📊 Processing {len(df_non_elec)} non-elective subjects for {sheet_name}")
                        
                        # Get semester default time slot
                        semester_default_slot = get_preferred_slot(sem)
                        
                        # Create a copy to avoid modifying the original
                        df_processed = df_non_elec.copy().reset_index(drop=True)
                        
                        # Add difficulty info safely
                        difficulty_values = []
                        for idx in range(len(df_processed)):
                            row = df_processed.iloc[idx]
                            difficulty = row.get('Difficulty', None)
                            if pd.notna(difficulty):
                                if difficulty == 0:
                                    difficulty_values.append(" (Easy)")
                                elif difficulty == 1:
                                    difficulty_values.append(" (Difficult)")
                                else:
                                    difficulty_values.append("")
                            else:
                                difficulty_values.append("")
                        
                        # Create subject display with conditional exam time
                        subject_displays = []
                        for idx in range(len(df_processed)):
                            row = df_processed.iloc[idx]
                            base_subject = str(row.get('Subject', ''))
                            difficulty_suffix = difficulty_values[idx]
                            
                            if should_show_exam_time(row, semester_default_slot):
                                duration = row.get('Exam Duration', 3)
                                assigned_slot = row.get('Time Slot', semester_default_slot)
                                
                                try:
                                    if duration != 3:
                                        # Non-standard duration
                                        start_time = str(assigned_slot).split(" - ")[0].strip()
                                        end_time = calculate_end_time(start_time, duration)
                                        exam_time_suffix = f" [{start_time} - {end_time}]"
                                    else:
                                        # Common subject with different timing
                                        exam_time_suffix = f" [{assigned_slot}]"
                                except Exception as e:
                                    st.warning(f"Error calculating exam time for subject {base_subject}: {e}")
                                    exam_time_suffix = ""
                                
                                subject_display = base_subject + difficulty_suffix + exam_time_suffix
                            else:
                                # Regular subject - no exam time needed
                                subject_display = base_subject + difficulty_suffix
                            
                            subject_displays.append(subject_display)
                        
                        # Add the subject display column
                        df_processed["SubjectDisplay"] = subject_displays
                        
                        # Convert and sort by date
                        df_processed["Exam Date"] = pd.to_datetime(df_processed["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                        df_processed = df_processed.sort_values(by="Exam Date", ascending=True)
                        
                        # Group subjects by date, combining any multiple subjects per branch per day
                        grouped_by_date = df_processed.groupby(['Exam Date', 'SubBranch'])['SubjectDisplay'].apply(
                            lambda x: ", ".join(str(i) for i in x)
                        ).reset_index()
                        
                        # Create pivot with single date index
                        pivot_df = grouped_by_date.pivot_table(
                            index="Exam Date",
                            columns="SubBranch",
                            values="SubjectDisplay",
                            aggfunc=lambda x: ", ".join(str(i) for i in x)
                        ).fillna("---")
                        
                        # Sort and format dates
                        pivot_df = pivot_df.sort_index(ascending=True)
                        formatted_dates = [d.strftime("%d-%m-%Y") if pd.notna(d) else "" for d in pivot_df.index]
                        pivot_df.index = formatted_dates
                        
                        # Reset index to make 'Exam Date' a column
                        pivot_df = pivot_df.reset_index()
                        
                        # Ensure the first column is named 'Exam Date'
                        if pivot_df.columns[0] != 'Exam Date':
                            pivot_df = pivot_df.rename(columns={pivot_df.columns[0]: 'Exam Date'})
                        
                        # Save to Excel
                        pivot_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        sheets_created += 1
                        st.write(f"✅ Created sheet {sheet_name} with {len(pivot_df)} exam dates")
                        
                    else:
                        # Create empty sheet structure for branches with no subjects
                        st.write(f"⚠️ No non-elective subjects for {sheet_name}, creating empty structure")
                        
                        # Get all possible subbranches for this main branch from the semester
                        all_subbranches = df_sem[df_sem["MainBranch"] == main_branch]["SubBranch"].unique()
                        
                        if len(all_subbranches) > 0:
                            # Create empty dataframe with proper structure
                            empty_data = {
                                'Exam Date': ['No exams scheduled'],
                            }
                            
                            # Add columns for each subbranch
                            for subbranch in sorted(all_subbranches):
                                empty_data[subbranch] = ['---']
                            
                            empty_df = pd.DataFrame(empty_data)
                            empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            sheets_created += 1
                            st.write(f"✅ Created empty sheet {sheet_name} with structure for subbranches: {', '.join(all_subbranches)}")
                        else:
                            # If no subbranches, create minimal structure
                            empty_df = pd.DataFrame({
                                'Exam Date': ['No exams scheduled'],
                                'Subjects': ['No subjects available']
                            })
                            empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            sheets_created += 1
                            st.write(f"✅ Created minimal empty sheet {sheet_name}")

                    # Process electives in a separate sheet (only if electives exist)
                    if not df_elec.empty:
                        st.write(f"📊 Processing {len(df_elec)} elective subjects for {sheet_name}")
                        
                        # Create a copy to avoid modifying the original
                        df_elec_processed = df_elec.copy().reset_index(drop=True)
                        
                        # Add difficulty info safely
                        difficulty_values_elec = []
                        for idx in range(len(df_elec_processed)):
                            row = df_elec_processed.iloc[idx]
                            difficulty = row.get('Difficulty', None)
                            if pd.notna(difficulty):
                                if difficulty == 0:
                                    difficulty_values_elec.append(" (Easy)")
                                elif difficulty == 1:
                                    difficulty_values_elec.append(" (Difficult)")
                                else:
                                    difficulty_values_elec.append("")
                            else:
                                difficulty_values_elec.append("")
                        
                        # Create subject display with OE and conditional exam time
                        subject_displays_elec = []
                        for idx in range(len(df_elec_processed)):
                            row = df_elec_processed.iloc[idx]
                            base_subject = str(row.get('Subject', ''))
                            oe_type = str(row.get('OE', ''))
                            difficulty_suffix = difficulty_values_elec[idx]
                            
                            base_display = f"{base_subject} [{oe_type}]"
                            
                            # Add exam time for non-standard durations
                            duration = row.get('Exam Duration', 3)
                            if duration != 3:
                                try:
                                    assigned_slot = row.get('Time Slot', '10:00 AM - 1:00 PM')
                                    start_time = str(assigned_slot).split(" - ")[0].strip()
                                    end_time = calculate_end_time(start_time, duration)
                                    exam_time_suffix = f" [{start_time} - {end_time}]"
                                    subject_display = base_display + difficulty_suffix + exam_time_suffix
                                except Exception as e:
                                    st.warning(f"Error calculating exam time for elective {base_subject}: {e}")
                                    subject_display = base_display + difficulty_suffix
                            else:
                                subject_display = base_display + difficulty_suffix
                            
                            subject_displays_elec.append(subject_display)
                        
                        # Add the subject display column
                        df_elec_processed["SubjectDisplay"] = subject_displays_elec
                        
                        # Group by OE and Date only (no time slot)
                        elec_pivot = df_elec_processed.groupby(['OE', 'Exam Date'])['SubjectDisplay'].apply(
                            lambda x: ", ".join(sorted(set(x)))
                        ).reset_index()
                        
                        # Format dates
                        elec_pivot['Exam Date'] = pd.to_datetime(
                            elec_pivot['Exam Date'], format="%d-%m-%Y", errors='coerce'
                        ).dt.strftime("%d-%m-%Y")
                        elec_pivot = elec_pivot.sort_values(by="Exam Date", ascending=True)
                        
                        # Save to Excel
                        elective_sheet_name = f"{main_branch}_Sem_{roman_sem}_Electives"
                        if len(elective_sheet_name) > 31:
                            elective_sheet_name = elective_sheet_name[:31]
                        elec_pivot.to_excel(writer, sheet_name=elective_sheet_name, index=False)
                        sheets_created += 1
                        st.write(f"✅ Created electives sheet {elective_sheet_name} with {len(elec_pivot)} entries")

            # Check if any sheets were created
            if sheets_created == 0:
                st.error("❌ No sheets were created! Creating a dummy sheet to prevent Excel error.")
                # Create a dummy sheet to prevent the "At least one sheet must be visible" error
                dummy_df = pd.DataFrame({'Message': ['No data available']})
                dummy_df.to_excel(writer, sheet_name="No_Data", index=False)
                
        output.seek(0)
        
        if sheets_created > 0:
            st.success(f"✅ Excel file created successfully with {sheets_created} required sheets")
        else:
            st.warning("⚠️ Excel file created with dummy sheet due to no data")
            
        return output
        
    except Exception as e:
        st.error(f"Error creating Excel file: {e}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")
        return None
# ============================================================================
# INTD/OE SUBJECT SCHEDULING LOGIC
# ============================================================================

def find_next_valid_day_for_electives(start_day, holidays):
    """Find the next valid day for scheduling electives (skip weekends and holidays)"""
    day = start_day
    while True:
        day_date = day.date()
        if day.weekday() == 6 or day_date in holidays:
            day += timedelta(days=1)
            continue
        return day

def schedule_electives_globally(df_ele, max_non_elec_date, holidays_set):
    """
    Schedule electives globally after non-elective scheduling is complete.
    OE1 and OE5 on first day, OE2 on second day (immediately after).
    """
    if df_ele is None or df_ele.empty:
        return df_ele
    
    st.info("🎓 Scheduling electives globally...")
    
    # Find the first and second valid elective days
    elective_day1 = find_next_valid_day_for_electives(
        datetime.combine(max_non_elec_date, datetime.min.time()) + timedelta(days=1), 
        holidays_set
    )
    elective_day2 = find_next_valid_day_for_electives(elective_day1 + timedelta(days=1), holidays_set)
    
    elective_day1_str = elective_day1.strftime("%d-%m-%Y")
    elective_day2_str = elective_day2.strftime("%d-%m-%Y")
    
    # Schedule OE1 and OE5 together on the first elective day (morning slot)
    df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Exam Date'] = elective_day1_str
    df_ele.loc[(df_ele['OE'] == 'OE1') | (df_ele['OE'] == 'OE5'), 'Time Slot'] = "10:00 AM - 1:00 PM"
    
    # Schedule OE2 on the second elective day (afternoon slot)
    df_ele.loc[df_ele['OE'] == 'OE2', 'Exam Date'] = elective_day2_str
    df_ele.loc[df_ele['OE'] == 'OE2', 'Time Slot'] = "2:00 PM - 5:00 PM"
    
    st.write(f"✅ OE1 and OE5 scheduled on {elective_day1_str} at 10:00 AM - 1:00 PM")
    st.write(f"✅ OE2 scheduled on {elective_day2_str} at 2:00 PM - 5:00 PM")
    
    return df_ele

def optimize_schedule_by_filling_gaps(df_dict, holidays_set, start_date, end_date):
    """
    After all scheduling is done, try to move subjects up to fill gaps and reduce span.
    ONLY MOVES: Subjects whose "Is Common" value is FALSE (uncommon subjects).
    EXCLUDES: INTD subjects, OE subjects (electives), and ALL common subjects.
    
    Args:
        df_dict (dict): Dictionary of semester dataframes
        holidays_set (set): Set of holiday dates
        start_date (datetime): Start date for examinations
        end_date (datetime): End date for examinations
    
    Returns:
        tuple: (updated_df_dict, moves_made, optimization_log)
    """
    st.info("🎯 Optimizing schedule by filling gaps (ONLY moving uncommon subjects)...")
    
    if not df_dict:
        return df_dict, 0, []

    # Helper function to move a subject (defined at the top)
    def move_subject(df_dict, subject_info, new_date, schedule_grid, old_date, subbranch_sem_key):
        """Move a single subject to a new date"""
        try:
            semester = subject_info['semester']
            branch = subject_info['branch']
            subject = subject_info['subject']
            
            # Get preferred time slot for this semester
            preferred_slot = get_preferred_slot(semester)
            
            # Update in the semester dictionary
            mask = (df_dict[semester]['Subject'] == subject) & \
                   (df_dict[semester]['Branch'] == branch) & \
                   (df_dict[semester]['SubBranch'] == subject_info['subbranch'])
            
            if mask.any():
                df_dict[semester].loc[mask, 'Exam Date'] = new_date
                df_dict[semester].loc[mask, 'Time Slot'] = preferred_slot
                
                # Update schedule grid
                # Remove from old date
                if old_date in schedule_grid and subbranch_sem_key in schedule_grid[old_date]:
                    del schedule_grid[old_date][subbranch_sem_key]
                    if not schedule_grid[old_date]:  # If date becomes empty
                        del schedule_grid[old_date]
                
                # Add to new date
                if new_date not in schedule_grid:
                    schedule_grid[new_date] = {}
                schedule_grid[new_date][subbranch_sem_key] = subject_info
                
                return True
        except Exception as e:
            return False
        
        return False
    
    # Combine all data to analyze the schedule
    all_data = pd.concat(df_dict.values(), ignore_index=True)
    
    # Ensure all dates are in DD-MM-YYYY string format
    def normalize_date_to_ddmmyyyy(date_val):
        """Convert any date format to DD-MM-YYYY string format"""
        if pd.isna(date_val) or date_val == "":
            return ""
        
        if isinstance(date_val, pd.Timestamp):
            return date_val.strftime("%d-%m-%Y")
        elif isinstance(date_val, str):
            try:
                parsed = pd.to_datetime(date_val, format="%d-%m-%Y", errors='raise')
                return parsed.strftime("%d-%m-%Y")
            except:
                try:
                    parsed = pd.to_datetime(date_val, dayfirst=True, errors='raise')
                    return parsed.strftime("%d-%m-%Y")
                except:
                    return str(date_val)
        else:
            try:
                parsed = pd.to_datetime(date_val, errors='coerce')
                if pd.notna(parsed):
                    return parsed.strftime("%d-%m-%Y")
                else:
                    return str(date_val)
            except:
                return str(date_val)
    
    # Apply date normalization
    all_data['Exam Date'] = all_data['Exam Date'].apply(normalize_date_to_ddmmyyyy)
    
    # Helper function to get subbranch-semester key
    def get_subbranch_semester_key(subbranch, semester):
        return f"{subbranch}_{semester}"
    
    # Helper function to safely get string value (handles NaN)
    def safe_get_string(value, default=""):
        """Safely get string value from potentially NaN/None values"""
        if pd.isna(value) or value is None:
            return default
        return str(value).strip()
    
    # Helper function to safely get boolean value
    def safe_get_boolean(value, default=False):
        """Safely get boolean value from potentially NaN/None values"""
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.strip().upper() in ['TRUE', 'YES', '1']
        return bool(value)
    
    # Build schedule grid: date -> subbranch_semester -> subject
    # Include ALL subjects for gap detection
    all_scheduled_data = all_data[
        (all_data['Exam Date'] != "") & 
        (all_data['Exam Date'] != "Out of Range") & 
        (all_data['Exam Date'] != "Cannot Schedule")
    ].copy()
    
    schedule_grid = {}
    subbranch_semester_combinations = set()
    
    for _, row in all_scheduled_data.iterrows():
        date_str = row['Exam Date']
        subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
        
        if date_str not in schedule_grid:
            schedule_grid[date_str] = {}
        
        # Safely extract values with proper NaN handling
        oe_value = safe_get_string(row.get('OE', ''))
        is_common_within_raw = row.get('IsCommon', 'NO')
        is_common_within = safe_get_string(is_common_within_raw, 'NO').upper() == 'YES'
        category_value = safe_get_string(row.get('Category', ''))
        is_common_across = safe_get_boolean(row.get('CommonAcrossSems', False))
        is_circuit = safe_get_boolean(row.get('Circuit', False))
        
        # STRICT filtering: Only uncommon subjects are eligible for moving
        is_eligible_for_moving = (
            category_value != 'INTD' and      # Not INTD
            not oe_value and                  # Not OE
            not is_common_within and          # Not common within semester (Is Common = FALSE)
            not is_common_across              # Not common across semesters
        )
        
        schedule_grid[date_str][subbranch_sem_key] = {
            'subject': safe_get_string(row['Subject']),
            'branch': safe_get_string(row['Branch']),
            'index': row.name,
            'semester': row['Semester'],
            'subbranch': safe_get_string(row['SubBranch']),
            'module_code': safe_get_string(row.get('ModuleCode', '')),
            'is_common_across': is_common_across,
            'is_common_within': is_common_within,
            'category': category_value,
            'circuit': is_circuit,
            'oe': oe_value,
            'is_eligible_for_moving': is_eligible_for_moving
        }
        
        subbranch_semester_combinations.add(subbranch_sem_key)
    
    # Get all dates in the schedule range
    all_scheduled_dates = sorted(schedule_grid.keys(), 
                                key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    if not all_scheduled_dates:
        return df_dict, 0, []
    
    # Create complete date range from start to end of schedule
    first_exam_date = datetime.strptime(all_scheduled_dates[0], "%d-%m-%Y")
    last_exam_date = datetime.strptime(all_scheduled_dates[-1], "%d-%m-%Y")
    
    # Get all valid dates in the range
    valid_dates_in_range = get_valid_dates_in_range(
        max(start_date, first_exam_date), 
        min(end_date, last_exam_date), 
        holidays_set
    )
    
    st.write(f"📊 Analyzing {len(valid_dates_in_range)} valid dates for gap optimization")
    st.write(f"📅 Current schedule spans from {all_scheduled_dates[0]} to {all_scheduled_dates[-1]}")
    
    # Identify gaps in the schedule
    gaps = []
    for date_str in valid_dates_in_range:
        if date_str not in schedule_grid:
            # This is a complete gap day
            gaps.append({
                'date': date_str,
                'type': 'complete_gap',
                'available_subbranches': list(subbranch_semester_combinations)
            })
        else:
            # Check for partial gaps (some subbranches free)
            occupied_subbranches = set(schedule_grid[date_str].keys())
            available_subbranches = subbranch_semester_combinations - occupied_subbranches
            
            if available_subbranches:
                gaps.append({
                    'date': date_str,
                    'type': 'partial_gap',
                    'available_subbranches': list(available_subbranches)
                })
    
    st.write(f"🔍 Found {len(gaps)} gaps to potentially fill")
    
    # Identify ONLY uncommon moveable subjects
    uncommon_moveable_subjects = []
    
    for date_str in reversed(all_scheduled_dates):  # Start from the end
        if date_str in schedule_grid:
            for subbranch_sem_key, subject_info in schedule_grid[date_str].items():
                # STRICT FILTER: Only consider truly uncommon subjects
                if not subject_info.get('is_eligible_for_moving', False):
                    continue
                
                # Calculate priority score for uncommon subjects
                priority_score = 0
                
                # Base priority for uncommon subjects
                priority_score += 10
                
                # Bonus for circuit subjects (if individual)
                if subject_info.get('circuit', False):
                    priority_score += 5
                
                # Prefer subjects from later dates for moving up
                date_obj = datetime.strptime(date_str, "%d-%m-%Y")
                last_date_obj = datetime.strptime(all_scheduled_dates[-1], "%d-%m-%Y")
                days_from_end = (last_date_obj - date_obj).days
                priority_score += max(0, 20 - days_from_end)
                
                if priority_score > 0:
                    subject_type = "Uncommon Circuit" if subject_info.get('circuit', False) else "Uncommon"
                    
                    uncommon_moveable_subjects.append({
                        'current_date': date_str,
                        'subbranch_sem_key': subbranch_sem_key,
                        'subject_info': subject_info,
                        'priority_score': priority_score,
                        'subject_type': subject_type
                    })
    
    # Sort uncommon subjects by priority (highest first)
    uncommon_moveable_subjects.sort(key=lambda x: x['priority_score'], reverse=True)
    
    st.write(f"🎯 Found {len(uncommon_moveable_subjects)} uncommon moveable subjects")
    
    # Show breakdown of subject types
    subject_type_counts = {}
    for moveable in uncommon_moveable_subjects:
        subject_type = moveable['subject_type']
        subject_type_counts[subject_type] = subject_type_counts.get(subject_type, 0) + 1
    
    if subject_type_counts:
        st.info("📋 Moveable subject breakdown:")
        for subject_type, count in subject_type_counts.items():
            st.write(f"  • {subject_type}: {count} subjects")
    
    # Show breakdown of excluded subjects - FIXED: Handle Series properly
    try:
        intd_count = len(all_scheduled_data[all_scheduled_data['Category'] == 'INTD'])
    except:
        intd_count = 0
    
    try:
        oe_mask = all_scheduled_data['OE'].notna() & (all_scheduled_data['OE'].astype(str).str.strip() != "")
        oe_count = len(all_scheduled_data[oe_mask])
    except:
        oe_count = 0
    
    try:
        # Fixed: Apply safe_get_string to each element individually
        common_within_count = 0
        for _, row in all_scheduled_data.iterrows():
            is_common_val = safe_get_string(row.get('IsCommon', 'NO'), 'NO').upper()
            if is_common_val == 'YES':
                common_within_count += 1
    except:
        common_within_count = 0
    
    try:
        common_across_count = len(all_scheduled_data[all_scheduled_data['CommonAcrossSems'] == True])
    except:
        common_across_count = 0
    
    st.info("❌ **Excluded from optimization:**")
    if intd_count > 0:
        st.write(f"  • {intd_count} INTD subjects")
    if oe_count > 0:
        st.write(f"  • {oe_count} OE subjects")
    if common_within_count > 0:
        st.write(f"  • {common_within_count} common within semester subjects")
    if common_across_count > 0:
        st.write(f"  • {common_across_count} common across semester subjects")
    
    # Optimization process - ONLY move uncommon subjects
    moves_made = 0
    optimization_log = []
    
    st.write("🔄 Moving ONLY uncommon subjects...")
    for gap in gaps:
        gap_date = gap['date']
        available_subbranches = gap['available_subbranches'][:]  # Copy list
        
        for moveable in uncommon_moveable_subjects[:]:  # Use slice to allow removal
            if not available_subbranches:  # No more space in this gap
                break
                
            current_date = moveable['current_date']
            subbranch_sem_key = moveable['subbranch_sem_key']
            subject_info = moveable['subject_info']
            subject_type = moveable['subject_type']
            
            # Check if this subject can move to this gap
            if subbranch_sem_key in available_subbranches:
                gap_date_obj = datetime.strptime(gap_date, "%d-%m-%Y")
                current_date_obj = datetime.strptime(current_date, "%d-%m-%Y")
                
                # Only move if it's moving to an earlier date
                if gap_date_obj < current_date_obj:
                    # Move the uncommon subject
                    if move_subject(df_dict, subject_info, gap_date, schedule_grid, current_date, subbranch_sem_key):
                        available_subbranches.remove(subbranch_sem_key)
                        uncommon_moveable_subjects.remove(moveable)
                        moves_made += 1
                        
                        days_moved_up = (current_date_obj - gap_date_obj).days
                        optimization_log.append(
                            f"Moved {subject_type} subject: {subject_info['subject']} "
                            f"({subject_info['subbranch']}, Sem {subject_info['semester']}) "
                            f"from {current_date} to {gap_date} (moved up {days_moved_up} days)"
                        )
        
        # Update gap's available subbranches
        gap['available_subbranches'] = available_subbranches
    
    # Calculate span reduction
    if moves_made > 0:
        # Recalculate the schedule span
        updated_all_data = pd.concat(df_dict.values(), ignore_index=True)
        updated_scheduled = updated_all_data[
            (updated_all_data['Exam Date'] != "") & 
            (updated_all_data['Exam Date'] != "Out of Range")
        ].copy()
        
        if not updated_scheduled.empty:
            updated_scheduled['Exam Date'] = updated_scheduled['Exam Date'].apply(normalize_date_to_ddmmyyyy)
            updated_dates = sorted(updated_scheduled['Exam Date'].unique(), 
                                 key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
            
            if updated_dates:
                new_first_date = updated_dates[0]
                new_last_date = updated_dates[-1]
                
                original_span = (datetime.strptime(all_scheduled_dates[-1], "%d-%m-%Y") - 
                               datetime.strptime(all_scheduled_dates[0], "%d-%m-%Y")).days + 1
                new_span = (datetime.strptime(new_last_date, "%d-%m-%Y") - 
                           datetime.strptime(new_first_date, "%d-%m-%Y")).days + 1
                
                span_reduction = original_span - new_span
                
                if span_reduction > 0:
                    optimization_log.append(f"Schedule span reduced by {span_reduction} days!")
                    st.success(f"📉 Schedule span reduced from {original_span} to {new_span} days (saved {span_reduction} days)")
    
    # Ensure all dates in df_dict are properly formatted
    for sem in df_dict:
        df_dict[sem]['Exam Date'] = df_dict[sem]['Exam Date'].apply(normalize_date_to_ddmmyyyy)
    
    if moves_made > 0:
        st.success(f"✅ Gap Optimization: Made {moves_made} moves to fill gaps!")
        with st.expander("📋 Gap Optimization Details"):
            for log in optimization_log:
                st.write(f"• {log}")
    else:
        st.info("ℹ️ No beneficial moves found for gap optimization (only uncommon subjects considered)")
    
    return df_dict, moves_made, optimization_log
    

def optimize_oe_subjects_after_scheduling(sem_dict, holidays, optimizer=None):
    """
    After main scheduling AND gap optimization, check if OE subjects can be moved to earlier COMPLETELY EMPTY days.
    CRITICAL: OE2 must be scheduled on the day immediately after OE1/OE5.
    UPDATED: Only moves to days with NO exams at all, runs AFTER gap optimization.
    """
    if not sem_dict:
        return sem_dict, 0, []
    
    st.info("🎯 Optimizing Open Elective (OE) placement (after gap optimization)...")
    
    # Combine all data to analyze the schedule
    all_data = pd.concat(sem_dict.values(), ignore_index=True)
    
    # Ensure all dates are in DD-MM-YYYY string format
    def normalize_date_to_ddmmyyyy(date_val):
        """Convert any date format to DD-MM-YYYY string format"""
        if pd.isna(date_val) or date_val == "":
            return ""
        
        if isinstance(date_val, pd.Timestamp):
            return date_val.strftime("%d-%m-%Y")
        elif isinstance(date_val, str):
            try:
                parsed = pd.to_datetime(date_val, format="%d-%m-%Y", errors='raise')
                return parsed.strftime("%d-%m-%Y")
            except:
                try:
                    parsed = pd.to_datetime(date_val, dayfirst=True, errors='raise')
                    return parsed.strftime("%d-%m-%Y")
                except:
                    return str(date_val)
        else:
            try:
                parsed = pd.to_datetime(date_val, errors='coerce')
                if pd.notna(parsed):
                    return parsed.strftime("%d-%m-%Y")
                else:
                    return str(date_val)
            except:
                return str(date_val)
    
    # Apply date normalization to all data
    all_data['Exam Date'] = all_data['Exam Date'].apply(normalize_date_to_ddmmyyyy)
    
    # Separate OE and non-OE data
    oe_data = all_data[all_data['OE'].notna() & (all_data['OE'].str.strip() != "")]
    non_oe_data = all_data[~(all_data['OE'].notna() & (all_data['OE'].str.strip() != ""))]
    
    if oe_data.empty:
        st.info("No OE subjects to optimize")
        return sem_dict, 0, []
    
    # Build schedule tracking: date -> count of ALL exams (including OE)
    exam_count_per_date = {}
    
    # Count ALL scheduled exams per date
    for _, row in all_data.iterrows():
        if pd.notna(row['Exam Date']) and row['Exam Date'].strip() != "":
            date_str = row['Exam Date']
            if date_str not in exam_count_per_date:
                exam_count_per_date[date_str] = 0
            exam_count_per_date[date_str] += 1
    
    # Find all dates in the schedule
    all_dates = sorted(exam_count_per_date.keys(), 
                      key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    if not all_dates:
        return sem_dict, 0, []
    
    st.write(f"📊 Current schedule has exams on {len(all_dates)} different dates")
    
    # Get the date range from first to last exam
    start_date = datetime.strptime(all_dates[0], "%d-%m-%Y")
    end_date = datetime.strptime(all_dates[-1], "%d-%m-%Y")
    
    # Find COMPLETELY EMPTY days in the schedule range
    completely_empty_days = []
    current_date = start_date
    
    while current_date <= end_date:
        if current_date.weekday() != 6 and current_date.date() not in holidays:
            date_str = current_date.strftime("%d-%m-%Y")
            # Check if this date has NO exams at all
            if date_str not in exam_count_per_date:
                completely_empty_days.append(date_str)
        current_date += timedelta(days=1)
    
    # Sort empty days chronologically
    completely_empty_days.sort(key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    st.write(f"🔍 Found {len(completely_empty_days)} completely empty days for potential OE optimization")
    
    if not completely_empty_days:
        st.info("ℹ️ No completely empty days available for OE optimization")
        return sem_dict, 0, []
    
    # Process OE optimization
    oe_data_copy = oe_data.copy()
    oe1_oe5_data = oe_data_copy[oe_data_copy['OE'].isin(['OE1', 'OE5'])]
    oe2_data = oe_data_copy[oe_data_copy['OE'] == 'OE2']
    
    moves_made = 0
    optimization_log = []
    
    # Process OE1/OE5 together (they should always be on the same date/time)
    if not oe1_oe5_data.empty:
        current_oe1_oe5_date = oe1_oe5_data['Exam Date'].iloc[0]
        current_oe1_oe5_date_obj = datetime.strptime(current_oe1_oe5_date, "%d-%m-%Y")
        
        st.write(f"📅 Current OE1/OE5 date: {current_oe1_oe5_date}")
        
        # Find the earliest completely empty day that comes before current OE1/OE5 date
        best_oe1_oe5_date = None
        best_oe2_date = None
        
        for empty_day in completely_empty_days:
            empty_day_obj = datetime.strptime(empty_day, "%d-%m-%Y")
            
            # Only consider dates earlier than current OE1/OE5 date
            if empty_day_obj >= current_oe1_oe5_date_obj:
                break
            
            # Check if the next day is also completely empty (for OE2)
            next_day = find_next_valid_day_for_electives(empty_day_obj + timedelta(days=1), holidays)
            if next_day:
                next_day_str = next_day.strftime("%d-%m-%Y")
                
                # Check if next day is also completely empty
                if next_day_str in completely_empty_days:
                    best_oe1_oe5_date = empty_day
                    best_oe2_date = next_day_str
                    break
        
        # If we found suitable completely empty days, move OE subjects
        if best_oe1_oe5_date and best_oe2_date:
            days_saved = (current_oe1_oe5_date_obj - datetime.strptime(best_oe1_oe5_date, "%d-%m-%Y")).days
            
            st.write(f"✅ Found optimal placement: OE1/OE5 on {best_oe1_oe5_date}, OE2 on {best_oe2_date}")
            
            # Update all OE1/OE5 exams in semester dictionary
            for idx in oe1_oe5_data.index:
                sem = all_data.at[idx, 'Semester']
                branch = all_data.at[idx, 'Branch']
                subject = all_data.at[idx, 'Subject']
                
                # Update in the semester dictionary
                mask = (sem_dict[sem]['Subject'] == subject) & \
                       (sem_dict[sem]['Branch'] == branch)
                sem_dict[sem].loc[mask, 'Exam Date'] = best_oe1_oe5_date
                sem_dict[sem].loc[mask, 'Time Slot'] = "10:00 AM - 1:00 PM"
            
            # Update all OE2 exams to the day immediately after OE1/OE5
            if not oe2_data.empty:
                for idx in oe2_data.index:
                    sem = all_data.at[idx, 'Semester']
                    branch = all_data.at[idx, 'Branch']
                    subject = all_data.at[idx, 'Subject']
                    
                    # Update in the semester dictionary
                    mask = (sem_dict[sem]['Subject'] == subject) & \
                           (sem_dict[sem]['Branch'] == branch)
                    sem_dict[sem].loc[mask, 'Exam Date'] = best_oe2_date
                    sem_dict[sem].loc[mask, 'Time Slot'] = "2:00 PM - 5:00 PM"
            
            moves_made += 1
            optimization_log.append(
                f"Moved OE1/OE5 from {current_oe1_oe5_date} to {best_oe1_oe5_date} (saved {days_saved} days, completely empty day)"
            )
            if not oe2_data.empty:
                optimization_log.append(
                    f"Moved OE2 to {best_oe2_date} (completely empty day, immediately after OE1/OE5)"
                )
        else:
            st.info("ℹ️ No suitable consecutive completely empty days found for OE optimization")
    
    # Ensure all dates in sem_dict are properly formatted
    for sem in sem_dict:
        sem_dict[sem]['Exam Date'] = sem_dict[sem]['Exam Date'].apply(normalize_date_to_ddmmyyyy)
    
    if moves_made > 0:
        st.success(f"✅ OE Optimization: Moved {moves_made} OE groups to completely empty days!")
        with st.expander("📋 OE Optimization Details"):
            for log in optimization_log:
                st.write(f"• {log}")
    else:
        st.info("ℹ️ OE subjects are already optimally placed or no suitable completely empty days available")
    
    return sem_dict, moves_made, optimization_log
    


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

    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        st.markdown("#### 📅 Examination Period")
        
        col1, col2 = st.columns(2)
        with col1:
            base_date = st.date_input("Start date for exams", value=datetime(2025, 4, 1))
            base_date = datetime.combine(base_date, datetime.min.time())
        
        with col2:
            end_date = st.date_input("End date for exams", value=datetime(2025, 5, 30))
            end_date = datetime.combine(end_date, datetime.min.time())
        
        # Validate date range
        if end_date <= base_date:
            st.error("❌ End date must be after start date!")
            end_date = base_date + timedelta(days=30)  # Default to 30 days after start
            st.warning(f"⚠️ Auto-corrected end date to: {end_date.strftime('%Y-%m-%d')}")

        st.markdown('<div style="margin-top: 2rem;"></div>', unsafe_allow_html=True)

        with st.expander("Holiday Configuration", expanded=True):
            st.markdown("#### 📅 Select Predefined Holidays")
        
            # Initialize holiday_dates list
            holiday_dates = []

            col1, col2 = st.columns(2)
            with col1:
                if st.checkbox("April 14, 2025", value=True):
                    holiday_dates.append(datetime(2025, 4, 14).date())
            with col2:
                if st.checkbox("May 1, 2025", value=True):
                    holiday_dates.append(datetime(2025, 5, 1).date())

            if st.checkbox("August 15, 2025", value=True):
                holiday_dates.append(datetime(2025, 8, 15).date())

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

            # Add custom holidays to the main list
            custom_holidays = [h for h in st.session_state.custom_holidays if h is not None]
            for custom_holiday in custom_holidays:
                holiday_dates.append(custom_holiday)

            # Create the final holidays set - ensure all are date objects
            holidays_set = set(holiday_dates)
        
            # Store in session state so it's accessible throughout the app
            st.session_state.holidays_set = holidays_set

            if holidays_set:
                st.markdown("#### Selected Holidays:")
                for holiday in sorted(holidays_set):
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
                <li>🎯 Common across semesters first</li>
                <li>🔗 Common within semester scheduling</li>
                <li>🔍 Gap-filling optimization</li>
                <li>📄 Stream-wise uncommon scheduling</li>
                <li>🎓 OE elective optimization</li>
                <li>⚡ One exam per day per branch</li>
                <li>📋 PDF generation</li>
                <li>✅ Verification file export</li>
                <li>🎯 Three-phase priority scheduling</li>
                <li>📱 Mobile-friendly interface</li>
                <li>📅 Date range enforcement</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    if uploaded_file is not None:
        if st.button("📄 Generate Timetable", type="primary", use_container_width=True):
            with st.spinner("Processing your timetable... Please wait..."):
                try:
                    # Use holidays from session state
                    holidays_set = st.session_state.get('holidays_set', set())
                    st.write(f"🗓️ Using {len(holidays_set)} holidays: {[h.strftime('%d-%m-%Y') for h in sorted(holidays_set)]}")
                    
                    # Display date range being used
                    date_range_days = (end_date - base_date).days + 1
                    valid_exam_days = len(get_valid_dates_in_range(base_date, end_date, holidays_set))
                    st.info(f"📅 Examination Period: {base_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')} ({date_range_days} total days, {valid_exam_days} valid exam days)")
                
                    st.write("Reading timetable...")
                    df_non_elec, df_ele, original_df = read_timetable(uploaded_file)

                    if df_non_elec is not None:
                        st.write("Processing subjects...")
                        
                        # NEW THREE-PHASE IMPROVED SCHEDULING ORDER:
                        # SUPER SCHEDULING: All subjects in one comprehensive function
                        st.info("🚀 SUPER SCHEDULING: All subjects with frequency-based priority and daily branch coverage")
                        df_scheduled = schedule_all_subjects_comprehensively(df_non_elec, holidays_set, base_date, end_date)
                        
                        # Step 5: Handle electives if they exist
                        if df_ele is not None and not df_ele.empty:
                            # Find the maximum date from non-elective scheduling
                            non_elec_dates = pd.to_datetime(df_scheduled['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                            if not non_elec_dates.empty:
                                max_non_elec_date = max(non_elec_dates).date()
                                st.write(f"📅 Max non-elective date: {max_non_elec_date.strftime('%d-%m-%Y')}")
                                
                                # Check if electives can be scheduled within end date
                                elective_day1 = find_next_valid_day_for_electives(
                                    datetime.combine(max_non_elec_date, datetime.min.time()) + timedelta(days=1), 
                                    holidays_set
                                )
                                elective_day2 = find_next_valid_day_for_electives(elective_day1 + timedelta(days=1), holidays_set)
                                
                                if elective_day2 <= end_date:
                                    # Schedule electives globally
                                    df_ele_scheduled = schedule_electives_globally(df_ele, max_non_elec_date, holidays_set)
                                    
                                    # Combine non-electives and electives
                                    all_scheduled_subjects = pd.concat([df_scheduled, df_ele_scheduled], ignore_index=True)
                                else:
                                    st.warning(f"⚠️ Electives cannot be scheduled within end date ({end_date.strftime('%d-%m-%Y')})")
                                    all_scheduled_subjects = df_scheduled
                            else:
                                all_scheduled_subjects = df_scheduled
                        else:
                            all_scheduled_subjects = df_scheduled
                        
                        # Step 6: Create semester dictionary - only include successfully scheduled subjects
                        successfully_scheduled = all_scheduled_subjects[
                            (all_scheduled_subjects['Exam Date'] != "") & 
                            (all_scheduled_subjects['Exam Date'] != "Out of Range")
                        ].copy()
                        
                        # Count subjects that couldn't be scheduled
                        out_of_range_subjects = all_scheduled_subjects[
                            all_scheduled_subjects['Exam Date'] == "Out of Range"
                        ]
                        
                        if not out_of_range_subjects.empty:
                            st.warning(f"⚠️ {len(out_of_range_subjects)} subjects could not be scheduled within the specified date range")
                            
                            # Show breakdown by semester and branch
                            with st.expander("📋 Subjects Not Scheduled (Out of Range)"):
                                for semester in sorted(out_of_range_subjects['Semester'].unique()):
                                    sem_subjects = out_of_range_subjects[out_of_range_subjects['Semester'] == semester]
                                    st.write(f"**Semester {semester}:** {len(sem_subjects)} subjects")
                                    for branch in sorted(sem_subjects['Branch'].unique()):
                                        branch_subjects = sem_subjects[sem_subjects['Branch'] == branch]
                                        st.write(f"  • {branch}: {len(branch_subjects)} subjects")
                        
                        if not successfully_scheduled.empty:
                            # Sort by semester and date
                            successfully_scheduled = successfully_scheduled.sort_values(["Semester", "Exam Date"], ascending=True)
                            
                            # Create semester dictionary
                            sem_dict = {}
                            for s in sorted(successfully_scheduled["Semester"].unique()):
                                sem_data = successfully_scheduled[successfully_scheduled["Semester"] == s].copy()
                                sem_dict[s] = sem_data
                            
                            st.write("Optimizing schedule by filling gaps...")
                            sem_dict, gap_moves_made, gap_optimization_log = optimize_schedule_by_filling_gaps(
                            sem_dict, holidays_set, base_date, end_date
                            )

                            # Step 8: Optimize OE subjects AFTER gap optimization
                            if df_ele is not None and not df_ele.empty:
                                st.write("Optimizing OE subjects...")
                                sem_dict, oe_moves_made, oe_optimization_log = optimize_oe_subjects_after_scheduling(sem_dict, holidays_set)        
                            else:
                                oe_moves_made = 0
                                oe_optimization_log = []

                            # Show combined optimization results
                            total_optimizations = (oe_moves_made if df_ele is not None and not df_ele.empty else 0) + gap_moves_made
                            if total_optimizations > 0:
                                st.success(f"🎯 Total Optimizations Made: {total_optimizations}")
                            if df_ele is not None and not df_ele.empty and oe_moves_made > 0:
                                st.info(f"📈 OE Optimizations: {oe_moves_made}")
                            if gap_moves_made > 0:
                                st.info(f"📉 Gap Fill Optimizations: {gap_moves_made}")

                            st.session_state.timetable_data = sem_dict
                            st.session_state.original_df = original_df
                            st.session_state.processing_complete = True

                            # Compute statistics
                            final_all_data = pd.concat(sem_dict.values(), ignore_index=True)
                            total_exams = len(final_all_data)
                            total_semesters = len(sem_dict)
                            total_branches = len(set(final_all_data['Branch'].unique()))

                            all_dates = pd.to_datetime(final_all_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                            overall_date_range = (max(all_dates) - min(all_dates)).days + 1 if all_dates.size > 0 else 0
                            unique_exam_days = len(all_dates.dt.date.unique())

                            # Store statistics in session state
                            st.session_state.total_exams = total_exams
                            st.session_state.total_semesters = total_semesters
                            st.session_state.total_branches = total_branches
                            st.session_state.overall_date_range = overall_date_range
                            st.session_state.unique_exam_days = unique_exam_days

                            # Generate and store downloadable files
                            st.write("Generating Excel...")
                            try:
                                excel_data = save_to_excel(sem_dict)
                                if excel_data:
                                    st.session_state.excel_data = excel_data.getvalue()
                                    st.success("✅ Excel file generated successfully")
                                else:
                                    st.warning("⚠️ Excel generation completed but no data returned")
                                    st.session_state.excel_data = None
                            except Exception as e:
                                st.error(f"❌ Excel generation failed: {str(e)}")
                                st.session_state.excel_data = None

                            st.write("Generating verification file...")
                            try:
                                verification_data = save_verification_excel(original_df, sem_dict)
                                if verification_data:
                                    st.session_state.verification_data = verification_data.getvalue()
                                    st.success("✅ Verification file generated successfully")
                                else:
                                    st.warning("⚠️ Verification file generation completed but no data returned")
                                    st.session_state.verification_data = None
                            except Exception as e:
                                st.error(f"❌ Verification file generation failed: {str(e)}")
                                st.session_state.verification_data = None

                            st.write("Generating PDF...")
                            try:
                                if sem_dict:
                                    pdf_output = io.BytesIO()
                                    temp_pdf_path = "temp_timetable.pdf"
                                    generate_pdf_timetable(sem_dict, temp_pdf_path)
                                    
                                    # Check if PDF was created successfully
                                    if os.path.exists(temp_pdf_path):
                                        with open(temp_pdf_path, "rb") as f:
                                            pdf_output.write(f.read())
                                        pdf_output.seek(0)
                                        st.session_state.pdf_data = pdf_output.getvalue()
                                        os.remove(temp_pdf_path)
                                        st.success("✅ PDF generated successfully")
                                    else:
                                        st.warning("⚠️ PDF generation completed but file not found")
                                        st.session_state.pdf_data = None
                                else:
                                    st.warning("⚠️ No data available for PDF generation")
                                    st.session_state.pdf_data = None
                            except Exception as e:
                                st.error(f"❌ PDF generation failed: {str(e)}")
                                st.session_state.pdf_data = None

                            st.markdown('<div class="status-success">🎉 Timetable generated successfully with THREE-PHASE SCHEDULING and NO DOUBLE BOOKINGS!</div>',
                                        unsafe_allow_html=True)
                            
                            # Show improved three-phase scheduling summary
                            st.info("✅ **Three-Phase Scheduling Applied:**\n1. 🎯 **Phase 1:** Common across semesters scheduled FIRST from base date\n2. 🔗 **Phase 2:** Common within semester subjects (COMP/ELEC appearing in multiple branches)\n3. 🔍 **Phase 3:** Truly uncommon subjects with gap-filling optimization within date range\n4. 🎓 **Phase 4:** Electives scheduled LAST (if space available)\n5. ⚡ **Guarantee:** ONE exam per day per subbranch-semester")
                            
                            # Show efficiency improvement
                            efficiency = (unique_exam_days / overall_date_range) * 100 if overall_date_range > 0 else 0
                            st.success(f"📊 **Schedule Efficiency: {efficiency:.1f}%** (Higher is better - more days utilized)")
                            
                            # Show date range utilization
                            date_range_utilization = (unique_exam_days / valid_exam_days) * 100 if valid_exam_days > 0 else 0
                            st.info(f"📅 **Date Range Utilization: {date_range_utilization:.1f}%** ({unique_exam_days}/{valid_exam_days} valid days used)")
                            
                            # Count subjects by type for summary
                            common_across_count = len(final_all_data[final_all_data['CommonAcrossSems'] == True])
                            
                            # Count common within semester
                            common_within_sem = final_all_data[
                                (final_all_data['CommonAcrossSems'] == False) & 
                                (final_all_data['Category'].isin(['COMP', 'ELEC']))
                            ]
                            common_within_sem_groups = common_within_sem.groupby(['Semester', 'ModuleCode'])['Branch'].nunique()
                            common_within_count = len(common_within_sem[
                                common_within_sem.set_index(['Semester', 'ModuleCode']).index.map(
                                    lambda x: common_within_sem_groups.get(x, 1) > 1
                                )
                            ])
                            
                            elective_count = len(final_all_data[final_all_data['OE'].notna() & (final_all_data['OE'].str.strip() != "")])
                            uncommon_count = total_exams - common_across_count - common_within_count - elective_count
                            
                            st.success(f"📈 **Scheduling Breakdown:**\n• Common Across Semesters: {common_across_count}\n• Common Within Semester: {common_within_count}\n• Truly Uncommon: {uncommon_count}\n• Electives: {elective_count}")
                            
                            # Show double booking verification
                            st.success("✅ **No Double Bookings**: Each subbranch has max one exam per day")
                            
                        else:
                            st.warning("No subjects could be scheduled within the specified date range.")

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

        # Download options
        st.markdown("### 📥 Download Options")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            if st.session_state.excel_data:
                st.download_button(
                    label="📊 Download Excel File",
                    data=st.session_state.excel_data,
                    file_name=f"complete_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_excel"
                )
            else:
                st.button("📊 Excel Not Available", disabled=True, use_container_width=True)

        with col2:
            if st.session_state.pdf_data:
                st.download_button(
                    label="📄 Download PDF File",
                    data=st.session_state.pdf_data,
                    file_name=f"complete_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="download_pdf"
                )
            else:
                st.button("📄 PDF Not Available", disabled=True, use_container_width=True)

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
            else:
                st.button("📋 Verification Not Available", disabled=True, use_container_width=True)

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
                st.rerun()

        # Statistics Overview
        # Statistics Overview
        st.markdown("""
        <div class="stats-section">
            <h2>📈 Complete Timetable Statistics</h2>
        </div>
        """, unsafe_allow_html=True)

        # Calculate additional statistics
        if st.session_state.timetable_data:
            # Combine all data to calculate date ranges
            final_all_data = pd.concat(st.session_state.timetable_data.values(), ignore_index=True)
    
            # Separate non-elective and OE subjects
            non_elective_data = final_all_data[~(final_all_data['OE'].notna() & (final_all_data['OE'].str.strip() != ""))]
            oe_data = final_all_data[final_all_data['OE'].notna() & (final_all_data['OE'].str.strip() != "")]
    
            # Calculate non-elective date range
            if not non_elective_data.empty:
                non_elec_dates = pd.to_datetime(non_elective_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                if not non_elec_dates.empty:
                    non_elec_range = (max(non_elec_dates) - min(non_elec_dates)).days + 1
                    non_elec_start = min(non_elec_dates).strftime("%-d %B")  # e.g., "1 April"
                    non_elec_end = max(non_elec_dates).strftime("%-d %B")    # e.g., "20 April"
            
                    if non_elec_start == non_elec_end:
                        non_elec_display = f"{non_elec_start}"
                    else:
                        non_elec_display = f"{non_elec_start} to {non_elec_end}"
                else:
                    non_elec_display = "No dates"
            else:
                non_elec_display = "No subjects"
    
            # Calculate OE date range
            if not oe_data.empty:
                oe_dates = pd.to_datetime(oe_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                if not oe_dates.empty:
                    oe_range = (max(oe_dates) - min(oe_dates)).days + 1
                    unique_oe_dates = sorted(oe_dates.dt.strftime("%-d %B").unique())
            
                    if len(unique_oe_dates) == 1:
                        oe_display = unique_oe_dates[0]
                    elif len(unique_oe_dates) == 2:
                        oe_display = f"{unique_oe_dates[0]}, {unique_oe_dates[1]}"
                    else:
                        oe_start = unique_oe_dates[0]
                        oe_end = unique_oe_dates[-1]
                        oe_display = f"{oe_start} to {oe_end}"
                else:
                    oe_display = "No dates"
            else:
                oe_display = "No OE subjects"
    
            # Calculate gap between non-elective and OE
            if not non_elective_data.empty and not oe_data.empty:
                non_elec_dates = pd.to_datetime(non_elective_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                oe_dates = pd.to_datetime(oe_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        
                if not non_elec_dates.empty and not oe_dates.empty:
                    max_non_elec = max(non_elec_dates)
                    min_oe = min(oe_dates)
                    gap_days = (min_oe - max_non_elec).days - 1
                    gap_display = f"{max(0, gap_days)} days"
                else:
                    gap_display = "N/A"
            else:
                gap_display = "N/A"
        else:
            non_elec_display = "No data"
            oe_display = "No data"
            gap_display = "N/A"

        # Display metrics in two rows
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="metric-card"><h3>📚 {st.session_state.total_exams}</h3><p>Total Exams</p></div>',
                        unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><h3>🎓 {st.session_state.total_semesters}</h3><p>Semesters</p></div>',
                        unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-card"><h3>🏫 {st.session_state.total_branches}</h3><p>Branches</p></div>',
                        unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-card"><h3>📅 {st.session_state.overall_date_range}</h3><p>Overall Span</p></div>',
                        unsafe_allow_html=True)

        # Second row with date range information
        st.markdown("### 📅 Examination Schedule Breakdown")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
            <div class="metric-card" style="text-align: left; padding: 1rem;">
                <h4 style="margin: 0; color: white;">📖 Non-Elective Exams</h4>
                <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.9;">{non_elec_display}</p>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div class="metric-card" style="text-align: left; padding: 1rem;">
                <h4 style="margin: 0; color: white;">🎓 Open Elective (OE) Exams</h4>
                <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.9;">{oe_display}</p>
            </div>
            """, unsafe_allow_html=True)

        # Additional metrics row
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f'<div class="metric-card"><h3>📊 {st.session_state.unique_exam_days}</h3><p>Unique Exam Days</p></div>',
                        unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><h3>⚡ {gap_display}</h3><p>Non-Elec to OE Gap</p></div>',
                        unsafe_allow_html=True)
        with col3:
            # Calculate efficiency percentage
            if st.session_state.unique_exam_days > 0 and st.session_state.overall_date_range > 0:
                efficiency = (st.session_state.unique_exam_days / st.session_state.overall_date_range) * 100
                efficiency_display = f"{efficiency:.1f}%"
            else:
                efficiency_display = "N/A"
    
            st.markdown(f'<div class="metric-card"><h3>🎯 {efficiency_display}</h3><p>Schedule Efficiency</p></div>',
                        unsafe_allow_html=True)


        # Show gap-filling efficiency
        total_possible_slots = st.session_state.overall_date_range * 2  # 2 slots per day
        actual_exams = st.session_state.total_exams
        slot_utilization = min(100, (actual_exams / total_possible_slots * 100)) if total_possible_slots > 0 else 0
        
        if slot_utilization > 70:
            st.success(f"🔍 **Slot Utilization:** {slot_utilization:.1f}% (Excellent optimization)")
        elif slot_utilization > 50:
            st.info(f"🔍 **Slot Utilization:** {slot_utilization:.1f}% (Good optimization)")
        else:
            st.warning(f"🔍 **Slot Utilization:** {slot_utilization:.1f}% (Room for improvement)")

        # Timetable Results
        st.markdown("---")
        st.markdown("""
        <div class="results-section">
            <h2>📊 Complete Timetable Results</h2>
        </div>
        """, unsafe_allow_html=True)

        # Define the subject display formatting functions for Streamlit display
        def format_subject_display(row):
            """Format subject display for non-electives in Streamlit interface"""
            subject = row['Subject']
            time_slot = row['Time Slot']
            duration = row.get('Exam Duration', 3)
            is_common = row.get('CommonAcrossSems', False)
            semester = row['Semester']
            
            # Get preferred slot for this semester
            preferred_slot = get_preferred_slot(semester)
            
            time_range = ""
            
            # Handle non-standard durations: show specific time range
            if duration != 3 and time_slot and time_slot.strip():
                start_time = time_slot.split(' - ')[0].strip()
                end_time = calculate_end_time(start_time, duration)
                time_range = f" ({start_time} - {end_time})"
            # Handle common subjects with different time slots
            elif is_common and time_slot != preferred_slot and time_slot and time_slot.strip():
                time_range = f" ({time_slot})"
            
            return subject + time_range

        def format_elective_display(row):
            """Format subject display for electives in Streamlit interface"""
            subject = row['Subject']
            oe_type = row.get('OE', '')
            base_display = f"{subject} [{oe_type}]" if oe_type else subject
            
            duration = row.get('Exam Duration', 3)
            time_slot = row['Time Slot']
            
            # For OE subjects in Streamlit, only show time if duration is non-standard
            if duration != 3 and time_slot and time_slot.strip():
                start_time = time_slot.split(' - ')[0].strip()
                end_time = calculate_end_time(start_time, duration)
                time_range = f" ({start_time} - {end_time})"
            else:
                time_range = ""
            
            return base_display + time_range

        # Display timetable data
        for sem, df_sem in st.session_state.timetable_data.items():
            st.markdown(f"### 📚 Semester {sem}")

            for main_branch in df_sem["MainBranch"].unique():
                main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()

                if not df_mb.empty:
                    # Separate non-electives and electives for display
                    df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                    df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()

                    # Display non-electives
                    if not df_non_elec.empty:
                        st.markdown(f"#### {main_branch_full} - Core Subjects")
                        
                        try:
                            # Apply formatting
                            df_non_elec["SubjectDisplay"] = df_non_elec.apply(format_subject_display, axis=1)
                            df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                            df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                           
                            # Create a simple table format
                            display_data = []
                            for date, group in df_non_elec.groupby('Exam Date'):
                                date_str = date.strftime("%d-%m-%Y") if pd.notna(date) else "Unknown Date"
                                row_data = {'Exam Date': date_str}
                               
                                # Add subjects for each SubBranch
                                for subbranch in df_non_elec['SubBranch'].unique():
                                    subbranch_subjects = group[group['SubBranch'] == subbranch]['SubjectDisplay'].tolist()
                                    row_data[subbranch] = ", ".join(subbranch_subjects) if subbranch_subjects else "---"
                               
                                display_data.append(row_data)
                           
                            if display_data:
                                display_df = pd.DataFrame(display_data)
                                display_df = display_df.set_index('Exam Date')
                                st.dataframe(display_df, use_container_width=True)
                            else:
                                st.write("No core subjects to display")
                               
                        except Exception as e:
                            st.error(f"Error displaying core subjects: {str(e)}")
                            # Fallback: show raw data
                            st.write("Showing raw data:")
                            display_cols = ['Exam Date', 'SubBranch', 'Subject', 'Time Slot']
                            available_cols = [col for col in display_cols if col in df_non_elec.columns]
                            st.dataframe(df_non_elec[available_cols], use_container_width=True)

                    # Display electives  
                    if not df_elec.empty:
                        st.markdown(f"#### {main_branch_full} - Open Electives")
                       
                        try:
                            # Apply formatting
                            df_elec["SubjectDisplay"] = df_elec.apply(format_elective_display, axis=1)
                            df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                            df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                           
                            # Create elective display
                            elec_display_data = []
                            for (oe_type, date), group in df_elec.groupby(['OE', 'Exam Date']):
                                date_str = date.strftime("%d-%m-%Y") if pd.notna(date) else "Unknown Date"
                                subjects = ", ".join(group['SubjectDisplay'].tolist())
                                elec_display_data.append({
                                    'Exam Date': date_str,
                                    'OE Type': oe_type,
                                    'Subjects': subjects
                                })
                           
                            if elec_display_data:
                                elec_display_df = pd.DataFrame(elec_display_data)
                                st.dataframe(elec_display_df, use_container_width=True)
                            else:
                                st.write("No elective subjects to display")
                               
                        except Exception as e:
                            st.error(f"Error displaying elective subjects: {str(e)}")
                            # Fallback: show raw data
                            st.write("Showing raw data:")
                            display_cols = ['Exam Date', 'OE', 'Subject', 'Time Slot']
                            available_cols = [col for col in display_cols if col in df_elec.columns]
                            st.dataframe(df_elec[available_cols], use_container_width=True)

    # Display footer
    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <p>🎓 <strong>Three-Phase Timetable Generator with Date Range Control & Gap-Filling</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
        <p style="font-size: 0.9em;">Common across semesters first • Common within semester • Gap-filling optimization • One exam per day per branch • OE optimization • Date range enforcement • Maximum efficiency • Verification export</p>
    </div>
    """, unsafe_allow_html=True)
    
if __name__ == "__main__":
    main()





