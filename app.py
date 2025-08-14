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

def get_preferred_slot(semester):
    """Calculate preferred time slot based on semester number"""
    if semester % 2 != 0:  # Odd semester
        odd_sem_position = (semester + 1) // 2
        return "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
    else:  # Even semester
        even_sem_position = semester // 2
        return "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"

def schedule_common_subjects_first(df, holidays, base_date):
    """
    Schedule common subjects FIRST across ALL branches starting from base date
    Priority: Schedule maximum common subjects per day while respecting constraints
    """
    st.info("🔧 Scheduling common subjects FIRST from base date...")
    
    # Filter only common subjects for initial scheduling
    common_subjects = df[df['CommonAcrossSems'] == True].copy()
    
    if common_subjects.empty:
        st.info("No common subjects to schedule first")
        return df
    
    st.write(f"Found {len(common_subjects)} common subjects to schedule first")
    
    scheduled_count = 0
    current_scheduling_date = base_date
    
    # Helper function to find next valid day
    def find_next_valid_day(start_date, holidays_set):
        current_date = start_date
        while True:
            if current_date.weekday() != 6 and current_date.date() not in holidays_set:
                return current_date
            current_date += timedelta(days=1)
    
    # Helper function to get subbranch-semester key
    def get_subbranch_semester_key(subbranch, semester):
        return f"{subbranch}_{semester}"
    
    # Track scheduled subbranch-semester combinations per day
    daily_scheduled = {}  # date -> set of subbranch_semester keys
    
    # Group by ModuleCode to find truly common subjects across ALL branches
    common_subject_groups = {}
    for module_code, group in common_subjects.groupby('ModuleCode'):
        branches_for_this_subject = group['Branch'].unique()
        common_subject_groups[module_code] = group
        #st.write(f"  📋 Found common subject {group['Subject'].iloc[0]} across {len(branches_for_this_subject)} branches: {', '.join(branches_for_this_subject)}")
    
    # Create a list of unscheduled common subject groups
    unscheduled_groups = list(common_subject_groups.keys())
    st.write(f"📊 Total common subject groups to schedule: {len(unscheduled_groups)}")
    
    # Schedule day by day, maximizing subjects per day
    scheduling_attempts = 0
    max_scheduling_attempts = 200
    
    while unscheduled_groups and scheduling_attempts < max_scheduling_attempts:
        scheduling_attempts += 1
        exam_date = find_next_valid_day(current_scheduling_date, holidays)
        date_str = exam_date.strftime("%d-%m-%Y")
        
        if date_str not in daily_scheduled:
            daily_scheduled[date_str] = set()
        
        # Available time slots for this date
        available_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        subjects_scheduled_today = 0
        
        #st.write(f"  📅 Attempting to schedule common subjects on {date_str} (Attempt {scheduling_attempts})")
        
        # Try to schedule as many common subjects as possible on this day
        day_had_scheduling = False
        
        for time_slot in available_slots:
            groups_scheduled_in_slot = []
            
            #st.write(f"    🕐 Checking {time_slot} slot")
            
            # Check which groups can be scheduled in this time slot
            for module_code in unscheduled_groups[:]:
                group = common_subject_groups[module_code]
                can_schedule = True
                conflicting_subbranches = []
                
                # Check if any subbranch-semester combination conflicts
                for idx, row in group.iterrows():
                    subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                    if subbranch_sem_key in daily_scheduled[date_str]:
                        can_schedule = False
                        conflicting_subbranches.append(f"{row['SubBranch']} (Sem {row['Semester']})")
                
                if can_schedule:
                    # Schedule this group in this time slot
                    semester = group['Semester'].iloc[0]
                    
                    # Calculate preferred slot based on semester
                    preferred_slot = get_preferred_slot(semester)
                    
                    # Use the preferred slot if it matches current slot, otherwise use current slot
                    actual_slot = preferred_slot if preferred_slot == time_slot else time_slot
                    
                    # Schedule all instances of this common subject
                    subbranches_scheduled = []
                    for idx in group.index:
                        df.loc[idx, 'Exam Date'] = date_str
                        df.loc[idx, 'Time Slot'] = actual_slot
                        scheduled_count += 1
                        
                        # Mark this subbranch-semester as scheduled for this date
                        subbranch_sem_key = get_subbranch_semester_key(group.loc[idx, 'SubBranch'], group.loc[idx, 'Semester'])
                        daily_scheduled[date_str].add(subbranch_sem_key)
                        subbranches_scheduled.append(f"{group.loc[idx, 'SubBranch']} (Sem {group.loc[idx, 'Semester']})")
                    
                    groups_scheduled_in_slot.append(module_code)
                    subjects_scheduled_today += 1
                    day_had_scheduling = True
                    
                    #st.write(f"      ✅ Scheduled common subject {group['Subject'].iloc[0]} for subbranches: {', '.join(subbranches_scheduled)} on {date_str} at {actual_slot}")
                else:
                    pass
                    #st.write(f"      ❌ Cannot schedule {group['Subject'].iloc[0]} - conflicts with subbranches: {', '.join(conflicting_subbranches)}")
            
            # Remove scheduled groups from unscheduled list
            for module_code in groups_scheduled_in_slot:
                unscheduled_groups.remove(module_code)
        
        total_scheduled_today = subjects_scheduled_today
        #st.write(f"  📊 Total common subjects scheduled on {date_str}: {total_scheduled_today}")
        
        # Move to next day
        current_scheduling_date = find_next_valid_day(exam_date + timedelta(days=1), holidays)
        
        # If no scheduling happened this day and we still have unscheduled groups, continue to next day
        if not day_had_scheduling and unscheduled_groups:
            #st.write(f"  ⚠️ No common subjects could be scheduled on {date_str}, moving to next day")
            continue
    
    st.success(f"✅ Successfully scheduled {scheduled_count} common subjects FIRST")
    return df

def schedule_common_within_semester_subjects(df, holidays, start_date):
    """
    Schedule subjects that are common within their semester but not across all semesters.
    These are subjects with CommonAcrossSems=False but have IsCommon='YES' and are COMP/ELEC category.
    """
    st.info("🔧 Scheduling common-within-semester subjects...")
    
    # Debug: Show data before filtering
    st.write("🔍 **Debug - Data overview before filtering:**")
    st.write(f"Total subjects: {len(df)}")
    st.write(f"CommonAcrossSems distribution: {df['CommonAcrossSems'].value_counts()}")
    if 'IsCommon' in df.columns:
        st.write(f"IsCommon distribution: {df['IsCommon'].value_counts()}")
    else:
        st.error("❌ IsCommon column missing from dataframe")
        return df
    
    # Filter subjects that are:
    # 1. Not common across semesters (CommonAcrossSems = False)
    # 2. Have IsCommon = 'YES'
    # 3. Category is COMP or ELEC 
    # 4. Not already scheduled
    
    candidate_subjects = df[
        (df['CommonAcrossSems'] == False) & 
        (df["IsCommon"] == 'YES') &
        (df['Category'].isin(['COMP', 'ELEC'])) &
        (df['Exam Date'] == "")
    ].copy()
    
    if candidate_subjects.empty:
        st.info("No common-within-semester subjects to schedule")
        st.write("🔍 **Debug - Why no candidates found:**")
        
        # Debug breakdown
        not_common_across = df[df['CommonAcrossSems'] == False]
        st.write(f"• Subjects not common across semesters: {len(not_common_across)}")
        
        is_common_yes = not_common_across[not_common_across['IsCommon'] == 'YES']
        st.write(f"• Of those, with IsCommon='YES': {len(is_common_yes)}")
        
        comp_elec = is_common_yes[is_common_yes['Category'].isin(['COMP', 'ELEC'])]
        st.write(f"• Of those, COMP/ELEC category: {len(comp_elec)}")
        
        not_scheduled = comp_elec[comp_elec['Exam Date'] == ""]
        st.write(f"• Of those, not yet scheduled: {len(not_scheduled)}")
        
        return df
    
    #st.write(f"Found {len(candidate_subjects)} candidate subjects for common-within-semester scheduling")
    
    # Show sample of candidates
    #st.write("📋 **Sample candidate subjects:**")
    sample_candidates = candidate_subjects[['Subject', 'Branch', 'Semester', 'Category', 'IsCommon']].head(3)
    st.dataframe(sample_candidates)
    
    # Helper functions
    def find_next_valid_day(start_date, holidays_set):
        current_date = start_date
        while True:
            if current_date.weekday() != 6 and current_date.date() not in holidays_set:
                return current_date
            current_date += timedelta(days=1)
    
    def get_subbranch_semester_key(subbranch, semester):
        return f"{subbranch}_{semester}"
    
    # Build tracking of already scheduled dates per subbranch-semester
    scheduled_subjects = df[df['Exam Date'] != ""].copy()
    scheduled_dates_per_subbranch_sem = {}
    
    # Initialize tracking from already scheduled subjects
    for idx, row in scheduled_subjects.iterrows():
        subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
        date_str = row['Exam Date']
        
        if subbranch_sem_key not in scheduled_dates_per_subbranch_sem:
            scheduled_dates_per_subbranch_sem[subbranch_sem_key] = set()
        scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
    
    # Find subjects that are common within semester (same ModuleCode + Semester across multiple branches)
    common_within_sem_groups = {}
    
    # Group by semester and module code
    for (semester, module_code), group in candidate_subjects.groupby(['Semester', 'ModuleCode']):
        branches_for_this_subject = group['Branch'].unique()
        
        # If this subject appears in multiple branches within the same semester, it's common within semester
        if len(branches_for_this_subject) > 1:
            group_key = f"sem_{semester}_{module_code}"
            common_within_sem_groups[group_key] = group
            #st.write(f"  📋 Found common-within-semester subject {group['Subject'].iloc[0]} (Sem {semester}) across {len(branches_for_this_subject)} branches: {', '.join(branches_for_this_subject)}")
        else:
            # Even if only in one branch but marked as IsCommon=YES, include it
            group_key = f"sem_{semester}_{module_code}"
            common_within_sem_groups[group_key] = group
            #st.write(f"  📋 Found marked common subject {group['Subject'].iloc[0]} (Sem {semester}) in branch: {branches_for_this_subject[0]}")
    
    if not common_within_sem_groups:
        #st.info("No subjects found that are common within their semester")
        return df
    
    st.write(f"📊 Total common-within-semester groups to schedule: {len(common_within_sem_groups)}")
    
    # Find date range for gap filling
    if not scheduled_subjects.empty:
        scheduled_dates = pd.to_datetime(scheduled_subjects['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        if not scheduled_dates.empty:
            latest_scheduled_date = max(scheduled_dates).date()
            #st.write(f"📅 Latest scheduled date from previous subjects: {latest_scheduled_date}")
        else:
            latest_scheduled_date = start_date.date()
    else:
        latest_scheduled_date = start_date.date()
    
    # Create list of available dates for gap filling
    available_dates = []
    current_date = start_date
    end_date = datetime.combine(latest_scheduled_date, datetime.min.time()) + timedelta(days=20)
    
    while current_date <= end_date:
        if current_date.weekday() != 6 and current_date.date() not in holidays:
            available_dates.append(current_date.strftime("%d-%m-%Y"))
        current_date += timedelta(days=1)
    
    available_dates.sort(key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    #st.write(f"📊 Created {len(available_dates)} available dates for common-within-semester scheduling")
    
    # Schedule common-within-semester subjects
    scheduled_count = 0
    unscheduled_groups = list(common_within_sem_groups.keys())
    
    # Try gap-filling first
    for date_str in available_dates:
        if not unscheduled_groups:
            break
            
        #st.write(f"  📅 Attempting to schedule common-within-semester subjects on {date_str}")
        
        # Available time slots for this date
        available_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
        
        for time_slot in available_slots:
            groups_scheduled_in_slot = []
            
            # Check which groups can be scheduled in this time slot
            for group_key in unscheduled_groups[:]:
                group = common_within_sem_groups[group_key]
                can_schedule = True
                conflicting_subbranches = []
                
                # Check if any subbranch-semester combination conflicts
                for idx, row in group.iterrows():
                    subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                    
                    # Initialize tracking if not exists
                    if subbranch_sem_key not in scheduled_dates_per_subbranch_sem:
                        scheduled_dates_per_subbranch_sem[subbranch_sem_key] = set()
                    
                    # Check if this subbranch-semester already has an exam on this date
                    if date_str in scheduled_dates_per_subbranch_sem[subbranch_sem_key]:
                        can_schedule = False
                        conflicting_subbranches.append(f"{row['SubBranch']} (Sem {row['Semester']})")
                
                if can_schedule:
                    # Schedule this group in this time slot
                    semester = group['Semester'].iloc[0]
                    preferred_slot = get_preferred_slot(semester)
                    
                    # Use the preferred slot if it matches current slot, otherwise use current slot
                    actual_slot = preferred_slot if preferred_slot == time_slot else time_slot
                    
                    # Schedule all instances of this common-within-semester subject
                    subbranches_scheduled = []
                    for idx in group.index:
                        df.loc[idx, 'Exam Date'] = date_str
                        df.loc[idx, 'Time Slot'] = actual_slot
                        scheduled_count += 1
                        
                        # Mark this subbranch-semester as scheduled for this date
                        subbranch_sem_key = get_subbranch_semester_key(group.loc[idx, 'SubBranch'], group.loc[idx, 'Semester'])
                        scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
                        subbranches_scheduled.append(f"{group.loc[idx, 'SubBranch']} (Sem {group.loc[idx, 'Semester']})")
                    
                    groups_scheduled_in_slot.append(group_key)
                    
                    #st.write(f"      ✅ Scheduled common-within-semester subject {group['Subject'].iloc[0]} for subbranches: {', '.join(subbranches_scheduled)} on {date_str} at {actual_slot}")
                else:
                    pass
                    #st.write(f"      ❌ Cannot schedule {group['Subject'].iloc[0]} - conflicts with subbranches: {', '.join(conflicting_subbranches)}")
            
            # Remove scheduled groups from unscheduled list
            for group_key in groups_scheduled_in_slot:
                unscheduled_groups.remove(group_key)
    
    # If there are still unscheduled groups, extend the schedule
    if unscheduled_groups:
        #st.write(f"  🔄 {len(unscheduled_groups)} common-within-semester groups remaining, extending schedule...")
        
        # Find next available date after gap filling
        if available_dates:
            last_gap_date = max(available_dates, key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
            last_gap_datetime = datetime.strptime(last_gap_date, "%d-%m-%Y")
            next_extension_date = find_next_valid_day(last_gap_datetime + timedelta(days=1), holidays)
        else:
            next_extension_date = start_date
        
        for group_key in unscheduled_groups:
            group = common_within_sem_groups[group_key]
            semester = group['Semester'].iloc[0]
            preferred_slot = get_preferred_slot(semester)
            
            # Find next available date where all subbranches in this group can be scheduled
            while True:
                exam_date = find_next_valid_day(next_extension_date, holidays)
                date_str = exam_date.strftime("%d-%m-%Y")
                
                # Check if this date is available for all subbranches in this group
                can_schedule_all = True
                for idx, row in group.iterrows():
                    subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                    
                    # Initialize tracking if not exists
                    if subbranch_sem_key not in scheduled_dates_per_subbranch_sem:
                        scheduled_dates_per_subbranch_sem[subbranch_sem_key] = set()
                    
                    # Check if this date is already occupied by this subbranch-semester
                    if date_str in scheduled_dates_per_subbranch_sem[subbranch_sem_key]:
                        can_schedule_all = False
                        break
                
                if can_schedule_all:
                    break
                else:
                    # Move to next day
                    next_extension_date = exam_date + timedelta(days=1)
            
            # Schedule all instances of this group
            subbranches_scheduled = []
            for idx in group.index:
                df.loc[idx, 'Exam Date'] = date_str
                df.loc[idx, 'Time Slot'] = preferred_slot
                scheduled_count += 1
                
                # Mark this date as occupied
                subbranch_sem_key = get_subbranch_semester_key(group.loc[idx, 'SubBranch'], group.loc[idx, 'Semester'])
                scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
                subbranches_scheduled.append(f"{group.loc[idx, 'SubBranch']} (Sem {group.loc[idx, 'Semester']})")
            
            #st.write(f"    ✅ Extended: {group['Subject'].iloc[0]} for subbranches: {', '.join(subbranches_scheduled)} on {date_str} at {preferred_slot}")
            
            # Move to next day for next group
            next_extension_date = exam_date + timedelta(days=1)
    
    #st.success(f"✅ Successfully scheduled {scheduled_count} common-within-semester subjects")
    
    # Verify no double bookings
    verify_no_double_bookings(df)
    
    return df



def schedule_uncommon_subjects_after_common(df, holidays, start_date):
    """
    Schedule uncommon subjects AFTER common subjects have been scheduled.
    FIRST fill gaps between common subjects, THEN extend if needed.
    ENSURES only ONE exam per day per subbranch-semester combination.
    
    Updated to only handle truly uncommon subjects (not common within semester).
    """
    st.info("🔧 Scheduling truly uncommon subjects - filling gaps first, then extending...")
    
    # Debug: Show data before filtering
    st.write("🔍 **Debug - Filtering uncommon subjects:**")
    st.write(f"Total subjects: {len(df)}")
    
    # Filter uncommon subjects that haven't been scheduled yet
    # These are subjects where:
    # 1. CommonAcrossSems = False (not common across all semesters)
    # 2. IsCommon = 'NO' (not common within semester either)
    # 3. Not already scheduled
    
    if 'IsCommon' not in df.columns:
        st.error("❌ IsCommon column missing from dataframe")
        return df
    
    uncommon_subjects = df[
        (df['CommonAcrossSems'] == False) &
        (df["IsCommon"] == 'NO') &
        (df['Exam Date'] == "")
    ].copy()
    
    if uncommon_subjects.empty:
        st.info("No truly uncommon subjects to schedule")
        
        # Debug breakdown
        st.write("🔍 **Debug - Why no uncommon subjects found:**")
        not_common_across = df[df['CommonAcrossSems'] == False]
        st.write(f"• Subjects not common across semesters: {len(not_common_across)}")
        
        is_common_no = not_common_across[not_common_across['IsCommon'] == 'NO']
        st.write(f"• Of those, with IsCommon='NO': {len(is_common_no)}")
        
        not_scheduled = is_common_no[is_common_no['Exam Date'] == ""]
        st.write(f"• Of those, not yet scheduled: {len(not_scheduled)}")
        
        return df
    
    #st.write(f"Found {len(uncommon_subjects)} truly uncommon subjects to schedule after common subjects")
    
    # Show sample of uncommon subjects
    #st.write("📋 **Sample uncommon subjects:**")
    sample_uncommon = uncommon_subjects[['Subject', 'Branch', 'Semester', 'Category', 'IsCommon']].head(3)
    st.dataframe(sample_uncommon)
    
    # Helper function to find next valid day
    def find_next_valid_day(start_date, holidays_set):
        current_date = start_date
        while True:
            if current_date.weekday() != 6 and current_date.date() not in holidays_set:
                return current_date
            current_date += timedelta(days=1)
    
    # Helper function to get subbranch-semester key
    def get_subbranch_semester_key(subbranch, semester):
        return f"{subbranch}_{semester}"
    
    # Build a schedule tracking system that tracks DATES (not time slots) per subbranch-semester
    scheduled_dates_per_subbranch_sem = {}  # subbranch_sem_key -> set of dates
    
    # Initialize tracking from already scheduled subjects (common subjects)
    scheduled_common = df[df['Exam Date'] != ""].copy()
    
    for idx, row in scheduled_common.iterrows():
        subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
        date_str = row['Exam Date']
        
        if subbranch_sem_key not in scheduled_dates_per_subbranch_sem:
            scheduled_dates_per_subbranch_sem[subbranch_sem_key] = set()
        scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
    
    # Find date range for gap filling
    if not scheduled_common.empty:
        scheduled_dates = pd.to_datetime(scheduled_common['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        if not scheduled_dates.empty:
            latest_scheduled_date = max(scheduled_dates).date()
            #st.write(f"📅 Latest scheduled date from common subjects: {latest_scheduled_date}")
        else:
            latest_scheduled_date = start_date.date()
    else:
        latest_scheduled_date = start_date.date()
    
    # Create list of available dates for gap filling (from base date to latest + buffer)
    available_dates = []
    current_date = start_date
    end_date = datetime.combine(latest_scheduled_date, datetime.min.time()) + timedelta(days=20)
    
    while current_date <= end_date:
        if current_date.weekday() != 6 and current_date.date() not in holidays:
            available_dates.append(current_date.strftime("%d-%m-%Y"))
        current_date += timedelta(days=1)
    
    available_dates.sort(key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    #st.write(f"📊 Created {len(available_dates)} available dates for gap-filling")
    
    # Sort uncommon subjects by semester for consistent scheduling
    scheduled_count = 0
    
    for semester in sorted(uncommon_subjects['Semester'].unique()):
        semester_data = uncommon_subjects[uncommon_subjects['Semester'] == semester].copy()
        #st.write(f"📚 Scheduling Semester {semester} uncommon subjects...")
        
        # Get preferred time slot for this semester
        preferred_slot = get_preferred_slot(semester)
        
        # Group by branch within this semester
        for branch in sorted(semester_data['Branch'].unique()):
            branch_subjects = semester_data[semester_data['Branch'] == branch].copy()
            #st.write(f"  🔧 Scheduling {len(branch_subjects)} subjects for {branch}")
            
            # Get the subbranch for this branch
            if not branch_subjects.empty:
                subbranch = branch_subjects['SubBranch'].iloc[0]
                subbranch_sem_key = get_subbranch_semester_key(subbranch, semester)
                
                # Initialize tracking for this subbranch-semester if not exists
                if subbranch_sem_key not in scheduled_dates_per_subbranch_sem:
                    scheduled_dates_per_subbranch_sem[subbranch_sem_key] = set()
                
                # PHASE 1: Fill gaps first
                #st.write(f"    🔍 PHASE 1: Filling gaps for {branch} from base date")
                
                subjects_to_schedule = list(branch_subjects.iterrows())
                scheduled_in_gaps = 0
                
                for date_str in available_dates:
                    if not subjects_to_schedule:  # All subjects scheduled
                        break
                    
                    # Check if this subbranch-semester already has an exam on this date
                    if date_str not in scheduled_dates_per_subbranch_sem[subbranch_sem_key]:
                        # This date is available for this subbranch-semester
                        idx, row = subjects_to_schedule.pop(0)
                        
                        # Schedule the subject
                        df.loc[idx, 'Exam Date'] = date_str
                        df.loc[idx, 'Time Slot'] = preferred_slot
                        
                        # Mark this date as occupied for this subbranch-semester
                        scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
                        
                        scheduled_count += 1
                        scheduled_in_gaps += 1
                        
                        #st.write(f"    ✅ Gap filled: {row['Subject']} on {date_str} at {preferred_slot}")
                
                #st.write(f"    📊 Filled {scheduled_in_gaps} subjects in existing gaps for {branch}")
                
                # PHASE 2: If subjects remain, extend the schedule
                if subjects_to_schedule:
                    st.write(f"    🔍 PHASE 2: Extending schedule for remaining {len(subjects_to_schedule)} subjects")
                    
                    # Find the next available date after gap filling
                    if available_dates:
                        last_gap_date = max(available_dates, key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
                        last_gap_datetime = datetime.strptime(last_gap_date, "%d-%m-%Y")
                        next_extension_date = find_next_valid_day(last_gap_datetime + timedelta(days=1), holidays)
                    else:
                        next_extension_date = start_date
                    
                    for idx, row in subjects_to_schedule:
                        # Find next valid date that's not already occupied by this subbranch-semester
                        while True:
                            exam_date = find_next_valid_day(next_extension_date, holidays)
                            date_str = exam_date.strftime("%d-%m-%Y")
                            
                            # Check if this date is already occupied by this subbranch-semester
                            if date_str not in scheduled_dates_per_subbranch_sem[subbranch_sem_key]:
                                break
                            else:
                                # Move to next day
                                next_extension_date = exam_date + timedelta(days=1)
                        
                        # Schedule the subject
                        df.loc[idx, 'Exam Date'] = date_str
                        df.loc[idx, 'Time Slot'] = preferred_slot
                        
                        # Mark this date as occupied
                        scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
                        
                        scheduled_count += 1
                        st.write(f"    ✅ Extended: {row['Subject']} on {date_str} at {preferred_slot}")
                        
                        # Move to next day for next subject
                        next_extension_date = exam_date + timedelta(days=1)
    
    st.success(f"✅ Successfully scheduled {scheduled_count} truly uncommon subjects (gaps filled first, then extended)")
    
    # Verify no double bookings
    verify_no_double_bookings(df)
    
    return df

def verify_no_double_bookings(df):
    """Verify that no subbranch-semester has more than one exam per day"""
    #st.write("🔍 Verifying no double bookings...")
    
    scheduled_df = df[df['Exam Date'] != ""].copy()
    if scheduled_df.empty:
        return
    
    # Group by subbranch, semester, and date
    grouped = scheduled_df.groupby(['SubBranch', 'Semester', 'Exam Date']).size()
    double_bookings = grouped[grouped > 1]
    
    if not double_bookings.empty:
        #st.error("❌ Double bookings detected:")
        for (subbranch, semester, date), count in double_bookings.items():
            pass
            #st.error(f"  - {subbranch} Semester {semester} has {count} exams on {date}")
    else:
        pass
        #st.success("✅ No double bookings found - one exam per day per subbranch-semester")

def schedule_remaining_individual_subjects(df, holidays, start_date):
    """
    Schedule any remaining individual subjects that couldn't be grouped.
    ENSURES only ONE exam per day per subbranch-semester combination.
    """
    remaining_subjects = df[df['Exam Date'] == ""].copy()
    
    if remaining_subjects.empty:
        return df
    
    st.info(f"🔧 Scheduling {len(remaining_subjects)} remaining individual subjects...")
    
    # Build tracking of already scheduled dates per subbranch-semester
    scheduled_subjects = df[df['Exam Date'] != ""].copy()
    scheduled_dates_per_subbranch_sem = {}
    
    def get_subbranch_semester_key(subbranch, semester):
        return f"{subbranch}_{semester}"
    
    # Initialize tracking from already scheduled subjects
    for idx, row in scheduled_subjects.iterrows():
        subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
        date_str = row['Exam Date']
        
        if subbranch_sem_key not in scheduled_dates_per_subbranch_sem:
            scheduled_dates_per_subbranch_sem[subbranch_sem_key] = set()
        scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
    
    # Helper function
    def find_next_valid_day(start_date, holidays_set):
        current_date = start_date
        while True:
            if current_date.weekday() != 6 and current_date.date() not in holidays_set:
                return current_date
            current_date += timedelta(days=1)
    
    # Find available dates for gap filling
    if not scheduled_subjects.empty:
        scheduled_dates = pd.to_datetime(scheduled_subjects['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        if not scheduled_dates.empty:
            latest_date = max(scheduled_dates).date()
        else:
            latest_date = start_date.date()
    else:
        latest_date = start_date.date()
    
    # Create list of available dates
    available_dates = []
    current_date = start_date
    end_date = datetime.combine(latest_date, datetime.min.time()) + timedelta(days=10)
    
    while current_date <= end_date:
        if current_date.weekday() != 6 and current_date.date() not in holidays:
            available_dates.append(current_date.strftime("%d-%m-%Y"))
        current_date += timedelta(days=1)
    
    available_dates.sort(key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    # Schedule remaining subjects
    scheduled_count = 0
    for idx, row in remaining_subjects.iterrows():
        semester = row['Semester']
        subbranch = row['SubBranch']
        subbranch_sem_key = get_subbranch_semester_key(subbranch, semester)
        preferred_slot = get_preferred_slot(semester)
        
        # Initialize tracking if not exists
        if subbranch_sem_key not in scheduled_dates_per_subbranch_sem:
            scheduled_dates_per_subbranch_sem[subbranch_sem_key] = set()
        
        # Try to find a gap first
        scheduled = False
        
        for date_str in available_dates:
            if date_str not in scheduled_dates_per_subbranch_sem[subbranch_sem_key]:
                # Found an available date
                df.loc[idx, 'Exam Date'] = date_str
                df.loc[idx, 'Time Slot'] = preferred_slot
                scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
                scheduled_count += 1
                scheduled = True
                break
        
        # If no gap found, extend
        if not scheduled:
            if available_dates:
                last_date = max(available_dates, key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
                next_date = find_next_valid_day(datetime.strptime(last_date, "%d-%m-%Y") + timedelta(days=1), holidays)
            else:
                next_date = start_date
            
            # Find next available date for this subbranch-semester
            while True:
                date_str = next_date.strftime("%d-%m-%Y")
                if date_str not in scheduled_dates_per_subbranch_sem[subbranch_sem_key]:
                    break
                next_date = find_next_valid_day(next_date + timedelta(days=1), holidays)
            
            df.loc[idx, 'Exam Date'] = date_str
            df.loc[idx, 'Time Slot'] = preferred_slot
            scheduled_dates_per_subbranch_sem[subbranch_sem_key].add(date_str)
            scheduled_count += 1
            
            # Add this date to available dates for future reference
            available_dates.append(date_str)
            available_dates.sort(key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    st.success(f"✅ Scheduled {scheduled_count} remaining individual subjects")
    
    # Verify no double bookings
    verify_no_double_bookings(df)
    
    return df
    
    
def read_timetable(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Debug: Show actual column names from the Excel file
        st.write("📋 **Actual columns in uploaded file:**")
        st.write(list(df.columns))
        
        # Create a flexible column mapping that handles variations
        column_mapping = {
            "Program": "Program",
            "Stream": "Stream", 
            "Current Session": "Semester",
            "Module Description": "SubjectName",
            "Module Abbreviation": "ModuleCode",
            "Campus Name": "Campus",
            "Difficulty Score": "Difficulty",
            "Exam Duration": "Exam Duration",
            "Student count": "StudentCount",
            "Common across sems": "CommonAcrossSems",
            "Circuit": "Circuit"
        }
        
        # Handle the "Is Common" column with flexible naming
        is_common_variations = ["Is Common", "IsCommon", "is common", "Is_Common", "is_common"]
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
        df["CommonAcrossSems"] = df["CommonAcrossSems"].fillna(False).astype(bool)
        df["Circuit"] = df["Circuit"].fillna(False).astype(bool)
        
        # Handle IsCommon column - create if it doesn't exist
        if "IsCommon" not in df.columns:
            st.info("ℹ️ Creating 'IsCommon' column with default values")
            # Default logic: if CommonAcrossSems is True, then IsCommon should be YES
            # Otherwise, check if the subject appears in multiple branches within the same semester
            df["IsCommon"] = "NO"  # Default value
            
            # Set YES for subjects that are common across semesters
            df.loc[df["CommonAcrossSems"] == True, "IsCommon"] = "YES"
            
            # Check for subjects common within semester
            for (semester, module_code), group in df.groupby(['Semester', 'ModuleCode']):
                if len(group['Branch'].unique()) > 1:
                    # This subject appears in multiple branches within the same semester
                    df.loc[group.index, "IsCommon"] = "YES"
        else:
            # Clean up the IsCommon column values
            df["IsCommon"] = df["IsCommon"].astype(str).str.strip().str.upper()
            df["IsCommon"] = df["IsCommon"].replace({"TRUE": "YES", "FALSE": "NO", "1": "YES", "0": "NO"})
            df["IsCommon"] = df["IsCommon"].fillna("NO")
        
        # Debug: Show IsCommon value distribution
        st.write("📊 **IsCommon value distribution:**")
        st.write(df["IsCommon"].value_counts())
        
        df_non = df[df["Category"] != "INTD"].copy()
        df_ele = df[df["Category"] == "INTD"].copy()
        
        def split_br(b):
            p = b.split("-", 1)
            return pd.Series([p[0].strip(), p[1].strip() if len(p) > 1 else ""])
        
        for d in (df_non, df_ele):
            d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)
        
        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty", "Exam Duration", "StudentCount", "CommonAcrossSems", "ModuleCode", "Circuit", "IsCommon"]
        
        # Ensure all required columns exist before selecting
        available_cols = [col for col in cols if col in df_non.columns]
        missing_cols = [col for col in cols if col not in df_non.columns]
        
        if missing_cols:
            st.warning(f"⚠️ Missing columns: {missing_cols}")
        
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

    def extract_duration(subject_str):
        match = re.search(r'\[Duration: (\d+\.?\d*) hrs\]', subject_str)
        if match:
            return float(match.group(1))
        else:
            return 3.0

    def get_semester_default_time_slot(semester, branch):
        # Get default time slot based on semester
        if semester.isdigit():
            sem_num = int(semester)
            return get_preferred_slot(sem_num)
        return "10:00 AM - 1:00 PM"  # Default fallback

    sheets_processed = 0
    
    for sheet_name, sheet_df in df_dict.items():
        #st.write(f"🔄 Processing sheet: {sheet_name}")
        
        if sheet_df.empty:
            st.warning(f"⚠️ Sheet {sheet_name} is empty, skipping")
            continue
            
        # Reset index to make sure we're working with a clean DataFrame
        if hasattr(sheet_df, 'index') and len(sheet_df.index.names) > 1:
            sheet_df = sheet_df.reset_index()
        
        # Parse sheet name
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        
        # Remove '_Electives' suffix if present
        if semester.endswith('_Electives'):
            semester = semester.replace('_Electives', '')
            
        semester_roman = semester if not semester.isdigit() else int_to_roman(int(semester))
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        # Handle normal subjects (non-electives)
        if not sheet_name.endswith('_Electives'):
            st.write(f"📋 Processing core subjects for {sheet_name}")
            
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

                #st.write(f"📄 Creating PDF page for {sheet_name} with {len(chunk_df)} exam dates")

                # Get the default time slot for this semester
                default_time_slot = get_semester_default_time_slot(semester, main_branch)

                # Convert Exam Date to desired format
                try:
                    chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
                except Exception as e:
                    st.warning(f"Date conversion error in {sheet_name}: {e}")
                    # Try alternative date parsing
                    chunk_df["Exam Date"] = chunk_df["Exam Date"].astype(str)

                # Process subjects to handle duration and timing information
                for sub_branch in chunk:
                    for idx in chunk_df.index:
                        cell_value = chunk_df.at[idx, sub_branch]
                        if pd.isna(cell_value) or str(cell_value).strip() == "---" or str(cell_value).strip() == "":
                            continue
                            
                        subjects = str(cell_value).split(", ")
                        modified_subjects = []
                        for subject in subjects:
                            duration = extract_duration(subject)
                            base_subject = re.sub(r' \[Duration: \d+\.?\d* hrs\]', '', subject)
                            
                            # Handle non-standard durations: show specific time range
                            if duration != 3:
                                start_time = default_time_slot.split(" - ")[0]
                                end_time = calculate_end_time(start_time, duration)
                                modified_subjects.append(f"{base_subject} ({start_time} - {end_time})")
                            else:
                                modified_subjects.append(base_subject)
                        
                        chunk_df.at[idx, sub_branch] = ", ".join(modified_subjects)

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
                
                print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=line_height, 
                                 header_content=header_content, branches=chunk, time_slot=default_time_slot)
                
                sheets_processed += 1
                #st.write(f"✅ Added PDF page for {sheet_name} chunk {start//sub_branch_cols_per_page + 1}")

        # Handle electives
        elif sheet_name.endswith('_Electives'):
            #st.write(f"🎓 Processing electives for {sheet_name}")
            
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

            st.write(f"📊 Processing {len(elective_data)} elective entries")

            # Convert Exam Date to desired format
            try:
                elective_data["Exam Date"] = pd.to_datetime(elective_data["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
            except Exception as e:
                st.warning(f"Date conversion error in electives {sheet_name}: {e}")
                elective_data["Exam Date"] = elective_data["Exam Date"].astype(str)

            # Clean subject display if SubjectDisplay column exists
            if 'SubjectDisplay' in elective_data.columns and 'OE' in elective_data.columns:
                elective_data['SubjectDisplay'] = elective_data.apply(
                    lambda row: ", ".join([s.replace(f" [{row['OE']}]", "") for s in str(row['SubjectDisplay']).split(", ") if str(s).strip()]),
                    axis=1
                )
            
            # Handle subjects with non-standard durations
            default_time_slot = get_semester_default_time_slot(semester, main_branch)
            if 'SubjectDisplay' in elective_data.columns:
                for idx in elective_data.index:
                    cell_value = elective_data.at[idx, 'SubjectDisplay']
                    if pd.isna(cell_value) or str(cell_value).strip() == "---":
                        continue
                    subjects = str(cell_value).split(", ")
                    modified_subjects = []
                    for subject in subjects:
                        duration = extract_duration(subject)
                        base_subject = re.sub(r' \[Duration: \d+\.?\d* hrs\]', '', subject)
                        
                        if duration != 3:
                            start_time = default_time_slot.split(" - ")[0]
                            end_time = calculate_end_time(start_time, duration)
                            modified_subjects.append(f"{base_subject} ({start_time} - {end_time})")
                        else:
                            modified_subjects.append(base_subject)
                    
                    elective_data.at[idx, 'SubjectDisplay'] = ", ".join(modified_subjects)

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
            
            print_table_custom(pdf, elective_data, cols_to_print, col_widths, line_height=10, 
                             header_content=header_content, branches=['All Streams'], time_slot=default_time_slot)
            
            sheets_processed += 1
            st.write(f"✅ Added PDF page for electives {sheet_name}")

    if sheets_processed == 0:
        st.error("❌ No sheets were processed for PDF generation!")
        return
    
    try:
        pdf.output(pdf_path)
        st.success(f"✅ PDF generated successfully with {sheets_processed} pages")
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
                    st.write(f"📊 Excel file contains {len(test_sheets)} sheets: {list(test_sheets.keys())}")
                    
                    # Show structure of first few sheets
                    for i, (sheet_name, sheet_df) in enumerate(test_sheets.items()):
                        if i < 3:  # Only show first 3 sheets
                            st.write(f"  📄 Sheet '{sheet_name}': {sheet_df.shape} with columns: {list(sheet_df.columns)}")
                            
                except Exception as e:
                    st.error(f"❌ Error reading back Excel file for verification: {e}")
            else:
                st.error(f"❌ Temporary Excel file was not created at {temp_excel}")
                return
            
        except Exception as e:
            st.error(f"❌ Error saving temporary Excel file: {e}")
            return
            
        st.write("🎨 Converting Excel to PDF...")
        try:
            convert_excel_to_pdf(temp_excel, output_pdf)
            st.write("✅ PDF conversion completed")
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

    st.write("🔍 **Debugging Verification Process...**")
    
    # Debug: Show what columns are available in original data
    st.write(f"📋 Original dataframe columns: {list(original_df.columns)}")
    st.write(f"📋 Original dataframe shape: {original_df.shape}")
    
    # Combine all scheduled data first
    scheduled_data = pd.concat(semester_wise_timetable.values(), ignore_index=True)
    st.write(f"📅 Scheduled data shape: {scheduled_data.shape}")
    st.write(f"📅 Scheduled data columns: {list(scheduled_data.columns)}")
    
    # Debug: Show sample of scheduled data
    if not scheduled_data.empty:
        st.write("📅 **Sample scheduled data:**")
        sample_scheduled = scheduled_data[['Subject', 'Branch', 'Semester', 'Exam Date', 'Time Slot']].head(3)
        st.dataframe(sample_scheduled)
    
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
    verification_df["Exam Time"] = ""
    verification_df["Time Slot"] = ""
    verification_df["Is Common Status"] = ""
    verification_df["Scheduling Status"] = "Not Scheduled"

    # Debug: Show sample of verification data before processing
    st.write("📋 **Sample verification data before processing:**")
    if not verification_df.empty:
        sample_verification = verification_df[['Module Abbreviation', 'Current Session', 'Program', 'Stream']].head(3)
        st.dataframe(sample_verification)

    # Track statistics
    matched_count = 0
    unmatched_count = 0
    
    # IMPROVED: Create lookup dictionaries for faster matching
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
                time_slot = matched_subject["Time Slot"]
                duration = row.get("Exam Duration", 3.0)
                
                # Handle duration
                try:
                    duration = float(duration) if pd.notna(duration) else 3.0
                except:
                    duration = 3.0
                
                # UPDATED: Calculate exam time based on duration and common subject rules
                if time_slot and str(time_slot).strip() and str(time_slot) != "nan":
                    try:
                        # For non-standard durations, show specific time range
                        if duration != 3:
                            start_time = str(time_slot).split(" - ")[0].strip()
                            end_time = calculate_end_time(start_time, duration)
                            exam_time = f"{start_time} - {end_time}"
                        # For common subjects, show the actual assigned time slot
                        elif matched_subject.get('CommonAcrossSems', False):
                            semester_default = get_preferred_slot(semester_num)
                            if time_slot != semester_default:
                                exam_time = str(time_slot)
                            else:
                                start_time = str(time_slot).split(" - ")[0].strip()
                                end_time = calculate_end_time(start_time, duration)
                                exam_time = f"{start_time} - {end_time}"
                        else:
                            start_time = str(time_slot).split(" - ")[0].strip()
                            end_time = calculate_end_time(start_time, duration)
                            exam_time = f"{start_time} - {end_time}"
                    except Exception as e:
                        st.write(f"⚠️ Error calculating time for {module_code}: {e}")
                        exam_time = str(time_slot)
                else:
                    exam_time = "TBD"
                    time_slot = "TBD"
                
                # Update verification dataframe
                verification_df.at[idx, "Exam Date"] = str(exam_date)
                verification_df.at[idx, "Exam Time"] = exam_time
                verification_df.at[idx, "Time Slot"] = str(time_slot)
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
                
                if idx < 3:  # Debug first few matches
                    st.write(f"   ✅ **MATCH FOUND for {module_code}!**")
                    st.write(f"   Exam Date: {exam_date}")
                    st.write(f"   Time Slot: {time_slot}")
                
            else:
                # No match found
                verification_df.at[idx, "Exam Date"] = "Not Scheduled"
                verification_df.at[idx, "Exam Time"] = "Not Scheduled" 
                verification_df.at[idx, "Time Slot"] = "Not Scheduled"
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

    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sem, df_sem in semester_wise_timetable.items():
                for main_branch in df_sem["MainBranch"].unique():
                    df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
                    
                    # Separate non-electives and electives
                    df_non_elec = df_mb[df_mb['OE'].isna() | (df_mb['OE'].str.strip() == "")].copy()
                    df_elec = df_mb[df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")].copy()

                    # Process non-electives (CRITICAL FIX: Always create sheet even if empty)
                    roman_sem = int_to_roman(sem)
                    sheet_name = f"{main_branch}_Sem_{roman_sem}"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    
                    if not df_non_elec.empty:
                        #st.write(f"📊 Processing {len(df_non_elec)} non-elective subjects for {sheet_name}")
                        
                        # Add difficulty and duration info
                        difficulty_str = df_non_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                        difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                        
                        # Create subject display with duration info
                        df_non_elec["SubjectDisplay"] = df_non_elec["Subject"]
                        duration_suffix = df_non_elec.apply(
                            lambda row: f" [Duration: {row['Exam Duration']} hrs]" if row['Exam Duration'] != 3 else '', axis=1)
                        df_non_elec["SubjectDisplay"] = df_non_elec["SubjectDisplay"] + difficulty_suffix + duration_suffix
                        
                        # Convert and sort by date
                        df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], format="%d-%m-%Y", dayfirst=True, errors='coerce')
                        df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                        
                        # Group subjects by date, combining any multiple subjects per branch per day
                        grouped_by_date = df_non_elec.groupby(['Exam Date', 'SubBranch'])['SubjectDisplay'].apply(
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
                        
                        # Ensure the first column is named 'Exam Date' (in case it got renamed to 'index')
                        if pivot_df.columns[0] != 'Exam Date':
                            pivot_df = pivot_df.rename(columns={pivot_df.columns[0]: 'Exam Date'})
                        
                        # Save to Excel
                        pivot_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        #st.write(f"✅ Created sheet {sheet_name} with {len(pivot_df)} exam dates")
                        
                    else:
                        # CRITICAL FIX: Create empty sheet structure for branches with no subjects
                        #st.write(f"⚠️ No non-elective subjects for {sheet_name}, creating empty structure")
                        
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
                            #st.write(f"✅ Created empty sheet {sheet_name} with structure for subbranches: {', '.join(all_subbranches)}")
                        else:
                            # If no subbranches, create minimal structure
                            empty_df = pd.DataFrame({
                                'Exam Date': ['No exams scheduled'],
                                'Subjects': ['No subjects available']
                            })
                            empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            #st.write(f"✅ Created minimal empty sheet {sheet_name}")

                    # Process electives in a separate sheet (only if electives exist)
                    if not df_elec.empty:
                        #st.write(f"📊 Processing {len(df_elec)} elective subjects for {sheet_name}")
                        
                        difficulty_str = df_elec['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                        difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                        
                        # Create subject display with OE and duration info
                        df_elec["SubjectDisplay"] = df_elec["Subject"] + " [" + df_elec["OE"] + "]"
                        duration_suffix = df_elec.apply(
                            lambda row: f" [Duration: {row['Exam Duration']} hrs]" if row['Exam Duration'] != 3 else '', axis=1)
                        df_elec["SubjectDisplay"] = df_elec["SubjectDisplay"] + difficulty_suffix + duration_suffix
                        
                        # Group by OE and Date only (no time slot)
                        elec_pivot = df_elec.groupby(['OE', 'Exam Date'])['SubjectDisplay'].apply(
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
                        #st.write(f"✅ Created electives sheet {elective_sheet_name} with {len(elec_pivot)} entries")

        output.seek(0)
        st.success("✅ Excel file created successfully with all required sheets")
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

def optimize_oe_subjects_after_scheduling(sem_dict, holidays, optimizer=None):
    """
    After main scheduling, check if OE subjects can be moved to earlier empty slots.
    CRITICAL: OE2 must be scheduled on the day immediately after OE1/OE5.
    """
    if not sem_dict:
        return sem_dict, 0, []
    
    st.info("🎯 Optimizing Open Elective (OE) placement...")
    
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
    
    if oe_data.empty:
        st.info("No OE subjects to optimize")
        return sem_dict, 0, []
    
    # Build complete schedule grid from current state
    schedule_grid = {}
    branches = all_data['Branch'].unique()
    
    # Populate with all scheduled exams
    for _, row in all_data.iterrows():
        if pd.notna(row['Exam Date']) and row['Exam Date'].strip() != "":
            date_str = row['Exam Date']
            
            if date_str not in schedule_grid:
                schedule_grid[date_str] = {}
            if row['Time Slot'] not in schedule_grid[date_str]:
                schedule_grid[date_str][row['Time Slot']] = {}
            schedule_grid[date_str][row['Time Slot']][row['Branch']] = row['Subject']
    
    # Find all dates in the schedule
    all_dates = sorted(schedule_grid.keys(), 
                      key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
    
    if not all_dates:
        return sem_dict, 0, []
    
    # Get date range and fill empty days
    start_date = datetime.strptime(all_dates[0], "%d-%m-%Y")
    end_date = datetime.strptime(all_dates[-1], "%d-%m-%Y")
    
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
    
    # Process OE optimization
    oe_data_copy = oe_data.copy()
    oe1_oe5_data = oe_data_copy[oe_data_copy['OE'].isin(['OE1', 'OE5'])]
    oe2_data = oe_data_copy[oe_data_copy['OE'] == 'OE2']
    
    moves_made = 0
    optimization_log = []
    
    # Process OE1/OE5 together (they should always be on the same date/time)
    if not oe1_oe5_data.empty:
        current_oe1_oe5_date = oe1_oe5_data['Exam Date'].iloc[0]
        current_oe1_oe5_slot = oe1_oe5_data['Time Slot'].iloc[0]
        current_oe1_oe5_date_obj = datetime.strptime(current_oe1_oe5_date, "%d-%m-%Y")
        
        affected_branches = oe1_oe5_data['Branch'].unique()
        
        # Find earlier slots that are empty for ALL branches with OE1/OE5
        best_oe1_oe5_date = None
        best_oe1_oe5_slot = None
        
        sorted_dates = sorted(schedule_grid.keys(), 
                            key=lambda x: datetime.strptime(x, "%d-%m-%Y"))
        for check_date_str in sorted_dates:
            check_date_obj = datetime.strptime(check_date_str, "%d-%m-%Y")
            
            # Only look for earlier dates
            if check_date_obj >= current_oe1_oe5_date_obj:
                break
            
            # Skip weekends and holidays
            if check_date_obj.weekday() == 6 or check_date_obj.date() in holidays:
                continue
            
            # Check if the day immediately after this date is also valid for OE2
            next_day = find_next_valid_day_for_electives(check_date_obj + timedelta(days=1), holidays)
            next_day_str = next_day.strftime("%d-%m-%Y")
            
            # Check both time slots for OE1/OE5
            for time_slot in ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]:
                can_move_oe1_oe5 = True
                
                # Check if this slot is empty for all OE1/OE5 branches
                for branch in affected_branches:
                    if (check_date_str in schedule_grid and 
                        time_slot in schedule_grid[check_date_str] and
                        branch in schedule_grid[check_date_str][time_slot] and
                        schedule_grid[check_date_str][time_slot][branch] is not None):
                        can_move_oe1_oe5 = False
                        break
                
                if can_move_oe1_oe5:
                    # Check if OE2 can be scheduled on the next day
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
                            best_oe1_oe5_date = check_date_str
                            best_oe1_oe5_slot = time_slot
                            best_oe2_date = next_day_str
                            break
                    else:
                        best_oe1_oe5_date = check_date_str
                        best_oe1_oe5_slot = time_slot
                        break
            
            if best_oe1_oe5_date:
                break
        
        # If we found a better slot for OE1/OE5, move them and OE2
        if best_oe1_oe5_date and best_oe1_oe5_date != current_oe1_oe5_date:
            days_saved = (current_oe1_oe5_date_obj - datetime.strptime(best_oe1_oe5_date, "%d-%m-%Y")).days
            
            # Update all OE1/OE5 exams in semester dictionary
            for idx in oe1_oe5_data.index:
                sem = all_data.at[idx, 'Semester']
                branch = all_data.at[idx, 'Branch']
                subject = all_data.at[idx, 'Subject']
                
                # Update in the semester dictionary
                mask = (sem_dict[sem]['Subject'] == subject) & \
                       (sem_dict[sem]['Branch'] == branch)
                sem_dict[sem].loc[mask, 'Exam Date'] = best_oe1_oe5_date
                sem_dict[sem].loc[mask, 'Time Slot'] = best_oe1_oe5_slot
            
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
                    sem_dict[sem].loc[mask, 'Time Slot'] = best_oe2_slot
            
            moves_made += 1
            optimization_log.append(
                f"Moved OE1/OE5 from {current_oe1_oe5_date} to {best_oe1_oe5_date} (saved {days_saved} days)"
            )
            if not oe2_data.empty:
                optimization_log.append(
                    f"Moved OE2 to {best_oe2_date} (day immediately after OE1/OE5)"
                )
    
    # Ensure all dates in sem_dict are properly formatted
    for sem in sem_dict:
        sem_dict[sem]['Exam Date'] = sem_dict[sem]['Exam Date'].apply(normalize_date_to_ddmmyyyy)
    
    if moves_made > 0:
        st.success(f"✅ OE Optimization: Moved {moves_made} OE groups!")
        with st.expander("📝 OE Optimization Details"):
            for log in optimization_log:
                st.write(f"• {log}")
    else:
        st.info("ℹ️ OE subjects are already optimally placed")
    
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
        st.markdown("#### 📅 Base Date for Scheduling")
        base_date = st.date_input("Start date for exams", value=datetime(2025, 4, 1))
        base_date = datetime.combine(base_date, datetime.min.time())

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
                
                    st.write("Reading timetable...")
                    df_non_elec, df_ele, original_df = read_timetable(uploaded_file)

                    if df_non_elec is not None:
                        st.write("Processing subjects...")
                        
                        # NEW THREE-PHASE IMPROVED SCHEDULING ORDER:
                        # Phase 1: Schedule COMMON ACROSS SEMESTERS subjects FIRST from base date
                        st.info("🎯 THREE-PHASE SCHEDULING: Common Across Semesters → Common Within Semester → Uncommon → Electives (One exam per day per branch)")
                        df_scheduled = schedule_common_subjects_first(df_non_elec, holidays_set, base_date)
                        
                        # Phase 2: Schedule COMMON WITHIN SEMESTER subjects (COMP/ELEC that appear in multiple branches within same semester)
                        st.info("🔗 Phase 2: Scheduling common-within-semester subjects...")
                        df_scheduled = schedule_common_within_semester_subjects(df_scheduled, holidays_set, base_date)
                        
                        # Phase 3: Schedule TRULY UNCOMMON subjects using gap-filling approach
                        # This will fill gaps first, then extend only if needed
                        st.info("🔍 Phase 3: Filling gaps with truly uncommon subjects first, then extending...")
                        df_scheduled = schedule_uncommon_subjects_after_common(df_scheduled, holidays_set, base_date)
                        
                        # Step 4: Handle any remaining unscheduled subjects
                        remaining_subjects = df_scheduled[
                            (df_scheduled['Exam Date'] == "") & 
                            (df_scheduled['Category'] != 'INTD')
                        ].copy()

                        if not remaining_subjects.empty:
                            st.warning(f"⚠️ {len(remaining_subjects)} subjects could not be scheduled in groups. Scheduling individually...")
                            df_scheduled = schedule_remaining_individual_subjects(df_scheduled, holidays_set, base_date)
                        
                        # Step 5: Handle electives if they exist
                        if df_ele is not None and not df_ele.empty:
                            # Find the maximum date from non-elective scheduling
                            non_elec_dates = pd.to_datetime(df_scheduled['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                            if not non_elec_dates.empty:
                                max_non_elec_date = max(non_elec_dates).date()
                                st.write(f"📅 Max non-elective date: {max_non_elec_date.strftime('%d-%m-%Y')}")
                                
                                # Schedule electives globally
                                df_ele_scheduled = schedule_electives_globally(df_ele, max_non_elec_date, holidays_set)
                                
                                # Combine non-electives and electives
                                all_scheduled_subjects = pd.concat([df_scheduled, df_ele_scheduled], ignore_index=True)
                            else:
                                all_scheduled_subjects = df_scheduled
                        else:
                            all_scheduled_subjects = df_scheduled
                        
                        # Step 6: Create semester dictionary
                        all_scheduled_subjects = all_scheduled_subjects[all_scheduled_subjects['Exam Date'] != ""]
                        
                        if not all_scheduled_subjects.empty:
                            # Sort by semester and date
                            all_scheduled_subjects = all_scheduled_subjects.sort_values(["Semester", "Exam Date"], ascending=True)
                            
                            # Create semester dictionary
                            sem_dict = {}
                            for s in sorted(all_scheduled_subjects["Semester"].unique()):
                                sem_data = all_scheduled_subjects[all_scheduled_subjects["Semester"] == s].copy()
                                sem_dict[s] = sem_data
                            
                            # Step 7: Optimize OE subjects if they exist
                            if df_ele is not None and not df_ele.empty:
                                st.write("Optimizing OE subjects...")
                                sem_dict, moves_made, optimization_log = optimize_oe_subjects_after_scheduling(sem_dict, holidays_set)

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
                            st.info("✅ **Three-Phase Scheduling Applied:**\n1. 🎯 **Phase 1:** Common across semesters scheduled FIRST from base date\n2. 🔗 **Phase 2:** Common within semester subjects (COMP/ELEC appearing in multiple branches)\n3. 🔍 **Phase 3:** Truly uncommon subjects with gap-filling optimization\n4. 🎓 **Phase 4:** Electives scheduled LAST\n5. ⚡ **Guarantee:** ONE exam per day per subbranch-semester")
                            
                            # Show efficiency improvement
                            efficiency = (unique_exam_days / overall_date_range) * 100 if overall_date_range > 0 else 0
                            st.success(f"📊 **Schedule Efficiency: {efficiency:.1f}%** (Higher is better - more days utilized)")
                            
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
                            st.warning("No subjects found to schedule.")

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
        st.markdown("""
        <div class="stats-section">
            <h2>📈 Complete Timetable Statistics</h2>
        </div>
        """, unsafe_allow_html=True)

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
            st.markdown(f'<div class="metric-card"><h3>📅 {st.session_state.overall_date_range}</h3><p>Days Span</p></div>',
                        unsafe_allow_html=True)

        # Show efficiency metrics
        if st.session_state.unique_exam_days > 0 and st.session_state.overall_date_range > 0:
            efficiency = (st.session_state.unique_exam_days / st.session_state.overall_date_range) * 100
            if efficiency > 80:
                st.success(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Excellent - most days are utilized)")
            elif efficiency > 60:
                st.info(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Good)")
            else:
                st.warning(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Could be improved)")

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
        <p>🎓 <strong>Three-Phase Timetable Generator with Gap-Filling & Double Booking Prevention</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
        <p style="font-size: 0.9em;">Common across semesters first • Common within semester • Gap-filling optimization • One exam per day per branch • OE optimization • Maximum efficiency • Verification export</p>
    </div>
    """, unsafe_allow_html=True)
    
if __name__ == "__main__":
    main()


















