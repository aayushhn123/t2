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

def schedule_uncommon_subjects_first(df, holidays, base_date):
    """
    Schedule uncommon subjects (CommonAcrossSems = FALSE) first, semester-wise for each stream
    Base date resets for each semester AND each sub-branch/stream
    """
    st.info("🔧 Scheduling uncommon subjects first...")
    
    # Filter uncommon subjects only
    uncommon_subjects = df[df['CommonAcrossSems'] == False].copy()
    
    if uncommon_subjects.empty:
        st.info("No uncommon subjects to schedule")
        return df
    
    st.write(f"Found {len(uncommon_subjects)} uncommon subjects to schedule")
    
    # Initialize exam_days tracker for each branch (this will track across all scheduling)
    all_branches = df['Branch'].unique()
    exam_days = {branch: set() for branch in all_branches}
    
    # Track all scheduled dates for span calculation
    all_scheduled_dates = []
    semester_statistics = {}
    
    # Helper function to find next valid day (skip Sundays and holidays)
    def find_next_valid_day(start_date, holidays_set):
        current_date = start_date
        while True:
            # Skip Sundays (weekday 6) and holidays
            if current_date.weekday() != 6 and current_date.date() not in holidays_set:
                return current_date
            current_date += timedelta(days=1)
    
    # Schedule by semester, then by stream within each semester
    scheduled_count = 0
    
    for semester in sorted(uncommon_subjects['Semester'].unique()):
        semester_data = uncommon_subjects[uncommon_subjects['Semester'] == semester]
        
        st.write(f"📚 Scheduling Semester {semester} uncommon subjects...")
        
        # Track dates for this semester
        semester_dates = []
        semester_branch_stats = {}
        
        # Get time slot based on semester
        if semester % 2 != 0:  # Odd semester
            odd_sem_position = (semester + 1) // 2
            preferred_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
        else:  # Even semester
            even_sem_position = semester // 2
            preferred_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
        
        # Group by stream/branch within this semester
        for branch in sorted(semester_data['Branch'].unique()):
            branch_subjects = semester_data[semester_data['Branch'] == branch]
            
            st.write(f"  🔧 Scheduling {len(branch_subjects)} subjects for {branch}")
            
            # RESET TO BASE DATE FOR EACH BRANCH/STREAM - KEY CHANGE!
            branch_scheduling_date = base_date
            branch_dates = []
            
            st.write(f"    📅 Starting {branch} from base date: {base_date.strftime('%d-%m-%Y')}")
            
            # Schedule each subject for this branch back-to-back
            for idx, row in branch_subjects.iterrows():
                # Find a valid day starting from base date for this branch
                exam_date = find_next_valid_day(branch_scheduling_date, holidays)
                
                # Since each branch starts independently, no need to check exam_days conflicts
                # Just find next valid calendar day
                
                # Schedule the exam
                date_str = exam_date.strftime("%d-%m-%Y")
                df.loc[idx, 'Exam Date'] = date_str
                df.loc[idx, 'Time Slot'] = preferred_slot
                
                # Mark this date as occupied for this branch (for future reference)
                exam_days[branch].add(exam_date.date())
                
                # Track the scheduled date
                all_scheduled_dates.append(exam_date.date())
                semester_dates.append(exam_date.date())
                branch_dates.append(exam_date.date())
                
                scheduled_count += 1
                st.write(f"    ✅ Scheduled {row['Subject']} on {date_str}")
                
                # Move to next day for next subject of this branch
                branch_scheduling_date = find_next_valid_day(exam_date + timedelta(days=1), holidays)
            
            # Track branch statistics
            if branch_dates:
                branch_start = min(branch_dates)
                branch_end = max(branch_dates)
                branch_span = (branch_end - branch_start).days + 1
                branch_unique_days = len(set(branch_dates))
                
                semester_branch_stats[branch] = {
                    'count': len(branch_subjects),
                    'start': branch_start,
                    'end': branch_end,
                    'span': branch_span,
                    'unique_days': branch_unique_days
                }
                
                st.write(f"    📊 {branch} span: {branch_span} days ({branch_start.strftime('%d-%m-%Y')} to {branch_end.strftime('%d-%m-%Y')})")
                st.write(f"    📅 {branch} subjects scheduled on {branch_unique_days} days")
        
        # Store semester statistics
        if semester_dates:
            semester_start = min(semester_dates)
            semester_end = max(semester_dates)
            semester_span = (semester_end - semester_start).days + 1
            semester_unique_days = len(set(semester_dates))
            
            semester_statistics[semester] = {
                'count': len(semester_data),
                'start': semester_start,
                'end': semester_end,
                'span': semester_span,
                'unique_days': semester_unique_days,
                'branches': semester_branch_stats
            }
            
            st.write(f"  📊 Semester {semester} overall span: {semester_span} days ({semester_start.strftime('%d-%m-%Y')} to {semester_end.strftime('%d-%m-%Y')})")
            st.write(f"  📅 Semester {semester} unique exam days: {semester_unique_days}")
        
        st.write(f"✅ Completed Semester {semester} scheduling")
    
    # Calculate and display overall span for uncommon subjects
    if all_scheduled_dates:
        overall_start = min(all_scheduled_dates)
        overall_end = max(all_scheduled_dates)
        overall_span = (overall_end - overall_start).days + 1
        unique_exam_days = len(set(all_scheduled_dates))
        total_exam_dates = len(all_scheduled_dates)
        
        st.success(f"✅ Successfully scheduled {scheduled_count} uncommon subjects")
        
        # Display comprehensive statistics
        st.markdown("### 📊 Uncommon Subjects Scheduling Statistics")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Subjects", scheduled_count)
        with col2:
            st.metric("Overall Span", f"{overall_span} days")
        with col3:
            st.metric("Unique Exam Days", unique_exam_days)
        with col4:
            st.metric("Total Streams", len(set(uncommon_subjects['Branch'])))
        
        st.info(f"📅 **Overall Date Range:** {overall_start.strftime('%d %B %Y')} to {overall_end.strftime('%d %B %Y')}")
        
        # Show efficiency metrics
        efficiency = (unique_exam_days / overall_span) * 100 if overall_span > 0 else 0
        if efficiency > 80:
            st.success(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Excellent - most days are utilized)")
        elif efficiency > 60:
            st.info(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Good)")
        else:
            st.warning(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Could be improved)")
        
        # Additional insights
        with st.expander("📈 Detailed Statistics"):
            st.write(f"• **Base date used for all streams:** {base_date.strftime('%d %B %Y')}")
            st.write(f"• **First exam scheduled on:** {overall_start.strftime('%d %B %Y')}")
            st.write(f"• **Last exam scheduled on:** {overall_end.strftime('%d %B %Y')}")
            st.write(f"• **Days with no exams:** {overall_span - unique_exam_days}")
            st.write(f"• **Average exams per active day:** {total_exam_dates / unique_exam_days:.1f}")
            
            # Show detailed semester and branch breakdown
            st.write("\n**📚 Semester & Stream-wise Breakdown:**")
            for semester in sorted(semester_statistics.keys()):
                stats = semester_statistics[semester]
                st.write(f"\n**Semester {semester}:** {stats['count']} subjects, {stats['span']} days span")
                
                for branch, branch_stats in stats['branches'].items():
                    st.write(f"  • **{branch}:** {branch_stats['count']} subjects")
                    st.write(f"    - Date range: {branch_stats['start'].strftime('%d-%m-%Y')} to {branch_stats['end'].strftime('%d-%m-%Y')}")
                    st.write(f"    - Span: {branch_stats['span']} days, Active days: {branch_stats['unique_days']}")
        
        # Show stream independence verification
        with st.expander("🔍 Stream Independence Verification"):
            st.write("**Verification that each stream started from base date:**")
            for semester in sorted(semester_statistics.keys()):
                st.write(f"\n**Semester {semester}:**")
                for branch, branch_stats in semester_statistics[semester]['branches'].items():
                    days_from_base = (branch_stats['start'] - base_date.date()).days
                    st.write(f"  • {branch}: Started {days_from_base} days from base date")
                    if days_from_base == 0:
                        st.write(f"    ✅ Started exactly on base date")
                    else:
                        st.write(f"    ℹ️ Started {days_from_base} days after base date (due to holidays/weekends)")
    else:
        st.success(f"✅ Successfully scheduled {scheduled_count} uncommon subjects")
    
    return df


def schedule_common_subjects_after_uncommon(df, holidays, base_date):
    """
    Schedule common subjects across ALL branches (circuit and non-circuit) and remaining individual subjects
    Priority: Schedule maximum common subjects per day while respecting constraints
    CRITICAL CONSTRAINT: Only one exam per day per subbranch per semester (not just branch per semester)
    FIXED: Properly handle all common subjects including those that couldn't be scheduled initially
    """
    st.info("🔧 Scheduling common subjects across all branches with maximum subjects per day priority...")
    
    # Filter subjects that need to be scheduled (all remaining subjects except electives)
    subjects_to_schedule = df[
        (df['Exam Date'] == "") &  # Not yet scheduled
        (df['Category'] != 'INTD')  # Not electives
    ].copy()
    
    if subjects_to_schedule.empty:
        st.info("No subjects to schedule")
        return df
    
    st.write(f"Found {len(subjects_to_schedule)} subjects to schedule")
    
    # Separate common and non-common subjects (across ALL branches now)
    common_subjects = subjects_to_schedule[subjects_to_schedule['CommonAcrossSems'] == True].copy()
    individual_subjects = subjects_to_schedule[subjects_to_schedule['CommonAcrossSems'] == False].copy()
    
    st.write(f"📋 Common subjects to schedule: {len(common_subjects)}")
    st.write(f"📋 Individual subjects to schedule: {len(individual_subjects)}")
    
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
    
    # Schedule common subjects with maximum subjects per day priority
    if not common_subjects.empty:
        st.write("📚 Scheduling common subjects across all branches...")
        
        # Group by ModuleCode to find truly common subjects across ALL branches
        common_subject_groups = {}
        for module_code, group in common_subjects.groupby('ModuleCode'):
            branches_for_this_subject = group['Branch'].unique()
            # FIXED: Accept subjects that are marked as common even if they appear in only one branch in this batch
            # This handles cases where other instances might exist in other semesters or have been filtered out
            common_subject_groups[module_code] = group
            st.write(f"  📋 Found common subject {group['Subject'].iloc[0]} across {len(branches_for_this_subject)} branches: {', '.join(branches_for_this_subject)}")
        
        # Create a list of unscheduled common subject groups
        unscheduled_groups = list(common_subject_groups.keys())
        st.write(f"📊 Total common subject groups to schedule: {len(unscheduled_groups)}")
        
        # Schedule day by day, maximizing subjects per day
        scheduling_attempts = 0
        max_scheduling_attempts = 200  # Prevent infinite loops
        
        while unscheduled_groups and scheduling_attempts < max_scheduling_attempts:
            scheduling_attempts += 1
            exam_date = find_next_valid_day(current_scheduling_date, holidays)
            date_str = exam_date.strftime("%d-%m-%Y")
            
            if date_str not in daily_scheduled:
                daily_scheduled[date_str] = set()
            
            # Available time slots for this date
            available_slots = ["10:00 AM - 1:00 PM", "2:00 PM - 5:00 PM"]
            subjects_scheduled_today = 0
            
            st.write(f"  📅 Attempting to schedule on {date_str} (Attempt {scheduling_attempts})")
            
            # Try to schedule as many common subjects as possible on this day
            day_had_scheduling = False
            
            for time_slot in available_slots:
                groups_scheduled_in_slot = []
                
                st.write(f"    🕐 Checking {time_slot} slot")
                
                # Check which groups can be scheduled in this time slot
                for module_code in unscheduled_groups[:]:  # Use slice to avoid modification during iteration
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
                        if semester % 2 != 0:  # Odd semester
                            odd_sem_position = (semester + 1) // 2
                            preferred_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
                        else:  # Even semester
                            even_sem_position = semester // 2
                            preferred_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
                        
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
                        
                        st.write(f"      ✅ Scheduled common subject {group['Subject'].iloc[0]} for subbranches: {', '.join(subbranches_scheduled)} on {date_str} at {actual_slot}")
                    else:
                        st.write(f"      ❌ Cannot schedule {group['Subject'].iloc[0]} - conflicts with subbranches: {', '.join(conflicting_subbranches)}")
                
                # Remove scheduled groups from unscheduled list
                for module_code in groups_scheduled_in_slot:
                    unscheduled_groups.remove(module_code)
                
                # If no more subjects can be scheduled in remaining slots for this time slot, continue to next slot
                if not groups_scheduled_in_slot:
                    st.write(f"      ℹ️ No more common subjects can be scheduled in {time_slot}")
            
            # Try to fill remaining slots with individual subjects
            remaining_slots = 2 - subjects_scheduled_today
            if remaining_slots > 0 and not individual_subjects.empty:
                st.write(f"    🔄 Trying to fill {remaining_slots} remaining slots with individual subjects")
                
                individual_scheduled_today = 0
                for time_slot in available_slots:
                    if individual_scheduled_today >= remaining_slots:
                        break
                    
                    # Find individual subjects that can be scheduled in this slot
                    available_individual = []
                    for idx, row in individual_subjects.iterrows():
                        if df.loc[idx, 'Exam Date'] == "":  # Not yet scheduled
                            subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                            if subbranch_sem_key not in daily_scheduled[date_str]:
                                available_individual.append(idx)
                    
                    if available_individual:
                        # Schedule one individual subject per available slot
                        for slot_fill in range(min(len(available_individual), remaining_slots - individual_scheduled_today)):
                            idx_to_schedule = available_individual[slot_fill]
                            row = individual_subjects.loc[idx_to_schedule]
                            
                            # Calculate preferred slot for this subject
                            semester = row['Semester']
                            if semester % 2 != 0:  # Odd semester
                                odd_sem_position = (semester + 1) // 2
                                preferred_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
                            else:  # Even semester
                                even_sem_position = semester // 2
                                preferred_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
                            
                            # Use the preferred slot if it matches current slot, otherwise use current slot
                            actual_slot = preferred_slot if preferred_slot == time_slot else time_slot
                            
                            # Schedule the individual subject
                            df.loc[idx_to_schedule, 'Exam Date'] = date_str
                            df.loc[idx_to_schedule, 'Time Slot'] = actual_slot
                            scheduled_count += 1
                            individual_scheduled_today += 1
                            
                            subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                            daily_scheduled[date_str].add(subbranch_sem_key)
                            
                            st.write(f"      ✅ Scheduled individual subject {row['Subject']} for {row['SubBranch']} (Sem {row['Semester']}) on {date_str} at {actual_slot}")
                
                # Update individual_subjects to remove scheduled ones
                individual_subjects = individual_subjects[individual_subjects['Exam Date'] == ""]
            
            # If no scheduling happened today, move to next day
            if not day_had_scheduling and subjects_scheduled_today == 0:
                current_scheduling_date = exam_date + timedelta(days=1)
            else:
                # Move to next day only if we scheduled something
                current_scheduling_date = exam_date + timedelta(days=1)
        
        if unscheduled_groups:
            st.warning(f"⚠️ {len(unscheduled_groups)} common subject groups could not be scheduled after {scheduling_attempts} attempts")
            for module_code in unscheduled_groups:
                group = common_subject_groups[module_code]
                st.write(f"  • Unscheduled: {group['Subject'].iloc[0]} for branches: {', '.join(group['Branch'].unique())}")
    
    # Schedule any remaining individual subjects
    if not individual_subjects.empty:
        st.write("📚 Scheduling remaining individual subjects...")
        
        for idx, row in individual_subjects.iterrows():
            if df.loc[idx, 'Exam Date'] != "":  # Skip if already scheduled
                continue
            
            # Find next available date for this subbranch-semester
            scheduling_date = current_scheduling_date
            scheduled = False
            
            while not scheduled:
                exam_date = find_next_valid_day(scheduling_date, holidays)
                date_str = exam_date.strftime("%d-%m-%Y")
                
                if date_str not in daily_scheduled:
                    daily_scheduled[date_str] = set()
                
                subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                if subbranch_sem_key not in daily_scheduled[date_str]:
                    # Schedule in preferred slot
                    semester = row['Semester']
                    if semester % 2 != 0:
                        odd_sem_position = (semester + 1) // 2
                        preferred_slot = "10:00 AM - 1:00 PM" if odd_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
                    else:
                        even_sem_position = semester // 2
                        preferred_slot = "10:00 AM - 1:00 PM" if even_sem_position % 2 == 1 else "2:00 PM - 5:00 PM"
                    
                    df.loc[idx, 'Exam Date'] = date_str
                    df.loc[idx, 'Time Slot'] = preferred_slot
                    daily_scheduled[date_str].add(subbranch_sem_key)
                    scheduled_count += 1
                    scheduled = True
                    
                    st.write(f"✅ Scheduled remaining individual subject {row['Subject']} for {row['SubBranch']} (Sem {row['Semester']}) on {date_str} at {preferred_slot}")
                
                scheduling_date = exam_date + timedelta(days=1)
            
            current_scheduling_date = scheduling_date  # Update global scheduling date
    
    st.success(f"✅ Successfully scheduled {scheduled_count} subjects in total")
    
    return df

# Integrated PDF Generation Logic Starts Here

# Key Helper Functions for PDF
def calculate_end_time(start_time, duration_hours):
    """Calculate the end time given a start time and duration in hours."""
    start = datetime.strptime(start_time, "%I:%M %p")
    duration = timedelta(hours=duration_hours)
    end = start + duration
    return end.strftime("%I:%M %p").replace("AM", "am").replace("PM", "pm")

def extract_duration(subject_str):
    """Extract duration from subject string using regex."""
    match = re.search(r'\[Duration: (\d+\.?\d*) hrs\]', subject_str)
    if match:
        return float(match.group(1))
    else:
        return 3.0  # Default to 3 hours

def get_semester_default_time_slot(semester, branch):
    """Get default time slot based on semester (odd/even logic)."""
    # Placeholder function: Replace with actual logic or data source
    # Example: B.Tech Semester 5 EXTC defaults to "10:00 AM - 1:00 PM"
    default_timings = {
        ("5", "EXTC"): "10:00 AM - 1:00 PM",
        ("6", "COMP"): "10:00 AM - 1:00 PM",  # Add for other semesters/branches
        # Add more as needed
    }
    return default_timings.get((str(semester), branch), "10:00 AM - 1:00 PM")  # Default fallback

def int_to_roman(num):
    """Convert integer to Roman numeral (for semesters)."""
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

def wrap_text(pdf, text, col_width):
    """Wrap text to fit column width (with caching)."""
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

def generate_pdf_timetable(sem_dict, pdf_path, sub_branch_cols_per_page=4):
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))  # Wide landscape format
    pdf.set_auto_page_break(auto=False, margin=15)
    
    # Enable automatic page numbering with alias
    pdf.alias_nb_pages()
    
    # Assume save_to_excel is defined elsewhere to create df_dict
    # For integration, we'll process sem_dict directly into pivot form similar to extraction
    # Here, we'll adapt the logic to work with sem_dict instead of reading from Excel

    for sem, df_sem in sem_dict.items():
        if df_sem.empty:
            continue
        main_branch = df_sem['MainBranch'].iloc[0] if 'MainBranch' in df_sem.columns else ""
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester_roman = int_to_roman(int(sem)) if str(sem).isdigit() else sem
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}

        # Separate non-electives and electives
        df_non_elec = df_sem[df_sem['OE'].isna() | (df_sem['OE'].str.strip() == "")].copy()
        df_elec = df_sem[df_sem['OE'].notna() & (df_sem['OE'].str.strip() != "")].copy()

        # Handle non-electives (core subjects)
        if not df_non_elec.empty:
            # Create pivot for non-electives
            pivot_df = df_non_elec.pivot_table(
                index=["Exam Date", "Time Slot"],
                columns="SubBranch",
                values="Subject",
                aggfunc=lambda x: ", ".join(x)
            ).fillna("---")
            pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
            fixed_cols = ["Exam Date", "Time Slot"]
            sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols]
            exam_date_width = 60
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

                # Modify subjects to show only the specific time range, not both formats
                default_time_slot = get_semester_default_time_slot(sem, main_branch)
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
                            
                            # Only show custom time if different from 3 hours, don't show both formats
                            if duration != 3 and subject_time_slot and subject_time_slot.strip():
                                start_time = subject_time_slot.split(" - ")[0]
                                end_time = calculate_end_time(start_time, duration)
                                modified_subjects.append(f"{base_subject} ({start_time} to {end_time})")
                            else:
                                # For 3-hour exams, check if time slot is different from default
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
                
                # Add page before printing the table
                pdf.add_page()
                # Add footer with page number to the new page
                footer_height = 25
                add_footer_with_page_number(pdf, footer_height)
                
                print_table_custom(pdf, chunk_df[cols_to_print], cols_to_print, col_widths, line_height=line_height, 
                                   header_content=header_content, branches=chunk, time_slot=time_slot)

        # Handle electives with updated table structure
        if not df_elec.empty:
            # Group by 'OE' and 'Exam Date' to handle multiple subjects per OE type
            elective_data = df_elec.groupby(['OE', 'Exam Date', 'Time Slot']).agg({
                'Subject': lambda x: ", ".join(x)
            }).reset_index()
            elective_data = elective_data.rename(columns={'Subject': 'SubjectDisplay'})

            # Convert Exam Date to desired format
            elective_data["Exam Date"] = pd.to_datetime(elective_data["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")

            # Clean 'SubjectDisplay' to remove [OE] from each subject if present
            elective_data['SubjectDisplay'] = elective_data.apply(
                lambda row: ", ".join([s.replace(f" [{row['OE']}]", "") for s in row['SubjectDisplay'].split(", ")]),
                axis=1
            )

            # Modify subjects for timing overrides - only show specific time range, not both formats
            default_time_slot = get_semester_default_time_slot(sem, main_branch)
            time_slot = df_elec['Time Slot'].iloc[0] if 'Time Slot' in df_elec.columns and not df_elec['Time Slot'].empty else None
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
                    
                    # Only show custom time if different from 3 hours, don't show both formats
                    if duration != 3 and subject_time_slot and subject_time_slot.strip():
                        start_time = subject_time_slot.split(" - ")[0]
                        end_time = calculate_end_time(start_time, duration)
                        modified_subjects.append(f"{base_subject} ({start_time} to {end_time})")
                    else:
                        # For 3-hour exams, check if time slot is different from default
                        if subject_time_slot and default_time_slot and subject_time_slot.strip() and default_time_slot.strip():
                            if subject_time_slot != default_time_slot:
                                modified_subjects.append(f"{base_subject} ({subject_time_slot})")
                            else:
                                modified_subjects.append(base_subject)
                        else:
                            modified_subjects.append(base_subject)
                
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
    
    # Post-processing: Remove blank pages
    try:
        reader = PdfReader(pdf_path)
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
            with open(pdf_path, 'wb') as output_file:
                writer.write(output_file)
        else:
            print("Warning: All pages were filtered out - keeping original PDF")
    except Exception as e:
        print(f"Error during PDF post-processing: {str(e)}")

# Assume other functions like read_timetable, schedule_electives_globally, optimize_oe_subjects_after_scheduling, save_to_excel, save_verification_excel are defined here or in the truncated part.
# For completeness, the code assumes they exist as in the original.

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
                <li>📅 Uncommon subject scheduling</li>
                <li>🔄 All-branch common scheduling</li>
                <li>🎓 OE elective optimization</li>
                <li>⚡ Maximum subjects per day</li>
                <li>📋 PDF generation</li>
                <li>✅ Verification file export</li>
                <li>🎯 Stream-wise scheduling</li>
                <li>📱 Mobile-friendly interface</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    if uploaded_file is not None:
        if st.button("🔄 Generate Timetable", type="primary", use_container_width=True):
            with st.spinner("Processing your timetable... Please wait..."):
                try:
                    # Use holidays from session state
                    holidays_set = st.session_state.get('holidays_set', set())
                    st.write(f"🗓️ Using {len(holidays_set)} holidays: {[h.strftime('%d-%m-%Y') for h in sorted(holidays_set)]}")
                
                    st.write("Reading timetable...")
                    df_non_elec, df_ele, original_df = read_timetable(uploaded_file)

                    if df_non_elec is not None:
                        st.write("Processing subjects...")
                        
                        # Step 1: Schedule uncommon subjects
                        df_scheduled = schedule_uncommon_subjects_first(df_non_elec, holidays_set, base_date)
                        
                        # Step 2: Find the latest date from scheduled uncommon subjects
                        scheduled_uncommon = df_scheduled[df_scheduled['Exam Date'] != ""]
                        if not scheduled_uncommon.empty:
                            scheduled_dates = pd.to_datetime(scheduled_uncommon['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
                            if not scheduled_dates.empty:
                                latest_uncommon_date = max(scheduled_dates)
                                # Start common subject scheduling from the day after the latest uncommon subject
                                common_start_date = latest_uncommon_date + timedelta(days=1)
                                st.write(f"📅 Latest uncommon subject scheduled on: {latest_uncommon_date.strftime('%d-%m-%Y')}")
                                st.write(f"📅 Will start common subjects from: {common_start_date.strftime('%d-%m-%Y')}")
                                
                                # Step 3: Schedule common subjects
                                df_scheduled = schedule_common_subjects_after_uncommon(df_scheduled, holidays_set, common_start_date)
                            else:
                                common_start_date = base_date
                                df_scheduled = schedule_common_subjects_after_uncommon(df_scheduled, holidays_set, common_start_date)
                        else:
                            common_start_date = base_date
                            df_scheduled = schedule_common_subjects_after_uncommon(df_scheduled, holidays_set, common_start_date)
                        
                        # Step 4: Handle electives if they exist
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
                        
                        # Step 5: Create semester dictionary
                        all_scheduled_subjects = all_scheduled_subjects[all_scheduled_subjects['Exam Date'] != ""]
                        
                        if not all_scheduled_subjects.empty:
                            # Sort by semester and date
                            all_scheduled_subjects = all_scheduled_subjects.sort_values(["Semester", "Exam Date"], ascending=True)
                            
                            # Create semester dictionary
                            sem_dict = {}
                            for s in sorted(all_scheduled_subjects["Semester"].unique()):
                                sem_data = all_scheduled_subjects[all_scheduled_subjects["Semester"] == s].copy()
                                sem_dict[s] = sem_data
                            
                            # Step 6: Optimize OE subjects if they exist
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
                            excel_data = save_to_excel(sem_dict)
                            if excel_data:
                                st.session_state.excel_data = excel_data.getvalue()

                            st.write("Generating verification file...")
                            verification_data = save_verification_excel(original_df, sem_dict)
                            if verification_data:
                                st.session_state.verification_data = verification_data.getvalue()

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

                            st.markdown('<div class="status-success">🎉 Complete timetable generated successfully!</div>',
                                        unsafe_allow_html=True)
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
                st.rerun()

        # Statistics Overview
        st.markdown("""
        <div class="stats-section">
            <h2>📈 Complete Timetable Statistics</h2>
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

        # Show efficiency metrics
        if st.session_state.unique_exam_days > 0 and st.session_state.overall_date_range > 0:
            efficiency = (st.session_state.unique_exam_days / st.session_state.overall_date_range) * 100
            if efficiency > 80:
                st.success(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Excellent - most days are utilized)")
            elif efficiency > 60:
                st.info(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Good)")
            else:
                st.warning(f"🎯 **Scheduling Efficiency:** {efficiency:.1f}% (Could be improved)")

        # Timetable Results
        st.markdown("---")
        st.markdown("""
        <div class="results-section">
            <h2>📊 Complete Timetable Results</h2>
        </div>
        """, unsafe_allow_html=True)

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
                        # Format subject display
                        def format_subject_display(row):
                            subject = row['Subject']
                            time_slot = row['Time Slot']
                            duration = row['Exam Duration']
                            
                            # If duration is not 3 hours, show the specific time range
                            if duration != 3 and time_slot and time_slot.strip():
                                start_time = time_slot.split(' - ')[0]
                                end_time = calculate_end_time(start_time, duration)
                                time_range = f" ({start_time} to {end_time})"
                            else:
                                time_range = ""
                            
                            return subject + time_range
                        
                        df_non_elec["SubjectDisplay"] = df_non_elec.apply(format_subject_display, axis=1)
                        df_non_elec["Exam Date"] = pd.to_datetime(df_non_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                        df_non_elec = df_non_elec.sort_values(by="Exam Date", ascending=True)
                        
                        # Create pivot table for non-electives
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
                        # Format elective display
                        def format_elective_display(row):
                            subject = row['Subject']
                            oe_type = row['OE']
                            time_slot = row['Time Slot']
                            duration = row['Exam Duration']
                            
                            base_display = f"{subject} [{oe_type}]"
                            
                            # If duration is not 3 hours, show the specific time range
                            if duration != 3 and time_slot and time_slot.strip():
                                start_time = time_slot.split(' - ')[0]
                                end_time = calculate_end_time(start_time, duration)
                                time_range = f" ({start_time} to {end_time})"
                            else:
                                time_range = ""
                            
                            return base_display + time_range
                        
                        df_elec["SubjectDisplay"] = df_elec.apply(format_elective_display, axis=1)
                        df_elec["Exam Date"] = pd.to_datetime(df_elec["Exam Date"], format="%d-%m-%Y", errors='coerce')
                        df_elec = df_elec.sort_values(by="Exam Date", ascending=True)
                        
                        # Create elective pivot
                        elec_pivot = df_elec.groupby(['OE', 'Exam Date', 'Time Slot'])['SubjectDisplay'].apply(
                            lambda x: ", ".join(x)
                        ).reset_index()
                        
                        if not elec_pivot.empty:
                            st.markdown(f"#### {main_branch_full} - Open Electives")
                            # Format dates for display
                            elec_pivot['Formatted_Date'] = elec_pivot['Exam Date'].dt.strftime("%d-%m-%Y")
                            elec_pivot_display = elec_pivot[['Formatted_Date', 'Time Slot', 'OE', 'SubjectDisplay']].rename(columns={
                                'Formatted_Date': 'Exam Date',
                                'OE': 'OE Type',
                                'SubjectDisplay': 'Subjects'
                            })
                            st.dataframe(elec_pivot_display, use_container_width=True)

    # Display footer
    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <p>🎓 <strong>Complete Timetable Generator</strong></p>
        <p>Developed for MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
        <p style="font-size: 0.9em;">Stream-wise scheduling • All-branch commonality • OE optimization • Maximum efficiency • Verification export</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
