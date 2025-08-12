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
                            preferred_slot = get_preferred_slot(semester)
                            
                            df.loc[idx_to_schedule, 'Exam Date'] = date_str
                            df.loc[idx_to_schedule, 'Time Slot'] = preferred_slot
                            scheduled_count += 1
                            individual_scheduled_today += 1
                            day_had_scheduling = True
                            
                            # Mark this subbranch-semester as scheduled
                            subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                            daily_scheduled[date_str].add(subbranch_sem_key)
                            
                            st.write(f"      ✅ Filled slot with individual subject {row['Subject']} ({row['SubBranch']}, Sem {row['Semester']}) at {preferred_slot}")
                            
                            # Remove from individual_subjects
                            individual_subjects = individual_subjects.drop([idx_to_schedule])
            
            total_scheduled_today = subjects_scheduled_today + (individual_scheduled_today if 'individual_scheduled_today' in locals() else 0)
            st.write(f"  📊 Total subjects scheduled on {date_str}: {total_scheduled_today} ({subjects_scheduled_today} common + {individual_scheduled_today if 'individual_scheduled_today' in locals() else 0} individual)")
            
            # Move to next day
            current_scheduling_date = find_next_valid_day(exam_date + timedelta(days=1), holidays)
            
            # If no scheduling happened this day and we still have unscheduled groups, continue to next day
            if not day_had_scheduling and unscheduled_groups:
                st.write(f"  ⚠️ No subjects could be scheduled on {date_str}, moving to next day")
                continue
    
    # FIXED: Handle any remaining unscheduled common subjects as individual subjects
    remaining_common_subjects = df[
        (df['Exam Date'] == "") & 
        (df['CommonAcrossSems'] == True) & 
        (df['Category'] != 'INTD')
    ].copy()
    
    if not remaining_common_subjects.empty:
        st.warning(f"⚠️ {len(remaining_common_subjects)} common subjects could not be scheduled in groups. Scheduling them individually...")
        
        # Add remaining common subjects to individual subjects for scheduling
        individual_subjects = pd.concat([individual_subjects, remaining_common_subjects], ignore_index=True)
    
    # Schedule remaining individual subjects (including previously unscheduled common subjects)
    if not individual_subjects.empty:
        st.write(f"📚 Scheduling remaining {len(individual_subjects)} individual subjects...")
        
        for idx, row in individual_subjects.iterrows():
            if df.loc[idx, 'Exam Date'] == "":  # Not yet scheduled
                # Find a day where this subbranch-semester is not scheduled
                found_slot = False
                test_date = current_scheduling_date
                attempts = 0
                max_attempts = 200  # Increased from 100
                
                while not found_slot and attempts < max_attempts:
                    test_date = find_next_valid_day(test_date, holidays)
                    date_str = test_date.strftime("%d-%m-%Y")
                    
                    if date_str not in daily_scheduled:
                        daily_scheduled[date_str] = set()
                    
                    subbranch_sem_key = get_subbranch_semester_key(row['SubBranch'], row['Semester'])
                    
                    if subbranch_sem_key not in daily_scheduled[date_str]:
                        # Can schedule here
                        semester = row['Semester']
                        preferred_slot = get_preferred_slot(semester)
                        
                        df.loc[idx, 'Exam Date'] = date_str
                        df.loc[idx, 'Time Slot'] = preferred_slot
                        scheduled_count += 1
                        
                        daily_scheduled[date_str].add(subbranch_sem_key)
                        
                        # Check if this was originally a common subject
                        original_common = "✓" if row['CommonAcrossSems'] else ""
                        st.write(f"    ✅ Scheduled subject {row['Subject']} ({row['SubBranch']}, Sem {row['Semester']}) on {date_str} {original_common}")
                        found_slot = True
                    else:
                        test_date += timedelta(days=1)
                        attempts += 1
                
                if not found_slot:
                    st.error(f"❌ Could not schedule {row['Subject']} ({row['SubBranch']}, Sem {row['Semester']}) after {max_attempts} attempts")
    
    # Final count of unscheduled subjects
    final_unscheduled = df[
        (df['Exam Date'] == "") & 
        (df['Category'] != 'INTD')
    ]
    
    if not final_unscheduled.empty:
        st.error(f"❌ {len(final_unscheduled)} subjects remain unscheduled:")
        for idx, row in final_unscheduled.iterrows():
            common_status = "Common" if row['CommonAcrossSems'] else "Individual"
            st.write(f"  • {row['Subject']} ({row['SubBranch']}, Sem {row['Semester']}) - {common_status}")
    else:
        st.success("✅ All subjects successfully scheduled!")
    
    st.success(f"✅ Successfully scheduled {scheduled_count} subjects across all branches")
    
    # Display scheduling efficiency summary with subbranch details
    if daily_scheduled:
        st.markdown("### 📊 Daily Scheduling Summary")
        total_days = len(daily_scheduled)
        total_subjects_scheduled = sum(len([idx for idx in df.index if df.loc[idx, 'Exam Date'] == date_str]) for date_str in daily_scheduled.keys())
        avg_subjects_per_day = total_subjects_scheduled / total_days if total_days > 0 else 0
        
        st.write(f"📈 **Overall Statistics:**")
        st.write(f"  • Total exam days: {total_days}")
        st.write(f"  • Total subjects scheduled: {total_subjects_scheduled}")
        st.write(f"  • Average subjects per day: {avg_subjects_per_day:.1f}")
        
        st.write(f"\n📅 **Day-by-day breakdown (with subbranch constraint verification):**")
        for date_str in sorted(daily_scheduled.keys(), key=lambda x: datetime.strptime(x, "%d-%m-%Y")):
            subjects_on_date = len([idx for idx in df.index if df.loc[idx, 'Exam Date'] == date_str])
            date_obj = datetime.strptime(date_str, "%d-%m-%Y")
            day_name = date_obj.strftime("%A")
            
            # Verify no subbranch-semester conflicts on this date
            scheduled_subbranch_sems = daily_scheduled[date_str]
            st.write(f"  • {day_name}, {date_str}: {subjects_on_date} subjects, {len(scheduled_subbranch_sems)} unique subbranch-semester combinations")
    
    return df
    
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

                # FIXED: Modify subjects to show only the specific time range, not both formats
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
                            
                            # FIXED: Only show custom time if different from 3 hours, don't show both formats
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

            # FIXED: Modify subjects for timing overrides - only show specific time range, not both formats
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
                    
                    # FIXED: Only show custom time if different from 3 hours, don't show both formats
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
        
        # Standardize column names first
        df.columns = df.columns.str.strip()
        
        # Create a flexible column mapping system
        column_mapping = {
            "Program": ["Program", "Programme"],
            "Stream": ["Stream", "Specialization", "Branch"],
            "Current Session": ["Current Session", "Semester", "Current Academic Session"],
            "Module Description": ["Module Description", "SubjectName", "Subject Name", "Subject", "Module Information"],
            "Module Abbreviation": ["Module Abbreviation", "ModuleCode", "Module Code", "Code"],
            "Campus Name": ["Campus Name", "Campus"],
            "Difficulty Score": ["Difficulty Score", "Difficulty"],
            "Exam Duration": ["Exam Duration", "Duration", "Exam_Duration"],
            "Student count": ["Student count", "StudentCount", "Student_count", "Count"],
            "Common across sems": ["Common across sems", "CommonAcrossSems", "Common_across_sems", "IsCommon"],
            "Circuit": ["Circuit", "Is_Circuit", "CircuitBranch"],
            "Category": ["Category"],
            "OE": ["OE"],
            "School Name": ["School Name", "School"]
        }
        
        # Apply flexible mapping
        rename_dict = {}
        for standard_name, possible_names in column_mapping.items():
            for possible_name in possible_names:
                if possible_name in df.columns:
                    rename_dict[possible_name] = standard_name
                    break
        
        # Rename columns using the mapping
        df = df.rename(columns=rename_dict)
        
        # Convert semester format
        def convert_sem(sem):
            if pd.isna(sem):
                return 0
            sem_str = str(sem).strip()
            if 'Sem' in sem_str:
                m = {
                    "Sem I": 1, "Sem II": 2, "Sem III": 3, "Sem IV": 4,
                    "Sem V": 5, "Sem VI": 6, "Sem VII": 7, "Sem VIII": 8,
                    "Sem IX": 9, "Sem X": 10, "Sem XI": 11
                }
                return m.get(sem_str, 0)
            else:
                # Extract number from string
                import re
                match = re.search(r'(\d+)', sem_str)
                return int(match.group(1)) if match else 0
        
        df["Semester"] = df["Current Session"].apply(convert_sem).astype(int)
        
        # Create Branch identifier - handle different input formats
        if "School Name" in df.columns and "Program" in df.columns:
            df["Branch"] = df["School Name"].astype(str).str.strip() + " " + df["Program"].astype(str).str.strip()
        elif "Program" in df.columns and "Stream" in df.columns:
            df["Branch"] = df["Program"].astype(str).str.strip() + "-" + df["Stream"].astype(str).str.strip()
        else:
            df["Branch"] = df.get("Program", "Unknown").astype(str).str.strip()
        
        # Create Subject field
        if "Module Description" in df.columns and "Module Abbreviation" in df.columns:
            df["Subject"] = df["Module Description"].astype(str) + " - (" + df["Module Abbreviation"].astype(str) + ")"
        else:
            df["Subject"] = df.get("Module Description", df.get("Subject", "Unknown Subject")).astype(str)
        
        # Handle difficulty scoring
        if "Category" in df.columns and "Difficulty Score" in df.columns:
            comp_mask = (df["Category"] == "COMP") & df["Difficulty Score"].notna()
            df["Difficulty"] = None
            df.loc[comp_mask, "Difficulty"] = df.loc[comp_mask, "Difficulty Score"]
        else:
            df["Difficulty"] = None
        
        # Initialize scheduling columns
        df["Exam Date"] = ""
        df["Time Slot"] = ""
        df["Exam Duration"] = df.get("Exam Duration", 3).fillna(3).astype(float)
        df["StudentCount"] = df.get("Student count", 0).fillna(0).astype(int)
        df["CommonAcrossSems"] = df.get("Common across sems", False).fillna(False).astype(bool)
        df["Circuit"] = df.get("Circuit", False).fillna(False).astype(bool)
        
        # Add ModuleCode if missing
        if "ModuleCode" not in df.columns:
            df["ModuleCode"] = df.get("Module Abbreviation", "")
        
        # Separate electives and non-electives
        if "Category" in df.columns:
            df_ele = df[df["Category"] == "INTD"].copy()
            df_non = df[df["Category"] != "INTD"].copy()
        else:
            df_ele = pd.DataFrame()
            df_non = df.copy()
        
        # Create MainBranch and SubBranch
        def split_br(b):
            b_str = str(b).strip()
            if " " in b_str:  # Handle "School Program" format
                parts = b_str.split(" ", 1)
                return pd.Series([parts[0].strip(), parts[1].strip() if len(parts) > 1 else ""])
            elif "-" in b_str:  # Handle "Program-Stream" format
                parts = b_str.split("-", 1)
                return pd.Series([parts[0].strip(), parts[1].strip() if len(parts) > 1 else ""])
            else:
                return pd.Series([b_str, ""])
        
        for d in [df_non, df_ele]:
            if not d.empty:
                d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)
                # If SubBranch is empty, use Stream if available
                if "Stream" in d.columns:
                    empty_subbranch = d["SubBranch"] == ""
                    d.loc[empty_subbranch, "SubBranch"] = d.loc[empty_subbranch, "Stream"].fillna("")
        
        # Define required columns
        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty", "Exam Duration", "StudentCount", "CommonAcrossSems", "ModuleCode", "Circuit"]
        
        # Ensure all required columns exist
        for col in cols:
            if col not in df_non.columns:
                df_non[col] = "" if col in ["Category", "OE", "Exam Date", "Time Slot"] else 0
            if col not in df_ele.columns:
                df_ele[col] = "" if col in ["Category", "OE", "Exam Date", "Time Slot"] else 0
        
        return df_non[cols], df_ele[cols], df
        
    except Exception as e:
        st.error(f"Error reading the Excel file: {str(e)}")
        return None, None, None

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
        "Program": ["Program", "Programme", "School Name"],
        "Stream": ["Stream", "Specialization", "Branch"],
        "Module Description": ["Module Description", "SubjectName", "Subject Name", "Subject", "Module Information"],
        "Exam Duration": ["Exam Duration", "Duration", "Exam_Duration"],
        "Student count": ["Student count", "StudentCount", "Student_count", "Count"],
        "Common across sems": ["Common across sems", "CommonAcrossSems", "Common_across_sems", "IsCommon"],
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
    verification_df["Is Common"] = ""
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
            elif ' ' in branch:
                program, stream = branch.split(' ', 1)
                keys_to_try.append(f"{module_code}_{semester}_{program.strip()} {stream.strip()}")
            
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
            
            # Create branch identifier - handle different formats
            program = str(row.get("Program", "")).strip()
            stream = str(row.get("Stream", "")).strip()
            
            # Try different branch format combinations
            branch_options = []
            if program and stream and program != "nan" and stream != "nan":
                branch_options.append(f"{program}-{stream}")  # Program-Stream format
                branch_options.append(f"{program} {stream}")  # Program Stream format
            elif program and program != "nan":
                branch_options.append(program)
            
            if not branch_options:
                st.write(f"⚠️ Empty branch for module {module_code}")
                unmatched_count += 1
                continue
            
            # Try to find match using lookup
            match_found = False
            matched_subject = None
            
            for branch in branch_options:
                lookup_key = f"{module_code}_{semester_num}_{branch}"
                
                if lookup_key in scheduled_lookup:
                    # Found exact match
                    matched_subject = scheduled_lookup[lookup_key][0]  # Take first match
                    match_found = True
                    break
            
            # If no exact match, try alternative# If no exact match, try alternative matching approaches
           if not match_found:
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
                               for branch_option in branch_options:
                                   if branch_option in key_branch or key_branch in branch_option:
                                       matched_subject = subjects[0]
                                       match_found = True
                                       break
                           if match_found:
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
               
               # Calculate exam time
               if time_slot and str(time_slot).strip() and str(time_slot) != "nan":
                   try:
                       start_time = str(time_slot).split(" - ")[0].strip()
                       end_time = calculate_end_time(start_time, duration)
                       exam_time = f"{start_time} to {end_time}"
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
               
               # Check commonality - improved logic
               same_module_scheduled = scheduled_data[scheduled_data["ExtractedModuleCode"] == module_code]
               unique_branches = same_module_scheduled["Branch"].nunique()
               is_marked_common = row.get("Common across sems", False)
               verification_df.at[idx, "Is Common"] = "YES" if (unique_branches > 1 or is_marked_common) else "NO"
               
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
               verification_df.at[idx, "Is Common"] = "N/A"
               verification_df.at[idx, "Scheduling Status"] = "Not Scheduled"
               unmatched_count += 1
               
               if unmatched_count <= 10:  # Show first 10 unmatched for debugging
                   st.write(f"   ❌ **NO MATCH** for {module_code} ({branch_options}, Sem {semester_num})")
                       
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
       
       # Add a summary sheet
       summary_data = {
           "Metric": ["Total Subjects", "Scheduled Subjects", "Unscheduled Subjects", "Match Rate (%)"],
           "Value": [matched_count + unmatched_count, matched_count, unmatched_count, 
                    round(matched_count/(matched_count+unmatched_count)*100, 1) if (matched_count+unmatched_count) > 0 else 0]
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

        # Define the subject display formatting functions
        def format_subject_display(row):
            """Format subject display for non-electives"""
            subject = row['Subject']
            time_slot = row['Time Slot']
            duration = row.get('Exam Duration', 3)
            is_common = row.get('CommonAcrossSems', False)
            semester = row['Semester']
            
            # Get preferred slot for this semester
            preferred_slot = get_preferred_slot(semester)
            
            # Determine if we need to show time info
            add_time = is_common and time_slot != preferred_slot
            time_range = ""
            
            if pd.notna(duration):
                if duration != 3 and time_slot and time_slot.strip():
                    # Non-standard duration - show specific time range
                    start_time = time_slot.split(' - ')[0].strip()
                    end_time = calculate_end_time(start_time, duration)
                    time_range = f" ({start_time} to {end_time})"
                elif add_time:
                    # Common subject with non-preferred slot
                    time_range = f" ({time_slot})"
            
            return subject + time_range

        def format_elective_display(row):
            """Format subject display for electives"""
            subject = row['Subject']
            oe_type = row.get('OE', '')
            base_display = f"{subject} [{oe_type}]" if oe_type else subject
            
            duration = row.get('Exam Duration', 3)
            time_slot = row['Time Slot']
            time_range = ""
            
            if pd.notna(duration) and duration != 3 and time_slot and time_slot.strip():
                start_time = time_slot.split(' - ')[0].strip()
                end_time = calculate_end_time(start_time, duration)
                time_range = f" ({start_time} to {end_time})"
            
            return base_display + time_range

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
                        # Apply formatting using the new function
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
                        # Apply formatting using the new function
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




