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

# PDF generation utilities
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
    if pdf.get_y() + row_h > pdf.page_break_trigger:
        pdf.add_page(pdf.cur_orientation)

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

def print_table_custom(pdf, df, columns, col_widths, line_height=5):
    if df.empty:
        return
    setattr(pdf, '_row_counter', 0)
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        if any(cell.strip() for cell in row):
            print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=True, margin=15)
    df_dict = pd.read_excel(excel_path, sheet_name=None, index_col=[0, 1])
    for sheet_name, pivot_df in df_dict.items():
        if pivot_df.empty:
            continue
        pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
        fixed_cols = ["Exam Date", "Time Slot"]
        sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols]
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        exam_date_width = 30
        time_slot_width = 40
        table_font_size = 8
        line_height = 8
        for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
            chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
            cols_to_print = fixed_cols + chunk
            chunk_df = pivot_df[cols_to_print].copy()
            mask = chunk_df[chunk].apply(lambda row: row.astype(str).str.strip() != "").any(axis=1)
            chunk_df = chunk_df[mask].reset_index(drop=True)
            if chunk_df.empty:
                continue
            pdf.add_page()
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
            pdf.cell(pdf.w - 20, 8, f"{main_branch_full} - Semester {semester}", 0, 1, 'C')
            ts = chunk_df['Time Slot'].iloc[0]
            if ts.strip():
                pdf.set_font("Arial", 'I', 13)
                pdf.set_xy(10, 61)
                pdf.cell(pdf.w - 20, 6, f"Time Slot: {ts}", 0, 1, 'C')
            pdf.set_font("Arial", '', 12)
            pdf.set_xy(10, 69)
            pdf.cell(pdf.w - 20, 6, f"Branches: {', '.join(chunk)}", 0, 1, 'C')
            page_width = pdf.w - 2 * pdf.l_margin
            remaining = page_width - (exam_date_width + time_slot_width)
            sub_width = remaining / max(len(chunk), 1)
            col_widths = [exam_date_width, time_slot_width] + [sub_width] * len(chunk)
            total_w = sum(col_widths)
            if total_w > page_width:
                factor = page_width / total_w
                col_widths = [w * factor for w in col_widths]
            pdf.set_y(83)
            pdf.set_font("Arial", size=table_font_size)
            print_table_custom(pdf, chunk_df, cols_to_print, col_widths, line_height=line_height)
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
            "Difficulty Score": "Difficulty"
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
        df = df[df["Campus"].str.strip() == "Mumbai"]

        # Handle Difficulty column: assign only for COMP category, else None
        comp_mask = (df["Category"] == "COMP") & df["Difficulty"].notna()
        df["Difficulty"] = None
        df.loc[comp_mask, "Difficulty"] = df.loc[comp_mask, "Difficulty"]

        df["Exam Date"] = ""
        df["Time Slot"] = ""
        df_non = df[df["Category"] != "INTD"].copy()
        df_ele = df[df["Category"] == "INTD"].copy()

        def split_br(b):
            p = b.split("-", 1)
            return pd.Series([p[0].strip(), p[1].strip() if len(p) > 1 else ""])

        for d in (df_non, df_ele):
            d[["MainBranch", "SubBranch"]] = d["Branch"].apply(split_br)

        cols = ["MainBranch", "SubBranch", "Branch", "Semester", "Subject", "Category", "OE", "Exam Date", "Time Slot",
                "Difficulty"]
        return df_non[cols], df_ele[cols]
    except Exception as e:
        st.error(f"Error reading the Excel file: {str(e)}")
        return None, None

def schedule_semester_non_electives(df_sem, holidays, base_date, schedule_by_difficulty=False):
    df_sem['SubjectCode'] = df_sem['Subject'].str.extract(r'\((.*?)\)', expand=False)

    # Convert base_date to datetime if it's a date object
    if isinstance(base_date, date) and not isinstance(base_date, datetime):
        base_date = datetime.combine(base_date, datetime.min.time())

    # Convert holidays to a set of dates (ignoring time part) for comparison
    holidays_dates = {h.date() for h in holidays}

    all_branches = df_sem['Branch'].unique()
    last_exam = {branch: None for branch in all_branches}
    exam_days = {branch: set() for branch in all_branches}

    def find_next_valid_day(start_day, for_branches):
        day = start_day
        while True:
            day_date = day.date()
            if day.weekday() == 6 or day_date in holidays_dates:
                day += timedelta(days=1)
                continue
            if all(
                    (day_date not in {d.date() for d in exam_days[branch]}) and
                    (last_exam[branch] is None or (day - last_exam[branch]).days >= 2)
                    for branch in for_branches
            ):
                return day
            day += timedelta(days=1)

    if schedule_by_difficulty:
        # Precompute mappings using vectorized operations
        subject_info = df_sem.groupby('SubjectCode').agg({
            'Difficulty': 'first',
            'Branch': set
        }).reset_index()
        subject_to_difficulty = dict(zip(subject_info['SubjectCode'], subject_info['Difficulty']))
        subject_to_branches = dict(zip(subject_info['SubjectCode'], subject_info['Branch']))

        common_subject_codes = {sc for sc, branches in subject_to_branches.items() if len(branches) > 1}

        common_subjects = {'easy': [], 'difficult': [], 'neutral': []}
        for subj_code in common_subject_codes:
            difficulty = subject_to_difficulty.get(subj_code)
            if difficulty == 0:
                common_subjects['easy'].append(subj_code)
            elif difficulty == 1:
                common_subjects['difficult'].append(subj_code)
            else:
                common_subjects['neutral'].append(subj_code)

        common_schedule_order = []
        last_difficulty = None
        easy, difficult, neutral = common_subjects['easy'], common_subjects['difficult'], common_subjects['neutral']
        while easy or difficult or neutral:
            if last_difficulty is None or last_difficulty == 1:
                if easy:
                    subj_code = easy.pop(0)
                    common_schedule_order.append(subj_code)
                    last_difficulty = 0
                elif neutral:
                    subj_code = neutral.pop(0)
                    common_schedule_order.append(subj_code)
                    last_difficulty = None
                elif difficult:
                    subj_code = difficult.pop(0)
                    common_schedule_order.append(subj_code)
                    last_difficulty = 1
            else:
                if difficult:
                    subj_code = difficult.pop(0)
                    common_schedule_order.append(subj_code)
                    last_difficulty = 1
                elif neutral:
                    subj_code = neutral.pop(0)
                    common_schedule_order.append(subj_code)
                    last_difficulty = None
                elif easy:
                    subj_code = easy.pop(0)
                    common_schedule_order.append(subj_code)
                    last_difficulty = 0

        current_day = base_date
        for subj_code in common_schedule_order:
            branches = subject_to_branches[subj_code]
            exam_day = find_next_valid_day(current_day, branches)
            mask = (df_sem['SubjectCode'] == subj_code) & (df_sem['Branch'].isin(branches))
            df_sem.loc[mask, 'Exam Date'] = exam_day.strftime("%d-%m-%Y")
            for branch in branches:
                exam_days[branch].add(exam_day)
                last_exam[branch] = exam_day
            current_day = exam_day + timedelta(days=1)

        scheduled_subjects = set(df_sem[df_sem['Exam Date'] != '']['Subject'])
        remaining_subjects_per_branch = {}
        for branch in all_branches:
            branch_subjects = set(df_sem[df_sem['Branch'] == branch]['Subject'])
            remaining_subjects_per_branch[branch] = list(branch_subjects - scheduled_subjects)

        remaining_df = df_sem[df_sem['Exam Date'] == ''].copy()
        remaining_df['DifficultyCategory'] = remaining_df['Difficulty'].map({0: 'easy', 1: 'difficult'}).fillna(
            'neutral')
        subjects_by_difficulty = remaining_df.groupby(['Branch', 'DifficultyCategory'])['Subject'].apply(list).unstack(
            fill_value=[])

        remaining_by_difficulty = {branch: {'easy': deque(), 'difficult': deque(), 'neutral': deque()} for branch in
                                   all_branches}
        for branch in all_branches:
            for cat in ['easy', 'difficult', 'neutral']:
                if branch in subjects_by_difficulty.index:
                    subjects = subjects_by_difficulty.loc[branch].get(cat, [])
                    remaining_by_difficulty[branch][cat].extend(subjects)

        next_available_date = {branch: current_day for branch in all_branches}
        last_difficulty_branch = {branch: None for branch in all_branches}

        while any(
                len(remaining_by_difficulty[branch]['easy']) > 0 or
                len(remaining_by_difficulty[branch]['difficult']) > 0 or
                len(remaining_by_difficulty[branch]['neutral']) > 0
                for branch in all_branches
        ):
            branches_with_subjects = [
                branch for branch in all_branches
                if (remaining_by_difficulty[branch]['easy'] or
                    remaining_by_difficulty[branch]['difficult'] or
                    remaining_by_difficulty[branch]['neutral'])
            ]
            if not branches_with_subjects:
                break

            earliest_date = min(next_available_date[branch] for branch in branches_with_subjects)
            branches_to_schedule = [
                branch for branch in branches_with_subjects
                if next_available_date[branch] == earliest_date
            ]

            individual_assignments = []
            for branch in branches_to_schedule:
                if not (remaining_by_difficulty[branch]['easy'] or
                        remaining_by_difficulty[branch]['difficult'] or
                        remaining_by_difficulty[branch]['neutral']):
                    continue

                exam_day = find_next_valid_day(earliest_date, [branch])
                if last_difficulty_branch[branch] is None or last_difficulty_branch[branch] == 1:
                    if remaining_by_difficulty[branch]['easy']:
                        subj = remaining_by_difficulty[branch]['easy'].popleft()
                        difficulty = 0
                    elif remaining_by_difficulty[branch]['neutral']:
                        subj = remaining_by_difficulty[branch]['neutral'].popleft()
                        difficulty = None
                    elif remaining_by_difficulty[branch]['difficult']:
                        subj = remaining_by_difficulty[branch]['difficult'].popleft()
                        difficulty = 1
                    else:
                        continue
                else:
                    if remaining_by_difficulty[branch]['difficult']:
                        subj = remaining_by_difficulty[branch]['difficult'].popleft()
                        difficulty = 1
                    elif remaining_by_difficulty[branch]['neutral']:
                        subj = remaining_by_difficulty[branch]['neutral'].popleft()
                        difficulty = None
                    elif remaining_by_difficulty[branch]['easy']:
                        subj = remaining_by_difficulty[branch]['easy'].popleft()
                        difficulty = 0
                    else:
                        continue

                individual_assignments.append((branch, subj, exam_day.strftime("%d-%m-%Y")))
                exam_days[branch].add(exam_day)
                last_exam[branch] = exam_day
                last_difficulty_branch[branch] = difficulty
                next_available_date[branch] = exam_day + timedelta(days=2)

            for branch, subj, exam_date in individual_assignments:
                df_sem.loc[(df_sem['Branch'] == branch) & (df_sem['Subject'] == subj), 'Exam Date'] = exam_date

    else:
        subject_branch_count = df_sem.groupby('SubjectCode')['Branch'].nunique()
        common_subject_codes = subject_branch_count[subject_branch_count > 1].index.tolist()
        common_subject_branches = {}
        for subj_code in common_subject_codes:
            branches = df_sem[df_sem['SubjectCode'] == subj_code]['Branch'].unique()
            common_subject_branches[subj_code] = branches
        common_subject_codes_sorted = sorted(common_subject_codes, key=lambda x: len(common_subject_branches[x]),
                                             reverse=True)

        for subj_code in common_subject_codes_sorted:
            branches = common_subject_branches[subj_code]
            exam_day = find_next_valid_day(base_date, branches)
            df_sem.loc[
                (df_sem['SubjectCode'] == subj_code) & (
                    df_sem['Branch'].isin(branches)), 'Exam Date'] = exam_day.strftime("%d-%m-%Y")
            for branch in branches:
                exam_days[branch].add(exam_day)
                last_exam[branch] = exam_day

        remaining_subjects = {}
        for branch in all_branches:
            branch_df = df_sem[df_sem['Branch'] == branch]
            scheduled_subjs = branch_df[branch_df['Exam Date'] != '']['Subject'].unique()
            all_subjs = branch_df['Subject'].unique()
            remaining = [subj for subj in all_subjs if subj not in scheduled_subjs]
            remaining_subjects[branch] = list(remaining)

        current_day = base_date
        while any(remaining_subjects[branch] for branch in all_branches):
            current_day_date = current_day.date()
            if current_day.weekday() == 6 or current_day_date in holidays_dates:
                current_day += timedelta(days=1)
                continue
            branch_list = list(all_branches)
            random.shuffle(branch_list)
            for branch in branch_list:
                if remaining_subjects[branch]:
                    if (current_day.date() not in {d.date() for d in exam_days[branch]}) and (
                            last_exam[branch] is None or (current_day - last_exam[branch]).days >= 2):
                        subj = remaining_subjects[branch].pop(0)
                        df_sem.loc[
                            (df_sem['Branch'] == branch) & (
                                        df_sem['Subject'] == subj), 'Exam Date'] = current_day.strftime("%d-%m-%Y")
                        exam_days[branch].add(current_day)
                        last_exam[branch] = current_day
            current_day += timedelta(days=1)

    sem = df_sem["Semester"].iloc[0]
    slot_str = "10:00 AM - 1:00 PM" if sem % 2 == 0 else "2:00 PM - 5:00 PM"
    df_sem['Time Slot'] = slot_str
    return df_sem

def process_constraints(df, holidays, base_date, schedule_by_difficulty=False):
    final_list = []
    for sem in sorted(df["Semester"].unique()):
        if sem == 0:
            continue
        df_sem = df[df["Semester"] == sem].copy()
        if df_sem.empty:
            continue
        scheduled_sem = schedule_semester_non_electives(df_sem, holidays, base_date, schedule_by_difficulty)
        final_list.append(scheduled_sem)
    if not final_list:
        return {}
    df_combined = pd.concat(final_list, ignore_index=True)
    sem_dict = {}
    for sem in sorted(df_combined["Semester"].unique()):
        sem_dict[sem] = df_combined[df_combined["Semester"] == sem].copy()
    return sem_dict

def schedule_electives_mainbranch(df_elec, elective_base_date, holidays, max_days=90):
    sub_branches = df_elec["SubBranch"].unique().tolist()
    oe_to_subjects = df_elec.groupby("OE")["Subject"].apply(list).to_dict()
    oe_to_subbranches = {
        oe: set(df_elec.loc[df_elec["OE"] == oe, "SubBranch"].unique())
        for oe in oe_to_subjects
    }
    common_oe = [oe for oe, sbs in oe_to_subbranches.items() if len(sbs) == len(sub_branches)]
    individual_oe = [oe for oe, sbs in oe_to_subbranches.items() if len(sbs) < len(sub_branches)]
    schedule_oe_date = {}
    day = elective_base_date

    if isinstance(day, date) and not isinstance(day, datetime):
        day = datetime.combine(day, datetime.min.time())

    holidays_dates = {h.date() for h in holidays}

    def advance_to_next_valid(d):
        while True:
            d_date = d.date()
            if d.weekday() == 6 or d_date in holidays_dates:
                d += timedelta(days=1)
                continue
            return d

    for oe in common_oe:
        day = advance_to_next_valid(day)
        schedule_oe_date[oe] = day
        day += timedelta(days=2)

    for oe in individual_oe:
        day = advance_to_next_valid(day)
        schedule_oe_date[oe] = day
        day += timedelta(days=2)

    sem = df_elec["Semester"].iloc[0]
    time_slot = "10:00 AM - 1:00 PM" if sem % 2 == 0 else "2:00 PM - 5:00 PM"
    for idx, row in df_elec.iterrows():
        oe = row["OE"]
        exam_day = schedule_oe_date.get(oe)
        if exam_day:
            df_elec.at[idx, "Exam Date"] = exam_day.strftime("%d-%m-%Y")
            df_elec.at[idx, "Time Slot"] = time_slot
        else:
            df_elec.at[idx, "Exam Date"] = "Not Scheduled"
            df_elec.at[idx, "Time Slot"] = "N/A"
    return df_elec

def save_to_excel(semester_wise_timetable):
    if not semester_wise_timetable:
        return None

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sem, df_sem in semester_wise_timetable.items():
            for main_branch in df_sem["MainBranch"].unique():
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()
                has_oe = df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")
                difficulty_str = df_mb['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                df_mb["SubjectDisplay"] = df_mb["Subject"]
                df_mb.loc[has_oe, "SubjectDisplay"] = (
                        df_mb.loc[has_oe, "Subject"] + " [" + df_mb.loc[has_oe, "OE"] + "]"
                )
                df_mb["SubjectDisplay"] = df_mb["SubjectDisplay"] + difficulty_suffix
                df_mb["Exam Date"] = pd.to_datetime(
                    df_mb["Exam Date"], format="%d-%m-%Y", errors='coerce'
                )
                df_mb = df_mb.sort_values(by="Exam Date", ascending=True)
                pivot_df = df_mb.pivot_table(
                    index=["Exam Date", "Time Slot"],
                    columns="SubBranch",
                    values="SubjectDisplay",
                    aggfunc=lambda x: ", ".join(x)
                ).fillna("")
                pivot_df = pivot_df.sort_index(level="Exam Date", ascending=True)
                formatted_dates = [d.strftime("%d-%m-%Y") for d in pivot_df.index.levels[0]]
                pivot_df.index = pivot_df.index.set_levels(formatted_dates, level=0)
                sheet_name = f"{main_branch}_Sem_{sem}"
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                pivot_df.to_excel(writer, sheet_name=sheet_name)

    output.seek(0)
    return output

def main():
    # Header section
    st.markdown("""
    <div class="main-header">
        <h1>📅 Exam Timetable Generator</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    # Initialize session state for custom holidays
    if 'num_custom_holidays' not in st.session_state:
        st.session_state.num_custom_holidays = 1
    if 'custom_holidays' not in st.session_state:
        st.session_state.custom_holidays = [None]

    # Sidebar configuration
    with st.sidebar:
        st.markdown("### ⚙️ Configuration")

        # Holiday selection
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

        # Custom holidays
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

        # Base date selection
        st.markdown("#### 📅 Base Date for Scheduling")
        base_date = st.date_input("Start date for exams", value=datetime(2025, 4, 1))
        base_date = datetime.combine(base_date, datetime.min.time())

        # Scheduling mode selection
        st.markdown("#### 🛠️ Scheduling Mode")
        schedule_by_difficulty = st.checkbox("Schedule by Difficulty (Alternate Easy/Difficult)", value=False)
        if schedule_by_difficulty:
            st.markdown('<div class="status-info">ℹ️ Exams will alternate between Easy and Difficult subjects.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-info">ℹ️ Normal scheduling without considering difficulty.</div>',
                        unsafe_allow_html=True)

    # Main content area
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
                <li>📅 Smart scheduling</li>
                <li>📋 Multiple output formats</li>
                <li>🎯 Conflict resolution</li>
                <li>📱 Mobile-friendly interface</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    # Process button and results
    if uploaded_file is not None:
        if st.button("🔄 Generate Timetable", type="primary", use_container_width=True):
            with st.spinner("Processing your timetable... Please wait..."):
                try:
                    holidays_set = set(holiday_dates)
                    df_non_elec, df_elec = read_timetable(uploaded_file)

                    if df_non_elec is not None and df_elec is not None:
                        # Process non-electives with the selected scheduling mode
                        non_elec_sched = process_constraints(df_non_elec, holidays_set, base_date,
                                                             schedule_by_difficulty)

                        # Process electives
                        final = []
                        for sem in sorted(non_elec_sched.keys()):
                            dfn = non_elec_sched[sem]
                            for mb in dfn["MainBranch"].unique():
                                non_df = dfn[dfn["MainBranch"] == mb].copy()
                                non_df["Exam Date Parsed"] = pd.to_datetime(non_df["Exam Date"], format="%d-%m-%Y",
                                                                            errors='coerce')
                                maxd = non_df["Exam Date Parsed"].max()
                                base = maxd + timedelta(days=2) if pd.notna(maxd) else base_date
                                ele_subset = df_elec[
                                    (df_elec["Semester"] == sem) & (df_elec["MainBranch"] == mb)].copy()
                                if not ele_subset.empty:
                                    sched_e = schedule_electives_mainbranch(ele_subset, base, holidays_set)
                                    comb = pd.concat([non_df, sched_e], ignore_index=True)
                                else:
                                    comb = non_df
                                final.append(comb)

                        if final:
                            final_df = pd.concat(final, ignore_index=True)
                            final_df["Exam Date"] = pd.to_datetime(final_df["Exam Date"], format="%d-%m-%Y",
                                                                   errors='coerce')
                            final_df = final_df.sort_values(["Exam Date", "Semester", "MainBranch"], ascending=True,
                                                            na_position='last')
                            sem_dict = {s: final_df[final_df["Semester"] == s].copy() for s in
                                        sorted(final_df["Semester"].unique())}

                            st.session_state.timetable_data = sem_dict
                            st.session_state.processing_complete = True

                            st.markdown('<div class="status-success">🎉 Timetable generated successfully!</div>',
                                        unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="status-error">❌ No valid data found to process.</div>',
                                        unsafe_allow_html=True)
                    else:
                        st.markdown(
                            '<div class="status-error">❌ Failed to read the Excel file. Please check the format.</div>',
                            unsafe_allow_html=True)

                except Exception as e:
                    st.markdown(f'<div class="status-error">❌ An error occurred: {str(e)}</div>',
                                unsafe_allow_html=True)

    # Display results if available
    if hasattr(st.session_state, 'processing_complete') and st.session_state.processing_complete:
        st.markdown("---")

        # Enhanced Statistics Section
        sem_dict = st.session_state.timetable_data
        total_exams = sum(len(df) for df in sem_dict.values())
        total_semesters = len(sem_dict)
        total_branches = len(set(branch for df in sem_dict.values() for branch in df['MainBranch'].unique()))

        # Combine all semester data into a single DataFrame for analysis
        all_data = pd.concat(sem_dict.values(), ignore_index=True)

        # Overall date range
        all_dates = pd.to_datetime(all_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        overall_date_range = (max(all_dates) - min(all_dates)).days if all_dates.size > 0 else 0

        # 1. Date range of non-elective exams (OE is NaN or empty)
        non_elective_data = all_data[all_data['OE'].isna() | (all_data['OE'].str.strip() == "")]
        non_elective_dates = pd.to_datetime(non_elective_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        non_elective_range = "N/A"
        if non_elective_dates.size > 0:
            non_elective_start = min(non_elective_dates).strftime("%d %b %Y")
            non_elective_end = max(non_elective_dates).strftime("%d %b %Y")
            non_elective_range = f"{non_elective_start} to {non_elective_end}"

        # 2. Dates of elective exams (OE is not NaN and not empty)
        elective_data = all_data[all_data['OE'].notna() & (all_data['OE'].str.strip() != "")]
        elective_dates = pd.to_datetime(elective_data['Exam Date'], format="%d-%m-%Y", errors='coerce').dropna()
        elective_dates_list = sorted(set(elective_dates.dt.strftime("%d %b %Y")))
        elective_dates_str = ", ".join(elective_dates_list) if elective_dates_list else "N/A"

        # 3. Number of subjects per stream (MainBranch + SubBranch)
        stream_counts = all_data.groupby(['MainBranch', 'SubBranch'])['Subject'].nunique().reset_index()
        stream_counts['Stream'] = stream_counts['MainBranch'] + " " + stream_counts['SubBranch']
        stream_counts = stream_counts[['Stream', 'Subject']].rename(columns={'Subject': 'Subject Count'}).sort_values('Stream')

        # Download section
        st.markdown("---")
        st.markdown("### 📥 Download Options")

        col1, col2, col3 = st.columns(3)

        with col1:
            # Excel download
            excel_data = save_to_excel(sem_dict)
            if excel_data:
                st.download_button(
                    label="📊 Download Excel File",
                    data=excel_data.getvalue(),
                    file_name=f"timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        with col2:
            # PDF download
            if sem_dict:
                pdf_output = io.BytesIO()
                temp_pdf_path = "temp_timetable.pdf"
                generate_pdf_timetable(sem_dict, temp_pdf_path)
                with open(temp_pdf_path, "rb") as f:
                    pdf_output.write(f.read())
                pdf_output.seek(0)
                if os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)
                st.download_button(
                    label="📄 Download PDF File",
                    data=pdf_output.getvalue(),
                    file_name=f"timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

        with col3:
            if st.button("🔄 Generate New Timetable", use_container_width=True):
                if hasattr(st.session_state, 'processing_complete'):
                    del st.session_state.processing_complete
                if hasattr(st.session_state, 'timetable_data'):
                    del st.session_state.timetable_data
                st.rerun()

        st.markdown("""
        <div class="stats-section">
            <h2>📈 Statistics Overview</h2>
        </div>
        """, unsafe_allow_html=True)

        # Display basic stats in cards
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="metric-card"><h3>📝 {total_exams}</h3><p>Total Exams</p></div>',
                        unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><h3>🎓 {total_semesters}</h3><p>Semesters</p></div>',
                        unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-card"><h3>🏫 {total_branches}</h3><p>Branches</p></div>',
                        unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-card"><h3>📅 {overall_date_range}</h3><p>Days Span</p></div>',
                        unsafe_allow_html=True)

        # Display new stats in cards
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f'<div class="metric-card"><h3>📆 {non_elective_range}</h3><p>Non-Elective Exam Range</p></div>',
                        unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><h3>📆 {elective_dates_str}</h3><p>Elective Exam Dates</p></div>',
                        unsafe_allow_html=True)

        # Display subjects per stream using st.dataframe
        st.markdown("#### Subjects Per Stream")
        if not stream_counts.empty:
            st.dataframe(stream_counts, use_container_width=True)
        else:
            st.markdown('<div class="status-info">ℹ️ No stream data available.</div>', unsafe_allow_html=True)

        st.markdown("---")

        # Results section
        st.markdown("""
        <div class="results-section">
            <h2>📊 Timetable Results</h2>
        </div>
        """, unsafe_allow_html=True)

        # Timetable display
        for sem, df_sem in sem_dict.items():
            st.markdown(f"### 📚 Semester {sem}")

            for main_branch in df_sem["MainBranch"].unique():
                main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
                df_mb = df_sem[df_sem["MainBranch"] == main_branch].copy()

                # Create display format with difficulty labels
                has_oe = df_mb['OE'].notna() & (df_mb['OE'].str.strip() != "")
                difficulty_str = df_mb['Difficulty'].map({0: 'Easy', 1: 'Difficult'}).fillna('')
                difficulty_suffix = difficulty_str.apply(lambda x: f" ({x})" if x else '')
                df_mb["SubjectDisplay"] = df_mb["Subject"]
                df_mb.loc[has_oe, "SubjectDisplay"] = (
                        df_mb.loc[has_oe, "Subject"] + " [" + df_mb.loc[has_oe, "OE"] + "]"
                )
                df_mb["SubjectDisplay"] = df_mb["SubjectDisplay"] + difficulty_suffix

                df_mb["Exam Date"] = pd.to_datetime(df_mb["Exam Date"], format="%d-%m-%Y", errors='coerce')
                df_mb = df_mb.sort_values(by="Exam Date", ascending=True)

                # Create pivot table
                pivot_df = df_mb.pivot_table(
                    index=["Exam Date", "Time Slot"],
                    columns="SubBranch",
                    values="SubjectDisplay",
                    aggfunc=lambda x: ", ".join(x)
                ).fillna("")

                if not pivot_df.empty:
                    st.markdown(f"#### {main_branch_full}")

                    formatted_pivot = pivot_df.copy()
                    if len(formatted_pivot.index.levels) > 0:
                        formatted_dates = [d.strftime("%d-%m-%Y") if pd.notna(d) else "" for d in
                                           formatted_pivot.index.levels[0]]
                        formatted_pivot.index = formatted_pivot.index.set_levels(formatted_dates, level=0)

                    st.dataframe(formatted_pivot, use_container_width=True)



    # Footer
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
