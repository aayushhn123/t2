import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
import re

# Set page configuration
st.set_page_config(
    page_title="Re-Exam Timetable Generator",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
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
        .status-warning {
            background: #fff3cd;
            color: #856404;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #ffc107;
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
        .status-warning {
            background: #5c4b2d;
            color: #fff3cd;
            padding: 1rem;
            border-radius: 5px;
            border-left: 4px solid #ffca28;
        }
    }
</style>
""", unsafe_allow_html=True)

# Constants
BRANCH_FULL_FORM = {
    "B TECH": "BACHELOR OF TECHNOLOGY",
    "B TECH INTG": "BACHELOR OF TECHNOLOGY SIX YEAR INTEGRATED PROGRAM",
    "M TECH": "MASTER OF TECHNOLOGY",
    "MBA TECH": "MASTER OF BUSINESS ADMINISTRATION IN TECHNOLOGY MANAGEMENT",
    "MCA": "MASTER OF COMPUTER APPLICATIONS"
}
LOGO_PATH = "logo.png"  # Adjust as needed

# Data reading functions
def read_re_exam_data(file):
    df = pd.read_excel(file)
    def convert_sem(sem):
        if pd.isna(sem):
            return None
        m = {
            "Semester I": 1, "Semester II": 2, "Semester III": 3, "Semester IV": 4,
            "Semester V": 5, "Semester VI": 6, "Semester VII": 7, "Semester VIII": 8,
            "Semester IX": 9, "Semester X": 10, "Semester XI": 11
        }
        return m.get(sem.strip(), None)
    df["Semester"] = df["Semester"].apply(convert_sem)
    # Fixed: Handle Series properly
    df["Academic Year"] = df["Academic Year"].str.extract(r'(\d{4}-\d{4})').iloc[:, 0].apply(
        lambda x: str(int(x.split('-')[1])) if pd.notna(x) else None
    )
    df["Semester"] = pd.to_numeric(df["Semester"], errors='coerce')
    df["Subject name"] = df["Subject name"].astype(str).str.strip()
    return df

def read_verification_data(file):
    df = pd.read_excel(file)
    df["Semester"] = pd.to_numeric(df["Semester"], errors='coerce')
    df["Current Academic Year"] = df["Current Academic Year"].astype(str)
    df["SubjectName"] = df["SubjectName"].astype(str).str.strip()
    return df

# Fixed matching logic
def match_subjects(re_exam_df, verification_df):
    matched_data = []
    unmatched_subjects = []
    for _, re_row in re_exam_df.iterrows():
        subject = re_row['Subject name']
        semester = re_row['Semester']
        academic_year = re_row['Academic Year']
        
        # Skip if any key field is NaN
        if pd.isna(semester) or pd.isna(academic_year) or pd.isna(subject):
            unmatched_subjects.append(re_row.to_dict())
            continue
        
        # Fixed: Properly handle Series filtering
        mask = (
            (verification_df['Semester'] == semester) &
            (verification_df['Current Academic Year'] == str(academic_year)) &
            (verification_df['SubjectName'] == subject)
        )
        potential_matches = verification_df[mask]
        
        if len(potential_matches) == 0:  # Fixed: Use len() instead of .empty
            unmatched_subjects.append(re_row.to_dict())
        else:
            matched_row = potential_matches.iloc[0]
            matched_data.append({
                'Campus': matched_row['Campus'],
                'Program': re_row['Program'],
                'Stream': re_row['Stream'],
                'Branch': f"{re_row['Program']}-{re_row['Stream']}",
                'Semester': semester,
                'Subject': f"{matched_row['SubjectName']} - ({matched_row['ModuleCode']})",
                'SubjectName': matched_row['SubjectName'],
                'ModuleCode': matched_row['ModuleCode'],
                'Exam Date': matched_row['Exam Date'],
                'Exam Time': matched_row['Exam Time'],
                'Time Slot': f"{matched_row['Exam Time'].split(' to ')[0]} - {matched_row['Exam Time'].split(' to ')[1].replace('pm', 'PM').replace('am', 'AM')}",
                'Exam Duration': matched_row['Exam Duration'],
                'Is Common': matched_row['Is Common'],
                'Academic Year': academic_year,
                'SubjectDisplayPDF': f"{matched_row['SubjectName']} ({academic_year})"
            })
    return pd.DataFrame(matched_data), unmatched_subjects

def split_br(b):
    p = b.split("-", 1)
    return pd.Series([p[0].strip(), p[1].strip() if len(p) > 1 else ""])

# Reused functions from final exam generator
def wrap_text(pdf, text, col_width):
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
    footer_height = 25
    pdf.set_xy(10, pdf.h - footer_height)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 5, "Controller of Examinations", 0, 1, 'L')
    pdf.line(10, pdf.h - footer_height + 5, 60, pdf.h - footer_height + 5)
    pdf.set_font("Arial", size=13)
    pdf.set_xy(10, pdf.h - footer_height + 7)
    pdf.cell(0, 5, "Signature", 0, 1, 'L')
    pdf.set_font("Arial", size=14)
    page_text = f"{pdf.page_no()} of {{nb}}"
    text_width = pdf.get_string_width(page_text.replace("{nb}", "99"))
    pdf.set_xy(pdf.w - 10 - text_width, pdf.h - footer_height + 12)
    pdf.cell(text_width, 5, page_text, 0, 0, 'R')
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
    available_height = pdf.h - pdf.t_margin - footer_height - header_height
    pdf.set_font("Arial", size=12)
    print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
    for idx in range(len(df)):
        row = [str(df.iloc[idx][c]) if pd.notna(df.iloc[idx][c]) else "" for c in columns]
        # Fixed: Check if any cell is non-empty properly
        if not any(cell.strip() for cell in row if isinstance(cell, str)):
            continue
        wrapped_cells = []
        max_lines = 0
        for i, cell_text in enumerate(row):
            text = str(cell_text) if cell_text is not None else ""
            avail_w = col_widths[i] - 2 * 2
            lines = wrap_text(pdf, text, avail_w)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_h = line_height * max_lines
        if pdf.get_y() + row_h > pdf.h - footer_height:
            pdf.add_page()
            print_row_custom(pdf, columns, col_widths, line_height=line_height, header=True)
        print_row_custom(pdf, row, col_widths, line_height=line_height, header=False)

def convert_excel_to_pdf(excel_path, pdf_path, sub_branch_cols_per_page=4):
    pdf = FPDF(orientation='L', unit='mm', format=(210, 500))
    pdf.set_auto_page_break(auto=False, margin=15)
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
    for sheet_name, pivot_df in df_dict.items():
        if pivot_df.empty:
            continue
        parts = sheet_name.split('_Sem_')
        main_branch = parts[0]
        main_branch_full = BRANCH_FULL_FORM.get(main_branch, main_branch)
        semester = parts[1] if len(parts) > 1 else ""
        semester_roman = semester if not semester.isdigit() else int_to_roman(int(semester))
        header_content = {'main_branch_full': main_branch_full, 'semester_roman': semester_roman}
        pivot_df = pivot_df.reset_index().dropna(how='all', axis=0).reset_index(drop=True)
        fixed_cols = ["Exam Date", "Time Slot"]
        sub_branch_cols = [c for c in pivot_df.columns if c not in fixed_cols]
        exam_date_width = 60
        for start in range(0, len(sub_branch_cols), sub_branch_cols_per_page):
            chunk = sub_branch_cols[start:start + sub_branch_cols_per_page]
            cols_to_print = fixed_cols[:1] + chunk
            chunk_df = pivot_df[fixed_cols + chunk].copy()
            # Fixed: Handle DataFrame boolean checking properly
            mask = chunk_df[chunk].apply(lambda row: (row.astype(str).str.strip() != "").any(), axis=1)
            chunk_df = chunk_df[mask].reset_index(drop=True)
            if chunk_df.empty:
                continue
            time_slot = pivot_df['Time Slot'].iloc[0] if 'Time Slot' in pivot_df.columns and len(pivot_df['Time Slot']) > 0 else None
            chunk_df["Exam Date"] = pd.to_datetime(chunk_df["Exam Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%A, %d %B, %Y")
            page_width = pdf.w - 2 * pdf.l_margin
            remaining = page_width - exam_date_width
            sub_width = remaining / max(len(chunk), 1)
            col_widths = [exam_date_width] + [sub_width] * len(chunk)
            total_w = sum(col_widths)
            if total_w > page_width:
                factor = page_width / total_w
                col_widths = [w * factor for w in col_widths]
            pdf.add_page()
            print_table_custom(pdf, chunk_df[cols_to_print], cols_to_print, col_widths, line_height=10,
                             header_content=header_content, branches=chunk, time_slot=time_slot)
    pdf.output(pdf_path)

def generate_pdf_timetable(semester_wise_timetable, output_pdf):
    temp_excel = os.path.join(os.path.dirname(output_pdf), "temp_re_exam_timetable.xlsx")
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
                df_mb["SubjectDisplay"] = df_mb["SubjectDisplayPDF"] + df_mb["Exam Duration"].apply(
                    lambda x: f" [Duration: {x} hrs]" if x != 3 else ''
                )
                df_mb["Exam Date"] = pd.to_datetime(df_mb["Exam Date"], format="%d-%m-%Y", errors='coerce')
                df_mb = df_mb.sort_values(by="Exam Date", ascending=True)
                pivot_df = df_mb.pivot_table(
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
    output.seek(0)
    return output

def save_verification_excel(matched_df):
    verification_columns = [
        "Campus", "Program", "Stream", "Academic Year", "Semester",
        "ModuleCode", "SubjectName", "Exam Date", "Exam Time", "Is Common"
    ]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        matched_df[verification_columns].to_excel(writer, sheet_name="Re-Exam Verification", index=False)
    output.seek(0)
    return output

def main():
    st.markdown("""
    <div class="main-header">
        <h1>📅 Re-Exam Timetable Generator</h1>
        <p>MUKESH PATEL SCHOOL OF TECHNOLOGY MANAGEMENT & ENGINEERING</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("""
        <div class="upload-section">
            <h3>📁 Upload Re-Exam Data Excel File</h3>
            <p>Upload the re-exam data file (.xlsx format)</p>
        </div>
        """, unsafe_allow_html=True)
        re_exam_file = st.file_uploader("Choose Re-Exam Data Excel file", type=['xlsx'])

        st.markdown("""
        <div class="upload-section">
            <h3>📁 Upload Final Exam Verification Excel File</h3>
            <p>Upload the final exam verification file (.xlsx format)</p>
        </div>
        """, unsafe_allow_html=True)
        verification_file = st.file_uploader("Choose Final Exam Verification Excel file", type=['xlsx'])

    if re_exam_file and verification_file:
        if st.button("🔄 Generate Re-Exam Timetable", type="primary", use_container_width=True):
            with st.spinner("Processing your re-exam timetable... Please wait..."):
                try:
                    re_exam_df = read_re_exam_data(re_exam_file)
                    verification_df = read_verification_data(verification_file)
                    matched_df, unmatched = match_subjects(re_exam_df, verification_df)

                    if not matched_df.empty:
                        matched_df[["MainBranch", "SubBranch"]] = matched_df["Branch"].apply(split_br)
                        sem_dict = {}
                        for sem in matched_df["Semester"].unique():
                            sem_dict[sem] = matched_df[matched_df["Semester"] == sem].copy()

                        # Generate outputs
                        excel_data = save_to_excel(sem_dict)
                        pdf_output = io.BytesIO()
                        temp_pdf_path = "temp_re_exam_timetable.pdf"
                        generate_pdf_timetable(sem_dict, temp_pdf_path)
                        with open(temp_pdf_path, "rb") as f:
                            pdf_output.write(f.read())
                        pdf_output.seek(0)
                        if os.path.exists(temp_pdf_path):
                            os.remove(temp_pdf_path)
                        verification_data = save_verification_excel(matched_df)

                        # Store in session state
                        st.session_state.excel_data = excel_data.getvalue()
                        st.session_state.pdf_data = pdf_output.getvalue()
                        st.session_state.verification_data = verification_data.getvalue()

                        # Display timetable
                        display_df = matched_df[['Branch', 'Semester', 'Subject', 'Exam Date', 'Exam Time']]
                        st.dataframe(display_df, use_container_width=True)

                        # Download buttons
                        st.download_button(
                            label="📊 Download Re-Exam Timetable Excel",
                            data=st.session_state.excel_data,
                            file_name=f"re_exam_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.download_button(
                            label="📄 Download Re-Exam Timetable PDF",
                            data=st.session_state.pdf_data,
                            file_name=f"re_exam_timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                            mime="application/pdf"
                        )
                        st.download_button(
                            label="📋 Download Re-Exam Verification Excel",
                            data=st.session_state.verification_data,
                            file_name=f"re_exam_verification_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                        st.markdown('<div class="status-success">🎉 Re-Exam Timetable generated successfully!</div>', unsafe_allow_html=True)

                    if unmatched:
                        st.markdown(f'<div class="status-warning">{len(unmatched)} subjects could not be matched and require manual scheduling.</div>', unsafe_allow_html=True)
                        with st.expander("View unmatched subjects"):
                            st.dataframe(pd.DataFrame(unmatched))

                except Exception as e:
                    st.markdown(f'<div class="status-error">❌ An error occurred: {str(e)}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-warning">Please upload both files to proceed.</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
