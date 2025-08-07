import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
from fpdf import FPDF
import io
import os

# -- Page config and CSS --
st.set_page_config(
    page_title="Exam Timetable Generator",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)
st.markdown("""
<style>
    /* Custom button hover styling */
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
</style>
""", unsafe_allow_html=True)

# -- PDF Helpers --
def wrap_text(pdf, text, col_width):
    words = text.split()
    lines = []
    current = ""
    for w in words:
        test = w if not current else current + " " + w
        if pdf.get_string_width(test) <= col_width:
            current = test
        else:
            lines.append(current)
            current = w
    if current:
        lines.append(current)
    return lines


def print_row(pdf, row, col_widths, line_h=6):
    wrapped = [wrap_text(pdf, str(cell), w) for cell, w in zip(row, col_widths)]
    max_lines = max(len(lines) for lines in wrapped)
    total_h = line_h * max_lines
    for i, lines in enumerate(wrapped):
        x, y = pdf.get_x(), pdf.get_y()
        for line in lines:
            pdf.multi_cell(col_widths[i], line_h, line, border=0, align='C')
            pdf.set_xy(x + col_widths[i], pdf.get_y())
        pdf.set_xy(pdf.get_x() - col_widths[i], y)
    pdf.ln(total_h)

# -- Scheduling Logic --
def schedule_stream_exams(df_stream, holidays, start_date, window_days=15):
    df = df_stream.copy().reset_index(drop=True)
    df['Exam Date'] = ''
    df['Time Slot'] = ''
    df = df.sort_values(['Semester', 'SubBranch', 'Subject']).reset_index(drop=True)
    window_end = start_date + timedelta(days=window_days - 1)
    current = start_date
    for idx, row in df.iterrows():
        while True:
            if current.date() > window_end.date():
                raise RuntimeError(f"Cannot finish {row['Branch']} in {window_days} days.")
            # skip Sundays and custom holidays
            if current.weekday() == 6 or current.date() in holidays:
                current += timedelta(days=1)
                continue
            date_str = current.strftime("%d-%m-%Y")
            # ensure one exam per day per branch
            if date_str in df['Exam Date'].values:
                current += timedelta(days=1)
                continue
            break
        df.at[idx, 'Exam Date'] = date_str
        df.at[idx, 'Time Slot'] = row['Preferred Slot']
        current += timedelta(days=1)
    return df


def process_constraints_streamwise(df_all, holidays, base_date):
    parts = []
    for branch in sorted(df_all['Branch'].unique()):
        df_branch = df_all[df_all['Branch'] == branch]
        scheduled = schedule_stream_exams(df_branch, holidays, base_date)
        parts.append(scheduled)
    return pd.concat(parts, ignore_index=True)

# -- PDF Generation --
def generate_pdf(df):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Exam Timetable', 0, 1, 'C')
    pdf.set_font('Arial', 'B', 12)
    cols = ['Branch', 'Semester', 'SubBranch', 'Subject', 'Exam Date', 'Time Slot']
    widths = [50, 30, 30, 100, 40, 40]
    for col, w in zip(cols, widths):
        pdf.cell(w, 10, col, 1, 0, 'C')
    pdf.ln()
    pdf.set_font('Arial', '', 11)
    for _, row in df[cols].iterrows():
        print_row(pdf, row.tolist(), widths)
    return pdf.output(dest='S').encode('latin-1')

# -- Excel Generation --
def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Timetable')
        writer.save()
    return output.getvalue()

# -- Main App --
def main():
    st.title('Exam Timetable Generator')
    uploaded = st.file_uploader('Upload Exam Details Excel', type=['xlsx', 'xls'])
    if not uploaded:
        st.info('Please upload your exam input file.')
        return

    # Read and map columns
    df_raw = pd.read_excel(uploaded, engine='openpyxl')
    # Create scheduling-specific fields
    df_raw['Branch'] = df_raw['Program'].astype(str) + ' - ' + df_raw['Stream'].astype(str)
    df_raw['Semester'] = df_raw['Current Session']
    df_raw['SubBranch'] = df_raw['Category']
    df_raw['Subject'] = df_raw['Module Description']
    # Generate time slots based on duration (start at 10 AM)
    def slot_from_dur(d):
        start_hr = 10
        end_hr = start_hr + int(d)
        suffix = 'AM' if end_hr < 12 else 'PM'
        end_12 = end_hr if end_hr <= 12 else end_hr - 12
        return f"10:00 AM - {end_12}:00 {suffix}"
    df_raw['Preferred Slot'] = df_raw['Exam Duration'].apply(slot_from_dur)

    # Sidebar: start date & holidays
    st.sidebar.header('Configuration')
    base_date = st.sidebar.date_input('Exam Start Date', value=date.today())
    hol_text = st.sidebar.text_area('Custom Holidays (comma-separated YYYY-MM-DD)')
    holidays = set()
    for part in hol_text.split(','):
        try:
            holidays.add(datetime.strptime(part.strip(), '%Y-%m-%d').date())
        except:
            pass

    # Run scheduling
    try:
        scheduled_df = process_constraints_streamwise(df_raw, holidays, base_date)
    except Exception as e:
        st.error(f"Scheduling Error: {e}")
        return

    st.success('Scheduling complete!')
    st.dataframe(scheduled_df)

    # Download PDF & Excel
    pdf_out = generate_pdf(scheduled_df)
    st.download_button('Download PDF', data=pdf_out,
                       file_name='Exam_Timetable.pdf', mime='application/pdf')
    xlsx_out = generate_excel(scheduled_df)
    st.download_button('Download Excel', data=xlsx_out,
                       file_name='Exam_Timetable.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    main()
