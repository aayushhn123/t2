import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
from fpdf import FPDF
import io
import os
import re

# -- Page config and CSS --
st.set_page_config(
    page_title="Exam Timetable Generator",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    /* Custom styling (same as original) */
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

# -- Helper: wrap text for PDF table --
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

# -- Helper: print row with uniform height --
def print_row(pdf, row, col_widths, line_h=6):
    # wrap each cell
    wrapped = [wrap_text(pdf, str(cell), w) for cell, w in zip(row, col_widths)]
    max_lines = max(len(lines) for lines in wrapped)
    total_h = line_h * max_lines
    for i, lines in enumerate(wrapped):
        x = pdf.get_x(); y = pdf.get_y()
        for line in lines:
            pdf.multi_cell(col_widths[i], line_h, line, border=0, align='C')
            pdf.set_xy(pdf.get_x() - col_widths[i], pdf.get_y())
        pdf.set_xy(x + col_widths[i], y)
    pdf.ln(total_h)

# -- New Scheduling Logic --
def schedule_stream_exams(df_stream, holidays, start_date, window_days=15):
    df = df_stream.copy().reset_index(drop=True)
    df['Exam Date'] = ''
    df['Time Slot'] = ''
    df = df.sort_values(['Semester','SubBranch','Subject']).reset_index(drop=True)
    window_end = start_date + timedelta(days=window_days - 1)
    current = start_date
    for idx, row in df.iterrows():
        while True:
            if current.date() > window_end.date():
                st.error(f"Cannot finish {row['Branch']} in {window_days} days.")
                return df
            if current.weekday() == 6 or current in holidays:
                current += timedelta(days=1)
                continue
            date_str = current.strftime("%d-%m-%Y")
            if date_str in df['Exam Date'].values:
                current += timedelta(days=1)
                continue
            break
        df.at[idx,'Exam Date'] = date_str
        df.at[idx,'Time Slot'] = row.get('Preferred Slot','10:00 AM - 1:00 PM')
        current += timedelta(days=1)
    return df

def process_constraints_streamwise(df_all, holidays, base_date):
    result = []
    for branch in sorted(df_all['Branch'].unique()):
        sub = df_all[df_all['Branch']==branch]
        sched = schedule_stream_exams(sub, holidays, base_date)
        result.append(sched)
    return pd.concat(result, ignore_index=True)

# -- PDF Generation --
def generate_pdf(df):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font('Arial','B',16)
    pdf.cell(0,10,'Exam Timetable',0,1,'C')
    # table header
    pdf.set_font('Arial','B',12)
    cols = ['Branch','Semester','SubBranch','Subject','Exam Date','Time Slot']
    widths = [40,30,40,80,40,40]
    for col,w in zip(cols,widths): pdf.cell(w,10,col,1,0,'C')
    pdf.ln()
    # rows
    pdf.set_font('Arial','',11)
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
    uploaded = st.file_uploader('Upload Exam Details Excel', type=['xlsx','xls'])
    if not uploaded:
        st.info('Awaiting file upload...')
        return
    df = pd.read_excel(uploaded, engine='openpyxl')
    # ensure datetime conversions
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    # Sidebar inputs
    st.sidebar.header('Configuration')
    base_date = st.sidebar.date_input('Exam Start Date', value=date.today())
    hols = st.sidebar.text_area('Custom Holidays (comma-separated YYYY-MM-DD)')
    holidays = set()
    for s in hols.split(','):
        try: holidays.add(datetime.strptime(s.strip(), '%Y-%m-%d').date())
        except: pass
    # Run scheduling
    try:
        scheduled = process_constraints_streamwise(df, holidays, base_date)
    except Exception as e:
        st.error(str(e))
        return
    st.success('Scheduling complete!')
    st.dataframe(scheduled)
    # Download buttons
    pdf_bytes = generate_pdf(scheduled)
    st.download_button('Download PDF', data=pdf_bytes, file_name='Exam_Timetable.pdf', mime='application/pdf')
    xlsx_bytes = generate_excel(scheduled)
    st.download_button('Download Excel', data=xlsx_bytes, file_name='Exam_Timetable.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__=='__main__':
    main()
