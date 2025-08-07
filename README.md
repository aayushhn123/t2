# Exam Timetable Schedulers

This repository contains two Streamlit applications designed to generate exam timetables efficiently:

- **Final Exam Timetable Scheduler**: Creates timetables for final exams based on input data.
- **Re-Exam Timetable Scheduler**: Generates re-exam timetables using final exam data and verification.

Both tools are hosted online and provide downloadable outputs in Excel and PDF formats for easy use.

## Links to Applications

- [Final Exam Timetable Scheduler](https://examtimetablescheduler-dqttcyf5vzakkjfpkt6xhp.streamlit.app/)
- [Re-Exam Timetable Scheduler](https://re-examtimetablescheduler-gndknuqn7whtdxe6cvubaw.streamlit.app/)

## Final Exam Timetable Scheduler

### Overview
This tool helps you create a final exam timetable. It takes your input data, schedules exams, and generates an Excel timetable, a PDF timetable, and a verification Excel file.

### Input File
- **File Name**: `Final Exam Input Data.xlsx`
- **Format**: Excel file containing details like subject, semester, branch, and other scheduling preferences.

### Outputs
- **Excel Timetable**: A detailed schedule in Excel format.
- **PDF Timetable**: A printable version of the schedule.
- **Verification Excel File**: A file to check the accuracy of the timetable.

### How to Use
1. **Visit the Link**: Go to [Final Exam Timetable Scheduler](https://examtimetablescheduler-dqttcyf5vzakkjfpkt6xhp.streamlit.app/).
2. **Upload Input File**: Click on "Upload Excel File" and select your `Final Exam Input Data.xlsx` file.
3. **Set Base Date**: Choose the starting date for the exam schedule using the date picker.
4. **Add Holidays**: Enter any holiday dates (e.g., `15-08-2025`) in the "Enter Holidays" box, one per line.
5. **Generate Timetable**: Click the "Generate Timetable" button and wait for the process to complete.
6. **Download Outputs**: Once generated, download the Excel timetable, PDF timetable, and verification file from the app.

## Re-Exam Timetable Scheduler

### Overview
This tool creates a re-exam timetable using the final exam timetable data and a verification file. It ensures re-exams are scheduled without conflicts.

### Input Files
- **Re-Exam Input File**: `Re-exam Input File.xlsx` with re-exam details (e.g., subjects, students).
- **Verification File**: The verification Excel file generated from the final exam timetable.

### Outputs
- **Excel Timetable**: A re-exam schedule in Excel format.
- **PDF Timetable**: A printable re-exam schedule.

### How to Use
1. **Visit the Link**: Go to [Re-Exam Timetable Scheduler](https://re-examtimetablescheduler-gndknuqn7whtdxe6cvubaw.streamlit.app/).
2. **Upload Files**: 
   - Click "Upload Re-Exam Input File" and select `Re-exam Input File.xlsx`.
   - Click "Upload Verification File" and select the verification Excel from the final exam scheduler.
3. **Generate Timetable**: Click the "Generate Timetable" button and wait for the process to complete.
4. **Download Outputs**: Download the Excel and PDF re-exam timetables from the app.

## Requirements
- Python 3.x
- Streamlit
- Pandas
- FPDF
- PyPDF2
- (Install via `pip install streamlit pandas fpdf PyPDF2`)

## Notes
- Ensure input files are in the correct Excel format with required columns (e.g., Subject, Semester, Branch).
- Holidays must be entered in `dd-mm-yyyy` format.
- The apps are optimized for the Mukesh Patel School of Technology Management & Engineering schedule.
