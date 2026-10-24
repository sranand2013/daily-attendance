An `IndentationError` in Python means the spacing or tabs at the beginning of the line got mixed up when the code was copied and pasted. Since Python relies heavily on exact spacing, even one extra space can cause this error.

To fix this completely, please **delete all the text** in your `hourlyReportV3.py` file, and copy-paste the entire block below using the "Copy" button at the top right of the code box. This ensures no hidden spaces or tabs are accidentally added.

Here is the clean, correctly indented code:

```python
import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="CMS Attendance Generator", layout="wide", page_icon="📊")

# Customizing the UI styling
st.markdown("""
    <style>
    .main-header { font-size: 2.5rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0;}
    .sub-header { font-size: 1.2rem; color: #4B5563; margin-bottom: 2rem;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">JAIN CMS - Daily Attendance Report Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload Attendance, Timetable, and Course Matrix files to generate the structured hourly summary.</p>', unsafe_allow_html=True)

with st.sidebar:
    st.header("1. Input Files")
    att_files = st.file_uploader("Upload Attendance Files (.xlsx)", accept_multiple_files=True, help="e.g., bai-a.xlsx, corp-b.xlsx")
    tt_files = st.file_uploader("Upload Time Table Files (.xlsx)", accept_multiple_files=True, help="e.g., bai-tt.xlsx, corp-tt.xlsx")
    course_matrix = st.file_uploader("Upload Course Matrix (.xlsx)", accept_multiple_files=False, help="e.g., coursematrix1sem.xlsx")
    
    st.header("2. Report Details")
    date_input = st.text_input("Date (DD-MM-YYYY)", "16-10-2026")
    out_filename = st.text_input("Output Filename", "Attendance_Report_Output.xlsx")
    
    process_btn = st.button("Generate Report", type="primary", use_container_width=True)

if process_btn:
    if not att_files or not tt_files or not course_matrix:
        st.error("⚠️ Please upload the Attendance files, Time Table files, AND the Course Matrix file.")
    else:
        with st.spinner('Processing files & generating report...'):
            try:
                # ---------------------------------------------------------
                # STEP 0: Process Course Matrix
                # ---------------------------------------------------------
                course_map = {}
                xls_matrix = pd.ExcelFile(course_matrix)
                for sheet in xls_matrix.sheet_names:
                    df_cm = pd.read_excel(xls_matrix, sheet_name=sheet)
                    if 'Course Code' in df_cm.columns and 'Course Name' in df_cm.columns:
                        df_cm_clean = df_cm.dropna(subset=['Course Code', 'Course Name'])
                        for _, row in df_cm_clean.iterrows():
                            code = str(row['Course Code']).strip()
                            name = str(row['Course Name']).strip()
                            course_map[code] = name

                # ---------------------------------------------------------
                # STEP 1: Process Timetables
                # ---------------------------------------------------------
                tt_dict = {} 
                hour_headers = [] 
                
                for ttf in tt_files:
                    spec = ttf.name.lower().replace('-tt.xlsx', '').replace('.xlsx', '').strip()
                    df_tt = pd.read_excel(ttf, sheet_name='Cumulative', header=4)
                    
                    df_tt['Date'] = df_tt['Date'].ffill().astype(str)
                    df_tt_date = df_tt[df_tt['Date'].str.contains(date_input)]
                    
                    current_tt_hours = [str(c) for c in df_tt.columns if 'HOUR' in str(c).upper()]
                    if not hour_headers and current_tt_hours:
                        hour_headers = current_tt_hours
                    
                    for _, row in df_tt_date.iterrows():
                        sec = str(row['Section']).strip().upper()
                        schedule = {}
                        for col in current_tt_hours:
                            val = str(row[col]).replace('\n', '').strip()
                            if val != 'nan' and val != '':
                                schedule[col] = val
                        tt_dict[(spec, sec)] = schedule

                if not hour_headers:
                    hour_headers = [f"Hour {i}" for i in range(1, 6)]

                # ---------------------------------------------------------
                # STEP 2: Process Attendance Files & Map to Timetable
                # ---------------------------------------------------------
                report_data = []
                sl_no = 1
                all_output_cols = ['Sl.No.', 'Specialization', 'SEC', 'ROOM', 'STRENGTH'] + hour_headers
                
                for attf in att_files:
                    base = attf.name.lower().replace('.xlsx', '')
                    if '-' in base:
                        parts = base.rsplit('-', 1)
                        spec = parts[0].strip()
                        sec = parts[1].strip().upper()
                    else:
                        spec = base[:-1] if len(base) > 1 else base
                        sec = base[-1].upper() if len(base) > 0 else ''
                        
                    df_att = pd.read_excel(attf, sheet_name=0, skiprows=4)
                    df_att.columns = [str(c).strip() for c in df_att.columns]
                    
                    student_col = next((c for c in df_att.columns if 'NAME' in c.upper() or 'STUDENT' in c.upper()), None)
                    if student_col:
                        strength = len(df_att[student_col].dropna())
                    else:
                        strength = len(df_att) - 1 if len(df_att) > 0 else 0
                    if strength < 0: strength = 0
                    
                    subject_cols = []
                    ignore_cols = ['S NO', 'REG/USN', 'STUDENT NAME', 'SECTION', 'OVERALL %', 'OVERALL', 'SIGNATURE']
                    
                    for i, col in enumerate(df_att.columns):
                        if 'UNNAMED' not in col.upper() and not any(ign == col.upper() for ign in ignore_cols):
                            if i + 1 < len(df_att.columns) and str(df_att.iloc[0, i+1]).strip().upper() == 'A':
                                subject_cols.append((i, col))
                                
                    att_counts = {}
                    for i, col in subject_cols:
                        code = str(col).split(' - ')[0].strip()
                        attended_series = pd.to_numeric(df_att.iloc[1:, i+1], errors='coerce').fillna(0)
                        att_counts[code] = int((attended_series > 0).sum())
                        
                    schedule = tt_dict.get((spec, sec), {})
                    
                    row1 = {c: '' for c in all_output_cols}
                    row2 = {c: '' for c in all_output_cols}
                    
                    row1.update({'Sl.No.': sl_no, 'Specialization': spec, 'SEC': sec, 'ROOM': '', 'STRENGTH': strength})
                    row2.update({'Sl.No.': '', 'Specialization': f"[{spec} - {sec}] Subject Code / Name", 'SEC': '', 'ROOM': '', 'STRENGTH': ''})
                    
                    for h_key in hour_headers:
                        if h_key in schedule:
                            code_from_tt = schedule[h_key]
                            
                            matched_code = None
                            for ac in att_counts.keys():
                                if ac == code_from_tt or ac in code_from_tt or code_from_tt in ac:
                                    matched_code = ac
                                    break
                                    
                            if matched_code:
                                row1[h_key] = att_counts[matched_code]
                                c_name = course_map.get(code_from_tt, course_map.get(matched_code, "Unknown Subject"))
                                row2[h_key] = f"{code_from_tt} - {c_name}"
                            else:
                                row1[h_key] = 0
                                c_name = course_map.get(code_from_tt, "Unrecorded/Unknown")
                                row2[h_key] = f"{code_from_tt} - {c_name}"
                        else:
                            row1[h_key] = 0
                            row2[h_key] = '-'
                        
                    report_data.append(row1)
                    report_data.append(row2)
                    sl_no += 1
                
                # Calculate Totals & Percentages for the Bottom Row
                total_strength = sum(r['STRENGTH'] for r in report_data if isinstance(r.get('STRENGTH'), (int, float)))
                total_row = {c: '' for c in all_output_cols}
                total_row['Sl.No.'] = 'Total'
                total_row['STRENGTH'] = total_strength
                
                for h_key in hour_headers:
                    total_hour_att = sum(r[h_key] for r in report_data if isinstance(r.get(h_key), (int, float)))
                    if total_strength > 0:
                        perc = (total_hour_att / total_strength) * 100
                        total_row[h_key] = f"{total_hour_att} ({perc:.1f}%)"
                    else:
                        total_row[h_key] = f"{total_hour_att} (0.0%)"
                        
                report_data.append(total_row)
                
                df_out = pd.DataFrame(report_data, columns=all_output_cols)
                df_out.columns = [c.replace('\n', ' ') for c in df_out.columns]
                
                # ---------------------------------------------------------
                # STEP 3: Generate Excel File and Format Beautifully
                # ---------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet('Hourly Attendance')
                    
                    bold_format = workbook.add_format({'bold': True, 'font_size': 12})
                    header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D3D3D3', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                    cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                    text_left_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
                    footer_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#E8E8E8', 'align': 'center', 'valign': 'vcenter'})
                    
                    worksheet.write(0, 0, "JAIN - CENTER FOR MANAGEMENT STUDIES", bold_format)
                    
                    try:
                        date_obj = datetime.strptime(date_input, "%d-%m-%Y")
                        day_name = date_obj.strftime("%A")
                    except:
                        day_name = ""
                        
                    worksheet.write(1, 0, f"Attendance Report - 1 Sem BBA ({date_input} - {day_name})", bold_format)
                    
                    for col_num, value in enumerate(df_out.columns.values):
                        worksheet.write(4, col_num, value, header_format)
                        
                    for row_num, row_data in enumerate(df_out.values):
                        is_last_row = (row_num == len(df_out) - 1)
                        for col_num, value in enumerate(row_data):
                            if is_last_row:
                                fmt = footer_format
                            else:
                                fmt = text_left_format if col_num == 1 else cell_format
                            worksheet.write(row_num + 5, col_num, value if pd.notna(value) else "", fmt)
                    
                    worksheet.set_column('A:A', 6)
                    worksheet.set_column('B:B', 50)  
                    worksheet.set_column('C:E', 12)
                    worksheet.set_column(5, len(df_out.columns) - 1, 24) 
                    worksheet.set_row(4, 30) 
                        
                output.seek(0)
                
                st.success(f"✅ Successfully compiled {len(report_data)//2} section records with correctly counted students!")
                
                st.download_button(
                    label="⬇️ Download Excel Output",
                    data=output,
                    file_name=out_filename if out_filename.endswith('.xlsx') else f"{out_filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.markdown("### Output Preview")
                st.dataframe(df_out.tail(10), use_container_width=True)
                
            except Exception as e:
                import traceback
                st.error(f"❌ An error occurred during processing: {str(e)}")
                with st.expander("Show Detailed Error"):
                    st.code(traceback.format_exc())

```
