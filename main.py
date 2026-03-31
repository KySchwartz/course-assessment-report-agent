import pandas as pd
from docx import Document
from docx.shared import Inches

def create_aggregate_report(excel_path, output_path):
    # 1. Load Data
    df = pd.read_excel(excel_path)
    total_students = len(df)
    
    # Identify Assignment columns
    assignment_cols = [col for col in df.columns if 'Assignment' in col]
    
    # 2. Initialize Word Doc
    doc = Document()
    doc.add_heading('Course Assessment Summary', 0)
    
    # 3. Create Table (Rows = assignments + 1 for header, Columns = 7)
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    
    # Define Headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'COURSE OBJECTIVE'
    hdr_cells[1].text = 'ACTIVITY'
    hdr_cells[2].text = '# STUDENTS ASSESSED'
    hdr_cells[3].text = '# MET'
    hdr_cells[4].text = '# DID NOT MEET'
    hdr_cells[5].text = '% MET'
    hdr_cells[6].text = 'COMMENTS'
    
    # 4. Process each assignment and add a row
    for i, col_name in enumerate(assignment_cols):
        # Calculate Stats
        met_count = df[col_name].isin(['A', 'B', 'C']).sum()
        not_met_count = total_students - met_count
        percent_met = (met_count / total_students) * 100
        
        # Add a new row to the table
        row_cells = table.add_row().cells
        row_cells[0].text = f"CO-{i+1}" # Placeholder objective
        row_cells[1].text = col_name    # The Assignment Name
        row_cells[2].text = str(total_students)
        row_cells[3].text = str(met_count)
        row_cells[4].text = str(not_met_count)
        row_cells[5].text = f"{percent_met:.0f}%"
        
        # Logic-based comment
        if percent_met >= 70:
            row_cells[6].text = "Most students met the objective. No change needed."
        else:
            row_cells[6].text = "Review material for better understanding."

    # 5. Save
    doc.save(output_path)
    print(f"Report generated: {output_path}")

# Run the function
create_aggregate_report('test_student_data.xlsx', 'Assessment_Report.docx')