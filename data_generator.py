import pandas as pd
import random
import os

def generate_dynamic_test_suite(filename="test_student_data.xlsx"):
    # Check if the file is open to avoid PermissionErrors
    if os.path.exists(filename):
        try:
            os.rename(filename, filename)
        except OSError:
            print(f"ERROR: The file '{filename}' is open in another program. Please close it and run the script again.")
            return

    # Define variables
    num_students = 30
    first_names = ["Kyle", "Joe", "Matthew", "Luyanda", "Reshmi", "Nipesh", "Sam", "Marilla", "Ivy", "Larry", "Lilly"]
    last_names = ["Schwartz", "Samples", "Clippard", "Chitofu", "Mitra", "Pant", "Brucker", "Anderson", "DeRousse", "Bartlet", "Stevens"]
    majors = ["CS", "CY", "CIS"]
    grades = ["A", "B", "C", "D", "E", "F"]
    
    # 1. GENERATE STUDENT DATA
    student_records = []
    for i in range(num_students):
        f_name = random.choice(first_names)
        l_name = random.choice(last_names)
        record = {
            "First Name": f_name,
            "Last Name": l_name,
            "SO Number": f"SO{3000 + i}",
            "Email Address": f"s{i}@university.edu",
            "Major": random.choice(majors)
        }
        for a in range(1, 7):
            record[f"Assignment {a}"] = random.choices(grades, weights=[50, 25, 10, 5, 5, 5])[0]
        student_records.append(record)
    
    df_students = pd.DataFrame(student_records)

    # 2. GENERATE OBJECTIVE MAPPING
    mapping_data = [
        {"Objective ID": "CO-1", "Objective Description": "Understanding Cyberspace attacks.", "Assignments": "Assignment 1"},
        {"Objective ID": "CO-2", "Objective Description": "Demonstrate cyber defense techniques.", "Assignments": "Assignment 2, Assignment 3"},
        {"Objective ID": "CO-3", "Objective Description": "Evaluate connected systems.", "Assignments": "Assignment 4"},
        {"Objective ID": "CO-4", "Objective Description": "Tool development and scripting.", "Assignments": "Assignment 5, Assignment 6"}
    ]
    df_mapping = pd.DataFrame(mapping_data)

    # 3. SAVE TO MULTIPLE SHEETS
    # 'openpyxl' is required for multiple sheets
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_students.to_excel(writer, sheet_name='Grades', index=False)
        df_mapping.to_excel(writer, sheet_name='Objectives', index=False)
    
    print(f"Successfully created '{filename}' with sheets: 'Grades' and 'Objectives'.")

if __name__ == "__main__":
    generate_dynamic_test_suite()