import pandas as pd
import random

def generate_test_data(filename="test_student_data.xlsx", num_students=30, num_assignments=5):
    # Lists for synthetic data generation
    first_names = ["James", "Mary", "Robert", "Patricia", "John", "Jennifer", "Michael", "Linda"]
    last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis"]
    majors = ["Computer Science", "Cyber Security", "Information Technology", "Data Science"]
    grades = ["A", "B", "C", "D", "E", "F"]
    
    data = []

    for i in range(num_students):
        f_name = random.choice(first_names)
        l_name = random.choice(last_names)
        so_num = f"SO{100000 + i}"
        email = f"{f_name.lower()}.{l_name.lower()}{i}@university.edu"
        major = random.choice(majors)
        
        # Create the base record
        record = {
            "First Name": f_name,
            "Last Name": l_name,
            "SO Number": so_num,
            "Email Address": email,
            "Major": major
        }
        
        # Dynamically add assignment columns
        for a in range(1, num_assignments + 1):
            # Weighted choice to make 'Met' (A, B, C) more likely for realistic testing
            record[f"Assignment {a}"] = random.choices(grades, weights=[30, 25, 20, 10, 10, 5])[0]
            
        data.append(record)

    # Convert to DataFrame and save to Excel
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Successfully generated {filename} with {num_students} students and {num_assignments} assignments.")

# Run the generator
if __name__ == "__main__":
    generate_test_data()