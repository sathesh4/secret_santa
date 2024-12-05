import pandas as pd
import random
import logging
from typing import List, Dict

logging.basicConfig(filename='secret_santa.log', level=logging.INFO)

class Employee:
    def __init__(self, name: str, email: str):
        self.name = name
        self.email = email

class SecretSanta:
    def __init__(self, employees: List[Employee], previous_assignments: Dict[str, str] = None):
        self.employees = employees
        self.previous_assignments = previous_assignments if previous_assignments else {}
        self.assignments = {}

    def generate_assignments(self, max_attempts=1000):
        for attempt in range(max_attempts):
            try:
                # Reset assignments
                self.assignments = {}
                available_secret_children = self.employees.copy()
                random.shuffle(available_secret_children)

                # Attempt to assign each employee a unique secret child
                for employee in self.employees:
                    available_choices = [child for child in available_secret_children if child.name != employee.name and child.name != self.previous_assignments.get(employee.name)]
                    if not available_choices:
                        raise ValueError("Unable to find valid assignment")
                    assigned_child = random.choice(available_choices)
                    self.assignments[employee] = assigned_child
                    available_secret_children.remove(assigned_child)

                # If all employees have been assigned a unique child, return
                if len(self.assignments) == len(self.employees):
                    logging.info(f"Assignments successfully generated on attempt {attempt + 1}")
                    return  # Successful assignment
            except ValueError:
                # If assignment fails, retry
                continue

        raise Exception("Assignment not possible due to constraints after multiple attempts.")

    def save_assignments_to_excel(self, output_filename: str):
        data = []
        for employee, secret_child in self.assignments.items():
            data.append({
                'Employee_Name': employee.name,
                'Employee_EmailID': employee.email,
                'Secret_Child_Name': secret_child.name,
                'Secret_Child_EmailID': secret_child.email
            })
        df = pd.DataFrame(data)
        df.to_excel(output_filename, index=False)

def read_employee_data(input_filename: str) -> List[Employee]:
    employees = []
    df = pd.read_excel(input_filename)
    for _, row in df.iterrows():
        employees.append(Employee(row['Employee_Name'], row['Employee_EmailID']))
    return employees

def read_previous_assignments(input_filename: str) -> Dict[str, str]:
    previous_assignments = {}
    df = pd.read_excel(input_filename)
    for _, row in df.iterrows():
        previous_assignments[row['Employee_Name']] = row['Secret_Child_Name']
    return previous_assignments

def main():
    employees = read_employee_data('Employee-List.xlsx')
    previous_assignments = read_previous_assignments('Secret-Santa-Game-Result-2023.xlsx')

    secret_santa = SecretSanta(employees, previous_assignments)
    try:
        secret_santa.generate_assignments()
        secret_santa.save_assignments_to_excel('Secret-Santa-Assignments-2024.xlsx')
        logging.info("Assignments generated successfully!")
        print("Assignments generated successfully!")
    except Exception as e:
        logging.error(f"Error during assignment generation: {e}")
        print(f"Error during assignment generation: {e}")

if __name__ == "__main__":
    main()
