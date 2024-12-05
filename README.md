# Secret Santa Automation Project

## Description
This project automates the process of assigning Secret Santa pairings for employees in a company. The Secret Santa event involves each employee being anonymously assigned another employee to whom they give a gift. The program ensures that the assignments adhere to certain constraints, such as avoiding self-assignments and ensuring that employees do not receive the same secret child as the previous year.

The solution is implemented in Python, utilizing object-oriented principles for modularity and reusability. It handles employee data from Excel files and produces a new Excel file with the assigned pairings.

## Features
- Reads employee information from an Excel file (`Employee-List.xlsx`).
- Reads last year's Secret Santa assignments (`Secret-Santa-Game-Result-2023.xlsx`) to ensure no repeats.
- Assigns each employee a unique Secret Santa partner.
- Handles constraints such as avoiding self-assignments and ensuring no repeated assignments from last year.
- Outputs the new Secret Santa assignments to an Excel file (`Secret-Santa-Assignments-2024.xlsx`).
- Includes error handling and logging for transparency in the assignment generation process.

## Requirements
- Python 3.x
- Pandas library for handling Excel files
- OpenPyXL library for Excel file reading and writing

## Installation
To install the required dependencies, run the following command:

```sh
pip install pandas openpyxl
```

## Usage
1. Prepare the input files:
   - `Employee-List.xlsx`: Contains employee data with columns `Employee_Name` and `Employee_EmailID`.
   - `Secret-Santa-Game-Result-2023.xlsx`: Contains last year's assignments with columns `Employee_Name`, `Employee_EmailID`, `Secret_Child_Name`, and `Secret_Child_EmailID`.

2. Run the script:

```sh
python secret_santa.py
```

3. The output file (`Secret-Santa-Assignments-2024.xlsx`) will be generated with the new assignments.

## Project Structure
- **secret_santa.py**: The main script that generates the Secret Santa assignments.
- **Employee-List.xlsx**: Input file containing employee data.
- **Secret-Santa-Game-Result-2023.xlsx**: Input file containing last year's Secret Santa assignments.
- **Secret-Santa-Assignments-2024.xlsx**: Output file with the new Secret Santa pairings.
- **secret_santa.log**: Log file capturing the assignment generation process and any errors encountered.

## Logging
A log file (`secret_santa.log`) is created to track the status of the assignment generation process, including successful attempts and any errors that occur.

## Error Handling
The script includes error handling for scenarios where the assignments cannot be generated due to constraints (e.g., no valid child left for an employee). If the assignment generation fails after multiple attempts, an exception is raised, and the error is logged.

## Contributing
If you would like to contribute to this project, please feel free to submit a pull request or raise an issue on GitHub.

## License
This project is licensed under the MIT License.

