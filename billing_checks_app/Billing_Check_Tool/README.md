# Billing Check Tool

## Overview
The Billing Check Tool is a Python application designed to process billing data from CSV files and generate a consolidated Excel report. It provides a user-friendly GUI for selecting input files and configuring processing options.

## Project Structure
```
Billing_Check_Tool
├── src
│   ├── main.py          # Entry point for the application
│   ├── gui.py           # GUI implementation using tkinter
│   ├── processor.py      # Logic for processing CSV files
│   └── utils
│       └── __init__.py  # Initialization file for utils package
├── requirements.txt      # List of dependencies
├── .gitignore            # Files and directories to ignore by Git
└── README.md             # Documentation for the project
```

## Installation
1. Clone the repository:
   ```
   git clone <repository-url>
   ```
2. Navigate to the project directory:
   ```
   cd Billing_Check_Tool
   ```
3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage
1. Run the application:
   ```
   python src/main.py
   ```
2. Use the GUI to select the required CSV files for processing:
   - Prior Month CSV
   - Current Month CSV
   - Sales CSV
   - Delayed Billing Client List CSV
   - Client Device Report CSV
3. Select the month and year for processing.
4. Choose the location to save the processed Excel file.
5. Click the "Process" button to generate the report.

## Dependencies
- pandas
- openpyxl
- tkinter

## Contributing
Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.