# Automation of GUI(TacticalGISTest Application)

# Coordinate Conversion & Line of Sight (LOS) Automation

## Project Overview
This project focuses on automating GUI interactions with the TacticalGis Test Application using Python and PyAutoGUI. The automation processes coordinate conversions and line-of-sight (LOS) calculations by interacting with the application’s UI, reading input data from CSV files, and writing processed results to output CSV files.

## Key Features
- Automates GUI interactions for coordinate conversion and LOS calculations.
- Reads input data from CSV files and processes it in the TacticalGis Test Application.
- Copies results from the application and saves them in output CSV files.
- Uses XML configurations to define application paths, workspace files, and input/output CSV paths.
- Implements real-time validation to ensure data accuracy.

## Technologies Used
- **Python** – Primary language for automation scripting.
- **PyAutoGUI** – GUI automation for keyboard and mouse interactions.
- **Pyperclip** – Clipboard operations for copying values.
- **PyGetWindow** – Window management and activation.
- **XML Parsing (ElementTree)** – Reads application configuration settings.
- **CSV Module** – Reads and writes input/output data.
- **OS & Subprocess** – File handling and application execution.

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo-url.git
   cd your-repo
   ```
2. Install required dependencies:
   ```bash
   pip install pyautogui pyperclip pygetwindow
   ```
3. Ensure the TacticalGis Test Application is installed and accessible.
4. Configure XML file paths accordingly.

## Usage
### 1. Coordinate Conversion Automation
- Reads input coordinates from a CSV file.
- Navigates the application’s UI to input values.
- Copies and validates converted coordinates.
- Writes processed values to an output CSV file.

### 2. Line of Sight (LOS) Calculation Automation
- Reads user-defined values and coordinates from CSV.
- Automates vector layer addition and workspace selection.
- Copies LOS calculations from the application.
- Writes LOS validation results to an output CSV file.

## Execution
Run the automation script:
```bash
python coordinate_conversion.py  # For coordinate conversion
python los_automation.py         # For LOS calculations
python rawFormats.py
```

## Skills & Concepts Used
- GUI Automation & Testing
- Data Parsing & Validation
- File Handling (CSV & XML)
- Process Management & Window Activation
- Error Handling & Debugging

## Future Enhancements
- Enhance exception handling for improved stability.
- Implement a GUI-based dashboard for monitoring automation progress.
- Extend support for additional GIS-based automation tasks.


## Demo
- Automation of Coordinate_Conversion
[Watch the Video](https://github.com/Princek469/Automation_TacticalGISTester_Application/raw/main/coordinate_conversion.mp4)

- Automation of LOS
[Watch the Video](https://github.com/Princek469/Automation_TacticalGISTester_Application/raw/main/LOS.mp4)


