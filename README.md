# Data Processing in Excel

This project focuses on analyzing the relationship between alcohol consumption and study time using Excel and VBA (Visual Basic for Applications). The goal is to process data from an Excel spreadsheet, calculate average alcohol consumption, and present the results graphically.

## Contents

- [Project Description](#project-description)
- [Getting Started](#getting-started)
- [File Structure](#file-structure)

## Project Description

The project analyzes student data to determine the impact of study time on alcohol consumption during weekdays and weekends. Based on the data, the project automatically generates a report and graphical representation of average alcohol consumption across different study time categories.

## Getting Started

1. **Prepare the Data**:
   - Ensure you have an Excel file containing student data.
   - If your data is in CSV format, use the provided Jupyter Notebook to convert the CSV file to Excel.

2. **Convert CSV to Excel**:
   - Open the `csv_to_xlsx.ipynb` file in Jupyter Notebook.
   - Follow the steps in the notebook to convert your CSV file to an Excel Workbook (`.xlsx`).

3. **Open the VBA Editor**:
   - Open your newly created Excel file.
   - Press `Alt + F11` to open the VBA Editor.

4. **Add VBA Code**:
   - You can view the VBA code for the analysis and chart creation in this module.

5. **Run the Macro**:
   - In the VBA Editor, press `F5` to run the macro or return to Excel and run the macro by clicking on "Gumb 1" located below the last row in the Excel sheet.
   - This will execute the VBA code and perform the analysis and chart creation.

6. **Review the Results**:
   - After running the macro, a new worksheet named `AlcoholStudytimeSummary` will be added to your Excel file.
   - This worksheet will contain the results of the analysis in tabular form, as well as a chart showing average alcohol consumption versus study time.

## File Structure

- **Excel File**: Contains a worksheet named `Math` with student data. If necessary, it may also include an additional worksheet named `AlcoholStudytimeSummary` for displaying results.
- **VBA Module**: Contains the macro for data analysis and chart generation.
- **Jupyter Notebook**: Includes a file named `csv_to_xlsx.ipynb` for converting CSV data to an Excel file.





