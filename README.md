# School Power BI Data Viewer

A web application that displays and visualizes data from the SchoolPowerBIData.xlsx Excel file.

## Features

- View different sheets in the Excel file in a responsive table format
- Advanced visualization for the "InternalExam" sheet data
- Filter students by year, class, and number
- Interactive chart visualization showing student rank progression across different subjects and assessments
- Chart features:
  - Toggle Y-axis orientation
  - Zoom and pan capabilities
  - Show/hide specific subjects
  - Detailed tooltips showing rank and score information

## How to Use

1. Open the index.html file in a web browser
2. Select a sheet from the dropdown to view its data
3. For the "InternalExam" sheet:
   - Use the filters to narrow down student selection
   - Select a specific student to view their rank progression chart
   - Interact with the chart using the provided controls

## Technologies Used

- HTML, CSS, and JavaScript
- SheetJS library for Excel file parsing
- Chart.js for data visualization
- Chart.js Zoom plugin for interactive chart manipulation