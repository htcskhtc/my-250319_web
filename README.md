# School Power BI Data Viewer

A web application that displays and visualizes data from the SchoolPowerBIData.xlsx Excel file, specializing in student performance analytics.

## Features

- View different sheets in the Excel file in a responsive table format
- Advanced visualization for the "InternalExam" sheet data
- Filter students by year, class, and number for easy student selection
- Interactive chart visualizations:
  - Rank progression chart showing student performance across different assessments
  - Subject-specific bar chart for individual assessment analysis
  - Detailed rank difference table with color-coded performance changes
- Data analysis features:
  - Performance change tracking with color-coded indicators
  - Summary statistics highlighting overall improvements and declines
  - Best and most challenging subject identification

## How to Use

1. Open the index.html file in a web browser
2. Select a sheet from the dropdown to view its data
3. For the "InternalExam" sheet:
   - Use the filters to narrow down student selection
   - Select a specific student to view their rank progression chart
   - Use the assessment selector to see subject-specific performance in bar chart format
   - Review the rank difference table to identify performance trends across assessments

## Visualization Features

- **Rank Progression Chart**: Tracks student performance in multiple subjects across all assessments
- **Assessment Bar Chart**: Displays subject rankings for a specific assessment with color-coded best/worst performance
- **Rank Difference Table**: Shows rank changes between assessments with color gradient indicators:
  - Green shades: Improvements in rank
  - Red shades: Declines in rank
  - Intensity based on the magnitude of change

## Technologies Used

- HTML, CSS, and JavaScript
- SheetJS library for Excel file parsing
- Chart.js for data visualization
- Chart.js plugins:
  - Data Labels plugin for enhanced data point labeling
  - Annotation plugin for chart annotations