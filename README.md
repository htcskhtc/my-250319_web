# School Power BI Data Viewer

A web application that displays and visualizes data from the SchoolPowerBIData.xlsx Excel file, specializing in student performance analytics.

## Overview

This application provides educators and administrators with an interactive tool to analyze student performance data across multiple assessments and subjects. The system includes user authentication, data visualization, and performance trend analysis.

## System Requirements

- Modern web browser with JavaScript enabled
- Local or remote web server for hosting the application files
- Excel data file (SchoolPowerBIData.xlsx) with required sheets and structure

## Installation

1. Clone or download this repository
2. Place all files on a web server or use a local server
3. Ensure SchoolPowerBIData.xlsx and userDB.xlsx are in the root directory
4. Open index.html through your web server

## Features

### Authentication System
- Excel-based user database (userDB.xlsx)
- Fallback to hardcoded credentials if Excel file is not available
- Session-based authentication using browser sessionStorage

### Data Visualization
- **Data Tables**: View raw data from different Excel sheets
- **Student Filtering**: Filter students by year, class, and number
- **Interactive Charts**:
  - Rank progression chart showing performance across assessments
  - Subject-specific bar chart for individual assessment analysis
  - Rank difference table with color-coded performance changes
  
### Performance Analysis
- Performance change tracking with color-coded indicators
- Summary statistics for improvements and declines
- Best and most challenging subject identification
- Student awards and achievements display
- Pre-Secondary 1 data and primary school information display

## Excel Data Structure

The application expects the following sheets in the Excel file:
- **InternalExam**: Contains student assessment data with columns for Name, Subj, Assessment, Rank, and RankDiff
- **InternalAct**: Contains student activity/awards data with columns for Name, Year, and Act_Award
- **Pre_S1**: Contains pre-secondary school assessment data with columns for Name, Assessment, and Rank
- **IntPriSch**: Contains primary school information with columns for Name and PriSch
- Additional sheets can be included and will be displayed as tables

## How to Use

1. Open the index.html file in a web browser
2. Log in using one of these credentials:
   - Username: `admin` | Password: `admin123`
   - Username: `user` | Password: `user123`
   - Username: `jackchui` | Password: `jackchui123456`
3. Select a sheet from the dropdown to view its data
4. For the "InternalExam" sheet:
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
- **Pre-S1 Profile**: Displays student's pre-secondary assessment performance and primary school background

## Technologies Used

- HTML5, CSS3, and JavaScript (ES6+)
- [SheetJS](https://sheetjs.com/) library for Excel file parsing
- [Chart.js](https://www.chartjs.org/) for data visualization
- Chart.js plugins:
  - Data Labels plugin for enhanced data point labeling
  - Annotation plugin for chart annotations
- Browser sessionStorage for maintaining authentication state

## Troubleshooting

- If the Excel files fail to load, check that they're properly formatted with the required columns
- Make sure SchoolPowerBIData.xlsx and userDB.xlsx are in the same directory as the HTML file
- If authentication fails, the application will fall back to default credentials
- For visualization issues, check browser console for error messages

## Future Development

- Export visualizations as images or PDF
- Add more advanced filtering options
- Implement comparative analysis between students
- Add admin dashboard for managing user accounts

## License

This project is provided for educational purposes only.