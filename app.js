let rankChart = null; // For storing chart instance
let activityData = null; // For storing activity/awards data
let workbook = null; // Global workbook variable to use across functions

document.addEventListener('DOMContentLoaded', function() {
    // Register Chart.js plugins if they exist
    if (Chart) {
        if (Chart.register) {
            // Register plugins if they exist
            if (window.ChartDataLabels) {
                Chart.register(ChartDataLabels);
                console.log("Chart.js DataLabels plugin registered");
            }
            if (window.annotationPlugin) {
                Chart.register(annotationPlugin);
                console.log("Chart.js Annotation plugin registered");
            }
        }
    }
    
    // Load the Excel file
    fetch('SchoolPowerBIData.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            workbook = XLSX.read(data, { type: 'array' });
            
            // Store the InternalAct sheet data for later use
            if (workbook.SheetNames.includes('InternalAct')) {
                const actSheet = workbook.Sheets['InternalAct'];
                activityData = XLSX.utils.sheet_to_json(actSheet);
                console.log("Activity data loaded:", activityData.length, "records");
            } else {
                console.warn("InternalAct sheet not found in the Excel file");
            }
            
            // Get all sheet names
            const sheetSelector = document.getElementById('sheetSelector');
            workbook.SheetNames.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelector.appendChild(option);
            });
            
            // Event listener for sheet selection
            sheetSelector.addEventListener('change', function() {
                if (this.value) {
                    // Only display the sheet if it's not "InternalExam"
                    if (this.value !== "InternalExam") {
                        displaySheet(workbook, this.value);
                    } else {
                        // Hide the table container for "InternalExam" sheet
                        document.getElementById('tableContainer').innerHTML = '';
                    }
                    
                    // If "InternalExam" sheet is selected, prepare for chart visualization
                    if (this.value === "InternalExam") {
                        prepareStudentSelector(workbook, this.value);
                        document.getElementById('tableContainer').style.display = 'none';
                    } else {
                        document.getElementById('studentSelector').style.display = 'none';
                        document.getElementById('chartContainer').style.display = 'none';
                        document.getElementById('tableContainer').style.display = 'block';
                    }
                } else {
                    document.getElementById('tableContainer').innerHTML = '';
                    document.getElementById('studentSelector').style.display = 'none';
                    document.getElementById('chartContainer').style.display = 'none';
                }
            });
            
            // Remove the automatic selection of the first sheet
            // Leave the sheetSelector empty so the user must make a choice
        })
        .catch(error => {
            console.error('Error loading Excel file:', error);
            document.getElementById('tableContainer').innerHTML = 
                `<div class="error">Error loading Excel file: ${error.message}</div>`;
        });
});

function displaySheet(workbook, sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Create table
    let tableHTML = '<table>';
    
    // Process headers (first row)
    if (jsonData.length > 0) {
        tableHTML += '<thead><tr>';
        jsonData[0].forEach(header => {
            tableHTML += `<th>${header}</th>`;
        });
        tableHTML += '</tr></thead>';
    }
    
    // Process data rows
    tableHTML += '<tbody>';
    for (let i = 1; i < jsonData.length; i++) {
        tableHTML += '<tr>';
        jsonData[i].forEach(cell => {
            tableHTML += `<td>${cell !== undefined ? cell : ''}</td>`;
        });
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody></table>';
    
    document.getElementById('tableContainer').innerHTML = tableHTML;
}

// Add these new functions
function prepareStudentSelector(workbook, sheetName) {
    console.log("prepareStudentSelector called", sheetName);
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    console.log("Data loaded:", jsonData.length, "rows");
    
    // Get unique student names and parse their components
    const students = [];
    const years = new Set();
    const classes = new Set();
    const numbers = new Set();
    
    jsonData.forEach(row => {
        if (students.some(s => s.fullName === row.Name)) return;
        
        // Parse student name components
        // Format: "2324 6A 03 Chan Yuen Yan Yolanda"
        const nameParts = row.Name.split(' ');
        if (nameParts.length >= 4) {
            const year = nameParts[0];
            const className = nameParts[1];
            const number = nameParts[2];
            const actualName = nameParts.slice(3).join(' ');
            
            students.push({
                fullName: row.Name,
                year: year,
                class: className,
                number: number,
                name: actualName
            });
            
            years.add(year);
            classes.add(className);
            numbers.add(number);
        } else {
            // Handle unexpected format
            students.push({
                fullName: row.Name,
                year: '',
                class: '',
                number: '',
                name: row.Name
            });
        }
    });
    
    // Sort and populate filter dropdowns
    const yearFilter = document.getElementById('yearFilter');
    const classFilter = document.getElementById('classFilter');
    const numberFilter = document.getElementById('numberFilter');
    
    // Populate year filter
    [...years].sort().forEach(year => {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        yearFilter.appendChild(option);
    });
    
    // Populate class filter
    [...classes].sort().forEach(cls => {
        const option = document.createElement('option');
        option.value = cls;
        option.textContent = cls;
        classFilter.appendChild(option);
    });
    
    // Populate number filter
    [...numbers].sort((a, b) => parseInt(a) - parseInt(b)).forEach(num => {
        const option = document.createElement('option');
        option.value = num;
        option.textContent = num;
        numberFilter.appendChild(option);
    });
    
    // Show filters
    document.getElementById('studentFilters').style.display = 'block';
    document.getElementById('studentSelector').style.display = 'block';
    
    // Function to update student selector based on filters
    function updateStudentSelector() {
        const selectedYear = yearFilter.value;
        const selectedClass = classFilter.value;
        const selectedNumber = numberFilter.value;
        
        // Filter students based on selected filters
        const filteredStudents = students.filter(student => {
            return (!selectedYear || student.year === selectedYear) && 
                   (!selectedClass || student.class === selectedClass) && 
                   (!selectedNumber || student.number === selectedNumber);
        });
        
        // Sort students by their full name in ascending order
        filteredStudents.sort((a, b) => a.fullName.localeCompare(b.fullName));
        
        // Update student selector
        const studentSelector = document.getElementById('studentSelector');
        studentSelector.innerHTML = '<option value="">Select a student...</option>';
        
        filteredStudents.forEach(student => {
            const option = document.createElement('option');
            option.value = student.fullName;
            option.textContent = student.fullName;
            studentSelector.appendChild(option);
        });
    }
    
    // Initial population of student selector
    updateStudentSelector();
    
    // Add event listeners for filters
    yearFilter.addEventListener('change', updateStudentSelector);
    classFilter.addEventListener('change', updateStudentSelector);
    numberFilter.addEventListener('change', updateStudentSelector);
    
    // Add event listener for student selection
    const studentSelector = document.getElementById('studentSelector');
    studentSelector.addEventListener('change', function() {
        if (this.value) {
            displayStudentChart(jsonData, this.value, workbook);
        } else {
            document.getElementById('chartContainer').style.display = 'none';
        }
    });
}

function displayStudentChart(data, studentName, workbook) {
    console.log("displayStudentChart called", studentName);
    // Filter data for selected student
    const studentData = data.filter(row => row.Name === studentName);
    console.log("Student data:", studentData.length, "rows");
    
    // Get unique assessments and subjects, and sort assessments
    const assessments = [...new Set(studentData.map(row => row.Assessment))].sort();
    const subjects = [...new Set(studentData.map(row => row.Subj))];
    
    // Group by subject and assessment
    const groupedData = {};
    
    studentData.forEach(row => {
        if (!groupedData[row.Assessment]) {
            groupedData[row.Assessment] = {};
        }
        groupedData[row.Assessment][row.Subj] = row.Rank;
    });
    
    // Create a dataset for each subject
    const datasets = subjects.map((subject, index) => {
        // Generate a color based on index
        const hue = (index * 137) % 360;
        const color = `hsla(${hue}, 70%, 50%, 0.7)`;
        
        return {
            label: subject,
            data: assessments.map(assessment => groupedData[assessment][subject] || null),
            borderColor: color,
            backgroundColor: color,
            tension: 0.1,
            pointRadius: 6,
            pointHoverRadius: 9,
            pointBackgroundColor: color,
            pointBorderColor: 'white',
            pointBorderWidth: 2,
        };
    });
    
    // Display chart
    const chartContainer = document.getElementById('chartContainer');
    chartContainer.style.display = 'block';
    
    // Add chart controls before the chart
    // FIXED: Properly structure the HTML with clear separation between chart sections
    let controlsHTML = `
        <div class="chart-controls" style="margin-bottom: 15px;">
            <select id="assessmentSelector" style="padding: 8px; margin-left: 10px; min-width: 150px;">
                <option value="">Select Assessment...</option>
                ${assessments.map(a => `<option value="${a}">${a}</option>`).join('')}
            </select>
            <div class="checkbox-container" style="margin-top: 10px;">
                <span>Subject Legend: </span>
                <div id="subjectToggles" style="display: inline-flex; flex-wrap: wrap; gap: 10px;"></div>
            </div>
        </div>
        
        <!-- FIXED: Bar chart container with complete structure including controls -->
        <div id="subjectRankBarChart" style="height: 400px; display: none; margin-bottom: 30px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
            <h4 style="text-align: center; margin-top: 0;">Subject Performance for Selected Assessment</h4>
            <div style="height: 300px;">
                <canvas id="barChart"></canvas>
            </div>
            <div id="barChartControls" style="margin-top: 15px;">
                <!-- Controls will be added here by JS -->
            </div>
        </div>
        
        <!-- FIXED: Clear separation with the rank progression chart -->
        <div style="height: 400px; margin-top: 20px; border-top: 1px solid #eee; padding-top: 15px;">
            <h4 style="text-align: center; margin-top: 0;">Static Rank Progression Chart</h4>
            <canvas id="rankChart"></canvas>
        </div>
        <div id="rankDiffTable" style="margin-top: 30px;"></div>
    `;
    
    // Replace the entire chart container content
    chartContainer.innerHTML = controlsHTML;
    
    const ctx = document.getElementById('rankChart').getContext('2d');
    
    // Destroy existing chart if it exists
    if (rankChart) {
        rankChart.destroy();
    }
    
    // Set initial y-axis state
    let yAxisReversed = false;
    
    // Create new chart with assessments as x-axis
    rankChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: assessments,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    reverse: false,
                    min: 0,
                    max: 11,
                    beginAtZero: false,
                    ticks: {
                        stepSize: 1,
                        callback: function(value, index, ticks) {
                            return (value === 0 || value === 11) ? '' : value;
                        }
                    },
                    title: {
                        display: true,
                        text: 'Rank (lower number is better)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Assessment'
                    }
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: `Rank Progression for ${studentName}`
                },
                tooltip: {
                    enabled: false // Disable tooltips
                },
                legend: {
                    title: {
                        display: true,
                        text: 'Subject'
                    }
                },
                zoom: {
                    enabled: false // Disable zoom plugin
                }
            },
            interaction: {
                mode: 'none' // Disable all interactions
            },
            events: [] // Disable all event listeners (hover, clicks)
        }
    });
    
    // Display information about subjects instead of toggle checkboxes
    const subjectToggles = document.getElementById('subjectToggles');
    datasets.forEach((dataset, i) => {
        const hue = (i * 137) % 360;
        const color = `hsla(${hue}, 70%, 50%, 0.7)`;
        
        const label = document.createElement('div');
        label.style.display = 'inline-flex';
        label.style.alignItems = 'center';
        label.style.marginRight = '10px';
        label.style.marginBottom = '5px';
        
        const colorIndicator = document.createElement('span');
        colorIndicator.style.display = 'inline-block';
        colorIndicator.style.width = '12px';
        colorIndicator.style.height = '12px';
        colorIndicator.style.backgroundColor = color;
        colorIndicator.style.marginRight = '5px';
        
        const text = document.createTextNode(dataset.label);
        
        label.appendChild(colorIndicator);
        label.appendChild(text);
        subjectToggles.appendChild(label);
    });
    
    // Create RankDiff table
    createRankDiffTable(studentData, subjects, assessments);
    
    // Add this line to display student awards
    displayStudentAwards(studentName);
    
    // Initialize the Assessment Bar Chart functionality
    initializeAssessmentBarChart(studentData, groupedData, subjects);
    
    // Display student awards
    displayStudentAwards(studentName);
    
    // Display Pre-S1 and primary school data
    displayPreS1AndPrimarySchool(studentName, workbook);
}

// Add new function to create the RankDiff table
function createRankDiffTable(studentData, subjects, assessments) {
    const tableContainer = document.getElementById('rankDiffTable');
    
    // Add CSS for the table styling
    const cssStyle = `
        <style>
            .rank-diff-table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                box-shadow: 0 2px 6px rgba(0,0,0,0.1);
            }
            
            .rank-diff-table th, .rank-diff-table td {
                padding: 10px;
                text-align: center;
                border: 1px solid #e0e0e0;
            }
            
            .rank-diff-table th {
                background-color: #f5f5f5;
                font-weight: bold;
                position: sticky;
                top: 0;
            }
            
            .rank-diff-table th:first-child {
                text-align: left;
            }
            
            .rank-diff-table td:first-child {
                text-align: left;
                font-weight: bold;
                background-color: #f9f9f9;
            }
            
            /* Color gradient for positive changes */
            .positive-change-1 { background-color: rgba(220, 255, 220, 0.6); color: #006400; }
            .positive-change-2 { background-color: rgba(180, 255, 180, 0.7); color: #006400; }
            .positive-change-3 { background-color: rgba(140, 255, 140, 0.8); color: #006400; }
            .positive-change-4 { background-color: rgba(100, 255, 100, 0.9); color: #006400; }
            .positive-change-5 { background-color: rgba(50, 200, 50, 1.0); color: white; }
            
            /* Color gradient for negative changes */
            .negative-change-1 { background-color: rgba(255, 220, 220, 0.6); color: #8b0000; }
            .negative-change-2 { background-color: rgba(255, 180, 180, 0.7); color: #8b0000; }
            .negative-change-3 { background-color: rgba(255, 140, 140, 0.8); color: #8b0000; }
            .negative-change-4 { background-color: rgba(255, 100, 100, 0.9); color: #8b0000; }
            .negative-change-5 { background-color: rgba(200, 50, 50, 1.0); color: white; }
            
            .no-change { background-color: #f0f0f0; color: #555; }
            
            .rank-diff-table td:not(:first-child) {
                font-weight: bold;
                min-width: 60px;
            }
            
            .legend-container {
                display: flex;
                justify-content: center;
                margin-bottom: 15px;
                flex-wrap: wrap;
            }
            
            .legend-item {
                display: flex;
                align-items: center;
                margin-right: 15px;
                margin-bottom: 5px;
            }
            
            .legend-color {
                width: 20px;
                height: 20px;
                margin-right: 5px;
                border: 1px solid #ddd;
            }
            
            .summary-stats {
                margin-bottom: 15px;
                padding: 10px;
                background-color: #f9f9f9;
                border-radius: 4px;
                border-left: 4px solid #4CAF50;
            }
            
            .highlight {
                font-weight: bold;
            }
            
            .assessment-group {
                border-bottom: 2px solid #aaa;
            }
            
            .assessment-group:last-child {
                border-bottom: none;
            }
        </style>
    `;
    
    // Create table heading
    let tableHTML = '<h3>Rank Difference by Subject and Assessment</h3>';
    
    // Add color legend
    tableHTML += `
        <div class="legend-container">
            <div class="legend-item">
                <div class="legend-color positive-change-5"></div>
                <span>Large Improvement (+5)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color positive-change-3"></div>
                <span>Medium Improvement (+3)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color positive-change-1"></div>
                <span>Small Improvement (+1)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color no-change"></div>
                <span>No Change (0)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color negative-change-1"></div>
                <span>Small Decline (-1)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color negative-change-3"></div>
                <span>Medium Decline (-3)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color negative-change-5"></div>
                <span>Large Decline (-5)</span>
            </div>
        </div>
    `;
    
    // Add summary statistics
    let totalImprovements = 0;
    let totalDeclines = 0;
    let bestSubject = '';
    let bestImprovement = 0;
    let worstSubject = '';
    let worstDecline = 0;
    
    // Calculate summary statistics
    subjects.forEach(subject => {
        let subjectImprovements = 0;
        let subjectDeclines = 0;
        
        assessments.forEach(assessment => {
            const dataPoint = studentData.find(row => 
                row.Assessment === assessment && 
                row.Subj === subject
            );
            
            if (dataPoint && dataPoint.RankDiff) {
                if (dataPoint.RankDiff > 0) {
                    subjectImprovements += dataPoint.RankDiff;
                } else if (dataPoint.RankDiff < 0) {
                    subjectDeclines += dataPoint.RankDiff;
                }
            }
        });
        
        totalImprovements += subjectImprovements;
        totalDeclines += subjectDeclines;
        
        if (subjectImprovements > bestImprovement) {
            bestImprovement = subjectImprovements;
            bestSubject = subject;
        }
        
        if (subjectDeclines < worstDecline) {
            worstDecline = subjectDeclines;
            worstSubject = subject;
        }
    });
    
    tableHTML += `
        <div class="summary-stats">
            <p>Total improvements: <span class="highlight">+${totalImprovements}</span> | 
               Total declines: <span class="highlight">${totalDeclines}</span></p>
            <p>Best performing subject: <span class="highlight">${bestSubject}</span> (Total improvement: +${bestImprovement})</p>
            <p>Most challenging subject: <span class="highlight">${worstSubject}</span> (Total change: ${worstDecline})</p>
        </div>
    `;
    
    tableHTML += cssStyle;
    tableHTML += '<div style="overflow-x: auto;"><table class="rank-diff-table">';
    
    // Header row with subject names
    tableHTML += '<tr><th>Assessment</th>';
    subjects.forEach(subject => {
        tableHTML += `<th>${subject}</th>`;
    });
    tableHTML += '</tr>';
    
    // Create rows for each assessment
    let currentYear = '';
    assessments.forEach((assessment, index) => {
        // Extract year/term from assessment name if possible
        const assessmentParts = assessment.split(' ');
        const assessmentYear = assessmentParts[0] || '';
        
        // Add a visual separator between different years/terms
        let rowClass = '';
        if (assessmentYear !== currentYear) {
            currentYear = assessmentYear;
            rowClass = ' class="assessment-group"';
        }
        
        tableHTML += `<tr${rowClass}><td>${assessment}</td>`;
        
        // For each subject, find the RankDiff value
        subjects.forEach(subject => {
            const dataPoint = studentData.find(row => 
                row.Assessment === assessment && 
                row.Subj === subject
            );
            
            const rankDiff = dataPoint && dataPoint.RankDiff !== undefined ? dataPoint.RankDiff : '';
            
            // Add color coding based on rank difference value and intensity
            let cellClass = '';
            if (rankDiff !== '') {
                if (rankDiff > 0) {
                    // Positive changes (improvements)
                    if (rankDiff >= 5) {
                        cellClass = 'positive-change-5';
                    } else if (rankDiff >= 3) {
                        cellClass = 'positive-change-4';
                    } else if (rankDiff >= 2) {
                        cellClass = 'positive-change-3';
                    } else if (rankDiff >= 1) {
                        cellClass = 'positive-change-1';
                    }
                } else if (rankDiff < 0) {
                    // Negative changes (declines)
                    if (rankDiff <= -5) {
                        cellClass = 'negative-change-5';
                    } else if (rankDiff <= -3) {
                        cellClass = 'negative-change-4';
                    } else if (rankDiff <= -2) {
                        cellClass = 'negative-change-3';
                    } else if (rankDiff <= -1) {
                        cellClass = 'negative-change-1';
                    }
                } else {
                    cellClass = 'no-change';
                }
            }
            
            // Add + sign prefix for positive values
            const displayValue = rankDiff > 0 ? `+${rankDiff}` : rankDiff;
            
            tableHTML += `<td class="${cellClass}">${displayValue}</td>`;
        });
        
        tableHTML += '</tr>';
    });
    
    tableHTML += '</table></div>';
    tableContainer.innerHTML = tableHTML;
}

// Add new function to create and manage the assessment bar chart
function initializeAssessmentBarChart(studentData, groupedData, subjects) {
    let barChart = null;
    const assessmentSelector = document.getElementById('assessmentSelector');
    const barChartContainer = document.getElementById('subjectRankBarChart');
    
    console.log("Assessment Bar Chart functionality initialized");
    
    // Add event listener for assessment selection
    assessmentSelector.addEventListener('change', function() {
        const selectedAssessment = this.value;
        console.log("Selected assessment:", selectedAssessment);
        
        if (!selectedAssessment) {
            barChartContainer.style.display = 'none';
            return;
        }
        
        // Display the bar chart container
        barChartContainer.style.display = 'block';
        
        // Get the ranks for the selected assessment
        const assessmentData = groupedData[selectedAssessment];
        console.log("Assessment data:", assessmentData);
        
        if (!assessmentData) {
            barChartContainer.innerHTML = '<p>No data available for this assessment</p>';
            return;
        }
        
        // Prepare data for the bar chart
        const subjectRanks = subjects.map(subject => {
            return {
                subject: subject,
                rank: assessmentData[subject] || null
            };
        }).filter(item => item.rank !== null);
        
        console.log("Subject ranks data:", subjectRanks);
        
        if (subjectRanks.length === 0) {
            barChartContainer.innerHTML = '<p>No rank data available for this assessment</p>';
            return;
        }
        
        // Sort by rank (ascending)
        subjectRanks.sort((a, b) => a.rank - b.rank);
        
        // Generate colors - best (lowest rank) is green, worst is red
        const colors = subjectRanks.map((item, index, array) => {
            if (index === 0) return 'rgba(75, 192, 75, 0.8)'; // Best - green
            if (index === array.length - 1) return 'rgba(255, 99, 71, 0.8)'; // Worst - red
            return 'rgba(54, 162, 235, 0.7)'; // Mid - blue
        });
        
        const borderColors = colors.map(color => color.replace('0.7', '1').replace('0.8', '1'));
        
        // Destroy existing chart if it exists
        if (barChart) {
            barChart.destroy();
        }
        
        // FIXED: Ensure chart structure is preserved and controls container exists
        if (!document.getElementById('barChart')) {
            barChartContainer.innerHTML = `
                <h4 style="text-align: center; margin-top: 0;">Subject Performance for Selected Assessment</h4>
                <div style="height: 300px;">
                    <canvas id="barChart"></canvas>
                </div>
                <div id="barChartControls" style="margin-top: 15px;"></div>
            `;
        }
        
        // Make sure we have a canvas element
        const barChartCanvas = document.getElementById('barChart');
        if (!barChartCanvas) {
            console.error("Bar chart canvas element not found!");
            return;
        }
        
        // Create the bar chart with simpler options first
        const ctx = barChartCanvas.getContext('2d');
        console.log("Creating bar chart...");
        
        try {
            barChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: subjectRanks.map(item => item.subject),
                    datasets: [{
                        label: 'Rank',
                        data: subjectRanks.map(item => item.rank),
                        backgroundColor: colors,
                        borderColor: borderColors,
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            max: 11,
                            ticks: {
                                stepSize: 1
                            },
                            title: {
                                display: true,
                                text: 'Rank (lower is better)'
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: 'Subject'
                            }
                        }
                    },
                    plugins: {
                        title: {
                            display: true,
                            text: `Subject Performance for ${selectedAssessment}`,
                            font: {
                                size: 16
                            }
                        },
                        legend: {
                            display: false
                        }
                    }
                }
            });
            
            console.log("Bar chart created successfully");
            
            // FIXED: Place controls in the dedicated container
            const controlsContainer = document.getElementById('barChartControls');
            controlsContainer.innerHTML = ''; // Clear previous controls
            
            // Add rank labels on top of bars
            const rankLabels = document.createElement('div');
            rankLabels.style.textAlign = 'center';
            rankLabels.innerHTML = '<p><strong>Ranks:</strong> ' + 
                subjectRanks.map(item => `${item.subject}: ${item.rank}`).join(' | ') + '</p>';
            controlsContainer.appendChild(rankLabels);
            
            // Add sorting buttons
            const sortingControls = document.createElement('div');
            sortingControls.style.textAlign = 'center';
            sortingControls.style.marginTop = '10px';
            sortingControls.innerHTML = `
                <button id="sortByRankBtn" class="chart-btn" style="margin-right: 10px;">Sort by Rank</button>
                <button id="sortBySubjectBtn" class="chart-btn">Sort by Subject</button>
            `;
            controlsContainer.appendChild(sortingControls);
            
            // Add event listeners for sorting buttons
            document.getElementById('sortByRankBtn').addEventListener('click', function() {
                subjectRanks.sort((a, b) => a.rank - b.rank);
                updateBarChart();
            });
            
            document.getElementById('sortBySubjectBtn').addEventListener('click', function() {
                subjectRanks.sort((a, b) => a.subject.localeCompare(b.subject));
                updateBarChart();
            });
            
            function updateBarChart() {
                barChart.data.labels = subjectRanks.map(item => item.subject);
                barChart.data.datasets[0].data = subjectRanks.map(item => item.rank);
                
                // Update colors based on new sorting
                const newColors = subjectRanks.map((item, index, array) => {
                    if (index === 0 && array[0].rank === Math.min(...array.map(i => i.rank))) {
                        return 'rgba(75, 192, 75, 0.8)'; // Best - green
                    }
                    if (index === array.length - 1 && array[array.length - 1].rank === Math.max(...array.map(i => i.rank))) {
                        return 'rgba(255, 99, 71, 0.8)'; // Worst - red
                    }
                    return 'rgba(54, 162, 235, 0.7)'; // Mid - blue
                });
                
                barChart.data.datasets[0].backgroundColor = newColors;
                barChart.data.datasets[0].borderColor = newColors.map(color => color.replace('0.7', '1').replace('0.8', '1'));
                
                barChart.update();
                
                // Update rank labels
                rankLabels.innerHTML = '<p><strong>Ranks:</strong> ' + 
                    subjectRanks.map(item => `${item.subject}: ${item.rank}`).join(' | ') + '</p>';
            }
        } catch (error) {
            console.error("Error creating bar chart:", error);
            barChartContainer.innerHTML += `<p style="color: red;">Error creating chart: ${error.message}</p>`;
        }
    });
    
    // Display initial instructions
    const chartContent = barChartContainer.querySelector('div:not(h4)');
    if (!chartContent) {
        barChartContainer.innerHTML = `
            <h4 style="text-align: center; margin-top: 0;">Subject Performance for Selected Assessment</h4>
            <div style="height: 300px;">
                <canvas id="barChart"></canvas>
            </div>
            <div id="barChartControls" style="margin-top: 15px;">
                <p style="text-align: center; color: #666;">
                    Select an assessment from the dropdown above to see the student's performance across subjects.
                </p>
            </div>
        `;
    }
}

// Add this new function to display student awards
function displayStudentAwards(studentName) {
    console.log("Displaying awards for student:", studentName);
    
    // Create a container for the awards table if it doesn't exist
    let awardsContainer = document.getElementById('studentAwardsContainer');
    if (!awardsContainer) {
        awardsContainer = document.createElement('div');
        awardsContainer.id = 'studentAwardsContainer';
        awardsContainer.style.marginTop = '30px';
        document.getElementById('rankDiffTable').after(awardsContainer);
    }
    
    // Check if we have activity data
    if (!activityData) {
        awardsContainer.innerHTML = '<p>Award data is not available</p>';
        return;
    }
    
    // Filter awards for the selected student
    const studentAwards = activityData.filter(record => record.Name === studentName);
    
    // Generate the awards table
    let tableHTML = `
        <h3>Awards and Achievements for ${studentName}</h3>
        <style>
            .awards-table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                box-shadow: 0 2px 6px rgba(0,0,0,0.1);
            }
            
            .awards-table th, .awards-table td {
                padding: 10px;
                text-align: left;
                border: 1px solid #e0e0e0;
            }
            
            .awards-table th {
                background-color: #f5f5f5;
                font-weight: bold;
            }
            
            .awards-table tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            
            .no-awards {
                padding: 15px;
                background-color: #f9f9f9;
                border-left: 4px solid #607D8B;
                margin: 10px 0;
            }
        </style>
    `;
    
    if (studentAwards.length > 0) {
        tableHTML += `
            <table class="awards-table">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Year</th>
                        <th>Award</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        studentAwards.forEach((award, index) => {
            // Get the year value, display empty string if not available
            const yearValue = award.Year !== undefined ? award.Year : '';
            
            tableHTML += `
                <tr>
                    <td>${index + 1}</td>
                    <td>${yearValue}</td>
                    <td>${award.Act_Award}</td>
                </tr>
            `;
        });
        
        tableHTML += `
                </tbody>
            </table>
        `;
    } else {
        tableHTML += `
            <div class="no-awards">
                <p>No awards recorded for this student.</p>
            </div>
        `;
    }
    
    awardsContainer.innerHTML = tableHTML;
}

// Add this new function to display Pre-S1 data and primary school
function displayPreS1AndPrimarySchool(studentName, workbook) {
    console.log("Displaying Pre-S1 and primary school data for:", studentName);
    
    // Create a container for the Pre-S1 and primary school info
    let preS1Container = document.getElementById('preS1Container');
    if (!preS1Container) {
        preS1Container = document.createElement('div');
        preS1Container.id = 'preS1Container';
        preS1Container.style.marginTop = '30px';
        preS1Container.style.marginBottom = '30px';
        
        // Insert between rank diff table and awards container
        const rankDiffTable = document.getElementById('rankDiffTable');
        if (rankDiffTable) {
            rankDiffTable.after(preS1Container);
        }
    }
    
    // Check if we have the necessary sheets in the workbook
    if (!workbook || !workbook.SheetNames.includes('Pre_S1') || !workbook.SheetNames.includes('IntPriSch')) {
        preS1Container.innerHTML = '<p>Pre-S1 or Primary School data is not available</p>';
        return;
    }
    
    // Extract data from Pre_S1 sheet
    const preS1Sheet = workbook.Sheets['Pre_S1'];
    const preS1Data = XLSX.utils.sheet_to_json(preS1Sheet);
    
    // Filter for the selected student
    const studentPreS1Data = preS1Data.filter(record => record.Name === studentName);
    
    // Extract data from IntPriSch sheet
    const priSchSheet = workbook.Sheets['IntPriSch'];
    const priSchData = XLSX.utils.sheet_to_json(priSchSheet);
    
    // Find the student's primary school
    const studentPriSchInfo = priSchData.find(record => record.Name === studentName);
    const primarySchool = studentPriSchInfo ? studentPriSchInfo.PriSch : 'Not available';
    
    // Generate the HTML content
    let contentHTML = `
        <style>
            .pre-s1-container {
                background-color: white;
                border-radius: 8px;
                box-shadow: 0 2px 6px rgba(0,0,0,0.1);
                padding: 20px;
                margin-bottom: 20px;
            }
            
            .pre-s1-header {
                color: #333;
                border-bottom: 2px solid #4CAF50;
                padding-bottom: 10px;
                margin-top: 0;
            }
            
            .pre-s1-school {
                background-color: #f5f7fa;
                padding: 15px;
                border-radius: 6px;
                margin-bottom: 15px;
                border-left: 4px solid #2196F3;
            }
            
            .pre-s1-school h4 {
                margin-top: 0;
                color: #2196F3;
            }
            
            .pre-s1-summary {
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
                margin-top: 15px;
            }
            
            .pre-s1-subject {
                flex: 1;
                min-width: 150px;
                background-color: #f9f9f9;
                padding: 15px;
                border-radius: 6px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            
            .pre-s1-subject h4 {
                margin-top: 0;
                color: #555;
                font-size: 16px;
                border-bottom: 1px solid #ddd;
                padding-bottom: 5px;
            }
            
            .pre-s1-rank {
                font-size: 28px;
                font-weight: bold;
                color: #4CAF50;
                text-align: center;
                margin: 10px 0;
            }
            
            .pre-s1-note {
                font-size: 12px;
                color: #777;
                text-align: center;
            }
            
            .no-data {
                color: #999;
                font-style: italic;
                text-align: center;
                padding: 20px;
            }
        </style>
        
        <div class="pre-s1-container">
            <h3 class="pre-s1-header">Pre-Secondary 1 Academic Profile</h3>
            
            <div class="pre-s1-school">
                <h4>Primary School Background</h4>
                <p><strong>School:</strong> ${primarySchool}</p>
            </div>
    `;
    
    if (studentPreS1Data.length > 0) {
        contentHTML += `
            <h4>Pre-S1 Assessment Performance</h4>
            <div class="pre-s1-summary">
        `;
        
        // Group by assessment subject
        const subjectGroups = {};
        studentPreS1Data.forEach(record => {
            if (!subjectGroups[record.Assessment]) {
                subjectGroups[record.Assessment] = record;
            }
        });
        
        // Display each subject
        Object.keys(subjectGroups).forEach(subject => {
            const record = subjectGroups[subject];
            contentHTML += `
                <div class="pre-s1-subject">
                    <h4>${subject}</h4>
                    <div class="pre-s1-rank">${record.Rank}</div>
                    <div class="pre-s1-note">Class Rank (1st is best)</div>
                </div>
            `;
        });
        
        contentHTML += `
            </div>
        `;
    } else {
        contentHTML += `
            <div class="no-data">No Pre-S1 assessment data available for this student</div>
        `;
    }
    
    contentHTML += `</div>`;
    preS1Container.innerHTML = contentHTML;
}