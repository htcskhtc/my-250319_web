let rankChart = null; // For storing chart instance

document.addEventListener('DOMContentLoaded', function() {
    // Load the Excel file
    fetch('SchoolPowerBIData.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            
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
                    displaySheet(workbook, this.value);
                    
                    // If "InternalExam" sheet is selected, prepare for chart visualization
                    if (this.value === "InternalExam") {
                        prepareStudentSelector(workbook, this.value);
                    } else {
                        document.getElementById('studentSelector').style.display = 'none';
                        document.getElementById('chartContainer').style.display = 'none';
                    }
                } else {
                    document.getElementById('tableContainer').innerHTML = '';
                    document.getElementById('studentSelector').style.display = 'none';
                    document.getElementById('chartContainer').style.display = 'none';
                }
            });
            
            // Display the first sheet by default if available
            if (workbook.SheetNames.length > 0) {
                sheetSelector.value = workbook.SheetNames[0];
                displaySheet(workbook, workbook.SheetNames[0]);
            }
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
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
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
            displayStudentChart(jsonData, this.value);
        } else {
            document.getElementById('chartContainer').style.display = 'none';
        }
    });
}

function displayStudentChart(data, studentName) {
    // Filter data for selected student
    const studentData = data.filter(row => row.Name === studentName);
    
    // Get unique assessments and subjects
    const assessments = [...new Set(studentData.map(row => row.Assessment))];
    const subjects = [...new Set(studentData.map(row => row.Subj))];
    
    // Group by subject and assessment, but reversed from before
    const groupedData = {};
    
    studentData.forEach(row => {
        if (!groupedData[row.Assessment]) {
            groupedData[row.Assessment] = {};
        }
        groupedData[row.Assessment][row.Subj] = row.Rank;
    });
    
    // Create a dataset for each subject (instead of assessment)
    const datasets = subjects.map((subject, index) => {
        // Generate a color based on index
        const hue = (index * 137) % 360;
        const color = `hsla(${hue}, 70%, 50%, 0.7)`;
        
        return {
            label: subject, // Subject is now the legend label
            data: assessments.map(assessment => groupedData[assessment][subject] || null),
            borderColor: color,
            backgroundColor: color,
            tension: 0.1
        };
    });
    
    // Display chart
    const chartContainer = document.getElementById('chartContainer');
    chartContainer.style.display = 'block';
    
    const ctx = document.getElementById('rankChart').getContext('2d');
    
    // Destroy existing chart if it exists
    if (rankChart) {
        rankChart.destroy();
    }
    
    // Create new chart with assessments as x-axis
    rankChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: assessments, // Changed from subjects to assessments for x-axis
            datasets: datasets
        },
        options: {
            scales: {
                y: {
                    reverse: false,
                    title: {
                        display: true,
                        text: 'Rank (higher is better)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Assessment' // Changed from 'Subject' to 'Assessment'
                    }
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: `Rank Progression for ${studentName}`
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `${context.dataset.label}: Rank ${context.raw}`;
                        }
                    }
                },
                legend: {
                    title: {
                        display: true,
                        text: 'Subject' // Add legend title for subjects
                    }
                }
            }
        }
    });
}