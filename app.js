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
    
    // Get unique student names
    const studentNames = [...new Set(jsonData.map(row => row.Name))];
    
    // Populate student selector
    const studentSelector = document.getElementById('studentSelector');
    studentSelector.innerHTML = '<option value="">Select a student...</option>';
    
    studentNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        studentSelector.appendChild(option);
    });
    
    // Show student selector
    studentSelector.style.display = 'block';
    
    // Add event listener for student selection
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
    
    // Group by subject and assessment
    const groupedData = {};
    
    studentData.forEach(row => {
        if (!groupedData[row.Subj]) {
            groupedData[row.Subj] = {};
        }
        groupedData[row.Subj][row.Assessment] = row.Rank;
    });
    
    // Prepare data for Chart.js
    const subjects = Object.keys(groupedData);
    const assessments = [...new Set(studentData.map(row => row.Assessment))];
    
    const datasets = assessments.map((assessment, index) => {
        // Generate a color based on index
        const hue = (index * 137) % 360;
        const color = `hsla(${hue}, 70%, 50%, 0.7)`;
        
        return {
            label: assessment,
            data: subjects.map(subj => groupedData[subj][assessment] || null),
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
    
    // Create new chart
    rankChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: subjects,
            datasets: datasets
        },
        options: {
            scales: {
                y: {
                    reverse: true, // Higher rank is actually a lower number
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
                    text: `Rank Progression for ${studentName}`
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `${context.dataset.label}: Rank ${context.raw}`;
                        }
                    }
                }
            }
        }
    });
}