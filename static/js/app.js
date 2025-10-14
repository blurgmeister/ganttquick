// Global state
let employees = [];
let tasks = [];
let currentStep = 1;

// Utility functions
function showMessage(message, type = 'success') {
    const msgEl = document.getElementById('message');
    msgEl.textContent = message;
    msgEl.className = `message ${type}`;
    setTimeout(() => {
        msgEl.style.display = 'none';
    }, 3000);
}

function goToStep(step) {
    // Hide all steps
    for (let i = 1; i <= 4; i++) {
        document.getElementById(`step${i}`).style.display = 'none';
    }
    // Show requested step
    document.getElementById(`step${step}`).style.display = 'block';
    currentStep = step;
}

// Step 1: Create Project
async function createProject() {
    const name = document.getElementById('projectName').value.trim();
    const startDate = document.getElementById('startDate').value;
    const globalHolidaysStr = document.getElementById('globalHolidays').value.trim();

    if (!name || !startDate) {
        showMessage('Please fill in all required fields', 'error');
        return;
    }

    const globalHolidays = globalHolidaysStr
        ? globalHolidaysStr.split(',').map(d => d.trim())
        : [];

    try {
        const response = await fetch('/api/project', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, start_date: startDate, global_holidays: globalHolidays })
        });

        const data = await response.json();

        if (response.ok) {
            showMessage('Project created successfully!', 'success');
            goToStep(2);
        } else {
            showMessage(data.error || 'Error creating project', 'error');
        }
    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
    }
}

// Step 2: Add Employees
function addEmployee() {
    const name = document.getElementById('employeeName').value.trim();
    const holidaysStr = document.getElementById('employeeHolidays').value.trim();

    if (!name) {
        showMessage('Please enter employee name', 'error');
        return;
    }

    // Get work pattern
    const checkboxes = document.querySelectorAll('.checkbox-group input[type="checkbox"]:checked');
    const workPattern = Array.from(checkboxes).map(cb => parseInt(cb.value));

    if (workPattern.length === 0) {
        showMessage('Please select at least one working day', 'error');
        return;
    }

    const holidays = holidaysStr ? holidaysStr.split(',').map(d => d.trim()) : [];

    employees.push({ name, work_pattern: workPattern, holidays });

    // Update display
    updateEmployeesList();

    // Clear form
    document.getElementById('employeeName').value = '';
    document.getElementById('employeeHolidays').value = '';

    showMessage(`Employee "${name}" added!`, 'success');
}

function updateEmployeesList() {
    const list = document.getElementById('employeesList');
    list.innerHTML = employees.map(emp =>
        `<span class="item-badge">${emp.name}</span>`
    ).join('');

    // Update task assignment dropdown
    const select = document.getElementById('taskAssignedTo');
    select.innerHTML = '<option value="">Select employee</option>' +
        employees.map(emp => `<option value="${emp.name}">${emp.name}</option>`).join('');
}

async function submitEmployees() {
    if (employees.length === 0) {
        showMessage('Please add at least one employee', 'error');
        return false;
    }

    try {
        const response = await fetch('/api/employees', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ employees })
        });

        const data = await response.json();

        if (response.ok) {
            return true;
        } else {
            showMessage(data.error || 'Error adding employees', 'error');
            return false;
        }
    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
        return false;
    }
}

// Step 3: Add Tasks
function addTask() {
    const name = document.getElementById('taskName').value.trim();
    const assignedTo = document.getElementById('taskAssignedTo').value;
    const estimatedDuration = parseInt(document.getElementById('taskDuration').value);
    const availability = parseInt(document.getElementById('taskAvailability').value);
    const contingencyMargin = parseInt(document.getElementById('taskContingency').value);
    const dependency = document.getElementById('taskDependency').value;

    if (!name || !assignedTo || !estimatedDuration) {
        showMessage('Please fill in all required fields', 'error');
        return;
    }

    if (availability < 1 || availability > 100) {
        showMessage('Availability must be between 1 and 100', 'error');
        return;
    }

    if (contingencyMargin < 0) {
        showMessage('Contingency margin cannot be negative', 'error');
        return;
    }

    const task = {
        name,
        assigned_to: assignedTo,
        estimated_duration: estimatedDuration,
        availability,
        contingency_margin: contingencyMargin,
        dependency: dependency || null
    };

    tasks.push(task);

    // Update display
    updateTasksList();

    // Clear form
    document.getElementById('taskName').value = '';
    document.getElementById('taskDuration').value = '1';
    document.getElementById('taskAvailability').value = '100';
    document.getElementById('taskContingency').value = '0';
    document.getElementById('taskDependency').value = '';

    showMessage(`Task "${name}" added!`, 'success');
}

function updateTasksList() {
    const list = document.getElementById('tasksList');
    list.innerHTML = tasks.map((task, idx) =>
        `<div class="item-badge">${idx + 1}. ${task.name} (${task.assigned_to}, ${task.estimated_duration}d)</div>`
    ).join('');

    // Update dependency dropdown
    const select = document.getElementById('taskDependency');
    select.innerHTML = '<option value="">No dependency</option>' +
        tasks.map(task => `<option value="${task.name}">${task.name}</option>`).join('');
}

async function submitTasks() {
    if (tasks.length === 0) {
        showMessage('Please add at least one task', 'error');
        return false;
    }

    try {
        const response = await fetch('/api/tasks', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ tasks })
        });

        const data = await response.json();

        if (response.ok) {
            return true;
        } else {
            showMessage(data.error || 'Error adding tasks', 'error');
            return false;
        }
    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
        return false;
    }
}

async function calculateAndView() {
    // Submit employees and tasks if not already done
    if (!(await submitEmployees())) return;
    if (!(await submitTasks())) return;

    try {
        // Calculate schedule
        const calcResponse = await fetch('/api/calculate', {
            method: 'POST'
        });

        const calcData = await calcResponse.json();

        if (!calcResponse.ok) {
            showMessage(calcData.error || 'Error calculating schedule', 'error');
            return;
        }

        // Get Gantt data
        const ganttResponse = await fetch('/api/gantt');
        const ganttData = await ganttResponse.json();

        if (ganttResponse.ok) {
            displayGanttChart(ganttData);
            goToStep(4);
            showMessage('Schedule calculated successfully!', 'success');
        } else {
            showMessage(ganttData.error || 'Error loading Gantt chart', 'error');
        }
    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
    }
}

// Step 4: Display Gantt Chart
function displayGanttChart(data) {
    const container = document.getElementById('ganttChart');

    // Parse dates
    const startDate = new Date(data.start_date);
    const endDate = new Date(data.end_date);

    // Generate date range
    const dates = [];
    let currentDate = new Date(startDate);
    while (currentDate <= endDate) {
        dates.push(new Date(currentDate));
        currentDate.setDate(currentDate.getDate() + 1);
    }

    // Build table
    let html = '<table class="gantt-table">';

    // Header row
    html += '<tr>';
    html += '<th>Task</th>';
    html += '<th>Assigned To</th>';
    html += '<th>Est. Days</th>';
    html += '<th>Actual Days</th>';
    html += '<th>Start Date</th>';
    html += '<th>End Date</th>';

    dates.forEach(date => {
        const dateStr = `${date.getMonth() + 1}/${date.getDate()}`;
        const dayStr = ['S', 'M', 'T', 'W', 'T', 'F', 'S'][date.getDay()];
        html += `<th class="date-header">${dateStr}<br>${dayStr}</th>`;
    });

    html += '</tr>';

    // Task rows
    data.tasks.forEach(task => {
        html += '<tr>';
        html += `<td>${task.name}</td>`;
        html += `<td>${task.assigned_to}</td>`;
        html += `<td>${task.estimated_duration}</td>`;
        html += `<td>${task.actual_duration}</td>`;
        html += `<td>${task.start_date || ''}</td>`;
        html += `<td>${task.end_date || ''}</td>`;

        const workingDatesSet = new Set(task.working_dates);

        dates.forEach(date => {
            const dateStr = date.toISOString().split('T')[0];
            const isWorking = workingDatesSet.has(dateStr);
            html += `<td class="date-cell ${isWorking ? 'working' : ''}"></td>`;
        });

        html += '</tr>';
    });

    html += '</table>';
    container.innerHTML = html;
}

// Export to Excel
async function exportToExcel() {
    try {
        const response = await fetch('/api/export');

        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'gantt_chart.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            showMessage('Excel file downloaded successfully!', 'success');
        } else {
            const data = await response.json();
            showMessage(data.error || 'Error exporting to Excel', 'error');
        }
    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
    }
}

// Reset project
async function resetProject() {
    if (!confirm('Are you sure you want to start a new project? All current data will be lost.')) {
        return;
    }

    try {
        await fetch('/api/reset', { method: 'POST' });

        employees = [];
        tasks = [];
        currentStep = 1;

        // Clear all forms
        document.getElementById('projectName').value = '';
        document.getElementById('startDate').value = '';
        document.getElementById('globalHolidays').value = '';

        goToStep(1);
        showMessage('Project reset successfully!', 'success');
    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
    }
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    // Set default date to today
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('startDate').value = today;
});
