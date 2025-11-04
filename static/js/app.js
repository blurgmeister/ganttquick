// Global state
let employees = [];
let tasks = [];
let currentStep = 1;
let dataAlreadySubmitted = false; // Track if employees/tasks already submitted to backend
let projectCreated = false; // Track if project has been created

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

    // Update step 1 button text based on whether project exists
    if (step === 1 && projectCreated) {
        const createBtn = document.querySelector('#step1 .btn-primary');
        if (createBtn) {
            createBtn.textContent = 'Update Project';
        }
    }
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
            const wasUpdating = projectCreated;
            projectCreated = true;
            document.getElementById('nextToStep2').style.display = 'inline-block';

            if (wasUpdating) {
                showMessage('Project updated successfully!', 'success');
                // If we have employees/tasks already, mark them for resubmission
                if (employees.length > 0 || tasks.length > 0) {
                    dataAlreadySubmitted = false;
                }
            } else {
                showMessage('Project created successfully!', 'success');
            }
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

    if (employees.length === 0) {
        list.innerHTML = '<p style="color: #666; font-style: italic;">No employees added yet.</p>';
    } else {
        list.innerHTML = employees.map((emp, idx) =>
            `<div class="item-badge" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                <span>${emp.name} (Works: ${emp.work_pattern.map(d => ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][d]).join(', ')})</span>
                <div>
                    <button onclick="editEmployee(${idx})" class="btn-small" style="margin-left: 5px;">Edit</button>
                    <button onclick="deleteEmployee(${idx})" class="btn-small btn-danger" style="margin-left: 5px;">Delete</button>
                </div>
            </div>`
        ).join('');
    }

    // Update task assignment dropdown
    const select = document.getElementById('taskAssignedTo');
    select.innerHTML = '<option value="">Select employee</option>' +
        employees.map(emp => `<option value="${emp.name}">${emp.name}</option>`).join('');
}

function editEmployee(idx) {
    const emp = employees[idx];

    // Populate the form with employee data
    document.getElementById('employeeName').value = emp.name;
    document.getElementById('employeeHolidays').value = emp.holidays.join(', ');

    // Check the appropriate work pattern checkboxes
    const checkboxes = document.querySelectorAll('.checkbox-group input[type="checkbox"]');
    checkboxes.forEach(cb => {
        cb.checked = emp.work_pattern.includes(parseInt(cb.value));
    });

    // Remove the employee from the array (will be re-added when user clicks "Add Employee")
    employees.splice(idx, 1);
    dataAlreadySubmitted = false; // Mark that data needs to be resubmitted
    updateEmployeesList();

    showMessage('Edit employee details and click "Add Employee" to save changes', 'info');
}

function deleteEmployee(idx) {
    const emp = employees[idx];
    if (confirm(`Are you sure you want to delete employee "${emp.name}"?`)) {
        employees.splice(idx, 1);
        dataAlreadySubmitted = false; // Mark that data needs to be resubmitted
        updateEmployeesList();
        showMessage(`Employee "${emp.name}" deleted`, 'success');
    }
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
    const customStartDate = document.getElementById('taskCustomStartDate').value;

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
        dependency: dependency || null,
        custom_start_date: customStartDate || null
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
    document.getElementById('taskCustomStartDate').value = '';

    showMessage(`Task "${name}" added!`, 'success');
}

function updateTasksList() {
    const list = document.getElementById('tasksList');

    if (tasks.length === 0) {
        list.innerHTML = '<p style="color: #666; font-style: italic;">No tasks added yet.</p>';
    } else {
        list.innerHTML = tasks.map((task, idx) =>
            `<div class="item-badge" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                <span>${idx + 1}. ${task.name} (${task.assigned_to}, ${task.estimated_duration}d, ${task.availability}% avail.)</span>
                <div>
                    <button onclick="editTask(${idx})" class="btn-small" style="margin-left: 5px;">Edit</button>
                    <button onclick="deleteTask(${idx})" class="btn-small btn-danger" style="margin-left: 5px;">Delete</button>
                </div>
            </div>`
        ).join('');
    }

    // Update dependency dropdown
    const select = document.getElementById('taskDependency');
    select.innerHTML = '<option value="">No dependency</option>' +
        tasks.map(task => `<option value="${task.name}">${task.name}</option>`).join('');
}

function editTask(idx) {
    const task = tasks[idx];

    // Populate the form with task data
    document.getElementById('taskName').value = task.name;
    document.getElementById('taskAssignedTo').value = task.assigned_to;
    document.getElementById('taskDuration').value = task.estimated_duration;
    document.getElementById('taskAvailability').value = task.availability;
    document.getElementById('taskContingency').value = task.contingency_margin;
    document.getElementById('taskDependency').value = task.dependency || '';
    document.getElementById('taskCustomStartDate').value = task.custom_start_date || '';

    // Remove the task from the array (will be re-added when user clicks "Add Task")
    tasks.splice(idx, 1);
    dataAlreadySubmitted = false; // Mark that data needs to be resubmitted
    updateTasksList();

    showMessage('Edit task details and click "Add Task" to save changes', 'info');
}

function deleteTask(idx) {
    const task = tasks[idx];
    if (confirm(`Are you sure you want to delete task "${task.name}"?`)) {
        tasks.splice(idx, 1);
        dataAlreadySubmitted = false; // Mark that data needs to be resubmitted
        updateTasksList();
        showMessage(`Task "${task.name}" deleted`, 'success');
    }
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
    if (!dataAlreadySubmitted) {
        if (!(await submitEmployees())) return;
        if (!(await submitTasks())) return;
        dataAlreadySubmitted = true; // Mark as submitted
    }

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

    // Sort tasks by start date (earliest first)
    const sortedTasks = [...data.tasks].sort((a, b) => {
        if (!a.start_date) return 1;  // Tasks without start date go to end
        if (!b.start_date) return -1;
        return new Date(a.start_date) - new Date(b.start_date);
    });

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
    sortedTasks.forEach(task => {
        html += '<tr>';
        html += `<td>${task.name}</td>`;
        html += `<td>${task.assigned_to}</td>`;
        html += `<td>${task.estimated_duration}</td>`;
        html += `<td>${task.actual_duration}</td>`;
        html += `<td>${task.start_date || ''}</td>`;
        html += `<td>${task.end_date || ''}</td>`;

        const workingDatesSet = new Set(task.working_dates);
        const holidayDatesSet = new Set(task.holiday_dates);

        dates.forEach(date => {
            const dateStr = date.toISOString().split('T')[0];
            const isWorking = workingDatesSet.has(dateStr);
            const isHoliday = holidayDatesSet.has(dateStr);

            let cellClass = 'date-cell';
            if (isWorking) {
                cellClass += ' working';
            } else if (isHoliday) {
                cellClass += ' holiday';
            }

            html += `<td class="${cellClass}"></td>`;
        });

        html += '</tr>';
    });

    html += '</table>';
    container.innerHTML = html;
}

// Recalculate schedule (for use on Step 4)
async function recalculateSchedule() {
    // Always resubmit employees and tasks to ensure backend has latest data
    dataAlreadySubmitted = false;

    // Call calculateAndView but stay on step 4
    await calculateAndView();
}

// Export to Excel
async function exportToExcel() {
    // Prompt user for filename
    const defaultFilename = document.getElementById('projectName').value.replace(/\s+/g, '_') + '_gantt';
    const userFilename = prompt('Enter a name for the Excel file:', defaultFilename);

    // If user cancels, return
    if (userFilename === null) {
        return;
    }

    // Use default if user enters empty string
    const filename = userFilename.trim() || defaultFilename;

    try {
        const response = await fetch('/api/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filename })
        });

        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            // Ensure .xlsx extension for download
            a.download = filename.endsWith('.xlsx') ? filename : filename + '.xlsx';
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
        dataAlreadySubmitted = false; // Reset flag
        projectCreated = false; // Reset flag

        // Clear all forms
        document.getElementById('projectName').value = '';
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('startDate').value = today;
        document.getElementById('globalHolidays').value = '';
        document.getElementById('nextToStep2').style.display = 'none';

        // Reset button text
        const createBtn = document.querySelector('#step1 .btn-primary');
        if (createBtn) {
            createBtn.textContent = 'Create Project';
        }

        updateEmployeesList();
        updateTasksList();

        goToStep(1);
        showMessage('Project reset successfully!', 'success');
    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
    }
}

// File import handling
document.addEventListener('DOMContentLoaded', () => {
    // Set default date to today
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('startDate').value = today;

    // Handle file selection
    document.getElementById('importFile').addEventListener('change', function() {
        const fileName = this.files.length > 0 ? this.files[0].name : '';
        document.getElementById('importFileName').textContent = fileName;
        document.getElementById('importBtn').disabled = !fileName;
    });
});

// Import project from Excel file
async function importProject() {
    const fileInput = document.getElementById('importFile');
    const file = fileInput.files[0];

    if (!file) {
        showMessage('Please select a file to import', 'error');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    try {
        showMessage('Importing project...', 'info');

        const response = await fetch('/api/import', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (!response.ok) {
            showMessage(data.error || 'Error importing file', 'error');
            return;
        }

        // Successfully imported - populate the app
        await populateFromImport(data);

        showMessage('Project imported successfully! You can now edit and continue.', 'success');

    } catch (error) {
        showMessage('Network error: ' + error.message, 'error');
    }
}

// Populate app state from imported data
async function populateFromImport(data) {
    // Reset current state
    employees = [];
    tasks = [];

    // Step 1: Create project with imported info
    const projectInfo = data.project_info;
    document.getElementById('projectName').value = projectInfo.name;
    document.getElementById('startDate').value = projectInfo.start_date;

    if (projectInfo.global_holidays && projectInfo.global_holidays.length > 0) {
        document.getElementById('globalHolidays').value = projectInfo.global_holidays.join(', ');
    } else {
        document.getElementById('globalHolidays').value = '';
    }

    // Create the project on backend
    const projectResponse = await fetch('/api/project', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            name: projectInfo.name,
            start_date: projectInfo.start_date,
            global_holidays: projectInfo.global_holidays || []
        })
    });

    if (!projectResponse.ok) {
        const errorData = await projectResponse.json();
        throw new Error(errorData.error || 'Failed to create project');
    }

    // Step 2: Add employees
    employees = data.employees;
    updateEmployeesList();

    // Submit employees to backend
    const employeesResponse = await fetch('/api/employees', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ employees })
    });

    if (!employeesResponse.ok) {
        const errorData = await employeesResponse.json();
        throw new Error(errorData.error || 'Failed to add employees');
    }

    // Step 3: Add tasks
    tasks = data.tasks;
    updateTasksList();

    // Submit tasks to backend
    const tasksResponse = await fetch('/api/tasks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tasks })
    });

    if (!tasksResponse.ok) {
        const errorData = await tasksResponse.json();
        throw new Error(errorData.error || 'Failed to add tasks');
    }

    // Mark data as already submitted since we sent it to backend
    dataAlreadySubmitted = true;
    projectCreated = true;

    // Show the next button on step 1
    document.getElementById('nextToStep2').style.display = 'inline-block';

    // Navigate to step 3 (tasks) so user can edit
    goToStep(3);
}
