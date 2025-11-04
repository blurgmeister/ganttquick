# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a new repository for "ganttquick" - currently empty with no codebase initialized.
This project aims to develop a simple and easy to use gantt charting software that exports an excel file showing the gantt chart.
The key features it needs are:
- let user name the project this is a gantt chart for
- allow the user to create and name tasks, in a vertical list preferably
- allow user to enter the estimated duration for each task (in whole/integer number days) 
- allow user to specify for each task the immediately previous/antecedant task it depends on for completion so it can then begin (each task begins the next avaialble working day after an antecedant task finishes)
- allow user to enter the name of the employee assigned to do the task
- allow user to input a work pattern (over a 2 week window using monday to friday as the available days). Specifying which days of the week each user works.
- allow user to input a holiday schedule globally, and for each user individually if so desired
- allow user to input a starting date for the project
- allow user to enter a percentage availablity (intereger between 1 and 100) to act as a multiplier on the estimated duration input. This reflects the allocated employees availability to do the work. For example a 50% availability means the actual duration of work will be double. 100% will be the default value set for the user the change as they wish.
- allow user to enter a contingency margin percentage (integer equal or exceeding 0 is allowed). This acts a straight multiplier on top of the duration adjusted for availability. The default value for the contingency margin is 0%. Which the user can modify if desired. 
- The actual duration to allocate in the gantt chart will use a formula like this: actual_duration = estimated_duration/availability * 100 * (1 + contingency_margin/100), rounded to the nearest whole number. 
- then the app must calculate the actual duration (accounting for availability and contingency margin) from the estimated duration the user input
- the app must then calculate the start dates and end dates of each task, accounting for the specified task dependencies, and project start date, and critically adjusting for the work schedules specified for each employee and adjusting for global and individual employee holidays as specified
- the app must then show the gantt chart with each row becorresponding to a task, and each column showing the name of the task, who it is assigned to, the task duration input, the actual duration calculated for it by the app, its calculated start and end dates and highlighting each working day being worked to complete the task. The chart will extend out to the right for as many days as necessary to complete the project.
- the app must include an export button that will save the gantt chart with all those details as an excel file.


## Current State

**Status**: IMPLEMENTED AND FUNCTIONAL

The GanttQuick application has been successfully implemented with all required features.

### Architecture

**Tech Stack**:
- Backend: Python 3 with Flask web framework
- Frontend: Vanilla JavaScript with HTML/CSS
- Excel Export: openpyxl library
- Date Handling: Python datetime and dateutil

**File Structure**:
```
ganttquick/
├── app.py                 # Flask application with REST API
├── models.py              # Data models (Project, Task, Employee)
├── excel_export.py        # Excel export functionality
├── excel_import.py        # Excel import functionality (NEW)
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Main UI (4-step wizard)
└── static/
    ├── css/
    │   └── style.css     # Application styles
    └── js/
        └── app.js        # Frontend logic
```

### Key Components

1. **models.py**: Core business logic
   - `Employee` class: Manages employee work patterns and holidays
   - `Task` class: Handles task properties and duration calculations
   - `Project` class: Orchestrates scheduling and dependency resolution

2. **app.py**: REST API endpoints
   - `/api/project` - Create/update project
   - `/api/employees` - Add employees
   - `/api/tasks` - Add tasks
   - `/api/calculate` - Calculate schedule
   - `/api/gantt` - Get chart data
   - `/api/export` - Export to Excel
   - `/api/import` - Import from Excel (NEW)
   - `/api/reset` - Reset current project

3. **excel_export.py**: Excel generation
   - Creates formatted Gantt chart with colored cells
   - Includes 4 sheets: Gantt Chart, Project Info, Work Schedules, Holiday Schedule
   - Freezes header rows/columns for easy navigation

4. **excel_import.py**: Excel import (NEW)
   - Validates required tabs (Gantt Chart, Project Info, Work Schedules)
   - Optionally reads Holiday Schedule tab if present
   - Extracts project info, employees, tasks, and all configuration
   - Provides detailed error messages for validation failures

### Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py

# Access at http://localhost:5000
```

### Testing

```bash
# Run original test example
python test_example.py

# Test import functionality
python test_import.py

# Test import error handling
python test_import_errors.py
```

### Features Implemented

All required features have been implemented:
- Project naming and start date configuration
- Task creation with vertical list display
- Task dependency management (antecedent tasks)
- Employee assignment per task
- Customizable work patterns (2-week window, Mon-Fri defaults)
- Global and per-employee holiday schedules
- Availability percentage (1-100%, default 100%)
- Contingency margin percentage (0+%, default 0%)
- Automatic actual duration calculation: `actual_duration = estimated_duration/availability * 100 * (1 + contingency_margin/100)`
- Start/end date calculation with dependency resolution
- Work schedule and holiday adjustment
- Visual Gantt chart with task details
- Excel export with formatted chart and highlighted working days
- **Excel import functionality (NEW)**:
  - Import previously exported Excel files
  - Validates file structure (required tabs: Gantt Chart, Project Info, Work Schedules)
  - Extracts all project data including tasks, employees, schedules, and holidays
  - Allows users to continue editing imported projects
  - Comprehensive error handling and validation

### Design Patterns

- **REST API**: Stateless endpoints for frontend communication
- **Model-View separation**: Business logic in models, UI in templates
- **Step-by-step wizard**: Guides users through project setup
- **Date calculation engine**: Handles complex scheduling logic with dependencies
