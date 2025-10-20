from flask import Flask, render_template, request, jsonify, send_file
from datetime import datetime
import json
import os
from models import Project, Task, Employee
from excel_export import export_to_excel

app = Flask(__name__)

# In-memory storage for the current project
current_project = None


@app.route('/')
def index():
    """Render the main application page."""
    return render_template('index.html')


@app.route('/api/project', methods=['POST'])
def create_project():
    """Create or update the project with basic information."""
    global current_project

    data = request.json
    project_name = data.get('name', 'Unnamed Project')
    start_date_str = data.get('start_date')

    try:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
    except (ValueError, TypeError):
        return jsonify({'error': 'Invalid start date format. Use YYYY-MM-DD'}), 400

    current_project = Project(project_name, start_date)

    # Add global holidays if provided
    global_holidays = data.get('global_holidays', [])
    for holiday in global_holidays:
        current_project.add_global_holiday(holiday)

    return jsonify({'message': 'Project created successfully', 'name': project_name})


@app.route('/api/employees', methods=['POST'])
def add_employees():
    """Add employees to the project."""
    global current_project

    if not current_project:
        return jsonify({'error': 'Project not initialized'}), 400

    data = request.json
    employees_data = data.get('employees', [])

    for emp_data in employees_data:
        emp = Employee(emp_data['name'])

        # Set work pattern if provided (list of weekday numbers)
        if 'work_pattern' in emp_data:
            emp.set_work_pattern(emp_data['work_pattern'])

        # Add individual holidays if provided
        if 'holidays' in emp_data:
            for holiday in emp_data['holidays']:
                emp.add_holiday(holiday)

        current_project.add_employee(emp)

    return jsonify({'message': f'{len(employees_data)} employee(s) added successfully'})


@app.route('/api/tasks', methods=['POST'])
def add_tasks():
    """Add tasks to the project."""
    global current_project

    if not current_project:
        return jsonify({'error': 'Project not initialized'}), 400

    data = request.json
    tasks_data = data.get('tasks', [])

    for task_data in tasks_data:
        task = Task(
            name=task_data['name'],
            estimated_duration=task_data['estimated_duration'],
            assigned_to=task_data['assigned_to']
        )

        # Set optional fields
        if 'dependency' in task_data and task_data['dependency']:
            task.dependency = task_data['dependency']

        task.availability = task_data.get('availability', 100)
        task.contingency_margin = task_data.get('contingency_margin', 0)

        current_project.add_task(task)

    return jsonify({'message': f'{len(tasks_data)} task(s) added successfully'})


@app.route('/api/calculate', methods=['POST'])
def calculate_schedule():
    """Calculate the project schedule."""
    global current_project

    if not current_project:
        return jsonify({'error': 'Project not initialized'}), 400

    if not current_project.tasks:
        return jsonify({'error': 'No tasks defined'}), 400

    try:
        current_project.calculate_schedule()
        return jsonify({'message': 'Schedule calculated successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 400


@app.route('/api/gantt', methods=['GET'])
def get_gantt_data():
    """Get the Gantt chart data for visualization."""
    global current_project

    if not current_project:
        return jsonify({'error': 'Project not initialized'}), 400

    # Prepare data for frontend
    tasks_data = []
    for task in current_project.tasks:
        task_info = {
            'name': task.name,
            'assigned_to': task.assigned_to,
            'estimated_duration': task.estimated_duration,
            'actual_duration': task.actual_duration,
            'availability': task.availability,
            'contingency_margin': task.contingency_margin,
            'dependency': task.dependency,
            'start_date': task.start_date.strftime('%Y-%m-%d') if task.start_date else None,
            'end_date': task.end_date.strftime('%Y-%m-%d') if task.end_date else None,
            'working_dates': [d.strftime('%Y-%m-%d') for d in task.working_dates]
        }
        tasks_data.append(task_info)

    start_date, end_date = current_project.get_date_range()

    response = {
        'project_name': current_project.name,
        'start_date': start_date.strftime('%Y-%m-%d'),
        'end_date': end_date.strftime('%Y-%m-%d') if end_date else None,
        'tasks': tasks_data
    }

    return jsonify(response)


@app.route('/api/export', methods=['POST'])
def export_excel():
    """Export the Gantt chart to Excel."""
    global current_project

    if not current_project:
        return jsonify({'error': 'Project not initialized'}), 400

    try:
        data = request.json
        custom_filename = data.get('filename', '').strip()

        # Use custom filename or default to project name
        if custom_filename:
            # Add .xlsx extension if not provided
            if not custom_filename.endswith('.xlsx'):
                filename = f"{custom_filename}.xlsx"
            else:
                filename = custom_filename
        else:
            filename = f"{current_project.name.replace(' ', '_')}_gantt.xlsx"

        # Sanitize filename (remove potentially problematic characters)
        filename = filename.replace('/', '_').replace('\\', '_')

        filepath = export_to_excel(current_project, filename)
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/reset', methods=['POST'])
def reset_project():
    """Reset the current project."""
    global current_project
    current_project = None
    return jsonify({'message': 'Project reset successfully'})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
