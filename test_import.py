#!/usr/bin/env python3
"""
Test script for Excel import functionality
"""

from datetime import datetime
from models import Project, Task, Employee
from excel_export import export_to_excel
from excel_import import import_from_excel, ExcelImportError

def test_import():
    """Test the import functionality by creating a project, exporting it, and importing it back."""

    print("=" * 60)
    print("Testing Excel Import Functionality")
    print("=" * 60)

    # Step 1: Create a test project
    print("\n1. Creating test project...")
    project = Project("Test Import Project", datetime(2025, 1, 15))
    project.add_global_holiday("2025-01-20")
    project.add_global_holiday("2025-02-14")

    # Add employees
    emp1 = Employee("Alice")
    emp1.set_work_pattern([0, 1, 2, 3, 4])  # Mon-Fri
    emp1.add_holiday("2025-01-22")

    emp2 = Employee("Bob")
    emp2.set_work_pattern([0, 1, 2, 3])  # Mon-Thu

    project.add_employee(emp1)
    project.add_employee(emp2)

    # Add tasks
    task1 = Task("Design Phase", 5, "Alice")
    task1.availability = 80
    task1.contingency_margin = 10

    task2 = Task("Development", 10, "Bob")
    task2.dependency = "Design Phase"
    task2.availability = 100
    task2.contingency_margin = 20

    task3 = Task("Testing", 3, "Alice")
    task3.dependency = "Development"
    task3.custom_start_date = datetime(2025, 2, 10)

    project.add_task(task1)
    project.add_task(task2)
    project.add_task(task3)

    # Calculate schedule
    project.calculate_schedule()

    print(f"   Project: {project.name}")
    print(f"   Start Date: {project.start_date.strftime('%Y-%m-%d')}")
    print(f"   Employees: {len(project.employees)}")
    print(f"   Tasks: {len(project.tasks)}")
    print(f"   Global Holidays: {len(project.global_holidays)}")

    # Step 2: Export to Excel
    print("\n2. Exporting to Excel...")
    test_filename = "test_import_sample.xlsx"
    export_to_excel(project, test_filename)
    print(f"   Exported to: {test_filename}")

    # Step 3: Import from Excel
    print("\n3. Importing from Excel...")
    try:
        imported_data = import_from_excel(test_filename)
        print("   Import successful!")

        # Verify imported data
        print("\n4. Verifying imported data...")

        # Check project info
        project_info = imported_data['project_info']
        print(f"   Project Name: {project_info['name']}")
        assert project_info['name'] == "Test Import Project", "Project name mismatch"
        print(f"   Start Date: {project_info['start_date']}")
        assert project_info['start_date'] == "2025-01-15", "Start date mismatch"
        print(f"   Global Holidays: {project_info['global_holidays']}")
        assert len(project_info['global_holidays']) == 2, "Global holidays count mismatch"

        # Check employees
        employees = imported_data['employees']
        print(f"\n   Employees: {len(employees)}")
        assert len(employees) == 2, "Employee count mismatch"

        alice = next(emp for emp in employees if emp['name'] == 'Alice')
        print(f"   - {alice['name']}: works {len(alice['work_pattern'])} days/week, {len(alice['holidays'])} holidays")
        assert alice['work_pattern'] == [0, 1, 2, 3, 4], "Alice work pattern mismatch"
        assert len(alice['holidays']) == 1, "Alice holidays count mismatch"

        bob = next(emp for emp in employees if emp['name'] == 'Bob')
        print(f"   - {bob['name']}: works {len(bob['work_pattern'])} days/week, {len(bob['holidays'])} holidays")
        assert bob['work_pattern'] == [0, 1, 2, 3], "Bob work pattern mismatch"

        # Check tasks
        tasks = imported_data['tasks']
        print(f"\n   Tasks: {len(tasks)}")
        assert len(tasks) == 3, "Task count mismatch"

        design = tasks[0]
        print(f"   - {design['name']}: {design['estimated_duration']}d, {design['availability']}% avail, {design['contingency_margin']}% margin")
        assert design['name'] == "Design Phase", "Task 1 name mismatch"
        assert design['estimated_duration'] == 5, "Task 1 duration mismatch"
        assert design['availability'] == 80, "Task 1 availability mismatch"
        assert design['contingency_margin'] == 10, "Task 1 contingency mismatch"
        assert design['assigned_to'] == "Alice", "Task 1 assignment mismatch"

        dev = tasks[1]
        print(f"   - {dev['name']}: {dev['estimated_duration']}d, depends on '{dev['dependency']}'")
        assert dev['dependency'] == "Design Phase", "Task 2 dependency mismatch"
        assert dev['contingency_margin'] == 20, "Task 2 contingency mismatch"

        testing = tasks[2]
        print(f"   - {testing['name']}: {testing['estimated_duration']}d, custom start: {testing['custom_start_date']}")
        assert testing['custom_start_date'] == "2025-02-10", "Task 3 custom start date mismatch"

        print("\n" + "=" * 60)
        print("ALL TESTS PASSED!")
        print("=" * 60)

        return True

    except ExcelImportError as e:
        print(f"   ERROR: {e}")
        return False
    except Exception as e:
        print(f"   UNEXPECTED ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    success = test_import()
    exit(0 if success else 1)
