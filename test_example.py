"""
Example test script to verify GanttQuick functionality
"""
from datetime import datetime
from models import Project, Task, Employee
from excel_export import export_to_excel


def main():
    # Create a project
    project = Project("Website Redesign", datetime(2024, 1, 8))  # Monday

    # Add global holidays
    project.add_global_holiday("2024-01-15")  # MLK Day

    # Add employees with different work patterns
    alice = Employee("Alice")
    alice.set_work_pattern([0, 1, 2, 3, 4])  # Mon-Fri
    alice.add_holiday("2024-01-10")  # Personal holiday
    project.add_employee(alice)

    bob = Employee("Bob")
    bob.set_work_pattern([0, 1, 2, 3, 4])  # Mon-Fri
    project.add_employee(bob)

    # Add tasks with dependencies
    task1 = Task("Requirements Gathering", 3, "Alice")
    task1.availability = 100
    task1.contingency_margin = 10
    project.add_task(task1)

    task2 = Task("Design Mockups", 5, "Bob")
    task2.dependency = "Requirements Gathering"
    task2.availability = 80
    task2.contingency_margin = 0
    project.add_task(task2)

    task3 = Task("Frontend Development", 10, "Alice")
    task3.dependency = "Design Mockups"
    task3.availability = 100
    task3.contingency_margin = 20
    project.add_task(task3)

    task4 = Task("Backend Development", 8, "Bob")
    task4.dependency = "Design Mockups"
    task4.availability = 100
    task4.contingency_margin = 15
    project.add_task(task4)

    task5 = Task("Testing & QA", 5, "Alice")
    task5.dependency = "Frontend Development"
    task5.availability = 100
    task5.contingency_margin = 0
    project.add_task(task5)

    # Calculate schedule
    print("Calculating project schedule...")
    project.calculate_schedule()

    # Print results
    print(f"\nProject: {project.name}")
    print(f"Start Date: {project.start_date.strftime('%Y-%m-%d')}")
    print(f"End Date: {project.get_project_end_date().strftime('%Y-%m-%d')}")
    print(f"\nTasks:")
    print("-" * 100)

    for task in project.tasks:
        print(f"\nTask: {task.name}")
        print(f"  Assigned to: {task.assigned_to}")
        print(f"  Estimated Duration: {task.estimated_duration} days")
        print(f"  Availability: {task.availability}%")
        print(f"  Contingency Margin: {task.contingency_margin}%")
        print(f"  Actual Duration: {task.actual_duration} days")
        print(f"  Start Date: {task.start_date.strftime('%Y-%m-%d') if task.start_date else 'N/A'}")
        print(f"  End Date: {task.end_date.strftime('%Y-%m-%d') if task.end_date else 'N/A'}")
        print(f"  Working Days: {len(task.working_dates)}")

    # Export to Excel
    print("\n\nExporting to Excel...")
    filename = export_to_excel(project, "test_gantt_chart.xlsx")
    print(f"Excel file created: {filename}")


if __name__ == "__main__":
    main()
