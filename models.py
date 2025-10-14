from datetime import datetime, timedelta
from typing import List, Dict, Optional, Set


class Employee:
    """Represents an employee with their work schedule and holidays."""

    def __init__(self, name: str):
        self.name = name
        # Work pattern: list of weekday numbers (0=Monday, 4=Friday)
        # Default is Mon-Fri (0,1,2,3,4)
        self.work_pattern: List[int] = [0, 1, 2, 3, 4]
        # Individual holidays: set of date strings in 'YYYY-MM-DD' format
        self.holidays: Set[str] = set()

    def set_work_pattern(self, work_days: List[int]):
        """Set which days of the week the employee works (0=Monday, 6=Sunday)."""
        self.work_pattern = sorted(work_days)

    def add_holiday(self, date_str: str):
        """Add a holiday date for this employee (format: 'YYYY-MM-DD')."""
        self.holidays.add(date_str)

    def is_working_day(self, date: datetime, global_holidays: Set[str]) -> bool:
        """Check if the employee works on a given date."""
        date_str = date.strftime('%Y-%m-%d')

        # Check if it's a holiday (global or personal)
        if date_str in global_holidays or date_str in self.holidays:
            return False

        # Check if it's in their work pattern
        return date.weekday() in self.work_pattern


class Task:
    """Represents a task in the Gantt chart."""

    def __init__(self, name: str, estimated_duration: int, assigned_to: str):
        self.name = name
        self.estimated_duration = estimated_duration  # in days
        self.assigned_to = assigned_to  # employee name
        self.dependency: Optional[str] = None  # name of antecedent task
        self.availability: int = 100  # percentage (1-100)
        self.contingency_margin: int = 0  # percentage (0+)

        # Calculated fields
        self.actual_duration: int = 0
        self.start_date: Optional[datetime] = None
        self.end_date: Optional[datetime] = None
        self.working_dates: List[datetime] = []

    def calculate_actual_duration(self) -> int:
        """Calculate actual duration based on availability and contingency margin."""
        # Formula: actual_duration = estimated_duration / availability * 100 * (1 + contingency_margin/100)
        duration = (self.estimated_duration / self.availability * 100) * (1 + self.contingency_margin / 100)
        return round(duration)


class Project:
    """Represents the entire project with tasks, employees, and scheduling."""

    def __init__(self, name: str, start_date: datetime):
        self.name = name
        self.start_date = start_date
        self.tasks: List[Task] = []
        self.employees: Dict[str, Employee] = {}
        self.global_holidays: Set[str] = set()

    def add_employee(self, employee: Employee):
        """Add an employee to the project."""
        self.employees[employee.name] = employee

    def add_task(self, task: Task):
        """Add a task to the project."""
        self.tasks.append(task)

    def add_global_holiday(self, date_str: str):
        """Add a global holiday (format: 'YYYY-MM-DD')."""
        self.global_holidays.add(date_str)

    def get_next_working_day(self, employee: Employee, from_date: datetime) -> datetime:
        """Get the next working day for an employee starting from a given date."""
        current = from_date
        while not employee.is_working_day(current, self.global_holidays):
            current += timedelta(days=1)
        return current

    def calculate_task_schedule(self, task: Task, earliest_start: datetime):
        """Calculate the schedule for a single task."""
        # Get the employee
        employee = self.employees.get(task.assigned_to)
        if not employee:
            raise ValueError(f"Employee '{task.assigned_to}' not found")

        # Calculate actual duration
        task.actual_duration = task.calculate_actual_duration()

        # Find the first working day from earliest_start
        current_date = self.get_next_working_day(employee, earliest_start)
        task.start_date = current_date

        # Count working days until we reach actual_duration
        working_days_count = 0
        task.working_dates = []

        while working_days_count < task.actual_duration:
            if employee.is_working_day(current_date, self.global_holidays):
                task.working_dates.append(current_date)
                working_days_count += 1
                task.end_date = current_date
            current_date += timedelta(days=1)

    def calculate_schedule(self):
        """Calculate the schedule for all tasks, respecting dependencies."""
        # Build a map of task names to tasks
        task_map = {task.name: task for task in self.tasks}

        # Track completed tasks
        scheduled_tasks = set()

        # Keep scheduling tasks until all are done
        while len(scheduled_tasks) < len(self.tasks):
            made_progress = False

            for task in self.tasks:
                if task.name in scheduled_tasks:
                    continue

                # Check if dependency is met
                if task.dependency:
                    if task.dependency not in task_map:
                        raise ValueError(f"Dependency '{task.dependency}' not found for task '{task.name}'")

                    dep_task = task_map[task.dependency]
                    if dep_task.name not in scheduled_tasks:
                        continue  # Wait for dependency

                    # Start the next available working day after dependency ends
                    earliest_start = dep_task.end_date + timedelta(days=1)
                else:
                    # No dependency, start from project start date
                    earliest_start = self.start_date

                # Calculate this task's schedule
                self.calculate_task_schedule(task, earliest_start)
                scheduled_tasks.add(task.name)
                made_progress = True

            if not made_progress:
                raise ValueError("Circular dependency detected or invalid dependency chain")

    def get_project_end_date(self) -> Optional[datetime]:
        """Get the end date of the entire project."""
        if not self.tasks:
            return None
        return max(task.end_date for task in self.tasks if task.end_date)

    def get_date_range(self) -> tuple:
        """Get the full date range of the project."""
        if not self.tasks:
            return (self.start_date, self.start_date)

        end_date = self.get_project_end_date()
        return (self.start_date, end_date)
