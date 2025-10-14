from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime, timedelta
from models import Project


def export_to_excel(project: Project, filename: str = "gantt_chart.xlsx"):
    """Export the project Gantt chart to an Excel file."""

    wb = Workbook()
    ws = wb.active
    ws.title = "Gantt Chart"

    # Define styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    task_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Get date range
    start_date, end_date = project.get_date_range()
    if not end_date:
        end_date = start_date

    # Create date list (all calendar days from start to end)
    date_list = []
    current = start_date
    while current <= end_date:
        date_list.append(current)
        current += timedelta(days=1)

    # Header row 1: Fixed columns
    headers = ["Task Name", "Assigned To", "Estimated Duration", "Actual Duration", "Start Date", "End Date"]
    col_offset = len(headers) + 1  # +1 for spacing

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Header row 1: Date columns
    for idx, date in enumerate(date_list):
        col = col_offset + idx
        cell = ws.cell(row=1, column=col)
        cell.value = date.strftime("%m/%d")
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
        cell.border = border
        ws.column_dimensions[cell.column_letter].width = 3

    # Header row 2: Day of week
    for idx, date in enumerate(date_list):
        col = col_offset + idx
        cell = ws.cell(row=2, column=col)
        cell.value = date.strftime("%a")[0]  # M, T, W, T, F, S, S
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Set column widths for fixed columns
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 12

    # Data rows
    for task_idx, task in enumerate(project.tasks, start=3):
        # Fixed columns
        ws.cell(row=task_idx, column=1).value = task.name
        ws.cell(row=task_idx, column=2).value = task.assigned_to
        ws.cell(row=task_idx, column=3).value = task.estimated_duration
        ws.cell(row=task_idx, column=4).value = task.actual_duration
        ws.cell(row=task_idx, column=5).value = task.start_date.strftime("%Y-%m-%d") if task.start_date else ""
        ws.cell(row=task_idx, column=6).value = task.end_date.strftime("%Y-%m-%d") if task.end_date else ""

        # Apply borders to fixed columns
        for col in range(1, 7):
            ws.cell(row=task_idx, column=col).border = border
            ws.cell(row=task_idx, column=col).alignment = Alignment(vertical="center")

        # Date columns - highlight working days
        working_dates_set = {d.strftime("%Y-%m-%d") for d in task.working_dates}
        for idx, date in enumerate(date_list):
            col = col_offset + idx
            cell = ws.cell(row=task_idx, column=col)
            cell.border = border

            date_str = date.strftime("%Y-%m-%d")
            if date_str in working_dates_set:
                cell.fill = task_fill

    # Freeze panes (freeze first 6 columns and first 2 rows)
    ws.freeze_panes = ws.cell(row=3, column=7)

    # Add project info sheet
    info_ws = wb.create_sheet("Project Info")
    info_ws['A1'] = "Project Name:"
    info_ws['B1'] = project.name
    info_ws['A2'] = "Start Date:"
    info_ws['B2'] = project.start_date.strftime("%Y-%m-%d")
    info_ws['A3'] = "End Date:"
    info_ws['B3'] = end_date.strftime("%Y-%m-%d") if end_date else ""
    info_ws['A4'] = "Total Duration (days):"
    info_ws['B4'] = (end_date - project.start_date).days + 1 if end_date else 0

    # Bold the labels
    for row in range(1, 5):
        info_ws.cell(row=row, column=1).font = Font(bold=True)

    info_ws.column_dimensions['A'].width = 25
    info_ws.column_dimensions['B'].width = 25

    # Save the file
    wb.save(filename)
    return filename
