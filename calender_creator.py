import xlsxwriter
import datetime
from constants import *
from shiftPlanner_class import ShiftPlanner
import json
from employee_class import Employee


# get the max number of employees which are working in a area from the year.
# is needed to create the excel file. so each sheet looks the same
def find_max_employee_count(data, area_in):
    max_count = 0

    for week_data in data.values():
        for shift_data in week_data.values():
            for area, area_data in shift_data.items():
                if area == area_in:
                    max_count = max(max_count, len(area_data))

    return max_count


def load_employees():
    try:
        with open(FILE_NAME, "r", encoding="utf-8") as file:
            data = json.load(file)
            employees_data = data["employees"]
    except FileNotFoundError:
        return print("Keine Mitarbeiterliste verfügbar")

    employees = []

    # Create the employees classes
    list_of_last_shift_in_model_2 = []
    for employee_data in employees_data:
        employee = Employee(
            employee_data[EMPLOYEE_KEY[0]],
            employee_data[EMPLOYEE_KEY[1]],
            employee_data[EMPLOYEE_KEY[2]],
            employee_data[EMPLOYEE_KEY[3]],
            employee_data[EMPLOYEE_KEY[4]],
            employee_data[EMPLOYEE_KEY[5]],
        )
        # set the start shift for the emp with the schichtmodel 2
        if employee.schicht_model == SCHICHT_MODELS[2]:
            result = employee.set_start_shift(list_of_last_shift_in_model_2)
            if result in list_of_last_shift_in_model_2:
                list_of_last_shift_in_model_2.remove(result)
            list_of_last_shift_in_model_2.insert(0, result)

        # Add emp to list
        employees.append(employee)

    return employees


def add_days_of_the_week_to_worksheet(worksheet, bordered_format, col, row, start_date_of_the_week):
    for day in range(7):
        start_date_of_the_week_temp = start_date_of_the_week + datetime.timedelta(days=day)

        worksheet.write(
            row,
            col + day,
            start_date_of_the_week_temp.strftime("%A")[:2],
            bordered_format,
        )

        worksheet.write(
            row + 1,
            col + day,
            start_date_of_the_week_temp.strftime("%d.%m"),
            bordered_format,
        )


def create_calender(year, start_week=1, end_week=52):
    # Erstelle eine Excel-Datei
    workbook = xlsxwriter.Workbook(f"Schichtplan_{year}.xlsx")

    # Definiere einen Formatierer mit Rahmen
    bordered_format = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})

    employees = load_employees()

    plan = ShiftPlanner(employees, start_week, end_week)

    shift_plan = plan.get_shift_plan()
    employees = plan.employees

    max_len_areas = {
        BEREICHE[1]: find_max_employee_count(shift_plan, BEREICHE[1]),
        BEREICHE[2]: find_max_employee_count(shift_plan, BEREICHE[2]),
        BEREICHE[3]: find_max_employee_count(shift_plan, BEREICHE[3]),
    }

    for week in range(start_week, end_week + 1):
        # Create list of employees in vacation
        vacation_employees = [employee for employee in employees if week in employee.urlaub_kw]

        # Erstelle ein neues Arbeitsblatt für jede Woche
        worksheet = workbook.add_worksheet(f"KW{week}")

        # Berechne den Starttag der Woche
        start_date_of_the_week = datetime.date(year, 1, 1) + datetime.timedelta(
            days=((week - 1) * 7) - datetime.date(year, 1, 1).weekday()
        )

        # Setze die Breite der Spalten auf 15
        worksheet.set_column(2, 100, 6)
        worksheet.set_column(0, 1, 12)
        worksheet.set_column(9, 10, 12)
        worksheet.set_column(18, 19, 12)

        start_row = 1
        start_col = 1

        for run, (shift) in enumerate(SCHICHT_RHYTHMUS.values()):
            row = start_row
            col = run * 9 + start_col

            # Excel display shift
            worksheet.merge_range(
                row,
                col,
                row + 1,
                col,
                shift,
                bordered_format,
            )
            row += 2

            # Add Bereich to Excel
            for bereich in BEREICHE.values():
                worksheet.merge_range(
                    row,
                    col - 1,
                    row + max_len_areas[bereich] - 1,
                    col - 1,
                    bereich,
                    bordered_format,
                )

                # Add employees to Excel
                for nr in range(max_len_areas[bereich]):
                    # check for error if the list has less employees and then find the employee name
                    if nr < len(shift_plan[week][shift][bereich]):
                        employee = shift_plan[week][shift][bereich][nr]
                        employee_name = employee.name

                        # Check if employee has one day off
                        for day in range(7):
                            start_date_of_the_week_temp = start_date_of_the_week + datetime.timedelta(days=day)

                            if start_date_of_the_week_temp.strftime("%d.%m") in employee.urlaub_tage or day >= 5:
                                worksheet.write(
                                    row + nr,
                                    col + 1 + day,
                                    "X",
                                    bordered_format,
                                )
                            else:
                                worksheet.write(
                                    row + nr,
                                    col + 1 + day,
                                    "",
                                    bordered_format,
                                )

                    else:
                        employee_name = ""
                        # add blanko format for the lane with no emp found
                        for day in range(7):
                            worksheet.write(
                                row + nr,
                                col + 1 + day,
                                "",
                                bordered_format,
                            )

                    # Add employee to excel
                    worksheet.write(
                        row + nr,
                        col,
                        employee_name,
                        bordered_format,
                    )

                row += max_len_areas[bereich]

            col += 1

            # Füge Wochentage und Datum hinzu (Overlay)
            add_days_of_the_week_to_worksheet(
                worksheet=worksheet,
                bordered_format=bordered_format,
                col=col,
                row=start_row,
                start_date_of_the_week=start_date_of_the_week,
            )

            # Add Abwesenheitslist zu worksheet Seite
            col = 29
            row = start_row
            worksheet.set_column(col, col, 15)
            worksheet.write(row, col, "Abwesenheit", bordered_format)
            if vacation_employees:
                for emp in vacation_employees:
                    row += 1
                    worksheet.write(row, col, emp.name, bordered_format)

    # --------------------------------------------------------------------------------
    # Create Data view on the last page
    # --------------------------------------------------------------------------------
    worksheet = workbook.add_worksheet("Information")
    start_col = 1
    row = 1
    col = start_col

    header_info = [
        "Name",
        SCHICHT_RHYTHMUS[1],
        SCHICHT_RHYTHMUS[2],
        SCHICHT_RHYTHMUS[3],
        "Gesamt",
    ]

    # Add header to excel
    for string in header_info:
        worksheet.write(row, col, string, bordered_format)
        col += 1

    row += 1

    # Add information from emp to excel
    for employee in employees:
        col = start_col
        data = employee.count_all()
        worksheet.write(row, col, employee.name, bordered_format)
        col += 1
        for i in data.values():
            worksheet.write(row, col, i, bordered_format)
            col += 1
        row += 1

    # Schließe die Excel-Datei
    workbook.close()
    print(f"Der Schichtplaner für das Jahr {year} wurde erstellt.")


if __name__ == "__main__":
    create_calender(2024)
