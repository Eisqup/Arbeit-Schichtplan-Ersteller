import xlsxwriter
from constants import SETTINGS_VARIABLES, SCHICHT_RHYTHMUS
from shiftPlanner_class import ShiftPlanner
from get_holidays import get_holiday_dates
import datetime
import locale


# Set the locale for German (Germany)
locale.setlocale(locale.LC_TIME, "de_DE.utf8")


def find_max_employee_count(data, area_in, areas_setting, default_extra_rows):
    max_count = 0

    for week_data in data.values():
        for shift_data in week_data.values():
            for area, area_data in shift_data.items():
                if area == area_in:
                    max_count = max(max_count, len(area_data))
    if max_count == 0:
        max_count = 1
    if max_count < areas_setting[area_in]["max"]:
        max_count = areas_setting[area_in]["max"]
    if "fill" in areas_setting[area_in]:
        max_count += default_extra_rows
    return max_count


class CalendarCreator:
    def __init__(self, year, start_week, end_week, program_runs, areas_setting, variable_settings):
        self.year = year
        self.start_week = start_week
        self.end_week = end_week
        self.program_runs = program_runs
        self.areas_setting = areas_setting
        self.variable_settings = variable_settings

        # Variablen
        self.cell_formatting = {}
        self.week_sheets = {}
        self.file_name = f"Schichtplan_{self.year}.xlsx"
        self.sheet_pw = "123"
        self.default_extra_rows_for_areas = 2
        self.day_col_width = 5
        self.area_and_name_col_width = 21
        self.tabel_row_hight = 20
        self.week_length = self.variable_settings["WEEK_RANGE"]
        self.start_row = 2
        self.start_col = 1
        self.author = "Steven Stegink"

        # Colors
        self.school_holiday_color = "#FFFF99"
        self.date_color = "#D3D3D3"
        self.ma_problem_color = "red"
        self.ma_ok_color = "green"
        na_color = "#FF6666"
        gz_color = "#3399FF"
        tu_color = "#B2FF66"
        x_color = "#808080"
        tz_color = "#99AABB"
        l_color = "#B09DD6"
        self.legend_color = {"TU": tu_color, "GZ": gz_color, "NA": na_color, "X": x_color, "TZ": tz_color, "L": l_color}
        self.legend_description = {
            "TU": "Tarifurlaub",
            "GZ": "Gleitzeit",
            "NA": "Nicht Anwesend",
            "X": "Kein Arbeitstag",
            "TZ": "T-Zug",
            "L": "Lehrgang",
        }

        # font size
        self.font_text = {"font_size": 15}  # like names and areas

        # Key init
        print("Creating shift plan....")
        self.create_shifts = ShiftPlanner(self.start_week, self.end_week, self.program_runs, self.areas_setting, self.variable_settings)
        print("Shift plan done.")

        self.school_holiday, self.public_holidays = get_holiday_dates(self.year)
        self.shift_plan = self.create_shifts.get_shift_plan()
        self.max_length_areas = {
            area: find_max_employee_count(self.shift_plan, area, self.areas_setting, self.default_extra_rows_for_areas)
            for area in self.areas_setting.keys()
        }
        # Add a default value for an empty area name
        self.emp_list_length = len(self.create_shifts.employees) + self.default_extra_rows_for_areas * 2
        self.tabel_row_length = sum(self.max_length_areas.values()) + 1  # +1 for count emp
        self.tabel_col_length = self.week_length * 3 + 4  # +4 for 3 times names list and 1 time area list
        self.excel_row_length_max = int(self.tabel_row_length * 3 / 2)

        # Calculate the maximum name length
        max_name_length_area_name = max(len(area) * 1.5 for area in self.max_length_areas.keys())
        max_name_length_emp_name = max(len(emp.name) * 1.5 for emp in self.create_shifts.employees)

        # Calculate the column width based on the maximum name length
        self.text_col_width = max(self.area_and_name_col_width, max_name_length_area_name, max_name_length_emp_name)

        self.run()

    def create_format(self, *format_specs):
        new_format = {"align": "center", "valign": "vcenter"}  # default settings
        for format_spec in format_specs:
            new_format.update(format_spec)
        return self.workbook.add_format(new_format)

    def create_week_sheet_overlay(self):
        def add_header(sheet):
            sheet.set_row(1, 40)
            sheet.write(1, 1, self.year, self.create_format({"font_size": 18, "bold": True}))
            sheet.write(1, 2, sheet.name, self.create_format({"font_size": 18, "bold": True}))

        def add_emp_count(sheet, start_row, end_row, start_col):
            border_format_big = {"left": 2, "right": 2, "top": 2, "bottom": 2, "hidden": 1, "locked": 1}
            # create emp count row
            border_format = {"top": 2, "left": 1, "right": 1, "bottom": 2, "font_size": 13, "hidden": 1, "locked": 1}
            col = start_col + 1
            row = end_row
            for runs in SCHICHT_RHYTHMUS.values():
                sheet.write(row, col, "Anzahl MA", self.create_format(border_format_big))
                excel_range_1 = xlsxwriter.utility.xl_range(start_row + 2, col, row - 1, col)
                col += 1
                for day in range(self.week_length):
                    excel_range_2 = xlsxwriter.utility.xl_range(start_row + 2, col + day, row - 1, col + day)
                    count_formula = f'SUMPRODUCT(({excel_range_1}<>"")*({excel_range_2}=""))'
                    if day == 0:
                        sheet.write_array_formula(row, col + day, row, col + day, count_formula, self.create_format(border_format, {"left": 2}))
                    elif day == self.week_length - 1:
                        sheet.write_array_formula(row, col + day, row, col + day, count_formula, self.create_format(border_format, {"right": 2}))
                    else:
                        sheet.write_array_formula(row, col + day, row, col + day, count_formula, self.create_format(border_format))
                col += self.week_length

        def add_areas_tabel(sheet, start_row, start_col):
            row = start_row
            col = start_col
            # Areas tabel
            sheet.merge_range(row, col, row + 1, col, "Bereich", self.create_format(border_format_big_lock, {"bold": True}, self.font_text))
            sheet.set_column(col, col, self.text_col_width)
            sheet.set_row(row, 20)
            sheet.set_row(row + 1, 20)
            row += 2
            for area, length in self.max_length_areas.items():
                if area == "-b":
                    area = ""
                if length == 1:
                    sheet.write(row, col, area, self.create_format(border_format_big_lock, self.font_text))
                else:
                    sheet.merge_range(row, col, row + length - 1, col, area, self.create_format(border_format_big_lock, self.font_text))
                row += length

        def add_tabel(sheet, start_row, start_col, start_date_of_the_week, week):
            row = start_row
            col = start_col

            # EMP-Name tabel, date tabel, presence tabel, count emp view
            border_format_names = {"left": 2, "right": 2, "font_size": 15, "hidden": 1, "locked": 0}

            # Create EMP_NAME tabel
            col = start_col + 1

            cell_format = self.create_format({"bg_color": "red"})
            for runs in SCHICHT_RHYTHMUS.values():
                row = start_row
                # width
                sheet.set_column(col, col, self.text_col_width)
                # Header
                sheet.merge_range(
                    row, col, row + 1, col, f"{runs}schicht", self.create_format(border_format_big_lock, self.font_text, {"bold": True})
                )
                row += 2
                for area, length in self.max_length_areas.items():
                    format_ = self.create_format(border_format_names, {"top": 2, "bottom": 1})
                    sheet.write_blank(row, col, None, format_)
                    self.cell_formatting[f"{week}-{row}-{col}"] = format_
                    for i in range(row + 1, row + length - 1):
                        sheet.set_row(i, self.tabel_row_hight)
                        format_ = self.create_format(border_format_names, {"top": 1, "bottom": 1})
                        sheet.write_blank(i, col, None, format_)
                        self.cell_formatting[f"{week}-{i}-{col}"] = format_
                    format_ = self.create_format(border_format_names, {"top": 1, "bottom": 2})
                    sheet.write_blank(row + length - 1, col, None, format_)
                    self.cell_formatting[f"{week}-{row + length - 1}-{col}"] = format_
                    sheet.set_row(row + length - 1, self.tabel_row_hight)
                    row += length
                x = xlsxwriter.utility.xl_range(start_row + 2, col, row - 1, col)
                sheet.conditional_format(
                    x,
                    {"type": "formula", "criteria": f'=IF(AND({x[:2]}<>"", COUNTIF({tabel_range}, {x[:2]}) > 1), 1, 0)', "format": cell_format},
                )
                sheet.conditional_format(
                    x,
                    {
                        "type": "formula",
                        "criteria": f'=AND({x[:2]}<>"", COUNTIF({ma_list_range}, {x[:2]}) =0)',
                        "format": cell_format,
                    },
                )
                col += self.week_length + 1

            # Create date tabel
            col = start_col + 2
            for runs in SCHICHT_RHYTHMUS.values():
                row = start_row
                for day in range(self.week_length):
                    sheet.set_column(col + day, col + day, self.day_col_width)
                    date_of_the_week = start_date_of_the_week + datetime.timedelta(days=day)
                    sheet.write(row, col + day, date_of_the_week.strftime("%A")[:2], self.create_format(border_format_big_lock))
                    date_string = date_of_the_week.strftime("%d.%m")

                    # Check if the date is in the school holiday list
                    if date_string in self.public_holidays or date_of_the_week.weekday() in [5, 6]:
                        bg_color = self.legend_color["X"]

                    elif date_string in self.school_holiday:
                        bg_color = self.school_holiday_color
                    else:
                        bg_color = self.date_color

                    sheet.write(row + 1, col + day, date_string, self.create_format(border_format_big_lock, {"bg_color": bg_color}))
                col += self.week_length + 1

            # create presence tabel
            def get_border_format(day, area_index, bot_line):
                left = right = top = bottom = 1

                if day == 0:
                    left = 2
                elif day == self.week_length - 1:
                    right = 2

                if area_index == 0:
                    top = 2
                elif area_index == bot_line - 1:
                    bottom = 2

                return {"left": left, "right": right, "top": top, "bottom": bottom, "font_size": 13, "hidden": 1, "locked": 0}

            col = start_col + 2
            for runs in SCHICHT_RHYTHMUS.values():
                for day in range(self.week_length):
                    row = start_row + 2
                    for key, range_ in self.max_length_areas.items():
                        for row_add in range(range_):
                            border_format = get_border_format(day, row_add, range_)
                            format_ = self.create_format(border_format)
                            sheet.write_blank(row + row_add, col + day, None, format_)
                            self.cell_formatting[f"{week}-{row + row_add}-{col + day}"] = format_
                        row += range_
                for key, item in self.legend_color.items():
                    x = xlsxwriter.utility.xl_range(start_row + 2, col, row - 1, col - 1 + self.week_length)
                    formal = xlsxwriter.utility.xl_rowcol_to_cell(start_row + 2, col)
                    cell_format = self.create_format({"bg_color": item})
                    sheet.conditional_format(
                        x,
                        {"type": "formula", "criteria": f'=EXACT({formal},"{key}")', "format": cell_format},
                    )
                col += self.week_length + 1

        def add_away_list(sheet, start_row, start_col, week):
            row = start_row
            col = start_col + self.tabel_col_length + 1

            sheet.merge_range(
                row,
                col,
                row + 1,
                col,
                "Abwesende MA",
                self.create_format(border_format_big_lock, self.font_text, {"bold": True}),
            )
            sheet.set_column(col, col, self.text_col_width)
            sheet.set_column(col - 1, col - 1, 2)
            row += 2
            for length in range(start_row - 1 + self.excel_row_length_max):
                sheet.set_row(row + length, self.tabel_row_hight)
                format_ = self.create_format(border_format_big_unlock, self.font_text)
                sheet.write_blank(row + length, col, None, format_)
                self.cell_formatting[f"{week}-{row + length}-{col}"] = format_

            cell_format = self.create_format({"bg_color": "red"})
            x = xlsxwriter.utility.xl_range(row, col, row + self.excel_row_length_max, col)
            sheet.conditional_format(
                x,
                {"type": "formula", "criteria": f'=IF(AND({x[:2]}<>"", COUNTIF({tabel_range}, {x[:2]}) > 1), 1, 0)', "format": cell_format},
            )

        def add_emp_list(sheet, start_row, start_col, week):
            row = start_row
            col = start_col + self.tabel_col_length + 1 + 2

            for add in range(2):
                col = col + (add * 2)
                row = start_row
                sheet.set_column(col - 1, col - 1, 2)
                sheet.set_column(col, col, self.text_col_width)
                sheet.merge_range(
                    row,
                    col,
                    row + 1,
                    col,
                    "MA Liste",
                    self.create_format(border_format_big_lock, self.font_text, {"bold": True}),
                )
                for _ in range(row + self.excel_row_length_max):
                    row += 1
                    format_ = self.create_format(border_format_big_unlock, self.font_text)
                    sheet.write_blank(row, col, None, format_)
                    self.cell_formatting[f"{week}-{row}-{col}"] = format_

                cell_format_r = self.create_format({"bg_color": self.ma_problem_color})
                cell_format_g = self.create_format({"bg_color": self.ma_ok_color})
                x = xlsxwriter.utility.xl_range(start_row + 2, col, row, col)
                first_group = x.split(":")[0]
                formula = f'=AND({first_group}<>"", ' f"COUNTIF({tabel_range}, {first_group}) = 1, " f"COUNTIF({ma_list_range}, {first_group}) = 1)"
                sheet.conditional_format(
                    x,
                    {"type": "formula", "criteria": formula, "format": cell_format_g},
                )
                formula = (
                    f'=AND({first_group}<>"", OR('
                    f"COUNTIF({tabel_range}, {first_group}) <> 1, "
                    f"COUNTIF({ma_list_range}, {first_group}) <> 1"
                    f"))"
                )
                sheet.conditional_format(
                    x,
                    {"type": "formula", "criteria": formula, "format": cell_format_r},
                )

        def add_legend(sheet, start_row, start_col):
            row = start_row + 3 + self.tabel_row_length
            col = start_col + 1
            sheet.merge_range(row, col, row, col + 1, "Legende", self.create_format(border_format_big_lock, {"bold": True}))
            for key, item in self.legend_color.items():
                row += 1
                sheet.set_row(row, self.tabel_row_hight)
                sheet.write(row, col, f"{key} => {self.legend_description[key]}", self.create_format(border_format_big_lock))
                sheet.write(row, col + 1, key, self.create_format(border_format_big_lock, {"bg_color": item}))
            row += 1
            sheet.write(row, col, f"Schulfrei", self.create_format(border_format_big_lock))
            sheet.write(row, col + 1, "", self.create_format(border_format_big_lock, {"bg_color": self.school_holiday_color}))
            sheet.set_row(row, self.tabel_row_hight)
            row += 1
            sheet.write(row, col, f"Probleme bei MA", self.create_format(border_format_big_lock))
            sheet.write(row, col + 1, "", self.create_format(border_format_big_lock, {"bg_color": self.ma_problem_color}))
            sheet.set_row(row, self.tabel_row_hight)
            row += 1
            sheet.write(row, col, f"MA in Tabelle", self.create_format(border_format_big_lock))
            sheet.write(row, col + 1, "", self.create_format(border_format_big_lock, {"bg_color": self.ma_ok_color}))
            sheet.set_row(row, self.tabel_row_hight)

        # -------------------------------------------------------------------------------------------------------------------------------
        # --------------------------Programm Start---------------------------------------------------------------------------------------
        # -------------------------------------------------------------------------------------------------------------------------------
        row = self.start_row
        col = self.start_col
        # Define conditional format rules for multiple cell ranges
        tabel_range = xlsxwriter.utility.xl_range_abs(
            row + 2,
            col + 1,
            row + self.excel_row_length_max + 2,
            col + self.tabel_col_length + 1,
        )
        ma_list_range = xlsxwriter.utility.xl_range_abs(
            row + 2,
            col + self.tabel_col_length + 3,
            row + 2 + self.excel_row_length_max,
            col + self.tabel_col_length + 5,
        )
        # Formats
        border_format_big_lock = {"left": 2, "right": 2, "top": 2, "bottom": 2, "hidden": 1, "locked": 1}
        border_format_big_unlock = {"left": 2, "right": 2, "top": 2, "bottom": 2, "hidden": 1, "locked": 0}

        for week, week_sheet in self.week_sheets.items():
            # Berechne den Starttag der Woche
            start_date_of_the_week = datetime.date(self.year, 1, 1) + datetime.timedelta(
                days=((week - 1) * 7) - datetime.date(self.year, 1, 1).weekday()
            )

            self.sheet_protection = {
                "objects": False,
                "scenarios": False,
                "format_cells": False,
                "format_columns": False,
                "format_rows": False,
                "insert_columns": False,
                "insert_rows": False,
                "insert_hyperlinks": False,
                "delete_columns": False,
                "delete_rows": False,
                "select_locked_cells": False,
                "sort": False,
                "autofilter": False,
                "pivot_tables": False,
                "select_unlocked_cells": True,
            }

            week_sheet.protect(
                self.sheet_pw,
                self.sheet_protection,
            )
            week_sheet.set_row(0, 6)
            week_sheet.set_column(0, 0, 1)
            add_header(week_sheet)
            add_areas_tabel(week_sheet, start_row=row, start_col=col)
            add_tabel(week_sheet, start_row=row, start_col=col, start_date_of_the_week=start_date_of_the_week, week=week)
            add_emp_count(week_sheet, start_row=row, end_row=row + self.tabel_row_length + 1, start_col=col)
            add_away_list(week_sheet, start_row=row, start_col=col, week=week)
            add_emp_list(week_sheet, start_row=row, start_col=col, week=week)
            add_legend(week_sheet, start_row=row, start_col=col)

    def create_sheets_for_weeks(self):
        for week in range(1, 52 + 1):
            week_sheet = self.workbook.add_worksheet(f"KW{week}")
            self.week_sheets[week] = week_sheet

    def insert_emp_to_weeks(self):
        for week, week_shift in self.shift_plan.items():
            col = self.start_col + 1
            week_sheet = self.week_sheets[week]
            for shift, areas in week_shift.items():
                row = self.start_row + 2
                for area, row_length in self.max_length_areas.items():
                    for index in range(row_length):
                        if area in areas.keys():
                            if len(areas[area]) > index:
                                week_sheet.write(row + index, col, areas[area][index].name, self.cell_formatting[f"{week}-{row + index}-{col}"])
                    row += row_length
                col += self.week_length + 1

    def insert_emp_to_ma_list(self):
        def sort_by_name(emp):
            return emp.name

        self.employees = sorted(self.create_shifts.employees, key=sort_by_name)
        for week in self.shift_plan.keys():
            row = self.start_row + 2
            col = self.start_col + self.tabel_col_length + 3
            week_sheet = self.week_sheets[week]
            for emp in self.employees:
                if row == self.excel_row_length_max + self.start_row + 3:
                    row = self.start_row + 2
                    col += 2
                week_sheet.write(row, col, emp.name, self.cell_formatting[f"{week}-{row}-{col}"])
                row += 1

    def insert_emp_to_away_list(self):
        employees = self.create_shifts.employees
        for week in self.shift_plan.keys():
            row = self.start_row + 2
            col = self.start_col + self.tabel_col_length + 1
            week_sheet = self.week_sheets[week]
            for emp in employees:
                if emp.emp_shifts_dict[week] == "X":
                    week_sheet.write(row, col, emp.name, self.cell_formatting[f"{week}-{row}-{col}"])
                    row += 1

    def insert_public_holidays_and_weekend(self):
        for week, sheet in self.week_sheets.items():
            start_date_of_the_week = datetime.date(self.year, 1, 1) + datetime.timedelta(
                days=((week - 1) * 7) - datetime.date(self.year, 1, 1).weekday()
            )
            col = self.start_col + 2
            for runs in SCHICHT_RHYTHMUS.values():
                for day in range(self.week_length):
                    row = self.start_row + 2
                    date_of_the_week = start_date_of_the_week + datetime.timedelta(days=day)
                    if date_of_the_week.strftime("%d.%m") in self.public_holidays or date_of_the_week.weekday() in [5, 6]:
                        for row in range(row, row + self.tabel_row_length - 1):
                            sheet.write(row, col, "X", self.cell_formatting[f"{week}-{row}-{col}"])
                    col += 1
                col += 1

    def create_area_statistic_sheet(self):
        sheet = self.workbook.add_worksheet("Bereichs Informationen")
        col = self.start_col + 1
        sheet.protect(self.sheet_pw, self.sheet_protection)
        sheet.freeze_panes(3, 0)
        sheet.set_row(0, 6)
        sheet.set_column(0, 0, 1)
        border_format_big_lock = {"left": 2, "right": 2, "top": 2, "bottom": 2, "hidden": 1, "locked": 1, "font_size": 16}
        row_size = 20

        # Create Range Dict to create formals later
        area_range_dict = {}
        for shift in SCHICHT_RHYTHMUS.values():
            row = 4
            for area, length in self.max_length_areas.items():
                range_ = xlsxwriter.utility.xl_range(row, col, row + length - 1, col)
                area_range_dict[f"{shift}-{area}"] = range_
                row += length

            col += self.week_length + 1

        row = 1
        col = 2
        sheet.merge_range(
            row,
            col,
            row,
            col + 2 + len(self.max_length_areas.keys()),
            "Bereichs Informationen",
            self.create_format(border_format_big_lock, {"font_size": 15, "bold": True}),
        )
        ### HEADER ####
        row += 1
        sheet.write(row, col, "Woche", self.create_format(border_format_big_lock, {"bold": True}))
        sheet.set_column(col, col, 10)
        sheet.set_row(1, row_size)
        sheet.set_row(2, row_size)
        sheet.set_row(3, row_size)

        col += 1
        sheet.write(row, col, f"Schicht", self.create_format(border_format_big_lock, {"bold": True}))
        sheet.set_column(col, col, self.text_col_width)
        col += 1
        sheet.set_column(col, col, row_size)
        for area in self.max_length_areas.keys():
            sheet.write(row, col, area, self.create_format(border_format_big_lock, {"bold": True}))
            sheet.set_column(col, col, self.text_col_width)
            col += 1
        sheet.write(row, col, "Gesamt", self.create_format(border_format_big_lock, {"bold": True}))
        sheet.set_column(col, col, self.text_col_width)
        col += 1

        row += 1

        def create_border(row, col):
            start_col = 2
            start_row = 3

            if row == start_row or (row - start_row) % (len(SCHICHT_RHYTHMUS)) == 0:
                top = 2
            else:
                top = 1

            if (row - start_row) % (len(SCHICHT_RHYTHMUS)) == len(SCHICHT_RHYTHMUS) - 1:
                bottom = 2
            else:
                bottom = 1

            if col == start_col:
                left = 2
            else:
                left = 1

            if col == start_col + 2 + len(self.max_length_areas.keys()):
                right = 2
            else:
                right = 1

            return {"left": left, "right": right, "top": top, "bottom": bottom}

        # Define cell colors as format objects
        red_format = self.create_format({"bg_color": "red"})
        green_format = self.create_format({"bg_color": "green"})
        yellow_format = self.create_format({"bg_color": "yellow"})
        dark_green_format = self.create_format({"bg_color": "#006400"})  # Dark Green in hex

        for week in range(1, 53):
            for shift in SCHICHT_RHYTHMUS.values():
                col = 2
                sheet.set_row(row, row_size)
                sheet.write(row, col, week, self.create_format(border_format_big_lock, create_border(row, col)))
                col += 1
                sheet.write(row, col, f"{shift}schicht", self.create_format(border_format_big_lock, create_border(row, col)))
                col += 1
                range_gesamt = xlsxwriter.utility.xl_range(row, col, row, col - 1 + len(self.max_length_areas.keys()))
                cor_gesamt = xlsxwriter.utility.xl_range(row, col + len(self.max_length_areas.keys()), row, col + len(self.max_length_areas.keys()))
                for area in self.max_length_areas.keys():
                    min_value = self.areas_setting[area]["min"]
                    max_value = self.areas_setting[area]["max"]
                    cor_area = xlsxwriter.utility.xl_range(row, col, row, col)

                    sheet.write(
                        row,
                        col,
                        f'=COUNTA(KW{week}!{area_range_dict[f"{shift}-{area}"]})',
                        self.create_format(border_format_big_lock, create_border(row, col)),
                    )
                    if min_value > 0:
                        sheet.conditional_format(
                            cor_area,
                            {
                                "type": "formula",
                                "criteria": f"AND({cor_area}>={min_value}, {cor_area}<{max_value}, {cor_gesamt}>0)",
                                "format": yellow_format,
                            },
                        )
                        sheet.conditional_format(
                            cor_area,
                            {
                                "type": "formula",
                                "criteria": f"AND({cor_area}>{max_value}, {cor_gesamt}>0)",
                                "format": dark_green_format,
                            },
                        )
                        sheet.conditional_format(
                            cor_area,
                            {
                                "type": "formula",
                                "criteria": f"AND({cor_area}={max_value}, {cor_gesamt}>0)",
                                "format": green_format,
                            },
                        )
                        sheet.conditional_format(
                            cor_area,
                            {
                                "type": "formula",
                                "criteria": f"AND({cor_area}<{min_value}, {cor_gesamt}>0)",
                                "format": red_format,
                            },
                        )

                    col += 1
                sheet.write(row, col, f"=SUM({range_gesamt})", self.create_format(border_format_big_lock, create_border(row, col)))
                col += 1
                row += 1

    def create_employee_statistic_sheet(self):
        sheet = self.workbook.add_worksheet("MA Informationen")
        col = self.start_col + 1
        sheet.protect(
            self.sheet_pw,
            {
                "objects": False,
                "scenarios": False,
                "format_cells": False,
                "format_columns": False,
                "format_rows": False,
                "insert_columns": False,
                "insert_rows": False,
                "insert_hyperlinks": False,
                "delete_columns": False,
                "delete_rows": False,
                "select_locked_cells": False,
                "sort": True,  # Allow sorting within the table.
                "autofilter": True,  # Allow applying autofilter within the table.
                "pivot_tables": False,
                "select_unlocked_cells": True,
            },
        )
        sheet.freeze_panes(3, 0)
        sheet.set_row(0, 6)
        sheet.set_column(0, 0, 1)
        border_format_big_lock = self.create_format(
            {
                "left": 2,
                "right": 2,
                "top": 2,
                "bottom": 2,
                "font_size": 14,
                "hidden": 1,
                "locked": 1,
            }
        )
        row_size = 20

        # Create Range Dict to create formals later
        sum_area_range_dict = {}
        for shift in SCHICHT_RHYTHMUS.values():
            row = self.start_row + 2
            end_row = sum(self.max_length_areas.values()) + self.start_row + 2
            range_ = xlsxwriter.utility.xl_range(row, col, end_row - 1, col)
            sum_area_range_dict[f"{shift}"] = range_
            col += self.week_length + 1

        ### HEADER ###
        column_names = []
        sheet.merge_range(1, 2, 1, 10, "MA Informationen", border_format_big_lock)
        sheet.write(2, 2, "Name", border_format_big_lock)
        column_names.append("Name")
        sheet.set_column(2, 2, self.text_col_width)
        sheet.set_row(2, row_size)
        col = 3
        for shift in SCHICHT_RHYTHMUS.values():
            sheet.set_column(col, col, 15)
            sheet.write(2, col, f"{shift}", border_format_big_lock)
            column_names.append(shift)
            col += 1
        sheet.write(2, col, "Gesamt", border_format_big_lock)
        column_names.append("Gesamt")
        sheet.set_column(col, col, 15)
        col += 1
        sheet.write(2, col, "Abwesend", border_format_big_lock)
        column_names.append("Abwesend")
        sheet.set_column(col, col, 15)
        col += 1
        for shift in SCHICHT_RHYTHMUS.values():
            sheet.set_column(col, col, 15)
            sheet.write(2, col, f"{shift} in %", border_format_big_lock)
            column_names.append(f"{shift} in %")
            col += 1

        #### Overlay #####
        row = 3
        for row in range(row, row + self.excel_row_length_max * 2):
            col = 2
            sheet.set_row(row, row_size)
            if row < len(self.employees) + 3:
                sheet.write(
                    row,
                    col,
                    self.employees[row - 3].name,
                    self.create_format(
                        {
                            "left": 2,
                            "right": 2,
                            "top": 2,
                            "bottom": 2,
                            "font_size": 14,
                            "hidden": 1,
                            "locked": 0,
                        }
                    ),
                )
            else:
                sheet.write(
                    row,
                    col,
                    "",
                    self.create_format(
                        {
                            "left": 2,
                            "right": 2,
                            "top": 2,
                            "bottom": 2,
                            "hidden": 1,
                            "locked": 0,
                        }
                    ),
                )
            range_name = xlsxwriter.utility.xl_range(row, col, row, col)
            col += 1
            for shift in SCHICHT_RHYTHMUS.values():
                sheet.write_array_formula(
                    row,
                    col,
                    row,
                    col,
                    f'SUMPRODUCT(SUM(COUNTIF(INDIRECT("\'" & "KW" & ROW(INDIRECT("1:52")) & "\'!{sum_area_range_dict[shift]}"), {range_name})))',
                    border_format_big_lock,
                )
                col += 1
            range_g = xlsxwriter.utility.xl_range(row, col - len(SCHICHT_RHYTHMUS), row, col - 1)
            range_gesamt = xlsxwriter.utility.xl_range(row, col, row, col)
            sheet.write(row, col, f"=SUM({range_g})", border_format_big_lock)
            col += 1
            range_away_list = xlsxwriter.utility.xl_range(
                self.start_row + 2,
                self.start_col + self.tabel_col_length + 1,
                self.start_row + 2 + self.excel_row_length_max,
                self.start_col + self.tabel_col_length + 1,
            )
            sheet.write_array_formula(
                row,
                col,
                row,
                col,
                f'SUMPRODUCT(SUM(COUNTIF(INDIRECT("\'" & "KW" & ROW(INDIRECT("1:52")) & "\'!{range_away_list}"), {range_name})))',
                border_format_big_lock,
            )
            col += 1
            for i, shift in enumerate(SCHICHT_RHYTHMUS.values()):
                col_n = ""
                row_n = ""
                for char in range_name:
                    if char.isalpha():
                        col_n += char
                    else:
                        row_n += char
                col_n = chr(ord(col_n) + 1 + i)
                new_cell_ref = col_n + str(row_n)
                sheet.write(
                    row,
                    col,
                    f'=IF({range_gesamt}=0, "", ROUND({new_cell_ref} * 100 / {range_gesamt}, 0))',
                    border_format_big_lock,
                )
                col += 1
        table_range = xlsxwriter.utility.xl_range(2, 2, 2 + self.excel_row_length_max * 2, 2 + len(SCHICHT_RHYTHMUS) * 2 + 2)
        sheet.add_table(table_range, {"name": "MAInfo", "columns": [{"header": name} for name in column_names]})
        sheet.unprotect_range(table_range)

    def run(self):
        print("Creating Excel file...")
        self.workbook = xlsxwriter.Workbook(self.file_name, {"author": self.author})
        self.create_sheets_for_weeks()
        self.create_week_sheet_overlay()
        self.insert_emp_to_weeks()
        self.insert_emp_to_away_list()
        self.insert_emp_to_ma_list()
        self.insert_public_holidays_and_weekend()
        self.create_area_statistic_sheet()
        self.create_employee_statistic_sheet()
        self.workbook.close()
        print("Done")


if __name__ == "__main__":
    CalendarCreator(
        year=2023,
        start_week=2,
        end_week=50,
        program_runs=10,
        areas_setting={
            "FFK": {"prio": 5, "min": 0, "max": 1},
            "Drehen": {"prio": 1, "min": 3, "max": 3, "fill": True},
            "Bohren": {"prio": 3, "min": 2, "max": 3},
            "Fräsen": {"prio": 2, "min": 3, "max": 3},
        },
        variable_settings=SETTINGS_VARIABLES,
    )
