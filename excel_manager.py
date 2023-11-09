import xlsxwriter
import xlwings as xw
import os
import datetime
import locale
import time

from constants import SETTINGS_VARIABLES, SCHICHT_RHYTHMUS
from shift_planner_manager import ShiftPlanner
from get_holidays import get_holiday_dates
from vba import vba_code


# Set the locale for German (Germany)
locale.setlocale(locale.LC_TIME, "de_DE.utf8")


def find_max_employee_count(data, area_in, areas_setting, default_extra_rows):
    max_count = 0

    if data:
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


class ExcelCreator:
    def __init__(self, year, start_week, end_week, program_runs, areas_setting, variable_settings, excel_pw):
        self.year = year
        self.start_week = start_week
        self.end_week = end_week
        self.program_runs = program_runs
        self.areas_setting = areas_setting
        self.variable_settings = variable_settings
        self.excel_sheets_pw = str(excel_pw)

        # Variablen
        self.cell_formatting = {}
        self.week_sheets = {}
        self.all_sheets = []
        self.timestamp = time.strftime("%Y%m%d%H%M%S")
        self.file_name_xlsx = f"Schichtplan_{self.year}_{self.timestamp}.xlsx"
        self.file_name_xlsm = f"Schichtplan_{self.year}_{self.timestamp}.xlsm"
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
        self.count_ma_color = "#EAEAEA"
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
        self.max_length_areas = {area: find_max_employee_count(self.shift_plan, area, self.areas_setting, self.default_extra_rows_for_areas) for area in self.areas_setting.keys()}
        # Add a default value for an empty area name
        self.emp_list_length = len(self.create_shifts.employees) + self.default_extra_rows_for_areas * 2
        self.tabel_row_length = sum(self.max_length_areas.values()) + 1  # +1 for count emp
        self.tabel_col_length = self.week_length * 3 + 4  # +4 for 3 times names list and 1 time area list
        self.excel_row_length_max = int(self.tabel_row_length * 3 / 2)

        # Calculate the maximum name length
        max_name_length_area_name = max(len(area) * 1.5 for area in self.max_length_areas.keys())
        if self.create_shifts.employees:
            max_name_length_emp_name = max(len(emp.name) * 1.5 for emp in self.create_shifts.employees)
        else:
            max_name_length_emp_name = 0

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
            border_format_big = {
                "left": 2,
                "right": 2,
                "top": 2,
                "bottom": 2,
                "hidden": 1,
                "locked": 1,
                "bold": True,
                "bg_color": self.count_ma_color,
            }
            # create emp count row
            border_format = {"top": 2, "left": 1, "right": 1, "bottom": 2, "font_size": 13, "hidden": 1, "locked": 1, "bg_color": self.count_ma_color}
            col = start_col + 1
            row = end_row
            for i, runs in enumerate(SCHICHT_RHYTHMUS.values()):
                if i == 0:
                    sheet.write(row, col, "Anzahl MA", self.create_format(border_format_big, self.font_text))
                else:
                    sheet.write(row, col, "", self.create_format(border_format_big, self.font_text))
                excel_range_1 = xlsxwriter.utility.xl_range(start_row + 2, col, row - 1, col)
                col += 1
                for day in range(self.week_length):
                    excel_range_2 = xlsxwriter.utility.xl_range(start_row + 2, col + day, row - 1, col + day)
                    count_formula = f'SUMPRODUCT(({excel_range_1}<>"")*({excel_range_2}=""))'
                    if day == 0:
                        sheet.write_array_formula(row, col + day, row, col + day, count_formula, self.create_format(border_format, {"left": 2}, self.font_text))
                    elif day == self.week_length - 1:
                        sheet.write_array_formula(row, col + day, row, col + day, count_formula, self.create_format(border_format, {"right": 2}, self.font_text))
                    else:
                        sheet.write_array_formula(row, col + day, row, col + day, count_formula, self.create_format(border_format, self.font_text))
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
                if area.startswith("-b"):
                    area = ""
                if length == 1:
                    sheet.write(row, col, area, self.create_format(border_format_big_lock, self.font_text, {"bold": True}))
                else:
                    sheet.merge_range(row, col, row + length - 1, col, area, self.create_format(border_format_big_lock, self.font_text, {"bold": True}))
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
                sheet.merge_range(row, col, row + 1, col, f"{runs}schicht", self.create_format(border_format_big_lock, self.font_text, {"bold": True}))
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
                    {
                        "type": "formula",
                        "criteria": f'=IF(AND({x[:2]}<>"",{x[:2]}<>"Anzahl MA", COUNTIF({tabel_range}, {x[:2]}) > 1), 1, 0)',
                        "format": cell_format,
                    },
                )
                sheet.conditional_format(
                    x,
                    {
                        "type": "formula",
                        "criteria": f'=AND({x[:2]}<>"",{x[:2]}<>"Anzahl MA", COUNTIF({ma_list_range}, {x[:2]}) =0)',
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
                formula = f'=AND({first_group}<>"", OR(' f"COUNTIF({tabel_range}, {first_group}) <> 1, " f"COUNTIF({ma_list_range}, {first_group}) <> 1" f"))"
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
            start_date_of_the_week = datetime.date(self.year, 1, 1) + datetime.timedelta(days=((week - 1) * 7) - datetime.date(self.year, 1, 1).weekday())

            week_sheet.set_row(0, 6)
            week_sheet.set_column(0, 0, 1)
            add_header(week_sheet)
            add_areas_tabel(week_sheet, start_row=row, start_col=col)
            add_tabel(week_sheet, start_row=row, start_col=col, start_date_of_the_week=start_date_of_the_week, week=week)

            ##### Creating breaks #####
            if any(key.startswith("-b") for key in self.areas_setting.keys()):
                count_up = 0
                start_r = row
                for key, length in self.max_length_areas.items():
                    count_up += length
                    if key.startswith("-b"):
                        add_emp_count(week_sheet, start_row=start_r, end_row=start_r + count_up + 1, start_col=col)
                        start_r += count_up
                        count_up = 0
                add_emp_count(week_sheet, start_row=start_r, end_row=start_r + count_up + 1 + 1, start_col=col)
            else:
                add_emp_count(week_sheet, start_row=row, end_row=row + self.tabel_row_length + 1, start_col=col)

            add_away_list(week_sheet, start_row=row, start_col=col, week=week)
            add_emp_list(week_sheet, start_row=row, start_col=col, week=week)
            add_legend(week_sheet, start_row=row, start_col=col)

    def create_sheets_for_weeks(self):
        for week in range(1, 52 + 1):
            week_sheet = self.workbook.add_worksheet(f"KW{week}")
            self.week_sheets[week] = week_sheet
            self.all_sheets.append(week_sheet)

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
        ## Check for br
        result_row_postion_list = []
        row_postion = self.start_row + 2
        for key, value in self.max_length_areas.items():
            row_postion += value
            if key.startswith("-b"):
                result_row_postion_list.append(row_postion - 1)

        for week, sheet in self.week_sheets.items():
            start_date_of_the_week = datetime.date(self.year, 1, 1) + datetime.timedelta(days=((week - 1) * 7) - datetime.date(self.year, 1, 1).weekday())
            col = self.start_col + 2
            for runs in SCHICHT_RHYTHMUS.values():
                for day in range(self.week_length):
                    row = self.start_row + 2

                    date_of_the_week = start_date_of_the_week + datetime.timedelta(days=day)
                    if date_of_the_week.strftime("%d.%m") in self.public_holidays or date_of_the_week.weekday() in [5, 6]:
                        for row in range(row, row + self.tabel_row_length - 1):
                            if row not in result_row_postion_list:
                                sheet.write(row, col, "X", self.cell_formatting[f"{week}-{row}-{col}"])
                    col += 1
                col += 1

    def create_area_statistic_sheet(self):
        sheet = self.workbook.add_worksheet("Bereichs Informationen")
        col = self.start_col + 1
        self.all_sheets.append(sheet)
        sheet.freeze_panes(3, 3)
        sheet.set_row(0, 6)
        sheet.set_column(0, 0, 1)
        sheet.set_column("A:A", None, None, {"hidden": 1})
        border_format_big_lock = {"left": 2, "right": 2, "top": 2, "bottom": 2, "hidden": 1, "locked": 1, "font_size": 16}
        row_size = 20

        # Create Range Dict to create formals later
        area_range_dict = {}
        keys_to_remove = []
        for shift in SCHICHT_RHYTHMUS.values():
            row = 4
            for area, length in self.max_length_areas.items():
                range_ = xlsxwriter.utility.xl_range(row, col, row + length - 1, col)
                area_range_dict[f"{shift}-{area}"] = range_
                row += length

            col += self.week_length + 1

        dict_for_sheet = {}
        for key, item in self.areas_setting.items():
            if item["min"] != 0 or key.startswith("-b"):
                dict_for_sheet[key] = self.max_length_areas[key]

        for key in reversed(list(dict_for_sheet.keys())):
            if key.startswith("-b"):
                del dict_for_sheet[key]
            else:
                break

        row = 1
        col = 1
        sheet.merge_range(
            row,
            col,
            row,
            col + 2 + len(dict_for_sheet.keys()),
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
        for area in dict_for_sheet.keys():
            if area.startswith("-b"):
                sheet.write(row, col, "Gesamt", self.create_format(border_format_big_lock, {"bold": True}))
                sheet.set_column(col, col, 10)
            else:
                sheet.write(row, col, area, self.create_format(border_format_big_lock, {"bold": True}))
                sheet.set_column(col, col, self.text_col_width)
            col += 1
        sheet.write(row, col, "Gesamt", self.create_format(border_format_big_lock, {"bold": True}))
        sheet.set_column(col, col, 10)
        col += 1

        row += 1

        def create_border(row, col):
            start_col = 1
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

            if col == start_col + 2 + len(dict_for_sheet.keys()):
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
                start_col = 3
                end_col = start_col - 1
                col = 1
                sheet.set_row(row, row_size)
                sheet.write(row, col, week, self.create_format(border_format_big_lock, create_border(row, col)))
                col += 1
                sheet.write(row, col, f"{shift}schicht", self.create_format(border_format_big_lock, create_border(row, col)))
                col += 1

                # Find count coordinate
                postion_dict = {}
                found_areas = []
                for i, area in enumerate(dict_for_sheet.keys()):
                    if area.startswith("-b"):
                        found_areas.append(area)
                        for area2 in found_areas:
                            cor_gesamt = xlsxwriter.utility.xl_range(row, col + i, row, col + i)
                            postion_dict[area2] = cor_gesamt
                        found_areas = []
                    else:
                        found_areas.append(area)

                    if len(postion_dict) + len(found_areas) == len(dict_for_sheet.keys()):
                        for area2 in found_areas:
                            cor_gesamt = xlsxwriter.utility.xl_range(row, col + i + 1, row, col + i + 1)
                            postion_dict[area2] = cor_gesamt

                for area in dict_for_sheet.keys():
                    cor_gesamt = postion_dict[area]
                    end_col += 1
                    range_gesamt = xlsxwriter.utility.xl_range(row, start_col, row, end_col - 1)
                    if area.startswith("-b"):
                        sheet.write(
                            row,
                            col,
                            f"=SUM({range_gesamt})",
                            self.create_format(border_format_big_lock, {"bg_color": self.count_ma_color}, create_border(row, col)),
                        )
                        start_col += col - start_col + 1
                    else:
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
                range_gesamt = xlsxwriter.utility.xl_range(row, start_col, row, end_col)
                sheet.write(
                    row,
                    col,
                    f"=SUM({range_gesamt})",
                    self.create_format(border_format_big_lock, {"bg_color": self.count_ma_color}, create_border(row, col)),
                )
                col += 1
                row += 1

    def create_employee_statistic_sheet(self):
        ma_information_sheet = "MA Informationen"
        sheet = self.workbook.add_worksheet(ma_information_sheet)
        col = self.start_col + 1
        self.all_sheets.append(sheet)
        sheet.freeze_panes(3, 0)
        sheet.set_row(0, 6)
        sheet.set_column(0, 0, 1)
        sheet.set_column("B:B", None, None, {"hidden": 1})
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
        sheet.write(2, col, "Anwesende KW´s", border_format_big_lock)
        column_names.append("Anwesende KW´s")
        sheet.set_column(col, col, 18)
        col += 1
        sheet.write(2, col, "Abwesende KW´s", border_format_big_lock)
        column_names.append("Abwesende KW´s")
        sheet.set_column(col, col, 18)
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
                            "font_size": 14,
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

    def add_protection(self):
        protect = {
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
        }
        for sheet in self.all_sheets:
            sheet.protect(self.excel_sheets_pw, protect)

    def create_error_sheet(self):
        def find_sentences_with_word(sentences, search_word):
            found_sentences = [sentence for sentence in sentences if search_word in sentence]
            return found_sentences

        format_2 = self.create_format({"left": 1, "right": 1, "top": 1, "bottom": 1, "hidden": 1, "locked": 0, "font_size": 13})
        format_1 = self.create_format({"left": 2, "right": 2, "top": 2, "bottom": 2, "hidden": 1, "locked": 0, "font_size": 13})

        sheet = self.workbook.add_worksheet("Fehler Meldungen")
        # self.all_sheets.append(sheet)

        all_errors = self.create_shifts.error_areas + self.create_shifts.error_shift

        list_area_with_understaffed_emp = find_sentences_with_word(all_errors, "Überprüfen: Bereich unterbesetzt.")
        list_model_2_emp_moved_from_shift = find_sentences_with_word(all_errors, "vom Modell 2 eingetragen.")
        list_missing_quali_for_area = find_sentences_with_word(all_errors, "Info: MA")
        list_shifts_area_not_even = find_sentences_with_word(all_errors, "Info: Schichten sind nicht ausgeglichen.")
        list_area_has_more_then_max_emp = find_sentences_with_word(all_errors, " hat mehr als")
        list_emp_force_to_send = find_sentences_with_word(all_errors, "Überprüfen: MA")

        col = 2
        row = 2
        col_size = 90

        if list_area_with_understaffed_emp:
            sheet.set_column(col, col, col_size)
            row += 1
            for massage in list_area_with_understaffed_emp:
                sheet.write(row, col, str(massage), format_2)
                row += 1
            row = 2
            sheet.write(row, col, "Bereich unterbesetzt", format_1)
            col += 1

        if list_model_2_emp_moved_from_shift:
            sheet.set_column(col, col, col_size * 2)
            row += 1
            for massage in list_model_2_emp_moved_from_shift:
                sheet.write(row, col, str(massage), format_2)
                row += 1
            row = 2
            sheet.write(row, col, "Mitarbeiter mit Model 2 verschoben", format_1)
            col += 1
        if list_missing_quali_for_area:
            sheet.set_column(col, col, col_size * 2)
            row += 1
            for massage in list_missing_quali_for_area:
                sheet.write(row, col, str(massage), format_2)
                row += 1
            row = 2
            sheet.write(row, col, "Mitarbeiter fehlt die Qualifizierung", format_1)
            col += 1
        if list_shifts_area_not_even:
            sheet.set_column(col, col, col_size)
            row += 1
            for massage in list_shifts_area_not_even:
                sheet.write(row, col, str(massage), format_2)
                row += 1
            row = 2
            sheet.write(row, col, "Schichten sind nicht ausgeglichen besetzt", format_1)
            col += 1
        if list_area_has_more_then_max_emp:
            sheet.set_column(col, col, col_size)
            row += 1
            for massage in list_area_has_more_then_max_emp:
                sheet.write(row, col, str(massage), format_2)
                row += 1
            row = 2
            sheet.write(row, col, "Bereich hat mehr als die maximale Anzahl an Mitarbeiter", format_1)
            col += 1
        if list_emp_force_to_send:
            sheet.set_column(col, col, col_size * 2)
            row += 1
            for massage in list_emp_force_to_send:
                sheet.write(row, col, str(massage), format_2)
                row += 1
            row = 2
            sheet.write(row, col, "Bereich hat zu viele Mitarbeiter, weil nicht gewechselt werden konnte", format_1)
            col += 1

    def add_log(self):
        sheet = self.workbook.add_worksheet("Log")
        self.all_sheets.append(sheet)

    def delete_xlsx_file(self):
        try:
            xlsx_file_path = os.path.abspath(self.file_name_xlsx)
            # Check if the XLSM file already exists
            if os.path.exists(xlsx_file_path):
                time.sleep(2)
                os.remove(xlsx_file_path)
                time.sleep(1)
        except Exception as rename_error:
            print(f"Error renaming existing XLSX file: {str(rename_error)}")

    def delete_xlsm_file(self):
        try:
            file = self.file_name_xlsm.replace(f"_{self.timestamp}", "")
            xlsm_file_path = os.path.abspath(file)
            # Check if the XLSM file already exists
            if os.path.exists(xlsm_file_path):
                # Rename the existing file with a timestamp
                timestamp = time.strftime("%Y%m%d%H%M%S")
                new_xlsm_file_path = xlsm_file_path.replace(".xlsm", f"_{timestamp}_OLD.xlsm")
                os.rename(xlsm_file_path, new_xlsm_file_path)
                time.sleep(2)
                os.remove(new_xlsm_file_path)
                time.sleep(1)
        except Exception as err:
            print(f"Error renaming existing XLSM file: {str(err)}")

    def change_workbook_type_and_add_vba(self):
        # Specify the paths for XLSX and XLSM files
        xlsx_file_path = self.file_name_xlsx
        xlsm_file_path = self.file_name_xlsm
        self.delete_xlsm_file()

        xlsm_file_path = os.path.abspath(xlsm_file_path)
        vba_code_new = vba_code
        vba_code_new = vba_code_new.format(self.excel_sheets_pw)
        try:
            # Initialize Excel application
            app = None
            try:
                app = xw.App(visible=False)
            except Exception as e:
                print(f"Open App fails: {str(e)}")

            # Check if Excel application is initialized
            if app:
                wb = None
                try:
                    # Open the XLSX file with xlwings
                    wb = app.books.open(xlsx_file_path)

                    try:
                        wb.api.VBProject.VBComponents(wb.api.CodeName).CodeModule.AddFromString(vba_code_new)
                    except Exception as e:
                        print(
                            f"Cant add Macro: {str(e)}\nExcel macros are not trusted. "
                            "To enable macros, please check Microsoft's documentation: "
                            "https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6"
                        )
                    # Save it as XLSM with error handling for savSe
                    try:
                        wb.save(xlsm_file_path)
                        self.delete_xlsx_file()
                    except Exception as e:
                        print(f"An error occurred while saving the XLSM file: {str(e)}")
                        # Create a manual VBA file
                        manual_vba_file_path = "manual_macro.vba"
                        with open(manual_vba_file_path, "w") as vba_file:
                            vba_file.write(vba_code_new)
                        print(f"VBA code saved to '{manual_vba_file_path}' for manual import and enable macros.")

                except Exception as e:
                    print(f"An error occurred while converting XLSX to XLSM and adding Macro: {str(e)}")
                finally:
                    # Close workbook
                    if wb:
                        wb.close()

                    try:
                        new_xlsm_file_path = xlsm_file_path.replace(f"_{self.timestamp}", "")
                        os.rename(xlsm_file_path, new_xlsm_file_path)
                    except Exception as e:
                        print(f"Error by rename the XLSM File: {e}")

                # Quit Excel application
                app.quit()
        except Exception as e:
            print(f"An error occurred: {str(e)}")

    def run(self):
        print("Creating Excel file...")
        self.workbook = xlsxwriter.Workbook(self.file_name_xlsx)
        self.create_sheets_for_weeks()
        self.create_week_sheet_overlay()
        self.insert_emp_to_weeks()
        self.insert_emp_to_away_list()
        self.insert_emp_to_ma_list()
        self.insert_public_holidays_and_weekend()
        self.create_area_statistic_sheet()
        self.create_employee_statistic_sheet()
        self.create_error_sheet()
        self.workbook.set_properties({"author": self.author})
        self.add_log()
        self.add_protection()
        self.workbook.close()
        print("Add Macro to Excel...")
        self.change_workbook_type_and_add_vba()
        time.sleep(2)
        print(f"Schichtplan {self.year} erstellt.")


if __name__ == "__main__":
    ExcelCreator(
        year=2024,
        start_week=2,
        end_week=50,
        program_runs=10,
        areas_setting={
            "FFK": {"prio": 199, "min": 0, "max": 1},
            "Drehen": {"prio": 1, "min": 3, "max": 3, "fill": True},
            "Bohren": {"prio": 3, "min": 2, "max": 3},
            "Fr\u00e4sen": {"prio": 2, "min": 3, "max": 3},
            "Platzhalter": {"prio": 1111, "min": 0, "max": 3},
            "-b1": {"prio": 1000, "min": 0, "max": 0},
            "B&O": {"prio": 42, "min": 1, "max": 1},
            "\u00d6VH": {"prio": 32, "min": 1, "max": 1},
            "Modul 2": {"prio": 122, "min": 1, "max": 2},
            "Comau": {"prio": 22, "min": 1, "max": 1},
            "Modul 3": {"prio": 52, "min": 2, "max": 3},
            "Dienstleister": {"prio": 622, "min": 0, "max": 3},
            "-b2": {"prio": 8888, "min": 0, "max": 0},
            "B\u00fcro": {"prio": 123123, "min": 0, "max": 4},
        },
        variable_settings=SETTINGS_VARIABLES,
        excel_pw="123",
    )
