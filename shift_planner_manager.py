from constants import *
import random
import json
from employee_manager import Employee
import os


class ShiftPlanner:
    def __init__(self, start_week=1, end_week=52, program_run=PROGRAM_RUNS, areas=None, variable_settings=None):
        self.start_week = start_week
        self.end_week = end_week

        self.employees = []
        self.shift_plan = {}
        self.error_shift = []
        self.error_areas = []
        self.error_areas_temp = []
        self.error_areas_small = []
        self.error_model_2_emp_move_to_other_shift = []
        self.areas = areas
        self.settings = variable_settings

        # Program
        self.run(program_run)
        self.update_employees_with_best_plan()

    def load_employees(self):
        try:
            with open(FILE_NAME, "r", encoding="utf-8") as file:
                data = json.load(file)
                employees_data = data["employees"]
        except FileNotFoundError:
            return print("Keine Mitarbeiterliste verfügbar")

        employees = []
        # Use a list comprehension to filter employees based on areas
        for emp in employees_data:
            emp_areas = emp[EMPLOYEE_KEY[3]][:]  # Create a copy of the list
            for area in emp_areas:
                if area not in self.areas.keys():
                    emp[EMPLOYEE_KEY[3]].remove(area)

        # Create a set to store all unique areas from employees_data
        all_areas = set()
        for emp in employees_data:
            all_areas.update(emp[EMPLOYEE_KEY[3]])
        # Create a new dictionary for self.areas with only the areas present in employees_data
        filtered_areas = {area: value for area, value in self.areas.items() if area in all_areas}
        # Update self.areas to contain only the filtered areas
        self.areas = filtered_areas

        random.shuffle(employees_data)

        count = 0

        for employee_data in employees_data:
            emp = Employee(
                name=employee_data.get("name"),
                schicht_model=employee_data.get("schicht_model"),
                schicht_rhythmus=employee_data.get("schicht_rhythmus"),
                bereiche=employee_data.get("bereiche"),
                urlaub_kw=employee_data.get("urlaub_kw"),
                link=employee_data.get("link"),
                start_kw_model_2=employee_data.get("start_kw_model_2"),
            )

            # add placeholder to emp dict for the program
            if self.start_week > 1:
                for i in range(1, self.start_week):
                    emp.add_shift(i, "no work")
            if self.end_week < 52:
                for i in range(self.end_week, 53):
                    emp.add_shift(i, "no work")

            # Add emp to the list
            employees.append(emp)

        model_2_start_shift_kw_1_dict = {value: 0 for value in SCHICHT_RHYTHMUS.values()}

        random.shuffle(employees)

        for emp in employees:
            # set the start shift for the emp with the schichtmodel 2
            if emp.schicht_model == SCHICHT_MODELS[2] and emp.start_shift_index_num == None:
                if SETTINGS_VARIABLES["IGNORE_MODEL_2_START_WEEK_LOGIC"] or emp.start_kw_model_2 == "kw1_start":
                    emp.start_shift_index_num = 0
                    model_2_start_shift_kw_1_dict[emp.schicht_rhythmus[0]] += 1
                elif emp.start_kw_model_2 == "arbeitswoche_start":
                    index = (len(emp.schicht_rhythmus) - ((self.start_week - 1) % len(emp.schicht_rhythmus))) % len(emp.schicht_rhythmus)
                    emp.start_shift_index_num = index
                    model_2_start_shift_kw_1_dict[emp.schicht_rhythmus[index]] += 1
                else:
                    shift = emp.set_start_shift(model_2_start_shift_kw_1_dict)
                    model_2_start_shift_kw_1_dict[shift] += 1

                if emp.link:
                    for emp_link in employees:
                        if emp.link == emp_link.name and emp_link.start_shift_index_num == None:
                            emp_link.start_shift_index_num = emp.start_shift_index_num
                            model_2_start_shift_kw_1_dict[emp_link.schicht_rhythmus[emp_link.start_shift_index_num]] += 1

        return employees

    def get_shift_plan(self):
        return self.shift_plan

    def update_employees_with_best_plan(self):
        for week, shifts in self.shift_plan.items():
            for shift, areas in shifts.items():
                for area, emp_list in areas.items():
                    for emp in emp_list:
                        # Find the corresponding Employee object
                        matching_emp = next((e for e in self.employees if e.name == emp.name), None)
                        if matching_emp:
                            # Update the shift for the corresponding Employee object
                            matching_emp.add_shift(week, shift)

    def assign_shift(self, employee, shift):
        if shift == SCHICHT_RHYTHMUS[1]:
            self.shift_1_list.append(employee)
        elif shift == SCHICHT_RHYTHMUS[2]:
            self.shift_2_list.append(employee)
        elif shift == SCHICHT_RHYTHMUS[3]:
            self.shift_3_list.append(employee)

    def move_employee_to_area(self, shift_with_emp, shift, week):
        def find_employees_with_area_ability(shift_with_emp, areas, area_needed):
            for employee in shift_with_emp:
                for assigned_area, employees_list in areas.items():
                    for employee_to_switch in employees_list:
                        if area_needed in employee_to_switch.bereiche and employee_to_switch not in areas[area_needed] and assigned_area in employee.bereiche:
                            shift_with_emp.remove(employee)
                            areas[assigned_area].remove(employee_to_switch)

                            areas[area_needed].append(employee_to_switch)
                            areas[assigned_area].append(employee)

                            return False  # Found a suitable employee

            return True  # No suitable employee found

        # creating areas
        areas = {value: [] for value in self.areas.keys()}

        shift_with_emp = [employee for employee in shift_with_emp if employee != "No employee found"]

        random.shuffle(shift_with_emp)

        # Add employee with just one area to work
        remove_list = []
        for employee in shift_with_emp:
            if len(employee.bereiche) == 1:
                areas[employee.bereiche[0]].append(employee)
                remove_list.append(employee)
        if remove_list:
            for employee in remove_list:
                shift_with_emp.remove(employee)

        # Add the rest of the employees
        while shift_with_emp:
            # -----------------------------------------------------------------------------------------------------------
            area = None
            # Find area with fill key if all areas have max number on emp
            if all(len(areas[area]) >= self.areas[area]["max"] for area in self.areas.keys()):
                # Find the key where "fill" is True
                for key, value in self.areas.items():
                    if "fill" in value:
                        area = key

            else:
                if all(len(areas[area]) >= self.areas[area]["min"] for area in self.areas.keys()):
                    min_or_max = "max"
                else:
                    min_or_max = "min"
                # Calculate the number of employees in each area and store them as (area_name, num_employees) tuples
                area_employees = [(area, len(areas[area])) for area in self.areas.keys()]

                # Calculate the maximum number of employees missing compared to the minimum required
                max_missing_employees = max(self.areas[area][min_or_max] - len(areas[area]) for area in self.areas.keys())

                # Find the areas with the same maximum missing employees
                areas_with_max_missing = [area for area in self.areas.keys() if self.areas[area][min_or_max] - len(areas[area]) == max_missing_employees]

                # Find the area with the lowest priority among areas with the lowest number of employees
                area = min(
                    [area for area, num_employees in area_employees if area in areas_with_max_missing],
                    key=lambda x: self.areas[x]["prio"],
                )

            # -----------------------------------------------------------------------------------------------------------
            # continue_loop = False
            # for length in range(2, len(areas.keys()) + 1):
            filtered_employees = [employee for employee in shift_with_emp if area in employee.bereiche]
            if filtered_employees:
                selected_employee = random.choice(filtered_employees)
                areas[area].append(selected_employee)
                shift_with_emp.remove(selected_employee)
                # continue_loop = True
                continue

            # if continue_loop:
            #     continue

            # if the available emp don´t have the ability for the area check the other areas if we can change one
            result = find_employees_with_area_ability(shift_with_emp, areas, area)

            if result:
                # Check if all areas have max emp or send a emp which don´t have the ability for the area to the fill area
                selected_employee = random.choice(shift_with_emp)
                num = 0
                area_ran = None
                for area_temp in selected_employee.bereiche:
                    # check for the area where the most emp still are missed
                    if len(areas[area_temp]) < self.areas[area_temp]["max"] and num < self.areas[area_temp]["max"] - len(areas[area_temp]):
                        area_ran = area_temp
                        num = self.areas[area_temp]["max"] - len(areas[area_temp])
                # if all areas have max number emp add emp to the fill area
                if area_ran is None:
                    for key, value in self.areas.items():
                        if "fill" in value:
                            area_ran = key
                            # if fill area not in ability in emp.bereiche check other areas for a swap
                            if area_ran not in selected_employee.bereiche:
                                change_emp = []
                                for area_temp in selected_employee.bereiche:
                                    for emp in areas[area_temp]:
                                        if area_ran in emp.bereiche:
                                            change_emp.append((emp, area_temp))
                                # change emp
                                if change_emp:
                                    select_emp = random.choice(change_emp)
                                    areas[area_ran].append(select_emp[0])
                                    areas[select_emp[1]].remove(select_emp[0])
                                    area_ran = select_emp[1]
                                # if no one to change is founded hard error and send emp to fill area
                                else:
                                    self.error_areas.append(f"Überprüfen: MA ({selected_employee.name}) fehlt die Quali fürs {area} in der Woche:{week} Schicht: {shift}. MA zum {area_ran} geschickt trotz fehlender Quali")

                if area_ran is None:
                    area_ran = random.choice(selected_employee.bereiche)
                    self.error_areas.append(f"Überprüfen: Alle Bereich voll und keine MA kann gewechselt werden MA({selected_employee.name}) zum {area_ran} geschickt in der Woche:{week} Schicht: {shift}.")

                areas[area_ran].append(selected_employee)
                shift_with_emp.remove(selected_employee)

                # -------------------------------------------------------------------------------------------------
                #                                 ERROR Creating Area
                # -------------------------------------------------------------------------------------------------

                # Check if the length of all areas is less than or equal to 3
                if all(len(areas[area]) <= self.areas[area]["min"] for area in self.areas.keys()):
                    self.error_areas_small.append(f"Info: MA ({selected_employee.name}) fehlt die Quali fürs {area} in der Woche:{week} Schicht: {shift}. MA zum {area_ran} geschickt")

        # Check if the length of emp is less of the min value an create a error
        for area_key in self.areas.keys():
            if len(areas[area_key]) < self.areas[area_key]["min"]:
                self.error_areas_temp.append([shift, area_key])

        # -------------------------------------------------------------------------------------------------

        return {shift: {v: areas[v] for v in self.areas.keys()}}

    def create_shift_plan(self):
        # Function to check for a buddy an add him to the shift
        def search_linked_emp(selected_employee, available_employees, shift, week):
            if getattr(selected_employee, EMPLOYEE_KEY[6]) is not None:
                for emp in available_employees:
                    if emp.name == selected_employee.link:
                        emp_count = emp.get_count_of_shifts(shift, week)
                        if abs(emp_count - selected_employee.get_count_of_shifts(shift, week)) <= 1 or emp.schicht_model == SCHICHT_MODELS[2] or emp.schicht_model == SCHICHT_MODELS[1]:  # Check if Buddy hast +-1 of the same shift
                            # add buddy if the parameter are ok
                            self.assign_shift(emp, shift)
                            available_employees.remove(emp)
                            emp.add_shift(week, shift)

        def creating_errors():
            if self.error_areas_temp:
                for err in self.error_areas_temp:
                    self.error_areas.append(f"Überprüfen: Bereich unterbesetzt. Woche:{week} Schicht:{err[0]} Bereich:{err[1]}")
                self.error_areas_temp = []

            for shift, areas in {**self.shift_1_list, **self.shift_2_list, **self.shift_3_list}.items():
                for area, emp_list in areas.items():
                    if len(emp_list) > self.areas[area]["max"] and "fill" not in self.areas[area]:
                        max_ = self.areas[area]["max"]
                        self.error_areas_small.append(f"Info: Bereich {area} in der Woche:{week} hat mehr als {max_} Mitarbeiter")

            # check if the list not fair sorted and create error for it
            len_shift_1 = sum(len(employee_list) for employee_list in self.shift_1_list.get(SCHICHT_RHYTHMUS[1], {}).values())
            len_shift_2 = sum(len(employee_list) for employee_list in self.shift_2_list.get(SCHICHT_RHYTHMUS[2], {}).values())
            len_shift_3 = sum(len(employee_list) for employee_list in self.shift_3_list.get(SCHICHT_RHYTHMUS[3], {}).values())

            # Check if the shifts are evenly balanced (+-1)
            shift_counts = [len_shift_1, len_shift_2, len_shift_3]
            max_shifts = max(shift_counts)
            min_shifts = min(shift_counts)

            if max_shifts - min_shifts > 1:
                # The shifts are not evenly balanced, create an error
                self.error_shift.append(f"Info: Schichten sind nicht ausgeglichen. Woche:{week}")

        self.employees = self.load_employees()
        # Iterate over weeks
        for week in range(self.start_week, self.end_week + 1):
            available_employees = []
            self.shift_1_list = []  # Früh
            self.shift_2_list = []  # Spät
            self.shift_3_list = []  # Nacht

            # Check each employee for vacation weeks
            vacation_employees = [employee for employee in self.employees if week in employee.urlaub_kw]
            for emp in vacation_employees:
                emp.add_shift(week, "X")

            available_employees = [employee for employee in self.employees if week not in employee.urlaub_kw]
            removed_employees = []

            random.shuffle(available_employees)

            for employee in available_employees:
                # Check for SCHICHT_MODELS[1]
                if employee.schicht_model == SCHICHT_MODELS[1]:
                    shift = employee.schicht_rhythmus[0]

                # Check for SCHICHT_MODELS[2]
                elif employee.schicht_model == SCHICHT_MODELS[2]:
                    shift = employee.get_shift_from_model_2(week)

                else:
                    continue  # Skip the rest of the code for this employee and move on to the next

                self.assign_shift(employee, shift)
                removed_employees.append(employee)
                employee.add_shift(week, shift)

            if removed_employees:
                for emp in removed_employees:
                    available_employees.remove(emp)

            # Add the rest of them employees with SCHICHT_RHYTHMUS[3]
            while available_employees:
                # Check which shift has the lowest number of employees / sot by prio
                if len(self.shift_1_list) <= len(self.shift_3_list) and len(self.shift_1_list) <= len(self.shift_2_list):
                    shift = SCHICHT_RHYTHMUS[1]
                elif len(self.shift_3_list) <= len(self.shift_2_list) and len(self.shift_3_list) <= len(self.shift_1_list):
                    shift = SCHICHT_RHYTHMUS[3]
                else:
                    shift = SCHICHT_RHYTHMUS[2]

                # Find the "best" employee based on your criteria
                # Sort available employees by the number of shifts
                available_employees.sort(key=lambda emp: emp.get_count_of_shifts(shift, week))

                # Create a list of employees
                matching_employees = []

                # Iterate over schicht_rhythmus indices
                for i in range(6):
                    matching_employees = [employee for employee in available_employees if employee.get_prio_rhythmus_by_index(i) == shift]

                    if matching_employees:
                        break  # Exit the loop if matching employees are found

                if matching_employees:
                    # Get the lowest count of shifts (first employee in the sorted list)
                    lowest_count = matching_employees[0].get_count_of_shifts(shift, week)

                    # Find the lowest count and the lowest count + 1
                    lowest_count = min(emp.get_count_of_shifts(shift, week) for emp in matching_employees)
                    lowest_count_plus_one = lowest_count + 0.5

                    # Filter matching employees with counts in the range [lowest_count, lowest_count + 1]
                    matching_employees = [emp for emp in matching_employees if lowest_count <= emp.get_count_of_shifts(shift, week) <= lowest_count_plus_one]

                    # Create a list to insert emp double for a better chance to get picked
                    random_choose_list = matching_employees.copy()

                    # create a list where Employee had the same shift last week but not the week before
                    if week > 2:
                        for emp in matching_employees:
                            if emp.emp_shifts_dict[week - 1] == shift and emp.emp_shifts_dict[week - 2] != shift and shift == emp.schicht_rhythmus[0]:
                                # Employee had the same shift last week and not the week before
                                # Considered a better match
                                random_choose_list.append(emp)
                                random_choose_list.append(emp)

                    selected_employee = random.choice(random_choose_list)

                    # Assign the best employee to the shift
                    self.assign_shift(selected_employee, shift)
                    available_employees.remove(selected_employee)
                    selected_employee.add_shift(week, shift)
                    search_linked_emp(selected_employee, available_employees, shift, week)

                else:
                    # Handle the case where no suitable employee is found for any shift
                    self.assign_shift("No employee found", shift)

            self.shift_1_list = self.move_employee_to_area(self.shift_1_list, SCHICHT_RHYTHMUS[1], week)
            self.shift_2_list = self.move_employee_to_area(self.shift_2_list, SCHICHT_RHYTHMUS[2], week)
            self.shift_3_list = self.move_employee_to_area(self.shift_3_list, SCHICHT_RHYTHMUS[3], week)

            if self.error_areas_temp:
                # logic to handel changes after shifts are created and still error in area
                def find_and_change_emp_if_error():
                    def emp_without_area_max_run(emp, err):
                        max_number_of_runs = self.settings["MAX_RUNS_IN_YEAR_FOR_MODEL_3_EMP_WITHOUT_SHIFT"]
                        if emp.get_count_of_shifts(err[0], week) >= max_number_of_runs and err[0] not in emp.schicht_rhythmus:
                            return True
                        else:
                            False

                    def use_model_2_emp(emp):
                        if emp.schicht_model == SCHICHT_MODELS[2]:
                            if self.settings["USE_ALL_MODEL_2_EMP_IF_AREAS_IN_OTHER_SHIFTS_UNDERSTAFFED"]:
                                return True
                            if self.settings["USE_MODEL_2_EMP_WITH_ALL_3_SHIFTS_IF_AREAS_IN_OTHER_SHIFTS_UNDERSTAFFED"]:
                                if all(shift in emp.schicht_rhythmus for shift in SCHICHT_RHYTHMUS.values()):
                                    return True
                        return False

                    err_list = self.error_areas_temp
                    for err in err_list:
                        available_employees = []
                        all_shifts = {
                            **self.shift_1_list,
                            **self.shift_2_list,
                            **self.shift_3_list,
                        }
                        for shift, shift_dict in all_shifts.items():
                            if shift != err[0]:
                                # Check if all values are greater than or equal to the respective min values
                                all_values_greater_or_equal = all(len(shift_dict[area]) >= self.areas[area]["min"] for area in self.areas.keys())

                                # Check if at least one value is greater than the minimum for any area
                                at_least_one_value_greater = any(len(shift_dict[area]) > self.areas[area]["min"] for area in self.areas.keys())

                                if all_values_greater_or_equal and at_least_one_value_greater:
                                    for area, emp_list_ in shift_dict.items():
                                        for emp in emp_list_:
                                            if emp.schicht_model == SCHICHT_MODELS[3] or use_model_2_emp(emp):
                                                if err[1] in emp.bereiche:
                                                    if len(shift_dict[area]) > self.areas[area]["min"]:
                                                        if emp_without_area_max_run(emp, err):
                                                            continue
                                                        available_employees.append(
                                                            [
                                                                emp,
                                                                emp.name,
                                                                area,
                                                                shift,
                                                                emp.get_count_of_shifts(err[0], week),
                                                                False,
                                                            ]
                                                        )
                                                    else:
                                                        emp_list = []
                                                        for area_2, emp_list_2 in shift_dict.items():
                                                            if len(emp_list_2) > self.areas[area_2]["min"]:
                                                                for emp_2 in emp_list_2:
                                                                    if area in emp_2.bereiche:
                                                                        emp_list.append([emp_2, area_2])

                                                        if err[0] in emp.schicht_rhythmus:
                                                            factor = 1  # + emp.schicht_rhythmus.index(err[0])
                                                        else:
                                                            if emp_without_area_max_run(emp, err):
                                                                continue
                                                            factor = self.settings["FACTOR_FOR_MODEL_3_EMP_WITHOUT_SHIFT"]  # factor to let even the ppl without schicht this rhythmus the shift do if they can swap
                                                        if emp_list:
                                                            selected_emp = random.choice(emp_list)
                                                            available_employees.append(
                                                                [
                                                                    emp,
                                                                    emp.name,
                                                                    area,
                                                                    shift,
                                                                    emp.get_count_of_shifts(err[0], week) * factor,
                                                                ]
                                                                + selected_emp
                                                            )

                        model_2 = False
                        filtered_employees = [emp for emp in available_employees if err[0] in emp[0].schicht_rhythmus and emp[0].schicht_model == SCHICHT_MODELS[3]]
                        if filtered_employees:
                            available_employees = filtered_employees  # Filer the list for emp which have the shift in there schicht rhythmus
                        else:
                            filtered_employees = [emp for emp in available_employees if err[0] not in emp[0].schicht_rhythmus and emp[0].schicht_model == SCHICHT_MODELS[3]]
                            if filtered_employees:
                                available_employees = filtered_employees  # Filer the list for emp which have the shift in there schicht rhythmus
                            else:
                                # filter last time if just model_2 emp left in available list
                                model_2 = True

                        if available_employees:
                            if model_2:
                                # sort model 2 emp and select one. sort for how often the emp was added to a shift
                                available_employees.sort(key=lambda emp: emp[0].model_2_count_for_shift_switch)
                                lowest_count = available_employees[0][0].model_2_count_for_shift_switch
                                available_employees = [emp for emp in available_employees if lowest_count == emp[0].model_2_count_for_shift_switch]

                                selected_emp = random.choice(available_employees)
                                selected_emp[0].model_2_count_for_shift_switch += 1
                            else:
                                if len(available_employees) == 1:
                                    min_value = available_employees[0][4]
                                    second_min_value = available_employees[0][4]
                                else:
                                    min_value = min(item[4] for item in available_employees)
                                    second_min_value = float("inf")

                                    for item in available_employees:
                                        if item[4] < min_value:
                                            second_min_value = min_value
                                            min_value = item[4]
                                        elif min_value < item[4] < second_min_value:
                                            second_min_value = item[4]

                                # Find all lists with the same lowest value at index 4
                                min_value_lists = [item for item in available_employees if item[4] == min_value]
                                second_min_value_lists = [item for item in available_employees if item[4] == second_min_value]

                                selected_emp = random.choice(min_value_lists + second_min_value_lists)
                            # Move emp from shift to the other shift
                            if use_model_2_emp(selected_emp[0]):
                                # creating hard error if emp add from model 2 to compensate understaffed shift
                                self.error_model_2_emp_move_to_other_shift.append(f"Überprüfen: Bereich {err[1]} unterbesetzt in der {err[0]} Woche {week}. MA {selected_emp[1]} aus der {selected_emp[3]} vom Modell 2 eingetragen. Bitte Überprüfen!!!")

                            for shift, shift_dict in all_shifts.items():
                                if selected_emp[5]:
                                    if shift == selected_emp[3]:
                                        shift_dict[selected_emp[6]].remove(selected_emp[5])
                                        shift_dict[selected_emp[2]].append(selected_emp[5])

                                if shift == err[0]:
                                    # print(f"move {selected_emp[0].name}")
                                    selected_emp[0].add_shift(week, shift)
                                    shift_dict[err[1]].append(selected_emp[0])
                                    if len(shift_dict[err[1]]) >= self.areas[err[1]]["min"]:
                                        self.error_areas_temp.remove(err)

                                # remove emp from the where he was
                                if shift == selected_emp[3]:
                                    shift_dict[selected_emp[2]].remove(selected_emp[0])

                                if shift == SCHICHT_RHYTHMUS[1]:
                                    self.shift_1_list = {shift: shift_dict}
                                elif shift == SCHICHT_RHYTHMUS[2]:
                                    self.shift_2_list = {shift: shift_dict}
                                elif shift == SCHICHT_RHYTHMUS[3]:
                                    self.shift_3_list = {shift: shift_dict}

                find_and_change_emp_if_error()

            # Creating errors
            creating_errors()

            self.shift_plan[week] = {**self.shift_1_list, **self.shift_2_list, **self.shift_3_list}

    def reset(self):
        self.shift_plan = {}
        self.error_shift = []
        self.error_areas = []
        self.error_areas_small = []
        self.error_model_2_emp_move_to_other_shift = []

    def run(self, max_iterations):
        def decode(text):
            return text.encode("utf-8").decode("utf-8")

        # Define a function to clear the screen
        def clear_terminal():
            os.system("clear" if os.name == "posix" else "cls")

        def error_points_for_model_3_balance():
            # here i check the max value of the most wanted rhythmus(emp.rhythmus[0]) form model 3 - average of all
            # I want the differences will be added as error-points
            err_points = 0
            for shift in SCHICHT_RHYTHMUS.values():
                count_emp = 0
                sum_of_shifts = 0
                max_num = 0
                for emp in self.employees:
                    if emp.schicht_model == SCHICHT_MODELS[3]:
                        if shift == emp.schicht_rhythmus[0]:
                            number = emp.get_count_of_shifts(shift, self.end_week)
                            sum_of_shifts += number
                            count_emp += 1
                            if number > max_num:
                                max_num = number
                if count_emp > 0:
                    err_points += max_num - sum_of_shifts / count_emp

            return err_points

        min_length = float("inf")
        best_shiftPlan = None
        best_error_areas = None
        best_error_shift = None
        best_employees = None
        best_run = None

        for iteration in range(1, max_iterations + 1):
            self.create_shift_plan()
            points_error_hard = len(self.error_areas) + len(self.error_model_2_emp_move_to_other_shift) / 2
            points_error_soft = len(self.error_areas_small) / 20
            points_error_shift = len(self.error_shift) / 50
            points_emp_balance = error_points_for_model_3_balance()
            all_points = points_emp_balance + points_error_shift + points_error_soft + points_error_hard  ##all error points

            # best cast of errors areas + 1, shift + 1/3
            if all_points < min_length:
                best_shiftPlan = self.shift_plan
                best_error_areas = self.error_areas + self.error_areas_small + self.error_model_2_emp_move_to_other_shift
                best_error_shift = self.error_shift
                best_employees = self.employees
                best_run = iteration

                best_points_error_hard = points_error_hard
                best_points_error_soft = points_error_soft + points_error_shift
                best_points_balance = points_emp_balance

                min_length = all_points

            # Call the clear_screen function to clear the screen
            if iteration == 1:
                print("Bitte warten...")

            if iteration % 10 == 0:
                clear_terminal()
                print("Bitte warten...")
                print(f"{iteration}/{max_iterations}")
                print(
                    f"Best run number: {best_run}  error_areas:{len(best_error_areas)} error_shift: {len(best_error_shift)} error_points_hard: {round(best_points_error_hard,2)} error_points_soft: {round(best_points_error_soft,2)} balance_points: {round(best_points_balance,2)}"
                )

            if all_points == 0:  # Break if all is perfect
                break

            # Reset data
            self.reset()

        # Save the best data in the class
        self.shift_plan = best_shiftPlan
        self.error_areas = best_error_areas
        self.error_shift = best_error_shift
        self.employees = best_employees

        if self.error_areas or self.error_shift:
            print("Max iterations reached.")
            print(
                f"Best run number: {best_run}\nerror_areas:{len(self.error_areas)}\nerror_shift:{len(self.error_shift)}\nerror_points_hard: {round(best_points_error_hard,2)}\nerror_points_soft: {round(best_points_error_soft,2)}\nbalance_points: {round(best_points_balance,2)}"
            )
            # if self.error_areas:
            #     for error in self.error_areas:
            #         print(decode(error))
            # if self.error_shift:
            #     for error in self.error_shift:
            #         print(decode(error))
        else:
            print("Shift planning completed successfully with 0 errors.")
