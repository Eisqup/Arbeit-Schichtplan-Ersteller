from constants import *
import random
import json
from employee_class import Employee


class ShiftPlanner:
    def __init__(self, start_week=1, end_week=52):
        self.employees = []
        self.start_week = start_week
        self.end_week = end_week
        self.shift_plan = {}
        self.error_shift = []
        self.error_shift_balance = []
        self.error_areas = []
        self.error_areas_run = []
        self.error_areas_balance = []
        self.error_areas_send = []
        self.run()
        self.update_employees_with_best_plan()

    def load_employees(self):
        try:
            with open(FILE_NAME, "r", encoding="utf-8") as file:
                data = json.load(file)
                employees_data = data["employees"]
        except FileNotFoundError:
            return print("Keine Mitarbeiterliste verfügbar")

        employees = []

        random.shuffle(employees_data)
        # Create the employees classes
        list_of_last_shift_in_model_2 = []
        for employee_data in employees_data:
            # Check if EMPLOYEE_KEY[6] exists in employee_data
            additional_data = employee_data.get(EMPLOYEE_KEY[6], None)

            employee = Employee(
                employee_data[EMPLOYEE_KEY[0]],
                employee_data[EMPLOYEE_KEY[1]],
                employee_data[EMPLOYEE_KEY[2]],
                employee_data[EMPLOYEE_KEY[3]],
                employee_data[EMPLOYEE_KEY[4]],
                employee_data[EMPLOYEE_KEY[5]],
                additional_data,
            )

            # set the start shift for the emp with the schichtmodel 2
            if employee.schicht_model == SCHICHT_MODELS[2]:
                result = employee.set_start_shift(list_of_last_shift_in_model_2)
                if result in list_of_last_shift_in_model_2:
                    list_of_last_shift_in_model_2.remove(result)
                list_of_last_shift_in_model_2.insert(0, result)

            if self.start_week > 1:
                for i in range(1, self.start_week):
                    employee.add_shift(i, "X")
            if self.end_week < 52:
                for i in range(self.end_week, 53):
                    employee.add_shift(i, "X")

            # Add emp to list
            employees.append(employee)

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

    def determine_priority_key_by_lengths(self, lists, keys):
        if not lists or not keys:
            return None  # Return None or raise an exception if either lists or keys are empty

        if len(lists) != len(keys):
            raise ValueError("The number of lists must match the number of keys.")

        min_length = float("inf")  # Initialize with positive infinity
        min_key = None

        for current_list, current_key in zip(lists, keys):
            if len(current_list) < min_length:
                min_length = len(current_list)
                min_key = current_key

        return min_key

    def create_shift_plan(self):
        # Function to check for a buddy an add him to the shift
        def search_linked_emp(selected_employee, available_employees, shift, week):
            if getattr(selected_employee, EMPLOYEE_KEY[6]) is not None:
                for emp in available_employees:
                    if emp.name == selected_employee.link:
                        emp_count = emp.get_count_of_shifts(shift)
                        if (
                            abs(emp_count - selected_employee.get_count_of_shifts(shift)) < 1
                        ):  # Check if Buddy hast +-1 of the same shift
                            # add buddy if the parameter are ok
                            self.assign_shift(emp, shift)
                            available_employees.remove(emp)
                            emp.add_shift(week, shift)

        def creating_errors():
            # check if the list not fair sorted and create error for it
            len_shift_1 = len(self.shift_1_list)
            len_shift_2 = len(self.shift_2_list)
            len_shift_3 = len(self.shift_3_list)

            # Check if the shifts are evenly balanced (+-1)
            shift_counts = [len_shift_1, len_shift_2, len_shift_3]
            max_shifts = max(shift_counts)
            min_shifts = min(shift_counts)

            if max_shifts - min_shifts > 1:
                # The shifts are not evenly balanced, create an error
                # print("here", week)
                self.error_shift_balance.append(f"Shifts are not evenly balanced. Week:{week}")

            shift_lists = [self.shift_1_list, self.shift_2_list, self.shift_3_list]
            for i, shift_list in enumerate(shift_lists, start=1):
                if "No employee found" in shift_list:
                    self.error_shift.append(f"No employee found in shift {SCHICHT_RHYTHMUS[i]} for week {week}")

        self.employees = self.load_employees()
        # Iterate over weeks
        for week in range(self.start_week, self.end_week + 1):
            available_employees = []
            self.shift_1_list = []  # Früh
            self.shift_2_list = []  # Spät
            self.shift_3_list = []  # Nacht

            # Check each employee for vacation weeks
            available_employees = [employee for employee in self.employees if week not in employee.urlaub_kw]
            removed_employees = []

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
                search_linked_emp(employee, available_employees, shift, week)

            if removed_employees:
                for emp in removed_employees:
                    available_employees.remove(emp)

            # Add the rest of them employees with SCHICHT_RHYTHMUS[3] // Add by pro Früh > Spät > Nacht
            # average_shifts = None
            # first_run = True
            while available_employees:
                # Check which shift has the lowest number of employees
                shift = self.determine_priority_key_by_lengths(
                    [self.shift_1_list, self.shift_3_list, self.shift_2_list],
                    [SCHICHT_RHYTHMUS[1], SCHICHT_RHYTHMUS[3], SCHICHT_RHYTHMUS[2]],
                )

                # Find the "best" employee based on your criteria
                # Sort available employees by the number of shifts
                available_employees.sort(key=lambda emp: emp.get_count_of_shifts(shift))

                # Create a list of employees
                matching_employees = []

                # Iterate over schicht_rhythmus indices
                for i in range(6):
                    matching_employees = [
                        employee for employee in available_employees if employee.get_prio_rhythmus_by_index(i) == shift
                    ]

                    if matching_employees:
                        break  # Exit the loop if matching employees are found

                if matching_employees:
                    # Get the lowest count of shifts (first employee in the sorted list)
                    lowest_count = matching_employees[0].get_count_of_shifts(shift)

                    # Filter matching employees with the lowest count
                    matching_employees = [
                        emp for emp in matching_employees if emp.get_count_of_shifts(shift) == lowest_count
                    ]

                    # Create a list to insert emp double for a better chance to get picked
                    random_choose_list = matching_employees.copy()

                    # create a list where Employee had the same shift last week but not the week before
                    if week > 2:
                        for emp in matching_employees:
                            if emp.emp_shifts_dict[week - 1] == shift and emp.emp_shifts_dict[week - 2] != shift:
                                # Employee had the same shift last week and not the week before
                                # Considered a better match
                                random_choose_list.append(emp)

                    selected_employee = random.choice(random_choose_list)

                    # Assign the best employee to the shift
                    self.assign_shift(selected_employee, shift)
                    available_employees.remove(selected_employee)
                    selected_employee.add_shift(week, shift)
                    search_linked_emp(selected_employee, available_employees, shift, week)

                else:
                    # Handle the case where no suitable employee is found for any shif
                    self.assign_shift("No employee found", shift)

            creating_errors()

            self.shift_1_list = self.move_employee_to_area(self.shift_1_list, SCHICHT_RHYTHMUS[1], week)
            self.shift_2_list = self.move_employee_to_area(self.shift_2_list, SCHICHT_RHYTHMUS[2], week)
            self.shift_3_list = self.move_employee_to_area(self.shift_3_list, SCHICHT_RHYTHMUS[3], week)

            if self.error_areas_run:
                # logic to handel changes after shifts are created and still error in area
                def find_and_change_emp_if_error():
                    err_list = self.error_areas_run
                    for err in err_list:
                        available_employees = []
                        all_shifts = {
                            **self.shift_1_list,
                            **self.shift_2_list,
                            **self.shift_3_list,
                        }
                        for shift, shift_dict in all_shifts.items():
                            if shift != err[0]:
                                if all(len(value) >= 3 for value in shift_dict.values()) and any(
                                    len(value) >= 4 for value in shift_dict.values()
                                ):
                                    for area, emp_list in shift_dict.items():
                                        for emp in emp_list:
                                            if emp.schicht_model == SCHICHT_MODELS[3]:
                                                if err[1] in emp.bereiche:
                                                    if len(shift_dict[area]) > 3:
                                                        if (
                                                            emp.get_count_of_shifts(err[0]) > 5
                                                            and err[0] not in emp.schicht_rhythmus
                                                        ):
                                                            break  # same like row 300
                                                        available_employees.append(
                                                            [
                                                                emp,
                                                                emp.name,
                                                                area,
                                                                shift,
                                                                emp.get_count_of_shifts(err[0]),
                                                                False,
                                                            ]
                                                        )
                                                    else:
                                                        ava_emp = []
                                                        for area_2, emp_list_2 in shift_dict.items():
                                                            if area_2 == [BEREICHE[3]]:  # Bohren
                                                                if len(emp_list_2) > 2:
                                                                    for emp_2 in emp_list_2:
                                                                        if area in emp_2.bereiche:
                                                                            ava_emp.append([emp_2, area_2])
                                                            else:
                                                                if len(emp_list_2) > 3:  # Drehen Fräsen
                                                                    for emp_2 in emp_list_2:
                                                                        if area in emp_2.bereiche:
                                                                            ava_emp.append([emp_2, area_2])
                                                        if err[0] in emp.schicht_rhythmus:
                                                            factor = 1  # + emp.schicht_rhythmus.index(err[0])
                                                        else:
                                                            if (
                                                                emp.get_count_of_shifts(err[0]) > 4
                                                                and err[0] not in emp.schicht_rhythmus
                                                            ):
                                                                break
                                                            factor = 4  # factor to let even the ppl without spät schicht this rhythmus the shift do if they can swap
                                                        if ava_emp:
                                                            selected_emp = random.choice(ava_emp)
                                                            available_employees.append(
                                                                [
                                                                    emp,
                                                                    emp.name,
                                                                    area,
                                                                    shift,
                                                                    emp.get_count_of_shifts(err[0]) * factor,
                                                                ]
                                                                + selected_emp
                                                            )
                        if available_employees:
                            # Find the minimum value at index 4 in the original list
                            min_value = min(item[4] for item in available_employees)

                            # Find all lists with the same lowest value at index 4
                            min_value_lists = [item for item in available_employees if item[4] == min_value]

                            selected_emp = random.choice(min_value_lists)
                            # Move emp from shift to the other shift
                            if selected_emp[5]:
                                for shift, shift_dict in all_shifts.items():
                                    if shift == selected_emp[3]:
                                        shift_dict[selected_emp[2]].remove(selected_emp[0])
                                        if selected_emp[5]:
                                            shift_dict[selected_emp[6]].remove(selected_emp[5])
                                            shift_dict[selected_emp[2]].append(selected_emp[5])

                                    if shift == err[0]:
                                        # print(f"move {selected_emp[0].name}")
                                        selected_emp[0].add_shift(week, shift)
                                        shift_dict[err[1]].append(selected_emp[0])
                                        if len(shift_dict[err[1]]) >= 3:
                                            self.error_areas_run.remove(err)

                                    if shift == SCHICHT_RHYTHMUS[1]:
                                        self.shift_1_list = {shift: shift_dict}
                                    elif shift == SCHICHT_RHYTHMUS[2]:
                                        self.shift_2_list = {shift: shift_dict}
                                    elif shift == SCHICHT_RHYTHMUS[3]:
                                        self.shift_3_list = {shift: shift_dict}

                find_and_change_emp_if_error()

            # Crate error for area if are hasnt enogh emp
            if self.error_areas_run:
                for err in self.error_areas_run:
                    self.error_areas.append(f"Bereich unterbesetzt. Week:{week} shift:{err[0]} area:{err[1]}")
                self.error_areas_run = []

            self.shift_plan[week] = {**self.shift_1_list, **self.shift_2_list, **self.shift_3_list}

    def move_employee_to_area(self, shift_with_emp, shift, week):
        # Prio order Drehen > Fräsen > Bohren
        areas = {value: [] for value in BEREICHE.values()}

        shift_with_emp = [employee for employee in shift_with_emp if employee != "No employee found"]

        # Add employee with just one area to work
        for employee in shift_with_emp:
            if len(employee.bereiche) == 1:
                areas[employee.bereiche[0]].append(employee)
                shift_with_emp.remove(employee)

        # Add the rest of the employees
        while shift_with_emp:
            # Find employees with two areas in their bereiche
            employees_with_two_areas = [employee for employee in shift_with_emp if len(employee.bereiche) == 2]

            # Find area with the lowest number sorted by prio Drehen > Fräsen > Bohren
            if all(len(areas[BEREICHE[i]]) >= 3 for i in range(1, 4)):
                area = BEREICHE[1]  # Add rest employees to BEREICH1
            else:
                area = self.determine_priority_key_by_lengths(
                    list(areas.values()),
                    list(BEREICHE.values()),
                )

            area_with_most_employees = max(areas, key=lambda area: len(areas[area]))

            if employees_with_two_areas:
                # Find employees with both area and area_with_most_employees in their bereiche
                employees_with_both_areas = [
                    employee
                    for employee in shift_with_emp
                    if area in employee.bereiche and area_with_most_employees in employee.bereiche
                ]
                employees_with_area = [employee for employee in employees_with_two_areas if area in employee.bereiche]

                if employees_with_both_areas:
                    selected_employee = random.choice(employees_with_both_areas)
                    areas[area].append(selected_employee)
                    shift_with_emp.remove(selected_employee)
                    continue

                elif employees_with_area:
                    selected_employee = random.choice(employees_with_area)
                    areas[area].append(selected_employee)
                    shift_with_emp.remove(selected_employee)
                    continue

            # Add emp with all 3 Area as ability
            employees_with_area = [employee for employee in shift_with_emp if area in employee.bereiche]

            if employees_with_area:
                selected_employee = random.choice(employees_with_area)
                areas[area].append(selected_employee)
                shift_with_emp.remove(selected_employee)

            else:

                def find_employees_with_area_ability(shift_with_emp, areas, area_needed):
                    for employee in shift_with_emp:
                        for assigned_area, employees_list in areas.items():
                            for employee_to_switch in employees_list:
                                if (
                                    area_needed in employee_to_switch.bereiche
                                    and employee_to_switch not in areas[area_needed]
                                    and assigned_area in employee.bereiche
                                ):
                                    shift_with_emp.remove(employee)
                                    areas[assigned_area].remove(employee_to_switch)

                                    areas[area_needed].append(employee_to_switch)
                                    areas[assigned_area].append(employee)

                                    return False  # Found a suitable employee

                    return True  # No suitable employee found

                result = find_employees_with_area_ability(shift_with_emp, areas, area)

                if result:
                    selected_employee = random.choice(shift_with_emp)
                    area_ran = random.choice(selected_employee.bereiche)
                    areas[area_ran].append(selected_employee)
                    shift_with_emp.remove(selected_employee)

                    # -------------------------------------------------------------------------------------------------
                    #                                 ERROR Creating Area
                    # -------------------------------------------------------------------------------------------------

                    # Check if the length of all areas is less than or equal to 3
                    if all(len(areas[area]) <= 3 for area in BEREICHE.values()):
                        self.error_areas_send.append(
                            f"Can't find an employee for the areas in week: {week} shift {shift} emp:{selected_employee.name} send to {area_ran}"
                        )
                    # break  # rest of emp cant be sort in

        # Check if the length of areas is increasing from the last index to the first index
        if len(areas[BEREICHE[1]]) < 3:
            self.error_areas_run.append([shift, BEREICHE[1]])
        if len(areas[BEREICHE[2]]) < 3:
            self.error_areas_run.append([shift, BEREICHE[2]])
        if len(areas[BEREICHE[3]]) < 2:
            self.error_areas_run.append([shift, BEREICHE[3]])

        # Get the list of areas in the order of the Constants
        sorted_areas = [areas[area] for area in BEREICHE.values()]
        for i in range(len(sorted_areas) - 1):
            if abs(len(sorted_areas[i + 1]) - len(sorted_areas[i])) >= 2 and len(sorted_areas[i]) <= 3:
                self.error_areas_balance.append(f"Areas are not evenly balanced. Week:{week} shift:{shift}")
                break  # Exit the loop if the condition is not met

        # -------------------------------------------------------------------------------------------------

        return {shift: {v: areas[v] for v in BEREICHE_EXCEL.values()}}

    def reset(self):
        self.shift_plan = {}
        self.error_shift = []
        self.error_areas = []
        self.error_areas_balance = []
        self.error_areas_send = []
        self.error_shift_balance = []

    def run(self, max_iterations=PROGRAM_RUNS):
        def decode(text):
            return text.encode("utf-8").decode("utf-8")

        min_length = float("inf")
        best_shiftPlan = None
        best_error_areas = None
        best_error_shift = None
        best_employees = None
        best_run = None

        for iteration in range(1, max_iterations):
            self.create_shift_plan()
            print(f"run:{iteration}\nerror_areas:{len(self.error_areas)}\nerror_shift:{len(self.error_shift)}")

            # best cast of errors areas + 1, shift + 1/3
            if len(self.error_areas) + len(self.error_shift) / 10 < min_length:
                best_shiftPlan = self.shift_plan
                best_error_areas = self.error_areas + self.error_areas_balance + self.error_areas_send
                best_error_shift = self.error_shift + self.error_shift_balance
                best_employees = self.employees
                best_run = iteration

                min_length = len(self.error_areas) + len(self.error_shift) / 10

            if not self.error_areas and not self.error_shift:  # Break if no errors
                break

            # Reset data
            self.reset()

        # Save the best data in the class
        self.shift_plan = best_shiftPlan
        self.error_areas = best_error_areas
        self.error_shift = best_error_shift
        self.employees = best_employees

        if self.error_areas or self.error_shift:
            print("Max iterations reached or errors still present.")
            print(
                f"Best run number: {best_run}\nerror_areas:{len(self.error_areas)}\nerror_shift:{len(self.error_shift)}"
            )
            if self.error_areas:
                for error in self.error_areas:
                    print(decode(error))
            if self.error_shift:
                for error in self.error_shift:
                    print(decode(error))
        else:
            print("Shift planning completed successfully with 0 errors.")
