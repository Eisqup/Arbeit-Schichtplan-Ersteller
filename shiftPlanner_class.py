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
        self.error_areas = []
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
                    # Handle the case where no suitable employee is found for any shift
                    self.error_shift.append(f"No employee found in shift: {shift} for week: {week}")
                    self.assign_shift("No employee found", shift)

            # Remove "No employee found" employees from shifts
            self.shift_1_list = [employee for employee in self.shift_1_list if employee != "No employee found"]
            self.shift_2_list = [employee for employee in self.shift_2_list if employee != "No employee found"]
            self.shift_3_list = [employee for employee in self.shift_3_list if employee != "No employee found"]

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
                self.error_shift.append(f"Shifts are not evenly balanced. Week:{week}")

            self.shift_1_list = self.move_employee_to_area(self.shift_1_list, SCHICHT_RHYTHMUS[1], week)
            self.shift_2_list = self.move_employee_to_area(self.shift_2_list, SCHICHT_RHYTHMUS[2], week)
            self.shift_3_list = self.move_employee_to_area(self.shift_3_list, SCHICHT_RHYTHMUS[3], week)

            self.shift_plan[week] = {
                SCHICHT_RHYTHMUS[1]: self.shift_1_list,
                SCHICHT_RHYTHMUS[2]: self.shift_2_list,
                SCHICHT_RHYTHMUS[3]: self.shift_3_list,
            }

    def move_employee_to_area(self, shift_with_emp, shift, week):
        # Prio order Drehen > Fräsen > Bohren
        areas = {
            BEREICHE[1]: [],  # Drehen
            BEREICHE[2]: [],  # Bohren
            BEREICHE[3]: [],  # Fräsen
        }

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
                    [areas[BEREICHE[1]], areas[BEREICHE[3]], areas[BEREICHE[2]]],
                    [BEREICHE[1], BEREICHE[3], BEREICHE[2]],
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
                    if (
                        not len(areas[BEREICHE[1]]) >= 3
                        and not len(areas[BEREICHE[2]]) >= 3
                        and not len(areas[BEREICHE[3]]) >= 3
                    ):
                        self.error_areas.append(
                            f"Can't find an employee for the area: {area} in week: {week} shift {shift} emp:{selected_employee.name} send to {area_ran}"
                        )
                    # break  # rest of emp cant be sort in

        if (len(areas[BEREICHE[3]]) > len(areas[BEREICHE[1]])) or (len(areas[BEREICHE[2]]) > len(areas[BEREICHE[3]])):
            self.error_areas.append(f"Areas are not evenly balanced. Week:{week}")
        return areas

    def reset(self):
        self.shift_plan = {}
        self.error_shift = []
        self.error_areas = []

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
            if len(self.error_areas) + len(self.error_shift) / 5 < min_length:
                best_shiftPlan = self.shift_plan
                best_error_areas = self.error_areas
                best_error_shift = self.error_shift
                best_employees = self.employees
                best_run = iteration

                min_length = len(self.error_areas) + len(self.error_shift) / 5

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
            print(f"Best run number: {best_run}")
            if self.error_areas:
                for error in self.error_areas:
                    print(decode(error))
            if self.error_shift:
                for error in self.error_shift:
                    print(decode(error))
        else:
            print("Shift planning completed successfully with 0 errors.")
