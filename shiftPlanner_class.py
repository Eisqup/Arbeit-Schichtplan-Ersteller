from constants import *
import random


class ShiftPlanner:
    def __init__(self, employees, start_week=1, end_week=52):
        self.employees = employees
        self.start_week = start_week
        self.end_week = end_week
        self.shift_plan = {}
        self.error_shift = []
        self.error_areas = []
        self.run()

    def get_shift_plan(self):
        return self.shift_plan

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

    def shift_creator(self):
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

            if removed_employees:
                for emp in removed_employees:
                    available_employees.remove(emp)

            # Add the rest of them employees with SCHICHT_RHYTHMUS[3] // Add by pro Früh > Spät > Nacht
            while available_employees:
                # Check which shift has the lowest number of employees
                shift = self.determine_priority_key_by_lengths(
                    [self.shift_1_list, self.shift_2_list, self.shift_3_list],
                    [SCHICHT_RHYTHMUS[1], SCHICHT_RHYTHMUS[2], SCHICHT_RHYTHMUS[3]],
                )

                # Find the "best" employee based on your criteria
                # Sort available employees by the number of shifts
                available_employees.sort(key=lambda emp: emp.get_count_of_shifts(shift))

                # Create a list of employees
                matching_employees = []

                # Iterate over schicht_rhythmus indices
                for i in range(5):
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

                    # Randomly select one employee from the filtered list
                    selected_employee = random.choice(matching_employees)

                    # Assign the best employee to the shift
                    self.assign_shift(selected_employee, shift)
                    available_employees.remove(selected_employee)
                    selected_employee.add_shift(week, shift)
                else:
                    # Handle the case where no suitable employee is found for any shift

                    self.error_shift.append(f"No employee found in shift: {shift} for week: {week}")
                    self.assign_shift("No employee found", shift)

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

        # Remove "No employee found" strings
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

                                    return True  # Found a suitable employee

                    return False  # No suitable employee found

                if not find_employees_with_area_ability(shift_with_emp, areas, area):
                    for emp in shift_with_emp:
                        area_ran = random.choice(emp.bereiche)
                        areas[area].append(emp)
                        self.error_areas.append(
                            f"Can't find an employee for the area: {area} in week: {week} emp:{emp.name} send to {area_ran}"
                        )
                        shift_with_emp.remove(emp)
                    # break  # rest of emp cant be sort in

        return areas

    def reset(self):
        self.shift_plan = {}
        self.error_shift = []
        self.error_areas = []

    def run(self, max_iterations=100):
        min_length = float("inf")
        best_shiftPlan = None
        best_error_areas = None
        best_error_shift = None
        best_employees = None

        for iteration in range(max_iterations):
            self.shift_creator()
            print(f"run:{iteration + 1}\nerror_areas:{len(self.error_areas)}\nerror_shift:{len(self.error_shift)}")

            # best cast of errors areas + 1, shift + 1/3
            if len(self.error_areas) + len(self.error_shift) / 3 < min_length:
                best_shiftPlan = self.shift_plan
                best_error_areas = self.error_areas
                best_error_shift = self.error_shift
                best_employees = self.employees
                min_length = len(self.error_areas) + len(self.error_shift) / 3

            if not self.error_areas and not self.error_shift:  # Break if no errors
                break

            # Reset data
            self.reset()

        # Save the best data in the class
        self.shift_plan = best_shiftPlan
        self.error_areas = best_error_areas
        self.error_shift = best_error_shift
        self.employees = best_employees

        if best_error_areas or best_error_shift:
            print("Max iterations reached or errors still present.")
            if self.error_areas:
                for error in self.error_areas:
                    print(error)
            if self.error_shift:
                for error in self.error_shift:
                    print(error)
        else:
            print("Shift planning completed successfully with 0 errors.")
