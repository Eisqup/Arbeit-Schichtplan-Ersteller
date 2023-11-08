from constants import *


class Employee:
    def __init__(self, name, schicht_model, schicht_rhythmus=[], bereiche=[], urlaub_kw=[], link=None, start_kw_model_2=None):
        self.name = name  # "name": "Employee1"
        self.schicht_model = schicht_model  # "schicht_model": "individuelle Schicht"
        self.schicht_rhythmus = schicht_rhythmus  # "schicht_rhythmus": "Früh, Spät, Nacht"

        self.bereiche = bereiche  # "bereiche": "Drehen, Bohren"
        self.urlaub_kw = [int(kw) for kw in urlaub_kw]  # "urlaub_kw": [2, 8, 14]

        self.link = link  # add here a buddy if he want to work with someone most time together
        self.start_kw_model_2 = start_kw_model_2

        self.start_shift_index_num = None
        self.counter_start_shift = 0
        self.emp_shifts_dict = {}
        self.shift_count_all = {}
        self.employee_list = []
        self.model_2_count_for_shift_switch = 0

    def set_start_shift(self, model_2_start_shift_kw_1_dict):
        # Find lowest values
        lowest_values = sorted(model_2_start_shift_kw_1_dict.values())
        for run in range(len(model_2_start_shift_kw_1_dict) - 1):
            lowest_keys = [key for key, value in model_2_start_shift_kw_1_dict.items() if value == lowest_values[run]]
            if any(lowest_key in self.schicht_rhythmus for lowest_key in lowest_keys):
                # Find the index of the first lowest key in self.schicht_rhythmus
                index = next(i for i, key in enumerate(self.schicht_rhythmus) if key in lowest_keys)
                self.start_shift_index_num = index
                return self.schicht_rhythmus[index]
        self.start_shift_index_num = 0
        return self.schicht_rhythmus[0]

    def get_shift_from_model_2(self, week):
        if self.start_shift_index_num is not None:
            return self.schicht_rhythmus[(week + self.start_shift_index_num - 1) % len(self.schicht_rhythmus)]
        else:
            return False

    def add_shift(self, week, shift):
        # prozentual aufs Jahr berechnen eventuell mit weiteren input der default true ist. anzahle * Jahres Arbeitswochen / reale Arbeitswochen
        self.emp_shifts_dict[week] = shift
        self.emp_shifts_dict = dict(sorted(self.emp_shifts_dict.items()))

    def get_count_of_shifts(self, shift_to_count, week_for_factor=False):
        factor = 1

        if self.emp_shifts_dict and week_for_factor > 1:
            # create a factor to check work weeks / (work week- vacation week)
            count = sum(1 for key, shift in self.emp_shifts_dict.items() if shift == "X" and key != week_for_factor)
            factor = week_for_factor - 1 - count
            if factor == 0:
                factor = 1
            factor = (week_for_factor - 1) / factor
            factor = sum(1 for key, shift in self.emp_shifts_dict.items() if shift == shift_to_count and key != week_for_factor) * factor
            return factor

        factor = sum(1 for shift in self.emp_shifts_dict.values() if shift == shift_to_count) * factor
        return factor

    def get_prio_rhythmus_by_index(self, index):
        if index >= len(self.schicht_rhythmus):
            index = len(self.schicht_rhythmus) - 1
        return self.schicht_rhythmus[index]

    def count_all(self):
        # Initialize counts for each shift type
        fruh_count = 0
        spat_count = 0
        nacht_count = 0
        vacation_count = 0

        # Iterate through emp_shifts_dict to count shifts
        for week, shift in self.emp_shifts_dict.items():
            if shift == SCHICHT_RHYTHMUS[1]:
                fruh_count += 1
            elif shift == SCHICHT_RHYTHMUS[2]:
                spat_count += 1
            elif shift == SCHICHT_RHYTHMUS[3]:
                nacht_count += 1
            elif shift == "X":
                vacation_count += 1

        # Return the counts
        return {SCHICHT_RHYTHMUS[1]: fruh_count, SCHICHT_RHYTHMUS[2]: spat_count, SCHICHT_RHYTHMUS[3]: nacht_count, "Urlaub": vacation_count}
