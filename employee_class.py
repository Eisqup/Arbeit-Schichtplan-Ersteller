from constants import *


class Employee:
    def __init__(self, name, schicht_model, schicht_rhythmus=[], bereiche=[], urlaub_kw=[], urlaub_tage=[], link=None):
        self.name = name  # "name": "Employee1"
        self.schicht_model = schicht_model  # "schicht_model": "individuelle Schicht"
        self.schicht_rhythmus = schicht_rhythmus  # "schicht_rhythmus": "Früh, Spät, Nacht"

        self.bereiche = bereiche  # "bereiche": "Drehen, Bohren"
        self.urlaub_kw = [int(kw) for kw in urlaub_kw]  # "urlaub_kw": [2, 8, 14]
        self.urlaub_tage = urlaub_tage  # "urlaub_tage": ["10.01", "12.02"]

        self.link = link  # add here a buddy if he want to work with someone most time together

        self.start_shift_index_num = None
        self.counter_start_shift = 0
        self.emp_shifts_dict = {}
        self.shift_count_all = {}
        self.employee_list = []

        # Add vacation to the dict
        for vacation_week in self.urlaub_kw:
            self.add_shift(int(vacation_week), "x")

    def set_start_shift(self, list_for_last_rhythmus_in_model_2):
        for rhythmus in self.schicht_rhythmus:
            if rhythmus not in list_for_last_rhythmus_in_model_2:
                self.start_shift_index_num = self.schicht_rhythmus.index(rhythmus)
                return rhythmus

        for rhythmus in reversed(list_for_last_rhythmus_in_model_2):
            if rhythmus in self.schicht_rhythmus:
                self.start_shift_index_num = self.schicht_rhythmus.index(rhythmus)
                return rhythmus

    def get_shift_from_model_2(self, week):
        if self.start_shift_index_num is not None:
            return self.schicht_rhythmus[(week + self.start_shift_index_num - 1) % len(self.schicht_rhythmus)]
        else:
            return False

    def to_dict(self):
        return {
            EMPLOYEE_KEY[0]: self.name,
            EMPLOYEE_KEY[1]: self.schicht_model,
            EMPLOYEE_KEY[2]: self.schicht_rhythmus,
            EMPLOYEE_KEY[3]: self.bereiche,
            EMPLOYEE_KEY[4]: self.urlaub_kw,
            EMPLOYEE_KEY[5]: self.urlaub_tage,
        }

    def add_shift(self, week, shift):
        self.emp_shifts_dict[week] = shift
        self.emp_shifts_dict = dict(sorted(self.emp_shifts_dict.items()))

    def get_count_of_shifts(self, shift_to_count):
        return sum(1 for shift in self.emp_shifts_dict.values() if shift == shift_to_count)

    def get_prio_rhythmus_by_index(self, index):
        if index >= len(self.schicht_rhythmus):
            index = len(self.schicht_rhythmus) - 1
        return self.schicht_rhythmus[index]

    def count_all(self):
        # Initialize counts for each shift type
        fruh_count = 0
        spat_count = 0
        nacht_count = 0

        # Iterate through emp_shifts_dict to count shifts
        for week, shift in self.emp_shifts_dict.items():
            if shift == SCHICHT_RHYTHMUS[1]:
                fruh_count += 1
            elif shift == SCHICHT_RHYTHMUS[2]:
                spat_count += 1
            elif shift == SCHICHT_RHYTHMUS[3]:
                nacht_count += 1

        # Return the counts
        return {
            SCHICHT_RHYTHMUS[1]: fruh_count,
            SCHICHT_RHYTHMUS[2]: spat_count,
            SCHICHT_RHYTHMUS[3]: nacht_count,
            "gesamt": fruh_count + spat_count + nacht_count,
        }
