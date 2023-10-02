import json

from constants import *


class DataManager:
    def __init__(self):
        self.data_file = FILE_NAME  # File name
        self.data = None
        self.load_data()

    def load_data(self):
        try:
            with open(self.data_file, "r", encoding="utf-8") as file:
                data = json.load(file)
                data = self.sort_data(data)
        except FileNotFoundError:
            data = {"employees": [], "settings": {}}
            self.save_data(data)

        if "settings" not in data:
            data["settings"] = {}
            data = self.save_data(data)
        self.data = data
        return data

    def save_data(self, data):
        with open(self.data_file, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=2)
        return data

    def load_emp_list(self):
        return self.data["employees"]

    def sort_data(self, data):
        if "employees" in data:

            def custom_sort_key(employee):
                return employee["name"]

            data["employees"] = sorted(data["employees"], key=custom_sort_key)
        return data

    def save_settings_areas(self, data):
        self.data["settings"]["areas"] = data
        self.save_data(self.data)

    def load_settings_areas(self):
        if "areas" in self.data["settings"]:
            return self.data["settings"]["areas"]
        else:
            return {}
