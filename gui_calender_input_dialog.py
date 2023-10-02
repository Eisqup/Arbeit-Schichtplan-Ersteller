import tkinter as tk
from tkinter import Label, Entry, Button, messagebox, Checkbutton
from constants import *
from calender_manager import CalendarCreator
import textwrap

MAX_LINE_LENGTH = 85


class CalendarInputDialog:
    def __init__(self, parent, settings_areas):
        self.parent = parent
        self.settings_areas = settings_areas
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Calendar Input")
        self.variable_settings = SETTINGS_VARIABLES
        self.settings_entry = {}
        self.a_window = None

        # Create a menu
        menu = tk.Menu(self.dialog)
        self.dialog.config(menu=menu)

        # Add a "Settings" option to the menu
        settings_menu = tk.Menu(menu)
        menu.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Settings", command=self.create_settings_area)

        insert_frame = tk.Frame(self.dialog)
        insert_frame.pack(side="left", padx=10, fill="both", expand=True)

        year_label = Label(insert_frame, text="Kalender Jahr:")
        year_label.pack()

        self.year_entry = Entry(insert_frame)
        self.year_entry.pack()
        self.year_entry.insert(0, "2024")

        start_week_label = Label(insert_frame, text="Start Arbeitswoche (1-52):")
        start_week_label.pack()

        self.start_week_entry = Entry(insert_frame)
        self.start_week_entry.pack()
        self.start_week_entry.insert(0, "1")

        end_week_label = Label(insert_frame, text="End Arbeitswoche (1-52):")
        end_week_label.pack()

        self.end_week_entry = Entry(insert_frame)
        self.end_week_entry.pack()
        self.end_week_entry.insert(0, "52")

        program_run_label = Label(insert_frame, text="Wie oft soll das Programm laufen(desto öfter desto besser)?")
        program_run_label.pack()

        self.program_run_entry = Entry(insert_frame)
        self.program_run_entry.pack()
        self.program_run_entry.insert(0, str(PROGRAM_RUNS))

        ok_button = Button(insert_frame, text="Erstelle Kalender", command=self.ok_clicked, justify="center")
        ok_button.pack(pady=20)

    def create_settings_area(self):
        self.a_window = tk.Toplevel()
        self.a_window.title("Settings")

        # Create a settings area on the right using grid
        settings_frame = tk.Frame(self.a_window)
        settings_frame.pack(side="right", padx=10)

        for i, (variable_name, default_value) in enumerate(self.variable_settings.items()):
            label = Label(settings_frame, text=variable_name)
            label.grid(row=i, column=0, sticky="w")

            # Check if the variable name exists in self.settings_entry
            if variable_name in self.settings_entry:
                var = self.settings_entry[variable_name]
                default_value = var.get()

            if isinstance(default_value, bool):
                var = tk.BooleanVar()
                var.set(default_value)
                widget = Checkbutton(settings_frame, variable=var)
            else:
                var = tk.StringVar()
                var.set(str(default_value))
                widget = Entry(settings_frame, textvariable=var)

            widget.grid(row=i, column=1, padx=5, sticky="e")

            # Get the description and wrap it to multiple lines
            description = SETTING_DESCRIPTION.get(variable_name, "")
            wrapped_description = textwrap.fill(description, MAX_LINE_LENGTH)

            # Display "Beschreibung:" and the wrapped description
            description_label = Label(settings_frame, text=f"{wrapped_description}")
            description_label.grid(row=i, column=2, sticky="w")

            # Store the variable name and the corresponding widget for later use
            self.settings_entry[variable_name] = var

    def ok_clicked(self):
        year = self.year_entry.get()
        start_week = self.start_week_entry.get()
        end_week = self.end_week_entry.get()
        program_run = self.program_run_entry.get()

        for variable_name, var in self.settings_entry.items():
            value = var.get()
            # You can update the constants here
            if variable_name in self.variable_settings:
                if isinstance(value, bool):
                    new_value = bool(value)
                    self.variable_settings[variable_name] = new_value
                else:
                    try:
                        new_value = int(value)
                        self.variable_settings[variable_name] = new_value
                    except ValueError:
                        messagebox.showerror("Invalid Input", f"Please enter a valid numeric value for {variable_name}.")
                        return

        try:
            year = int(year)
            start_week = int(start_week)
            end_week = int(end_week)
            program_run = int(program_run)

            if 1 <= start_week <= 52 and 1 <= end_week <= 52:
                # Valid input, close the dialog and call create_calendar
                if self.a_window is not None:
                    self.a_window.destroy()
                self.dialog.destroy()
                CalendarCreator(year, start_week, end_week, program_run, self.settings_areas, self.variable_settings)
            else:
                messagebox.showerror("Invalid Input", "Please enter valid values or check areas.")
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numeric values.")
