import tkinter as tk
from tkinter import Label, Button, StringVar, Radiobutton, Checkbutton, ttk, messagebox, Entry
import json
import re

from calender_creator import create_calender
from constants import *


class GUIManager:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Employee Creator")
        self.start_row = 1
        self.button_size = 16
        self.data = {}
        # Create an empty list to store selected "Rhythmus" options
        self.selected_rhythmus_list = []

        # Create a dictionary to hold the Checkbutton variables
        self.selected_bereich_list = []
        self.bereich_vars = []

        self.load_data()

        # reader grid
        self.create_gui()

        # Populate the list of employee names
        self.load_employee_names_for_listbox()

    def sort_data(self):
        if "employees" in self.data:

            def custom_sort_key(employee):
                name = employee[EMPLOYEE_KEY[0]]
                # Split the name into a tuple of string and numeric parts
                parts = re.split(r"(\d+)", name)
                # Convert numeric parts to integers for sorting
                numeric_parts = [int(part) if part.isdigit() else part for part in parts]
                return numeric_parts

            self.data["employees"] = sorted(self.data["employees"], key=custom_sort_key)

    def load_data(self):
        try:
            with open(FILE_NAME, "r", encoding="utf-8") as file:
                self.data = json.load(file)
                self.sort_data()
        except FileNotFoundError:
            self.data = {"employees": []}

            # Save the initial data to the file
            with open(FILE_NAME, "w", encoding="utf-8") as file:
                json.dump(self.data, file, indent=2)

    def save_data_to_file(self):
        # Save the updated data back to the file
        with open(FILE_NAME, "w", encoding="utf-8") as file:
            json.dump(self.data, file, indent=2)

    def create_gui(self):
        row = self.start_row
        # Label for titles on the left and right
        title_font = ("TkDefaultFont", 14, "bold underline")
        title_justify = "center"

        # Label for titles on the left and right
        self.left_title_label = Label(
            self.window, text="Mitarbeiter Erstellen/Bearbeiten", font=title_font, justify=title_justify
        )
        self.left_title_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.right_title_label = Label(
            self.window, text="Vorhandene Mitarbeiter", font=title_font, justify=title_justify
        )
        self.right_title_label.grid(row=0, column=4, padx=10, pady=5, sticky="w")

        row += 2
        # Label for entering the employee's name
        self.name_label = Label(self.window, text="Enter Name:")
        self.name_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")

        self.name_entry = tk.Entry(self.window)
        self.name_entry.grid(row=row, column=1, padx=10, pady=5)

        row += 1

        # Separator between sections
        self.separator1 = ttk.Separator(self.window, orient="horizontal")
        self.separator1.grid(row=row + 1, columnspan=3, sticky="ew", padx=10, pady=10)
        row += 2

        # Label for selecting shift model
        self.shift_model_label = Label(self.window, text="Schicht Modell:")
        self.shift_model_label.grid(row=row + 1, column=0, padx=10, pady=5, sticky="w")

        # Radio buttons for selecting shift model
        self.shift_model_var = tk.IntVar()
        self.shift_model_var.set(1)  # Set an initial value to None
        for key, value in SCHICHT_MODELS.items():
            self.shift_model_radio = Radiobutton(self.window, text=value, variable=self.shift_model_var, value=key)
            self.shift_model_radio.grid(row=row, column=1, padx=10, pady=5, sticky="w")
            row += 1

        # Separator between sections
        self.separator1 = ttk.Separator(self.window, orient="horizontal")
        self.separator1.grid(row=row + 1, columnspan=3, sticky="ew", padx=10, pady=10)
        row += 2

        # Label for selecting Rhythmus
        self.rhythmus_label = Label(self.window, text="Schicht Zyklus:\n (bei Individuell nach prio sortieren)")
        self.rhythmus_label.grid(row=row + 1, column=0, padx=10, pady=5, sticky="w")

        # Buttons for selecting Rhythmus
        self.rhythmus_options = list(SCHICHT_RHYTHMUS.values())
        self.selected_rhythmus_var = StringVar()
        self.selected_rhythmus_var.set(None)  # Set an initial value
        for option in self.rhythmus_options:
            self.rhythmus_button = Button(
                self.window,
                text=option,
                width=self.button_size,
                command=lambda o=option: self.add_to_selected_rhythmus(o),
            )
            self.rhythmus_button.grid(row=row, column=1, padx=12, pady=5, sticky="w")
            row += 1

        # Label for displaying selected Rhythmus
        self.selected_rhythmus_label = Label(self.window, text="Selected Rhythmus:")
        self.selected_rhythmus_label.grid(row=row - 3, column=2, padx=10, pady=5, sticky="w")

        self.selected_rhythmus_display = Label(self.window, text="")
        self.selected_rhythmus_display.grid(row=row - 2, column=2, padx=10, pady=5, sticky="w")

        # Clear button for Rhythmus
        self.clear_rhythmus_button = Button(
            self.window,
            text="RESET Zyklus",
            width=self.button_size,
            command=self.reset_selected_rhythmus_label,
        )
        self.clear_rhythmus_button.grid(row=row - 1, column=2, padx=10, pady=5, sticky="w")

        # Separator between sections
        self.separator1 = ttk.Separator(self.window, orient="horizontal")
        self.separator1.grid(row=row + 1, columnspan=3, sticky="ew", padx=10, pady=10)
        row += 2

        # Label for selecting Bereich
        self.bereich_label = Label(self.window, text="Mitarbeiter Fähigkeiten:")
        self.bereich_label.grid(row=row + 2, column=0, padx=10, pady=5, sticky="w")

        # Checkbuttons for selecting Bereich
        row += 1
        for key, value in BEREICHE.items():
            self.bereich_var = tk.BooleanVar()
            self.bereich_vars.append(self.bereich_var)
            self.bereich_checkbutton = Checkbutton(
                self.window,
                text=value,
                variable=self.bereich_var,
                command=lambda o=value: self.add_to_selected_bereich(o),
            )
            self.bereich_checkbutton.grid(row=row, column=1, padx=10, pady=5, sticky="w")
            row += 1

        # Separator between sections
        self.separator1 = ttk.Separator(self.window, orient="horizontal")
        self.separator1.grid(row=row + 1, columnspan=3, sticky="ew", padx=10, pady=10)
        row += 2

        # Label for entering urlaub_kw
        self.urlaub_kw_label = Label(self.window, text="Enter Urlaub (KW) (sepertatet with space):")
        self.urlaub_kw_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")

        self.urlaub_kw_entry = tk.Entry(self.window)
        self.urlaub_kw_entry.grid(row=row, column=1, padx=10, pady=5)

        row += 1

        # Label for entering urlaub_day
        self.urlaub_day_label = Label(
            self.window, text="Enter Urlaub (Tage)(sepertatet with space):\nBespiel: DD.MM DD.MM"
        )
        self.urlaub_day_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")

        self.urlaub_day_entry = tk.Entry(self.window)
        self.urlaub_day_entry.grid(row=row, column=1, padx=10, pady=5)

        # Submit button
        row += 1
        self.submit_button = tk.Button(
            self.window, text="Daten Speichern", width=self.button_size, command=self.save_employee_data
        )
        self.submit_button.grid(row=row, columnspan=2, pady=10)

        # reset all
        self.reset_button = tk.Button(
            self.window, text="RESET ALL", width=self.button_size, command=self.reset_form_fields
        )
        self.reset_button.grid(row=row, column=1, columnspan=2, pady=10)

        # Vertical Separator in column 3
        self.separator1_vertical = ttk.Separator(self.window, orient="vertical")
        self.separator1_vertical.grid(row=0, column=3, rowspan=100, sticky="ns")

        # Listbox widget to display employee names
        self.employee_listbox = tk.Listbox(self.window, height=5, width=30)
        self.employee_listbox.grid(row=self.start_row + 1, column=4, padx=10, pady=5, rowspan=row - 4, sticky="nsew")

        # Link employee
        self.delete_button = tk.Button(
            self.window,
            text="Hinzufügen/Entfernen\nMitarbeiter link",
            width=self.button_size,
            command=self.link_employees_window,
        )
        self.delete_button.grid(row=row - 2, column=3, columnspan=2, pady=10)

        # Delete button
        self.delete_button = tk.Button(
            self.window, text="Mitarbeiter löschen", width=self.button_size, command=self.delete_employee
        )
        self.delete_button.grid(row=row, column=3, columnspan=2, pady=10)

        # Create a Load button
        self.load_button = tk.Button(
            self.window, text="Daten Laden", width=self.button_size, command=self.load_selected_employee_to_GUI
        )
        self.load_button.grid(row=row - 1, column=4, padx=10, pady=5)

        # Create a Shift plan button
        self.load_button = tk.Button(
            self.window, text="Erstelle Schichtplan", width=self.button_size, command=self.create_calender_button
        )
        self.load_button.grid(row=row + 1, column=4, padx=10, pady=5)

        # Create a frame to display the sneak peek
        self.sneak_peek_frame = tk.Frame(self.window)
        self.sneak_peek_frame.grid(row=self.start_row + 1, column=5, padx=10, pady=5, rowspan=row - 3, sticky="nsew")

        # Label to display selected employee's info in the sneak peek
        self.sneak_peek_info_label = Label(self.sneak_peek_frame, text="", justify="left")
        self.sneak_peek_info_label.pack()

        # Bind the Listbox selection event to update the sneak peek
        self.employee_listbox.bind("<<ListboxSelect>>", self.update_sneak_peek)

    def save_employee_data(self):
        name = self.name_entry.get()
        selected_shift_model = self.shift_model_var.get()
        urlaub_kw = self.split_and_check_input(self.urlaub_kw_entry.get(), is_weeks=True)
        urlaub_day = self.split_and_check_input(self.urlaub_day_entry.get())
        self.selected_bereich_list = [
            list(BEREICHE.values())[index] for index, var in enumerate(self.bereich_vars) if var.get()
        ]

        if urlaub_kw is None or urlaub_day is None:
            return

        # Check if required fields are empty
        if not name or not selected_shift_model or not self.selected_rhythmus_list or not self.selected_bereich_list:
            empty_fields = []
            if not name:
                empty_fields.append("Name")
            if not selected_shift_model:
                empty_fields.append("Shift Model")
            if not self.selected_rhythmus_list:
                empty_fields.append("Schicht Rhythmus")
            if not self.selected_bereich_list:
                empty_fields.append("Bereiche")

            messagebox.showerror("Error", f"The following fields are empty:\n\n {',   '.join(empty_fields)}")
        else:
            # Check if the name already exists in the database
            employee_exists = False
            for employee in self.data.get("employees", []):
                if employee.get(EMPLOYEE_KEY[0]) == name:
                    employee_exists = True
                    break

            if employee_exists:
                result = messagebox.askquestion("Confirm", "Mitarbeiter schon vorhanden. Daten überschreiben?")
                if result == "yes":
                    # Overwrite the existing employee's data
                    for index, employee in enumerate(self.data.get("employees", [])):
                        if employee.get(EMPLOYEE_KEY[0]) == name:
                            # Remove unchecked "Bereich" abilities from the list
                            self.data["employees"][index] = {
                                EMPLOYEE_KEY[0]: name,
                                EMPLOYEE_KEY[1]: SCHICHT_MODELS[selected_shift_model],
                                EMPLOYEE_KEY[2]: self.selected_rhythmus_list,
                                EMPLOYEE_KEY[3]: self.selected_bereich_list,
                                EMPLOYEE_KEY[4]: urlaub_kw,
                                EMPLOYEE_KEY[5]: urlaub_day,
                            }
            else:
                # Create the employee data dictionary
                employee_data = {
                    EMPLOYEE_KEY[0]: name,
                    EMPLOYEE_KEY[1]: SCHICHT_MODELS[selected_shift_model],
                    EMPLOYEE_KEY[2]: self.selected_rhythmus_list,
                    EMPLOYEE_KEY[3]: self.selected_bereich_list,
                    EMPLOYEE_KEY[4]: urlaub_kw,
                    EMPLOYEE_KEY[5]: urlaub_day,
                }
                self.data["employees"].append(employee_data)
                self.sort_data()

            # Save the updated data back to the file
            self.save_data_to_file()

            if not employee_exists:
                message = "Mitarbeiter hinzugefügt:\n"
                for key, value in employee_data.items():
                    message += f"{key}: {value}\n"
            else:
                message = "Mitarbeiterdaten aktualisiert\n"

            print(message)

        # Insert the newly created/updated employee's name into the Listbox
        self.employee_listbox.delete(0, tk.END)
        self.load_employee_names_for_listbox()

        # Clear the form fields after creating or updating an employee
        self.reset_form_fields()
        self.load_data()

    def update_sneak_peek(self, event):
        # Get the selected index from the listbox
        selected_index = self.employee_listbox.curselection()
        if selected_index:
            selected_index = int(selected_index[0])

            if "employees" in self.data and selected_index < len(self.data["employees"]):
                selected_employee = self.data["employees"][selected_index]
                # Update the sneak peek with the selected employee's info
                info_text = "\n".join([f"{key}: {value}" for key, value in selected_employee.items()])
                self.sneak_peek_info_label.config(text=info_text)

    # Create a method to delete an employee
    def delete_employee(self):
        selected_index = self.employee_listbox.curselection()
        if selected_index:
            selected_index = int(selected_index[0])
            if "employees" in self.data and selected_index < len(self.data["employees"]):
                employee_to_delete = self.data["employees"][selected_index]
                # Ask for confirmation before deleting
                result = messagebox.askquestion(
                    "Confirm Deletion",
                    f"Do you want to delete the employee:\n{employee_to_delete[EMPLOYEE_KEY[0]]}?",
                )
                if result == "yes":
                    # Delete the employee
                    deleted_employee = self.data["employees"].pop(selected_index)

                    # Save the updated data back to the file
                    self.save_data_to_file()

                    # Clear the form fields
                    self.reset_form_fields()
                    # Update the Listbox by removing the deleted employee's name
                    self.employee_listbox.delete(selected_index)
                    # Display a message confirming the deletion
                    messagebox.showinfo("Deleted", f"Employee deleted:\n{deleted_employee[EMPLOYEE_KEY[0]]}")

    def load_selected_employee_to_GUI(self):
        # Get the selected index from the listbox
        selected_index = self.employee_listbox.curselection()
        if selected_index:
            selected_index = int(selected_index[0])
            if "employees" in self.data and selected_index < len(self.data["employees"]):
                # Get the selected employee's data
                selected_employee = self.data["employees"][selected_index]
                # Populate the input fields on the left side
                self.name_entry.delete(0, tk.END)
                self.name_entry.insert(0, selected_employee[EMPLOYEE_KEY[0]])
                # Convert the selected shift model to title case before checking its index
                shift_model = selected_employee[EMPLOYEE_KEY[1]]
                shift_model_index = list(SCHICHT_MODELS.values()).index(shift_model) + 1
                self.shift_model_var.set(shift_model_index)
                # Clear previous selections
                for var in self.bereich_vars:
                    var.set(False)
                # Check the Bereich checkboxes based on the selected employee's data
                for bereich in selected_employee[EMPLOYEE_KEY[3]]:
                    index = list(BEREICHE.values()).index(bereich)
                    if 0 <= index < len(self.bereich_vars):
                        self.bereich_vars[index].set(True)
                # Populate the Rhythmus buttons based on the selected employee's data
                self.reset_selected_rhythmus_label()
                for rhythmus in selected_employee[EMPLOYEE_KEY[2]]:
                    if rhythmus in self.rhythmus_options:
                        self.add_to_selected_rhythmus(rhythmus)
                # Populate the Urlaub fields
                self.urlaub_kw_entry.delete(0, tk.END)
                self.urlaub_kw_entry.insert(0, " ".join(map(str, selected_employee[EMPLOYEE_KEY[4]])))
                self.urlaub_day_entry.delete(0, tk.END)
                self.urlaub_day_entry.insert(0, " ".join(selected_employee[EMPLOYEE_KEY[5]]))

    # Add this method to populate the Listbox with employee names
    def load_employee_names_for_listbox(self):
        employee_names = []
        if "employees" in self.data:
            employee_names = [employee[EMPLOYEE_KEY[0]] for employee in self.data["employees"]]
        for name in employee_names:
            self.employee_listbox.insert(tk.END, name)

    def add_to_selected_bereich(self, option):
        self.selected_bereich_list.append(option)

    # Function to split the input into a list
    def split_and_check_input(self, input_str, is_weeks=False):
        values = input_str.split()

        if is_weeks:
            # Validate weeks (1-52)
            for value in values:
                try:
                    week = int(value)
                    if not (1 <= week <= 52):
                        raise ValueError("Week must be between 1 and 52")
                except ValueError:
                    messagebox.showerror(
                        "Error", f"Invalid week value: {value}. Weeks must be integers between 1 and 52."
                    )
                    return None  # Exit and handle the error

        else:
            # Validate dates (DD.MM)
            date_pattern = re.compile(r"^(0[1-9]|[12][0-9]|3[01])\.(0[1-9]|1[0-2])$")
            for value in values:
                if not date_pattern.match(value):
                    messagebox.showerror("Error", f"Invalid date format: {value}. Dates must be in DD.MM format.")
                    return None  # Exit and handle the error

        return values

    def reset_form_fields(self):
        self.name_entry.delete(0, tk.END)
        self.shift_model_var.set(1)
        self.reset_selected_rhythmus_label()
        self.selected_bereich_list.clear()

        # Uncheck all checkboxes
        for var in self.bereich_vars:
            var.set(False)

        self.urlaub_kw_entry.delete(0, tk.END)
        self.urlaub_day_entry.delete(0, tk.END)

    def add_to_selected_rhythmus(self, option):
        self.selected_rhythmus_list.append(option)
        self.update_selected_rhythmus_label()

    def update_selected_rhythmus_label(self):
        self.selected_rhythmus_display.config(text=", ".join(self.selected_rhythmus_list))

    # Reset list for "Rhythmus"
    def reset_selected_rhythmus_label(self):
        self.selected_rhythmus_display.config(text="")
        self.selected_rhythmus_list.clear()

    def run(self):
        self.window.mainloop()

    def create_calender_button(self):
        # Create a custom dialog to get all three values
        dialog = tk.Toplevel(self.window)
        dialog.title("Calendar Input")

        year_label = Label(dialog, text="Kalender Jahr:")
        year_label.pack()

        year_entry = Entry(dialog)
        year_entry.pack()
        year_entry.insert(0, "2024")

        start_week_label = Label(dialog, text="Start Arbeitswoche (1-52):")
        start_week_label.pack()

        start_week_entry = Entry(dialog)
        start_week_entry.pack()
        start_week_entry.insert(0, "1")

        end_week_label = Label(dialog, text="Letzte Arbeitswoche (1-52):")
        end_week_label.pack()

        end_week_entry = Entry(dialog)
        end_week_entry.pack()
        end_week_entry.insert(0, "52")

        # Function to handle the OK button click
        def ok_clicked():
            year = year_entry.get()
            start_week = start_week_entry.get()
            end_week = end_week_entry.get()

            try:
                year = int(year)
                start_week = int(start_week)
                end_week = int(end_week)

                if 1 <= start_week <= 52 and 1 <= end_week <= 52:
                    # Valid input, close the dialog and call create_calender
                    dialog.destroy()
                    create_calender(year, start_week, end_week)
                else:
                    messagebox.showerror("Invalid Input", "Please enter valid values.")
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter valid numeric values.")

        ok_button = Button(dialog, text="OK", command=ok_clicked)
        ok_button.pack()

        dialog.mainloop()

    def link_employees_window(self):
        # Get the index of the selected employee in the main listbox
        selected_index = self.employee_listbox.curselection()

        if not selected_index:
            return  # No employee selected

        selected_index = int(selected_index[0])
        selected_employee = self.data["employees"][selected_index]

        # Check if the selected employee has a link
        if EMPLOYEE_KEY[6] in selected_employee:
            # Ask for confirmation to remove the link
            result = messagebox.askquestion(
                "Bestätigen",
                f"{selected_employee[EMPLOYEE_KEY[0]]} hat bereits einen Link zu {selected_employee[EMPLOYEE_KEY[6]]}. Möchten Sie den Link entfernen?",
            )

            if result == "no":
                return  # Link removal canceled

            # Remove the link from both employees
            linked_employee_name = selected_employee[EMPLOYEE_KEY[6]]
            for employee in self.data["employees"]:
                if employee[EMPLOYEE_KEY[0]] == linked_employee_name:
                    del employee[EMPLOYEE_KEY[6]]
                    break  # Assuming there is only one link
            del selected_employee[EMPLOYEE_KEY[6]]

            # Save the updated data back to the file
            self.save_data_to_file()

            # Update the Listbox to reflect the change
            self.employee_listbox.delete(0, tk.END)
            self.load_employee_names_for_listbox()

            # Show a confirmation message
            messagebox.showinfo(
                "Link Entfernt",
                f"Der Link zwischen {selected_employee[EMPLOYEE_KEY[0]]} und {linked_employee_name} wurde entfernt.",
            )
        else:
            # Create a new window for linking employees
            link_window = tk.Toplevel(self.window)
            link_window.title("Link Employees")

            # Filter available employees excluding the selected one
            available_employees = [
                employee for index, employee in enumerate(self.data["employees"]) if index != selected_index
            ]

            # Filter available employees with the same rhythm and model
            matching_employees = [
                employee
                for employee in available_employees
                if (
                    employee[EMPLOYEE_KEY[1]] == selected_employee[EMPLOYEE_KEY[1]]
                    and set(employee[EMPLOYEE_KEY[2]]) == set(selected_employee[EMPLOYEE_KEY[2]])
                )
            ]

            if not matching_employees:
                messagebox.showerror("Error", "Kein Mitarbeiter mit demselben Rhythmus und Modell verfügbar.")
                link_window.destroy()
                return

            # Create a label to display the information
            info_label = tk.Label(
                link_window,
                text="Mitarbeiter mit selbem Rhythmus und selbem Modell:",
                font=("TkDefaultFont", 10, "bold"),
            )
            info_label.pack()

            # Create a listbox to display matching available employees
            available_listbox = tk.Listbox(link_window, height=5, width=30)
            available_listbox.pack()

            # Populate the listbox with matching available employees
            for employee in matching_employees:
                available_listbox.insert(tk.END, employee[EMPLOYEE_KEY[0]])

            # Function to link the selected employees
            def link_selected_employees():
                selected_available_index = available_listbox.curselection()
                if selected_available_index:
                    selected_available_index = int(selected_available_index[0])
                    selected_available_employee = matching_employees[selected_available_index]

                    # Update both employees' data with link information
                    selected_employee[EMPLOYEE_KEY[6]] = selected_available_employee[EMPLOYEE_KEY[0]]
                    selected_available_employee[EMPLOYEE_KEY[6]] = selected_employee[EMPLOYEE_KEY[0]]

                    # Save the updated data back to the file
                    self.save_data_to_file()

                    # Update the Listbox to reflect the change
                    self.employee_listbox.delete(0, tk.END)
                    self.load_employee_names_for_listbox()

                    # Close the link window
                    link_window.destroy()

                    # Show a confirmation message
                    messagebox.showinfo(
                        "Verlinkt",
                        f"{selected_employee[EMPLOYEE_KEY[0]]} und {selected_available_employee[EMPLOYEE_KEY[0]]} sind jetzt verlinkt.",
                    )

            link_button = tk.Button(link_window, text="Link Mitarbeiter", command=link_selected_employees)
            link_button.pack()


if __name__ == "__main__":
    app = GUIManager()
    app.run()
