# Employee Creator GUI Application
# Version: 1.1
# Created by: Steven Stegink

import tkinter as tk
from tkinter import Label, Button, StringVar, Radiobutton, Checkbutton, ttk, messagebox, Entry, Frame
import re
import webbrowser

from constants import *

from data_manager import DataManager
from gui_edit_areas_dialog import EditAreas
from gui_calender_input_dialog import CalendarInputDialog


class GUIManager:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Employee Creator")
        self.button_size = 20

        self.data_manager = DataManager()
        self.data = self.data_manager.data
        self.settings_areas = self.data_manager.load_settings_areas()

        # Add a "Help" menu to your GUI's menu bar
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)

        # Create a "Help" menu item with an "About" option
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Informationen", command=self.show_program_infos)
        help_menu.add_command(label="About", command=self.show_about_info)

        # reader grid
        self.create_gui()
        self.show_program_infos()

    def show_program_infos(self):
        messagebox.showinfo(
            "Programm Informationen",
            PROGRAMM_INFOS,
        )

    def show_about_info(self):
        def open_github(event):
            webbrowser.open("https://github.com/Eisqup/Arbeit-Schichtplan-Ersteller")

        about_window = tk.Toplevel()
        about_window.title("About Employee Creator")
        about_label = tk.Label(about_window, text="Employee Calender Creator Application\nVersion: 1.1\nCreated by: Steven Stegink")
        about_label.pack(padx=20, pady=20)

        # Create a link label
        link_label = tk.Label(about_window, text="GitHub Repository", fg="blue", cursor="hand2")
        link_label.pack()
        link_label.bind("<Button-1>", open_github)

    def create_gui(self):
        # Create a dictionary to hold the Checkbutton variables
        self.selected_bereich_list = []
        self.bereich_vars = []

        # Create an empty list to store selected "Rhythmus" options
        self.selected_rhythmus_list = []

        # Label for titles on the left and right
        left_title_label = Label(self.window, text="Mitarbeiter Erstellen/Bearbeiten", font=("TkDefaultFont", 14, "bold underline"), justify="center", pady="20px")
        left_title_label.grid(row=0, column=0, sticky="w")

        right_title_label = Label(self.window, text="Vorhandene Mitarbeiter", font=("TkDefaultFont", 14, "bold underline"), justify="center")
        right_title_label.grid(row=0, column=4, sticky="w")

        # Label for entering the employee's name
        name_label = Label(self.window, text="Mitarbeiter Name:")
        name_label.grid(row=1, column=0, sticky="w")

        self.name_entry = tk.Entry(self.window)
        self.name_entry.grid(row=1, column=1)

        # Separator between sections
        separator1 = ttk.Separator(self.window, orient="horizontal")
        separator1.grid(row=2, columnspan=3, sticky="ew", pady=10)

        # Label for selecting shift model
        shift_model_label = Label(self.window, text="Schicht Modell:")
        shift_model_label.grid(row=3, column=0, sticky="w")

        # Radio buttons for selecting shift model
        self.shift_model_var = tk.IntVar()
        self.shift_model_var.set(1)  # Set an initial value to None
        for key, value in SCHICHT_MODELS.items():
            self.shift_model_radio = Radiobutton(self.window, text=value, variable=self.shift_model_var, value=key)
            self.shift_model_radio.grid(row=2 + key, column=1, sticky="w")
            if key == 2:
                # Bind the function to the second radio button
                self.shift_model_radio.bind("<ButtonRelease-1>", lambda event, i=key: self.on_model_2_radio_selected())
            else:
                # Bind the function to the first and third radio buttons
                self.shift_model_radio.bind("<ButtonRelease-1>", lambda event, i=key: self.on_other_model_radio_selected())
        for i, item in enumerate(SCHICHT_MODELS_BESCHREIBUNG):
            i += 1
            shift_description = Label(self.window, text=item)
            shift_description.grid(row=2 + i, column=2, sticky="w")

        # Separator between sections
        separator1 = ttk.Separator(self.window, orient="horizontal")
        separator1.grid(row=6, columnspan=3, sticky="ew", padx=10, pady=10)

        # Label for selecting Rhythmus
        rhythmus_label = Label(self.window, text="Schicht Rhythmus:")
        rhythmus_label.grid(row=7, column=0, padx=10, sticky="nw")

        # Create a frame to contain the buttons
        rhythmus_frame = Frame(self.window)

        # Buttons for selecting Rhythmus
        self.rhythmus_options = list(SCHICHT_RHYTHMUS.values())
        self.selected_rhythmus_var = StringVar()
        self.selected_rhythmus_var.set(None)  # Set an initial value
        for i, option in enumerate(self.rhythmus_options):
            self.rhythmus_button = Button(
                rhythmus_frame,
                text=option,
                width=self.button_size,
                command=lambda o=option: self.add_to_selected_rhythmus(o),
            )
            self.rhythmus_button.grid(row=i, column=0, sticky="w", pady=5)

        rhythmus_frame.grid(row=7, rowspan=len(SCHICHT_RHYTHMUS), column=1)

        # Frame für die beiden Labels
        label_frame = Frame(self.window)
        label_frame.grid(row=7, column=2, padx=10, pady=5, sticky="w", rowspan=3)

        # Label für "Ausgewählter Rhythmus" im Frame
        selected_rhythmus_label = Label(label_frame, text="Ausgewählter Rhythmus:")
        selected_rhythmus_label.pack(side="top", anchor="nw")

        # Label für die Anzeige des ausgewählten Rhythmus im Frame
        self.selected_rhythmus_display = Label(label_frame, text="")
        self.selected_rhythmus_display.pack(side="top", pady=10)

        # Clear button for Rhythmus
        clear_rhythmus_button = Button(
            label_frame,
            text="RESET Zyklus",
            width=self.button_size,
            command=self.reset_selected_rhythmus_label,
        )
        clear_rhythmus_button.pack(side="bottom", anchor="sw")

        # Separator between sections
        separator1 = ttk.Separator(self.window, orient="horizontal")
        separator1.grid(row=10, columnspan=3, sticky="ew", padx=10, pady=10)

        # Label for selecting Bereich
        bereich_label = Label(self.window, text="Mitarbeiter Bereichsfähigkeiten:")
        bereich_label.grid(row=11, column=0, padx=10, pady=5, sticky="w")

        # Create a frame to contain the Checkbuttons
        bereich_frame = Frame(self.window)

        # Checkbuttons for selecting Bereich
        if not self.settings_areas:
            no_bereich_label = Label(bereich_frame, text="Keine Bereiche vorhanden")
            no_bereich_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
            add_len = 1
        else:
            filtered_data = {key: value for key, value in self.settings_areas.items() if value.get("min", 0) > 0}
            add_len = len(filtered_data)
            for i, (key, value) in enumerate(filtered_data.items()):
                self.bereich_var = tk.BooleanVar()
                self.bereich_vars.append(self.bereich_var)
                self.bereich_checkbutton = Checkbutton(
                    bereich_frame,
                    text=key,
                    variable=self.bereich_var,
                    command=lambda o=key: self.selected_bereich_list.append(o),
                )
                self.bereich_checkbutton.grid(row=i, column=0, padx=10, pady=5, sticky="w")

        # Place the frame in the main window with the specified grid settings
        bereich_frame.grid(row=11, rowspan=add_len, column=1)

        # Create a button for "Bereich bearbeiten"
        bereich_bearbeiten_button = tk.Button(self.window, text="Bereiche bearbeiten", command=self.edit_areas)  # Define a callback function for the button
        bereich_bearbeiten_button.grid(row=11, column=2, sticky="w")

        # Separator between sections
        separator1 = ttk.Separator(self.window, orient="horizontal")
        separator1.grid(row=11 + add_len, columnspan=3, sticky="ew", padx=10, pady=10)

        # Label for entering urlaub_kw
        urlaub_kw_label = Label(self.window, text="Mitarbeiter Urlaub in KW (sepertatet with space):")
        urlaub_kw_label.grid(row=12 + add_len, column=0, padx=10, pady=5, sticky="w")

        self.urlaub_kw_entry = tk.Entry(self.window)
        self.urlaub_kw_entry.grid(row=12 + add_len, column=1, padx=10, pady=5)

        # # Label for entering urlaub_day
        # urlaub_day_label = Label(self.window, text="Enter Urlaub (Tage)(sepertatet with space):\nBespiel: DD.MM DD.MM")
        # urlaub_day_label.grid(row=13 + add_len, column=0, padx=10, pady=5, sticky="w")

        # self.urlaub_day_entry = tk.Entry(self.window)
        # self.urlaub_day_entry.grid(row=13 + add_len, column=1, padx=10, pady=5)

        # Submit button
        submit_button = tk.Button(self.window, text="Daten Speichern", width=self.button_size, command=self.save_employee_data)
        submit_button.grid(row=14 + add_len, columnspan=2, pady=10)

        # reset all
        reset_button = tk.Button(self.window, text="RESET OVERLAY", width=self.button_size, command=self.reset_form_fields)
        reset_button.grid(row=14 + add_len, column=1, columnspan=2, pady=10)

        # Vertical Separator in column 3
        separator1_vertical = ttk.Separator(self.window, orient="vertical")
        separator1_vertical.grid(row=0, column=3, rowspan=100, sticky="ns")

        # Listbox widget to display employee names
        self.employee_listbox = tk.Listbox(self.window, height=5, width=30)
        self.employee_listbox.grid(row=1, column=4, rowspan=12, sticky="nsew")

        # Link employee
        delete_button = tk.Button(
            self.window,
            text="Hinzufügen/Entfernen\nMitarbeiter link",
            width=self.button_size,
            command=self.link_employees_window,
        )
        delete_button.grid(row=14, column=4, columnspan=2, sticky="w")

        # Delete button
        delete_button = tk.Button(self.window, text="Mitarbeiter löschen", width=self.button_size, command=self.delete_employee)
        delete_button.grid(row=15, column=4, columnspan=2, sticky="w")

        # Create a Load button
        load_button = tk.Button(self.window, text="Mitarbeiter Daten Laden", width=self.button_size, command=self.load_selected_employee_to_GUI)
        load_button.grid(row=16, column=4, columnspan=2, sticky="w")

        # Create a Shift plan button
        load_button = tk.Button(self.window, text="Erstelle Schichtplan", width=self.button_size, command=self.btn_create_calender)
        load_button.grid(row=17, column=4, columnspan=2, sticky="w")

        # Create a frame to display the sneak peek
        self.sneak_peek_frame = tk.Frame(self.window)
        self.sneak_peek_frame.grid(row=1, column=6, rowspan=5, sticky="nsew")

        # Label to display selected employee's info in the sneak peek
        self.sneak_peek_info_label = Label(self.sneak_peek_frame, text="", justify="left")
        self.sneak_peek_info_label.pack()

        # Bind the Listbox selection event to update the sneak peek
        self.employee_listbox.bind("<<ListboxSelect>>", self.update_sneak_peek)

        # Populate the list of employee names
        self.load_employee_names_in_listbox()

    def on_model_2_radio_selected(self, default_value_for_radiobutton="default"):
        if hasattr(self, "radiobutton_frame"):
            self.radiobutton_frame.destroy()

        # Create a frame for Checkbuttons
        self.radiobutton_frame = Frame(self.window)

        label_only_model_2 = Label(
            self.radiobutton_frame,
            text="\nBestimme in welcher Woche\nder ausgewählte Rhythmus startet.\n(Default=Automatisch via Algorithmus)",
            justify="left",
        )
        label_only_model_2.pack(side="top", anchor="w")

        # Create a variable to track the selected option
        self.selected_option_start_kw_model2 = tk.StringVar()
        self.selected_option_start_kw_model2.set(default_value_for_radiobutton)  # Set an initial value

        # Create the radio buttons
        default_radio = Radiobutton(self.radiobutton_frame, text="Default", variable=self.selected_option_start_kw_model2, value="default")
        kw1_start_radio = Radiobutton(self.radiobutton_frame, text="KW1", variable=self.selected_option_start_kw_model2, value="kw1_start")
        arbeitswoche_start_radio = Radiobutton(self.radiobutton_frame, text="Arbeitswoche start", variable=self.selected_option_start_kw_model2, value="arbeitswoche_start")

        default_radio.pack(side="left", padx=5)
        kw1_start_radio.pack(side="left", padx=5)
        arbeitswoche_start_radio.pack(side="left", padx=5)

        self.radiobutton_frame.grid(row=8, column=0)

    def on_other_model_radio_selected(self):
        # Delete the checkbutton frame if it exists
        if hasattr(self, "radiobutton_frame"):
            self.radiobutton_frame.destroy()

    def save_employee_data(self):
        # Function to split the input into a list
        def split_and_check_input(input_str, is_weeks=False):
            values = input_str.split()
            if is_weeks:
                # Validate weeks (1-52)
                for value in values:
                    try:
                        week = int(value)
                        if not (1 <= week <= 52):
                            raise ValueError("Week must be between 1 and 52")
                    except ValueError:
                        messagebox.showerror("Error", f"Invalid week value: {value}. Weeks must be integers between 1 and 52.")
                        return None  # Exit and handle the error

            else:
                # Validate dates (DD.MM)
                date_pattern = re.compile(r"^(0[1-9]|[12][0-9]|3[01])\.(0[1-9]|1[0-2])$")
                for value in values:
                    if not date_pattern.match(value):
                        messagebox.showerror("Error", f"Invalid date format: {value}. Dates must be in DD.MM format.")
                        return None  # Exit and handle the error

            return values

        name = self.name_entry.get()
        selected_shift_model = self.shift_model_var.get()
        urlaub_kw = split_and_check_input(self.urlaub_kw_entry.get(), is_weeks=True)
        # urlaub_day = split_and_check_input(self.urlaub_day_entry.get())

        filtered_data = [key for key, value in self.settings_areas.items() if value.get("min", 0) > 0]
        self.selected_bereich_list = [filtered_data[index] for index, var in enumerate(self.bereich_vars) if var.get()]

        # if urlaub_kw is None or urlaub_day is None:
        if urlaub_kw is None:
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
                    emp = employee
                    break

            employee_data = {
                EMPLOYEE_KEY[0]: name,
                EMPLOYEE_KEY[1]: SCHICHT_MODELS[selected_shift_model],
                EMPLOYEE_KEY[2]: self.selected_rhythmus_list,
                EMPLOYEE_KEY[3]: self.selected_bereich_list,
                EMPLOYEE_KEY[4]: urlaub_kw,
                # EMPLOYEE_KEY[5]: urlaub_day,
            }

            if SCHICHT_MODELS[selected_shift_model] == SCHICHT_MODELS[2]:
                model_2_checkbutton_start_kw = self.selected_option_start_kw_model2.get()
                employee_data["start_kw_model_2"] = model_2_checkbutton_start_kw

            if employee_exists:
                if "link" in emp:
                    employee_data["link"] = emp["link"]

                result = messagebox.askquestion("Confirm", "Mitarbeiter schon vorhanden. Daten überschreiben?")
                if result == "yes":
                    # Overwrite the existing employee's data
                    for index, employee in enumerate(self.data.get("employees", [])):
                        if employee.get(EMPLOYEE_KEY[0]) == name:
                            self.data["employees"][index] = {}
                            # Add the new employee data
                            self.data["employees"][index].update(employee_data)
            else:
                self.data["employees"].append(employee_data)
                self.data = self.data_manager.sort_data(self.data)

            # Save the updated data back to the file
            self.data_manager.save_data(self.data)

            if not employee_exists:
                message = "Mitarbeiter hinzugefügt:\n"
                for key, value in employee_data.items():
                    message += f"{key}: {value}\n"

                messagebox.showinfo(
                    "Info",
                    f"{message}",
                )
            else:
                if result == "yes":
                    message = "Mitarbeiterdaten aktualisiert\n"
                    messagebox.showinfo(
                        "Info",
                        f"{message}",
                    )
                else:
                    message = "Abbruch"

            print(message)

        # Insert the newly created/updated employee's name into the Listbox
        self.employee_listbox.delete(0, tk.END)
        self.load_employee_names_in_listbox()

        # Clear the form fields after creating or updating an employee
        self.reset_form_fields()
        self.data = self.data_manager.load_data()

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
                    self.data_manager.save_data(self.data)

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
                    filtered_data = [key for key, value in self.settings_areas.items() if value.get("min", 0) > 0]
                    index = filtered_data.index(bereich)
                    # index = list(self.settings_areas.keys()).index(bereich)
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
                # self.urlaub_day_entry.delete(0, tk.END)
                # self.urlaub_day_entry.insert(0, " ".join(selected_employee[EMPLOYEE_KEY[5]]))

                if shift_model == SCHICHT_MODELS[2]:
                    if "start_kw_model_2" in selected_employee:
                        self.on_model_2_radio_selected(selected_employee["start_kw_model_2"])
                    else:
                        self.on_model_2_radio_selected()
                else:
                    self.on_other_model_radio_selected()

    def load_employee_names_in_listbox(self):
        employee_names = []
        if "employees" in self.data:
            employee_names = [employee[EMPLOYEE_KEY[0]] for employee in self.data["employees"]]
        for name in employee_names:
            self.employee_listbox.insert(tk.END, name)

    def edit_areas(self):
        def on_data_saved():
            # This function will be called when data is successfully saved
            edit_window.destroy()
            self.settings_areas = self.data_manager.load_settings_areas()
            for widget in self.window.grid_slaves():
                widget.grid_forget()
            self.create_gui()

        edit_window = tk.Toplevel(self.window)
        edit_window.title("Bereich bearbeiten")
        edit_areas = EditAreas(edit_window, areas=self.settings_areas, data_manager=self.data_manager, save_callback=on_data_saved)

    def reset_form_fields(self):
        self.name_entry.delete(0, tk.END)
        self.shift_model_var.set(1)
        self.reset_selected_rhythmus_label()
        self.selected_bereich_list.clear()
        self.on_other_model_radio_selected()

        # Uncheck all checkboxes
        for var in self.bereich_vars:
            var.set(False)

        self.urlaub_kw_entry.delete(0, tk.END)
        # self.urlaub_day_entry.delete(0, tk.END)

    def add_to_selected_rhythmus(self, option):
        self.selected_rhythmus_list.append(option)
        self.selected_rhythmus_display.config(text=", ".join(self.selected_rhythmus_list))

    # Reset list for "Rhythmus"
    def reset_selected_rhythmus_label(self):
        self.selected_rhythmus_display.config(text="")
        self.selected_rhythmus_list.clear()

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
                    if EMPLOYEE_KEY[6] in employee:
                        del employee[EMPLOYEE_KEY[6]]
                        break  # Assuming there is only one link
            if EMPLOYEE_KEY[6] in selected_employee:
                del selected_employee[EMPLOYEE_KEY[6]]

            # Save the updated data back to the file
            self.data_manager.save_data(self.data)

            # Update the Listbox to reflect the change
            self.employee_listbox.delete(0, tk.END)
            self.load_employee_names_in_listbox()

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
            available_employees = [employee for index, employee in enumerate(self.data["employees"]) if index != selected_index]

            # Filter available employees with the same rhythm and model
            matching_employees = [employee for employee in available_employees if (employee[EMPLOYEE_KEY[1]] == selected_employee[EMPLOYEE_KEY[1]] and set(employee[EMPLOYEE_KEY[2]]) == set(selected_employee[EMPLOYEE_KEY[2]]))]

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
                    self.data_manager.save_data(self.data)

                    # Update the Listbox to reflect the change
                    self.employee_listbox.delete(0, tk.END)
                    self.load_employee_names_in_listbox()

                    # Close the link window
                    link_window.destroy()

                    # Show a confirmation message
                    messagebox.showinfo(
                        "Verlinkt",
                        f"{selected_employee[EMPLOYEE_KEY[0]]} und {selected_available_employee[EMPLOYEE_KEY[0]]} sind jetzt verlinkt.",
                    )

            link_button = tk.Button(link_window, text="Link Mitarbeiter", command=link_selected_employees)
            link_button.pack()

    def btn_create_calender(self):
        # check if data there to run teh program
        if not self.settings_areas:
            messagebox.showerror("Missing Data", "Please insert areas and make sure employees have the areas as ability.")
            return

        if len(self.data["employees"]) == 0:
            print("Erstelle nur Excel Datei.")

        # Create the calendar input dialog
        calendar_dialog = CalendarInputDialog(self.window, self.settings_areas)

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = GUIManager()
    app.run()
