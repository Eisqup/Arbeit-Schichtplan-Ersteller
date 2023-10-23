import tkinter as tk
from tkinter import messagebox
from constants import AREAS_DESCRIPTION


class EditAreas:
    def __init__(self, parent, areas, data_manager=None, save_callback=None):
        self.parent = parent
        self.areas = areas
        self.save_callback = save_callback
        self.data_manager = data_manager

        self.create_entries()

    def create_entries(self):
        # Add an information label at the top
        info_label = tk.Label(self.parent, text=AREAS_DESCRIPTION, anchor="w", justify="left")
        info_label.grid(row=0, column=0, columnspan=12, pady=10, sticky="w")
        self.entry_vars = {}
        row = 1
        for key, values in self.areas.items():
            # Create a label with "Name" before the entry
            name_label = tk.Label(self.parent, text="Name")
            name_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")

            entry_var = tk.StringVar()
            if key.startswith("-b"):
                entry_var.set("-b")
            else:
                entry_var.set(key)
            entry = tk.Entry(self.parent, textvariable=entry_var)
            entry.grid(row=row, column=1, padx=10, pady=5, sticky="w")
            self.entry_vars[key] = entry_var

            for i, (sub_key, sub_value) in enumerate(values.items()):
                if sub_key == "fill":
                    continue
                sub_label = tk.Label(self.parent, text=sub_key)
                sub_label.grid(row=row, column=i * 2 + 2, padx=(0, 5), pady=5, sticky="e")

                sub_entry_var = tk.StringVar()
                sub_entry_var.set(sub_value)
                sub_entry = tk.Entry(self.parent, textvariable=sub_entry_var)
                sub_entry.grid(row=row, column=i * 2 + 3, padx=10, pady=5, sticky="w")
                self.entry_vars[f"{key}_{sub_key}"] = sub_entry_var

            # Check if the last value in 'values' is a boolean (for a checkbox)
            last_value = list(values.values())[-1]
            if isinstance(last_value, bool):
                fill_check_var = tk.BooleanVar(value=last_value)
            else:
                fill_check_var = tk.BooleanVar(value=False)

            checkbox_label = tk.Label(self.parent, text="Überhang")
            checkbox_label.grid(row=row, column=9, sticky="e")
            fill_checkbutton = tk.Checkbutton(self.parent, variable=fill_check_var, onvalue=True, offvalue=False)
            fill_checkbutton.grid(row=row, column=10, sticky="w")
            self.entry_vars[f"{key}_fill"] = fill_check_var

            # Create a Delete button for each row
            delete_button = tk.Button(self.parent, text="X", command=lambda k=key: self.delete_area(k))
            delete_button.grid(row=row, column=11, padx=10, pady=5)
            row += 1

        # Create Save and Add Row buttons
        save_button = tk.Button(self.parent, text="Speicher Bereiche", command=self.save_data)
        save_button.grid(row=len(self.areas) + 1, column=1, pady=10)

        add_row_button = tk.Button(self.parent, text="Erstelle einen neuen Bereich", command=self.add_row)
        add_row_button.grid(row=len(self.areas) + 1, column=2, pady=10)

    def save_data(self, conform=True):
        areas = {}
        fill_set = False  # Flag to track if one fill value has been set to True
        unique_prio_values = set()
        unique_number_for_break = 1

        for key in self.areas.keys():
            # Extract values from the entry fields
            name_value = self.entry_vars[key].get()
            if name_value == "-b":
                name_value = "-b" + str(unique_number_for_break)
                unique_number_for_break += 1
            prio_value = self.entry_vars[f"{key}_prio"].get()
            min_value = self.entry_vars[f"{key}_min"].get()
            max_value = self.entry_vars[f"{key}_max"].get()
            fill_value = self.entry_vars[f"{key}_fill"].get()

            if conform:
                # Check data types and perform conversions/validations
                try:
                    prio_value = int(prio_value)
                    min_value = int(min_value)
                    max_value = int(max_value)
                except ValueError:
                    tk.messagebox.showerror("Error", f"Invalid integer value for {self.entry_vars[key].get()}")
                    return

                # Check if the prio_value is unique
                if prio_value in unique_prio_values:
                    duplicate_rows = [k for k, v in areas.items() if v["prio"] == prio_value]
                    tk.messagebox.showerror("Error", f"Duplicate 'prio' value detected for {self.entry_vars[key].get()} and {duplicate_rows}")
                    return
                else:
                    unique_prio_values.add(prio_value)

                # Check if fill_value is True, and if it is, set the fill_set flag
                if fill_value:
                    if fill_set:
                        # If fill_set is already True, show an error message
                        tk.messagebox.showerror("Error", "Only one 'fill' value can be True")
                        return
                    else:
                        fill_set = True

            # Create the dictionary entry
            areas[name_value] = {"prio": prio_value, "min": min_value, "max": max_value}
            if fill_value:
                areas[name_value]["fill"] = fill_value

        if conform:
            confirmation_message = "Alle Bereiche richtig?\n\n" + "\n".join([f"-b: {value}" if key.startswith("-b") else f"{key}: {value}" for key, value in areas.items()])
            if messagebox.askyesno("Confirmation", confirmation_message):
                self.areas = areas
                self.data_manager.save_settings_areas(self.areas)
                if self.save_callback:
                    self.save_callback()
            else:
                return
        else:
            self.areas = areas

    def add_row(self):
        # Save the current entries to self.areas
        self.save_data(conform=False)

        # Create a new empty row
        new_key = f"New Area {len(self.areas)+ 1}"
        c = 1
        while new_key in self.areas.keys():
            new_key = f"New Area {len(self.areas)+ 1 + c}"
            c += 1
        self.areas[new_key] = {"prio": "Enter a number", "min": "Enter a number", "max": "Enter a number"}
        # Clear the old widgets from the window
        for widget in self.parent.winfo_children():
            widget.grid_forget()
        self.create_entries()

    def delete_area(self, key):
        # Save the current entries to self.areas
        # self.save_data(conform=False)

        # Remove the specified row by its key
        if key in self.areas:
            del self.areas[key]

            # Clear the old widgets from the window
            for widget in self.parent.winfo_children():
                widget.grid_forget()
            self.save_data(conform=False)
            # Recreate the entries to update the view
            self.create_entries()


if __name__ == "__main__":
    BEREICHE = {
        "Drehen": {"prio": 1, "min": 3, "max": 3, "fill": True},
        "Bohren": {"prio": 3, "min": 2, "max": 3},
        "Fräsen": {"prio": 2, "min": 3, "max": 3},
    }

    root = tk.Tk()
    edit_window = tk.Toplevel(root)
    edit_window.title("Bereich bearbeiten")
    edit_areas = EditAreas(edit_window, areas=BEREICHE)
    root.mainloop()
