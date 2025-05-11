import tkinter as tk
from tkinter import ttk, messagebox
from export_database import export_database_to_excel
from report import show_report
from tooltip import ToolTip


class WorkplaceManager:
    def __init__(self, parent, json_file):
        self.parent = parent
        self.root = self.parent
        self.json_file = json_file
        self.workplace_data = []
        self.filtered_workplace_data = None

        # Configure ttk style for buttons
        style = ttk.Style()
        style.configure(
            "Report.TButton",
            font=("Helvetica", 12, "bold"),
            foreground="black",
            background="#4682b4",
        )
        style.configure(
            "Close.TButton",
            font=("Helvetica", 12, "bold"),
            foreground="black",
            background="#d9534f",
        )
        style.map("Close.TButton", background=[("active", "#e57373")])
        style.configure("Action.TButton", font=("Helvetica", 10))

        # Workplace Frame
        self.workplace_frame = ttk.Frame(
            self.parent, relief="groove", borderwidth=2, padding=10
        )
        self.workplace_frame.grid(row=0, column=1, sticky="nsew")
        self.workplace_frame.columnconfigure(0, weight=1)
        self.workplace_frame.rowconfigure(2, weight=1)

        # Frame for Report and Close buttons
        btn_report_frame = ttk.Frame(self.workplace_frame)
        btn_report_frame.grid(row=0, column=0, sticky="ew")
        btn_report_frame.columnconfigure(0, weight=1)

        self.btn_report = ttk.Button(
            btn_report_frame,
            text="Report",
            command=self.show_report,
            style="Report.TButton",
        )
        self.btn_report.grid(row=0, column=1, sticky="e", padx=(0, 5))
        ToolTip(self.btn_report, "Generate a report of workplace tests")

        self.btn_close_main = ttk.Button(
            btn_report_frame,
            text="Close",
            command=self.close_application,
            style="Close.TButton",
        )
        self.btn_close_main.grid(row=0, column=2, sticky="e")
        ToolTip(self.btn_close_main, "Close the application")

        self.workplace_title = ttk.Label(
            self.workplace_frame, text="Workplace", font=("Helvetica", 14, "bold")
        )
        self.workplace_title.grid(row=0, column=0, sticky="w", pady=(0, 10))

        # Temperature and Limiter Filters
        filter_temp_frame = ttk.Frame(self.workplace_frame)
        filter_temp_frame.grid(row=1, column=0, sticky="w", pady=5)

        # Temperature Filter (Checkbuttons)
        ttk.Label(filter_temp_frame, text="Filter Temperature:").pack(
            side="left", padx=(0, 5)
        )
        self.temp_all_var = tk.BooleanVar(value=True)
        self.temp_rt_var = tk.BooleanVar(value=True)
        self.temp_lt_var = tk.BooleanVar(value=True)
        self.temp_ht_var = tk.BooleanVar(value=True)
        self.selected_temperatures = ["RT", "LT", "HT"]
        self.temp_all_chk = ttk.Checkbutton(
            filter_temp_frame,
            text="All",
            variable=self.temp_all_var,
            command=self.update_temp_checkbuttons,
        )
        self.temp_all_chk.pack(side="left", padx=2)
        ToolTip(self.temp_all_chk, "Select all temperatures")
        self.temp_rt_chk = ttk.Checkbutton(
            filter_temp_frame, text="RT", variable=self.temp_rt_var, state="disabled"
        )
        self.temp_rt_chk.pack(side="left", padx=2)
        ToolTip(self.temp_rt_chk, "Select Room Temperature (RT)")
        self.temp_lt_chk = ttk.Checkbutton(
            filter_temp_frame, text="LT", variable=self.temp_lt_var, state="disabled"
        )
        self.temp_lt_chk.pack(side="left", padx=2)
        ToolTip(self.temp_lt_chk, "Select Low Temperature (LT)")
        self.temp_ht_chk = ttk.Checkbutton(
            filter_temp_frame, text="HT", variable=self.temp_ht_var, state="disabled"
        )
        self.temp_ht_chk.pack(side="left", padx=2)
        ToolTip(self.temp_ht_chk, "Select High Temperature (HT)")

        # Limiter Filter
        ttk.Label(filter_temp_frame, text="Limiter:").pack(side="left", padx=(10, 5))
        self.limit_var = tk.StringVar()
        limit_values = ["All"] + [str(i) for i in range(0, 201, 1)]
        self.limit_combobox = ttk.Combobox(
            filter_temp_frame,
            textvariable=self.limit_var,
            state="readonly",
            values=limit_values,
            width=8,
        )
        self.limit_combobox.pack(side="left")
        self.limit_combobox.set("All")
        ToolTip(self.limit_combobox, "Limit number of displayed tests")

        self.btn_apply_filters = ttk.Button(
            filter_temp_frame,
            text="Apply Filters",
            command=self.apply_filters,
            style="Action.TButton",
        )
        self.btn_apply_filters.pack(side="left", padx=10)
        ToolTip(self.btn_apply_filters, "Apply temperature and limiter filters")

        # Results List with Scrollbars
        self.frame_results = ttk.Frame(self.workplace_frame)
        self.frame_results.grid(row=2, column=0, sticky="nsew", pady=5)
        self.frame_results.columnconfigure(0, weight=1)
        self.frame_results.rowconfigure(0, weight=1)

        separator = ttk.Separator(self.workplace_frame, orient="horizontal")
        separator.grid(row=3, column=0, sticky="ew", pady=(5, 5))

        self.counter_frame = ttk.Frame(self.workplace_frame, padding=(0, 8))
        self.counter_frame.grid(row=4, column=0, sticky="ew")
        self.counter_frame.columnconfigure((0, 1, 2, 3), weight=1)

        self.label_rt = ttk.Label(
            self.counter_frame,
            text="RT: 0",
            font=("Helvetica", 11, "bold"),
            foreground="#008800",
        )
        self.label_rt.grid(row=0, column=0, sticky="ew", padx=8)

        self.label_lt = ttk.Label(
            self.counter_frame,
            text="LT: 0",
            font=("Helvetica", 11, "bold"),
            foreground="#0055cc",
        )
        self.label_lt.grid(row=0, column=1, sticky="ew", padx=8)

        self.label_ht = ttk.Label(
            self.counter_frame,
            text="HT: 0",
            font=("Helvetica", 11, "bold"),
            foreground="#cc5500",
        )
        self.label_ht.grid(row=0, column=2, sticky="ew", padx=8)

        self.label_total = ttk.Label(
            self.counter_frame,
            text="Total: 0",
            font=("Helvetica", 11, "bold"),
            foreground="#222222",
        )
        self.label_total.grid(row=0, column=3, sticky="ew", padx=8)

        self.scrollbar_y = ttk.Scrollbar(self.frame_results, orient=tk.VERTICAL)
        self.scrollbar_x = ttk.Scrollbar(self.frame_results, orient=tk.HORIZONTAL)

        self.list_results = tk.Listbox(  # tk.Listbox remains
            self.frame_results,
            width=100,
            height=25,
            yscrollcommand=self.scrollbar_y.set,
            xscrollcommand=self.scrollbar_x.set,
        )
        self.list_results.grid(row=0, column=0, sticky="nsew")

        self.scrollbar_y.configure(command=self.list_results.yview)
        self.scrollbar_y.grid(row=0, column=1, sticky="ns")

        self.scrollbar_x.configure(command=self.list_results.xview)
        self.scrollbar_x.grid(row=1, column=0, sticky="ew")

    def update_temp_checkbuttons(self):
        """Update the state of temperature checkbuttons based on 'All' selection."""
        state = "disabled" if self.temp_all_var.get() else "normal"
        self.temp_rt_chk.configure(state=state)
        self.temp_lt_chk.configure(state=state)
        self.temp_ht_chk.configure(state=state)
        if self.temp_all_var.get():
            self.temp_rt_var.set(True)
            self.temp_lt_var.set(True)
            self.temp_ht_var.set(True)
            self.selected_temperatures = ["RT", "LT", "HT"]
        else:
            self.selected_temperatures = [
                temp
                for temp, var in [
                    ("RT", self.temp_rt_var),
                    ("LT", self.temp_lt_var),
                    ("HT", self.temp_ht_var),
                ]
                if var.get()
            ]

    def apply_filters(self):
        """Apply temperature and limiter filters to workplace data."""
        limit_filter = self.limit_var.get()

        # Determine selected temperatures
        if self.temp_all_var.get():
            self.selected_temperatures = ["RT", "LT", "HT"]
        else:
            self.selected_temperatures = [
                temp
                for temp, var in [
                    ("RT", self.temp_rt_var),
                    ("LT", self.temp_lt_var),
                    ("HT", self.temp_ht_var),
                ]
                if var.get()
            ]

        self.list_results.delete(0, tk.END)
        header = "Test | Inflator | Temperature | Type | Version | Order | Date"
        self.list_results.insert(tk.END, header)
        self.list_results.insert(tk.END, "-" * len(header))

        # Check for mixed versions
        versions = {reg["version"] for reg in self.workplace_data}
        if len(versions) > 1:
            messagebox.showerror(
                "Error",
                "Workplace contains mixed versions! Clear before applying filters.",
            )
            return

        # If no temperatures selected, show warning and clear display
        if not self.selected_temperatures and not self.temp_all_var.get():
            messagebox.showwarning(
                "Warning", "No temperatures selected. Please select at least one."
            )
            self.filtered_workplace_data = []
            self.update_workplace_counters([])
            return

        # Apply filters
        filtered_data = []
        for reg in self.workplace_data:
            if reg["type"] not in self.selected_temperatures:
                continue
            filtered_data.append(reg)

        # Apply limiter
        if limit_filter != "All":
            limit = int(limit_filter)
            filtered_data = filtered_data[:limit]

        # Update display
        for reg in filtered_data:
            line = f"{reg['test_no']} | {reg['inflator_no']} | {reg['temperature_c']}°C | {reg['type']} | {reg['version']} | {reg['order']} | {reg.get('test_date', 'N/A')}"
            if reg["pressures"]:
                line += " | Pressure data available"
            else:
                line += " | No pressure data"
            self.list_results.insert(tk.END, line)

        self.filtered_workplace_data = filtered_data
        self.update_workplace_counters(filtered_data)

    def export_database_to_excel(self):
        export_database_to_excel(self)

    def show_report(self):
        show_report(self)

    def close_application(self):
        try:
            self.parent.destroy()
            import sys

            sys.exit(0)
        except Exception as e:
            print(f"Error closing application: {str(e)}")
            raise

    def update_workplace_display(self):
        """Update the workplace listbox display."""
        self.list_results.delete(0, tk.END)
        header = "Test | Inflator | Temperature | Type | Version | Order | Date"
        self.list_results.insert(tk.END, header)
        self.list_results.insert(tk.END, "-" * len(header))

        for reg in self.workplace_data:
            line = f"{reg['test_no']} | {reg['inflator_no']} | {reg['temperature_c']}°C | {reg['type']} | {reg['version']} | {reg['order']} | {reg.get('test_date', 'N/A')}"
            if reg["pressures"]:
                line += " | Pressure data available"
            else:
                line += " | No pressure data"
            self.list_results.insert(tk.END, line)
        self.update_workplace_counters()

    def update_workplace_counters(self, data=None):
        """Update the workplace counters."""
        if data is None:
            data = self.workplace_data
        rt = lt = ht = 0
        for item in data:
            temp = item.get("type")
            if temp == "RT":
                rt += 1
            elif temp == "LT":
                lt += 1
            elif temp == "HT":
                ht += 1
        total = rt + lt + ht
        self.label_rt.configure(text=f"RT: {rt}")
        self.label_lt.configure(text=f"LT: {lt}")
        self.label_ht.configure(text=f"HT: {ht}")
        self.label_total.configure(text=f"Total: {total}")

    def clear_workplace(self):
        """Clear all data from the workplace."""
        self.workplace_data.clear()
        self.list_results.delete(0, tk.END)
        messagebox.showinfo("Success", "Workplace cleared successfully.")
        self.update_workplace_counters()
