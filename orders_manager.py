import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import database
from tooltip import ToolTip


class OrdersManager:
    def __init__(self, parent, json_file, excel_folder, workplace_manager):
        self.parent = parent
        self.json_file = json_file
        self.excel_folder = excel_folder
        self.workplace_manager = workplace_manager

        # Configure ttk style for buttons
        style = ttk.Style()
        style.configure("Action.TButton", font=("Helvetica", 10, "bold"))
        style.configure(
            "Send.TButton", font=("Helvetica", 10, "bold"), background="#90EE90"
        )
        style.configure(
            "Remove.TButton", font=("Helvetica", 10, "bold"), background="#FF6B6B"
        )
        style.configure(
            "Clear.TButton", font=("Helvetica", 10, "bold"), background="#D3D3D3"
        )
        style.configure("Filters.TFrame", background="#f0f0f0")
        style.configure("Filters.TLabel", background="#f0f0f0")

        # Orders Manager Frame
        self.frame_orders_manager = ttk.Frame(
            self.parent, relief="groove", borderwidth=2, padding=10
        )
        self.frame_orders_manager.grid(row=5, column=0, sticky="nsew")
        self.frame_orders_manager.columnconfigure(0, weight=1)
        self.frame_orders_manager.rowconfigure(3, weight=1)

        # Initialize Pagination Variables
        self.current_page = 1
        self.orders_per_page = 10
        self.total_pages = 1

        # Orders Manager Title
        title_orders_frame = ttk.Frame(self.frame_orders_manager)
        title_orders_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        title_orders_frame.columnconfigure(0, weight=1)
        ttk.Label(
            title_orders_frame, text="Orders Manager", font=("Helvetica", 12, "bold")
        ).grid(row=0, column=0, sticky="w")

        # Filters Frame
        filters_frame = ttk.Frame(
            self.frame_orders_manager,
            relief="groove",
            borderwidth=2,
            padding=10,
            style="Filters.TFrame",
        )
        filters_frame.grid(row=1, column=0, sticky="ew", pady=(0, 5))
        filters_frame.columnconfigure((0, 1, 2, 3), weight=1)

        # Version Filter
        ttk.Label(filters_frame, text="Version:", style="Filters.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )
        self.version_var = tk.StringVar()
        self.version_combobox = ttk.Combobox(
            filters_frame, textvariable=self.version_var, state="readonly", width=10
        )
        self.version_combobox.grid(row=0, column=1, sticky="w")
        self.version_combobox.bind("<<ComboboxSelected>>", self.on_version_filter)
        self.version_combobox["values"] = []
        self.version_combobox.set("All")
        ToolTip(self.version_combobox, "Filter orders by version")

        # Date Range Filter
        ttk.Label(filters_frame, text="Date Range:", style="Filters.TLabel").grid(
            row=0, column=2, sticky="w", padx=(10, 5)
        )
        self.date_range_var = tk.StringVar()
        self.date_range_combobox = ttk.Combobox(
            filters_frame,
            textvariable=self.date_range_var,
            state="readonly",
            values=["All", "Last 30 Days", "Last 60 Days", "Last 90 Days", "Custom"],
            width=12,
        )
        self.date_range_combobox.grid(row=0, column=3, sticky="w")
        self.date_range_combobox.set("All")
        self.date_range_combobox.bind("<<ComboboxSelected>>", self.on_date_range_filter)
        ToolTip(self.date_range_combobox, "Filter orders by date range")

        # Custom Date Inputs
        self.custom_dates_frame = ttk.Frame(filters_frame, style="Filters.TFrame")
        self.custom_dates_frame.grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(5, 0)
        )
        ttk.Label(self.custom_dates_frame, text="Start:", style="Filters.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )
        self.start_date_entry = ttk.Entry(self.custom_dates_frame, width=12)
        self.start_date_entry.grid(row=0, column=1, sticky="w")
        self.start_date_btn = ttk.Button(
            self.custom_dates_frame,
            text="📅",
            command=lambda: self.pick_date(self.start_date_entry),
            width=2,
            style="Action.TButton",
        )
        self.start_date_btn.grid(row=0, column=2, sticky="w", padx=2)
        ttk.Label(self.custom_dates_frame, text="End:", style="Filters.TLabel").grid(
            row=0, column=3, sticky="w", padx=(10, 5)
        )
        self.end_date_entry = ttk.Entry(self.custom_dates_frame, width=12)
        self.end_date_entry.grid(row=0, column=4, sticky="w")
        self.end_date_btn = ttk.Button(
            self.custom_dates_frame,
            text="📅",
            command=lambda: self.pick_date(self.end_date_entry),
            width=2,
            style="Action.TButton",
        )
        self.end_date_btn.grid(row=0, column=5, sticky="w", padx=2)
        self.custom_dates_frame.grid_remove()

        # Pagination Frame
        pagination_frame = ttk.Frame(self.frame_orders_manager)
        pagination_frame.grid(row=2, column=0, sticky="ew", pady=5)
        pagination_frame.columnconfigure(2, weight=1)

        ttk.Label(pagination_frame, text="Items per page:").grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )
        self.page_selector = ttk.Combobox(
            pagination_frame, values=[5, 10, 15, 20, 25, 30], width=5, state="readonly"
        )
        self.page_selector.set(self.orders_per_page)
        self.page_selector.grid(row=0, column=1, sticky="w")
        self.page_selector.bind("<<ComboboxSelected>>", self.on_items_per_page_changed)
        ToolTip(self.page_selector, "Select number of orders per page")

        self.select_all_var = tk.BooleanVar()
        self.select_all_chk = ttk.Checkbutton(
            pagination_frame,
            text="Select All",
            variable=self.select_all_var,
            command=self.toggle_select_all,
        )
        self.select_all_chk.grid(row=0, column=2, sticky="w", padx=10)
        ToolTip(self.select_all_chk, "Select all orders on current page")

        # Navigation Frame for Prev, Page Info, and Next
        self.nav_frame = ttk.Frame(pagination_frame)
        self.nav_frame.grid(row=0, column=3, sticky="e")
        self.prev_btn = ttk.Button(
            self.nav_frame,
            text="< Prev",
            command=lambda: self.change_page(-1),
            width=8,
            style="Action.TButton",
        )
        self.prev_btn.pack(side="left", padx=(0, 5))
        self.page_info = ttk.Label(
            self.nav_frame, text="Page 1/1", width=10, anchor="center"
        )
        self.page_info.pack(side="left", padx=(5, 5))
        self.next_btn = ttk.Button(
            self.nav_frame,
            text="Next >",
            command=lambda: self.change_page(1),
            width=8,
            style="Action.TButton",
        )
        self.next_btn.pack(side="left", padx=(5, 0))

        # Canvas and Scrollbar for Orders
        self.orders_canvas = tk.Canvas(self.frame_orders_manager)
        self.orders_canvas.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)
        self.orders_scrollbar = ttk.Scrollbar(
            self.frame_orders_manager,
            orient=tk.VERTICAL,
            command=self.orders_canvas.yview,
        )
        self.orders_scrollbar.grid(row=3, column=1, sticky="ns", pady=5)
        self.orders_canvas.configure(yscrollcommand=self.orders_scrollbar.set)

        self.orders_inner_frame = ttk.Frame(self.orders_canvas)
        self.orders_canvas.create_window(
            (0, 0), window=self.orders_inner_frame, anchor="nw"
        )

        self.orders_inner_frame.bind(
            "<Configure>",
            lambda e: self.orders_canvas.configure(
                scrollregion=self.orders_canvas.bbox("all")
            ),
        )
        self.orders_canvas.yview_moveto(0)
        self.orders_canvas.bind(
            "<Enter>",
            lambda e: self.orders_canvas.bind_all("<MouseWheel>", self._on_mousewheel),
        )
        self.orders_canvas.bind(
            "<Leave>", lambda e: self.orders_canvas.unbind_all("<MouseWheel>")
        )

        self.order_vars = {}
        self.order_checkbuttons = {}

        # Action Buttons Frame
        action_btn_frame = ttk.Frame(
            self.frame_orders_manager, relief="groove", borderwidth=2, padding=10
        )
        action_btn_frame.grid(row=4, column=0, sticky="ew", pady=5)
        action_btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_send_workplace = ttk.Button(
            action_btn_frame,
            text="Send to Workplace",
            command=self.send_to_workplace,
            style="Send.TButton",
            width=15,
        )
        self.btn_send_workplace.grid(row=0, column=0, sticky="ew", padx=5)
        ToolTip(self.btn_send_workplace, "Send selected orders to workplace")

        self.btn_remove_workplace_orders = ttk.Button(
            action_btn_frame,
            text="Remove Tests",
            command=self.remove_workplace_orders_selected,
            style="Remove.TButton",
            width=15,
        )
        self.btn_remove_workplace_orders.grid(row=0, column=1, sticky="ew", padx=5)
        ToolTip(
            self.btn_remove_workplace_orders, "Remove selected orders from workplace"
        )

        self.btn_clear_workplace = ttk.Button(
            action_btn_frame,
            text="Clear Workplace",
            command=self.clear_workplace,
            style="Clear.TButton",
            width=15,
        )
        self.btn_clear_workplace.grid(row=0, column=2, sticky="ew", padx=5)
        ToolTip(self.btn_clear_workplace, "Clear all tests from workplace")

    def _on_mousewheel(self, event):
        self.orders_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def pick_date(self, entry):
        """Open a DateEntry calendar to pick a date and insert it into the entry."""
        popup = tk.Toplevel(self.parent)
        popup.title("Select Date")
        popup.geometry("250x250")
        popup.resizable(False, False)
        date_picker = DateEntry(popup, date_pattern="yyyy-mm-dd", width=12)
        date_picker.pack(pady=10)
        ttk.Button(
            popup,
            text="Confirm",
            command=lambda: [
                entry.delete(0, tk.END),
                entry.insert(0, date_picker.get()),
                popup.destroy(),
                self.update_orders_list(),
            ],
            style="Action.TButton",
        ).pack(pady=5)
        ttk.Button(
            popup, text="Cancel", command=popup.destroy, style="Action.TButton"
        ).pack(pady=5)

    def on_date_range_filter(self, event=None):
        """Handle date range filter selection and show/hide custom date inputs."""
        date_range = self.date_range_var.get()
        if date_range == "Custom":
            self.custom_dates_frame.grid()
        else:
            self.custom_dates_frame.grid_remove()
        self.current_page = 1
        for var in self.order_vars.values():
            var.set(False)
        self.select_all_var.set(False)
        self.update_orders_list()

    def process_orders(self, orders_input):
        """Process orders entered by the user."""
        success, message = database.process_orders(
            orders_input, self.json_file, self.excel_folder
        )
        if success:
            self.update_orders_list()
        return success, message

    def update_orders_list(self):
        """Update the orders list UI with filtered data."""
        for widget in self.orders_inner_frame.winfo_children():
            widget.destroy()
        self.order_checkbuttons.clear()

        # Get date range filter
        date_range = self.date_range_var.get()
        today = datetime.now().date()
        start_date = None
        end_date = today

        if date_range == "Last 30 Days":
            start_date = today - timedelta(days=30)
        elif date_range == "Last 60 Days":
            start_date = today - timedelta(days=60)
        elif date_range == "Last 90 Days":
            start_date = today - timedelta(days=90)
        elif date_range == "Custom":
            try:
                start_date_str = self.start_date_entry.get()
                end_date_str = self.end_date_entry.get()
                if start_date_str and end_date_str:
                    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
                    end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
                    if start_date > end_date:
                        messagebox.showerror(
                            "Error", "Start date cannot be after end date."
                        )
                        return
                else:
                    start_date = None
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD.")
                return

        selected_version = self.version_combobox.get()
        orders_list, versions, error = database.get_orders_list(
            self.json_file, selected_version, start_date, end_date
        )

        if error:
            messagebox.showerror("Error", error)
            return

        self.version_combobox["values"] = ["All"] + versions
        current = self.version_combobox.get()
        if current not in self.version_combobox["values"]:
            self.version_combobox.set("All")

        total_orders = len(orders_list)
        self.total_pages = max(
            1, (total_orders + self.orders_per_page - 1) // self.orders_per_page
        )
        if self.current_page > self.total_pages:
            self.current_page = self.total_pages

        start_idx = (self.current_page - 1) * self.orders_per_page
        end_idx = start_idx + self.orders_per_page
        paginated_orders = orders_list[start_idx:end_idx]

        for idx, (version, order, test_date) in enumerate(
                paginated_orders, start=start_idx + 1
        ):
            key = (version, order)
            if key not in self.order_vars:
                self.order_vars[key] = tk.BooleanVar()
            var = self.order_vars[key]
            display_text = (
                f"{idx}. Version: {version}, Order: {order}, Date: {test_date}"
            )

            row_frame = ttk.Frame(self.orders_inner_frame)
            row_frame.grid(row=idx - start_idx, column=0, sticky="w", padx=5, pady=2)

            chk = ttk.Checkbutton(row_frame, text=display_text, variable=var, width=40)
            chk.pack(side=tk.LEFT)

            btn_view = ttk.Button(
                row_frame,
                text=" 👁️",
                width=3,
                command=lambda v=version, o=order: self.show_metadata_popup(v, o),
                style="Action.TButton",
            )
            btn_view.pack(side=tk.LEFT, padx=(10, 0))

            self.order_checkbuttons[key] = chk

        self.orders_canvas.configure(scrollregion=self.orders_canvas.bbox("all"))
        self.orders_canvas.yview_moveto(0)

        self.page_info.configure(text=f"Page {self.current_page}/{self.total_pages}")

        self.prev_btn.configure(state="normal" if self.current_page > 1 else "disabled")
        self.next_btn.configure(
            state="normal" if self.current_page < self.total_pages else "disabled"
        )

        current_page_orders = [
            (version, order) for version, order, _ in paginated_orders
        ]
        all_selected = all(
            self.order_vars.get(key, tk.BooleanVar(value=False)).get()
            for key in current_page_orders
        )
        self.select_all_var.set(all_selected)

    def show_metadata_popup(self, version, order):
        """Show metadata for a specific order in a popup."""
        data, error = database.get_metadata(self.json_file, version, order)
        if error:
            messagebox.showerror("Error", error)
            return

        info = ""
        for k, v in data["metadata"].items():
            info += f"{k}: {v}\n"
        info += "\nTemperatures (°C):\n"
        for tipo, tdata in data["temperatures"].items():
            info += f"  {tipo}: {tdata.get('temperature_c', 'N/A')}\n"

        popup = tk.Toplevel(self.parent)
        popup.title(f"Metadata - {order}")
        popup.geometry("350x250")
        popup.resizable(False, False)
        ttk.Label(
            popup,
            text=f"Metadata for Order {order} ({version})",
            font=("Arial", 11, "bold"),
        ).pack(pady=8)
        text = tk.Text(popup, width=40, height=10, wrap="word")
        text.insert("1.0", info)
        text.configure(state="disabled")
        text.pack(padx=8, pady=8)
        ttk.Button(
            popup, text="Close", command=popup.destroy, style="Action.TButton"
        ).pack(pady=5)

    def on_items_per_page_changed(self, event=None):
        """Handle changes to items per page."""
        try:
            self.orders_per_page = int(self.page_selector.get())
        except:
            self.orders_per_page = 10
        self.current_page = 1
        self.update_orders_list()

    def change_page(self, delta):
        """Change the current page of orders."""
        new_page = self.current_page + delta
        if 1 <= new_page <= self.total_pages:
            self.current_page = new_page
            self.update_orders_list()

    def toggle_select_all(self):
        """Toggle selection of all orders on the current page."""
        state = self.select_all_var.get()
        start_idx = (self.current_page - 1) * self.orders_per_page
        end_idx = start_idx + self.orders_per_page

        orders_list, _, _ = database.get_orders_list(
            self.json_file, self.version_combobox.get(), None, None
        )
        paginated_orders = orders_list[start_idx:end_idx]

        for version, order, _ in paginated_orders:
            key = (version, order)
            if key in self.order_vars:
                self.order_vars[key].set(state)

    def on_version_filter(self, event=None):
        """Handle version filter changes."""
        self.current_page = 1
        for var in self.order_vars.values():
            var.set(False)
        self.select_all_var.set(False)
        self.update_orders_list()

    def remove_orders_by_input(self, orders_input):
        """Remove orders specified in the input field."""
        success, message = database.remove_orders(orders_input, self.json_file)
        if success:
            self.update_orders_list()
        return success, message

    def send_to_workplace(self):
        """Send selected orders to the workplace, preventing duplicates."""
        selected_orders = [
            (version, order)
            for (version, order), var in self.order_vars.items()
            if var.get()
        ]
        if not selected_orders:
            messagebox.showwarning(
                "Warning", "No orders selected to send to workplace."
            )
            return

        # Check for version conflicts
        if (
                self.workplace_manager.workplace_data
                and self.workplace_manager.workplace_data[0]["version"]
                not in set(version for version, _ in selected_orders)
        ):
            messagebox.showerror(
                "Error",
                "Workplace already contains tests from another version. Clear workplace first.",
            )
            return

        # Check for duplicate orders
        existing_orders = {
            (reg["version"], reg["order"])
            for reg in self.workplace_manager.workplace_data
        }
        duplicates = [order for order in selected_orders if order in existing_orders]
        if duplicates:
            duplicate_str = ", ".join(
                f"{version}/{order}" for version, order in duplicates
            )
            messagebox.showwarning(
                "Warning",
                f"The following orders are already in the workplace: {duplicate_str}.",
            )
            # Filter out duplicates
            selected_orders = [
                order for order in selected_orders if order not in existing_orders
            ]
            if not selected_orders:
                return

        # Fetch data for non-duplicate orders
        new_data, message = database.get_workplace_data(self.json_file, selected_orders)
        if not new_data and message:
            messagebox.showerror("Error", message)
            return

        # Add new data to workplace
        self.workplace_manager.workplace_data.extend(new_data)
        self.workplace_manager.update_workplace_display()
        self.workplace_manager.update_workplace_counters()
        messagebox.showinfo("Success", message)

    def remove_workplace_orders_selected(self):
        """Remove selected orders from the workplace."""
        selected_orders = {
            (version, order)
            for (version, order), var in self.order_vars.items()
            if var.get()
        }
        if not selected_orders:
            messagebox.showwarning(
                "Warning", "No orders selected to remove from Workplace."
            )
            return

        before = len(self.workplace_manager.workplace_data)
        self.workplace_manager.workplace_data = [
            reg
            for reg in self.workplace_manager.workplace_data
            if (reg["version"], reg["order"]) not in selected_orders
        ]
        after = len(self.workplace_manager.workplace_data)
        self.workplace_manager.update_workplace_display()
        self.workplace_manager.update_workplace_counters()
        messagebox.showinfo(
            "Success", f"{before - after} records removed from Workplace."
        )

    def clear_workplace(self):
        """Clear all data from the workplace."""
        self.workplace_manager.clear_workplace()
