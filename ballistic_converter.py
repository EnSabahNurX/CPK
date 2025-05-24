import tkinter as tk
from tkinter import ttk, messagebox
import config
import database
from orders_manager import OrdersManager
from workplace_manager import WorkplaceManager
from export_database import export_database_to_excel


class ExcelToJsonConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Ballistic Tests Converter")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Configure ttk style for buttons to match original colors
        style = ttk.Style()
        style.configure(
            "Export.TButton",
            font=("Helvetica", 10, "bold"),
            foreground="black",
            background="#5cb85c",
        )
        style.map("Export.TButton", background=[("active", "#6fd66f")])
        style.configure("Action.TButton", font=("Helvetica", 10, "bold"))

        # Left Frame (Orders Manager)
        self.frame = ttk.Frame(self.root, relief="groove", padding=10, borderwidth=2)
        self.frame.grid(row=0, column=0, sticky="nsew")
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(5, weight=1)

        # Database Title Frame
        title_frame = ttk.Frame(self.frame)
        title_frame.grid(row=0, column=0, sticky="ew")
        title_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, minsize=150)

        # self.label_database_title = ttk.Label(
        #     title_frame, text="Database", font=("Helvetica", 14, "bold")
        # )
        # self.label_database_title.grid(row=0, column=0, sticky="w", pady=(0, 10))

        # self.btn_export_db = ttk.Button(
        #     title_frame,
        #     text="Export Database",
        #     command=self.export_database,
        #     style="Export.TButton",
        #     width=15,
        # )
        # self.btn_export_db.grid(row=0, column=1, sticky="e", padx=(0, 5))
        # ToolTip(self.btn_export_db, "Export entire database to Excel")

        # # Orders Input
        # ttk.Label(self.frame, text="Enter order numbers separated by commas:").grid(
        #     row=0, column=0, sticky="nsew"
        # )
        # self.entry_orders = ttk.Entry(self.frame)
        # self.entry_orders.grid(row=1, column=0, sticky="ew", pady=5)

        # # Process and Remove Orders Buttons Side by Side
        # btns_frame = ttk.Frame(self.frame)
        # btns_frame.grid(row=2, column=0, sticky="ew", pady=5)
        # btns_frame.columnconfigure((0, 1), weight=1)

        # self.btn_process = ttk.Button(
        #     btns_frame,
        #     text="Add Entered Orders",
        #     command=self.process_orders,
        #     style="Action.TButton",
        #     width=15,
        # )
        # self.btn_process.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        # ToolTip(self.btn_process, "Process and add orders to the database")

        # self.btn_remove_orders = ttk.Button(
        #     btns_frame,
        #     text="Remove Entered Orders",
        #     command=self.remove_orders_by_input,
        #     style="Action.TButton",
        #     width=15,
        # )
        # self.btn_remove_orders.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        # ToolTip(self.btn_remove_orders, "Remove specified orders from the database")

        self.status_label = ttk.Label(
            self.frame, text="", anchor="w", foreground="green"
        )
        self.status_label.grid(row=3, column=0, sticky="ew", pady=(5, 10))

        # Initializations
        self.json_file = config.JSON_FILE
        self.excel_folder = config.EXCEL_FOLDER
        database.create_daily_backup(self.json_file)

        # Instantiate Managers
        self.workplace_manager = WorkplaceManager(self.root, self.json_file)
        self.orders_manager = OrdersManager(
            self.frame, self.json_file, self.excel_folder, self.workplace_manager
        )
        self.orders_manager.update_orders_list()

    def process_orders(self):
        """Process orders entered by the user."""
        orders_input = self.entry_orders.get().strip()
        success, message = self.orders_manager.process_orders(orders_input)
        if success:
            self.entry_orders.delete(0, tk.END)
            self.status_label.configure(text=message)
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)

    def remove_orders_by_input(self):
        """Remove orders specified in the input field."""
        orders_input = self.entry_orders.get().strip()
        success, message = self.orders_manager.remove_orders_by_input(orders_input)
        if success:
            self.status_label.configure(text=message)
            messagebox.showinfo("Success", message)
            self.entry_orders.delete(0, tk.END)
        else:
            messagebox.showwarning("Warning", message)

    def export_database(self):
        """Export the database to Excel."""
        export_database_to_excel(self)
