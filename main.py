import tkinter as tk
from tkinter import messagebox
import config
import database
from orders_manager import OrdersManager
from workplace_manager import WorkplaceManager
from tooltip import ToolTip
from export_database import export_database_to_excel


class ExcelToJsonConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Ballistic Tests Converter")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Left Frame (Database)
        self.frame_db = tk.Frame(self.root, relief="groove", padx=10, pady=10, bd=2)
        self.frame_db.grid(row=0, column=0, sticky="nsew")
        self.frame_db.columnconfigure(0, weight=1)
        self.frame_db.rowconfigure(5, weight=1)

        # Database Title Frame
        title_frame = tk.Frame(self.frame_db)
        title_frame.grid(row=0, column=0, sticky="ew")
        title_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, minsize=150)

        self.label_database_title = tk.Label(
            title_frame, text="Database", font=("Helvetica", 14, "bold")
        )
        self.label_database_title.grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.btn_export_db = tk.Button(
            title_frame,
            text="Export Database",
            command=self.export_database,
            font=("Helvetica", 10, "bold"),
            width=15,
            bg="#5cb85c",
            fg="white",
        )
        self.btn_export_db.grid(row=0, column=1, sticky="e", padx=(0, 5))
        self.btn_export_db.bind(
            "<Enter>", lambda e: self.btn_export_db.config(bg="#6fd66f")
        )
        self.btn_export_db.bind(
            "<Leave>", lambda e: self.btn_export_db.config(bg="#5cb85c")
        )
        ToolTip(self.btn_export_db, "Export entire database to Excel")

        # Orders Input
        tk.Label(self.frame_db, text="Enter order numbers separated by commas:").grid(
            row=1, column=0, sticky="w", pady=(10, 5)
        )
        self.entry_orders = tk.Entry(self.frame_db)
        self.entry_orders.grid(row=2, column=0, sticky="ew", pady=5)

        # Process and Remove Orders Buttons Side by Side
        btns_frame = tk.Frame(self.frame_db)
        btns_frame.grid(row=3, column=0, sticky="ew", pady=5)
        btns_frame.columnconfigure((0, 1), weight=1)

        self.btn_process = tk.Button(
            btns_frame,
            text="Add Entered Orders",
            command=self.process_orders,
            font=("Helvetica", 10, "bold"),
            width=15,
        )
        self.btn_process.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ToolTip(self.btn_process, "Process and add orders to the database")

        self.btn_remove_orders = tk.Button(
            btns_frame,
            text="Remove Entered Orders",
            command=self.remove_orders_by_input,
            font=("Helvetica", 10, "bold"),
            width=15,
        )
        self.btn_remove_orders.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        ToolTip(self.btn_remove_orders, "Remove specified orders from the database")

        self.status_label = tk.Label(self.frame_db, text="", anchor="w", fg="green")
        self.status_label.grid(row=4, column=0, sticky="ew", pady=(5, 10))

        # Initializations
        self.json_file = config.JSON_FILE
        self.excel_folder = config.EXCEL_FOLDER
        database.create_daily_backup(self.json_file)

        # Instantiate Managers
        self.workplace_manager = WorkplaceManager(self.root, self.json_file)
        self.orders_manager = OrdersManager(
            self.frame_db, self.json_file, self.excel_folder, self.workplace_manager
        )

        self.orders_manager.update_orders_list()
        self.root.mainloop()

    def process_orders(self):
        """Process orders entered by the user."""
        orders_input = self.entry_orders.get().strip()
        success, message = self.orders_manager.process_orders(orders_input)
        if success:
            self.entry_orders.delete(0, tk.END)
            self.status_label.config(text=message)
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)

    def remove_orders_by_input(self):
        """Remove orders specified in the input field."""
        orders_input = self.entry_orders.get().strip()
        success, message = self.orders_manager.remove_orders_by_input(orders_input)
        if success:
            self.status_label.config(text=message)
            messagebox.showinfo("Success", message)
            self.entry_orders.delete(0, tk.END)
        else:
            messagebox.showwarning("Warning", message)

    def export_database(self):
        """Export the database to Excel."""
        export_database_to_excel(self)


if __name__ == "__main__":
    ExcelToJsonConverter()
