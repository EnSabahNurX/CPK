import database_front
import tkinter as tk
from tkinter import ttk
import ballistic_converter


class PlatformApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ballistic Converter Platform")
        self.root.geometry("600x400")
        self.root.minsize(500, 300)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Configuração de estilo ttk
        style = ttk.Style()
        style.configure(
            "StartButton.TButton",
            font=("Helvetica", 12, "bold"),
            foreground="black",
            background="#5cb85c",
        )
        style.map("StartButton.TButton", background=[("active", "#6fd66f")])
        style.configure(
            "DatabaseButton.TButton",
            font=("Helvetica", 12, "bold"),
            foreground="black",
            background="#5bc0de",
        )
        style.map("DatabaseButton.TButton", background=[("active", "#6fd6de")])
        style.configure(
            "VersionButton.TButton",
            font=("Helvetica", 12, "bold"),
            foreground="black",
            background="#de7c6f",
        )
        style.map("VersionButton.TButton", background=[("active", "#de7c6f")])

        # Frame principal
        self.main_frame = ttk.Frame(
            self.root, relief="groove", padding=10, borderwidth=2
        )
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.columnconfigure(0, weight=1)

        # Botão para iniciar o conversor
        self.btn_start_converter = ttk.Button(
            self.main_frame,
            text="Start Ballistic Converter",
            command=self.start_converter,
            style="StartButton.TButton",
            width=20,
        )
        self.btn_start_converter.grid(row=0, column=0, pady=10)

        # Botão para abrir o Database
        self.btn_open_database = ttk.Button(
            self.main_frame,
            text="Database",
            command=self.open_database,
            style="DatabaseButton.TButton",
            width=20,
        )
        self.btn_open_database.grid(row=1, column=0, pady=10)

        self.btn_version_control = ttk.Button(
            self.main_frame,
            text="Version Control",
            command=self.open_version_control,
            style="VersionButton.TButton",
            width=20,
        )
        self.btn_version_control.grid(row=2, column=0, pady=10)

    def start_converter(self):
        """Cria uma nova janela e mantém o foco nela."""
        new_window = tk.Toplevel(self.root)
        ballistic_converter.ExcelToJsonConverter(new_window)

        self.root.iconify()
        new_window.focus_set()

    def open_database(self):
        """Cria uma nova janela apenas para o frame de database e mantém o foco nela."""

        new_window = tk.Toplevel(self.root)
        database_front.DatabaseFront(new_window)

        self.root.iconify()
        new_window.focus_set()

    def open_version_control(self):
        """Abrir a janela de controle de versões."""
        import versions

        new_window = tk.Toplevel(self.root)
        versions.VersionManager(new_window)

        self.root.iconify()
        new_window.focus_set()


if __name__ == "__main__":
    root = tk.Tk()
    app = PlatformApp(root)
    root.mainloop()
