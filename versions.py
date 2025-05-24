import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import config


class VersionManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Version Control")
        self.root.geometry("900x600")
        self.root.minsize(900, 600)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Configuração de estilos modernos
        style = ttk.Style()
        style.configure("TButton", font=("Helvetica", 11, "bold"), padding=5)
        style.configure("Add.TButton", foreground="black", background="#5cb85c")
        style.map("Add.TButton", background=[("active", "#6fd66f")])
        style.configure("Edit.TButton", foreground="black", background="#f0ad4e")
        style.map("Edit.TButton", background=[("active", "#f5bf77")])
        style.configure("Delete.TButton", foreground="black", background="#d9534f")
        style.map("Delete.TButton", background=[("active", "#e67c73")])
        style.configure("Refresh.TButton", foreground="black", background="#0275d8")
        style.map("Refresh.TButton", background=[("active", "#2a84df")])

        self.json_file = config.JSON_FILE

        # Criar layout de divisão (PanedWindow)
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        # Criar área da lista de versões
        self.frame_versions = ttk.Frame(self.paned_window, relief="groove", padding=10)
        self.frame_versions.columnconfigure(0, weight=1)
        self.frame_versions.rowconfigure(0, weight=1)

        self.tree_versions = ttk.Treeview(
            self.frame_versions, columns=("Version"), show="headings", height=10
        )
        self.tree_versions.heading("Version", text="Versions")
        self.tree_versions.column("Version", width=150)
        self.tree_versions.grid(row=0, column=0, sticky="nsew")
        self.load_versions()

        self.paned_window.add(self.frame_versions, weight=1)

        # Botões de ação
        btn_frame = ttk.Frame(self.root)
        btn_frame.grid(row=2, column=0, pady=10)

        self.btn_add = ttk.Button(
            btn_frame,
            text="Add",
            command=self.add_version,
            style="Add.TButton",
            width=15,
        )
        self.btn_add.grid(row=0, column=0, padx=5)

        self.btn_edit = ttk.Button(
            btn_frame,
            text="Edit",
            command=self.edit_version,
            style="Edit.TButton",
            width=15,
        )
        self.btn_edit.grid(row=0, column=1, padx=5)

        self.btn_delete = ttk.Button(
            btn_frame,
            text="Delete",
            command=self.delete_version,
            style="Delete.TButton",
            width=15,
        )
        self.btn_delete.grid(row=0, column=2, padx=5)

        self.btn_refresh = ttk.Button(
            btn_frame,
            text="Refresh",
            command=self.load_versions,
            style="Refresh.TButton",
            width=15,
        )
        self.btn_refresh.grid(row=0, column=3, padx=5)

    def load_versions(self):
        """Carregar todas as versões existentes no Data.JSON."""
        for item in self.tree_versions.get_children():
            self.tree_versions.delete(item)

        try:
            with open(self.json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

                # Iterar sobre cada versão na raiz do JSON, ignorando "versions"
                for version_id, version_data in data.items():
                    if version_id != "versions" and isinstance(version_data, dict):
                        self.tree_versions.insert("", "end", values=(version_id,))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load versions: {str(e)}")

    def create_colored_table(self, parent, data=None):
        """Criar uma tabela colorida e pré-preencher valores dos limites encontrados."""
        frame_table = ttk.Frame(parent)
        frame_table.pack(fill="both", expand=True, padx=10, pady=10)

        labels = ["PK10", "PK25", "PK50", "PK75", "PK90", "PKMax"]
        temperatures = ["RT", "LT", "HT"]
        categories = ["minimums", "baseline", "maximums"]
        colors = {
            "RT": ["#DFFFD6", "#B8E2A7", "#8CC882"],
            "LT": ["#D6ECFF", "#A7D0E2", "#82B8CC"],
            "HT": ["#FFD6D6", "#E2A7A7", "#C88282"],
        }

        entries = {}

        table_grid = ttk.Frame(frame_table)
        table_grid.pack(fill="both", expand=True)

        header_frame = ttk.Frame(table_grid)
        header_frame.grid(row=0, column=0, columnspan=len(labels) + 1, sticky="ew")

        ttk.Label(header_frame, text="Temp", anchor="center", width=15).grid(
            row=0, column=0, sticky="ew", padx=5
        )
        for j, label in enumerate(labels):
            ttk.Label(header_frame, text=label, anchor="center", width=10).grid(
                row=0, column=j + 1, sticky="ew", padx=5
            )

        row_index = 1
        for temp in temperatures:
            for idx, category in enumerate(categories):
                color = colors[temp][idx]

                row_frame = tk.Frame(table_grid, bg=color)
                row_frame.grid(
                    row=row_index, column=0, columnspan=len(labels) + 1, sticky="ew"
                )

                label_category = ttk.Label(
                    row_frame,
                    text=f"{temp} - {category}",
                    background=color,
                    anchor="center",
                    width=15,
                )
                label_category.grid(row=0, column=0, sticky="ew", padx=5)

                entries[(temp, category)] = []
                for j, label in enumerate(labels):
                    value = (
                        data.get("temperatures", {})
                        .get(temp, {})
                        .get("limits", {})
                        .get(category, {})
                        .get(label, 0.0)
                        if data
                        else 0.0
                    )
                    entry = ttk.Entry(row_frame, width=10, justify="center")
                    entry.insert(0, str(value))
                    entry.grid(row=0, column=j + 1, sticky="ew", padx=2, pady=2)
                    entries[(temp, category)].append(entry)

                row_index += 1

        return frame_table, entries

    def add_version(self):
        """Adicionar uma nova versão utilizando uma tabela interativa."""
        add_window = tk.Toplevel(self.root)

        # Perguntar nome da versão antes de exibir a janela
        version_name = simpledialog.askstring("Add Version", "Enter version name:")
        if not version_name:
            add_window.destroy()
            return

        add_window.title(f"Create Version: {version_name}")
        add_window.geometry("900x500")
        add_window.transient(self.root)
        add_window.grab_set()
        add_window.focus_set()

        frame_table, entries = self.create_colored_table(add_window)

        def save_new_version():
            """Salvar nova versão no JSON."""
            new_version_data = {
                temp: {cat: {} for cat in ["minimums", "baseline", "maximums"]}
                for temp in ["RT", "LT", "HT"]
            }
            for temp in ["RT", "LT", "HT"]:
                for category in ["minimums", "baseline", "maximums"]:
                    for j, label in enumerate(
                        ["PK10", "PK25", "PK50", "PK75", "PK90", "PKMax"]
                    ):
                        value = entries[(temp, category)][j].get()
                        try:
                            new_version_data[temp][category][label] = float(
                                value
                            )  # Apenas números válidos
                        except ValueError:
                            new_version_data[temp][category][label] = (
                                0.0  # Substituir caracteres inválidos por 0.0
                            )

            with open(self.json_file, "r+", encoding="utf-8") as f:
                data = json.load(f)
                data[version_name] = new_version_data
                f.seek(0)
                json.dump(data, f, ensure_ascii=False, indent=2)
                f.truncate()

            messagebox.showinfo(
                "Success", f"Version '{version_name}' created successfully!"
            )
            add_window.destroy()
            self.load_versions()

        ttk.Button(
            add_window,
            text="Save Version",
            command=save_new_version,
            style="Add.TButton",
        ).pack(pady=10)

    def edit_version(self):
        """Editar uma versão existente com a tabela interativa."""
        selected_version = self.get_selected_version()
        if not selected_version:
            return

        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Edit Version: {selected_version}")
        edit_window.geometry("900x500")
        edit_window.transient(self.root)
        edit_window.grab_set()
        edit_window.focus_set()

        with open(self.json_file, "r", encoding="utf-8") as f:
            data = json.load(f).get(selected_version, {})

        frame_table, entries = self.create_colored_table(edit_window, data)

        def save_changes():
            """Salvar alterações no JSON."""
            updated_data = {
                temp: {cat: {} for cat in ["minimums", "baseline", "maximums"]}
                for temp in ["RT", "LT", "HT"]
            }
            for temp in ["RT", "LT", "HT"]:
                for category in ["minimums", "baseline", "maximums"]:
                    for j, label in enumerate(
                        ["PK10", "PK25", "PK50", "PK75", "PK90", "PKMax"]
                    ):
                        value = entries[(temp, category)][j].get()
                        try:
                            updated_data[temp][category][label] = float(
                                value
                            )  # Apenas números válidos
                        except ValueError:
                            updated_data[temp][category][label] = (
                                0.0  # Substituir caracteres inválidos por 0.0
                            )

            with open(self.json_file, "r+", encoding="utf-8") as f:
                full_data = json.load(f)
                full_data[selected_version] = updated_data
                f.seek(0)
                json.dump(full_data, f, ensure_ascii=False, indent=2)
                f.truncate()

            messagebox.showinfo(
                "Success", f"Version '{selected_version}' updated successfully!"
            )
            edit_window.destroy()
            self.load_versions()

        ttk.Button(
            edit_window, text="Save Changes", command=save_changes, style="Edit.TButton"
        ).pack(pady=10)

    def delete_version(self):
        """Remover uma versão."""
        selected_version = self.get_selected_version()
        if not selected_version:
            return

        if messagebox.askyesno(
            "Delete Version", f"Are you sure you want to delete '{selected_version}'?"
        ):
            try:
                with open(self.json_file, "r+", encoding="utf-8") as f:
                    data = json.load(f)
                    del data[selected_version]
                    f.seek(0)
                    json.dump(data, f, ensure_ascii=False, indent=2)
                    f.truncate()

                messagebox.showinfo(
                    "Success", f"Version '{selected_version}' deleted successfully!"
                )
                self.load_versions()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete version: {str(e)}")

    def get_selected_version(self):
        """Obter a versão selecionada na lista."""
        selected_item = self.tree_versions.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "No version selected!")
            return None
        return self.tree_versions.item(selected_item, "values")[0]
