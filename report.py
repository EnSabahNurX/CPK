import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import json
from datetime import datetime
from export_utils import export_to_excel, export_to_pdf
from tooltip import ToolTip


def show_report(self):
    try:
        # Use filtered data if available, otherwise use all workplace data
        data_to_use = (
            self.filtered_workplace_data
            if self.filtered_workplace_data is not None
            else self.workplace_data
        )

        if not data_to_use:
            messagebox.showwarning(
                "Warning",
                "Workplace empty or no data after filtering. Add tests before generating the report.",
            )
            return

        # Group data by temperature
        data_by_temp = {}
        for reg in data_to_use:
            temp = reg["type"]
            if temp not in data_by_temp:
                data_by_temp[temp] = []
            data_by_temp[temp].append(reg)

        # Define available temperatures, sorted to prioritize RT, LT, HT
        available_temps = sorted(
            [temp for temp in data_by_temp if data_by_temp[temp]],
            key=lambda x: (
                ["RT", "LT", "HT"].index(x)
                if x in ["RT", "LT", "HT"]
                else len(["RT", "LT", "HT"])
            ),
        )
        if not available_temps:
            messagebox.showwarning("Warning", "No valid temperature data to report.")
            return

        # Create report window
        report_win = tk.Toplevel(self.root)
        report_win.title("Ballistic Tests Report")
        report_win.geometry("1000x700")
        report_win.minsize(800, 600)
        report_win.focus_set()
        report_win.grab_set()

        # Configure ttk styles
        style = ttk.Style()
        style.configure(
            "Export.TButton",
            font=("Helvetica", 10, "bold"),
            foreground="black",
            background="#4682b4",
        )
        style.map("Export.TButton", background=[("active", "#5a9bd4")])
        style.configure(
            "Close.TButton",
            font=("Helvetica", 10, "bold"),
            foreground="black",
            background="#d9534f",
        )
        style.map("Close.TButton", background=[("active", "#e57373")])
        style.configure("Header.TFrame", background="#fafafa")
        style.configure("Header.TLabel", background="#fafafa")

        # Header
        header_frame = ttk.Frame(report_win, style="Header.TFrame")
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(
            header_frame,
            text="Ballistic Tests Report",
            font=("Helvetica", 16, "bold"),
            style="Header.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Label(
            header_frame,
            text=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            font=("Helvetica", 10),
            style="Header.TLabel",
        ).pack(side=tk.RIGHT)

        # Frame for buttons in top-right corner
        btn_frame = ttk.Frame(report_win, style="Header.TFrame")
        btn_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        # Export to PDF button
        btn_export_pdf = ttk.Button(
            btn_frame,
            text="Export to PDF",
            command=lambda: export_to_pdf(
                data_by_temp, table_data, ms_points_dict, self.json_file
            ),
            style="Export.TButton",
        )
        btn_export_pdf.pack(side=tk.RIGHT, padx=(5, 5))

        # Export to Excel button
        btn_export_excel = ttk.Button(
            btn_frame,
            text="Export to Excel",
            command=lambda: export_to_excel(data_by_temp, table_data, ms_points_dict),
            style="Export.TButton",
        )
        btn_export_excel.pack(side=tk.RIGHT, padx=(5, 5))

        # Close button
        btn_close = ttk.Button(
            btn_frame, text="Close", command=report_win.destroy, style="Close.TButton"
        )
        btn_close.pack(side=tk.RIGHT, padx=(5, 5))

        # Bind tooltips
        def bind_tooltips():
            ToolTip(btn_export_pdf, "Export report as PDF file")
            ToolTip(btn_export_excel, "Export report as Excel file")
            ToolTip(btn_close, "Close report window")

        report_win.after(100, bind_tooltips)

        # Frame with vertical scrollbar
        container = ttk.Frame(report_win)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg="#fafafa")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        canvas.create_window(
            (0, 0),
            window=scrollable_frame,
            anchor="nw",
            width=container.winfo_width(),
        )
        canvas.configure(yscrollcommand=scrollbar.set, bg="#fafafa")

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Add mouse wheel support
        def _on_report_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind(
            "<Enter>",
            lambda e: canvas.bind_all("<MouseWheel>", _on_report_mousewheel),
        )
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Update canvas width on container resize
        def update_canvas_width(event=None):
            if container.winfo_exists():
                canvas_width = container.winfo_width()
                canvas.itemconfig(
                    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw"),
                    width=canvas_width,
                )

        container.bind("<Configure>", update_canvas_width)

        # Store table data and ms_points for export
        table_data = [[] for _ in range(3)]  # [RT, LT, HT]
        ms_points_dict = {}

        # Fixed pressure points
        pressure_points = ["PK 10%", "PK 25%", "PK 50%", "PK 75%", "PK 90%", "PK MAX"]

        # For each temperature, plot graph and table
        for temp in available_temps:
            records = data_by_temp[temp]
            versions = set(r["version"] for r in records)
            version = ", ".join(versions) if len(versions) > 1 else list(versions)[0]
            total_inflators = len(records)

            temp_frame = ttk.LabelFrame(
                scrollable_frame,
                text=f"Temperature: {temp} | Version: {version} | Total Inflators: {total_inflators}",
                padding=5,
            )
            temp_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            temp_frame.columnconfigure(0, weight=1)
            temp_frame.rowconfigure(1, weight=1)

            ms_points = set()
            for r in records:
                if r["pressures"]:
                    ms_points.update(r["pressures"].keys())
            ms_points = sorted(int(ms) for ms in ms_points)
            ms_points_str = [str(ms) for ms in ms_points]
            ms_points_dict[temp] = ms_points_str

            if not ms_points:
                ttk.Label(
                    temp_frame,
                    text="No pressure data available for this temperature.",
                    font=("Helvetica", 10),
                ).pack(pady=10, fill=tk.BOTH, expand=True)
                table_data[["RT", "LT", "HT"].index(temp)] = []
                continue

            pressure_matrix = []
            for r in records:
                p = []
                if r["pressures"]:
                    for ms in ms_points_str:
                        val = r["pressures"].get(ms, np.nan)
                        p.append(val)
                else:
                    p = [np.nan] * len(ms_points)
                pressure_matrix.append(p)
            pressure_matrix = np.array(pressure_matrix, dtype=np.float64)

            if pressure_matrix.size == 0:
                ttk.Label(
                    temp_frame,
                    text="No valid pressure data available for this temperature.",
                    font=("Helvetica", 10),
                ).pack(pady=10, fill=tk.BOTH, expand=True)
                table_data[["RT", "LT", "HT"].index(temp)] = []
                continue

            limits_max = []
            limits_min = []
            try:
                with open(self.json_file, "r", encoding="utf-8") as f:
                    data_json = json.load(f)
                sample_order = records[0]["order"]
                limits = data_json[version][sample_order]["temperatures"][temp][
                    "limits"
                ]
                max_dict = limits.get("maximums", {})
                min_dict = limits.get("minimums", {})
                limits_max = [max_dict.get(str(ms), np.nan) for ms in ms_points]
                limits_min = [min_dict.get(str(ms), np.nan) for ms in ms_points]
            except Exception as e:
                limits_max = [np.nan] * len(ms_points)
                limits_min = [np.nan] * len(ms_points)

            mean = np.nanmean(pressure_matrix, axis=0)

            # Align ms_points_str with pressure_points
            if len(ms_points_str) > len(pressure_points):
                ms_points_str = ms_points_str[: len(pressure_points)]
            elif len(ms_points_str) < len(pressure_points):
                ms_points_str.extend(
                    ["-"] * (len(pressure_points) - len(ms_points_str))
                )

            # Initial figure
            fig, ax = plt.subplots(figsize=(8, 4))
            fig.patch.set_facecolor("#fafafa")
            ax.set_facecolor("#fafafa")

            # Initial plot
            for p in pressure_matrix:
                ax.plot(
                    ms_points,
                    p,
                    color="#444444",
                    linewidth=1,
                    alpha=0.5,
                )
            ax.plot(
                ms_points,
                limits_max,
                color="#d62728",
                linewidth=2,
                label="Maximum Limit",
                linestyle="--",
            )
            ax.plot(
                ms_points,
                limits_min,
                color="#1f77b4",
                linewidth=2,
                label="Minimum Limit",
                linestyle="--",
            )
            ax.plot(
                ms_points,
                mean,
                color="#7CE04A",
                linewidth=2.5,
                label="Mean",
                linestyle="-",
            )
            ax.set_title(f"Pressure Curves - Temperature {temp}", fontsize=12, pad=10)
            ax.set_xlabel("Time (ms)", fontsize=10)
            ax.set_ylabel("Pressure (bar)", fontsize=10)
            ax.legend(loc="lower right", fontsize=8)
            ax.grid(True, color="#cccccc", linestyle="--", linewidth=0.7)
            ax.minorticks_on()
            ax.grid(True, which="minor", color="#e0e0e0", linestyle=":", linewidth=0.5)

            def update_graph(event=None):
                if not temp_frame.winfo_exists():
                    return
                frame_width = temp_frame.winfo_width()
                frame_height = temp_frame.winfo_height()
                fig_width = max(4, frame_width / 100 * 0.98)
                fig_height = max(2, frame_height / 100 * 0.6)
                fig.set_size_inches(fig_width, fig_height)
                ax.clear()
                ax.set_facecolor("#fafafa")
                for p in pressure_matrix:
                    ax.plot(
                        ms_points,
                        p,
                        color="#444444",
                        linewidth=1,
                        alpha=0.5,
                    )
                ax.plot(
                    ms_points,
                    limits_max,
                    color="#d62728",
                    linewidth=2,
                    label="Maximum Limit",
                    linestyle="--",
                )
                ax.plot(
                    ms_points,
                    limits_min,
                    color="#1f77b4",
                    linewidth=2,
                    label="Minimum Limit",
                    linestyle="--",
                )
                ax.plot(
                    ms_points,
                    mean,
                    color="#7CE04A",
                    linewidth=2.5,
                    label="Mean",
                    linestyle="-",
                )
                ax.set_title(
                    f"Pressure Curves - Temperature {temp}", fontsize=12, pad=10
                )
                ax.set_xlabel("Time (ms)", fontsize=10)
                ax.set_ylabel("Pressure (bar)", fontsize=10)
                ax.legend(loc="lower right", fontsize=8)
                ax.grid(True, color="#cccccc", linestyle="--", linewidth=0.7)
                ax.minorticks_on()
                ax.grid(
                    True, which="minor", color="#e0e0e0", linestyle=":", linewidth=0.5
                )
                canvas_fig.draw()
                canvas_fig.get_tk_widget().update()

            temp_frame.bind("<Configure>", update_graph)

            canvas_fig = FigureCanvasTkAgg(fig, master=temp_frame)
            canvas_fig.get_tk_widget().grid(row=0, column=0, sticky="nsew", pady=5)
            temp_frame.rowconfigure(0, weight=3)
            temp_frame.rowconfigure(1, weight=1)

            # Combined Table with Label column
            table_columns = ["Label"] + pressure_points
            table = ttk.Treeview(
                temp_frame,
                columns=table_columns,
                show="headings",
                height=4,
            )
            table.heading("Label", text="")
            for pk in pressure_points:
                table.heading(pk, text=pk)
                table.column(pk, anchor="center", stretch=True)
            table.column("Label", anchor="center", stretch=True)

            def format_row(row):
                return [f"{v:.2f}" if not np.isnan(v) else "-" for v in row]

            # Insert rows with labels
            table.insert(
                "", "end", values=["Time (ms)"] + ms_points_str, tags=("time",)
            )
            table.insert(
                "",
                "end",
                values=["Maximum (bar)"] + format_row(limits_max),
                tags=("max",),
            )
            table.insert(
                "", "end", values=["Mean (bar)"] + format_row(mean), tags=("mean",)
            )
            table.insert(
                "",
                "end",
                values=["Minimum (bar)"] + format_row(limits_min),
                tags=("min",),
            )

            table.tag_configure("time", background="#f0f0f0")
            table.tag_configure("max", background="#ffcccc")
            table.tag_configure("mean", background="#7CE04A")
            table.tag_configure("min", background="#cce6ff")

            style = ttk.Style()
            style.configure("Treeview", font=("Helvetica", 6), rowheight=20)
            style.configure("Treeview.Heading", font=("Helvetica", 6, "bold"))

            def update_table_columns(
                event=None, tf=temp_frame, tbl=table, cols=table_columns
            ):
                if not tf.winfo_exists():
                    return
                frame_width = tf.winfo_width() - 10
                num_cols = len(cols)
                if num_cols == 0 or frame_width <= 1:
                    tf.after(100, lambda: update_table_columns(None, tf, tbl, cols))
                    return
                col_width = max(60, int(frame_width / num_cols))
                for col in cols:
                    tbl.column(col, width=col_width, anchor="center")

            temp_frame.bind(
                "<Configure>",
                lambda e, tf=temp_frame, tbl=table, cols=table_columns: update_table_columns(
                    e, tf, tbl, cols
                ),
            )
            temp_frame.after(
                200,
                lambda tf=temp_frame, tbl=table, cols=table_columns: update_table_columns(
                    None, tf, tbl, cols
                ),
            )
            update_table_columns(tf=temp_frame, tbl=table, cols=table_columns)

            table.grid(row=1, column=0, sticky="nsew", padx=(5, 15), pady=5)

            # Update table_data for export
            table_data[["RT", "LT", "HT"].index(temp)] = [
                ("Time (ms)", ["Time (ms)"] + ms_points_str),
                ("Maximum (bar)", ["Maximum (bar)"] + format_row(limits_max)),
                ("Mean (bar)", ["Mean (bar)"] + format_row(mean)),
                ("Minimum (bar)", ["Minimum (bar)"] + format_row(limits_min)),
            ]

        report_win.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"Error generating report: {str(e)}")
