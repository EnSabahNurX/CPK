# Ballistic Tests Reporting Application

## Overview
The **Ballistic Tests Reporting Application** is a Python-based desktop application designed to generate and visualize reports for ballistic test data, specifically pressure curves across different temperatures (Room Temperature [RT], Low Temperature [LT], and High Temperature [HT]). The application provides a user-friendly Tkinter GUI to display interactive graphs and tables, with options to export reports to PDF and Excel formats. It processes test data stored in JSON format, groups it by temperature, and visualizes pressure measurements over time, including maximum, minimum, and mean values, with dynamic resizing and scrolling capabilities.

The application is designed for engineers or researchers analyzing ballistic test data, ensuring that reports are consistently formatted and that temperature data is displayed without gaps, even when fewer than three temperatures are available. The codebase emphasizes modularity, with separate utilities for data export and a robust GUI for data visualization.

## Features
- **Interactive GUI**:
  - Displays ballistic test data in a Tkinter window with a scrollable frame.
  - Each temperature (RT, LT, HT) is presented in a dedicated `LabelFrame` containing:
    - A Matplotlib graph showing pressure curves over time (ms), with maximum, minimum, and mean lines.
    - A `Treeview` table summarizing maximum, minimum, and mean pressure values at each time point.
  - Dynamically resizes graphs and tables based on window size.
  - Supports mouse wheel scrolling for large datasets.
- **Temperature Positioning**:
  - Automatically positions available temperatures at the top of the report window, ensuring no empty spaces (e.g., LT at top if RT is missing).
  - Prioritizes RT, then LT, then HT in display order.
- **Data Export**:
  - Exports reports to PDF with consistent formatting, including graphs and tables for each temperature.
  - Exports data to Excel for further analysis.
  - Ensures PDF and Excel exports align with the GUI display (e.g., LT in top subplot if first in GUI).
- **Data Validation**:
  - Handles missing or invalid data gracefully, displaying warnings for empty datasets or missing pressure data.
  - Skips invalid temperature data to prevent errors like `RuntimeWarning: Mean of empty slice`.
- **Customizable Styling**:
  - Graphs use distinct colors for curves (`#444444`), maximum limits (`#d62728`), minimum limits (`#1f77b4`), and mean (`#2ca02c`).
  - Tables use colored backgrounds for maximum (`#ffcccc`), mean (`#ccffcc`), and minimum (`#cce6ff`) rows.
  - Uses `Helvetica` fonts for a professional appearance.
- **Error Handling**:
  - Displays user-friendly error messages via Tkinter `messagebox` for issues like missing JSON files or invalid data.
  - Includes traceback for debugging.

## Installation
### Prerequisites
- **Python**: Version 3.6 or higher.
- **Operating System**: Windows, macOS, or Linux.
- **Dependencies**:
  - `tkinter`: For the GUI (usually included with Python).
  - `matplotlib`: For plotting pressure curves.
  - `numpy`: For numerical computations.
  - `pandas`: For Excel export (assumed based on `export_utils.py`).
  - `reportlab`: For PDF generation (assumed based on `export_utils.py`).
  - `openpyxl` or `xlsxwriter`: For Excel file handling.
  - Custom `tooltip` module: For button tooltips.

### Steps
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/EnSabahNurX/CPK
   cd ballistic-tests-reporting
   ```

2. **Set Up a Virtual Environment** (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies**:
   Create a `requirements.txt` file with the following:
   ```
   matplotlib>=3.5.0
   numpy>=1.21.0
   pandas>=1.4.0
   reportlab>=3.6.0
   openpyxl>=3.0.0
   ```
   Then run:
   ```bash
   pip install -r requirements.txt
   ```

4. **Add Custom `tooltip` Module**:
   - Ensure the `tooltip.py` module (providing the `ToolTip` class) is in the project directory.
   - If unavailable, implement a basic `ToolTip` class or remove tooltip functionality (see [Customization](#customization)).

5. **Prepare JSON Data**:
   - Place your JSON file (containing test specifications and limits) in the project directory or specify its path in the application.
   - Expected JSON structure:
     ```json
     {
       "version_name": {
         "sample_order": {
           "temperatures": {
             "RT": {
               "limits": {
                 "maximums": {"time_ms": value, ...},
                 "minimums": {"time_ms": value, ...}
               }
             },
             "LT": {...},
             "HT": {...}
           }
         }
       }
     }
     ```

6. **Run the Application**:
   ```bash
   python main.py
   ```
   Replace `main.py` with the entry point script if different.

## Usage
1. **Launch the Application**:
   - Run the main script (e.g., `python main.py`).
   - A Tkinter window opens, allowing data input or selection (assumed to be handled by `main.py`).

2. **Load or Filter Data**:
   - Load ballistic test data (via `workplace_data` or `filtered_workplace_data`).
   - Data format (expected in `data_to_use`):
     ```python
     [
       {
         "type": "RT",  # or "LT", "HT"
         "version": "version_name",
         "order": "sample_order",
         "pressures": {"time_ms": pressure_value, ...}
       },
       ...
     ]
     ```

3. **Generate Report**:
   - Click a button or menu option (assumed in `main.py`) to call `show_report(self)`.
   - The Report window opens, displaying:
     - A header with the report title and generation timestamp.
     - A scrollable frame with `LabelFrame` widgets for each temperature.
     - Each `LabelFrame` contains a graph and table, or a warning if data is missing.
   - If only LT is present, LT appears at the top; if LT and HT, LT is at top, HT below, etc.

4. **Interact with the Report**:
   - **Scroll**: Use the mouse wheel or scrollbar to navigate.
   - **Resize**: Adjust the window to dynamically resize graphs and tables.
   - **Export**:
     - Click **Export to PDF** to save a PDF with graphs and tables.
     - Click **Export to Excel** to save data in Excel format.
     - Click **Close** to exit the Report window.

5. **Handle Errors**:
   - Warnings appear for empty data or missing pressure data.
   - Errors (e.g., invalid JSON) display a `messagebox` with details.

## File Structure
```
ballistic-tests-reporting/
├── main.py              # Entry point (assumed)
├── report.py            # Report generation and GUI logic
├── export_utils.py      # PDF and Excel export functions
├── tooltip.py           # Custom tooltip module (assumed)
├── data.json            # Sample JSON file with test specifications
├── requirements.txt     # Python dependencies
├── README.md            # This file
└── output/              # Directory for exported PDF and Excel files
```

### Key Files
- **`report.py`**:
  - Contains the `show_report(self)` function, which generates the Report window.
  - Groups data by temperature, creates `LabelFrame` widgets with Matplotlib graphs and `Treeview` tables.
  - Handles dynamic resizing, scrolling, and export button commands.
  - Ensures temperatures are displayed at the top without gaps.
- **`export_utils.py`**:
  - Provides `export_to_pdf`, `export_to_excel`, and `adjust_column_widths` functions.
  - Generates PDF reports with graphs and tables, matching the GUI layout.
  - Exports data to Excel, including pressure values and limits.
- **`tooltip.py`**:
  - Custom module for button tooltips (e.g., "Export report as PDF file").
- **`data.json`**:
  - Stores test specifications, including maximum and minimum pressure limits for each temperature.

## Dependencies
| Library       | Version   | Purpose                              |
|---------------|-----------|--------------------------------------|
| `tkinter`     | Built-in  | GUI framework for Report window      |
| `matplotlib`  | >=3.5.0   | Plotting pressure curves             |
| `numpy`       | >=1.21.0  | Numerical computations (e.g., mean)  |
| `pandas`      | >=1.4.0   | Excel export (assumed)               |
| `reportlab`   | >=3.6.0   | PDF generation (assumed)             |
| `openpyxl`    | >=3.0.0   | Excel file handling (assumed)        |

## Customization
To adapt the application:
- **Change Temperature Order**:
  - Modify the `available_temps` sorting in `report.py`:
    ```python
    available_temps = sorted(
        [temp for temp in data_by_temp if data_by_temp[temp]],
        key=lambda x: ["LT", "RT", "HT"].index(x) if x in ["LT", "RT", "HT"] else len(["LT", "RT", "HT"])
    )
    ```
    This prioritizes LT > RT > HT.
- **Adjust Styling**:
  - Update graph colors in `report.py` (e.g., change `#2ca02c` for mean line).
  - Modify table backgrounds in `table.tag_configure` (e.g., change `#ffcccc` for maximum row).
  - Adjust fonts in `style.configure("Treeview", font=("Arial", 12))`.
- **Remove Tooltips**:
  - If `tooltip.py` is unavailable, comment out:
    ```python
    def bind_tooltips():
        ToolTip(btn_export_pdf, "Export report as PDF file")
        ToolTip(btn_export_excel, "Export report as Excel file")
        ToolTip(btn_close, "Close report window")
    report_win.after(100, bind_tooltips)
    ```
- **Add Data Validation**:
  - Add checks in `report.py` to validate `pressures` keys (e.g., ensure they’re numeric).
  - Example:
    ```python
    ms_points = sorted(float(ms) for ms in ms_points if ms.replace(".", "").isdigit())
    ```

## Troubleshooting
- **No Graphs or Tables Displayed**:
  - Check `data_by_temp` by adding `print(data_by_temp)` after grouping data.
  - Verify `pressures` dictionaries contain valid keys and values.
  - Ensure `ms_points` is populated: `print(f"{temp} ms_points: {ms_points}")`.
  - Check `pressure_matrix.shape`: `print(f"{temp} pressure_matrix.shape: {pressure_matrix.shape}")`.
- **RuntimeWarning: Mean of empty slice**:
  - Indicates empty `pressure_matrix`. Ensure `records` have valid `pressures` data.
  - The updated `report.py` skips empty `pressure_matrix` with a warning label.
- **PDF Export Issues**:
  - Verify `table_data` and `ms_points_dict` by printing before export: `print(table_data, ms_points_dict)`.
  - Ensure `export_utils.py` is unchanged and compatible (artifact_id `a85d89dc-a352-445b-b6a2-2648b4bea24f`).
- **Invalid JSON**:
  - Check `data.json` structure matches the expected format.
  - Ensure file path in `self.json_file` is correct.
- **Error Messages**:
  - Errors are displayed via `messagebox`. Share the full traceback for debugging.

## Contributing
1. Fork the repository.
2. Create a feature branch: `git checkout -b feature-name`.
3. Commit changes: `git commit -m "Add feature"`.
4. Push to the branch: `git push origin feature-name`.
5. Open a pull request with a detailed description.

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.

## Acknowledgments
- Built with Python, Tkinter, and Matplotlib for robust data visualization.
- Inspired by the need for clear, professional ballistic test reports.
- Thanks to the open-source community for libraries like `reportlab` and `pandas`.

## Contact
For issues or feature requests, open an issue on the repository or contact the maintainer at ricklimadealmeida@gmail.com