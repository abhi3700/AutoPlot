# AutoPlot

## v1.8 - `25-02-2021`

- AutoPlot Web App is released with all data in Excel logbooks.

## v1.7.5 - `15-02-2021`

- "View time series chart with multiple size particle visualization for an Equipment chamber

## v1.7 - `12-December-2019`

- `run.sh` file is modified to update the charts (in the same directory) & opening in the browser. It has been linked to the `equipment_logbook.py` e.g. `CNT01_Ch_A_QC_LOG_BOOK.py` for execution. It opens the excel, update/opens the charts and then auto-close the excel file. Finally, it also kills the task - `excel.exe`.

## v1.6 - `27-November-2019`

- Added **Contour plot (2D)** for viewing Etch Rate (ER) parameter (customizable) variation across the wafer, on a specific date (of QC). Here, the contour plot is customized to 8-inch wafer (200 mm dia). Also, the data points are visible in 'x' shape.
- Fetch (x,y) coordinates of the wafer from Excel file.

## v1.5 - `14-November-2019`

#### <u>Major changes</u>

- **Control limit calculation (LCL, UCL):** Based on the last 30 (customizable) QC days's data points, the LCL & UCL of layers (like Nitride, Poly, etc..) can be automatically calculated for a specific date.
- **Mathematics:** 2 methods to calculate control limits.
  - <u>M-1 (Raw data):</u> based on (last_30_QC _ individual_points). E.g. for Nitride layer (13 points on wafer) ==> [30 _ 13 = 390]
  - <u>M-2 (Avg. data):</u> based on (last_30_QC_avg_points). E.g. for Nitride layer (13 points on wafer) ==> [30]
- **User Interface (UI):** Different buttons are added on the Excel's sheet for clearing data, fetching dates, calculating control limits, parsing LCL, UCL values. <kbd>Clear</kbd>, <kbd>Fetch</kbd>, <kbd>Calc</kbd>, <kbd>Change</kbd> buttons added for different layers in 1st example: CNT01 Ch# A
- **Change Control limit:** If required, the calculated LCL, UCL values can be parsed (on pressing a button) onto the excel sheet containing QC data.
- **Frequency:** The control limit calculation, data parsing is kept as manual as per our Fab's activity. For instance, after equipment's PM (scheduled/unscheduled) if there is a drift in the plot after QC repeatability, then the new control limits (LCL, UCL) is to be calculated and can also be parsed correspondingly onto the excel sheet containing QC data (like Etch Rate (ER), Thickness, ...).
- **New Plot:** Now, with the new calculated control limits (after parsing), the plot can be generated and viewed in the browser.

#### <u>Minor changes</u>

- Replaced 'NIL' remarks with '.'

## v1.1 - `06-June-2019`

- `auto_run.sh` script added to automatically generate plots via **Task Scheduler** (@ 0800 hrs daily) in respective Excel file directories.
- `input.py` file created for taking all the input params in a separate file.

## v1.0 - `24-May-2019`

- <kbd>RUN</kbd> Button added in Excel for Execution.
- `run.sh` file added in each Excel directory for execution without even opening the Excel file itself.
- Collapse all the logbooks' auto-execution in Task Scheduler (one `auto_run.bat` file per Section).
- Plotly (Offline) charts added, instead of Matplotlib charts (interactive missing- hovering, etc.)
- **`ViEW`** Software working as foundation for maintaining the **`AutoPlot`**'s script codebase.
