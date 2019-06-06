#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot

# CP Plot
cp_plot_title = 'CP Plot for ASBE1'  # title for CP plot
cp_plot_xlabel = 'Date (MM/DD/YYYY)'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'ASBE1_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 3    # no. of traces in CP plot

# ER Plot
er_plot_title = 'ER Plot for ASBE1'  # title for ER plot
er_plot_xlabel = 'Date (MM/DD/YYYY)'        # xaxis name for ER plot
er_plot_ylabel = 'Etch Rate (A/min)'   # yaxis name for ER plot
er_plot_html_file = 'ASBE1_ER-Plot.html'   # HTML filename for ER plot
er_plot_trace_count = 4    # no. of traces in ER plot

# Unif Plot
unif_plot_title = 'Uniformity Plot for ASBE1'  # title for UNIF plot
unif_plot_xlabel = 'Date (MM/DD/YYYY)'      # xaxis name for Unif plot
unif_plot_ylabel = 'Uniformity (%)'    # yaxis name for Unif plot
unif_plot_html_file = 'ASBE1_Unif-Plot.html'   # HTML filename for Unif plot
unif_plot_trace_count = 2    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'ASBE1-CP'
sht_name_er = 'ASBE1-ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "delta CP", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "Remarks"]

# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"

excel_file_directory = "I:\\github_repos\\AutoPlot\\examples\\dry_etch\\ASH10_QC_LOG_BOOK\\ASH10_QC_LOG_BOOK.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\ASH_09_10_LOG_BOOK\\ASH10_QC_LOG_BOOK\\ASH10_QC_LOG_BOOK.xlsm"
