#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot

# CP Plot
cp_plot_title = 'CP Plot for RESP1B'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'RESP1B_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot

# BARC ER Plot
er_barc_plot_title = 'BARC ER Plot for RESP1B'  # title for ER plot
er_barc_plot_xlabel = 'Date'        # xaxis name for ER plot
er_barc_plot_ylabel = 'BARC ER (A/min)'   # yaxis name for ER plot
er_barc_plot_html_file = 'RESP1B_BARC_ER-Plot.html'   # HTML filename for ER plot
er_barc_plot_trace_count = 5    # no. of traces in ER plot

# BARC Unif Plot
unif_barc_plot_title = 'BARC Uniformity Plot for RESP1B'  # title for Unif plot
unif_barc_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_barc_plot_ylabel = 'BARC Unif (%)'    # yaxis name for Unif plot
unif_barc_plot_html_file = 'RESP1B_BARC_Unif-Plot.html'   # HTML filename for Unif plot
unif_barc_plot_trace_count = 2    # no. of traces in Unif plot

# PR ER Plot
er_pr_plot_title = 'PR ER Plot for RESP1B'  # title for ER plot
er_pr_plot_xlabel = 'Date'        # xaxis name for ER plot
er_pr_plot_ylabel = 'PR ER (A/min)'   # yaxis name for ER plot
er_pr_plot_html_file = 'RESP1B_PR_ER-Plot.html'   # HTML filename for ER plot
er_pr_plot_trace_count = 5    # no. of traces in ER plot

# PR Unif Plot
unif_pr_plot_title = 'PR Uniformity Plot for RESP1B'  # title for Unif plot
unif_pr_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_pr_plot_ylabel = 'PR Unif (%)'    # yaxis name for Unif plot
unif_pr_plot_html_file = 'RESP1B_PR_Unif-Plot.html'   # HTML filename for Unif plot
unif_pr_plot_trace_count = 2    # no. of traces in Unif plot

# TEOS ER Plot
er_teos_plot_title = 'TEOS ER Plot for RESP1B'  # title for ER plot
er_teos_plot_xlabel = 'Date'        # xaxis name for ER plot
er_teos_plot_ylabel = 'TEOS ER (A/min)'   # yaxis name for ER plot
er_teos_plot_html_file = 'RESP1B_TEOS_ER-Plot.html'   # HTML filename for ER plot
er_teos_plot_trace_count = 5    # no. of traces in ER plot

# TEOS Unif Plot
unif_teos_plot_title = 'TEOS Uniformity Plot for RESP1B'  # title for Unif plot
unif_teos_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_teos_plot_ylabel = 'TEOS Unif (%)'    # yaxis name for Unif plot
unif_teos_plot_html_file = 'RESP1B_TEOS_Unif-Plot.html'   # HTML filename for Unif plot
unif_teos_plot_trace_count = 3    # no. of traces in Unif plot

# SiN ER Plot
er_sin_plot_title = 'SiN ER Plot for RESP1B'  # title for ER plot
er_sin_plot_xlabel = 'Date'        # xaxis name for ER plot
er_sin_plot_ylabel = 'SiN ER (A/min)'   # yaxis name for ER plot
er_sin_plot_html_file = 'RESP1B_SiN_ER-Plot.html'   # HTML filename for ER plot
er_sin_plot_trace_count = 5    # no. of traces in ER plot

# SiN Unif Plot
unif_sin_plot_title = 'SiN Uniformity Plot for RESP1B'  # title for Unif plot
unif_sin_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_sin_plot_ylabel = 'SiN Unif (%)'    # yaxis name for Unif plot
unif_sin_plot_html_file = 'RESP1B_SiN_Unif-Plot.html'   # HTML filename for Unif plot
unif_sin_plot_trace_count = 3    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'RESP1B-CP'
sht_name_er_barc_pr_teos = 'ER-BARC,PR & TEOS'
sht_name_er_sin = 'SIN ER'

sht_cp_columns = ["Date (MM/DD/YYYY)", "delta CP", "USL", "Remarks"]
sht_er_barc_pr_teos_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]
sht_er_sin_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]


# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"

excel_file_directory = "I:\\github_repos\\AutoPlot\\examples\\dry_etch\\UNT02_Ch_B_QC_LOG_BOOK\\UNT02_Ch_B_QC_LOG_BOOK.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\UNT_02_LOG_BOOK\\UNT02_Ch_B_QC_LOG_BOOK\\UNT02_Ch_B_QC_LOG_BOOK.xlsm"
