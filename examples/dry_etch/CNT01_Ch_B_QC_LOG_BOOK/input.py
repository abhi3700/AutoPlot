#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot

# CP Plot
cp_plot_title = 'CP Plot for REPL1B'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REPL1B_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot

# NIT ER Plot
er_nit_plot_title = 'Nit ER Plot for REPL1B'  # title for Nit ER plot
er_nit_plot_xlabel = 'Date'        # xaxis name for Nit ER plot
er_nit_plot_ylabel = 'Nit ER (A/min)'   # yaxis name for Nit ER plot
er_nit_plot_html_file = 'REPL1B_Nit_ER-Plot.html'   # HTML filename for Nit ER plot
er_nit_plot_trace_count = 5    # no. of traces in Nit ER plot

# NIT Unif Plot
unif_nit_plot_title = 'Nit Uniformity Plot for REPL1B'  # title for Nit Unif plot
unif_nit_plot_xlabel = 'Date'      # xaxis name for Nit Unif plot
unif_nit_plot_ylabel = 'Nit Unif (%)'    # yaxis name for Nit Unif plot
unif_nit_plot_html_file = 'REPL1B_Nit_Unif-Plot.html'   # HTML filename for Nit Unif plot
unif_nit_plot_trace_count = 3    # no. of traces in Nit Unif plot

# POLY ER Plot
er_poly_plot_title = 'Poly ER Plot for REPL1B'  # title for Poly ER plot
er_poly_plot_xlabel = 'Date'        # xaxis name for Poly ER plot
er_poly_plot_ylabel = 'Poly ER (A/min)'   # yaxis name for Poly ER plot
er_poly_plot_html_file = 'REPL1B_Poly_ER-Plot.html'   # HTML filename for Poly ER plot
er_poly_plot_trace_count = 5    # no. of traces in Poly ER plot

# POLY Unif Plot
unif_poly_plot_title = 'Poly Uniformity Plot for REPL1B'  # title for Poly Unif plot
unif_poly_plot_xlabel = 'Date'      # xaxis name for Poly Unif plot
unif_poly_plot_ylabel = 'Poly Unif (%)'    # yaxis name for Poly Unif plot
unif_poly_plot_html_file = 'REPL1B_Poly_Unif-Plot.html'   # HTML filename for Poly Unif plot
unif_poly_plot_trace_count = 3    # no. of traces in Poly Unif plot

# Sheet names
sht_name_cp = 'REPL1B-CP'
sht_name_er_nit = 'REPL1B-ERNit'
sht_name_er_poly = 'REPL1B-ERPoly'

# Columns
sht_cp_columns =  ["Date (MM/DD/YYYY)", "delta CP", "USL", "Remarks"]
sht_er_nit_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]
sht_er_poly_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"


excel_file_directory = "I:\\github_repos\\AutoPlot\\examples\\dry_etch\\CNT01_Ch_B_QC_LOG_BOOK\\CNT01_Ch_B_QC_LOG_BOOK.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_Ch_B_QC_LOG_BOOK\\CNT01_Ch_B_QC_LOG_BOOK.xlsm"
