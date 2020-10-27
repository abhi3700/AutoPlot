#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color for any plot
line_color_2 = '#80deea'      # line (trace1) color for any plot
line_color_3 = '#a5d6a7'      # line (trace2) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_color_2 = '#80deea'      # line (trace1) color for any plot
marker_color_3 = '#a5d6a7'      # line (trace2) color for any plot
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

sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.16u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_barc_pr_teos_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]
sht_er_sin_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]


# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"
date_format_contour = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_barc_range = 'F7:N7'
y_coord_barc_range = 'F8:N8'
x_coord_pr_range = 'F9:N9'
y_coord_pr_range = 'F10:N10'
x_coord_teos_range = 'F11:N11'
y_coord_teos_range = 'F12:N12'
x_coord_sin_range = 'E7:Q7'
y_coord_sin_range = 'E8:Q8'

# Skiprows
skiprows_cp = 8
skiprows_barc_pr_teos = 12
skiprows_nit = 8