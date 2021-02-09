#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color for any plot
line_color_2 = '#80deea'      # line (trace1) color for any plot
line_color_3 = '#a5d6a7'      # line (trace2) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot

# CP Plot
cp_plot_title = 'CP Plot for RESP1A'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'RESP1A_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 3    # no. of traces in CP plot

# SiN-1st ER Plot
er_sin_1st_plot_title = 'SiN-1st ER Plot for RESP1A'  # title for ER plot
er_sin_1st_plot_xlabel = 'Date'        # xaxis name for ER plot
er_sin_1st_plot_ylabel = 'SiN-1st ER (A/min)'   # yaxis name for ER plot
er_sin_1st_plot_html_file = 'RESP1A_SiN_1st_ER-Plot.html'   # HTML filename for ER plot
er_sin_1st_plot_trace_count = 5    # no. of traces in Nit ER plot

# SiN-1st Unif Plot
unif_sin_1st_plot_title = 'SiN-1st Uniformity Plot for RESP1A'  # title for Unif plot
unif_sin_1st_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_sin_1st_plot_ylabel = 'SiN-1st Unif (%)'    # yaxis name for Unif plot
unif_sin_1st_plot_html_file = 'RESP1A_SiN_1st_Unif-Plot.html'   # HTML filename for Unif plot
unif_sin_1st_plot_trace_count = 3    # no. of traces in Nit Unif plot

# SiN-2nd ER Plot
er_sin_2nd_plot_title = 'SiN-2nd ER Plot for RESP1A'  # title for ER plot
er_sin_2nd_plot_xlabel = 'Date'        # xaxis name for ER plot
er_sin_2nd_plot_ylabel = 'SiN-2nd ER (A/min)'   # yaxis name for ER plot
er_sin_2nd_plot_html_file = 'RESP1A_SiN_2nd_ER-Plot.html'   # HTML filename for ER plot
er_sin_2nd_plot_trace_count = 5    # no. of traces in Nit ER plot

# SiN-2nd Unif Plot
unif_sin_2nd_plot_title = 'SiN-2nd Uniformity Plot for RESP1A'  # title for Unif plot
unif_sin_2nd_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_sin_2nd_plot_ylabel = 'SiN-2nd Unif (%)'    # yaxis name for Unif plot
unif_sin_2nd_plot_html_file = 'RESP1A_SiN_2nd_Unif-Plot.html'   # HTML filename for Unif plot
unif_sin_2nd_plot_trace_count = 3    # no. of traces in Nit Unif plot

# TEOS-1st ER Plot
er_teos_1st_plot_title = 'TEOS-1st ER Plot for RESP1A'  # title for ER plot
er_teos_1st_plot_xlabel = 'Date'        # xaxis name for ER plot
er_teos_1st_plot_ylabel = 'TEOS-1st ER (A/min)'   # yaxis name for ER plot
er_teos_1st_plot_html_file = 'RESP1A_TEOS_1st_ER-Plot.html'   # HTML filename for ER plot
er_teos_1st_plot_trace_count = 5    # no. of traces in Nit ER plot

# TEOS-1st Unif Plot
unif_teos_1st_plot_title = 'TEOS-1st Uniformity Plot for RESP1A'  # title for Unif plot
unif_teos_1st_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_teos_1st_plot_ylabel = 'TEOS-1st Unif (%)'    # yaxis name for Unif plot
unif_teos_1st_plot_html_file = 'RESP1A_TEOS_1st_Unif-Plot.html'   # HTML filename for Unif plot
unif_teos_1st_plot_trace_count = 2    # no. of traces in Nit Unif plot

# TEOS-2nd ER Plot
er_teos_2nd_plot_title = 'TEOS-2nd ER Plot for RESP1A'  # title for ER plot
er_teos_2nd_plot_xlabel = 'Date'        # xaxis name for ER plot
er_teos_2nd_plot_ylabel = 'TEOS-2nd ER (A/min)'   # yaxis name for ER plot
er_teos_2nd_plot_html_file = 'RESP1A_TEOS_2nd_ER-Plot.html'   # HTML filename for ER plot
er_teos_2nd_plot_trace_count = 5    # no. of traces in Nit ER plot

# TEOS-2nd Unif Plot
unif_teos_2nd_plot_title = 'TEOS-2nd Uniformity Plot for RESP1A'  # title for Unif plot
unif_teos_2nd_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_teos_2nd_plot_ylabel = 'TEOS-2nd Unif (%)'    # yaxis name for Unif plot
unif_teos_2nd_plot_html_file = 'RESP1A_TEOS_2nd_Unif-Plot.html'   # HTML filename for Unif plot
unif_teos_2nd_plot_trace_count = 2    # no. of traces in Nit Unif plot

# Sheet names
sht_name_cp = 'RESP1A-CP'
sht_name_er = 'RESP1A-ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP > 0.16u", "Delta CP > 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer-Step", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_sin_range = 'E15:M15'
y_coord_sin_range = 'E16:M16'
x_coord_teos_range = 'E17:M17'
y_coord_teos_range = 'E18:M18'

# skiprows
skiprows_cp = 9
skiprows_er = 18