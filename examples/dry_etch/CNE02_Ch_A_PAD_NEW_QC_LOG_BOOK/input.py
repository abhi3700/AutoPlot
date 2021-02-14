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
cp_plot_title = 'CP Plot for REOX1A'  # title for CP plot
cp_plot_xlabel = 'Date (MM/DD/YYYY)'   # xaxis name for CP plot
cp_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REOX1A_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot

# SiN ER Plot
er_sin_plot_title = 'SiN ER Plot for REOX1A'  # title for ER plot
er_sin_plot_xlabel = 'Date (MM/DD/YYYY)'        # xaxis name for ER plot
er_sin_plot_ylabel = 'SiN ER (A/min)'   # yaxis name for ER plot
er_sin_plot_html_file = 'REOX1A_SiN_ER-Plot.html'   # HTML filename for ER plot
er_sin_plot_trace_count = 5    # no. of traces in ER plot

# SiN Unif Plot
unif_sin_plot_title = 'SiN Uniformity Plot for REOX1A'  # title for Unif plot
unif_sin_plot_xlabel = 'Date (MM/DD/YYYY)'      # xaxis name for Unif plot
unif_sin_plot_ylabel = 'SiN Unif (%)'    # yaxis name for Unif plot
unif_sin_plot_html_file = 'REOX1A_SiN_Unif-Plot.html'   # HTML filename for Unif plot
unif_sin_plot_trace_count = 3    # no. of traces in Unif plot

# TEOS ER Plot
er_teos_plot_title = 'TEOS ER Plot for REOX1A'  # title for ER plot
er_teos_plot_xlabel = 'Date (MM/DD/YYYY)'        # xaxis name for ER plot
er_teos_plot_ylabel = 'TEOS ER (A/min)'   # yaxis name for ER plot
er_teos_plot_html_file = 'REOX1A_TEOS_ER-Plot.html'   # HTML filename for ER plot
er_teos_plot_trace_count = 5    # no. of traces in ER plot

# TEOS Unif Plot
unif_teos_plot_title = 'TEOS Uniformity Plot for REOX1A'  # title for Unif plot
unif_teos_plot_xlabel = 'Date (MM/DD/YYYY)'      # xaxis name for Unif plot
unif_teos_plot_ylabel = 'TEOS Unif (%)'    # yaxis name for Unif plot
unif_teos_plot_html_file = 'REOX1A_TEOS_Unif-Plot.html'   # HTML filename for Unif plot
unif_teos_plot_trace_count = 3    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'REOX1A-CP'
sht_name_er = 'REOX1A-ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.2u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"
date_format_contour = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_sin_range = 'D9:L9'
y_coord_sin_range = 'D10:L10'
x_coord_teos_range = 'D11:L11'
y_coord_teos_range = 'D12:L12'

# Skiprows
skiprows_cp = 9
skiprows_er = 12