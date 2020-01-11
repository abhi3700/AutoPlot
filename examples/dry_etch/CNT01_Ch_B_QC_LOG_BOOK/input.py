#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot

# dtype1 Plot
dtype1_plot_title = 'dtype1 Plot for ch_1'  # title for dtype1 plot
dtype1_plot_xlabel = 'Date'   # xaxis name for dtype1 plot
dtype1_plot_ylabel = 'delta dtype1 (no.s)'     # yaxis name for dtype1 plot
dtype1_plot_html_file = 'ch_1_dtype1-Plot.html'   # HTML filename for dtype1 plot
dtype1_plot_trace_count = 2    # no. of traces in dtype1 plot

# dtype2 ER Plot
er_dtype2_plot_title = 'dtype2 ER Plot for ch_1'  # title for dtype2 ER plot
er_dtype2_plot_xlabel = 'Date'        # xaxis name for dtype2 ER plot
er_dtype2_plot_ylabel = 'dtype2 ER (A/min)'   # yaxis name for dtype2 ER plot
er_dtype2_plot_html_file = 'ch_1_dtype2_ER-Plot.html'   # HTML filename for dtype2 ER plot
er_dtype2_plot_trace_count = 5    # no. of traces in dtype2 ER plot

# dtype2 Unif Plot
unif_dtype2_plot_title = 'dtype2 Uniformity Plot for ch_1'  # title for dtype2 Unif plot
unif_dtype2_plot_xlabel = 'Date'      # xaxis name for dtype2 Unif plot
unif_dtype2_plot_ylabel = 'dtype2 Unif (%)'    # yaxis name for dtype2 Unif plot
unif_dtype2_plot_html_file = 'ch_1_dtype2_Unif-Plot.html'   # HTML filename for dtype2 Unif plot
unif_dtype2_plot_trace_count = 3    # no. of traces in dtype2 Unif plot

# Sheet names
sht_name_dtype1 = 'dtype_1'
sht_name_er_dtype2 = 'dtype_2'

# Columns
sht_dtype1_columns = ["Date (MM/DD/YYYY)", "delta dtype1", "USL", "UCL", "Remarks"]
sht_er_dtype2_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"

# Skiprows
skiprows_dtype1 = 8
skiprows_dtype2 = 10