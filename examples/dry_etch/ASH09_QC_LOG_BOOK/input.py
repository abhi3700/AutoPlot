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
cp_plot_title = 'CP Plot for ASFE1'  # title for CP plot
cp_plot_xlabel = 'Date (DD/MM/YYYY)'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'ASFE1_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 3    # no. of traces in CP plot

# PR ER Plot
er_plot_title = 'ER Plot for ASFE1'  # title for ER plot
er_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for ER plot
er_plot_ylabel = 'Etch Rate (A/min)'   # yaxis name for ER plot
er_plot_html_file = 'ASFE1_ER-Plot.html'   # HTML filename for ER plot
er_plot_trace_count = 4    # no. of traces in ER plot

# PR Unif Plot
unif_plot_title = 'Uniformity Plot for ASFE1'  # title for UNIF plot
unif_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Unif plot
unif_plot_ylabel = 'Uniformity (%)'    # yaxis name for Unif plot
unif_plot_html_file = 'ASFE1_Unif-Plot.html'   # HTML filename for Unif plot
unif_plot_trace_count = 2    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'ASFE1-CP'
sht_name_er = 'ASFE1-ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.16u", "Delta CP 0.5u", "Delta CP AC", "Remarks", "USL", "UCL", ]
sht_er_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni UCL", "% Uni USL"]

# Date formatting for Plotly charts
date_format = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_pr_range = 'D9:L9'
y_coord_pr_range = 'D10:L10'

# skiprows
skiprows_cp = 10
skiprows_pr = 10