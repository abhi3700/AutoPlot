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
cp_plot_title = 'CP Plot for REOX1C'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REOX1C_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot

# ARC ER Plot
er_arc_plot_title = 'ARC ER Plot for REOX1C'  # title for ER plot
er_arc_plot_xlabel = 'Date'        # xaxis name for ER plot
er_arc_plot_ylabel = 'ARC ER (A/min)'   # yaxis name for ER plot
er_arc_plot_html_file = 'REOX1C_ARC_ER-Plot.html'   # HTML filename for ER plot
er_arc_plot_trace_count = 5    # no. of traces in ER plot

# ARC Unif Plot
unif_arc_plot_title = 'ARC Uniformity Plot for REOX1C'  # title for Unif plot
unif_arc_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_arc_plot_ylabel = 'ARC Unif (%)'    # yaxis name for Unif plot
unif_arc_plot_html_file = 'REOX1C_ARC_Unif-Plot.html'   # HTML filename for Unif plot
unif_arc_plot_trace_count = 2    # no. of traces in Unif plot

# TEOS ER Plot
er_teos_plot_title = 'TEOS ER Plot for REOX1C'  # title for ER plot
er_teos_plot_xlabel = 'Date'        # xaxis name for ER plot
er_teos_plot_ylabel = 'TEOS ER (A/min)'   # yaxis name for ER plot
er_teos_plot_html_file = 'REOX1C_TEOS_ER-Plot.html'   # HTML filename for ER plot
er_teos_plot_trace_count = 5    # no. of traces in ER plot

# TEOS Unif Plot
unif_teos_plot_title = 'TEOS Uniformity Plot for REOX1C'  # title for Unif plot
unif_teos_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_teos_plot_ylabel = 'TEOS Unif (%)'    # yaxis name for Unif plot
unif_teos_plot_html_file = 'REOX1C_TEOS_Unif-Plot.html'   # HTML filename for Unif plot
unif_teos_plot_trace_count = 2    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'REOX1C-CP'
sht_name_er = 'REOX1C-ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "delta-CP 0.16u", "delta-CP 0.5u", "delta-CP AC", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]   # deleted "% Uni UCL" column
 
# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"
date_format_contour = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_arc_range = 'D9:L9'
y_coord_arc_range = 'D10:L10'
x_coord_teos_range = 'D11:L11'
y_coord_teos_range = 'D12:L12'

# Skiprows
skiprows_cp = 9
skiprows_er = 12