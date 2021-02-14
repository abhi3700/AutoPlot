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
cp_plot_title = 'CP Plot for REOX1B'  # title for CP plot
cp_plot_xlabel = 'Date (DD/MM/YYYY)'   # xaxis name for CP plot
cp_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REOX1B_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 3    # no. of traces in CP plot

# BPSG ER Plot
er_bpsgcs_plot_title = 'BPSG_CS ER Plot for REOX1B'  # title for ER plot
er_bpsgcs_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for ER plot
er_bpsgcs_plot_ylabel = 'BPSG_CS ER (A/min)'   # yaxis name for ER plot
er_bpsgcs_plot_html_file = 'REOX1B_BPSG_CS_ER-Plot.html'   # HTML filename for ER plot
er_bpsgcs_plot_trace_count = 5    # no. of traces in ER plot

# BPSG Unif Plot
unif_bpsgcs_plot_title = 'BPSG_CS Uniformity Plot for REOX1B'  # title for Unif plot
unif_bpsgcs_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Unif plot
unif_bpsgcs_plot_ylabel = 'BPSG_CS Unif (%)'    # yaxis name for Unif plot
unif_bpsgcs_plot_html_file = 'REOX1B_BPSG_CS_Unif-Plot.html'   # HTML filename for Unif plot
unif_bpsgcs_plot_trace_count = 2    # no. of traces in Unif plot

# SiNCS ER Plot
er_sincs_plot_title = 'SiN_CS ER Plot for REOX1B'  # title for ER plot
er_sincs_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for ER plot
er_sincs_plot_ylabel = 'SiN_CS ER (A/min)'   # yaxis name for ER plot
er_sincs_plot_html_file = 'REOX1B_SiN_CS_ER-Plot.html'   # HTML filename for ER plot
er_sincs_plot_trace_count = 5    # no. of traces in ER plot

# SiNCS Unif Plot
unif_sincs_plot_title = 'SiN_CS Uniformity Plot for REOX1B'  # title for Unif plot
unif_sincs_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Unif plot
unif_sincs_plot_ylabel = 'SiN_CS Unif (%)'    # yaxis name for Unif plot
unif_sincs_plot_html_file = 'REOX1B_SiN_CS_Unif-Plot.html'   # HTML filename for Unif plot
unif_sincs_plot_trace_count = 2    # no. of traces in Unif plot

# TEOSVIA ER Plot
er_teosvia_plot_title = 'TEOS_VIA ER Plot for REOX1B'  # title for ER plot
er_teosvia_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for ER plot
er_teosvia_plot_ylabel = 'TEOS_VIA ER (A/min)'   # yaxis name for ER plot
er_teosvia_plot_html_file = 'REOX1B_TEOS_VIA_ER-Plot.html'   # HTML filename for ER plot
er_teosvia_plot_trace_count = 5    # no. of traces in ER plot

# TEOSVIA Unif Plot
unif_teosvia_plot_title = 'TEOS_VIA Uniformity Plot for REOX1B'  # title for Unif plot
unif_teosvia_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Unif plot
unif_teosvia_plot_ylabel = 'TEOS_VIA Unif (%)'    # yaxis name for Unif plot
unif_teosvia_plot_html_file = 'REOX1B_TEOS_VIA_Unif-Plot.html'   # HTML filename for Unif plot
unif_teosvia_plot_trace_count = 3    # no. of traces in Unif plot

# ARC ER Plot
er_arc_plot_title = 'ARC ER Plot for REOX1B'  # title for ER plot
er_arc_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for ER plot
er_arc_plot_ylabel = 'ARC ER (A/min)'   # yaxis name for ER plot
er_arc_plot_html_file = 'REOX1B_ARC_ER-Plot.html'   # HTML filename for ER plot
er_arc_plot_trace_count = 5    # no. of traces in ER plot

# ARC Unif Plot
unif_arc_plot_title = 'ARC Uniformity Plot for REOX1B'  # title for Unif plot
unif_arc_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Unif plot
unif_arc_plot_ylabel = 'ARC Unif (%)'    # yaxis name for Unif plot
unif_arc_plot_html_file = 'REOX1B_ARC_Unif-Plot.html'   # HTML filename for Unif plot
unif_arc_plot_trace_count = 3    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'REOX1B-CP'
sht_name_er = 'REOX1B-ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.16u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%d-%m-%Y %H:%M:%S"
date_format_contour = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_bpsgcs_range = 'E9:M9'
y_coord_bpsgcs_range = 'D10:L10'
x_coord_sincs_range = 'E11:Q11'
y_coord_sincs_range = 'E12:Q12'
x_coord_teosvia_range = 'E13:M13'
y_coord_teosvia_range = 'E14:M14'
x_coord_arc_range = 'E15:M15'
y_coord_arc_range = 'E16:M16'

# Skiprows
skiprows_cp = 9
skiprows_er = 16