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
cp_plot_title = 'CP Plot for REPL1A'  # title for CP plot
cp_plot_xlabel = 'Date (DD/MM/YYYY)'   # xaxis name for CP plot
cp_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REPL1A_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot

# NIT ER Plot
er_nit_plot_title = 'Nit ER Plot for REPL1A'  # title for Nit ER plot
er_nit_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for Nit ER plot
er_nit_plot_ylabel = 'Nit ER (A/min)'   # yaxis name for Nit ER plot
er_nit_plot_html_file = 'REPL1A_Nit_ER-Plot.html'   # HTML filename for Nit ER plot
er_nit_plot_trace_count = 5    # no. of traces in Nit ER plot
er_nit_contour_fname = 'REPL1A_Nit_ER-Plot_Contour.html'
er_nit_contour_tname = 'NIT ER'

# NIT Unif Plot
unif_nit_plot_title = 'Nit Uniformity Plot for REPL1A'  # title for Nit Unif plot
unif_nit_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Nit Unif plot
unif_nit_plot_ylabel = 'Nit Unif (%)'    # yaxis name for Nit Unif plot
unif_nit_plot_html_file = 'REPL1A_Nit_Unif-Plot.html'   # HTML filename for Nit Unif plot
unif_nit_plot_trace_count = 3    # no. of traces in Nit Unif plot

# POLY ER Plot
er_poly_plot_title = 'Poly ER Plot for REPL1A'  # title for Poly ER plot
er_poly_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for Poly ER plot
er_poly_plot_ylabel = 'Poly ER (A/min)'   # yaxis name for Poly ER plot
er_poly_plot_html_file = 'REPL1A_Poly_ER-Plot.html'   # HTML filename for Poly ER plot
er_poly_plot_trace_count = 5    # no. of traces in Poly ER plot
er_poly_contour_fname = 'REPL1A_Poly_ER-Plot_Contour.html'
er_poly_contour_tname = 'POLY ER'

# POLY Unif Plot
unif_poly_plot_title = 'Poly Uniformity Plot for REPL1A'  # title for Poly Unif plot
unif_poly_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Poly Unif plot
unif_poly_plot_ylabel = 'Poly Unif (%)'    # yaxis name for Poly Unif plot
unif_poly_plot_html_file = 'REPL1A_Poly_Unif-Plot.html'   # HTML filename for Poly Unif plot
unif_poly_plot_trace_count = 3    # no. of traces in Poly Unif plot

# Sheet names
sht_name_cp = 'REPL1A-CP'
sht_name_er_nit = 'REPL1A-ERNit'
sht_name_er_poly = 'REPL1A-ERPoly'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.16u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_nit_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]
sht_er_nit_cl_columns = ["Date (MM/DD/YYYY)", "Site", "site_1", "site_2", "site_3", "site_4", "site_5", "site_6", "site_7", "site_8", "site_9", "site_10", "site_11", "site_12", "site_13", "Etch Rate (A/Min)", "Result"]
sht_er_poly_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]
sht_er_poly_cl_columns = ["Date (MM/DD/YYYY)", "Site", "site_1", "site_2", "site_3", "site_4", "site_5", "site_6", "site_7", "site_8", "site_9", "site_10", "site_11", "site_12", "site_13", "site_14", "site_15", "site_16", "site_17", "Etch Rate (A/Min)", "Result"]
N_cl = 30

# Date formatter
date_format = "%d-%m-%Y %H:%M:%S"
date_format_contour = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_nit_range = 'D9:P9'
y_coord_nit_range = 'D10:P10'
x_coord_poly_range = 'D9:T9'
y_coord_poly_range = 'D10:T10'

# Skiprows
skiprows_cp = 9
skiprows_nit = 10
skiprows_poly = 10