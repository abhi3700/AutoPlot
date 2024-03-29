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

# ChA CP Plot
cp_cha_plot_title = 'CP Plot for REML1A'  # title for CP plot
cp_cha_plot_xlabel = 'Date (DD/MM/YYYY)'   # xaxis name for CP plot
cp_cha_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
cp_cha_plot_html_file = 'REML1A_CP-Plot.html'   # HTML filename for CP plot
cp_cha_plot_trace_count = 2    # no. of traces in CP plot

# ChC CP Plot
cp_chc_plot_title = 'CP Plot for REML1C'  # title for CP plot
cp_chc_plot_xlabel = 'Date (DD/MM/YYYY)'   # xaxis name for CP plot
cp_chc_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
cp_chc_plot_html_file = 'REML1C_CP-Plot.html'   # HTML filename for CP plot
cp_chc_plot_trace_count = 2    # no. of traces in CP plot

# ChA PR ER Plot
er_cha_pr_plot_title = 'PR ER Plot for REML1A'  # title for ER plot
er_cha_pr_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for ER plot
er_cha_pr_plot_ylabel = 'PR ER (A/min)'   # yaxis name for ER plot
er_cha_pr_plot_html_file = 'REML1A_PR_ER-Plot.html'   # HTML filename for ER plot
er_cha_pr_plot_trace_count = 5    # no. of traces in Nit ER plot

# ChA PR Unif Plot
unif_cha_pr_plot_title = 'PR Uniformity Plot for REML1A'  # title for Unif plot
unif_cha_pr_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Unif plot
unif_cha_pr_plot_ylabel = 'PR Unif (%)'    # yaxis name for Unif plot
unif_cha_pr_plot_html_file = 'REML1A_PR_Unif-Plot.html'   # HTML filename for Unif plot
unif_cha_pr_plot_trace_count = 3    # no. of traces in Nit Unif plot

# ChC PR ER Plot
er_chc_pr_plot_title = 'PR ER Plot for REML1C'  # title for ER plot
er_chc_pr_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for ER plot
er_chc_pr_plot_ylabel = 'PR ER (A/min)'   # yaxis name for ER plot
er_chc_pr_plot_html_file = 'REML1C_PR_ER-Plot.html'   # HTML filename for ER plot
er_chc_pr_plot_trace_count = 5    # no. of traces in ER plot

# ChC PR Unif Plot
unif_chc_pr_plot_title = 'PR Uniformity Plot for REML1C'  # title for Unif plot
unif_chc_pr_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Unif plot
unif_chc_pr_plot_ylabel = 'PR Unif (%)'    # yaxis name for Unif plot
unif_chc_pr_plot_html_file = 'REML1C_PR_Unif-Plot.html'   # HTML filename for Unif plot
unif_chc_pr_plot_trace_count = 3    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'REML1-CP'
sht_name_er_cha_pr = 'PR Ch A ER'
sht_name_er_chc_pr = 'PR Ch C ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "Chamber", "Delta CP 0.2u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_reml1a_pr_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]
sht_er_reml1c_pr_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_pr_a_range = 'D9:L9'
y_coord_pr_a_range = 'D10:L10'
x_coord_pr_c_range = 'D9:L9'
y_coord_pr_c_range = 'D10:L10'


# skiprows
skiprows_cp = 8
skiprows_pr_a = 10
skiprows_pr_c = 10