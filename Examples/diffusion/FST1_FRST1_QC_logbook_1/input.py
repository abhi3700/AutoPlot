#=================================Global variables====================================================================#

# Colors (standard)
line_color = '#3f51b5'      # line (trace1) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot

# Thickness Plot
thick_plot_trace1_name = 'Top'        # Trace name for Top thickness
thick_plot_trace2_name = 'Center'        # Trace name for Top thickness
thick_plot_trace3_name = 'Bottom'        # Trace name for Top thickness
thick_plot_trace1_mode_top = 'lines+markers'    # Chart mode for Top thickness -> 'lines+markers' or 'lines' or 'markers'
thick_plot_trace2_mode_cnt = 'lines+markers'    # Chart mode for Center thickness -> 'lines+markers' or 'lines' or 'markers'
thick_plot_trace3_mode_btm = 'lines+markers'    # Chart mode for BTM thickness -> 'lines+markers' or 'lines' or 'markers'
thick_plot_trace1_line_color_top = '#3f51b5'    # Line color for Top thickness
thick_plot_trace2_line_color_cnt = '#3f51b5'    # Line color for Center thickness
thick_plot_trace3_line_color_btm = '#3f51b5'    # Line color for BTM thickness
thick_plot_trace1_marker_color_top = '#43a047'    # Line color for Top thickness
thick_plot_trace2_marker_color_cnt = '#43a047'    # Line color for Center thickness
thick_plot_trace3_marker_color_btm = '#43a047'    # Line color for BTM thickness
thick_plot_trace1_marker_border_color_top = '#ffffff'    # Line color for Top thickness
thick_plot_trace2_marker_border_color_cnt = '#ffffff'    # Line color for Center thickness
thick_plot_trace3_marker_border_color_btm = '#ffffff'    # Line color for BTM thickness
thick_plot_title = 'Thickness Plot for FRST1'  # title for Thickness plot
thick_plot_xlabel = 'Date (DD/MM/YYYY)'        # xaxis name for Thickness plot
thick_plot_ylabel = 'Thickness (A)'   # yaxis name for Thickness plot
thick_plot_html_file = 'FRST1_Thickness-Plot.html'   # HTML filename for Thickness plot
thick_plot_trace_count = 7    # no. of traces in Thickness plot

# Uniformity Plot
unif_plot_trace1_name = 'TOP'        # Trace name for Top thickness
unif_plot_trace2_name = 'CNT'        # Trace name for Top thickness
unif_plot_trace3_name = 'BTM'        # Trace name for Top thickness
unif_plot_trace1_mode_top = 'lines+markers'    # Chart mode for Top thickness -> 'lines+markers' or 'lines' or 'markers'
unif_plot_trace2_mode_cnt = 'lines+markers'    # Chart mode for Center thickness -> 'lines+markers' or 'lines' or 'markers'
unif_plot_trace3_mode_btm = 'lines+markers'    # Chart mode for BTM thickness -> 'lines+markers' or 'lines' or 'markers'
unif_plot_trace1_line_color_top = '#3f51b5'    # Line color for Top thickness
unif_plot_trace2_line_color_cnt = '#3f51b5'    # Line color for Center thickness
unif_plot_trace3_line_color_btm = '#3f51b5'    # Line color for BTM thickness
unif_plot_trace1_marker_color_top = '#43a047'    # Line color for Top thickness
unif_plot_trace2_marker_color_cnt = '#43a047'    # Line color for Center thickness
unif_plot_trace3_marker_color_btm = '#43a047'    # Line color for BTM thickness
unif_plot_trace1_marker_border_color_top = '#ffffff'    # Line color for Top thickness
unif_plot_trace2_marker_border_color_cnt = '#ffffff'    # Line color for Center thickness
unif_plot_trace3_marker_border_color_btm = '#ffffff'    # Line color for BTM thickness
unif_plot_title = 'Uniformity Plot for FRST1'  # title for Uniformity plot
unif_plot_xlabel = 'Date (DD/MM/YYYY)'      # xaxis name for Uniformity plot
unif_plot_ylabel = 'Uniformity (%)'    # yaxis name for Uniformity plot
unif_plot_html_file = 'FRST1_Uniformity-Plot.html'   # HTML filename for Uniformity plot
unif_plot_trace_count = 4    # no. of traces in Uniformity plot

# CP Plot
cp_plot_trace1_name = 'TOP'        # Trace name for Top thickness
cp_plot_trace2_name = 'CNT'        # Trace name for Top thickness
cp_plot_trace3_name = 'BTM'        # Trace name for Top thickness
cp_plot_trace1_mode_top = 'lines+markers'    # Chart mode for Top thickness -> 'lines+markers' or 'lines' or 'markers'
cp_plot_trace2_mode_cnt = 'lines+markers'    # Chart mode for Center thickness -> 'lines+markers' or 'lines' or 'markers'
cp_plot_trace3_mode_btm = 'lines+markers'    # Chart mode for BTM thickness -> 'lines+markers' or 'lines' or 'markers'
cp_plot_trace1_line_color_top = '#3f51b5'    # Line color for Top thickness
cp_plot_trace2_line_color_cnt = '#3f51b5'    # Line color for Center thickness
cp_plot_trace3_line_color_btm = '#3f51b5'    # Line color for BTM thickness
cp_plot_trace1_marker_color_top = '#43a047'    # Line color for Top thickness
cp_plot_trace2_marker_color_cnt = '#43a047'    # Line color for Center thickness
cp_plot_trace3_marker_color_btm = '#43a047'    # Line color for BTM thickness
cp_plot_trace1_marker_border_color_top = '#ffffff'    # Line color for Top thickness
cp_plot_trace2_marker_border_color_cnt = '#ffffff'    # Line color for Center thickness
cp_plot_trace3_marker_border_color_btm = '#ffffff'    # Line color for BTM thickness
cp_plot_title = 'CP Plot for FRST1'  # title for CP plot
cp_plot_xlabel = 'Date (DD/MM/YYYY)'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'FRST1_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 4    # no. of traces in CP plot

# Columns
sht_thick_columns = ["Date", "Top Avg", "Center Avg", "BTM Avg", "Thick LCL", "Thick UCL", "Thick LSL", "Thick USL", "Remarks"]
sht_unif_columns = ["Date", "Unif TOP", "Unif CNT", "Unif BTM", "Unif UCL", "Remarks"]
sht_cp_columns = ["Date", "CP TOP", "CP CNT", "CP BTM", "CP UCL", "Remarks"]

# Sheet name
sht_name = 'AA-DENS QC'     # same sheet for CP, Thickness, Uniformity

# Date formatting for Plotly charts
date_format = "%d-%m-%Y %H:%M:%S"

# Excel file directory
excel_file_directory = "I:\\github_repos\\AutoPlot\\examples\\diffusion\\FST1_FRST1_QC_logbook_1\\FST1_FRST1_QC_logbook_1.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_QC_LOG_BOOK_Ch_A_macro\\CNT01_Ch_A_QC_LOG_BOOK.xlsm"
