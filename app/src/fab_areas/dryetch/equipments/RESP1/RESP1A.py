# Import packages
import pandas as pd
import plotly as py
import plotly.graph_objs as go
# import datetime as dt


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
cp_plot_title = 'CP Plot for RESP1A'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
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
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.16u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer-Step", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_sin_range = 'E15:M15'
y_coord_sin_range = 'E16:M16'
x_coord_teos_range = 'E17:M17'
y_coord_teos_range = 'E18:M18'

# skiprows
skiprows_cp = 9
skiprows_er = 18

# =================================================================================================================================================================
excel_file_directory = "F:\\Coding\\github_repos\\AutoPlot\\examples\\dry_etch\\UNT02_Ch_A_QC_LOG_BOOK\\UNT02_Ch_A_QC_LOG_BOOK.xlsm"

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": Date formatter to format the excel date (issue: one date less in plotly chart) as "%m-%d-%Y %H:%M:%S"
"x": datetime list
"return": formatted datetime list
"""
def date_formatter(x):
    x_fmt = []
    for a in x:
        a = a.strftime(date_format)
        x_fmt.append(a)
    return x_fmt


#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP 0.16u (y-axis) for CP Chart
"y2": Delta-CP 0.5u (y-axis) for CP Chart
"y3": Delta-CP AC (y-axis) for CP Chart
"y4": USL (y-axis) for CP Chart
"y5": UCL (y-axis) for CP Chart
"""
def draw_plotly_resp1a_cp_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'Delta CP 0.16u',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'Delta CP 0.5u',
            mode = 'lines+markers',
            line = dict(
                    color = line_color_2,
                    width = 2),
            marker = dict(
                    color = marker_color_2,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'Delta CP AC',
            mode = 'lines+markers',
            line = dict(
                    color = line_color_3,
                    width = 2),
            marker = dict(
                    color = marker_color_3,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace4 = go.Scatter(
            x = x,
            y = y4,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace5 = go.Scatter(
            x = x,
            y = y5,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4, trace5]
    layout = dict(
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= cp_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_resp1a_er_sin_1st_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'ER',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'LSL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace4 = go.Scatter(
            x = x,
            y = y4,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    trace5 = go.Scatter(
            x = x,
            y = y5,
            name = 'LCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4, trace5]
    layout = dict(
            title = er_sin_1st_plot_title,
            xaxis = dict(title= er_sin_1st_plot_xlabel),
            yaxis = dict(title= er_sin_1st_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_sin_1st_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_sin_1st_plot(x, y1, y2, y3, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'Unif',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3]
    layout = dict(
            title = unif_sin_1st_plot_title,
            xaxis = dict(title= unif_sin_1st_plot_xlabel),
            yaxis = dict(title= unif_sin_1st_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_sin_1st_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_resp1a_er_sin_2nd_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'ER',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'LSL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace4 = go.Scatter(
            x = x,
            y = y4,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    trace5 = go.Scatter(
            x = x,
            y = y5,
            name = 'LCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4, trace5]
    layout = dict(
            title = er_sin_2nd_plot_title,
            xaxis = dict(title= er_sin_2nd_plot_xlabel),
            yaxis = dict(title= er_sin_2nd_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_sin_2nd_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_sin_2nd_plot(x, y1, y2, y3, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'Unif',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3]
    layout = dict(
            title = unif_sin_2nd_plot_title,
            xaxis = dict(title= unif_sin_2nd_plot_xlabel),
            yaxis = dict(title= unif_sin_2nd_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_sin_2nd_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_resp1a_er_teos_1st_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'ER',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'LSL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace4 = go.Scatter(
            x = x,
            y = y4,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    trace5 = go.Scatter(
            x = x,
            y = y5,
            name = 'LCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4, trace5]
    layout = dict(
            title = er_teos_1st_plot_title,
            xaxis = dict(title= er_teos_1st_plot_xlabel),
            yaxis = dict(title= er_teos_1st_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_teos_1st_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_teos_1st_plot(x, y1, y2, y3, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'Unif',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3]
    layout = dict(
            title = unif_teos_1st_plot_title,
            xaxis = dict(title= unif_teos_1st_plot_xlabel),
            yaxis = dict(title= unif_teos_1st_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_teos_1st_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_resp1a_er_teos_2nd_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'ER',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'LSL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace4 = go.Scatter(
            x = x,
            y = y4,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    trace5 = go.Scatter(
            x = x,
            y = y5,
            name = 'LCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4, trace5]
    layout = dict(
            title = er_teos_2nd_plot_title,
            xaxis = dict(title= er_teos_2nd_plot_xlabel),
            yaxis = dict(title= er_teos_2nd_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_teos_2nd_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_teos_2nd_plot(x, y1, y2, y3, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'Unif',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3]
    layout = dict(
            title = unif_teos_2nd_plot_title,
            xaxis = dict(title= unif_teos_2nd_plot_xlabel),
            yaxis = dict(title= unif_teos_2nd_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_teos_2nd_plot_html_file)
    return fig

#====================================================================================================================================================================
#####################################################################################################################################################################
# def init():
#     # Initialize the workbook
#     # wb = xw.Book.caller()
#     wb = xw.Book('UNT02_Ch_A_QC_LOG_BOOK.xlsm')
#     # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

#     #****************************************************************************************************************************************************************
#     # Define sheets
#     sht_resp1a_cp = wb.sheets[sht_name_cp]
#     sht_resp1a_er = wb.sheets[sht_name_er]
#     sht_run = wb.sheets['RUN_code']     # for testing purpose
#     #****************************************************************************************************************************************************************
#     x_coord_sin = sht_resp1a_er.range(x_coord_sin_range).value
#     y_coord_sin = sht_resp1a_er.range(y_coord_sin_range).value
#     x_coord_teos = sht_resp1a_er.range(x_coord_teos_range).value
#     y_coord_teos = sht_resp1a_er.range(y_coord_teos_range).value
#     #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
#     excel_file = pd.ExcelFile(excel_file_directory)

#     return wb, sht_resp1a_cp, sht_resp1a_er, sht_run, x_coord_sin, y_coord_sin, x_coord_teos, y_coord_teos, excel_file

excel_file = pd.ExcelFile(excel_file_directory)

def resp1a_cp_chart():
    # Fetch Dataframe for CP
    df_resp1a_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_cp = df_resp1a_cp[sht_cp_columns]        # The final dataframe with required columns
    df_resp1a_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_resp1a_cp['Delta CP 0.5u'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_resp1a_cp['Delta CP AC'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_resp1a_cp = df_resp1a_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_run.range('A25').options(index=False).value = df_resp1a_cp        # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_resp1a_cp_date = df_resp1a_cp["Date (MM/DD/YYYY)"]
    df_resp1a_cp_delta_cp_1 = df_resp1a_cp["Delta CP 0.16u"]
    df_resp1a_cp_delta_cp_2 = df_resp1a_cp["Delta CP 0.5u"]
    df_resp1a_cp_delta_cp_3 = df_resp1a_cp["Delta CP AC"]
    df_resp1a_cp_usl = df_resp1a_cp["USL"]
    df_resp1a_cp_ucl = df_resp1a_cp["UCL"]
    df_resp1a_cp_remarks = df_resp1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    return draw_plotly_resp1a_cp_plot(
        x = date_formatter(df_resp1a_cp_date), 
        y1 = df_resp1a_cp_delta_cp_1, 
        y2 = df_resp1a_cp_delta_cp_2, 
        y3 = df_resp1a_cp_delta_cp_3, 
        y4 = df_resp1a_cp_usl, 
        y5 = df_resp1a_cp_ucl,
        remarks = df_resp1a_cp_remarks
    )

def resp1a_sin1st_er_chart():
    # Fetch Dataframes for SiN-1st
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_sin_1st = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'SiN-1Step']
    df_resp1a_er_sin_1st = df_resp1a_er_sin_1st.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_sin_1st

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-1st ER & Unif PLot
    df_resp1a_er_sin_1st_date = df_resp1a_er_sin_1st["Date (MM/DD/YYYY)"]
    df_resp1a_er_sin_1st_er = df_resp1a_er_sin_1st["Etch Rate (A/Min)"]
    df_resp1a_er_sin_1st_usl = df_resp1a_er_sin_1st["USL"]
    df_resp1a_er_sin_1st_lsl = df_resp1a_er_sin_1st["LSL"]
    df_resp1a_er_sin_1st_ucl = df_resp1a_er_sin_1st["UCL"]
    df_resp1a_er_sin_1st_lcl = df_resp1a_er_sin_1st["LCL"]
    df_resp1a_er_sin_1st_remarks = df_resp1a_er_sin_1st["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-1st ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1a_er_sin_1st_plot(
        x = date_formatter(df_resp1a_er_sin_1st_date), 
        y1 = df_resp1a_er_sin_1st_er,
        y2 = df_resp1a_er_sin_1st_usl, 
        y3 = df_resp1a_er_sin_1st_lsl,
        y4 = df_resp1a_er_sin_1st_ucl,
        y5 = df_resp1a_er_sin_1st_lcl,
        remarks = df_resp1a_er_sin_1st_remarks
    )

def resp1a_sin1st_unif_chart():
    # Fetch Dataframes for SiN-1st
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_sin_1st = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'SiN-1Step']
    df_resp1a_er_sin_1st = df_resp1a_er_sin_1st.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_sin_1st

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-1st ER & Unif PLot
    df_resp1a_er_sin_1st_date = df_resp1a_er_sin_1st["Date (MM/DD/YYYY)"]
    df_resp1a_er_sin_1st_unif = df_resp1a_er_sin_1st["% Uni"]
    df_resp1a_er_sin_1st_unif_usl = df_resp1a_er_sin_1st["% Uni USL"]
    df_resp1a_er_sin_1st_unif_ucl = df_resp1a_er_sin_1st["% Uni UCL"]
    df_resp1a_er_sin_1st_remarks = df_resp1a_er_sin_1st["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-1st Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1a_unif_sin_1st_plot(
        x = date_formatter(df_resp1a_er_sin_1st_date), 
        y1 = df_resp1a_er_sin_1st_unif, 
        y2 = df_resp1a_er_sin_1st_unif_usl,
        y3 = df_resp1a_er_sin_1st_unif_ucl,
        remarks = df_resp1a_er_sin_1st_remarks
    )


def resp1a_sin2nd_er_chart():
    # Fetch Dataframes for SiN-2nd
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_sin_2nd = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'SiN-2Step']
    df_resp1a_er_sin_2nd = df_resp1a_er_sin_2nd.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_sin_2nd

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-2nd ER & Unif PLot
    df_resp1a_er_sin_2nd_date = df_resp1a_er_sin_2nd["Date (MM/DD/YYYY)"]
    df_resp1a_er_sin_2nd_er = df_resp1a_er_sin_2nd["Etch Rate (A/Min)"]
    df_resp1a_er_sin_2nd_usl = df_resp1a_er_sin_2nd["USL"]
    df_resp1a_er_sin_2nd_lsl = df_resp1a_er_sin_2nd["LSL"]
    df_resp1a_er_sin_2nd_ucl = df_resp1a_er_sin_2nd["UCL"]
    df_resp1a_er_sin_2nd_lcl = df_resp1a_er_sin_2nd["LCL"]
    df_resp1a_er_sin_2nd_remarks = df_resp1a_er_sin_2nd["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-2nd ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1a_er_sin_2nd_plot(
        x = date_formatter(df_resp1a_er_sin_2nd_date), 
        y1 = df_resp1a_er_sin_2nd_er,
        y2 = df_resp1a_er_sin_2nd_usl, 
        y3 = df_resp1a_er_sin_2nd_lsl,
        y4 = df_resp1a_er_sin_2nd_ucl,
        y5 = df_resp1a_er_sin_2nd_lcl,
        remarks = df_resp1a_er_sin_2nd_remarks
    )

def resp1a_sin2nd_unif_chart():
    # Fetch Dataframes for SiN-2nd
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_sin_2nd = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'SiN-2Step']
    df_resp1a_er_sin_2nd = df_resp1a_er_sin_2nd.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_sin_2nd

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-2nd ER & Unif PLot
    df_resp1a_er_sin_2nd_date = df_resp1a_er_sin_2nd["Date (MM/DD/YYYY)"]
    df_resp1a_er_sin_2nd_unif = df_resp1a_er_sin_2nd["% Uni"]
    df_resp1a_er_sin_2nd_unif_usl = df_resp1a_er_sin_2nd["% Uni USL"]
    df_resp1a_er_sin_2nd_unif_ucl = df_resp1a_er_sin_2nd["% Uni UCL"]
    df_resp1a_er_sin_2nd_remarks = df_resp1a_er_sin_2nd["Remarks"]
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-2nd Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1a_unif_sin_2nd_plot(
        x = date_formatter(df_resp1a_er_sin_2nd_date), 
        y1 = df_resp1a_er_sin_2nd_unif, 
        y2 = df_resp1a_er_sin_2nd_unif_usl,
        y3 = df_resp1a_er_sin_2nd_unif_ucl,
        remarks = df_resp1a_er_sin_2nd_remarks
    )

def resp1a_teos1st_er_chart():
    # Fetch Dataframes for TEOS-1st
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_teos_1st = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'TEOS-1Step']
    df_resp1a_er_teos_1st = df_resp1a_er_teos_1st.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_teos_1st

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-1st ER & Unif PLot
    df_resp1a_er_teos_1st_date = df_resp1a_er_teos_1st["Date (MM/DD/YYYY)"]
    df_resp1a_er_teos_1st_er = df_resp1a_er_teos_1st["Etch Rate (A/Min)"]
    df_resp1a_er_teos_1st_usl = df_resp1a_er_teos_1st["USL"]
    df_resp1a_er_teos_1st_lsl = df_resp1a_er_teos_1st["LSL"]
    df_resp1a_er_teos_1st_ucl = df_resp1a_er_teos_1st["UCL"]
    df_resp1a_er_teos_1st_lcl = df_resp1a_er_teos_1st["LCL"]
    df_resp1a_er_teos_1st_remarks = df_resp1a_er_teos_1st["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1a_er_teos_1st_plot(
        x = date_formatter(df_resp1a_er_teos_1st_date), 
        y1 = df_resp1a_er_teos_1st_er,
        y2 = df_resp1a_er_teos_1st_usl, 
        y3 = df_resp1a_er_teos_1st_lsl,
        y4 = df_resp1a_er_teos_1st_ucl,
        y5 = df_resp1a_er_teos_1st_lcl,
        remarks = df_resp1a_er_teos_1st_remarks
    )
   
def resp1a_teos1st_unif_chart():
    # Fetch Dataframes for TEOS-1st
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_teos_1st = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'TEOS-1Step']
    df_resp1a_er_teos_1st = df_resp1a_er_teos_1st.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_teos_1st

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-1st ER & Unif PLot
    df_resp1a_er_teos_1st_date = df_resp1a_er_teos_1st["Date (MM/DD/YYYY)"]
    df_resp1a_er_teos_1st_unif = df_resp1a_er_teos_1st["% Uni"]
    df_resp1a_er_teos_1st_unif_usl = df_resp1a_er_teos_1st["% Uni USL"]
    df_resp1a_er_teos_1st_unif_ucl = df_resp1a_er_teos_1st["% Uni UCL"]
    df_resp1a_er_teos_1st_remarks = df_resp1a_er_teos_1st["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1a_unif_teos_1st_plot(
        x = date_formatter(df_resp1a_er_teos_1st_date), 
        y1 = df_resp1a_er_teos_1st_unif, 
        y2 = df_resp1a_er_teos_1st_unif_usl,
        y3 = df_resp1a_er_teos_1st_unif_ucl,
        remarks = df_resp1a_er_teos_1st_remarks
    )

def resp1a_teos2nd_er_chart():
    # Fetch Dataframes for TEOS-2nd
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_teos_2nd = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'TEOS-2Step']
    df_resp1a_er_teos_2nd = df_resp1a_er_teos_2nd.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_teos_2nd

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-2nd ER & Unif PLot
    df_resp1a_er_teos_2nd_date = df_resp1a_er_teos_2nd["Date (MM/DD/YYYY)"]
    df_resp1a_er_teos_2nd_er = df_resp1a_er_teos_2nd["Etch Rate (A/Min)"]
    df_resp1a_er_teos_2nd_usl = df_resp1a_er_teos_2nd["USL"]
    df_resp1a_er_teos_2nd_lsl = df_resp1a_er_teos_2nd["LSL"]
    df_resp1a_er_teos_2nd_ucl = df_resp1a_er_teos_2nd["UCL"]
    df_resp1a_er_teos_2nd_lcl = df_resp1a_er_teos_2nd["LCL"]
    df_resp1a_er_teos_2nd_remarks = df_resp1a_er_teos_2nd["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1a_er_teos_2nd_plot(
        x = date_formatter(df_resp1a_er_teos_2nd_date), 
        y1 = df_resp1a_er_teos_2nd_er,
        y2 = df_resp1a_er_teos_2nd_usl, 
        y3 = df_resp1a_er_teos_2nd_lsl,
        y4 = df_resp1a_er_teos_2nd_ucl,
        y5 = df_resp1a_er_teos_2nd_lcl,
        remarks = df_resp1a_er_teos_2nd_remarks
    )

def resp1a_teos2nd_unif_chart():
    # Fetch Dataframes for TEOS-2nd
    df_resp1a_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_er_teos_2nd = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'TEOS-2Step']
    df_resp1a_er_teos_2nd = df_resp1a_er_teos_2nd.dropna()
    # Display the dataframes in respective sheets
    # sht_run.range('A25').options(index=False).value = df_resp1a_er_teos_2nd

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-2nd ER & Unif PLot
    df_resp1a_er_teos_2nd_date = df_resp1a_er_teos_2nd["Date (MM/DD/YYYY)"]
    df_resp1a_er_teos_2nd_unif = df_resp1a_er_teos_2nd["% Uni"]
    df_resp1a_er_teos_2nd_unif_usl = df_resp1a_er_teos_2nd["% Uni USL"]
    df_resp1a_er_teos_2nd_unif_ucl = df_resp1a_er_teos_2nd["% Uni UCL"]
    df_resp1a_er_teos_2nd_remarks = df_resp1a_er_teos_2nd["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1a_unif_teos_2nd_plot(
        x = date_formatter(df_resp1a_er_teos_2nd_date), 
        y1 = df_resp1a_er_teos_2nd_unif, 
        y2 = df_resp1a_er_teos_2nd_unif_usl,
        y3 = df_resp1a_er_teos_2nd_unif_ucl,
        remarks = df_resp1a_er_teos_2nd_remarks
    )



#--------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------
# @xw.func
# def hello(name):
#     return "hello {0}".format(name)

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# MAIN Function call
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# if __name__ == "__main__":
#     button_run()

