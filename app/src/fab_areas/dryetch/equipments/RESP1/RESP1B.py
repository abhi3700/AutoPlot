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
cp_plot_title = 'CP Plot for RESP1B'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'RESP1B_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot

# BARC ER Plot
er_barc_plot_title = 'BARC ER Plot for RESP1B'  # title for ER plot
er_barc_plot_xlabel = 'Date'        # xaxis name for ER plot
er_barc_plot_ylabel = 'BARC ER (A/min)'   # yaxis name for ER plot
er_barc_plot_html_file = 'RESP1B_BARC_ER-Plot.html'   # HTML filename for ER plot
er_barc_plot_trace_count = 5    # no. of traces in ER plot

# BARC Unif Plot
unif_barc_plot_title = 'BARC Uniformity Plot for RESP1B'  # title for Unif plot
unif_barc_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_barc_plot_ylabel = 'BARC Unif (%)'    # yaxis name for Unif plot
unif_barc_plot_html_file = 'RESP1B_BARC_Unif-Plot.html'   # HTML filename for Unif plot
unif_barc_plot_trace_count = 2    # no. of traces in Unif plot

# PR ER Plot
er_pr_plot_title = 'PR ER Plot for RESP1B'  # title for ER plot
er_pr_plot_xlabel = 'Date'        # xaxis name for ER plot
er_pr_plot_ylabel = 'PR ER (A/min)'   # yaxis name for ER plot
er_pr_plot_html_file = 'RESP1B_PR_ER-Plot.html'   # HTML filename for ER plot
er_pr_plot_trace_count = 5    # no. of traces in ER plot

# PR Unif Plot
unif_pr_plot_title = 'PR Uniformity Plot for RESP1B'  # title for Unif plot
unif_pr_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_pr_plot_ylabel = 'PR Unif (%)'    # yaxis name for Unif plot
unif_pr_plot_html_file = 'RESP1B_PR_Unif-Plot.html'   # HTML filename for Unif plot
unif_pr_plot_trace_count = 2    # no. of traces in Unif plot

# TEOS ER Plot
er_teos_plot_title = 'TEOS ER Plot for RESP1B'  # title for ER plot
er_teos_plot_xlabel = 'Date'        # xaxis name for ER plot
er_teos_plot_ylabel = 'TEOS ER (A/min)'   # yaxis name for ER plot
er_teos_plot_html_file = 'RESP1B_TEOS_ER-Plot.html'   # HTML filename for ER plot
er_teos_plot_trace_count = 5    # no. of traces in ER plot

# TEOS Unif Plot
unif_teos_plot_title = 'TEOS Uniformity Plot for RESP1B'  # title for Unif plot
unif_teos_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_teos_plot_ylabel = 'TEOS Unif (%)'    # yaxis name for Unif plot
unif_teos_plot_html_file = 'RESP1B_TEOS_Unif-Plot.html'   # HTML filename for Unif plot
unif_teos_plot_trace_count = 3    # no. of traces in Unif plot

# SiN ER Plot
er_sin_plot_title = 'SiN ER Plot for RESP1B'  # title for ER plot
er_sin_plot_xlabel = 'Date'        # xaxis name for ER plot
er_sin_plot_ylabel = 'SiN ER (A/min)'   # yaxis name for ER plot
er_sin_plot_html_file = 'RESP1B_SiN_ER-Plot.html'   # HTML filename for ER plot
er_sin_plot_trace_count = 5    # no. of traces in ER plot

# SiN Unif Plot
unif_sin_plot_title = 'SiN Uniformity Plot for RESP1B'  # title for Unif plot
unif_sin_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_sin_plot_ylabel = 'SiN Unif (%)'    # yaxis name for Unif plot
unif_sin_plot_html_file = 'RESP1B_SiN_Unif-Plot.html'   # HTML filename for Unif plot
unif_sin_plot_trace_count = 3    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'RESP1B-CP'
sht_name_er_barc_pr_teos = 'ER-BARC,PR & TEOS'
sht_name_er_sin = 'SIN ER'

sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP > 0.16u", "Delta CP > 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_barc_pr_teos_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]
sht_er_sin_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]


# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"
date_format_contour = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_barc_range = 'F7:N7'
y_coord_barc_range = 'F8:N8'
x_coord_pr_range = 'F9:N9'
y_coord_pr_range = 'F10:N10'
x_coord_teos_range = 'F11:N11'
y_coord_teos_range = 'F12:N12'
x_coord_sin_range = 'E7:Q7'
y_coord_sin_range = 'E8:Q8'

# Skiprows
skiprows_cp = 9
skiprows_barc_pr_teos = 12
skiprows_nit = 8

# ====================================================================================================
excel_file_directory = "F:\\Coding\\github_repos\\AutoPlot\\examples\\dry_etch\\UNT02_Ch_B_QC_LOG_BOOK\\UNT02_Ch_B_QC_LOG_BOOK.xlsm"

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
def draw_plotly_resp1b_cp_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'delta-CP > 0.16u',
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
            name = 'delta-CP > 0.5u',
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
            name = 'delta-CP AC',
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
def draw_plotly_resp1b_er_barc_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_barc_plot_title,
            xaxis = dict(title= er_barc_plot_xlabel),
            yaxis = dict(title= er_barc_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_barc_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1b_unif_barc_plot(x, y1, y2, y3, remarks):
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
            title = unif_barc_plot_title,
            xaxis = dict(title= unif_barc_plot_xlabel),
            yaxis = dict(title= unif_barc_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_barc_plot_html_file)
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
def draw_plotly_resp1b_er_pr_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_pr_plot_title,
            xaxis = dict(title= er_pr_plot_xlabel),
            yaxis = dict(title= er_pr_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_pr_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1b_unif_pr_plot(x, y1, y2, y3, remarks):
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
            title = unif_pr_plot_title,
            xaxis = dict(title= unif_pr_plot_xlabel),
            yaxis = dict(title= unif_pr_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_pr_plot_html_file)
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
def draw_plotly_resp1b_er_teos_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_teos_plot_title,
            xaxis = dict(title= er_teos_plot_xlabel),
            yaxis = dict(title= er_teos_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_teos_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
# "y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1b_unif_teos_plot(x, y1, y2, y3, remarks):
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
            title = unif_teos_plot_title,
            xaxis = dict(title= unif_teos_plot_xlabel),
            yaxis = dict(title= unif_teos_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_teos_plot_html_file)
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
def draw_plotly_resp1b_er_sin_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_sin_plot_title,
            xaxis = dict(title= er_sin_plot_xlabel),
            yaxis = dict(title= er_sin_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_sin_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1b_unif_sin_plot(x, y1, y2, y3, remarks):
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
            title = unif_sin_plot_title,
            xaxis = dict(title= unif_sin_plot_xlabel),
            yaxis = dict(title= unif_sin_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_sin_plot_html_file)
    return fig

#====================================================================================================================================================================
#####################################################################################################################################################################
excel_file = pd.ExcelFile(excel_file_directory)

def resp1b_cp_chart():
    # Fetch Dataframe for CP
    df_resp1b_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_cp = df_resp1b_cp[sht_cp_columns]        # The final dataframe with required columns
    df_resp1b_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_resp1b_cp['Delta CP > 0.5u'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_resp1b_cp['Delta CP AC'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_resp1b_cp = df_resp1b_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_run.range('A25').options(index=False).value = df_resp1b_cp         # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_resp1b_cp_date = df_resp1b_cp["Date (MM/DD/YYYY)"]
    df_resp1b_cp_delta_cp_1 = df_resp1b_cp["Delta CP > 0.16u"]
    df_resp1b_cp_delta_cp_2 = df_resp1b_cp["Delta CP > 0.5u"]
    df_resp1b_cp_delta_cp_3 = df_resp1b_cp["Delta CP AC"]
    df_resp1b_cp_usl = df_resp1b_cp["USL"]
    df_resp1b_cp_ucl = df_resp1b_cp["UCL"]
    df_resp1b_cp_remarks = df_resp1b_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    return draw_plotly_resp1b_cp_plot(
        x = date_formatter(df_resp1b_cp_date), 
        y1 = df_resp1b_cp_delta_cp_1, 
        y2 = df_resp1b_cp_delta_cp_2, 
        y3 = df_resp1b_cp_delta_cp_3, 
        y4 = df_resp1b_cp_usl, 
        y5 = df_resp1b_cp_ucl,
        remarks = df_resp1b_cp_remarks
    )

def resp1b_barc_er_chart():
    # Fetch Dataframe for BARC
    df_resp1b_er = excel_file.parse(sht_name_er_barc_pr_teos, skiprows=skiprows_barc_pr_teos)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er = df_resp1b_er[sht_er_barc_pr_teos_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_barc = df_resp1b_er[df_resp1b_er["Layer"] == 'BARC']
    df_resp1b_er_barc = df_resp1b_er_barc.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_resp1b_plot_barc.range('A25').options(index=False).value = df_resp1b_barc

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-1st ER & Unif PLot
    df_resp1b_er_barc_date = df_resp1b_er_barc["Date (MM/DD/YYYY)"]
    df_resp1b_er_barc_er = df_resp1b_er_barc["Etch Rate (A/Min)"]
    df_resp1b_er_barc_usl = df_resp1b_er_barc["USL"]
    df_resp1b_er_barc_lsl = df_resp1b_er_barc["LSL"]
    df_resp1b_er_barc_ucl = df_resp1b_er_barc["UCL"]
    df_resp1b_er_barc_lcl = df_resp1b_er_barc["LCL"]
    df_resp1b_er_barc_remarks = df_resp1b_er_barc["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw BARC ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1b_er_barc_plot(
        x = date_formatter(df_resp1b_er_barc_date), 
        y1 = df_resp1b_er_barc_er,
        y2 = df_resp1b_er_barc_usl, 
        y3 = df_resp1b_er_barc_lsl,
        y4 = df_resp1b_er_barc_ucl,
        y5 = df_resp1b_er_barc_lcl,
        remarks = df_resp1b_er_barc_remarks
    )

def resp1b_barc_unif_chart():
    # Fetch Dataframe for BARC
    df_resp1b_er = excel_file.parse(sht_name_er_barc_pr_teos, skiprows=skiprows_barc_pr_teos)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er = df_resp1b_er[sht_er_barc_pr_teos_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_barc = df_resp1b_er[df_resp1b_er["Layer"] == 'BARC']
    df_resp1b_er_barc = df_resp1b_er_barc.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_resp1b_plot_barc.range('A25').options(index=False).value = df_resp1b_barc

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-1st ER & Unif PLot
    df_resp1b_er_barc_date = df_resp1b_er_barc["Date (MM/DD/YYYY)"]
    df_resp1b_er_barc_unif = df_resp1b_er_barc["% Uni"]
    df_resp1b_er_barc_unif_usl = df_resp1b_er_barc["% Uni USL"]
    df_resp1b_er_barc_unif_ucl = df_resp1b_er_barc["% Uni UCL"]
    df_resp1b_er_barc_remarks = df_resp1b_er_barc["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw BARC Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1b_unif_barc_plot(
        x = date_formatter(df_resp1b_er_barc_date), 
        y1 = df_resp1b_er_barc_unif, 
        y2 = df_resp1b_er_barc_unif_usl,
        y3 = df_resp1b_er_barc_unif_ucl,
        remarks = df_resp1b_er_barc_remarks
    )

def resp1b_pr_er_chart():
    # Fetch Dataframe for PR
    df_resp1b_er = excel_file.parse(sht_name_er_barc_pr_teos, skiprows=skiprows_barc_pr_teos)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er = df_resp1b_er[sht_er_barc_pr_teos_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_pr = df_resp1b_er[df_resp1b_er["Layer"] == 'PR']
    df_resp1b_er_pr = df_resp1b_er_pr.dropna()          # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_resp1b_plot_pr.range('A25').options(index=False).value = df_resp1b_pr

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-2nd ER & Unif PLot
    df_resp1b_er_pr_date = df_resp1b_er_pr["Date (MM/DD/YYYY)"]
    df_resp1b_er_pr_er = df_resp1b_er_pr["Etch Rate (A/Min)"]
    df_resp1b_er_pr_usl = df_resp1b_er_pr["USL"]
    df_resp1b_er_pr_lsl = df_resp1b_er_pr["LSL"]
    df_resp1b_er_pr_ucl = df_resp1b_er_pr["UCL"]
    df_resp1b_er_pr_lcl = df_resp1b_er_pr["LCL"]
    df_resp1b_er_pr_remarks = df_resp1b_er_pr["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-2nd ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1b_er_pr_plot(
        x = date_formatter(df_resp1b_er_pr_date), 
        y1 = df_resp1b_er_pr_er,
        y2 = df_resp1b_er_pr_usl, 
        y3 = df_resp1b_er_pr_lsl,
        y4 = df_resp1b_er_pr_ucl,
        y5 = df_resp1b_er_pr_lcl,
        remarks = df_resp1b_er_pr_remarks
    )
   
def resp1b_pr_unif_chart():
    # Fetch Dataframe for PR
    df_resp1b_er = excel_file.parse(sht_name_er_barc_pr_teos, skiprows=skiprows_barc_pr_teos)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er = df_resp1b_er[sht_er_barc_pr_teos_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_pr = df_resp1b_er[df_resp1b_er["Layer"] == 'PR']
    df_resp1b_er_pr = df_resp1b_er_pr.dropna()          # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_resp1b_plot_pr.range('A25').options(index=False).value = df_resp1b_pr

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-2nd ER & Unif PLot
    df_resp1b_er_pr_date = df_resp1b_er_pr["Date (MM/DD/YYYY)"]
    df_resp1b_er_pr_unif = df_resp1b_er_pr["% Uni"]
    df_resp1b_er_pr_unif_usl = df_resp1b_er_pr["% Uni USL"]
    df_resp1b_er_pr_unif_ucl = df_resp1b_er_pr["% Uni UCL"]
    df_resp1b_er_pr_remarks = df_resp1b_er_pr["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-2nd Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1b_unif_pr_plot(
        x = date_formatter(df_resp1b_er_pr_date), 
        y1 = df_resp1b_er_pr_unif, 
        y2 = df_resp1b_er_pr_unif_usl,
        y3 = df_resp1b_er_pr_unif_ucl,
        remarks = df_resp1b_er_pr_remarks
    )

def resp1b_teos_er_chart():
    # Fetch Dataframe for TEOS
    df_resp1b_er = excel_file.parse(sht_name_er_barc_pr_teos, skiprows=skiprows_barc_pr_teos)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er = df_resp1b_er[sht_er_barc_pr_teos_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_teos = df_resp1b_er[df_resp1b_er["Layer"] == 'TEOS']
    df_resp1b_er_teos = df_resp1b_er_teos.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_resp1b_plot_teos.range('A25').options(index=False).value = df_resp1b_teos

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-1st ER & Unif PLot
    df_resp1b_er_teos_date = df_resp1b_er_teos["Date (MM/DD/YYYY)"]
    df_resp1b_er_teos_er = df_resp1b_er_teos["Etch Rate (A/Min)"]
    df_resp1b_er_teos_usl = df_resp1b_er_teos["USL"]
    df_resp1b_er_teos_lsl = df_resp1b_er_teos["LSL"]
    df_resp1b_er_teos_ucl = df_resp1b_er_teos["UCL"]
    df_resp1b_er_teos_lcl = df_resp1b_er_teos["LCL"]
    df_resp1b_er_teos_remarks = df_resp1b_er_teos["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1b_er_teos_plot(
        x = date_formatter(df_resp1b_er_teos_date), 
        y1 = df_resp1b_er_teos_er,
        y2 = df_resp1b_er_teos_usl, 
        y3 = df_resp1b_er_teos_lsl,
        y4 = df_resp1b_er_teos_ucl,
        y5 = df_resp1b_er_teos_lcl,
        remarks = df_resp1b_er_teos_remarks
    )

def resp1b_teos_unif_chart():
    # Fetch Dataframe for TEOS
    df_resp1b_er = excel_file.parse(sht_name_er_barc_pr_teos, skiprows=skiprows_barc_pr_teos)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er = df_resp1b_er[sht_er_barc_pr_teos_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_teos = df_resp1b_er[df_resp1b_er["Layer"] == 'TEOS']
    df_resp1b_er_teos = df_resp1b_er_teos.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_resp1b_plot_teos.range('A25').options(index=False).value = df_resp1b_teos

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-1st ER & Unif PLot
    df_resp1b_er_teos_date = df_resp1b_er_teos["Date (MM/DD/YYYY)"]
    df_resp1b_er_teos_unif = df_resp1b_er_teos["% Uni"]
    df_resp1b_er_teos_unif_usl = df_resp1b_er_teos["% Uni USL"]
    df_resp1b_er_teos_unif_ucl = df_resp1b_er_teos["% Uni UCL"]
    df_resp1b_er_teos_remarks = df_resp1b_er_teos["Remarks"]
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1b_unif_teos_plot(
        x = date_formatter(df_resp1b_er_teos_date), 
        y1 = df_resp1b_er_teos_unif, 
        y2 = df_resp1b_er_teos_unif_usl,
        y3 = df_resp1b_er_teos_unif_ucl,
        remarks = df_resp1b_er_teos_remarks
    )

def resp1b_sin_er_chart():
    # Fetch Dataframe for SiN
    df_resp1b_er_sin = excel_file.parse(sht_name_er_sin, skiprows=skiprows_nit)                            # copy a sheet and paste into another sheet and skiprows 5
    df_resp1b_er_sin['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er_sin = df_resp1b_er_sin[sht_er_sin_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_sin = df_resp1b_er_sin.dropna()            # dropping rows where at least one element is missing
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN ER & Unif PLot
    df_resp1b_er_sin_date = df_resp1b_er_sin["Date (MM/DD/YYYY)"]
    df_resp1b_er_sin_er = df_resp1b_er_sin["Etch Rate (A/Min)"]
    df_resp1b_er_sin_usl = df_resp1b_er_sin["USL"]
    df_resp1b_er_sin_lsl = df_resp1b_er_sin["LSL"]
    df_resp1b_er_sin_ucl = df_resp1b_er_sin["UCL"]
    df_resp1b_er_sin_lcl = df_resp1b_er_sin["LCL"]
    df_resp1b_er_sin_remarks = df_resp1b_er_sin["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN ER Plot (using Plotly) in Browser 
    return draw_plotly_resp1b_er_sin_plot(
        x = date_formatter(df_resp1b_er_sin_date), 
        y1 = df_resp1b_er_sin_er,
        y2 = df_resp1b_er_sin_usl, 
        y3 = df_resp1b_er_sin_lsl,
        y4 = df_resp1b_er_sin_ucl,
        y5 = df_resp1b_er_sin_lcl,
        remarks = df_resp1b_er_sin_remarks
    )

def resp1b_sin_unif_chart():
    # Fetch Dataframe for SiN
    df_resp1b_er_sin = excel_file.parse(sht_name_er_sin, skiprows=skiprows_nit)                            # copy a sheet and paste into another sheet and skiprows 5
    df_resp1b_er_sin['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_resp1b_er_sin = df_resp1b_er_sin[sht_er_sin_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1b_er_sin = df_resp1b_er_sin.dropna()            # dropping rows where at least one element is missing
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN ER & Unif PLot
    df_resp1b_er_sin_date = df_resp1b_er_sin["Date (MM/DD/YYYY)"]
    df_resp1b_er_sin_unif = df_resp1b_er_sin["% Uni"]
    df_resp1b_er_sin_unif_usl = df_resp1b_er_sin["% Uni USL"]
    df_resp1b_er_sin_unif_ucl = df_resp1b_er_sin["% Uni UCL"]
    df_resp1b_er_sin_remarks = df_resp1b_er_sin["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN Unif Plot (using Plotly) in Browser     
    return draw_plotly_resp1b_unif_sin_plot(
        x = date_formatter(df_resp1b_er_sin_date), 
        y1 = df_resp1b_er_sin_unif, 
        y2 = df_resp1b_er_sin_unif_usl,
        y3 = df_resp1b_er_sin_unif_ucl,
        remarks = df_resp1b_er_sin_remarks
    )


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# @xw.func
# def hello(name):
#     return "hello {0}".format(name)


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# MAIN Function call
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# if __name__ == "__main__":
#     button_run()