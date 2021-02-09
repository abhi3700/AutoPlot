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
cp_plot_title = 'CP Plot for REOX1B'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REOX1B_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 3    # no. of traces in CP plot

# BPSG ER Plot
er_bpsgcs_plot_title = 'BPSG_CS ER Plot for REOX1B'  # title for ER plot
er_bpsgcs_plot_xlabel = 'Date'        # xaxis name for ER plot
er_bpsgcs_plot_ylabel = 'BPSG_CS ER (A/min)'   # yaxis name for ER plot
er_bpsgcs_plot_html_file = 'REOX1B_BPSG_CS_ER-Plot.html'   # HTML filename for ER plot
er_bpsgcs_plot_trace_count = 5    # no. of traces in ER plot

# BPSG Unif Plot
unif_bpsgcs_plot_title = 'BPSG_CS Uniformity Plot for REOX1B'  # title for Unif plot
unif_bpsgcs_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_bpsgcs_plot_ylabel = 'BPSG_CS Unif (%)'    # yaxis name for Unif plot
unif_bpsgcs_plot_html_file = 'REOX1B_BPSG_CS_Unif-Plot.html'   # HTML filename for Unif plot
unif_bpsgcs_plot_trace_count = 2    # no. of traces in Unif plot

# SiNCS ER Plot
er_sincs_plot_title = 'SiN_CS ER Plot for REOX1B'  # title for ER plot
er_sincs_plot_xlabel = 'Date'        # xaxis name for ER plot
er_sincs_plot_ylabel = 'SiN_CS ER (A/min)'   # yaxis name for ER plot
er_sincs_plot_html_file = 'REOX1B_SiN_CS_ER-Plot.html'   # HTML filename for ER plot
er_sincs_plot_trace_count = 5    # no. of traces in ER plot

# SiNCS Unif Plot
unif_sincs_plot_title = 'SiN_CS Uniformity Plot for REOX1B'  # title for Unif plot
unif_sincs_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_sincs_plot_ylabel = 'SiN_CS Unif (%)'    # yaxis name for Unif plot
unif_sincs_plot_html_file = 'REOX1B_SiN_CS_Unif-Plot.html'   # HTML filename for Unif plot
unif_sincs_plot_trace_count = 2    # no. of traces in Unif plot

# TEOSVIA ER Plot
er_teosvia_plot_title = 'TEOS_VIA ER Plot for REOX1B'  # title for ER plot
er_teosvia_plot_xlabel = 'Date'        # xaxis name for ER plot
er_teosvia_plot_ylabel = 'TEOS_VIA ER (A/min)'   # yaxis name for ER plot
er_teosvia_plot_html_file = 'REOX1B_TEOS_VIA_ER-Plot.html'   # HTML filename for ER plot
er_teosvia_plot_trace_count = 5    # no. of traces in ER plot

# TEOSVIA Unif Plot
unif_teosvia_plot_title = 'TEOS_VIA Uniformity Plot for REOX1B'  # title for Unif plot
unif_teosvia_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_teosvia_plot_ylabel = 'TEOS_VIA Unif (%)'    # yaxis name for Unif plot
unif_teosvia_plot_html_file = 'REOX1B_TEOS_VIA_Unif-Plot.html'   # HTML filename for Unif plot
unif_teosvia_plot_trace_count = 3    # no. of traces in Unif plot

# ARC ER Plot
er_arc_plot_title = 'ARC ER Plot for REOX1B'  # title for ER plot
er_arc_plot_xlabel = 'Date'        # xaxis name for ER plot
er_arc_plot_ylabel = 'ARC ER (A/min)'   # yaxis name for ER plot
er_arc_plot_html_file = 'REOX1B_ARC_ER-Plot.html'   # HTML filename for ER plot
er_arc_plot_trace_count = 5    # no. of traces in ER plot

# ARC Unif Plot
unif_arc_plot_title = 'ARC Uniformity Plot for REOX1B'  # title for Unif plot
unif_arc_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_arc_plot_ylabel = 'ARC Unif (%)'    # yaxis name for Unif plot
unif_arc_plot_html_file = 'REOX1B_ARC_Unif-Plot.html'   # HTML filename for Unif plot
unif_arc_plot_trace_count = 3    # no. of traces in Unif plot

# Sheet names
sht_name_cp = 'REOX1B-CP'
sht_name_er = 'REOX1B-ER'

# Columns
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.16u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]

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

# =================================================================================================================================================================
excel_file_directory = "F:\\Coding\\github_repos\\AutoPlot\\examples\\dry_etch\\CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK\\CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK.xlsm"

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": Date formatter to format the excel date (issue: one date less in plotly chart) as "%m-%d-%Y %H:%M:%S"
"x": datetime list
"return": formatted datetime list
"""
def date_formatter(x):
    x_fmt = []
    for a in x:
        a = a.strftime("%m-%d-%Y %H:%M:%S")
        x_fmt.append(a)
    return x_fmt


#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
"y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_reox1b_cp_plot(x, y1, y2, y3, y4, y5, remarks):
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
def draw_plotly_reox1b_er_bpsgcs_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_bpsgcs_plot_title,
            xaxis = dict(title= er_bpsgcs_plot_xlabel),
            yaxis = dict(title= er_bpsgcs_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_bpsgcs_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1b_unif_bpsgcs_plot(x, y1, y2, y3, remarks):
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
            title = unif_bpsgcs_plot_title,
            xaxis = dict(title= unif_bpsgcs_plot_xlabel),
            yaxis = dict(title= unif_bpsgcs_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_bpsgcs_plot_html_file)
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
def draw_plotly_reox1b_er_sincs_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_sincs_plot_title,
            xaxis = dict(title= er_sincs_plot_xlabel),
            yaxis = dict(title= er_sincs_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_sincs_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1b_unif_sincs_plot(x, y1, y2, y3, remarks):
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
            title = unif_sincs_plot_title,
            xaxis = dict(title= unif_sincs_plot_xlabel),
            yaxis = dict(title= unif_sincs_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_sincs_plot_html_file)
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
def draw_plotly_reox1b_er_teosvia_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_teosvia_plot_title,
            xaxis = dict(title= er_teosvia_plot_xlabel),
            yaxis = dict(title= er_teosvia_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_teosvia_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1b_unif_teosvia_plot(x, y1, y2, y3, remarks):
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
            title = unif_teosvia_plot_title,
            xaxis = dict(title= unif_teosvia_plot_xlabel),
            yaxis = dict(title= unif_teosvia_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_teosvia_plot_html_file)
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
def draw_plotly_reox1b_er_arc_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_arc_plot_title,
            xaxis = dict(title= er_arc_plot_xlabel),
            yaxis = dict(title= er_arc_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_arc_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1b_unif_arc_plot(x, y1, y2, y3, remarks):
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
            title = unif_arc_plot_title,
            xaxis = dict(title= unif_arc_plot_xlabel),
            yaxis = dict(title= unif_arc_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_arc_plot_html_file)
    return fig

#====================================================================================================================================================================
#####################################################################################################################################################################
excel_file = pd.ExcelFile(excel_file_directory)

def reox1b_cp_chart():
    # Fetch Dataframe for CP
    df_reox1b_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows
    df_reox1b_cp = df_reox1b_cp[sht_cp_columns]        # The final dataframe with required columns
    df_reox1b_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_reox1b_cp['Delta CP 0.5u'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reox1b_cp['Delta CP AC'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reox1b_cp = df_reox1b_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_run.range('A25').options(index=False).value = df_reox1b_cp         # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reox1b_cp_date = df_reox1b_cp["Date (MM/DD/YYYY)"]
    df_reox1b_cp_delta_cp_1 = df_reox1b_cp["Delta CP 0.16u"]
    df_reox1b_cp_delta_cp_2 = df_reox1b_cp["Delta CP 0.5u"]
    df_reox1b_cp_delta_cp_3 = df_reox1b_cp["Delta CP AC"]
    df_reox1b_cp_usl = df_reox1b_cp["USL"]
    df_reox1b_cp_ucl = df_reox1b_cp["UCL"]
    df_reox1b_cp_remarks = df_reox1b_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    return draw_plotly_reox1b_cp_plot(
        x = date_formatter(df_reox1b_cp_date), 
        y1 = df_reox1b_cp_delta_cp_1, 
        y2 = df_reox1b_cp_delta_cp_2, 
        y3 = df_reox1b_cp_delta_cp_3, 
        y4 = df_reox1b_cp_usl, 
        y5 = df_reox1b_cp_ucl,
        remarks = df_reox1b_cp_remarks
    )

def reox1b_bpsgcs_er_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_bpsgcs = df_reox1b_er[df_reox1b_er["Layer"] == 'BPSG_CS']
    df_reox1b_er_bpsgcs = df_reox1b_er_bpsgcs.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_bpsgcs

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC-1st ER & Unif PLot
    df_reox1b_er_bpsgcs_date = df_reox1b_er_bpsgcs["Date (MM/DD/YYYY)"]
    df_reox1b_er_bpsgcs_er = df_reox1b_er_bpsgcs["Etch Rate (A/Min)"]
    df_reox1b_er_bpsgcs_usl = df_reox1b_er_bpsgcs["USL"]
    df_reox1b_er_bpsgcs_lsl = df_reox1b_er_bpsgcs["LSL"]
    df_reox1b_er_bpsgcs_ucl = df_reox1b_er_bpsgcs["UCL"]
    df_reox1b_er_bpsgcs_lcl = df_reox1b_er_bpsgcs["LCL"]
    df_reox1b_er_bpsgcs_remarks = df_reox1b_er_bpsgcs["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw BPSG_CS ER Plot (using Plotly) in Browser 
    return draw_plotly_reox1b_er_bpsgcs_plot(
        x = date_formatter(df_reox1b_er_bpsgcs_date), 
        y1 = df_reox1b_er_bpsgcs_er,
        y2 = df_reox1b_er_bpsgcs_usl, 
        y3 = df_reox1b_er_bpsgcs_lsl,
        y4 = df_reox1b_er_bpsgcs_ucl,
        y5 = df_reox1b_er_bpsgcs_lcl,
        remarks = df_reox1b_er_bpsgcs_remarks
    )
   
def reox1b_bpsgcs_unif_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_bpsgcs = df_reox1b_er[df_reox1b_er["Layer"] == 'BPSG_CS']
    df_reox1b_er_bpsgcs = df_reox1b_er_bpsgcs.dropna()      # dropping rows where at least one element is missing
    df_reox1b_er_sincs = df_reox1b_er[df_reox1b_er["Layer"] == 'SiN_CS']
    df_reox1b_er_sincs = df_reox1b_er_sincs.dropna()          # dropping rows where at least one element is missing
    df_reox1b_er_teosvia = df_reox1b_er[df_reox1b_er["Layer"] == 'TEOS_VIA']
    df_reox1b_er_teosvia = df_reox1b_er_teosvia.dropna()      # dropping rows where at least one element is missing
    df_reox1b_er_arc = df_reox1b_er[df_reox1b_er["Layer"] == 'ARC']
    df_reox1b_er_arc = df_reox1b_er_arc.dropna()      # dropping rows where at least one element is missing

    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_bpsgcs
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_sincs
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_teosvia
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_arc
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC-1st ER & Unif PLot
    df_reox1b_er_bpsgcs_date = df_reox1b_er_bpsgcs["Date (MM/DD/YYYY)"]
    df_reox1b_er_bpsgcs_unif = df_reox1b_er_bpsgcs["% Uni"]
    df_reox1b_er_bpsgcs_unif_usl = df_reox1b_er_bpsgcs["% Uni USL"]
    df_reox1b_er_bpsgcs_unif_ucl = df_reox1b_er_bpsgcs["% Uni UCL"]
    df_reox1b_er_bpsgcs_remarks = df_reox1b_er_bpsgcs["Remarks"]
     
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw BPSG_CS Unif Plot (using Plotly) in Browser     
    return draw_plotly_reox1b_unif_bpsgcs_plot(
        x = date_formatter(df_reox1b_er_bpsgcs_date), 
        y1 = df_reox1b_er_bpsgcs_unif, 
        y2 = df_reox1b_er_bpsgcs_unif_usl,
        y3 = df_reox1b_er_bpsgcs_unif_ucl,
        remarks = df_reox1b_er_bpsgcs_remarks
    )

def reox1b_sincs_er_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_sincs = df_reox1b_er[df_reox1b_er["Layer"] == 'SiN_CS']
    df_reox1b_er_sincs = df_reox1b_er_sincs.dropna()          # dropping rows where at least one element is missing

    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_sincs
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC-2nd ER & Unif PLot
    df_reox1b_er_sincs_date = df_reox1b_er_sincs["Date (MM/DD/YYYY)"]
    df_reox1b_er_sincs_er = df_reox1b_er_sincs["Etch Rate (A/Min)"]
    df_reox1b_er_sincs_usl = df_reox1b_er_sincs["USL"]
    df_reox1b_er_sincs_lsl = df_reox1b_er_sincs["LSL"]
    df_reox1b_er_sincs_ucl = df_reox1b_er_sincs["UCL"]
    df_reox1b_er_sincs_lcl = df_reox1b_er_sincs["LCL"]
    df_reox1b_er_sincs_remarks = df_reox1b_er_sincs["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ARC-2nd ER Plot (using Plotly) in Browser 
    return draw_plotly_reox1b_er_sincs_plot(
        x = date_formatter(df_reox1b_er_sincs_date), 
        y1 = df_reox1b_er_sincs_er,
        y2 = df_reox1b_er_sincs_usl, 
        y3 = df_reox1b_er_sincs_lsl,
        y4 = df_reox1b_er_sincs_ucl,
        y5 = df_reox1b_er_sincs_lcl,
        remarks = df_reox1b_er_sincs_remarks
    )
   
def reox1b_sincs_unif_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_sincs = df_reox1b_er[df_reox1b_er["Layer"] == 'SiN_CS']
    df_reox1b_er_sincs = df_reox1b_er_sincs.dropna()          # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_sincs
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC-2nd ER & Unif PLot
    df_reox1b_er_sincs_date = df_reox1b_er_sincs["Date (MM/DD/YYYY)"]
    df_reox1b_er_sincs_unif = df_reox1b_er_sincs["% Uni"]
    df_reox1b_er_sincs_unif_usl = df_reox1b_er_sincs["% Uni USL"]
    df_reox1b_er_sincs_unif_ucl = df_reox1b_er_sincs["% Uni UCL"]
    df_reox1b_er_sincs_remarks = df_reox1b_er_sincs["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ARC-2nd Unif Plot (using Plotly) in Browser     
    return draw_plotly_reox1b_unif_sincs_plot(
        x = date_formatter(df_reox1b_er_sincs_date), 
        y1 = df_reox1b_er_sincs_unif, 
        y2 = df_reox1b_er_sincs_unif_usl,
        y3 = df_reox1b_er_sincs_unif_ucl,
        remarks = df_reox1b_er_sincs_remarks
    )

def reox1b_teosvia_er_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_teosvia = df_reox1b_er[df_reox1b_er["Layer"] == 'TEOS_VIA']
    df_reox1b_er_teosvia = df_reox1b_er_teosvia.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_teosvia

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS_VIA-1st ER & Unif PLot
    df_reox1b_er_teosvia_date = df_reox1b_er_teosvia["Date (MM/DD/YYYY)"]
    df_reox1b_er_teosvia_er = df_reox1b_er_teosvia["Etch Rate (A/Min)"]
    df_reox1b_er_teosvia_usl = df_reox1b_er_teosvia["USL"]
    df_reox1b_er_teosvia_lsl = df_reox1b_er_teosvia["LSL"]
    df_reox1b_er_teosvia_ucl = df_reox1b_er_teosvia["UCL"]
    df_reox1b_er_teosvia_lcl = df_reox1b_er_teosvia["LCL"]
    df_reox1b_er_teosvia_remarks = df_reox1b_er_teosvia["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS_VIA ER Plot (using Plotly) in Browser 
    return draw_plotly_reox1b_er_teosvia_plot(
        x = date_formatter(df_reox1b_er_teosvia_date), 
        y1 = df_reox1b_er_teosvia_er,
        y2 = df_reox1b_er_teosvia_usl, 
        y3 = df_reox1b_er_teosvia_lsl,
        y4 = df_reox1b_er_teosvia_ucl,
        y5 = df_reox1b_er_teosvia_lcl,
        remarks = df_reox1b_er_teosvia_remarks
    )
   
def reox1b_teosvia_unif_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_teosvia = df_reox1b_er[df_reox1b_er["Layer"] == 'TEOS_VIA']
    df_reox1b_er_teosvia = df_reox1b_er_teosvia.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_teosvia

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS_VIA-1st ER & Unif PLot
    df_reox1b_er_teosvia_date = df_reox1b_er_teosvia["Date (MM/DD/YYYY)"]
    df_reox1b_er_teosvia_unif = df_reox1b_er_teosvia["% Uni"]
    df_reox1b_er_teosvia_unif_usl = df_reox1b_er_teosvia["% Uni USL"]
    df_reox1b_er_teosvia_unif_ucl = df_reox1b_er_teosvia["% Uni UCL"]
    df_reox1b_er_teosvia_remarks = df_reox1b_er_teosvia["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS_VIA Unif Plot (using Plotly) in Browser     
    return draw_plotly_reox1b_unif_teosvia_plot(
        x = date_formatter(df_reox1b_er_teosvia_date), 
        y1 = df_reox1b_er_teosvia_unif, 
        y2 = df_reox1b_er_teosvia_unif_usl,
        y3 = df_reox1b_er_teosvia_unif_ucl,
        remarks = df_reox1b_er_teosvia_remarks
    )

def reox1b_arc_er_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_arc = df_reox1b_er[df_reox1b_er["Layer"] == 'ARC']
    df_reox1b_er_arc = df_reox1b_er_arc.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_arc

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC ER & Unif PLot
    df_reox1b_er_arc_date = df_reox1b_er_arc["Date (MM/DD/YYYY)"]
    df_reox1b_er_arc_er = df_reox1b_er_arc["Etch Rate (A/Min)"]
    df_reox1b_er_arc_usl = df_reox1b_er_arc["USL"]
    df_reox1b_er_arc_lsl = df_reox1b_er_arc["LSL"]
    df_reox1b_er_arc_ucl = df_reox1b_er_arc["UCL"]
    df_reox1b_er_arc_lcl = df_reox1b_er_arc["LCL"]
    df_reox1b_er_arc_remarks = df_reox1b_er_arc["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ARC ER Plot (using Plotly) in Browser 
    return draw_plotly_reox1b_er_arc_plot(
        x = date_formatter(df_reox1b_er_arc_date), 
        y1 = df_reox1b_er_arc_er,
        y2 = df_reox1b_er_arc_usl, 
        y3 = df_reox1b_er_arc_lsl,
        y4 = df_reox1b_er_arc_ucl,
        y5 = df_reox1b_er_arc_lcl,
        remarks = df_reox1b_er_arc_remarks
    )

def reox1b_arc_unif_chart():
    # Fetch Dataframes for BPSG_CS
    df_reox1b_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 13
    df_reox1b_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1b_er = df_reox1b_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1b_er_arc = df_reox1b_er[df_reox1b_er["Layer"] == 'ARC']
    df_reox1b_er_arc = df_reox1b_er_arc.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_run.range('A5').options(index=False).value = df_reox1b_er_arc

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC ER & Unif PLot
    df_reox1b_er_arc_date = df_reox1b_er_arc["Date (MM/DD/YYYY)"]
    df_reox1b_er_arc_unif = df_reox1b_er_arc["% Uni"]
    df_reox1b_er_arc_unif_usl = df_reox1b_er_arc["% Uni USL"]
    df_reox1b_er_arc_unif_ucl = df_reox1b_er_arc["% Uni UCL"]
    df_reox1b_er_arc_remarks = df_reox1b_er_arc["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ARC Unif Plot (using Plotly) in Browser     
    return draw_plotly_reox1b_unif_arc_plot(
        x = date_formatter(df_reox1b_er_arc_date), 
        y1 = df_reox1b_er_arc_unif, 
        y2 = df_reox1b_er_arc_unif_usl,
        y3 = df_reox1b_er_arc_unif_ucl,
        remarks = df_reox1b_er_arc_remarks
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