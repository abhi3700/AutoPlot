# Import packages
import pandas as pd
import plotly as py
import plotly.graph_objs as go
# import datetime as dt

#========================================INPUT===========================================================================================================
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
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'ASFE1_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 3    # no. of traces in CP plot

# PR ER Plot
er_plot_title = 'ER Plot for ASFE1'  # title for ER plot
er_plot_xlabel = 'Date'        # xaxis name for ER plot
er_plot_ylabel = 'Etch Rate (A/min)'   # yaxis name for ER plot
er_plot_html_file = 'ASFE1_ER-Plot.html'   # HTML filename for ER plot
er_plot_trace_count = 4    # no. of traces in ER plot

# PR Unif Plot
unif_plot_title = 'Uniformity Plot for ASFE1'  # title for UNIF plot
unif_plot_xlabel = 'Date'      # xaxis name for Unif plot
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

# =========================DIR=====================================================================================================
excel_file_directory = "F:\\Coding\\github_repos\\AutoPlot\\examples\\dry_etch\\ASH09_QC_LOG_BOOK\\ASH09_QC_LOG_BOOK.xlsm"

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
"Description": This function plots CP Chart with `cp_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP 0.16u (y-axis) for CP Chart
"y2": Delta-CP 0.5u (y-axis) for CP Chart
"y3": Delta-CP AC (y-axis) for CP Chart
"y4": USL (y-axis) for CP Chart
"y5": UCL (y-axis) for CP Chart
"""
def draw_plotly_asfe1_cp_plot(x, y1, y2, y3, y4, y5, remarks):
    # x = dt.datetime(x)
    # x = x.strftime("%m-%d-%Y %H:%M:%S")
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
"Description": This function plots ER Chart with `er_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_asfe1_er_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_plot_title,
            xaxis = dict(title= er_plot_xlabel),
            yaxis = dict(title= er_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": UCL (y-axis) for Unif Chart
"y3": USL (y-axis) for Unif Chart
"""
def draw_plotly_asfe1_unif_plot(x, y1, y2, y3, remarks):
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
            title = unif_plot_title,
            xaxis = dict(title= unif_plot_xlabel),
            yaxis = dict(title= unif_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_plot_html_file)
    return fig

#====================================================================================================================================================================
#####################################################################################################################################################################
excel_file = pd.ExcelFile(excel_file_directory)

def asfe1_cp_chart():
    df_asfe1_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows 8
    df_asfe1_cp = df_asfe1_cp[sht_cp_columns]        # The final dataframe with required columns
    df_asfe1_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_cp['Delta CP 0.5u'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_cp['Delta CP AC'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_cp = df_asfe1_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_asfe1_plot_cp.range('A25').options(index=False).value = df_asfe1_cp           # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_asfe1_cp_date = df_asfe1_cp["Date (MM/DD/YYYY)"]
    df_asfe1_cp_delta_cp_1 = df_asfe1_cp["Delta CP 0.16u"]
    df_asfe1_cp_delta_cp_2 = df_asfe1_cp["Delta CP 0.5u"]
    df_asfe1_cp_delta_cp_3 = df_asfe1_cp["Delta CP AC"]
    df_asfe1_cp_usl = df_asfe1_cp["USL"]
    df_asfe1_cp_ucl = df_asfe1_cp["UCL"]
    df_asfe1_cp_remarks = df_asfe1_cp["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    return draw_plotly_asfe1_cp_plot(
        x = date_formatter(df_asfe1_cp_date), 
        y1 = df_asfe1_cp_delta_cp_1, 
        y2 = df_asfe1_cp_delta_cp_2, 
        y3 = df_asfe1_cp_delta_cp_3, 
        y4 = df_asfe1_cp_usl, 
        y5 = df_asfe1_cp_ucl,
        remarks = df_asfe1_cp_remarks
    )

def asfe1_er_chart():
    # Fetch Dataframe for ER & Unif Plot
    df_asfe1_er = excel_file.parse(sht_name_er, skiprows=skiprows_pr)                            # copy a sheet and paste into another sheet and skiprows 8
    df_asfe1_er = df_asfe1_er[sht_er_columns]             # The final Dataframe with 5 columns for plot: x-1, y-4
    df_asfe1_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_er = df_asfe1_er.dropna()                                              # dropping rows where at least one element is missing
    # sht_asfe1_plot_er.range('A28').options(index=False).value = df_asfe1_er        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for ER & Unif PLot
    df_asfe1_er_date = df_asfe1_er["Date (MM/DD/YYYY)"]
    df_asfe1_er_er = df_asfe1_er["Etch Rate (A/Min)"]
    df_asfe1_er_usl = df_asfe1_er["USL"]
    df_asfe1_er_ucl = df_asfe1_er["UCL"]
    df_asfe1_er_lsl = df_asfe1_er["LSL"]
    df_asfe1_er_lcl = df_asfe1_er["LCL"]
    df_asfe1_er_remarks = df_asfe1_er["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ER Plot (using Plotly) in Browser 
    return draw_plotly_asfe1_er_plot(
        x = date_formatter(df_asfe1_er_date), 
        y1 = df_asfe1_er_er,
        y2 = df_asfe1_er_usl, 
        y3 = df_asfe1_er_lsl, 
        y4 = df_asfe1_er_ucl,
        y5 = df_asfe1_er_lcl,
        remarks = df_asfe1_er_remarks
    )

def asfe1_unif_chart():
    # Fetch Dataframe for ER & Unif Plot
    df_asfe1_er = excel_file.parse(sht_name_er, skiprows=skiprows_pr)                            # copy a sheet and paste into another sheet and skiprows 8
    df_asfe1_er = df_asfe1_er[sht_er_columns]             # The final Dataframe with 5 columns for plot: x-1, y-4
    df_asfe1_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_er = df_asfe1_er.dropna()                                              # dropping rows where at least one element is missing
    # sht_asfe1_plot_er.range('A28').options(index=False).value = df_asfe1_er        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for ER & Unif PLot
    df_asfe1_er_date = df_asfe1_er["Date (MM/DD/YYYY)"]
    df_asfe1_er_unif = df_asfe1_er["% Uni"]
    df_asfe1_er_unif_ucl = df_asfe1_er["% Uni UCL"]
    df_asfe1_er_unif_usl = df_asfe1_er["% Uni USL"]
    df_asfe1_er_remarks = df_asfe1_er["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Unif Plot (using Plotly) in Browser     
    return draw_plotly_asfe1_unif_plot(
        x = date_formatter(df_asfe1_er_date), 
        y1 = df_asfe1_er_unif, 
        y2 = df_asfe1_er_unif_usl,
        y3 = df_asfe1_er_unif_ucl,
        remarks = df_asfe1_er_remarks
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
    # button_run()


