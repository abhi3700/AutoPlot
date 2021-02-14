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
cp_plot_title = 'CP Plot for REOX1C'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'Delta CP (no.s)'     # yaxis name for CP plot
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
sht_cp_columns = ["Date (MM/DD/YYYY)", "Delta CP 0.16u", "Delta CP 0.5u", "Delta CP AC", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]   # deleted "% Uni UCL" column
 
# Date formatter
date_format = "%d-%m-%Y %H:%M:%S"
date_format_contour = "%d-%m-%Y %H:%M:%S"

# Metrology tool measurement coordinates
x_coord_arc_range = 'D9:L9'
y_coord_arc_range = 'D10:L10'
x_coord_teos_range = 'D11:L11'
y_coord_teos_range = 'D12:L12'

# Skiprows
skiprows_cp = 9
skiprows_er = 12

# =================================================================================================================================================================
excel_file_directory = "F:\\Coding\\github_repos\\AutoPlot\\examples\\dry_etch\\CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK\\CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK.xlsm"
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
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
"y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_reox1c_cp_plot(x, y1, y2, y3, y4, y5, remarks):
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
def draw_plotly_reox1c_er_arc_plot(x, y1, y2, y3, y4, y5, remarks):
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
def draw_plotly_reox1c_unif_arc_plot(x, y1, y2, y3, remarks):
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
def draw_plotly_reox1c_er_teos_plot(x, y1, y2, y3, y4, y5, remarks):
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
def draw_plotly_reox1c_unif_teos_plot(x, y1, y2, y3, remarks):
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


#====================================================================================================================================================================
#####################################################################################################################################################################
excel_file = pd.ExcelFile(excel_file_directory)

def reox1c_cp_chart():
    df_reox1c_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reox1c_cp = df_reox1c_cp[sht_cp_columns]        # The final dataframe with required columns
    df_reox1c_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_reox1c_cp['Delta CP 0.5u'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_reox1c_cp['Delta CP AC'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_reox1c_cp = df_reox1c_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_run.range('A25').options(index=False).value = df_reox1c_cp         # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reox1c_cp_date = df_reox1c_cp["Date (MM/DD/YYYY)"]
    df_reox1c_cp_delta_cp_1 = df_reox1c_cp["Delta CP 0.16u"]
    df_reox1c_cp_delta_cp_2 = df_reox1c_cp["Delta CP 0.5u"]
    df_reox1c_cp_delta_cp_3 = df_reox1c_cp["Delta CP AC"]
    df_reox1c_cp_usl = df_reox1c_cp["USL"]
    df_reox1c_cp_ucl = df_reox1c_cp["UCL"]
    df_reox1c_cp_remarks = df_reox1c_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    return draw_plotly_reox1c_cp_plot(
        x = date_formatter(df_reox1c_cp_date), 
        y1 = df_reox1c_cp_delta_cp_1, 
        y2 = df_reox1c_cp_delta_cp_2, 
        y3 = df_reox1c_cp_delta_cp_3, 
        y4 = df_reox1c_cp_usl, 
        y5 = df_reox1c_cp_ucl,
        remarks = df_reox1c_cp_remarks
    )

def reox1c_arc_er_chart():
    # Fetch Dataframes for ARC
    df_reox1c_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reox1c_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1c_er = df_reox1c_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1c_er_arc = df_reox1c_er[df_reox1c_er["Layer"] == 'ARC']
    df_reox1c_er_arc = df_reox1c_er_arc.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_reox1c_plot_barc.range('A25').options(index=False).value = df_reox1c_barc

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC ER & Unif PLot
    df_reox1c_er_arc_date = df_reox1c_er_arc["Date (MM/DD/YYYY)"]
    df_reox1c_er_arc_er = df_reox1c_er_arc["Etch Rate (A/Min)"]
    df_reox1c_er_arc_usl = df_reox1c_er_arc["USL"]
    df_reox1c_er_arc_lsl = df_reox1c_er_arc["LSL"]
    df_reox1c_er_arc_ucl = df_reox1c_er_arc["UCL"]
    df_reox1c_er_arc_lcl = df_reox1c_er_arc["LCL"]
    df_reox1c_er_arc_remarks = df_reox1c_er_arc["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ARC ER Plot (using Plotly) in Browser 
    return draw_plotly_reox1c_er_arc_plot(
        x = date_formatter(df_reox1c_er_arc_date), 
        y1 = df_reox1c_er_arc_er,
        y2 = df_reox1c_er_arc_usl, 
        y3 = df_reox1c_er_arc_lsl,
        y4 = df_reox1c_er_arc_ucl,
        y5 = df_reox1c_er_arc_lcl,
        remarks = df_reox1c_er_arc_remarks
    )
   
def reox1c_arc_unif_chart():
    # Fetch Dataframes for ARC
    df_reox1c_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reox1c_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1c_er = df_reox1c_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1c_er_arc = df_reox1c_er[df_reox1c_er["Layer"] == 'ARC']
    df_reox1c_er_arc = df_reox1c_er_arc.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_reox1c_plot_barc.range('A25').options(index=False).value = df_reox1c_barc

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for ARC ER & Unif PLot
    df_reox1c_er_arc_date = df_reox1c_er_arc["Date (MM/DD/YYYY)"]
    df_reox1c_er_arc_unif = df_reox1c_er_arc["% Uni"]
    df_reox1c_er_arc_unif_usl = df_reox1c_er_arc["% Uni USL"]
    df_reox1c_er_arc_unif_ucl = df_reox1c_er_arc["% Uni UCL"]
    df_reox1c_er_arc_remarks = df_reox1c_er_arc["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ARC Unif Plot (using Plotly) in Browser     
    return draw_plotly_reox1c_unif_arc_plot(
        x = date_formatter(df_reox1c_er_arc_date), 
        y1 = df_reox1c_er_arc_unif, 
        y2 = df_reox1c_er_arc_unif_usl,
        y3 = df_reox1c_er_arc_unif_ucl,
        remarks = df_reox1c_er_arc_remarks
    )

def reox1c_teos_er_chart():
    # Fetch Dataframes for TEOS
    df_reox1c_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reox1c_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1c_er = df_reox1c_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1c_er_arc = df_reox1c_er[df_reox1c_er["Layer"] == 'ARC']
    df_reox1c_er_arc = df_reox1c_er_arc.dropna()      # dropping rows where at least one element is missing
    df_reox1c_er_teos = df_reox1c_er[df_reox1c_er["Layer"] == 'TEOS']
    df_reox1c_er_teos = df_reox1c_er_teos.dropna()      # dropping rows where at least one element is missing
    # Display the dataframes in respective sheets
    # sht_reox1c_plot_teos.range('A25').options(index=False).value = df_reox1c_teos

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS ER & Unif PLot
    df_reox1c_er_teos_date = df_reox1c_er_teos["Date (MM/DD/YYYY)"]
    df_reox1c_er_teos_er = df_reox1c_er_teos["Etch Rate (A/Min)"]
    df_reox1c_er_teos_usl = df_reox1c_er_teos["USL"]
    df_reox1c_er_teos_lsl = df_reox1c_er_teos["LSL"]
    df_reox1c_er_teos_ucl = df_reox1c_er_teos["UCL"]
    df_reox1c_er_teos_lcl = df_reox1c_er_teos["LCL"]
    df_reox1c_er_teos_remarks = df_reox1c_er_teos["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS ER Plot (using Plotly) in Browser 
    return draw_plotly_reox1c_er_teos_plot(
        x = date_formatter(df_reox1c_er_teos_date), 
        y1 = df_reox1c_er_teos_er,
        y2 = df_reox1c_er_teos_usl, 
        y3 = df_reox1c_er_teos_lsl,
        y4 = df_reox1c_er_teos_ucl,
        y5 = df_reox1c_er_teos_lcl,
        remarks = df_reox1c_er_teos_remarks
    )

def reox1c_teos_unif_chart():
    # Fetch Dataframes for TEOS
    df_reox1c_er = excel_file.parse(sht_name_er, skiprows=skiprows_er)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reox1c_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.' in "Remarks" column 
    df_reox1c_er = df_reox1c_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1c_er_teos = df_reox1c_er[df_reox1c_er["Layer"] == 'TEOS']
    df_reox1c_er_teos = df_reox1c_er_teos.dropna()      # dropping rows where at least one element is missing

    # Display the dataframes in respective sheets
    # sht_reox1c_plot_teos.range('A25').options(index=False).value = df_reox1c_teos
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS ER & Unif PLot
    df_reox1c_er_teos_date = df_reox1c_er_teos["Date (MM/DD/YYYY)"]
    df_reox1c_er_teos_unif = df_reox1c_er_teos["% Uni"]
    df_reox1c_er_teos_unif_usl = df_reox1c_er_teos["% Uni USL"]
    df_reox1c_er_teos_unif_ucl = df_reox1c_er_teos["% Uni UCL"]
    df_reox1c_er_teos_remarks = df_reox1c_er_teos["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS Unif Plot (using Plotly) in Browser     
    return draw_plotly_reox1c_unif_teos_plot(
        x = date_formatter(df_reox1c_er_teos_date), 
        y1 = df_reox1c_er_teos_unif, 
        y2 = df_reox1c_er_teos_unif_usl,
        y3 = df_reox1c_er_teos_unif_ucl,
        remarks = df_reox1c_er_teos_remarks
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