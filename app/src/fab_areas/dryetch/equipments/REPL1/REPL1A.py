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
cp_plot_title = 'CP Plot for REPL1A'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REPL1A_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot

# NIT ER Plot
er_nit_plot_title = 'Nit ER Plot for REPL1A'  # title for Nit ER plot
er_nit_plot_xlabel = 'Date'        # xaxis name for Nit ER plot
er_nit_plot_ylabel = 'Nit ER (A/min)'   # yaxis name for Nit ER plot
er_nit_plot_html_file = 'REPL1A_Nit_ER-Plot.html'   # HTML filename for Nit ER plot
er_nit_plot_trace_count = 5    # no. of traces in Nit ER plot
er_nit_contour_fname = 'REPL1A_Nit_ER-Plot_Contour.html'
er_nit_contour_tname = 'NIT ER'

# NIT Unif Plot
unif_nit_plot_title = 'Nit Uniformity Plot for REPL1A'  # title for Nit Unif plot
unif_nit_plot_xlabel = 'Date'      # xaxis name for Nit Unif plot
unif_nit_plot_ylabel = 'Nit Unif (%)'    # yaxis name for Nit Unif plot
unif_nit_plot_html_file = 'REPL1A_Nit_Unif-Plot.html'   # HTML filename for Nit Unif plot
unif_nit_plot_trace_count = 3    # no. of traces in Nit Unif plot

# POLY ER Plot
er_poly_plot_title = 'Poly ER Plot for REPL1A'  # title for Poly ER plot
er_poly_plot_xlabel = 'Date'        # xaxis name for Poly ER plot
er_poly_plot_ylabel = 'Poly ER (A/min)'   # yaxis name for Poly ER plot
er_poly_plot_html_file = 'REPL1A_Poly_ER-Plot.html'   # HTML filename for Poly ER plot
er_poly_plot_trace_count = 5    # no. of traces in Poly ER plot
er_poly_contour_fname = 'REPL1A_Poly_ER-Plot_Contour.html'
er_poly_contour_tname = 'POLY ER'

# POLY Unif Plot
unif_poly_plot_title = 'Poly Uniformity Plot for REPL1A'  # title for Poly Unif plot
unif_poly_plot_xlabel = 'Date'      # xaxis name for Poly Unif plot
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

# ================================================================================================================================================================
excel_file_directory = "F:\\Coding\\github_repos\\AutoPlot\\examples\\dry_etch\\CNT01_Ch_A_QC_LOG_BOOK\\CNT01_Ch_A_QC_LOG_BOOK.xlsm"

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
def draw_plotly_repl1a_cp_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'delta-CP 0.16u',
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
            name = 'delta-CP 0.5u',
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
"Description": This function plots ER Chart with `er_nit_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_repl1a_er_nit_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_nit_plot_title,
            xaxis = dict(title= er_nit_plot_xlabel),
            yaxis = dict(title= er_nit_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_nit_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_repl1a_unif_nit_plot(x, y1, y2, y3, remarks):
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
            title = unif_nit_plot_title,
            xaxis = dict(title= unif_nit_plot_xlabel),
            yaxis = dict(title= unif_nit_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_nit_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with `er_poly_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_repl1a_er_poly_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_poly_plot_title,
            xaxis = dict(title= er_poly_plot_xlabel),
            yaxis = dict(title= er_poly_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= er_poly_plot_html_file)
    return fig

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_poly_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_repl1a_unif_poly_plot(x, y1, y2, y3, remarks):
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
            title = unif_poly_plot_title,
            xaxis = dict(title= unif_poly_plot_xlabel),
            yaxis = dict(title= unif_poly_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    # py.offline.plot(fig, filename= unif_poly_plot_html_file)
    return fig

#====================================================================================================================================================================
#####################################################################################################################################################################
excel_file = pd.ExcelFile(excel_file_directory)

def repl1a_cp_chart():
    df_repl1a_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows 9
    df_repl1a_cp = df_repl1a_cp[sht_cp_columns]        # The final dataframe with required columns
    df_repl1a_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_repl1a_cp['Delta CP 0.5u'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_cp['Delta CP AC'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_cp.dropna(inplace=True)                                              # dropping rows where at least one element is missing
    # sht_run.range('A10').options(index=False).value = df_repl1a_cp         # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_repl1a_cp_date = df_repl1a_cp["Date (MM/DD/YYYY)"]
    df_repl1a_cp_delta_cp_1 = df_repl1a_cp["Delta CP 0.16u"]
    df_repl1a_cp_delta_cp_2 = df_repl1a_cp["Delta CP 0.5u"]
    df_repl1a_cp_delta_cp_3 = df_repl1a_cp["Delta CP AC"]
    df_repl1a_cp_usl = df_repl1a_cp["USL"]
    df_repl1a_cp_ucl = df_repl1a_cp["UCL"]
    df_repl1a_cp_remarks = df_repl1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    return draw_plotly_repl1a_cp_plot(
        x = date_formatter(df_repl1a_cp_date), 
        y1 = df_repl1a_cp_delta_cp_1, 
        y2 = df_repl1a_cp_delta_cp_2, 
        y3 = df_repl1a_cp_delta_cp_3, 
        y4 = df_repl1a_cp_usl, 
        y5 = df_repl1a_cp_ucl,
        remarks = df_repl1a_cp_remarks
    )

def repl1a_nit_er_chart():
    df_repl1a_er_nit = excel_file.parse(sht_name_er_nit, skiprows=skiprows_nit)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_nit = df_repl1a_er_nit[sht_er_nit_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_nit['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_nit.dropna(inplace=True)                                              # dropping rows where at least one element is misnitg
    # sht_run.range('A10').options(index=False).value = df_repl1a_er_nit        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for Nit ER & Unif PLot
    df_repl1a_er_nit_date = df_repl1a_er_nit["Date (MM/DD/YYYY)"]
    df_repl1a_er_nit_er = df_repl1a_er_nit["Etch Rate (A/Min)"]
    df_repl1a_er_nit_usl = df_repl1a_er_nit["USL"]
    df_repl1a_er_nit_lsl = df_repl1a_er_nit["LSL"]
    df_repl1a_er_nit_ucl = df_repl1a_er_nit["UCL"]
    df_repl1a_er_nit_lcl = df_repl1a_er_nit["LCL"]
    df_repl1a_er_nit_remarks = df_repl1a_er_nit["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Nit ER Plot (using Plotly) in Browser 
    return draw_plotly_repl1a_er_nit_plot(
        x = date_formatter(df_repl1a_er_nit_date), 
        y1 = df_repl1a_er_nit_er,
        y2 = df_repl1a_er_nit_usl, 
        y3 = df_repl1a_er_nit_lsl,
        y4 = df_repl1a_er_nit_ucl,
        y5 = df_repl1a_er_nit_lcl,
        remarks = df_repl1a_er_nit_remarks
    )

def repl1a_nit_unif_chart():
    df_repl1a_er_nit = excel_file.parse(sht_name_er_nit, skiprows=skiprows_nit)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_nit = df_repl1a_er_nit[sht_er_nit_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_nit['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_nit.dropna(inplace=True)                                              # dropping rows where at least one element is misnitg
    # sht_run.range('A10').options(index=False).value = df_repl1a_er_nit        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for Nit ER & Unif PLot
    df_repl1a_er_nit_date = df_repl1a_er_nit["Date (MM/DD/YYYY)"]
    df_repl1a_er_nit_unif = df_repl1a_er_nit["% Uni"]
    df_repl1a_er_nit_unif_usl = df_repl1a_er_nit["% Uni USL"]
    df_repl1a_er_nit_unif_ucl = df_repl1a_er_nit["% Uni UCL"]
    df_repl1a_er_nit_remarks = df_repl1a_er_nit["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Nit Unif Plot (using Plotly) in Browser     
    return draw_plotly_repl1a_unif_nit_plot(
        x = date_formatter(df_repl1a_er_nit_date), 
        y1 = df_repl1a_er_nit_unif, 
        y2 = df_repl1a_er_nit_unif_usl,
        y3 = df_repl1a_er_nit_unif_ucl,
        remarks = df_repl1a_er_nit_remarks
    )

def repl1a_poly_er_chart():
    df_repl1a_er_poly = excel_file.parse(sht_name_er_poly, skiprows= skiprows_poly)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_poly = df_repl1a_er_poly[sht_er_poly_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_poly['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_poly.dropna(inplace=True)                                              # dropping rows where at least one element is misnitg
    # sht_run.range('A10').options(index=False).value = df_repl1a_er_poly        # show the dataframe values into sheet- 'CP Plot'
 
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for Nit ER & Unif PLot
    df_repl1a_er_poly_date = df_repl1a_er_poly["Date (MM/DD/YYYY)"]
    df_repl1a_er_poly_er = df_repl1a_er_poly["Etch Rate (A/Min)"]
    df_repl1a_er_poly_usl = df_repl1a_er_poly["USL"]
    df_repl1a_er_poly_lsl = df_repl1a_er_poly["LSL"]
    df_repl1a_er_poly_ucl = df_repl1a_er_poly["UCL"]
    df_repl1a_er_poly_lcl = df_repl1a_er_poly["LCL"]
    df_repl1a_er_poly_remarks = df_repl1a_er_poly["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Poly ER Plot (using Plotly) in Browser 
    return draw_plotly_repl1a_er_poly_plot(
        x = date_formatter(df_repl1a_er_poly_date), 
        y1 = df_repl1a_er_poly_er,
        y2 = df_repl1a_er_poly_usl, 
        y3 = df_repl1a_er_poly_lsl,
        y4 = df_repl1a_er_poly_ucl,
        y5 = df_repl1a_er_poly_lcl,
        remarks = df_repl1a_er_poly_remarks
    )

def repl1a_poly_unif_chart():
    df_repl1a_er_poly = excel_file.parse(sht_name_er_poly, skiprows= skiprows_poly)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_poly = df_repl1a_er_poly[sht_er_poly_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_poly['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_poly.dropna(inplace=True)                                              # dropping rows where at least one element is misnitg
    # sht_run.range('A10').options(index=False).value = df_repl1a_er_poly        # show the dataframe values into sheet- 'CP Plot'
 
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for Nit ER & Unif PLot
    df_repl1a_er_poly_date = df_repl1a_er_poly["Date (MM/DD/YYYY)"]
    df_repl1a_er_poly_unif = df_repl1a_er_poly["% Uni"]
    df_repl1a_er_poly_unif_usl = df_repl1a_er_poly["% Uni USL"]
    df_repl1a_er_poly_unif_ucl = df_repl1a_er_poly["% Uni UCL"]
    df_repl1a_er_poly_remarks = df_repl1a_er_poly["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Poly Unif Plot (using Plotly) in Browser     
    return draw_plotly_repl1a_unif_poly_plot(
        x = date_formatter(df_repl1a_er_poly_date), 
        y1 = df_repl1a_er_poly_unif, 
        y2 = df_repl1a_er_poly_unif_usl,
        y3 = df_repl1a_er_poly_unif_ucl,
        remarks = df_repl1a_er_poly_remarks
    )


#********************************************************NIT*************************************************************************************************

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
