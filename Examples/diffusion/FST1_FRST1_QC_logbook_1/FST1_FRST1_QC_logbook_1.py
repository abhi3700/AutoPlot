# Import packages
import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
from input import *
# import datetime as dt
# import win32api
# import os
# from pathlib import Path




#==================================================================================================================================================================
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
"Description": This function plots Thickness Chart with traces v/s Date.
"x": Date (x-axis) for Thickness Chart
"y1": Top thick (y-axis) for  Chart
"y2": Center thick (y-axis) for Thickness Chart
"y3": Bottom thick (y-axis) for Thickness Chart
"y4": USL (y-axis) for Thickness Chart
"y5": LSL (y-axis) for Thickness Chart
"y6": UCL (y-axis) for Thickness Chart
"y7": LCL (y-axis) for Thickness Chart
"""
def draw_plotly_frst1_thick_plot(x, y1, y2, y3, y4, y5, y6, y7, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = thick_plot_trace1_name,
            mode = thick_plot_trace1_mode_top,
            line = dict(
                    color = thick_plot_trace1_line_color_top,
                    width = 2),
            marker = dict(
                    color = thick_plot_trace1_marker_color_top,
                    size = 8,
                    line = dict(
                        color = thick_plot_trace1_marker_border_color_top,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = thick_plot_trace2_name,
            mode = thick_plot_trace2_mode_cnt,
            line = dict(
                    color = thick_plot_trace2_line_color_cnt,
                    width = 2),
            marker = dict(
                    color = thick_plot_trace2_marker_color_cnt,
                    size = 8,
                    line = dict(
                        color = thick_plot_trace2_marker_border_color_cnt,
                        width = 0.5),
                    symbol = 'diamond'
                    ),
            # text = remarks
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = thick_plot_trace3_name,
            mode = thick_plot_trace3_mode_btm,
            line = dict(
                    color = thick_plot_trace3_line_color_btm,
                    width = 2),
            marker = dict(
                    color = thick_plot_trace3_marker_color_btm,
                    size = 8,
                    line = dict(
                        color = thick_plot_trace3_marker_border_color_btm,
                        width = 0.5),
                    symbol = 'x',
                    ),
            # text = remarks
    )
    go.Scatter

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
            name = 'LSL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
    )

    trace6 = go.Scatter(
            x = x,
            y = y6,
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    trace7 = go.Scatter(
            x = x,
            y = y7,
            name = 'LCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4, trace5, trace6, trace7]
    layout = dict(
            title = thick_plot_title,
            xaxis = dict(title= thick_plot_xlabel),
            yaxis = dict(title= thick_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= thick_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": TOP (y-axis) for Unif Chart
"y2": CNT (y-axis) for Unif Chart
"y3": BTM (y-axis) for Unif Chart
"y4": UCL (y-axis) for Unif Chart
"""
def draw_plotly_frst1_unif_plot(x, y1, y2, y3, y4, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = unif_plot_trace1_name,
            mode = unif_plot_trace1_mode_top,
            line = dict(
                    color = unif_plot_trace1_line_color_top,
                    width = 2),
            marker = dict(
                    color = unif_plot_trace1_marker_color_top,
                    size = 8,
                    line = dict(
                        color = unif_plot_trace1_marker_border_color_top,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = unif_plot_trace2_name,
            mode = unif_plot_trace2_mode_cnt,
            line = dict(
                    color = unif_plot_trace2_line_color_cnt,
                    width = 2),
            marker = dict(
                    color = unif_plot_trace2_marker_color_cnt,
                    size = 8,
                    line = dict(
                        color = unif_plot_trace2_marker_border_color_cnt,
                        width = 0.5),
                    symbol = 'diamond'
                    ),
            # text = remarks
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = unif_plot_trace3_name,
            mode = unif_plot_trace3_mode_btm,
            line = dict(
                    color = unif_plot_trace3_line_color_btm,
                    width = 2),
            marker = dict(
                    color = unif_plot_trace3_marker_color_btm,
                    size = 8,
                    line = dict(
                        color = unif_plot_trace3_marker_border_color_btm,
                        width = 0.5),
                    symbol = 'x'
                    ),
            # text = remarks
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

    data = [trace1, trace2, trace3, trace4]
    layout = dict(
            title = unif_plot_title,
            xaxis = dict(title= unif_plot_xlabel),
            yaxis = dict(title= unif_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": TOP (y-axis) for CP Chart
"y2": CNT (y-axis) for CP Chart
"y3": BTM (y-axis) for CP Chart
"y4": UCL (y-axis) for CP Chart
"""
def draw_plotly_frst1_cp_plot(x, y1, y2, y3, y4, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = cp_plot_trace1_name,
            mode = cp_plot_trace1_mode_top,
            line = dict(
                    color = cp_plot_trace1_line_color_top,
                    width = 2),
            marker = dict(
                    color = cp_plot_trace1_marker_color_top,
                    size = 8,
                    line = dict(
                        color = cp_plot_trace1_marker_border_color_top,
                        width = 0.5),
                    ),
            text = remarks
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = cp_plot_trace2_name,
            mode = cp_plot_trace2_mode_cnt,
            line = dict(
                    color = cp_plot_trace2_line_color_cnt,
                    width = 2),
            marker = dict(
                    color = cp_plot_trace2_marker_color_cnt,
                    size = 8,
                    line = dict(
                        color = cp_plot_trace2_marker_border_color_cnt,
                        width = 0.5),
                    symbol = 'diamond'
                    ),
            # text = remarks
    )

    trace3 = go.Scatter(
            x = x,
            y = y3,
            name = cp_plot_trace3_name,
            mode = cp_plot_trace3_mode_btm,
            line = dict(
                    color = cp_plot_trace3_line_color_btm,
                    width = 2),
            marker = dict(
                    color = cp_plot_trace3_marker_color_btm,
                    size = 8,
                    line = dict(
                        color = cp_plot_trace3_marker_border_color_btm,
                        width = 0.5),
                    symbol = 'x'
                    ),
            # text = remarks
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

    data = [trace1, trace2, trace3, trace4]
    layout = dict(
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_plot_html_file)




#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_frst1 = wb.sheets[sht_name]
    sht_run_code = wb.sheets['RUN_code'] 
    excel_file_sht = pd.ExcelFile(excel_file_directory)
    df_frst1 = excel_file_sht.parse(sht_name, skiprows=2)


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Thickness Plot
    df_frst1_thick = df_frst1[sht_thick_columns]        # The final dataframe with required columns
    df_frst1_thick['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_frst1_thick['Remarks'] = pd.Series(df_frst1_thick['Remarks']).fillna(method='ffill')       # fill merged cells with same values
    df_frst1_thick['Date'] = pd.Series(df_frst1_thick['Date']).fillna(method='ffill')       # fill merged cells with same values

    df_frst1_thick = df_frst1_thick.dropna()                                              # dropping rows where at least one element is missing
    # sht_run_code.range('A13').options(index=False).value = df_frst1_thick         # show the dataframe values into sheet- 'RUN_code'
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for Thickness PLot
    df_frst1_thick_date = df_frst1_thick["Date"]
    df_frst1_thick_top = df_frst1_thick["Top Avg"]
    df_frst1_thick_cnt = df_frst1_thick["Center Avg"]
    df_frst1_thick_btm = df_frst1_thick["BTM Avg"]
    df_frst1_thick_usl = df_frst1_thick["Thick USL"]
    df_frst1_thick_lsl = df_frst1_thick["Thick LSL"]
    df_frst1_thick_ucl = df_frst1_thick["Thick UCL"]
    df_frst1_thick_lcl = df_frst1_thick["Thick LCL"]
    df_frst1_thick_remarks = df_frst1_thick["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Poly Thickness Plot (using Plotly) in Browser 
    draw_plotly_frst1_thick_plot(
        x = date_formatter(df_frst1_thick_date), 
        y1 = df_frst1_thick_top,
        y2 = df_frst1_thick_cnt, 
        y3 = df_frst1_thick_btm,
        y4 = df_frst1_thick_usl,
        y5 = df_frst1_thick_lsl,
        y6 = df_frst1_thick_ucl,
        y7 = df_frst1_thick_lcl,
        remarks = df_frst1_thick_remarks
        )
   
    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Unif Plot
    df_frst1_unif = df_frst1[sht_unif_columns]        # The final dataframe with required columns
    df_frst1_unif['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_frst1_unif['Remarks'] = pd.Series(df_frst1_unif['Remarks']).fillna(method='ffill')
    df_frst1_unif['Date'] = pd.Series(df_frst1_unif['Date']).fillna(method='ffill')
    df_frst1_unif = df_frst1_unif.dropna()                                              # dropping rows where at least one element is missing
    # sht_run_code.range('A13').options(index=False).value = df_frst1_unif         # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_frst1_unif_date = df_frst1_unif["Date"]
    df_frst1_unif_top = df_frst1_unif["Unif TOP"]
    df_frst1_unif_cnt = df_frst1_unif["Unif CNT"]
    df_frst1_unif_btm = df_frst1_unif["Unif BTM"]
    df_frst1_unif_ucl = df_frst1_unif["Unif UCL"]
    df_frst1_unif_remarks = df_frst1_unif["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_frst1_unif_plot(
        x = date_formatter(df_frst1_unif_date), 
        y1 = df_frst1_unif_top, 
        y2 = df_frst1_unif_cnt, 
        y3 = df_frst1_unif_btm,
        y4 = df_frst1_unif_ucl,
        remarks = df_frst1_unif_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_frst1_cp = df_frst1[sht_cp_columns]        # The final dataframe with required columns
    df_frst1_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_frst1_cp['Remarks'] = pd.Series(df_frst1_cp['Remarks']).fillna(method='ffill')
    df_frst1_cp['Date'] = pd.Series(df_frst1_cp['Date']).fillna(method='ffill')
    df_frst1_cp = df_frst1_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_run_code.range('A13').options(index=False).value = df_frst1_cp         # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_frst1_cp_date = df_frst1_cp["Date"]
    df_frst1_cp_top = df_frst1_cp["CP TOP"]
    df_frst1_cp_cnt = df_frst1_cp["CP CNT"]
    df_frst1_cp_btm = df_frst1_cp["CP BTM"]
    df_frst1_cp_ucl = df_frst1_cp["CP UCL"]
    df_frst1_cp_remarks = df_frst1_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_frst1_cp_plot(
        x = date_formatter(df_frst1_cp_date), 
        y1 = df_frst1_cp_top, 
        y2 = df_frst1_cp_cnt, 
        y3 = df_frst1_cp_btm,
        y4 = df_frst1_cp_ucl,
        remarks = df_frst1_cp_remarks
        )

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



