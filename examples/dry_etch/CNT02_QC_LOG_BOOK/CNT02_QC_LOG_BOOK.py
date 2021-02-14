# Import packages
import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
import datetime as dt
from input import *
from dir import *
# import win32api
# import os
# from pathlib import Path




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
"y1": Delta-CP 0.2u (y-axis) for CP Chart
"y2": Delta-CP 0.5u (y-axis) for CP Chart
"y3": Delta-CP AC (y-axis) for CP Chart
"y4": USL (y-axis) for CP Chart
"y5": UCL (y-axis) for CP Chart
"""
def draw_plotly_reml1a_cp_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'Delta CP 0.2u',
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
            title = cp_cha_plot_title,
            xaxis = dict(title= cp_cha_plot_xlabel),
            yaxis = dict(title= cp_cha_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_cha_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
"Description": This function plots CP Chart with `cp_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP 0.2u (y-axis) for CP Chart
"y2": Delta-CP 0.5u (y-axis) for CP Chart
"y3": Delta-CP AC (y-axis) for CP Chart
"y4": USL (y-axis) for CP Chart
"y5": UCL (y-axis) for CP Chart
"""
def draw_plotly_reml1c_cp_plot(x, y1, y2, y3, y4, y5, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'Delta CP 0.2u',
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
            title = cp_chc_plot_title,
            xaxis = dict(title= cp_chc_plot_xlabel),
            yaxis = dict(title= cp_chc_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_chc_plot_html_file)

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
def draw_plotly_reml1a_er_pr_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_cha_pr_plot_title,
            xaxis = dict(title= er_cha_pr_plot_xlabel),
            yaxis = dict(title= er_cha_pr_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_cha_pr_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reml1a_unif_pr_plot(x, y1, y2, y3, remarks):
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
            title = unif_cha_pr_plot_title,
            xaxis = dict(title= unif_cha_pr_plot_xlabel),
            yaxis = dict(title= unif_cha_pr_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_cha_pr_plot_html_file)

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
def draw_plotly_reml1c_er_pr_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_chc_pr_plot_title,
            xaxis = dict(title= er_chc_pr_plot_xlabel),
            yaxis = dict(title= er_chc_pr_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_chc_pr_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reml1c_unif_pr_plot(x, y1, y2, y3, remarks):
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
            title = unif_chc_pr_plot_title,
            xaxis = dict(title= unif_chc_pr_plot_xlabel),
            yaxis = dict(title= unif_chc_pr_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_chc_pr_plot_html_file)

#====================================================================================================================================================================
#####################################################################################################################################################################
def init():
    # Initialize the workbook
    # wb = xw.Book.caller()
    wb = xw.Book('CNT02_QC_LOG_BOOK.xlsm')
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_reml1_cp = wb.sheets[sht_name_cp]
    sht_reml1a_er_pr = wb.sheets[sht_name_er_cha_pr]
    sht_reml1c_er_pr = wb.sheets[sht_name_er_chc_pr]
    sht_run = wb.sheets['RUN_code']     # for testing purpose
    #****************************************************************************************************************************************************************
    x_coord_pr_a = sht_reml1a_er_pr.range(x_coord_pr_a_range).value
    y_coord_pr_a = sht_reml1a_er_pr.range(y_coord_pr_a_range).value
    x_coord_pr_c = sht_reml1c_er_pr.range(x_coord_pr_c_range).value
    y_coord_pr_c = sht_reml1c_er_pr.range(y_coord_pr_c_range).value
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    excel_file = pd.ExcelFile(excel_file_directory)

    return wb, sht_reml1_cp, sht_reml1a_er_pr, sht_reml1c_er_pr, sht_run, x_coord_pr_a, y_coord_pr_a, x_coord_pr_c, y_coord_pr_c, excel_file


def button_run():
    wb, sht_reml1_cp, sht_reml1a_er_pr, sht_reml1c_er_pr, sht_run, x_coord_pr_a, y_coord_pr_a, x_coord_pr_c, y_coord_pr_c, excel_file = init()

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_reml1_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reml1_cp = df_reml1_cp[sht_cp_columns]        # The final dataframe with required columns
    df_reml1_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_reml1_cp['Delta CP 0.5u'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1_cp['Delta CP AC'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1_cp = df_reml1_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_run.range('A46').options(index=False).value = df_reml1_cp   	    # show the dataframe values into sheet- 'CP Plot'
    df_reml1a_cp = df_reml1_cp[df_reml1_cp["Chamber"] == 'DPS']					# dataframe for DPS (chamber A)
    # sht_run.range('A46').options(index=False).value = df_reml1a_cp           # show the dataframe values into sheet- 'CP Plot'
    df_reml1c_cp = df_reml1_cp[df_reml1_cp["Chamber"] == 'ASP']					# dataframe for ASP (chamber C)
    # sht_run.range('G46').options(index=False).value = df_reml1c_cp           # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reml1a_cp_date = df_reml1a_cp["Date (MM/DD/YYYY)"]
    df_reml1a_cp_delta_cp_1 = df_reml1a_cp["Delta CP 0.2u"]
    df_reml1a_cp_delta_cp_2 = df_reml1a_cp["Delta CP 0.5u"]
    df_reml1a_cp_delta_cp_3 = df_reml1a_cp["Delta CP AC"]
    df_reml1a_cp_usl = df_reml1a_cp["USL"]
    df_reml1a_cp_ucl = df_reml1a_cp["UCL"]
    df_reml1a_cp_remarks = df_reml1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reml1c_cp_date = df_reml1c_cp["Date (MM/DD/YYYY)"]
    df_reml1c_cp_delta_cp_1 = df_reml1c_cp["Delta CP 0.2u"]
    df_reml1c_cp_delta_cp_2 = df_reml1c_cp["Delta CP 0.5u"]
    df_reml1c_cp_delta_cp_3 = df_reml1c_cp["Delta CP AC"]
    df_reml1c_cp_usl = df_reml1c_cp["USL"]
    df_reml1c_cp_ucl = df_reml1c_cp["UCL"]
    df_reml1c_cp_remarks = df_reml1c_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw REML1A CP Plot (using Plotly) in Browser 
    draw_plotly_reml1a_cp_plot(
        x = date_formatter(df_reml1a_cp_date), 
        y1 = df_reml1a_cp_delta_cp_1,
        y2 = df_reml1a_cp_delta_cp_2,
        y3 = df_reml1a_cp_delta_cp_3,
        y4 = df_reml1a_cp_usl, 
        y5 = df_reml1a_cp_ucl,
        remarks = df_reml1a_cp_remarks
        )

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw REML1C CP Plot (using Plotly) in Browser 
    draw_plotly_reml1c_cp_plot(
        x = date_formatter(df_reml1c_cp_date), 
        y1 = df_reml1c_cp_delta_cp_1, 
        y2 = df_reml1c_cp_delta_cp_2, 
        y3 = df_reml1c_cp_delta_cp_3, 
        y4 = df_reml1c_cp_usl, 
        y5 = df_reml1c_cp_ucl,
        remarks = df_reml1c_cp_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Ch A PR i.e. DPS chamber ER & Unif PLot
    df_reml1a_er_pr = excel_file.parse(sht_name_er_cha_pr, skiprows=skiprows_pr_a)                            # copy a sheet and paste into another sheet and skiprows 8
    
    df_reml1a_er_pr = df_reml1a_er_pr[sht_er_reml1a_pr_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reml1a_er_pr['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_reml1a_er_pr = df_reml1a_er_pr.dropna()                                              # dropping rows where at least one element is missing
    # sht_run.range('A28').options(index=False).value = df_reml1a_er_pr        # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for REML1A PR ER & Unif PLot
    df_reml1a_er_pr_date = df_reml1a_er_pr["Date (MM/DD/YYYY)"]
    df_reml1a_er_pr_er = df_reml1a_er_pr["Etch Rate (A/Min)"]
    df_reml1a_er_pr_usl = df_reml1a_er_pr["USL"]
    df_reml1a_er_pr_lsl = df_reml1a_er_pr["LSL"]
    df_reml1a_er_pr_ucl = df_reml1a_er_pr["UCL"]
    df_reml1a_er_pr_lcl = df_reml1a_er_pr["LCL"]
    df_reml1a_er_pr_unif = df_reml1a_er_pr["% Uni"]
    df_reml1a_er_pr_unif_usl = df_reml1a_er_pr["% Uni USL"]
    df_reml1a_er_pr_unif_ucl = df_reml1a_er_pr["% Uni UCL"]
    df_reml1a_er_pr_remarks = df_reml1a_er_pr["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw REML1A PR ER Plot (using Plotly) in Browser 
    draw_plotly_reml1a_er_pr_plot(
        x = date_formatter(df_reml1a_er_pr_date), 
        y1 = df_reml1a_er_pr_er,
        y2 = df_reml1a_er_pr_usl, 
        y3 = df_reml1a_er_pr_lsl,
        y4 = df_reml1a_er_pr_ucl,
        y5 = df_reml1a_er_pr_lcl,
        remarks = df_reml1a_er_pr_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw REML1A PR Unif Plot (using Plotly) in Browser     
    draw_plotly_reml1a_unif_pr_plot(
        x = date_formatter(df_reml1a_er_pr_date), 
        y1 = df_reml1a_er_pr_unif, 
        y2 = df_reml1a_er_pr_unif_usl,
        y3 = df_reml1a_er_pr_unif_ucl,
        remarks = df_reml1a_er_pr_remarks
        )


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Ch C PR i.e. ASP Chamber ER & Unif PLot
    df_reml1c_er_pr = excel_file.parse(sht_name_er_chc_pr, skiprows=skiprows_pr_a)                            # copy a sheet and paste into another sheet and skiprows 8
    
    df_reml1c_er_pr = df_reml1c_er_pr[sht_er_reml1c_pr_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reml1c_er_pr['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_reml1c_er_pr = df_reml1c_er_pr.dropna()                                              # dropping rows where at least one element is missing
    # sht_run.range('A28').options(index=False).value = df_reml1c_er_pr        # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for REML1C PR ER & Unif PLot
    df_reml1c_er_pr_date = df_reml1c_er_pr["Date (MM/DD/YYYY)"]
    df_reml1c_er_pr_er = df_reml1c_er_pr["Etch Rate (A/Min)"]
    df_reml1c_er_pr_usl = df_reml1c_er_pr["USL"]
    df_reml1c_er_pr_lsl = df_reml1c_er_pr["LSL"]
    df_reml1c_er_pr_ucl = df_reml1c_er_pr["UCL"]
    df_reml1c_er_pr_lcl = df_reml1c_er_pr["LCL"]
    df_reml1c_er_pr_unif = df_reml1c_er_pr["% Uni"]
    df_reml1c_er_pr_unif_usl = df_reml1c_er_pr["% Uni USL"]
    df_reml1c_er_pr_unif_ucl = df_reml1c_er_pr["% Uni UCL"]
    df_reml1c_er_pr_remarks = df_reml1c_er_pr["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw REML1C PR ER Plot (using Plotly) in Browser 
    draw_plotly_reml1c_er_pr_plot(
        x = date_formatter(df_reml1c_er_pr_date), 
        y1 = df_reml1c_er_pr_er,
        y2 = df_reml1c_er_pr_usl, 
        y3 = df_reml1c_er_pr_lsl,
        y4 = df_reml1c_er_pr_ucl,
        y5 = df_reml1c_er_pr_lcl,
        remarks = df_reml1c_er_pr_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw REML1C PR Unif Plot (using Plotly) in Browser     
    draw_plotly_reml1c_unif_pr_plot(
        x = date_formatter(df_reml1c_er_pr_date), 
        y1 = df_reml1c_er_pr_unif, 
        y2 = df_reml1c_er_pr_unif_usl,
        y3 = df_reml1c_er_pr_unif_ucl,
        remarks = df_reml1c_er_pr_remarks
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
if __name__ == "__main__":
    button_run()




