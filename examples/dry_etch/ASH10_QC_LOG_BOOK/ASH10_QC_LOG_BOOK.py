# Import packages
import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
from input import *
from dir import *
# import datetime as dt
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
"Description": This function plots CP Chart with 3 traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
"y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_asbe1_cp_plot(x, y1, y2, y3, remarks):
    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'delta-CP',
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
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with 4 traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_asbe1_er_plot(x, y1, y2, y3, y4, y5, remarks):
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
    py.offline.plot(fig, filename= er_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with 2 traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_asbe1_unif_plot(x, y1, y2, y3, remarks):
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
    py.offline.plot(fig, filename= unif_plot_html_file)


#====================================================================================================================================================================
#####################################################################################################################################################################
def init():
    # Initialize the workbook
    # wb = xw.Book.caller()
    wb = xw.Book('ASH10_QC_LOG_BOOK.xlsm')
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    sht_asbe1_cp = wb.sheets[sht_name_cp]
    sht_asbe1_er = wb.sheets[sht_name_er]
    sht_run = wb.sheets['RUN_code']     # for testing purpose
    #****************************************************************************************************************************************************************
    x_coord_pr = sht_asbe1_er.range(x_coord_pr_range).value
    y_coord_pr = sht_asbe1_er.range(y_coord_pr_range).value
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    excel_file = pd.ExcelFile(excel_file_directory)

    return wb, sht_asbe1_cp, sht_asbe1_er, sht_run, x_coord_pr, y_coord_pr, excel_file


def button_run():
    wb, sht_asbe1_cp, sht_asbe1_er, sht_run, x_coord_pr, y_coord_pr, excel_file = init()

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_asbe1_cp = sht_asbe1_cp.range('A10').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- sht_name_cp
    df_asbe1_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_asbe1_cp = df_asbe1_cp[sht_cp_columns]        # The final dataframe with required columns
    df_asbe1_cp = df_asbe1_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_asbe1_plot_cp.range('A25').options(index=False).value = df_asbe1_cp   	    # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_asbe1_cp_date = df_asbe1_cp["Date (MM/DD/YYYY)"]
    df_asbe1_cp_delta_cp = df_asbe1_cp["delta CP"]
    df_asbe1_cp_usl = df_asbe1_cp["USL"]
    df_asbe1_cp_ucl = df_asbe1_cp["UCL"]
    df_asbe1_cp_remarks = df_asbe1_cp["Remarks"]

    # #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_asbe1_cp_plot(
        x = date_formatter(df_asbe1_cp_date), 
        y1 = df_asbe1_cp_delta_cp, 
        y2 = df_asbe1_cp_usl, 
        y3 = df_asbe1_cp_ucl,
        remarks = df_asbe1_cp_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for ER & Unif Plot
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    df_asbe1_er = excel_file.parse(sht_name_er, skiprows=skiprows_pr)                            # copy a sheet and paste into another sheet and skiprows 8
    df_asbe1_er['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with 'NIL'
    df_asbe1_er = df_asbe1_er[sht_er_columns]             # The final Dataframe with 5 columns for plot: x-1, y-4
    df_asbe1_er = df_asbe1_er.dropna()                                              # dropping rows where at least one element is missing
    # sht_asbe1_plot_er.range('A28').options(index=False).value = df_asbe1_er        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for ER & Unif PLot
    df_asbe1_er_date = df_asbe1_er["Date (MM/DD/YYYY)"]
    df_asbe1_er_er = df_asbe1_er["Etch Rate (A/Min)"]
    df_asbe1_er_usl = df_asbe1_er["USL"]
    df_asbe1_er_lsl = df_asbe1_er["LSL"]
    df_asbe1_er_ucl = df_asbe1_er["UCL"]
    df_asbe1_er_lcl = df_asbe1_er["LCL"]
    df_asbe1_er_unif = df_asbe1_er["% Uni"]
    df_asbe1_er_unif_usl = df_asbe1_er["% Uni USL"]
    df_asbe1_er_unif_ucl = df_asbe1_er["% Uni UCL"]
    df_asbe1_er_remarks = df_asbe1_er["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------        
    # Draw ER Plot (using Plotly) in Browser 
    draw_plotly_asbe1_er_plot(
        x = date_formatter(df_asbe1_er_date), 
        y1 = df_asbe1_er_er,
        y2 = df_asbe1_er_usl, 
        y3 = df_asbe1_er_lsl, 
        y4 = df_asbe1_er_ucl,
        y5 = df_asbe1_er_lcl,
        remarks = df_asbe1_er_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------        
    # Draw Unif Plot (using Plotly) in Browser     
    draw_plotly_asbe1_unif_plot(
        x = date_formatter(df_asbe1_er_date), 
        y1 = df_asbe1_er_unif, 
        y2 = df_asbe1_er_unif_usl,
        y3 = df_asbe1_er_unif_ucl,
        remarks = df_asbe1_er_remarks
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