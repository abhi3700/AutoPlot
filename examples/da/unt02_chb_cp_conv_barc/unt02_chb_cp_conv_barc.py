# Import packages
import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
from dir import *
from input import *
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
"Description": This function plots CP Chart with traces v/s Date.
"x": Date (x-axis) for CP Chart
"x_1": Date (x-axis) Conventional for CP Chart
"x_2": Date (x-axis) BARC for CP Chart
"y1_1": Delta-CP 0.16u (y-axis) Conventional for CP Chart
"y1_2": Delta-CP 0.16u (y-axis) BARC for CP Chart
"y2": Delta-CP 0.5u (y-axis) for CP Chart
"y3": Delta-CP AC (y-axis) for CP Chart
"y4": USL (y-axis) for CP Chart
"y5": UCL (y-axis) for CP Chart,
"remarks": remarks for entire dataframe
"remarks_1": remarks for conventional CP
"remarks_2": remarks for barc CP
"""
def draw_plotly_resp1b_cp_plot(x, x1_1, x1_2, y1_1, y1_2, y2, y3, y4, y5, remarks, remarks_1, remarks_2):
    trace1_1 = go.Scatter(
            x = x1_1,
            y = y1_1,
            name = 'Delta CP 0.16u CONV',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = line_color,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks_1
    )

    trace1_2 = go.Scatter(
            x = x1_2,
            y = y1_2,
            name = 'Delta CP 0.16u BARC',
            mode = 'lines+markers',
            line = dict(
                    color = line_color_4,
                    width = 2),
            marker = dict(
                    color = marker_color_4,
                    size = 8,
                    line = dict(
                        color = marker_border_color,
                        width = 0.5),
                    ),
            text = remarks_2
    )

    trace2 = go.Scatter(
            x = x,
            y = y2,
            name = 'Delta CP 0.5u',
            mode = 'markers',
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
            mode = 'markers',
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

    data = [trace2, trace1_1, trace1_2, trace3, trace4, trace5]
    layout = dict(
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_plot_html_file)


#====================================================================================================================================================================
#####################################################################################################################################################################
def init():
    # Initialize the workbook
    # wb = xw.Book.caller()
    wb = xw.Book('unt02_chb_cp_conv_barc.xlsm')
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_resp1b_cp = wb.sheets[sht_name_cp]
    # sht_resp1b_er_barc_pr_teos = wb.sheets[sht_name_er_barc_pr_teos]
    # sht_resp1b_sin = wb.sheets[sht_name_er_sin]
    sht_run = wb.sheets['RUN_code']     # for testing purpose
    sht_run_2 = wb.sheets['RUN_code_2']     # for testing purpose
    #****************************************************************************************************************************************************************
    # x_coord_barc = sht_resp1b_er_barc_pr_teos.range(x_coord_barc_range).value
    # y_coord_barc = sht_resp1b_er_barc_pr_teos.range(y_coord_barc_range).value
    # x_coord_pr = sht_resp1b_er_barc_pr_teos.range(x_coord_pr_range).value
    # y_coord_pr = sht_resp1b_er_barc_pr_teos.range(y_coord_pr_range).value
    # x_coord_teos = sht_resp1b_er_barc_pr_teos.range(x_coord_teos_range).value
    # y_coord_teos = sht_resp1b_er_barc_pr_teos.range(y_coord_teos_range).value
    # x_coord_sin = sht_resp1b_sin.range(x_coord_sin_range).value
    # y_coord_sin = sht_resp1b_sin.range(y_coord_sin_range).value
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    excel_file = pd.ExcelFile(excel_file_directory)

    return sht_resp1b_cp, sht_run, sht_run_2, excel_file

def button_run():
    sht_resp1b_cp, sht_run, sht_run_2, excel_file = init()
    # excel_file = pd.ExcelFile("./unt02_chb_cp_conv_barc.xlsm")

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_resp1b_cp = excel_file.parse(sht_name_cp, skiprows=skiprows_cp)                            # copy a sheet and paste into another sheet and skiprows 8
    df_resp1b_cp = df_resp1b_cp[sht_cp_columns]        # The final dataframe with required columns
    df_resp1b_cp['Remarks'].fillna('.', inplace=True)        # replacing the empty cells with '.'
    df_resp1b_cp['Delta CP 0.5u'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_resp1b_cp['Delta CP AC'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_resp1b_cp = df_resp1b_cp.dropna()                                              # dropping rows where at least one element is missing
    df_resp1b_cp_conv = df_resp1b_cp[df_resp1b_cp["CP type"] == 'CONVENTIONAL CP']          # dataframe - 'conventional CP'
    df_resp1b_cp_barc = df_resp1b_cp[df_resp1b_cp["CP type"] == 'BARC CP']                  # dataframe - 'barc CP'
    
    sht_run.range('A70').options(index=False).value = df_resp1b_cp         # show the dataframe values into sheet- 'CP Plot'
    sht_run.range('A1').options(index=False).value = df_resp1b_cp_conv         # show the dataframe values into sheet- 'CP Plot'
    sht_run_2.range('A1').options(index=False).value = df_resp1b_cp_barc         # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_resp1b_cp_date = df_resp1b_cp["Date (MM/DD/YYYY)"]
    df_resp1b_cp_date_conv = df_resp1b_cp_conv["Date (MM/DD/YYYY)"]
    df_resp1b_cp_date_barc = df_resp1b_cp_barc["Date (MM/DD/YYYY)"]
    df_resp1b_cp_delta_cp_1_conv = df_resp1b_cp_conv["Delta CP 0.16u"]
    df_resp1b_cp_delta_cp_1_barc = df_resp1b_cp_barc["Delta CP 0.16u"]
    df_resp1b_cp_delta_cp_2 = df_resp1b_cp["Delta CP 0.5u"]
    df_resp1b_cp_delta_cp_3 = df_resp1b_cp["Delta CP AC"]
    df_resp1b_cp_usl = df_resp1b_cp["USL"]
    df_resp1b_cp_ucl = df_resp1b_cp["UCL"]
    df_resp1b_cp_remarks = df_resp1b_cp["Remarks"]
    df_resp1b_cp_remarks_conv = df_resp1b_cp_conv["Remarks"]
    df_resp1b_cp_remarks_barc = df_resp1b_cp_barc["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_resp1b_cp_plot(
        x = date_formatter(df_resp1b_cp_date),
        x1_1 = date_formatter(df_resp1b_cp_date_conv),
        x1_2 = date_formatter(df_resp1b_cp_date_barc),
        y1_1 = df_resp1b_cp_delta_cp_1_conv, 
        y1_2 = df_resp1b_cp_delta_cp_1_barc, 
        y2 = df_resp1b_cp_delta_cp_2, 
        y3 = df_resp1b_cp_delta_cp_3, 
        y4 = df_resp1b_cp_usl, 
        y5 = df_resp1b_cp_ucl,
        remarks = df_resp1b_cp_remarks,
        remarks_1= df_resp1b_cp_remarks_conv,
        remarks_2= df_resp1b_cp_remarks_barc
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
if __name__ == "__main__":
    button_run()