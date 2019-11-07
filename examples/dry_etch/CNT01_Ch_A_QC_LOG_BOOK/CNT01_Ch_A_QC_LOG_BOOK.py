# Import packages
import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
import win32api         # for message box
import numpy as np
import math
import statistics as stat
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
        a = a.strftime("%m-%d-%Y %H:%M:%S")
        x_fmt.append(a)
    return x_fmt


#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with `cp_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
# "y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_repl1a_cp_plot(x, y1, y2, remarks):
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

    # trace3 = go.Scatter(
    #         x = x,
    #         y = y3,
    #         name = 'UCL',
    #         mode = 'lines',
    #         line = dict(
    #                 color = cl_color,
    #                 width = 3)
    # )

    # data = [trace1, trace2, trace3]
    data = [trace1, trace2]
    layout = dict(
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_plot_html_file)

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
    py.offline.plot(fig, filename= er_nit_plot_html_file)

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
    py.offline.plot(fig, filename= unif_nit_plot_html_file)

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
    py.offline.plot(fig, filename= er_poly_plot_html_file)

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
    py.offline.plot(fig, filename= unif_poly_plot_html_file)

# ==========================================Control limit calculation===================================
"""
TODO: 
+ control limit calculation of last 30 wafers. Each wafer has 'n' points. So, total sample size = 30 * n
- Integrate it on a button inside Excel files

"Description": calculate the control limits (LCL, UCL) of sample (size=30) 
               data points out of total data points
"data_list": total data points
"""
def control_limit_calc(data_list):
    mean = stat.mean(data_list)
    sum_sq_points_mean = 0
    for x in data_list:
        sum_sq_points_mean += (mean - x)**2

    sigma3 = 3 * math.sqrt(sum_sq_points_mean/len(data_list) - 1)
    # print(f'data_list: {data_list}')
    # print(f'sigma3: {sigma3}')
    
    ucl = mean + sigma3
    # print(f'UCL: {ucl}')
    lcl = mean - sigma3
    # print(f'LCL: {lcl}')

    return lcl, ucl


#====================================================================================================================================================================
#####################################################################################################################################################################
# Initialize the workbook
wb = xw.Book.caller()
# wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

#****************************************************************************************************************************************************************
# Define sheets
sht_repl1a_cp = wb.sheets[sht_name_cp]
sht_repl1a_er_nit = wb.sheets[sht_name_er_nit]
sht_repl1a_er_poly = wb.sheets[sht_name_er_poly]
sht_run = wb.sheets['RUN_code']     # for testing purpose
# sht_repl1a_plot_cp = wb.sheets['CP Plot']
# sht_repl1a_plot_er_nit = wb.sheets['Nit Plot']
# sht_repl1a_plot_er_poly = wb.sheets['Poly Plot']
#****************************************************************************************************************************************************************
# Fetch Dataframe for NIT ER & UNif PLot
# data_folder = Path(os.getcwd())
# file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
# excel_file = pd.ExcelFile(file_to_open)

excel_file_sht = pd.ExcelFile(excel_file_directory)

def main():
    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_repl1a_cp = sht_repl1a_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value                                                         # fetch the data from sheet- 'ASBE1-CP'
    df_repl1a_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_cp = df_repl1a_cp[sht_cp_columns]        # The final dataframe with required columns
    df_repl1a_cp.dropna(inplace=True)                                              # dropping rows where at least one element is missing
    # sht_run.range('A10').options(index=False).value = df_repl1a_cp         # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_repl1a_cp_date = df_repl1a_cp["Date (MM/DD/YYYY)"]
    df_repl1a_cp_delta_cp = df_repl1a_cp["delta CP"]
    df_repl1a_cp_usl = df_repl1a_cp["USL"]
    # df_repl1a_cp_ucl = df_repl1a_cp["UCL"]
    df_repl1a_cp_remarks = df_repl1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_repl1a_cp_plot(
        x = date_formatter(df_repl1a_cp_date), 
        y1 = df_repl1a_cp_delta_cp, 
        y2 = df_repl1a_cp_usl, 
        # y3 = df_repl1a_cp_ucl,
        remarks = df_repl1a_cp_remarks
        )

    #****************************************************************************************************************************************************************
    df_repl1a_er_nit = excel_file_sht.parse(sht_name_er_nit, skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_nit = df_repl1a_er_nit[sht_er_nit_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_nit['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
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
    df_repl1a_er_nit_unif = df_repl1a_er_nit["% Uni"]
    df_repl1a_er_nit_unif_usl = df_repl1a_er_nit["% Uni USL"]
    df_repl1a_er_nit_unif_ucl = df_repl1a_er_nit["% Uni UCL"]
    df_repl1a_er_nit_remarks = df_repl1a_er_nit["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Nit ER Plot (using Plotly) in Browser 
    draw_plotly_repl1a_er_nit_plot(
        x = date_formatter(df_repl1a_er_nit_date), 
        y1 = df_repl1a_er_nit_er,
        y2 = df_repl1a_er_nit_usl, 
        y3 = df_repl1a_er_nit_lsl,
        y4 = df_repl1a_er_nit_ucl,
        y5 = df_repl1a_er_nit_lcl,
        remarks = df_repl1a_er_nit_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Nit Unif Plot (using Plotly) in Browser     
    draw_plotly_repl1a_unif_nit_plot(
        x = date_formatter(df_repl1a_er_nit_date), 
        y1 = df_repl1a_er_nit_unif, 
        y2 = df_repl1a_er_nit_unif_usl,
        y3 = df_repl1a_er_nit_unif_ucl,
        remarks = df_repl1a_er_nit_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for POLY ER & Unif PLot  
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    df_repl1a_er_poly = excel_file_sht.parse(sht_name_er_poly, skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_poly = df_repl1a_er_poly[sht_er_poly_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_poly['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
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
    df_repl1a_er_poly_unif = df_repl1a_er_poly["% Uni"]
    df_repl1a_er_poly_unif_usl = df_repl1a_er_poly["% Uni USL"]
    df_repl1a_er_poly_unif_ucl = df_repl1a_er_poly["% Uni UCL"]
    df_repl1a_er_poly_remarks = df_repl1a_er_poly["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Poly ER Plot (using Plotly) in Browser 
    draw_plotly_repl1a_er_poly_plot(
        x = date_formatter(df_repl1a_er_poly_date), 
        y1 = df_repl1a_er_poly_er,
        y2 = df_repl1a_er_poly_usl, 
        y3 = df_repl1a_er_poly_lsl,
        y4 = df_repl1a_er_poly_ucl,
        y5 = df_repl1a_er_poly_lcl,
        remarks = df_repl1a_er_poly_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Poly Unif Plot (using Plotly) in Browser     
    draw_plotly_repl1a_unif_poly_plot(
        x = date_formatter(df_repl1a_er_poly_date), 
        y1 = df_repl1a_er_poly_unif, 
        y2 = df_repl1a_er_poly_unif_usl,
        y3 = df_repl1a_er_poly_unif_ucl,
        remarks = df_repl1a_er_poly_remarks
        )


#===================================================Control limit Calculation========================================================================
################################################################################################################################################################
#********************************************************NIT*************************************************************************************************
def date_search1_clear():   # for NIT
    sht_run.range('U3:EKS3').clear_contents()
    sht_run.range('U3').clear_contents()
    sht_run.range('J3').clear_contents()
    sht_run.range('K3').clear_contents()
    sht_run.range('L3').clear_contents()

'''
NOTE: Here, "Etch Rate (A/Min)" column has been considered because we have to compare the Control limit calculation from 2 methods:
    - M-1: take last 30 QC days. So, total data_sample = (30 * 13) points
    - M-2: take last 30 QC days. So, total data_sample = 30 points (all `ER_avg` taken)
'''
def fetch_date_nit():
    # ----------------------------------------------------------- 
    df_repl1a_er_nit = excel_file_sht.parse(sht_name_er_nit, skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9

    df_repl1a_er_nit = df_repl1a_er_nit[sht_er_nit_cl_columns]      # select desired columns
    df_repl1a_er_nit['Date (MM/DD/YYYY)'].fillna(method='ffill', inplace=True)        # forward fill the empty cells
    df_repl1a_er_nit["Etch Rate (A/Min)"].fillna('', inplace=True)        # replacing the empty cells with ''

    nit_er_list = df_repl1a_er_nit["Etch Rate (A/Min)"].tolist()    # extract the ER column into a list
    nit_er_list = list(filter(None, nit_er_list))        # filter '' i.e. None from the list
    df_repl1a_er_nit.drop(columns=["Etch Rate (A/Min)"], inplace= True)    # drop the ER column from dataframe
    df_repl1a_er_nit.dropna(inplace=True)                                              # dropping rows where at least one element is missing
    df_repl1a_er_nit = df_repl1a_er_nit[df_repl1a_er_nit['Site'] == 'ER_point\n']
    df_repl1a_er_nit.insert(loc= len(df_repl1a_er_nit.columns), column= "Etch Rate (A/Min)", value= nit_er_list)    # insert `nit_er_list` into the last col of current dataframe
    # sht_run.range('A20').options(index=False).value = df_repl1a_er_nit        # show the dataframe values into sheet- 'RUN_code'

    # populate the Date column cells with date
    sht_run.range('U3:EKS3').clear_contents()      # clear content only
    sht_run.range('U3').value = df_repl1a_er_nit['Date (MM/DD/YYYY)'].tolist()
    return df_repl1a_er_nit

def button_control_limit_calc_nit():
    df_repl1a_er_nit = fetch_date_nit()

    search_date_in = sht_run.range('J3').value     # input -- to be entered into search box in 'RUN_code' sheet
    df_repl1a_er_nit.index = pd.RangeIndex(len(df_repl1a_er_nit.index))     # reset index 
    index_no = df_repl1a_er_nit[df_repl1a_er_nit['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in datframe column
    
    lcl = 0     # initialize LCL
    ucl = 0     # initialize UCL
    if index_no != []:
        df_upto_search = df_repl1a_er_nit.iloc[0:index_no[-1]+1]     # returns a dataframe from index 0 to search_index. That's why 1 is added.
        if len(df_upto_search) > 30:    # ensure that the length of dataframe is min. 30
            data_last30 = []
            for col in sht_er_nit_cl_columns[2:]:
                data_last30 += df_upto_search[col].tolist()[-N_cl:]     # join the list with last 30 elements of site_1 to site_13
            lcl, ucl = control_limit_calc(data_last30)
            sht_run.range('K3').value = lcl
            sht_run.range('L3').value = ucl
        else:
            win32api.MessageBox(wb.app.hwnd, "There is lesser QC data points available for calculating Control limits.", "Search by Date")         

    elif index_no == []:
        win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    elif sht_run.range('J13').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter the Date in the search box", "Search by Date")
    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Date doesn't exist.", "Search by Date")

    return lcl, ucl, search_date_in

def change_control_limit_nit():
    lcl, ucl, search_date_in = button_control_limit_calc_nit()
    
    if lcl !=0 and ucl !=0 and search_date_in !=0:
        df_repl1a_er_nit = excel_file_sht.parse(sht_name_er_nit, skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
        
        df_repl1a_er_nit = df_repl1a_er_nit[sht_er_nit_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
        # sht_run.range('A23').options(index=False).value = df_repl1a_er_nit        # show the dataframe values into sheet- 'CP Plot'

        index_no = df_repl1a_er_nit[df_repl1a_er_nit['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in datframe column
        if index_no != []:
            # add 11 to get the exact cell's excel row no.
            index_no_final = 11 + index_no[-1]     # consider the latest/last index_no, when there is a failure in QC and it is done multiple times.
            for i in range(0, 300, 3):            
                sht_repl1a_er_nit.range('AB'+str(index_no_final + i)).value = lcl
                sht_repl1a_er_nit.range('AC'+str(index_no_final + i)).value = ucl
            win32api.MessageBox(wb.app.hwnd, "Now, save the excel file & then press RUN button to generate plots.", "Plot chart with new LCL, UCL")         
        elif index_no == []:
            win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    else:
        win32api.MessageBox(wb.app.hwnd, "Either LCL or UCL value is zero OR there is no date mentioned.", "Change control limit")

#********************************************************POLY*************************************************************************************************
def date_search2_clear():   # for POLY
    sht_run.range('U8:EKP8').clear_contents()
    sht_run.range('U8').clear_contents()
    sht_run.range('J13').clear_contents()
    sht_run.range('K13').clear_contents()
    sht_run.range('L13').clear_contents()

def fetch_date_poly():
    df_repl1a_er_poly = excel_file_sht.parse(sht_name_er_poly, skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9

    df_repl1a_er_poly = df_repl1a_er_poly[sht_er_poly_cl_columns]      # select desired columns
    df_repl1a_er_poly['Date (MM/DD/YYYY)'].fillna(method='ffill', inplace=True)        # forward fill the empty cells
    df_repl1a_er_poly["Etch Rate (A/Min)"].fillna('', inplace=True)        # replacing the empty cells with ''

    poly_er_list = df_repl1a_er_poly["Etch Rate (A/Min)"].tolist()    # extract the ER column into a list
    poly_er_list = list(filter(None, poly_er_list))        # filter '' i.e. None from the list
    df_repl1a_er_poly.drop(columns=["Etch Rate (A/Min)"], inplace= True)    # drop the ER column from dataframe
    df_repl1a_er_poly.dropna(inplace=True)                                              # dropping rows where at least one element is missing
    df_repl1a_er_poly = df_repl1a_er_poly[df_repl1a_er_poly['Site'] == 'ER_point\n']
    df_repl1a_er_poly.insert(loc= len(df_repl1a_er_poly.columns), column= "Etch Rate (A/Min)", value= poly_er_list)    # insert `poly_er_list` into the last col of current dataframe
    sht_run.range('U24').options(index=False).value = df_repl1a_er_poly        # show the dataframe values into sheet- 'RUN_code'

    # populate the Date column cells with date
    sht_run.range('U8:EKS8').clear_contents()      # clear content only
    sht_run.range('U8').value = df_repl1a_er_poly['Date (MM/DD/YYYY)'].tolist()
    return df_repl1a_er_poly


def button_control_limit_calc_poly():
    df_repl1a_er_poly = fetch_date_poly()

    search_date_in = sht_run.range('J13').value     # input -- to be entered into search box in 'RUN_code' sheet
    df_repl1a_er_poly.index = pd.RangeIndex(len(df_repl1a_er_poly.index))     # reset index 
    index_no = df_repl1a_er_poly[df_repl1a_er_poly['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in datframe column
    
    lcl = 0     # initialize LCL
    ucl = 0     # initialize UCL
    if index_no != []:
        df_upto_search = df_repl1a_er_poly.iloc[0:index_no[-1]+1]     # returns a dataframe from index 0 to search_index. That's why 1 is added.

        if len(df_upto_search) > 30:    # ensure that the length of dataframe is min. 30
            data_last30 = []
            for col in sht_er_poly_cl_columns[2:]:
                data_last30 += df_upto_search[col].tolist()[-N_cl:]     # join the list with last 30 elements of site_1 to site_17
            
            lcl, ucl = control_limit_calc(data_last30)
            sht_run.range('K13').value = lcl
            sht_run.range('L13').value = ucl
        else:
            win32api.MessageBox(wb.app.hwnd, "There is lesser QC data points available for calculating Control limits.", "Search by Date")         

    elif index_no == []:
        win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    elif sht_run.range('J13').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter the Date in the search box", "Search by Date")
    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Date doesn't exist.", "Search by Date")

    return lcl, ucl, search_date_in

def change_control_limit_poly():
    lcl, ucl, search_date_in = button_control_limit_calc_poly()
    
    if lcl !=0 and ucl !=0 and search_date_in !=0:
        df_repl1a_er_poly = excel_file_sht.parse(sht_name_er_poly, skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
        
        df_repl1a_er_poly = df_repl1a_er_poly[sht_er_nit_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
        # sht_run.range('A23').options(index=False).value = df_repl1a_er_poly        # show the dataframe values into sheet- 'CP Plot'

        index_no = df_repl1a_er_poly[df_repl1a_er_poly['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in datframe column
        if index_no != []:
            # add 11 to get the exact cell's excel row no.
            index_no_final = 11 + index_no[-1]     # consider the latest/last index_no, when there is a failure in QC and it is done multiple times.
            for i in range(0, 300, 3):            
                sht_repl1a_er_poly.range('AE'+str(index_no_final + i)).value = lcl
                sht_repl1a_er_poly.range('AF'+str(index_no_final + i)).value = ucl
            win32api.MessageBox(wb.app.hwnd, "Now, save the excel file & then press RUN button to generate plots.", "Plot chart with new LCL, UCL")         
        elif index_no == []:
            win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    else:
        win32api.MessageBox(wb.app.hwnd, "Either LCL or UCL value is zero OR there is no date mentioned.", "Change control limit")


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



