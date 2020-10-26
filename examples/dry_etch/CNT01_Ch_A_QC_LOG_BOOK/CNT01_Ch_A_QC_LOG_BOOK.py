# Import packages
import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
import win32api         # for message box
import numpy as np
import math
import statistics as stat
from dir import *
from input import *


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

#====================================================================================================================================================================
#####################################################################################################################################################################
def init():
    # Initialize the workbook
    # wb = xw.Book.caller()
    wb = xw.Book('CNT01_Ch_A_QC_LOG_BOOK.xlsm')
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_repl1a_cp = wb.sheets[sht_name_cp]
    sht_repl1a_er_nit = wb.sheets[sht_name_er_nit]
    sht_repl1a_er_poly = wb.sheets[sht_name_er_poly]
    sht_run = wb.sheets['RUN_code']     # for testing purpose
    #****************************************************************************************************************************************************************
    x_coord_nit = sht_repl1a_er_nit.range(x_coord_nit_range).value
    y_coord_nit = sht_repl1a_er_nit.range(y_coord_nit_range).value
    x_coord_poly = sht_repl1a_er_poly.range(x_coord_poly_range).value
    y_coord_poly = sht_repl1a_er_poly.range(y_coord_poly_range).value
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    excel_file = pd.ExcelFile(excel_file_directory)

    return wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file

def button_run():
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_repl1a_cp = sht_repl1a_cp.range('A10').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value                                                         # fetch the data from sheet- 'ASBE1-CP'
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
    draw_plotly_repl1a_cp_plot(
        x = date_formatter(df_repl1a_cp_date), 
        y1 = df_repl1a_cp_delta_cp_1, 
        y2 = df_repl1a_cp_delta_cp_2, 
        y3 = df_repl1a_cp_delta_cp_3, 
        y4 = df_repl1a_cp_usl, 
        y5 = df_repl1a_cp_ucl,
        remarks = df_repl1a_cp_remarks
        )
    #****************************************************************************************************************************************************************
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
# -------------------------------UTILITY--------------------------------------------------
"""
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


# -----------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": fetch dates from excel sheet
"sht_name": name of the sheet
"skiprows_no": skip the no. of rows
"sht_cols": the list of cols to be filtered out from all cols
"cell_cell_range": range of the cell from where past data is to be cleared
"parse_date_cell": parse the dates from this cell from where it will be shown in the drop down option
"""
def fetch_date_pass(sht_name, skiprows_no, sht_cols, clear_cell_range, parse_date_cell):
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    df = excel_file.parse(sht_name, skiprows= skiprows_no)                            # copy a sheet and paste into another sheet and skiprows 9

    df = df[sht_cols]      # select desired columns
    df['Date (MM/DD/YYYY)'].fillna(method='ffill', inplace=True)        # forward fill the empty cells
    df['Result'].fillna(method='ffill', inplace=True)         # forward fill the empty cells
    df['Etch Rate (A/Min)'].fillna(method='ffill', inplace=True)        # forward fill the empty cells

    df.dropna(inplace=True)                                              # dropping rows where at least one element is missing
    df = df[df['Site'] == 'ER_point\n']
    df = df[df['Result'] == 'Pass']
    # sht_run.range('A20').options(index=False).value = df      # show the dataframe values into sheet- 'RUN_code'

    # populate the Date column cells with date
    sht_run.range(clear_cell_range).clear_contents()      # clear content only
    sht_run.range(parse_date_cell).value = df['Date (MM/DD/YYYY)'].tolist()
    return df

# -----------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": fetch dates from excel sheet
"sht_name": name of the sheet
"skiprows_no": skip the no. of rows
"sht_cols": the list of cols to be filtered out from all cols
"cell_cell_range": range of the cell from where past data is to be cleared
"parse_date_cell": parse the dates from this cell from where it will be shown in the drop down option
"""
def fetch_date_all(sht_name, skiprows_no, sht_cols, clear_cell_range, parse_date_cell):
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()
    
    df = excel_file.parse(sht_name, skiprows= skiprows_no)                            # copy a sheet and paste into another sheet and skiprows 9

    df = df[sht_cols]      # select desired columns
    df['Date (MM/DD/YYYY)'].fillna(method='ffill', inplace=True)        # forward fill the empty cells
    df['Result'].fillna(method='ffill', inplace=True)         # forward fill the empty cells
    df['Etch Rate (A/Min)'].fillna(method='ffill', inplace=True)        # forward fill the empty cells

    df.dropna(inplace=True)                                              # dropping rows where at least one element is missing
    df = df[df['Site'] == 'ER_point\n']
    # df = df[df['Result'] == 'Pass']
    # sht_run.range('A20').options(index=False).value = df      # show the dataframe values into sheet- 'RUN_code'

    # populate the Date column cells with date
    sht_run.range(clear_cell_range).clear_contents()      # clear content only
    sht_run.range(parse_date_cell).value = df['Date (MM/DD/YYYY)'].tolist()
    return df


#********************************************************NIT*************************************************************************************************
def button_clear_nit():   # for NIT
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    sht_run.range('AA5:EKP5').clear_contents()
    sht_run.range('J5').clear_contents()
    sht_run.range('J6').clear_contents()
    sht_run.range('K5').clear_contents()
    sht_run.range('L5').clear_contents()
    sht_run.range('K6').clear_contents()
    sht_run.range('L6').clear_contents()

'''
NOTE: Here, "Etch Rate (A/Min)" column has been considered because we have to compare the Control limit calculation from 2 methods:
    - M-1: take last 30 QC days. So, total data_sample = (30 * 13) points
    - M-2: take last 30 QC days. So, total data_sample = 30 points (all `ER_avg` taken)
'''
def button_fetch_date_nit_pass():
    df = fetch_date_pass(
        sht_name = sht_name_er_nit,
        skiprows_no = skiprows_nit,
        sht_cols = sht_er_nit_cl_columns,
        clear_cell_range = 'AA5:EKP5',
        parse_date_cell = 'AA5'
        )
    return df

def button_calc_control_limit_nit():
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    df_repl1a_er_nit = button_fetch_date_nit_pass()

    search_date_in = sht_run.range('J5').value     # input -- to be entered into search box in 'RUN_code' sheet
    df_repl1a_er_nit.index = pd.RangeIndex(len(df_repl1a_er_nit.index))     # reset index 
    index_no = df_repl1a_er_nit[df_repl1a_er_nit['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in dataframe column
    
    lcl = 0     # initialize LCL
    ucl = 0     # initialize UCL
    if index_no != []:
        df_upto_search = df_repl1a_er_nit.iloc[0:index_no[-1]+1]     # returns a dataframe from index 0 to search_index. That's why 1 is added.
        
        if len(df_upto_search) > 30:    # ensure that the length of dataframe is min. 30
            data_last30 = []
            # ------------------M-1----------------------------------------
            for col in sht_er_nit_cl_columns[2:15]:
                data_last30 += df_upto_search[col].tolist()[-N_cl:]     # join the list with last 30 elements of site_1 to site_13
            lcl, ucl = control_limit_calc(data_last30)
            sht_run.range('K5').value = lcl
            sht_run.range('L5').value = ucl
            # ------------------M-2----------------------------------------
            # data_last30 = df_upto_search['Etch Rate (A/Min)'].tolist()
            # lcl, ucl = control_limit_calc(data_last30)
            # sht_run.range('K6').value = lcl
            # sht_run.range('L6').value = ucl
        else:
            win32api.MessageBox(wb.app.hwnd, "There is lesser QC data points available for calculating Control limits.", "Search by Date")         

    elif index_no == []:
        win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    elif sht_run.range('J5').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter the Date in the search box", "Search by Date")
    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Date doesn't exist.", "Search by Date")

    return lcl, ucl, search_date_in

def button_change_control_limit_nit():
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    lcl, ucl, search_date_in = button_calc_control_limit_nit()
    
    if lcl !=0 and ucl !=0 and search_date_in !=0:
        df_repl1a_er_nit = excel_file.parse(sht_name_er_nit, skiprows=skiprows_nit)                            # copy a sheet and paste into another sheet and skiprows 9
        
        df_repl1a_er_nit = df_repl1a_er_nit[sht_er_nit_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
        # sht_run.range('A23').options(index=False).value = df_repl1a_er_nit        # show the dataframe values into sheet- 'CP Plot'

        index_no = df_repl1a_er_nit[df_repl1a_er_nit['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in datframe column
        if index_no != []:
            # add 11 to get the exact cell's excel row no.
            index_no_final = 12 + index_no[-1]     # consider the latest/last index_no, when there is a failure in QC and it is done multiple times.
            for i in range(0, 300, 3):            
                sht_repl1a_er_nit.range('AB'+str(index_no_final + i)).value = lcl
                sht_repl1a_er_nit.range('AC'+str(index_no_final + i)).value = ucl
            win32api.MessageBox(wb.app.hwnd, "Now, save the excel file & then press RUN button to generate plots.", "Plot chart with new LCL, UCL")         
        elif index_no == []:
            win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    else:
        win32api.MessageBox(wb.app.hwnd, "Either LCL or UCL value is zero OR there is no date mentioned.", "Change control limit")


#********************************************************POLY*************************************************************************************************
def button_clear_poly():   # for POLY
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    sht_run.range('AA15:EKP15').clear_contents()
    sht_run.range('J15').clear_contents()
    sht_run.range('J16').clear_contents()
    sht_run.range('K15').clear_contents()
    sht_run.range('L15').clear_contents()
    sht_run.range('K16').clear_contents()
    sht_run.range('L16').clear_contents()


def button_fetch_date_poly_pass():
    df = fetch_date_pass(
        sht_name = sht_name_er_poly,
        skiprows_no = skiprows_poly, 
        sht_cols = sht_er_poly_cl_columns, 
        clear_cell_range = 'AA15:EKS15',
        parse_date_cell = 'AA15'
        )
    return df

def button_calc_control_limit_poly():
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    df_repl1a_er_poly = button_fetch_date_poly_pass()

    search_date_in = sht_run.range('J15').value     # input -- to be entered into search box in 'RUN_code' sheet
    df_repl1a_er_poly.index = pd.RangeIndex(len(df_repl1a_er_poly.index))     # reset index 
    index_no = df_repl1a_er_poly[df_repl1a_er_poly['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in dataframe column
    
    lcl = 0     # initialize LCL
    ucl = 0     # initialize UCL
    if index_no != []:
        df_upto_search = df_repl1a_er_poly.iloc[0:index_no[-1]+1]     # returns a dataframe from index 0 to search_index. That's why 1 is added.

        if len(df_upto_search) > 30:    # ensure that the length of dataframe is min. 30
            data_last30 = []
            # ------------------M-1----------------------------------------
            for col in sht_er_poly_cl_columns[2:19]:
                data_last30 += df_upto_search[col].tolist()[-N_cl:]     # join the list with last 30 elements of site_1 to site_17            
            lcl, ucl = control_limit_calc(data_last30)
            sht_run.range('K15').value = lcl
            sht_run.range('L15').value = ucl
            # ------------------M-2----------------------------------------
            # data_last30 = df_upto_search['Etch Rate (A/Min)'].tolist()
            # lcl, ucl = control_limit_calc(data_last30)
            # sht_run.range('K16').value = lcl
            # sht_run.range('L16').value = ucl
        else:
            win32api.MessageBox(wb.app.hwnd, "There is lesser QC data points available for calculating Control limits.", "Search by Date")         

    elif index_no == []:
        win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    elif sht_run.range('J15').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter the Date in the search box", "Search by Date")
    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Date doesn't exist.", "Search by Date")

    return lcl, ucl, search_date_in

def button_change_control_limit_poly():
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    lcl, ucl, search_date_in = button_calc_control_limit_poly()
    
    if lcl !=0 and ucl !=0 and search_date_in !=0:
        df_repl1a_er_poly = excel_file.parse(sht_name_er_poly, skiprows=skiprows_poly)                            # copy a sheet and paste into another sheet and skiprows 9
        
        df_repl1a_er_poly = df_repl1a_er_poly[sht_er_nit_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
        # sht_run.range('A23').options(index=False).value = df_repl1a_er_poly        # show the dataframe values into sheet- 'CP Plot'

        index_no = df_repl1a_er_poly[df_repl1a_er_poly['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in datframe column
        if index_no != []:
            # add 11 to get the exact cell's excel row no.
            index_no_final = 12 + index_no[-1]     # consider the latest/last index_no, when there is a failure in QC and it is done multiple times.
            for i in range(0, 300, 3):            
                sht_repl1a_er_poly.range('AE'+str(index_no_final + i)).value = lcl
                sht_repl1a_er_poly.range('AF'+str(index_no_final + i)).value = ucl
            win32api.MessageBox(wb.app.hwnd, "Now, save the excel file & then press RUN button to generate plots.", "Plot chart with new LCL, UCL")         
        elif index_no == []:
            win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    else:
        win32api.MessageBox(wb.app.hwnd, "Either LCL or UCL value is zero OR there is no date mentioned.", "Change control limit")

#===================================================Plot Wafer Map via Contour Plot========================================================================
################################################################################################################################################################
# --------------------------------------------------UTILITY---------------------------------------------
"""
"Description": This function plots Contour (2-D) chart of ER for different layers.
"x": x-coordinate for ER Chart
"y": y-coordinate for ER Chart
"z": z values for ER Chart
"fname": filename for ER Chart
"tname": title name for ER Chart
"date": date for ER Chart
"""
def contour_plot(x, y, z, fname, tname, date):
    trace1 = go.Contour(
                x= x,
                y= y,
                z= z,
                contours=dict(
                            # coloring ='heatmap',
                            showlabels = True, # show labels on contours
                            labelfont = dict( # label font properties
                                size = 12,
                                color = 'white',
                               )
                        ),
                )
    

    layout = dict(
                title = 'Contour plot of ' + tname + ' on ' + date + ' for REPL1 Ch A',
                xaxis = dict(
                            title= 'x-axis (in mm)',
                            # type='linear',
                            range= [-100, 100]
                            ),
                yaxis = dict(
                            title= 'y-axis (in mm)',
                            # type='linear',
                            range= [-100, 100]
                            ),
                autosize= True,
                width=800, height=800,

                )

    data = [trace1]
    # fig = dict(data= data, layout= layout)
    fig = go.Figure(data= data, layout= layout)     # an `graph_objs` object created
    fig.add_trace(      # added a scatter plot for showing the markers onto the contour plot
            go.Scatter(
                    mode='markers',
                    x= x,
                    y= y,
                    opacity=0.8,
                    marker=dict(
                        color='#ffffff',
                        size=10,
                        line=dict(
                            color='#ffffff',
                            width= 0.5
                        ),
                        symbol = 'x'
                    ),
                    showlegend=False,
                    text= list(range(1,14))
            )        
        )
    fig.update_layout(      # added a unfilled circle from (-100, -100) to (100, 100) diagonally
        shapes=[
            # unfilled circle
            go.layout.Shape(
                type="circle",
                xref="x",
                yref="y",
                x0= -100,
                y0= -100,
                x1= 100,
                y1= 100,
                line_color="#212121",
            ),
        ]
    )
    py.offline.plot(fig, filename= fname)

#********************************************************NIT*************************************************************************************************
def button_fetch_date_nit_all():
    df = fetch_date_all(
        sht_name = sht_name_er_nit,
        skiprows_no = skiprows_nit,
        sht_cols = sht_er_nit_cl_columns,
        clear_cell_range = 'AA5:EKP5',
        parse_date_cell = 'AA5'
        )
    return df

def button_plot_wafermap_nit():
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    df_repl1a_er_nit = button_fetch_date_nit_all()

    search_date_in = sht_run.range('J5').value     # input -- to be entered into search box in 'RUN_code' sheet
    df_repl1a_er_nit.index = pd.RangeIndex(len(df_repl1a_er_nit.index))     # reset index 
    index_no = df_repl1a_er_nit[df_repl1a_er_nit['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in dataframe column

    if index_no != []:
        if len(index_no) == 1:
            z = df_repl1a_er_nit.iloc[index_no[-1], 2:15].tolist()
            # sht_run.range('A25').value = z
            contour_plot(
                x = x_coord_nit,
                y = y_coord_nit,
                z = z,
                fname = er_nit_contour_fname,
                tname = er_nit_contour_tname,
                date = str(search_date_in.strftime(date_format_contour))[:-8]    # truncate last 8 digits i.e. 00:00:00 after formatting date
                )
        else:
            if sht_run.range('J6').value is None:
                win32api.MessageBox(wb.app.hwnd, "As, QC was repeated on this day. \nSo, please give entry no. from 0 to " + str(len(index_no) - 1) + " in cell-`J6`", "No. of Dates")       
            else:
                index_no_entry = sht_run.range('J6').value
                z = df_repl1a_er_nit.iloc[index_no[int(index_no_entry)], 2:15].tolist()
                # sht_run.range('A25').value = z
                del index_no_entry
                contour_plot(
                    x = x_coord_nit,
                    y = y_coord_nit,
                    z = z,
                    fname = er_nit_contour_fname,
                    tname = er_nit_contour_tname,
                    date = str(search_date_in.strftime(date_format_contour))[:-8]    # truncate last 8 digits i.e. 00:00:00 after formatting date
                    ) 
    
    elif index_no == []:
        win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    elif sht_run.range('J5').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter the Date in the search box", "Search by Date")
    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Date doesn't exist.", "Search by Date")
    

#********************************************************POLY*************************************************************************************************
def button_fetch_date_poly_all():
    df = fetch_date_all(
        sht_name = sht_name_er_poly,
        skiprows_no = skiprows_poly, 
        sht_cols = sht_er_poly_cl_columns, 
        clear_cell_range = 'AA15:EKS15',
        parse_date_cell = 'AA15'
        )
    return df

def button_plot_wafermap_poly():
    wb, sht_repl1a_cp, sht_repl1a_er_nit, sht_repl1a_er_poly, sht_run, x_coord_nit, y_coord_nit, x_coord_poly, y_coord_poly, excel_file = init()

    df_repl1a_er_poly = button_fetch_date_poly_all()

    search_date_in = sht_run.range('J15').value     # input -- to be entered into search box in 'RUN_code' sheet
    df_repl1a_er_poly.index = pd.RangeIndex(len(df_repl1a_er_poly.index))     # reset index 
    index_no = df_repl1a_er_poly[df_repl1a_er_poly['Date (MM/DD/YYYY)'] == search_date_in].index.tolist()     # returns a list with indices matching the search item in dataframe column

    if index_no != []:
        if len(index_no) == 1:
            z = df_repl1a_er_poly.iloc[index_no[-1], 2:19].tolist()
            # sht_run.range('A25').value = z
            contour_plot(
                x = x_coord_poly,
                y = y_coord_poly,
                z = z,
                fname = er_poly_contour_fname,
                tname = er_poly_contour_tname,
                date = str(search_date_in.strftime(date_format_contour))[:-8]    # truncate last 8 digits i.e. 00:00:00 after formatting date
                )
        else:
            if sht_run.range('J16').value is None:
                win32api.MessageBox(wb.app.hwnd, "As, QC was repeated on this day. \nSo, please give entry no. from 0 to " + str(len(index_no) - 1) + " in cell-`J6`", "No. of Dates")       
            else:
                index_no_entry = sht_run.range('J16').value
                z = df_repl1a_er_poly.iloc[index_no[int(index_no_entry)], 2:19].tolist()
                # sht_run.range('A25').value = z
                del index_no_entry
                contour_plot(
                    x = x_coord_poly,
                    y = y_coord_poly,
                    z = z,
                    fname = er_poly_contour_fname,
                    tname = er_poly_contour_tname,
                    date = str(search_date_in.strftime(date_format_contour))[:-8]    # truncate last 8 digits i.e. 00:00:00 after formatting date
                    ) 
    
    elif index_no == []:
        win32api.MessageBox(wb.app.hwnd, "SORRY!, the date was not found.", "Search by Date")         
    elif sht_run.range('J15').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter the Date in the search box", "Search by Date")
    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Date doesn't exist.", "Search by Date")

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
