# Import packages
# import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
from input import *
# import datetime as dt
# import win32api
# import os
# from pathlib import Path


#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# inputs for this `run.py` file
auto_open = False


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
# "y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_reox1a_cp_plot(x, y1, y2, remarks):
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

    data = [trace1, trace2]
    layout = dict(
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_plot_html_file, auto_open= auto_open)

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
def draw_plotly_reox1a_er_sin_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_sin_plot_title,
            xaxis = dict(title= er_sin_plot_xlabel),
            yaxis = dict(title= er_sin_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_sin_plot_html_file, auto_open= auto_open)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1a_unif_sin_plot(x, y1, y2, y3, remarks):
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
            title = unif_sin_plot_title,
            xaxis = dict(title= unif_sin_plot_xlabel),
            yaxis = dict(title= unif_sin_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_sin_plot_html_file, auto_open= auto_open)

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
def draw_plotly_reox1a_er_teos_plot(x, y1, y2, y3, y4, y5, remarks):
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
    py.offline.plot(fig, filename= er_teos_plot_html_file, auto_open= auto_open)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1a_unif_teos_plot(x, y1, y2, y3, remarks):
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
    py.offline.plot(fig, filename= unif_teos_plot_html_file, auto_open= auto_open)



#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    # wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    # sht_reox1a_cp = wb.sheets[sht_name_cp]
    # sht_reox1a_er = wb.sheets[sht_name_er]
    # sht_reox1a_plot_cp = wb.sheets['CP Plot']
    # sht_reox1a_plot_barc = wb.sheets['BARC Plot']
    # sht_reox1a_plot_pr = wb.sheets['SiN Plot']
    # sht_reox1a_plot_teos = wb.sheets['TEOS Plot']
    # sht_reox1a_plot_sin = wb.sheets['SiN Plot']
    excel_file = pd.ExcelFile(excel_file_directory)

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_reox1a_cp = excel_file.parse(sht_name_cp, skiprows=8)
    df_reox1a_cp = df_reox1a_cp[sht_cp_columns]        # The final dataframe with required columns
    df_reox1a_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reox1a_cp = df_reox1a_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_reox1a_plot_cp.range('A25').options(index=False).value = df_reox1a_cp         # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reox1a_cp_date = df_reox1a_cp["Date (MM/DD/YYYY)"]
    df_reox1a_cp_delta_cp = df_reox1a_cp["delta CP"]
    df_reox1a_cp_usl = df_reox1a_cp["USL"]
    # df_reox1a_cp_ucl = df_reox1a_cp["UCL"]
    df_reox1a_cp_remarks = df_reox1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_reox1a_cp_plot(
        x = date_formatter(df_reox1a_cp_date), 
        y1 = df_reox1a_cp_delta_cp, 
        y2 = df_reox1a_cp_usl, 
        # y3 = df_reox1a_cp_ucl,
        remarks = df_reox1a_cp_remarks
        )


    #****************************************************************************************************************************************************************
    """
    Fetch Dataframes for 
        - SiN, 
        - TEOS 
    """
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    df_reox1a_er = excel_file.parse(sht_name_er, skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reox1a_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL' in "Remarks" column 
    df_reox1a_er = df_reox1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1a_er_sin = df_reox1a_er[df_reox1a_er["Layer"] == 'SiN']
    df_reox1a_er_sin = df_reox1a_er_sin.dropna()      # dropping rows where at least one element is missing
    df_reox1a_er_teos = df_reox1a_er[df_reox1a_er["Layer"] == 'TEOS']
    df_reox1a_er_teos = df_reox1a_er_teos.dropna()      # dropping rows where at least one element is missing

    # Display the dataframes in respective sheets
    # sht_reox1a_plot_barc.range('A25').options(index=False).value = df_reox1a_barc
    # sht_reox1a_plot_pr.range('A25').options(index=False).value = df_reox1a_pr
    # sht_reox1a_plot_teos.range('A25').options(index=False).value = df_reox1a_teos


    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN ER & Unif PLot
    df_reox1a_er_sin_date = df_reox1a_er_sin["Date (MM/DD/YYYY)"]
    df_reox1a_er_sin_er = df_reox1a_er_sin["Etch Rate (A/Min)"]
    df_reox1a_er_sin_usl = df_reox1a_er_sin["USL"]
    df_reox1a_er_sin_lsl = df_reox1a_er_sin["LSL"]
    df_reox1a_er_sin_ucl = df_reox1a_er_sin["UCL"]
    df_reox1a_er_sin_lcl = df_reox1a_er_sin["LCL"]
    df_reox1a_er_sin_unif = df_reox1a_er_sin["% Uni"]
    df_reox1a_er_sin_unif_usl = df_reox1a_er_sin["% Uni USL"]
    df_reox1a_er_sin_unif_ucl = df_reox1a_er_sin["% Uni UCL"]
    df_reox1a_er_sin_remarks = df_reox1a_er_sin["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN ER Plot (using Plotly) in Browser 
    draw_plotly_reox1a_er_sin_plot(
        x = date_formatter(df_reox1a_er_sin_date), 
        y1 = df_reox1a_er_sin_er,
        y2 = df_reox1a_er_sin_usl, 
        y3 = df_reox1a_er_sin_lsl,
        y4 = df_reox1a_er_sin_ucl,
        y5 = df_reox1a_er_sin_lcl,
        remarks = df_reox1a_er_sin_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN Unif Plot (using Plotly) in Browser     
    draw_plotly_reox1a_unif_sin_plot(
        x = date_formatter(df_reox1a_er_sin_date), 
        y1 = df_reox1a_er_sin_unif, 
        y2 = df_reox1a_er_sin_unif_usl,
        y3 = df_reox1a_er_sin_unif_ucl,
        remarks = df_reox1a_er_sin_remarks
        )

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS ER & Unif PLot
    df_reox1a_er_teos_date = df_reox1a_er_teos["Date (MM/DD/YYYY)"]
    df_reox1a_er_teos_er = df_reox1a_er_teos["Etch Rate (A/Min)"]
    df_reox1a_er_teos_usl = df_reox1a_er_teos["USL"]
    df_reox1a_er_teos_lsl = df_reox1a_er_teos["LSL"]
    df_reox1a_er_teos_ucl = df_reox1a_er_teos["UCL"]
    df_reox1a_er_teos_lcl = df_reox1a_er_teos["LCL"]
    df_reox1a_er_teos_unif = df_reox1a_er_teos["% Uni"]
    df_reox1a_er_teos_unif_usl = df_reox1a_er_teos["% Uni USL"]
    df_reox1a_er_teos_unif_ucl = df_reox1a_er_teos["% Uni UCL"]
    df_reox1a_er_teos_remarks = df_reox1a_er_teos["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS ER Plot (using Plotly) in Browser 
    draw_plotly_reox1a_er_teos_plot(
        x = date_formatter(df_reox1a_er_teos_date), 
        y1 = df_reox1a_er_teos_er,
        y2 = df_reox1a_er_teos_usl, 
        y3 = df_reox1a_er_teos_lsl,
        y4 = df_reox1a_er_teos_ucl,
        y5 = df_reox1a_er_teos_lcl,
        remarks = df_reox1a_er_teos_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS Unif Plot (using Plotly) in Browser     
    draw_plotly_reox1a_unif_teos_plot(
        x = date_formatter(df_reox1a_er_teos_date), 
        y1 = df_reox1a_er_teos_unif, 
        y2 = df_reox1a_er_teos_unif_usl,
        y3 = df_reox1a_er_teos_unif_ucl,
        remarks = df_reox1a_er_teos_remarks
        )


if __name__ == '__main__':
    main()
