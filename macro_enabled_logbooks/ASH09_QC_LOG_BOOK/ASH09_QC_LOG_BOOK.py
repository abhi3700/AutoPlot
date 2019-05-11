# Import packages
import xlwings as xw
import pandas as pd
import plotly as py
import plotly.graph_objs as go
# import datetime as dt
# import win32api
# import os
# from pathlib import Path




#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color for any chart
marker_color = '#43a047'    # marker color for any chart
marker_border_color = '#ffffff'     # marker border color for any chart
cl_color = '#ffa000'    # control limit line color for any chart
sl_color = '#e53935'    # spec limit line color for any chart
cp_chart_title = 'CP Plot for ASFE1'  # title for CP chart
er_chart_title = 'ER Plot for ASFE1'  # title for ER chart
unif_chart_title = 'Uniformity Plot for ASFE1'  # title for UNIF chart
cp_chart_xlabel = 'Date'   # xaxis name for CP chart
cp_chart_ylabel = 'delta CP (no.s)'     # yaxis name for CP chart
cp_chart_html_file = 'ASFE1_CP-Plot.html'   # HTML filename for CP chart
er_chart_xlabel = 'Date'        # xaxis name for ER chart
er_chart_ylabel = 'Etch Rate (A/min)'   # yaxis name for ER chart
er_chart_html_file = 'ASFE1_ER-Plot.html'   # HTML filename for ER chart
unif_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_chart_ylabel = 'Uniformity (%)'    # yaxis name for Unif chart
unif_chart_html_file = 'ASFE1_Unif-Plot.html'   # HTML filename for Unif chart

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with 3 traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
"y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_asfe1_cp_plot(x, y1, y2, y3, remarks):
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
            title = cp_chart_title,
            xaxis = dict(title= cp_chart_xlabel),
            yaxis = dict(title= cp_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with 4 traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": LSL (y-axis) for ER Chart
"y3": UCL (y-axis) for ER Chart
"y4": LCL (y-axis) for ER Chart
"""
def draw_plotly_asfe1_er_plot(x, y1, y2, y3, y4, remarks):
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
            name = 'LSL',
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

    trace4 = go.Scatter(
            x = x,
            y = y4,
            name = 'LCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4]
    layout = dict(
            title = er_chart_title,
            xaxis = dict(title= er_chart_xlabel),
            yaxis = dict(title= er_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with 2 traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": UCL (y-axis) for Unif Chart
"""
def draw_plotly_asfe1_unif_plot(x, y1, y2, remarks):
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
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2]
    layout = dict(
            title = unif_chart_title,
            xaxis = dict(title= unif_chart_xlabel),
            yaxis = dict(title= unif_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_chart_html_file)

#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_asfe1_cp = wb.sheets['ASFE1-CP']
    sht_asfe1_er = wb.sheets['ASFE1-ER']
    # sht_asfe1_plot_cp = wb.sheets['CP Plot']
    # sht_asfe1_plot_er = wb.sheets['ER Plot']


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_asfe1_cp = sht_asfe1_cp.range('A10').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value                                                         # fetch the data from sheet- 'ASFE1-CP'
    df_asfe1_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_cp = df_asfe1_cp[["Date (MM/DD/YYYY)", "delta CP", "USL", "UCL", "Remarks"]]        # The final dataframe with required columns
    df_asfe1_cp = df_asfe1_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_asfe1_plot_cp.range('A25').options(index=False).value = df_asfe1_cp           # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_asfe1_cp_date = df_asfe1_cp["Date (MM/DD/YYYY)"]
    df_asfe1_cp_delta_cp = df_asfe1_cp["delta CP"]
    df_asfe1_cp_usl = df_asfe1_cp["USL"]
    df_asfe1_cp_ucl = df_asfe1_cp["UCL"]
    df_asfe1_cp_remarks = df_asfe1_cp["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_asfe1_cp_plot(
        x = df_asfe1_cp_date, 
        y1 = df_asfe1_cp_delta_cp, 
        y2 = df_asfe1_cp_usl, 
        y3 = df_asfe1_cp_ucl,
        remarks = df_asfe1_cp_remarks
        )
    #****************************************************************************************************************************************************************
    # Fetch Dataframe for ER & Unif Plot
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file = pd.ExcelFile("I:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\ASH09_QC_LOG_BOOK\\ASH09_QC_LOG_BOOK.xlsm")
    # excel_file = pd.ExcelFile("C:\\Users\\abhijit\\Desktop\\dryetch-excel-py-macros\\ASH09_QC_LOG_BOOK\\ASH09_QC_LOG_BOOK.xlsm")
    # excel_file = pd.ExcelFile("\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\ASH_09_10_LOG_BOOK\\ASH09_QC_LOG_BOOK_macro\\ASH09_QC_LOG_BOOK.xlsm")
    df_asfe1_er = excel_file.parse('ASFE1-ER', skiprows=8)                            # copy a sheet and paste into another sheet and skiprows 8
    df_asfe1_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_er = df_asfe1_er[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "LCL", "UCL", "Remarks", "% Uni UCL"]]             # The final Dataframe with 5 columns for plot: x-1, y-4
    df_asfe1_er = df_asfe1_er.dropna()                                              # dropping rows where at least one element is missing
    # sht_asfe1_plot_er.range('A28').options(index=False).value = df_asfe1_er        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for ER & Unif PLot
    df_asfe1_er_date = df_asfe1_er["Date (MM/DD/YYYY)"]
    df_asfe1_er_er = df_asfe1_er["Etch Rate (A/Min)"]
    df_asfe1_er_ucl = df_asfe1_er["UCL"]
    df_asfe1_er_lsl = df_asfe1_er["LSL"]
    df_asfe1_er_lcl = df_asfe1_er["LCL"]
    df_asfe1_er_unif = df_asfe1_er["% Uni"]
    df_asfe1_er_unif_ucl = df_asfe1_er["% Uni UCL"]
    df_asfe1_er_remarks = df_asfe1_er["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ER Plot (using Plotly) in Browser 
    draw_plotly_asfe1_er_plot(
        x = df_asfe1_er_date, 
        y1 = df_asfe1_er_er,
        y2 = df_asfe1_er_lsl, 
        y3 = df_asfe1_er_ucl,
        y4 = df_asfe1_er_lcl,
        remarks = df_asfe1_er_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Unif Plot (using Plotly) in Browser     
    draw_plotly_asfe1_unif_plot(
        x = df_asfe1_er_date, 
        y1 = df_asfe1_er_unif, 
        y2 = df_asfe1_er_unif_ucl,
        remarks = df_asfe1_er_remarks
        )



#--------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



