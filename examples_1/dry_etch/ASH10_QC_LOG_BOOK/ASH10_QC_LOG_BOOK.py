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
line_color = '#3f51b5'      # line (trace0) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot
cp_plot_title = 'CP Plot for ASBE1'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'ASBE1_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 3    # no. of traces in CP plot
er_plot_title = 'ER Plot for ASBE1'  # title for ER plot
er_plot_xlabel = 'Date'        # xaxis name for ER plot
er_plot_ylabel = 'Etch Rate (A/min)'   # yaxis name for ER plot
er_plot_html_file = 'ASBE1_ER-Plot.html'   # HTML filename for ER plot
er_plot_trace_count = 4    # no. of traces in ER plot
unif_plot_title = 'Uniformity Plot for ASBE1'  # title for UNIF plot
unif_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_plot_ylabel = 'Uniformity (%)'    # yaxis name for Unif plot
unif_plot_html_file = 'ASBE1_Unif-Plot.html'   # HTML filename for Unif plot
unif_plot_trace_count = 2    # no. of traces in Unif plot
sht_cp_columns = ["Date (MM/DD/YYYY)", "delta CP", "USL", "UCL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "Remarks"]

excel_file_directory = "I:\\github_repos\\AutoPlot\\Examples\\dry_etch\\ASH10_QC_LOG_BOOK\\ASH10_QC_LOG_BOOK.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\ASH_09_10_LOG_BOOK\\ASH10_QC_LOG_BOOK_macro\\ASH10_QC_LOG_BOOK.xlsm"

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
"y2": LSL (y-axis) for ER Chart
# "y3": UCL (y-axis) for ER Chart
# "y4": LCL (y-axis) for ER Chart
"""
# def draw_plotly_asbe1_er_plot(x, y1, y2, y3, y4, remarks):
def draw_plotly_asbe1_er_plot(x, y1, y2, remarks):
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

    # trace3 = go.Scatter(
    #         x = x,
    #         y = y3,
    #         name = 'UCL',
    #         mode = 'lines',
    #         line = dict(
    #                 color = cl_color,
    #                 width = 3)
    # )

    # trace4 = go.Scatter(
    #         x = x,
    #         y = y4,
    #         name = 'LCL',
    #         mode = 'lines',
    #         line = dict(
    #                 color = cl_color,
    #                 width = 3)
    # )

    # data = [trace1, trace2, trace3, trace4]
    data = [trace1, trace2]
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
# "y2": UCL (y-axis) for Unif Chart
"""
# def draw_plotly_asbe1_unif_plot(x, y1, y2, remarks):
def draw_plotly_asbe1_unif_plot(x, y1, remarks):
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

    # trace2 = go.Scatter(
    #         x = x,
    #         y = y2,
    #         name = 'UCL',
    #         mode = 'lines',
    #         line = dict(
    #                 color = cl_color,
    #                 width = 3)
    # )

    # data = [trace1, trace2]
    data = [trace1]
    layout = dict(
            title = unif_plot_title,
            xaxis = dict(title= unif_plot_xlabel),
            yaxis = dict(title= unif_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_plot_html_file)


#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_asbe1_cp = wb.sheets['ASBE1-CP']
    sht_asbe1_er = wb.sheets['ASBE1-ER']
    # sht_asbe1_plot_cp = wb.sheets['CP Plot']
    # sht_asbe1_plot_er = wb.sheets['ER Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_asbe1_cp = sht_asbe1_cp.range('A10').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'ASBE1-CP'
    df_asbe1_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
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

    excel_file = pd.ExcelFile(excel_file_directory)
    df_asbe1_er = excel_file.parse('ASBE1-ER', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 8
    df_asbe1_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asbe1_er = df_asbe1_er[sht_er_columns]             # The final Dataframe with 5 columns for plot: x-1, y-4
    df_asbe1_er = df_asbe1_er.dropna()                                              # dropping rows where at least one element is missing
    # sht_asbe1_plot_er.range('A28').options(index=False).value = df_asbe1_er        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for ER & Unif PLot
    df_asbe1_er_date = df_asbe1_er["Date (MM/DD/YYYY)"]
    df_asbe1_er_er = df_asbe1_er["Etch Rate (A/Min)"]
    # df_asbe1_er_ucl = df_asbe1_er["UCL"]
    df_asbe1_er_lsl = df_asbe1_er["LSL"]
    # df_asbe1_er_lcl = df_asbe1_er["LCL"]
    df_asbe1_er_unif = df_asbe1_er["% Uni"]
    # df_asbe1_er_unif_ucl = df_asbe1_er["% Uni UCL"]
    df_asbe1_er_remarks = df_asbe1_er["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------        
    # Draw ER Plot (using Plotly) in Browser 
    draw_plotly_asbe1_er_plot(
        x = date_formatter(df_asbe1_er_date), 
        y1 = df_asbe1_er_er,
        y2 = df_asbe1_er_lsl, 
        # y3 = df_asbe1_er_ucl,
        # y4 = df_asbe1_er_lcl,
        remarks = df_asbe1_er_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------        
    # Draw Unif Plot (using Plotly) in Browser     
    draw_plotly_asbe1_unif_plot(
        x = date_formatter(df_asbe1_er_date), 
        y1 = df_asbe1_er_unif, 
        # y2 = df_asbe1_er_unif_ucl,
        remarks = df_asbe1_er_remarks
        )


#--------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



