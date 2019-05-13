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
cp_chart_title = 'CP Plot for RESP1A'  # title for CP chart
cp_chart_xlabel = 'Date'   # xaxis name for CP chart
cp_chart_ylabel = 'delta CP (no.s)'     # yaxis name for CP chart
cp_chart_html_file = 'RESP1A_CP-Plot.html'   # HTML filename for CP chart
cp_chart_trace_count = 3    # no. of traces in CP chart
er_sin_1st_chart_title = 'SiN-1st ER Plot for RESP1A'  # title for ER chart
er_sin_1st_chart_xlabel = 'Date'        # xaxis name for ER chart
er_sin_1st_chart_ylabel = 'SiN-1st ER (A/min)'   # yaxis name for ER chart
er_sin_1st_chart_html_file = 'RESP1A_SiN_1st_ER-Plot.html'   # HTML filename for ER chart
er_sin_1st_chart_trace_count = 5    # no. of traces in Nit ER chart
unif_sin_1st_chart_title = 'SiN-1st Uniformity Plot for RESP1A'  # title for Unif chart
unif_sin_1st_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_sin_1st_chart_ylabel = 'SiN-1st Unif (%)'    # yaxis name for Unif chart
unif_sin_1st_chart_html_file = 'RESP1A_SiN_1st_Unif-Plot.html'   # HTML filename for Unif chart
unif_sin_1st_chart_trace_count = 3    # no. of traces in Nit Unif chart
er_sin_2nd_chart_title = 'SiN-2nd ER Plot for RESP1A'  # title for ER chart
er_sin_2nd_chart_xlabel = 'Date'        # xaxis name for ER chart
er_sin_2nd_chart_ylabel = 'SiN-2nd ER (A/min)'   # yaxis name for ER chart
er_sin_2nd_chart_html_file = 'RESP1A_SiN_2nd_ER-Plot.html'   # HTML filename for ER chart
er_sin_2nd_chart_trace_count = 5    # no. of traces in Nit ER chart
unif_sin_2nd_chart_title = 'SiN-2nd Uniformity Plot for RESP1A'  # title for Unif chart
unif_sin_2nd_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_sin_2nd_chart_ylabel = 'SiN-2nd Unif (%)'    # yaxis name for Unif chart
unif_sin_2nd_chart_html_file = 'RESP1A_SiN_2nd_Unif-Plot.html'   # HTML filename for Unif chart
unif_sin_2nd_chart_trace_count = 3    # no. of traces in Nit Unif chart
er_teos_1st_chart_title = 'TEOS-1st ER Plot for RESP1A'  # title for ER chart
er_teos_1st_chart_xlabel = 'Date'        # xaxis name for ER chart
er_teos_1st_chart_ylabel = 'TEOS-1st ER (A/min)'   # yaxis name for ER chart
er_teos_1st_chart_html_file = 'RESP1A_TEOS_1st_ER-Plot.html'   # HTML filename for ER chart
er_teos_1st_chart_trace_count = 5    # no. of traces in Nit ER chart
unif_teos_1st_chart_title = 'TEOS-1st Uniformity Plot for RESP1A'  # title for Unif chart
unif_teos_1st_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_teos_1st_chart_ylabel = 'TEOS-1st Unif (%)'    # yaxis name for Unif chart
unif_teos_1st_chart_html_file = 'RESP1A_TEOS_1st_Unif-Plot.html'   # HTML filename for Unif chart
unif_teos_1st_chart_trace_count = 2    # no. of traces in Nit Unif chart
er_teos_2nd_chart_title = 'TEOS-2nd ER Plot for RESP1A'  # title for ER chart
er_teos_2nd_chart_xlabel = 'Date'        # xaxis name for ER chart
er_teos_2nd_chart_ylabel = 'TEOS-2nd ER (A/min)'   # yaxis name for ER chart
er_teos_2nd_chart_html_file = 'RESP1A_TEOS_2nd_ER-Plot.html'   # HTML filename for ER chart
er_teos_2nd_chart_trace_count = 5    # no. of traces in Nit ER chart
unif_teos_2nd_chart_title = 'TEOS-2nd Uniformity Plot for RESP1A'  # title for Unif chart
unif_teos_2nd_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_teos_2nd_chart_ylabel = 'TEOS-2nd Unif (%)'    # yaxis name for Unif chart
unif_teos_2nd_chart_html_file = 'RESP1A_TEOS_2nd_Unif-Plot.html'   # HTML filename for Unif chart
unif_teos_2nd_chart_trace_count = 2    # no. of traces in Nit Unif chart

excel_file_directory = "I:\\github_repos\\AutoPlot\\macro_enabled_logbooks\\UNT02_CHA_QC_LOG_BOOK_as_Spacer_Chamber\\UNT02_CHA_QC_LOG_BOOK_as_Spacer_Chamber.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_QC_LOG_BOOK_Ch_A_macro\\CNT01_QC_LOG_BOOK_Ch_A.xlsm"

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
"y3": LSL (y-axis) for CP Chart
"""
def draw_plotly_resp1a_cp_plot(x, y1, y2, y3, remarks):
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

    data = [trace1, trace2]
    layout = dict(
            title = cp_chart_title,
            xaxis = dict(title= cp_chart_xlabel),
            yaxis = dict(title= cp_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_chart_html_file)

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
def draw_plotly_resp1a_er_sin_1st_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_sin_1st_chart_title,
            xaxis = dict(title= er_sin_1st_chart_xlabel),
            yaxis = dict(title= er_sin_1st_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_sin_1st_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_sin_1st_plot(x, y1, y2, y3, remarks):
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
            title = unif_sin_1st_chart_title,
            xaxis = dict(title= unif_sin_1st_chart_xlabel),
            yaxis = dict(title= unif_sin_1st_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_sin_1st_chart_html_file)

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
def draw_plotly_resp1a_er_sin_2nd_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_sin_2nd_chart_title,
            xaxis = dict(title= er_sin_2nd_chart_xlabel),
            yaxis = dict(title= er_sin_2nd_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_sin_2nd_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_sin_2nd_plot(x, y1, y2, y3, remarks):
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
            title = unif_sin_2nd_chart_title,
            xaxis = dict(title= unif_sin_2nd_chart_xlabel),
            yaxis = dict(title= unif_sin_2nd_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_sin_2nd_chart_html_file)

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
def draw_plotly_resp1a_er_teos_1st_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_teos_1st_chart_title,
            xaxis = dict(title= er_teos_1st_chart_xlabel),
            yaxis = dict(title= er_teos_1st_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_teos_1st_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
# "y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_teos_1st_plot(x, y1, y2, remarks):
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
            title = unif_teos_1st_chart_title,
            xaxis = dict(title= unif_teos_1st_chart_xlabel),
            yaxis = dict(title= unif_teos_1st_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_teos_1st_chart_html_file)

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
def draw_plotly_resp1a_er_teos_2nd_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_teos_2nd_chart_title,
            xaxis = dict(title= er_teos_2nd_chart_xlabel),
            yaxis = dict(title= er_teos_2nd_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_teos_2nd_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
# "y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_resp1a_unif_teos_2nd_plot(x, y1, y2, remarks):
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
            title = unif_teos_2nd_chart_title,
            xaxis = dict(title= unif_teos_2nd_chart_xlabel),
            yaxis = dict(title= unif_teos_2nd_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_teos_2nd_chart_html_file)

#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_resp1a_cp = wb.sheets['RESP1A-CP']
    # sht_resp1a_cp = wb.sheets['CP']
    sht_resp1a_er = wb.sheets['RESP1A-ER']
    # sht_resp1a_er = wb.sheets['ER']
    # sht_resp1a_plot_cp = wb.sheets['CP Plot']
    sht_resp1a_plot_sin_1st = wb.sheets['SiN 1st Step Plot']
    sht_resp1a_plot_sin_2nd = wb.sheets['SiN 2nd Step Plot']
    sht_resp1a_plot_teos_1st = wb.sheets['TEOS 1st Step Plot']
    sht_resp1a_plot_teos_2nd = wb.sheets['TEOS 2nd Step Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_resp1a_cp = sht_resp1a_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'ASBE1-CP'
    df_resp1a_cp = df_resp1a_cp[["Date (MM/DD/YYYY)", "delta CP", "LSL", "USL", "Remarks"]]        # The final dataframe with required columns
    df_resp1a_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_resp1a_cp = df_resp1a_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_resp1a_plot_cp.range('A25').options(index=False).value = df_resp1a_cp   	    # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_resp1a_cp_date = df_resp1a_cp["Date (MM/DD/YYYY)"]
    df_resp1a_cp_delta_cp = df_resp1a_cp["delta CP"]
    df_resp1a_cp_usl = df_resp1a_cp["USL"]
    df_resp1a_cp_lsl = df_resp1a_cp["LSL"]
    # df_resp1a_cp_ucl = df_resp1a_cp["UCL"]
    df_resp1a_cp_remarks = df_resp1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw CP Plot
    # fig_resp1a_cp, ax_resp1a_cp = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_cp = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_cp.xaxis.set_major_formatter(monthyearFmt_cp)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_cp.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    # ax_resp1a_cp.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_cp.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_cp["Date (MM/DD/YYYY)"], df_resp1a_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    # plt.plot(df_resp1a_cp["Date (MM/DD/YYYY)"], df_resp1a_cp["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_cp["Date (MM/DD/YYYY)"], df_resp1a_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('delta CP (no.s)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_cp = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),
    #     Line2D([0], [0], color='#0000CD', lw=4),
    #     Line2D([0], [0], color='#0000CD', lw=4),
    #     ]
    # ax_resp1a_cp.legend(custom_lines_cp, ['CP', 'LSL', 'USL'], fontsize=11, loc='upper right')  
    # # lines_cp = ax_resp1a_cp.plot(df_resp1a_cp["Date (MM/DD/YYYY)"], df_resp1a_cp["delta CP"], visible=False)
    # # datacursor(lines_cp, hover=True, point_labels=df_resp1a_cp['Remarks'])
    # # plt.show()
    # # sht_resp1a_cp_plot.activate()
    # # pic_cp = plt.show()
    # # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # # sht_resp1a_cp_plot.pictures.add(pic_cp, name= "ASFE1_CP_Plot", update= True)
    # sht_resp1a_plot_cp.pictures.add(fig_resp1a_cp, name= "RESP1A_CP_Plot", update= True)


    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_resp1a_cp_plot(
        x = df_resp1a_cp_date, 
        y1 = df_resp1a_cp_delta_cp, 
        y2 = df_resp1a_cp_usl, 
        y3 = df_resp1a_cp_lsl,
        remarks = df_resp1a_cp_remarks
        )


    #****************************************************************************************************************************************************************
    """
    Fetch Dataframes for 
        - SIN-1st step, 
        - SIN-2nd step, 
        - TEOS-1st step, 
        - TEOS-2nd step  
    
    """
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_er = pd.ExcelFile(excel_file_directory)
    df_resp1a_er = excel_file_sht_er.parse('RESP1A-ER', skiprows=14)                            # copy a sheet and paste into another sheet and skiprows 9
    df_resp1a_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL' in "Remarks" column 
    df_resp1a_er = df_resp1a_er[["Date (MM/DD/YYYY)", "Layer-Step", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_resp1a_teos = df_resp1a_er.drop(columns='% Uni UCL')      # in TEOS-1st, TEOS-2nd, '% Uni UCL' is not defined, so drop this column.
    df_resp1a_er_sin_1st = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'SiN-1Step']
    df_resp1a_er_sin_1st = df_resp1a_er_sin_1st.dropna()
    df_resp1a_er_sin_2nd = df_resp1a_er[df_resp1a_er["Layer-Step"] == 'SiN-2Step']
    df_resp1a_er_sin_2nd = df_resp1a_er_sin_2nd.dropna()
    df_resp1a_er_teos_1st = df_resp1a_teos[df_resp1a_teos["Layer-Step"] == 'TEOS-1Step']
    df_resp1a_er_teos_1st = df_resp1a_er_teos_1st.dropna()
    df_resp1a_er_teos_2nd = df_resp1a_teos[df_resp1a_teos["Layer-Step"] == 'TEOS-2Step']
    df_resp1a_er_teos_2nd = df_resp1a_er_teos_2nd.dropna()

    # Display the dataframes in respective sheets
    # sht_resp1a_plot_sin_1st.range('A25').options(index=False).value = df_resp1a_sin_1st
    # sht_resp1a_plot_sin_2nd.range('A25').options(index=False).value = df_resp1a_sin_2nd
    # sht_resp1a_plot_teos_1st.range('A25').options(index=False).value = df_resp1a_teos
    # sht_resp1a_plot_teos_2nd.range('A25').options(index=False).value = df_resp1a_teos_2nd


    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw SIN-1st ER PLot
    # fig_resp1a_er_sin_1st, ax_resp1a_er_sin_1st = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_sin_1st = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_er_sin_1st.xaxis.set_major_formatter(monthyearFmt_er_sin_1st)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_er_sin_1st.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_resp1a_er_sin_1st.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_er_sin_1st.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('SiN-1st ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_sin_1st = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),	# ER
    #     Line2D([0], [0], color='#0000CD', lw=4),	# USL
    #     Line2D([0], [0], color='#0000CD', lw=4),	# LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),	# UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)		# LCL        
    #     ]
    # ax_resp1a_er_sin_1st.legend(custom_lines_er_sin_1st, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_sin_1st = ax_resp1a_er_sin_1st.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["Etch Rate (A/Min)"], visible=False)
    # datacursor(lines_er_sin_1st, hover=True, point_labels=df_resp1a_sin_1st['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_resp1a_plot_sin_1st.pictures.add(fig_resp1a_er_sin_1st, name= "RESP1A_SIN_1st_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
	# # Draw SIN-1st Unif PLot
 #    fig_resp1a_unif_sin_1st, ax_resp1a_unif_sin_1st = plt.subplots(1,1, figsize=(20,6))
 #    monthyearFmt_unif_sin_1st = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
 #    ax_resp1a_unif_sin_1st.xaxis.set_major_formatter(monthyearFmt_unif_sin_1st)
 #    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
 #    # ax_resp1a_unif_sin_1st.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
 #    ax_resp1a_unif_sin_1st.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
 #    ax_resp1a_unif_sin_1st.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
 #    plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
 #    plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
 #    plt.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
 #    plt.xlabel('Date', fontsize=18)      # xlabel
 #    plt.ylabel('SiN-1st Unif (%)', fontsize=18)     # ylabel
 #    # Custom Legends
 #    custom_lines_unif_sin_1st = [
 #        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
 #        Line2D([0], [0], color='#0000CD', lw=4),	# USL
 #        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
 #        ]
 #    ax_resp1a_unif_sin_1st.legend(custom_lines_unif_sin_1st, ['Unif', 'USL', 'UCL'], fontsize=11, loc='upper right') 
 #    lines_er_sin_1st = ax_resp1a_unif_sin_1st.plot(df_resp1a_sin_1st["Date (MM/DD/YYYY)"], df_resp1a_sin_1st["% Uniformity"], visible=False)
 #    datacursor(lines_er_sin_1st, hover=True, point_labels=df_resp1a_sin_1st['Remarks'])
 #    # plt.show()        # shows 2 figures in different windows
 #    sht_plot_nit.pictures.add(fig_resp1a_unif_sin_1st, name= "RESP1A_SIN_1st_UNIF_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-1st ER & Unif PLot
    df_resp1a_er_sin_1st_date = df_resp1a_er_sin_1st["Date (MM/DD/YYYY)"]
    df_resp1a_er_sin_1st_er = df_resp1a_er_sin_1st["Etch Rate (A/Min)"]
    df_resp1a_er_sin_1st_usl = df_resp1a_er_sin_1st["USL"]
    df_resp1a_er_sin_1st_lsl = df_resp1a_er_sin_1st["LSL"]
    df_resp1a_er_sin_1st_ucl = df_resp1a_er_sin_1st["UCL"]
    df_resp1a_er_sin_1st_lcl = df_resp1a_er_sin_1st["LCL"]
    df_resp1a_er_sin_1st_unif = df_resp1a_er_sin_1st["% Uni"]
    df_resp1a_er_sin_1st_unif_usl = df_resp1a_er_sin_1st["% Uni USL"]
    df_resp1a_er_sin_1st_unif_ucl = df_resp1a_er_sin_1st["% Uni UCL"]
    df_resp1a_er_sin_1st_remarks = df_resp1a_er_sin_1st["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-1st ER Plot (using Plotly) in Browser 
    draw_plotly_resp1a_er_sin_1st_plot(
        x = df_resp1a_er_sin_1st_date, 
        y1 = df_resp1a_er_sin_1st_er,
        y2 = df_resp1a_er_sin_1st_usl, 
        y3 = df_resp1a_er_sin_1st_lsl,
        y4 = df_resp1a_er_sin_1st_ucl,
        y5 = df_resp1a_er_sin_1st_lcl,
        remarks = df_resp1a_er_sin_1st_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-1st Unif Plot (using Plotly) in Browser     
    draw_plotly_resp1a_unif_sin_1st_plot(
        x = df_resp1a_er_sin_1st_date, 
        y1 = df_resp1a_er_sin_1st_unif, 
        y2 = df_resp1a_er_sin_1st_unif_usl,
        y3 = df_resp1a_er_sin_1st_unif_ucl,
        remarks = df_resp1a_er_sin_1st_remarks
        )

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw SIN-2nd ER PLot
    # fig_resp1a_er_sin_2nd, ax_resp1a_er_sin_2nd = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_sin_2nd = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_er_sin_2nd.xaxis.set_major_formatter(monthyearFmt_er_sin_2nd)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_er_sin_2nd.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_resp1a_er_sin_2nd.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_er_sin_2nd.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('SiN-2nd ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_sin_2nd = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     Line2D([0], [0], color='#0000CD', lw=4),    # LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)     # LCL        
    #     ]
    # ax_resp1a_er_sin_2nd.legend(custom_lines_er_sin_2nd, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_sin_2nd = ax_resp1a_er_sin_2nd.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["Etch Rate (A/Min)"], visible=False)
    # datacursor(lines_er_sin_2nd, hover=True, point_labels=df_resp1a_sin_2nd['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_resp1a_plot_sin_2nd.pictures.add(fig_resp1a_er_sin_1st, name= "RESP1A_SIN_2nd_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw SIN-2nd Unif PLot
    # fig_resp1a_unif_sin_2nd, ax_resp1a_unif_sin_2nd = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_unif_sin_2nd = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_unif_sin_2nd.xaxis.set_major_formatter(monthyearFmt_unif_sin_2nd)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_unif_sin_2nd.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_resp1a_unif_sin_2nd.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_unif_sin_2nd.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('SiN-2nd Unif (%)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_unif_sin_2nd = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     ]
    # ax_resp1a_unif_sin_2nd.legend(custom_lines_unif_sin_2nd, ['Unif', 'USL', 'UCL'], fontsize=11, loc='upper right') 
    # lines_er_sin_2nd = ax_resp1a_unif_sin_2nd.plot(df_resp1a_sin_2nd["Date (MM/DD/YYYY)"], df_resp1a_sin_2nd["% Uniformity"], visible=False)
    # datacursor(lines_er_sin_2nd, hover=True, point_labels=df_resp1a_sin_2nd['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_resp1a_plot_sin_2nd.pictures.add(fig_resp1a_unif_sin_2nd, name= "RESP1A_SIN_2nd_UNIF_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN-2nd ER & Unif PLot
    df_resp1a_er_sin_2nd_date = df_resp1a_er_sin_2nd["Date (MM/DD/YYYY)"]
    df_resp1a_er_sin_2nd_er = df_resp1a_er_sin_2nd["Etch Rate (A/Min)"]
    df_resp1a_er_sin_2nd_usl = df_resp1a_er_sin_2nd["USL"]
    df_resp1a_er_sin_2nd_lsl = df_resp1a_er_sin_2nd["LSL"]
    df_resp1a_er_sin_2nd_ucl = df_resp1a_er_sin_2nd["UCL"]
    df_resp1a_er_sin_2nd_lcl = df_resp1a_er_sin_2nd["LCL"]
    df_resp1a_er_sin_2nd_unif = df_resp1a_er_sin_2nd["% Uni"]
    df_resp1a_er_sin_2nd_unif_usl = df_resp1a_er_sin_2nd["% Uni USL"]
    df_resp1a_er_sin_2nd_unif_ucl = df_resp1a_er_sin_2nd["% Uni UCL"]
    df_resp1a_er_sin_2nd_remarks = df_resp1a_er_sin_2nd["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-2nd ER Plot (using Plotly) in Browser 
    draw_plotly_resp1a_er_sin_2nd_plot(
        x = df_resp1a_er_sin_2nd_date, 
        y1 = df_resp1a_er_sin_2nd_er,
        y2 = df_resp1a_er_sin_2nd_usl, 
        y3 = df_resp1a_er_sin_2nd_lsl,
        y4 = df_resp1a_er_sin_2nd_ucl,
        y5 = df_resp1a_er_sin_2nd_lcl,
        remarks = df_resp1a_er_sin_2nd_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN-2nd Unif Plot (using Plotly) in Browser     
    draw_plotly_resp1a_unif_sin_2nd_plot(
        x = df_resp1a_er_sin_2nd_date, 
        y1 = df_resp1a_er_sin_2nd_unif, 
        y2 = df_resp1a_er_sin_2nd_unif_usl,
        y3 = df_resp1a_er_sin_2nd_unif_ucl,
        remarks = df_resp1a_er_sin_2nd_remarks
        )

    # #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw TEOS-1st ER PLot
    # fig_resp1a_er_teos_1st, ax_resp1a_er_teos_1st = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_teos_1st = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_er_teos_1st.xaxis.set_major_formatter(monthyearFmt_er_teos_1st)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_er_teos_1st.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_resp1a_er_teos_1st.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_er_teos_1st.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('TEOS-1st ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_teos_1st = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     Line2D([0], [0], color='#0000CD', lw=4),    # LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)     # LCL        
    #     ]
    # ax_resp1a_er_teos_1st.legend(custom_lines_er_teos_1st, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_teos_1st = ax_resp1a_er_teos_1st.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["Etch Rate (A/Min)"], visible=False)
    # datacursor(lines_er_teos_1st, hover=True, point_labels=df_resp1a_teos_1st['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_resp1a_plot_teos_1st.pictures.add(fig_resp1a_er_teos_1st, name= "RESP1A_TEOS_1st_ER_Plot", update= True)


    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw TEOS-1st Unif PLot
    # fig_resp1a_unif_teos_1st, ax_resp1a_unif_teos_1st = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_unif_teos_1st = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_unif_teos_1st.xaxis.set_major_formatter(monthyearFmt_unif_teos_1st)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_unif_teos_1st.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_resp1a_unif_teos_1st.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_unif_teos_1st.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # # plt.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('TEOS-1st Unif (%)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_unif_teos_1st = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     # Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     ]
    # ax_resp1a_unif_teos_1st.legend(custom_lines_unif_teos_1st, ['Unif', 'USL'], fontsize=11, loc='upper right') 
    # lines_er_teos_1st = ax_resp1a_unif_teos_1st.plot(df_resp1a_teos_1st["Date (MM/DD/YYYY)"], df_resp1a_teos_1st["% Uniformity"], visible=False)
    # datacursor(lines_er_teos_1st, hover=True, point_labels=df_resp1a_teos_1st['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_resp1a_plot_teos_1st.pictures.add(fig_resp1a_unif_teos_1st, name= "RESP1A_TEOS_1st_UNIF_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-1st ER & Unif PLot
    df_resp1a_er_teos_1st_date = df_resp1a_er_teos_1st["Date (MM/DD/YYYY)"]
    df_resp1a_er_teos_1st_er = df_resp1a_er_teos_1st["Etch Rate (A/Min)"]
    df_resp1a_er_teos_1st_usl = df_resp1a_er_teos_1st["USL"]
    df_resp1a_er_teos_1st_lsl = df_resp1a_er_teos_1st["LSL"]
    df_resp1a_er_teos_1st_ucl = df_resp1a_er_teos_1st["UCL"]
    df_resp1a_er_teos_1st_lcl = df_resp1a_er_teos_1st["LCL"]
    df_resp1a_er_teos_1st_unif = df_resp1a_er_teos_1st["% Uni"]
    df_resp1a_er_teos_1st_unif_usl = df_resp1a_er_teos_1st["% Uni USL"]
    # df_resp1a_er_teos_1st_unif_ucl = df_resp1a_er_teos_1st["% Uni UCL"]
    df_resp1a_er_teos_1st_remarks = df_resp1a_er_teos_1st["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st ER Plot (using Plotly) in Browser 
    draw_plotly_resp1a_er_teos_1st_plot(
        x = df_resp1a_er_teos_1st_date, 
        y1 = df_resp1a_er_teos_1st_er,
        y2 = df_resp1a_er_teos_1st_usl, 
        y3 = df_resp1a_er_teos_1st_lsl,
        y4 = df_resp1a_er_teos_1st_ucl,
        y5 = df_resp1a_er_teos_1st_lcl,
        remarks = df_resp1a_er_teos_1st_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st Unif Plot (using Plotly) in Browser     
    draw_plotly_resp1a_unif_teos_1st_plot(
        x = df_resp1a_er_teos_1st_date, 
        y1 = df_resp1a_er_teos_1st_unif, 
        y2 = df_resp1a_er_teos_1st_unif_usl,
        # y3 = df_resp1a_er_teos_1st_unif_ucl,
        remarks = df_resp1a_er_teos_1st_remarks
        )


    # #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw TEOS-2nd ER PLot
    # fig_resp1a_er_teos_2nd, ax_resp1a_er_teos_2nd = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_teos_2nd = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_er_teos_2nd.xaxis.set_major_formatter(monthyearFmt_er_teos_2nd)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_er_teos_2nd.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_resp1a_er_teos_2nd.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_er_teos_2nd.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('TEOS-2nd ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_teos_2nd = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     Line2D([0], [0], color='#0000CD', lw=4),    # LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)     # LCL        
    #     ]
    # ax_resp1a_er_teos_2nd.legend(custom_lines_er_teos_2nd, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_teos_2nd = ax_resp1a_er_teos_2nd.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["Etch Rate (A/Min)"], visible=False)
    # datacursor(lines_er_teos_2nd, hover=True, point_labels=df_resp1a_teos_2nd['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_resp1a_plot_teos_2nd.pictures.add(fig_resp1a_er_teos_2nd, name= "RESP1A_TEOS_2nd_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw TEOS-2nd Unif PLot
    # fig_resp1a_unif_teos_2nd, ax_resp1a_unif_teos_2nd = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_unif_teos_2nd = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_resp1a_unif_teos_2nd.xaxis.set_major_formatter(monthyearFmt_unif_teos_2nd)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_resp1a_unif_teos_2nd.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_resp1a_unif_teos_2nd.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_resp1a_unif_teos_2nd.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # # plt.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('TEOS-2nd Unif (%)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_unif_teos_2nd = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     # Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     ]
    # ax_resp1a_unif_teos_2nd.legend(custom_lines_unif_teos_2nd, ['Unif', 'USL'], fontsize=11, loc='upper right') 
    # lines_er_teos_2nd = ax_resp1a_unif_teos_2nd.plot(df_resp1a_teos_2nd["Date (MM/DD/YYYY)"], df_resp1a_teos_2nd["% Uniformity"], visible=False)
    # datacursor(lines_er_teos_2nd, hover=True, point_labels=df_resp1a_teos_2nd['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_resp1a_plot_teos_2nd.pictures.add(fig_resp1a_unif_teos_2nd, name= "RESP1A_TEOS_2nd_UNIF_Plot", update= True)
 
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS-2nd ER & Unif PLot
    df_resp1a_er_teos_2nd_date = df_resp1a_er_teos_2nd["Date (MM/DD/YYYY)"]
    df_resp1a_er_teos_2nd_er = df_resp1a_er_teos_2nd["Etch Rate (A/Min)"]
    df_resp1a_er_teos_2nd_usl = df_resp1a_er_teos_2nd["USL"]
    df_resp1a_er_teos_2nd_lsl = df_resp1a_er_teos_2nd["LSL"]
    df_resp1a_er_teos_2nd_ucl = df_resp1a_er_teos_2nd["UCL"]
    df_resp1a_er_teos_2nd_lcl = df_resp1a_er_teos_2nd["LCL"]
    df_resp1a_er_teos_2nd_unif = df_resp1a_er_teos_2nd["% Uni"]
    df_resp1a_er_teos_2nd_unif_usl = df_resp1a_er_teos_2nd["% Uni USL"]
    # df_resp1a_er_teos_2nd_unif_ucl = df_resp1a_er_teos_2nd["% Uni UCL"]
    df_resp1a_er_teos_2nd_remarks = df_resp1a_er_teos_2nd["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st ER Plot (using Plotly) in Browser 
    draw_plotly_resp1a_er_teos_2nd_plot(
        x = df_resp1a_er_teos_2nd_date, 
        y1 = df_resp1a_er_teos_2nd_er,
        y2 = df_resp1a_er_teos_2nd_usl, 
        y3 = df_resp1a_er_teos_2nd_lsl,
        y4 = df_resp1a_er_teos_2nd_ucl,
        y5 = df_resp1a_er_teos_2nd_lcl,
        remarks = df_resp1a_er_teos_2nd_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS-1st Unif Plot (using Plotly) in Browser     
    draw_plotly_resp1a_unif_teos_2nd_plot(
        x = df_resp1a_er_teos_2nd_date, 
        y1 = df_resp1a_er_teos_2nd_unif, 
        y2 = df_resp1a_er_teos_2nd_unif_usl,
        # y3 = df_resp1a_er_teos_2nd_unif_ucl,
        remarks = df_resp1a_er_teos_2nd_remarks
        )





#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



