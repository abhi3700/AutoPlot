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
cp_chart_title = 'CP Plot for ASBE1'  # title for CP chart
cp_chart_xlabel = 'Date'   # xaxis name for CP chart
cp_chart_ylabel = 'delta CP (no.s)'     # yaxis name for CP chart
cp_chart_html_file = 'ASBE1_CP-Plot.html'   # HTML filename for CP chart
cp_chart_trace_count = 3    # no. of traces in CP chart
er_chart_title = 'ER Plot for ASBE1'  # title for ER chart
er_chart_xlabel = 'Date'        # xaxis name for ER chart
er_chart_ylabel = 'Etch Rate (A/min)'   # yaxis name for ER chart
er_chart_html_file = 'ASBE1_ER-Plot.html'   # HTML filename for ER chart
er_chart_trace_count = 4    # no. of traces in ER chart
unif_chart_title = 'Uniformity Plot for ASBE1'  # title for UNIF chart
unif_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_chart_ylabel = 'Uniformity (%)'    # yaxis name for Unif chart
unif_chart_html_file = 'ASBE1_Unif-Plot.html'   # HTML filename for Unif chart
unif_chart_trace_count = 2    # no. of traces in Unif chart

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
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_asbe1_cp = wb.sheets['ASBE1-CP']
    sht_asbe1_er = wb.sheets['ASBE1-ER']
    sht_asbe1_plot_cp = wb.sheets['CP Plot']
    sht_asbe1_plot_er = wb.sheets['ER Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_asbe1_cp = sht_asbe1_cp.range('A10').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'ASBE1-CP'
    df_asbe1_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asbe1_cp = df_asbe1_cp[["Date (MM/DD/YYYY)", "delta CP", "USL", "UCL", "Remarks"]]        # The final dataframe with required columns
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
    # # Draw CP Plot
    # fig_asbe1_cp, ax_asbe1_cp = plt.subplots(1,1, figsize=(25,6))
    # monthyearFmt_cp = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_asbe1_cp.xaxis.set_major_formatter(monthyearFmt_cp)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_asbe1_cp.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    # ax_asbe1_cp.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_asbe1_cp["Date (MM/DD/YYYY)"], df_asbe1_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    # plt.plot(df_asbe1_cp["Date (MM/DD/YYYY)"], df_asbe1_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    # plt.plot(df_asbe1_cp["Date (MM/DD/YYYY)"], df_asbe1_cp["UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('delta CP (no.s)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_asbe1_cp = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),
    #     Line2D([0], [0], color='#0000CD', lw=4),
    #     Line2D([0], [0], color='#FF1493', lw=4)        
    #     ]
    # ax_asbe1_cp.legend(custom_lines_asbe1_cp, ['CP', 'USL', 'UCL'], fontsize=11, loc='upper right')  
    # lines_cp = ax_asbe1_cp.plot(df_asbe1_cp["Date (MM/DD/YYYY)"], df_asbe1_cp["delta CP"], visible=False)
    # # datacursor(lines_cp, hover=True, point_labels=df_asbe1_cp['Remarks'])
    # # plt.show()
    # # sht_asbe1_plot_cp.activate()
    # # pic_cp = plt.show()
    # # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # # sht_asbe1_plot_cp.pictures.add(pic_cp, name= "ASBE1_CP_Plot", update= True)
    # sht_asbe1_plot_cp.pictures.add(fig_asbe1_cp, name= "ASBE1_CP_Plot", update= True)

    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_asbe1_cp_plot(
        x = df_asbe1_cp_date, 
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

    excel_file = pd.ExcelFile("I:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\ASH10_QC_LOG_BOOK\\ASH10_QC_LOG_BOOK.xlsm")
    # excel_file = pd.ExcelFile("\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\ASH_09_10_LOG_BOOK\\ASH10_QC_LOG_BOOK_macro\\ASH10_QC_LOG_BOOK.xlsm")
    df_asbe1_er = excel_file.parse('ASBE1-ER', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 8
    
    df_asbe1_er = df_asbe1_er[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "Remarks"]]             # The final Dataframe with 5 columns for plot: x-1, y-4
    df_asbe1_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
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
    # # Draw ER PLot
    # fig_asbe1_er, ax_asbe1_er = plt.subplots(1,1, figsize=(25,6))
    # monthyearFmt_er = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_asbe1_er.xaxis.set_major_formatter(monthyearFmt_er)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_asbe1_er.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_asbe1_er.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_asbe1_er["Date (MM/DD/YYYY)"], df_asbe1_er["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_asbe1_er["Date (MM/DD/YYYY)"], df_asbe1_er["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('Etch Rate (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),
    #     Line2D([0], [0], color='#0000CD', lw=4),
    #     ]
    # ax_asbe1_er.legend(custom_lines_er, ['ER', 'LSL'], fontsize=11, loc='upper right') 
    # lines_er = ax_asbe1_er.plot(df_asbe1_er["Date (MM/DD/YYYY)"], df_asbe1_er["Etch Rate (A/Min)"], visible=False)
    # # datacursor(lines_er, hover=True, point_labels=df_asbe1_er['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_asbe1_plot_er.pictures.add(fig_asbe1_er, name= "ASBE1_ER_Plot", update= True)
    # Draw ER Plot (using Plotly) in Browser 
    draw_plotly_asbe1_er_plot(
        x = df_asbe1_er_date, 
        y1 = df_asbe1_er_er,
        y2 = df_asbe1_er_lsl, 
        # y3 = df_asbe1_er_ucl,
        # y4 = df_asbe1_er_lcl,
        remarks = df_asbe1_er_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------        
    # Draw Unif PLot
    # fig_asbe1_unif, ax_asbe1_unif = plt.subplots(1,1, figsize=(25,6))
    # monthyearFmt_unif = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_asbe1_unif.xaxis.set_major_formatter(monthyearFmt_unif)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_asbe1_unif.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_asbe1_unif.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_asbe1_er["Date (MM/DD/YYYY)"], df_asbe1_er["% Uni"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('Uniformity (%)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_unif = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),
    #     ]
    # ax_asbe1_unif.legend(custom_lines_unif, ['Unif'], fontsize=11, loc='upper right') 
    # lines_unif = ax_asbe1_unif.plot(df_asbe1_er["Date (MM/DD/YYYY)"], df_asbe1_er["% Uni"], visible=False)
    # # datacursor(lines_unif, hover=True, point_labels=df_asbe1_er['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_asbe1_plot_er.pictures.add(fig_asbe1_unif, name= "ASBE1_Unif_Plot", update= True)

    # Draw Unif Plot (using Plotly) in Browser     
    draw_plotly_asbe1_unif_plot(
        x = df_asbe1_er_date, 
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



