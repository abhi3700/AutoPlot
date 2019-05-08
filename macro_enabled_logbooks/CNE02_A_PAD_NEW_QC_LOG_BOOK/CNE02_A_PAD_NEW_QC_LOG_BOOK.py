# Import packages
import xlwings as xw
import matplotlib.pyplot as plt
import pandas as pd
import matplotlib.dates as mdates
from matplotlib.dates import MO, TU, WE, TH, FR, SA, SU
from matplotlib.lines import Line2D
import plotly as py
import plotly.graph_objs as go
# from matplotlib.figure import Figure
# import datetime as dt
# import win32api
# import os
# from pathlib import Path




#==================================================================================================================================================================
# Global variables
line_color = '#3f51b5'      # line (trace0) color
marker_color = '#43a047'  # marker color
cl_color = '#ffa000'    # control limit line color
sl_color = '#e53935'    # spec limit line color

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with traces v/s Date.
"plot_sht": Excel sheet for CP Chart
"x": Date (x-axis) for CP Chart
"y0": Delta-CP (y-axis) for CP Chart
"y1": USL (y-axis) for CP Chart
"y2": UCL (y-axis) for CP Chart
"""
def draw_reox1a_cp_plot(plot_sht, x, y0, y1, y2):
    # Draw CP Plot
    fig_reox1a_cp, fig_reox1a_cp = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_cp = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    fig_reox1a_cp.xaxis.set_major_formatter(monthyearFmt_cp)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # fig_reox1a_cp.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    fig_reox1a_cp.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    fig_reox1a_cp.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_reox1a_cp["Date (MM/DD/YYYY)"], df_reox1a_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    plt.plot(df_reox1a_cp["Date (MM/DD/YYYY)"], df_reox1a_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('delta CP (no.s)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_cp = [
        Line2D([0], [0], color='#FF7F50', lw=4),
        Line2D([0], [0], color='#0000CD', lw=4),
        ]
    fig_reox1a_cp.legend(custom_lines_cp, ['CP', 'USL'], fontsize=11, loc='upper right')  
    lines_cp = fig_reox1a_cp.plot(df_reox1a_cp["Date (MM/DD/YYYY)"], df_reox1a_cp["delta CP"], visible=False)
    # sht_reox1a_plot_cp.pictures.add(fig_reox1a_cp, name= "REPL1A_CP_Plot", update= True)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# def draw_plotly_repl1a_cp_plot(x, y0, y1, y2, remarks):
def draw_plotly_reox1a_cp_plot(x, y0, y1, remarks):
    trace0 = go.Scatter(
            x = x,
            y = y0,
            name = 'delta-CP',
            mode = 'lines+markers',
            line = dict(
                    color = line_color,
                    width = 2),
            marker = dict(
                    color = marker_color,
                    size = 8,
                    line = dict(
                        color = '#ffffff',
                        width = 0.5),
                    ),
            text = remarks
    )

    trace1 = go.Scatter(
            x = x,
            y = y1,
            name = 'USL',
            mode = 'lines',
            line = dict(
                    color = sl_color,
                    width = 3)
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

    # data = [trace0, trace1, trace2]
    data = [trace0, trace1]
    layout = dict(
            title = 'CP Plot for REOX1A',
            xaxis = dict(title= 'Date'),
            yaxis = dict(title= 'delta CP (no.s)')
        )
    fig = dict(data= data, layout = layout)
    py.offline.plot(fig, filename='REOX1A_CP-Plot.html')



#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_reox1a_cp = wb.sheets['REOX1A-CP']
    sht_reox1a_er = wb.sheets['REOX1A-ER']
    sht_reox1a_plot_cp = wb.sheets['CP Plot']
    sht_reox1a_plot_er_sin = wb.sheets['SiN Plot']
    sht_reox1a_plot_er_teos = wb.sheets['TEOS Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_reox1a_cp = sht_reox1a_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'REOX1A-CP'
    df_reox1a_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reox1a_cp = df_reox1a_cp[["Date (MM/DD/YYYY)", "delta CP", "USL", "Remarks"]]        # The final dataframe with required columns
    # sht_reox1a_plot_cp.range('A25').options(index=False).value = df_reox1a_cp   	    # show the dataframe values into sheet- 'CP Plot'

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
        x = df_reox1a_cp_date, 
        y0 = df_reox1a_cp_delta_cp, 
        y1 = df_reox1a_cp_usl, 
        # y2 = df_reox1a_cp_ucl,
        remarks = df_reox1a_cp_remarks
        )



#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



